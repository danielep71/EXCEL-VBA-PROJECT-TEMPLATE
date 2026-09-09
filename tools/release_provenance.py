"""Contract 1.2.0 release provenance; called only by the release gate.

Checks assertions and optionally authenticates their signer, not the builder.
Trust policy is read from the candidate Git object, never from release inputs.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import subprocess
import tempfile
from typing import Any

POLICY = ".github/release-provenance.json"
NAMESPACE = "excel-vba-release"
SHA = re.compile(r"[0-9a-f]{40}")


def require(condition: object, message: str) -> None:
    if not condition:
        raise ValueError(message)


def object_keys(value: Any, keys: str, label: str) -> None:
    require(isinstance(value, dict) and set(value) == set(keys.split()),
            f"{label}: invalid object fields")


def nonempty(value: Any) -> bool:
    return isinstance(value, str) and bool(value.strip())


def relative(value: Any) -> bool:
    return (nonempty(value) and "\\" not in value and "\0" not in value
            and not PurePosixPath(value).is_absolute() and ".." not in value.split("/")
            and PurePosixPath(value).as_posix() == value and value != ".")


def decode(raw: bytes) -> Any:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key: {key}")
            result[key] = value
        return result
    return json.loads(raw, object_pairs_hook=unique)


def committed(root: Path, sha: str, path: str) -> bytes:
    require(bool(SHA.fullmatch(sha)) and relative(path), "unsafe Git identity")
    result = subprocess.run(["git", "-C", str(root), "show", f"{sha}:{path}"],
                            capture_output=True, check=False, timeout=30)
    require(result.returncode == 0, f"candidate does not contain {path}")
    return result.stdout


def policy_for(root: Path, sha: str, configuration: dict[str, Any]) -> dict[str, Any]:
    policy = decode(committed(root, sha, POLICY))
    object_keys(policy, "schema_version workflow signature", "policy")
    require(type(policy["schema_version"]) is int and policy["schema_version"] == 1,
            "unsupported provenance policy schema")
    workflow = policy["workflow"]
    object_keys(workflow, "repository path sha", "workflow policy")
    if workflow["repository"] == "@repository":
        workflow["repository"] = configuration["repository"]
    if workflow["sha"] == "@candidate":
        workflow["sha"] = sha
    require(isinstance(workflow["repository"], str) and bool(re.fullmatch(
        r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", workflow["repository"])), "invalid workflow repository")
    require(relative(workflow["path"]) and workflow["path"].startswith(".github/workflows/")
            and workflow["path"].endswith((".yml", ".yaml")), "invalid workflow path")
    require(isinstance(workflow["sha"], str) and SHA.fullmatch(workflow["sha"]),
            "workflow must resolve to an immutable commit")
    if workflow["repository"] == configuration["repository"] and workflow["sha"] == sha:
        committed(root, sha, workflow["path"])
    signature = policy["signature"]
    require(isinstance(signature, dict), "invalid signature policy")
    mode = signature.get("mode")
    if mode == "none":
        object_keys(signature, "mode", "signature policy")
    else:
        require(mode == "ssh", "unsupported signature mode")
        object_keys(signature, "mode principal allowed_signers", "signature policy")
        require(nonempty(signature["principal"]) and relative(signature["allowed_signers"]),
                "signature policy requires a principal and committed allowed-signers path")
    return policy


def inventory(root: Path, evidence: dict[str, Any]) -> list[dict[str, str]]:
    """Dist is the complete payload boundary, including unexpected file types."""
    dist = root / "dist"
    require(not dist.is_symlink(), "dist must not be a symlink")
    actual: dict[str, str] = {}
    if dist.exists():
        require(dist.is_dir(), "dist must be a directory")
        for path in sorted(dist.rglob("*")):
            require(not path.is_symlink(), "release assets must not contain symlinks")
            if path.is_dir():
                continue
            require(path.is_file(), "release assets must be regular files")
            actual[path.relative_to(root).as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    assets = evidence.get("assets")
    require(isinstance(assets, list), "evidence assets must be a list")
    assert isinstance(assets, list)
    expected: dict[str, str] = {}
    for asset in assets:
        require(isinstance(asset, dict) and relative(asset.get("path")), "invalid asset record")
        path = asset["path"]
        require(path.startswith("dist/") and path not in expected, "invalid or duplicate asset path")
        expected[path] = asset.get("sha256")
    require(actual == expected, "dist inventory differs from evidence: missing, extra, or modified assets")
    return [{"path": path, "sha256": digest} for path, digest in sorted(actual.items())]


def validate_record(record: Any, configuration: dict[str, Any], evidence: dict[str, Any],
                    sha: str, policy: dict[str, Any], assets: list[dict[str, str]],
                    manifest_path: Path | None, evidence_bytes: bytes) -> None:
    object_keys(record, "schema_version repository candidate_sha template_contract profile "
                "tag distribution digest_algorithm evidence_sha256 manifest_sha256 assets workflow environment",
                "provenance")
    expected = {
        "schema_version": 1, "repository": configuration["repository"], "candidate_sha": sha,
        "template_contract": configuration["template_contract"], "profile": evidence.get("profile"),
        "tag": evidence.get("tag"), "distribution": evidence.get("distribution"),
        "digest_algorithm": "sha256", "assets": assets,
        "evidence_sha256": hashlib.sha256(evidence_bytes).hexdigest(),
        "manifest_sha256": (hashlib.sha256(manifest_path.read_bytes()).hexdigest()
                            if manifest_path is not None else None),
    }
    require(type(record["schema_version"]) is int, "invalid provenance schema")
    for key, value in expected.items():
        require(record[key] == value, f"provenance {key} does not match candidate inputs")
    workflow = record["workflow"]
    object_keys(workflow, "repository path sha run_id run_attempt", "workflow record")
    for key, value in policy["workflow"].items():
        require(workflow[key] == value, f"workflow {key} differs from candidate trust policy")
    for key in ("run_id", "run_attempt"):
        require(type(workflow[key]) is int and workflow[key] > 0, f"workflow {key} must be positive")
    environment = record["environment"]
    object_keys(environment, "os architecture host host_version office_bitness runtime builder procedure",
                "build environment")
    require(all(nonempty(value) for value in environment.values()), "build environment fields must be nonempty")


def verify_signature(root: Path, sha: str, policy: dict[str, Any], raw: bytes,
                     signature_path: Path | None) -> None:
    signature = policy["signature"]
    if signature["mode"] == "none":
        require(signature_path is None, "signature supplied but verification is not enabled")
        return
    require(signature_path is not None, "enabled SSH verification requires --provenance-signature")
    assert signature_path is not None
    signers = committed(root, sha, signature["allowed_signers"])
    with tempfile.TemporaryDirectory(prefix="release-signature-") as directory:
        trusted = Path(directory) / "allowed_signers"
        detached = Path(directory) / "provenance.sig"
        trusted.write_bytes(signers)
        detached.write_bytes(signature_path.read_bytes())
        result = subprocess.run([
            "ssh-keygen", "-Y", "verify", "-f", str(trusted), "-I", signature["principal"],
            "-n", NAMESPACE, "-s", str(detached),
        ], input=raw, capture_output=True, check=False, timeout=30)
        require(result.returncode == 0, "SSH provenance signature verification failed")


def validate(root: Path, configuration: dict[str, Any], sha: str, evidence_path: Path,
             manifest_path: Path | None, provenance_path: Path | None,
             signature_path: Path | None) -> list[dict[str, str]]:
    try:
        # Read the authority from Git even if the working copy attempts a downgrade.
        configuration = decode(committed(root, sha, ".github/repository-profile.json"))
        contract = configuration.get("template_contract", {})
        version = contract.get("version") if isinstance(contract, dict) else None
        if version in (None, "1.0.0", "1.1.0"):
            require(provenance_path is None and signature_path is None,
                    "advanced provenance requires contract 1.2.0")
            return []
        require(version == "1.2.0", "unsupported provenance contract version")
        policy = policy_for(root, sha, configuration)
        raw_evidence = evidence_path.read_bytes()
        evidence = decode(raw_evidence)
        require(isinstance(evidence, dict), "release evidence must be an object")
        assets = inventory(root, evidence)
        required = evidence.get("distribution") == "binary" or policy["signature"]["mode"] != "none"
        if provenance_path is None:
            require(not required, "this release requires --provenance")
            require(signature_path is None, "signature requires a provenance record")
            return []
        raw = provenance_path.read_bytes()
        validate_record(decode(raw), configuration, evidence, sha, policy, assets,
                        manifest_path, raw_evidence)
        verify_signature(root, sha, policy, raw, signature_path)
        return []
    except (ValueError, KeyError, TypeError, OSError, subprocess.SubprocessError) as error:
        return [{"code": "release-provenance", "path": POLICY, "message": str(error)}]
