#!/usr/bin/env python3
"""Build or verify deterministic, candidate-bound release certification bundles."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any

TOOL_VERSION = "1.0.0"
SCHEMA_VERSION = 1
REQUIRED_ROLES = {
    "excel-evidence",
    "external-links",
    "gate-evidence",
    "release-evidence",
    "release-integrity",
    "wiki-publication",
}
SHA256_RE = re.compile(r"[0-9a-f]{64}$")
SHA40_RE = re.compile(r"[0-9a-f]{40}$")
VERSION_RE = re.compile(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")
SAFE_ID_RE = re.compile(r"[a-z0-9]+(?:-[a-z0-9]+)*$")
SECRET_PATTERNS = (
    re.compile(rb"-----BEGIN [A-Z ]*PRIVATE KEY-----"),
    re.compile(rb"\bghp_[A-Za-z0-9]{20,}\b"),
    re.compile(rb"\bgithub_pat_[A-Za-z0-9_]{20,}\b"),
    re.compile(rb"\bAKIA[A-Z0-9]{16}\b"),
)


class CertificationError(RuntimeError):
    """Certification input or retained bundle violates the contract."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CertificationError(message)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def load_json(path: Path) -> dict[str, Any]:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key in {path}: {key}")
            result[key] = value
        return result

    try:
        value = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=unique)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise CertificationError(f"cannot read {path}: {error}") from error
    require(isinstance(value, dict), f"{path} must contain one JSON object")
    return value


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def git(root: Path, *args: str) -> str:
    try:
        completed = subprocess.run(
            ["git", "-C", str(root), *args],
            check=True,
            capture_output=True,
            text=True,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise CertificationError(f"git {' '.join(args)} failed") from error
    return completed.stdout.strip()


def is_within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def safe_relative(value: str, field: str) -> str:
    path = PurePosixPath(value)
    require(value == path.as_posix(), f"{field} must use normalized POSIX separators")
    require(not path.is_absolute(), f"{field} must be relative")
    require(value not in {"", "."}, f"{field} must not be empty")
    require(all(part not in {"", ".", ".."} for part in path.parts), f"unsafe {field}: {value}")
    return value


def validate_identity(root: Path, specification: dict[str, Any]) -> dict[str, str]:
    require(specification.get("schema_version") == SCHEMA_VERSION, "unsupported specification schema")
    repository = specification.get("repository")
    version = specification.get("version")
    tag = specification.get("tag")
    candidate = specification.get("candidate_sha")
    require(isinstance(repository, str) and re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository) is not None, "repository must use owner/name form")
    require(isinstance(version, str) and VERSION_RE.fullmatch(version) is not None, "version must be SemVer core")
    require(tag == f"v{version}", "tag must equal v + version")
    require(isinstance(candidate, str) and SHA40_RE.fullmatch(candidate) is not None, "candidate_sha must be full lowercase hex")
    require(git(root, "rev-parse", "HEAD") == candidate, "working tree HEAD does not equal candidate_sha")
    require(not git(root, "status", "--porcelain"), "candidate working tree must be clean")
    require((root / "VERSION").read_text(encoding="utf-8").strip() == version, "candidate VERSION disagrees with specification")
    profile = load_json(root / ".github/repository-profile.json")
    require(profile.get("repository") == repository, "candidate repository identity disagrees with specification")
    return {
        "repository": repository,
        "version": version,
        "tag": tag,
        "candidate_sha": candidate,
    }


def validate_public_bytes(data: bytes, record_id: str) -> None:
    for pattern in SECRET_PATTERNS:
        require(pattern.search(data) is None, f"public record {record_id} contains secret-like material")


def validate_records(
    specification: dict[str, Any], evidence_dir: Path
) -> tuple[list[dict[str, Any]], dict[str, bytes]]:
    raw_records = specification.get("records")
    require(isinstance(raw_records, list) and raw_records, "records must be a non-empty array")
    seen_ids: set[str] = set()
    seen_roles: set[str] = set()
    seen_archive_paths: set[str] = set()
    manifest_records: list[dict[str, Any]] = []
    payloads: dict[str, bytes] = {}

    for index, raw in enumerate(raw_records):
        require(isinstance(raw, dict), f"records[{index}] must be an object")
        record_id = raw.get("id")
        role = raw.get("role")
        visibility = raw.get("visibility")
        require(isinstance(record_id, str) and SAFE_ID_RE.fullmatch(record_id) is not None, f"records[{index}].id must be kebab-case")
        require(record_id not in seen_ids, f"duplicate record id: {record_id}")
        require(isinstance(role, str) and SAFE_ID_RE.fullmatch(role) is not None, f"records[{index}].role must be kebab-case")
        require(role not in seen_roles, f"duplicate record role: {role}")
        require(visibility in {"public-file", "restricted-reference"}, f"unsupported visibility for {record_id}")
        seen_ids.add(record_id)
        seen_roles.add(role)

        if visibility == "public-file":
            require(raw.get("redaction_reviewed") is True, f"public record {record_id} must confirm redaction_reviewed=true")
            source = safe_relative(str(raw.get("path") or ""), f"{record_id}.path")
            public_name = safe_relative(str(raw.get("public_name") or ""), f"{record_id}.public_name")
            source_path = evidence_dir / source
            require(source_path.is_file() and not source_path.is_symlink(), f"public record file is unavailable or symlinked: {source}")
            archive_path = f"evidence/{public_name}"
            require(archive_path not in seen_archive_paths, f"duplicate archive path: {archive_path}")
            seen_archive_paths.add(archive_path)
            data = source_path.read_bytes()
            validate_public_bytes(data, record_id)
            payloads[archive_path] = data
            manifest_records.append(
                {
                    "id": record_id,
                    "role": role,
                    "visibility": visibility,
                    "archive_path": archive_path,
                    "sha256": sha256_bytes(data),
                    "size": len(data),
                }
            )
        else:
            digest = raw.get("sha256")
            reference = raw.get("reference")
            reason = raw.get("reason")
            require(isinstance(digest, str) and SHA256_RE.fullmatch(digest) is not None, f"restricted record {record_id} requires a lowercase SHA-256")
            require(isinstance(reference, str) and reference.strip(), f"restricted record {record_id} requires a reference")
            require(isinstance(reason, str) and reason.strip(), f"restricted record {record_id} requires a reason")
            require("\n" not in reference and "\r" not in reference, f"restricted record {record_id} reference must be single-line")
            require("\n" not in reason and "\r" not in reason, f"restricted record {record_id} reason must be single-line")
            manifest_records.append(
                {
                    "id": record_id,
                    "role": role,
                    "visibility": visibility,
                    "sha256": digest,
                    "reference": reference,
                    "reason": reason,
                }
            )

    missing = sorted(REQUIRED_ROLES - seen_roles)
    require(not missing, "required certification roles are missing: " + ", ".join(missing))
    manifest_records.sort(key=lambda row: (str(row["role"]), str(row["id"])))
    return manifest_records, payloads


def zip_bytes(entries: dict[str, bytes]) -> bytes:
    import io

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_STORED) as archive:
        for name in sorted(entries):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_STORED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            archive.writestr(info, entries[name])
    return buffer.getvalue()


def build_bundle(root: Path, specification_path: Path, evidence_dir: Path, output_dir: Path) -> dict[str, Any]:
    root = root.resolve()
    specification_path = specification_path.resolve()
    evidence_dir = evidence_dir.resolve()
    output_dir = output_dir.resolve()
    require(root.is_dir(), "candidate root is unavailable")
    require(evidence_dir.is_dir() and not evidence_dir.is_symlink(), "evidence directory is unavailable or symlinked")
    require(not is_within(specification_path, root), "certification specification must stay outside the candidate tree")
    require(not is_within(evidence_dir, root), "certification evidence must stay outside the candidate tree")
    require(not is_within(output_dir, root), "certification output must stay outside the candidate tree")

    specification = load_json(specification_path)
    identity = validate_identity(root, specification)
    records, payloads = validate_records(specification, evidence_dir)
    base = f"certification-{identity['tag']}"
    bundle_name = base + ".zip"
    manifest_name = base + ".manifest.json"
    digest_name = base + ".sha256"
    tool_path = Path(__file__).resolve()
    manifest = {
        "schema_version": SCHEMA_VERSION,
        **identity,
        "distribution": "certification-evidence",
        "bundle_name": bundle_name,
        "builder": {
            "tool": "tools/release_certification.py",
            "version": TOOL_VERSION,
            "sha256": sha256_bytes(tool_path.read_bytes()),
        },
        "records": records,
        "scope_note": (
            "Certification evidence is retained separately from source/runtime distribution. "
            "restricted-reference records disclose identity/digest/reason without publishing bytes."
        ),
    }
    manifest_bytes = canonical_json(manifest)
    archive_entries = dict(payloads)
    archive_entries["certification-manifest.json"] = manifest_bytes
    bundle = zip_bytes(archive_entries)
    bundle_digest = sha256_bytes(bundle)
    digest_bytes = f"{bundle_digest}  {bundle_name}\n".encode("ascii")

    output_dir.mkdir(parents=True, exist_ok=True)
    for name in (bundle_name, manifest_name, digest_name):
        require(not (output_dir / name).exists(), f"refusing to overwrite existing output: {name}")
    (output_dir / bundle_name).write_bytes(bundle)
    (output_dir / manifest_name).write_bytes(manifest_bytes)
    (output_dir / digest_name).write_bytes(digest_bytes)
    return {
        "bundle": str(output_dir / bundle_name),
        "manifest": str(output_dir / manifest_name),
        "digest": str(output_dir / digest_name),
        "bundle_sha256": bundle_digest,
        "records": len(records),
    }


def verify_bundle(bundle: Path, manifest_path: Path, digest_path: Path) -> dict[str, Any]:
    bundle = bundle.resolve()
    manifest_path = manifest_path.resolve()
    digest_path = digest_path.resolve()
    require(bundle.is_file() and not bundle.is_symlink(), "bundle is unavailable or symlinked")
    require(manifest_path.is_file() and not manifest_path.is_symlink(), "manifest is unavailable or symlinked")
    require(digest_path.is_file() and not digest_path.is_symlink(), "digest file is unavailable or symlinked")
    manifest = load_json(manifest_path)
    require(manifest.get("schema_version") == SCHEMA_VERSION, "unsupported manifest schema")
    require(manifest.get("distribution") == "certification-evidence", "manifest distribution is not certification-evidence")
    require(manifest.get("bundle_name") == bundle.name, "manifest bundle_name mismatches bundle")
    expected_line = digest_path.read_text(encoding="ascii")
    match = re.fullmatch(r"([0-9a-f]{64})  ([^\r\n]+)\n", expected_line)
    require(match is not None, "digest file must contain one canonical SHA-256 line")
    require(match.group(2) == bundle.name, "digest file names another bundle")
    actual_digest = sha256_bytes(bundle.read_bytes())
    require(match.group(1) == actual_digest, "bundle SHA-256 mismatch")

    records = manifest.get("records")
    require(isinstance(records, list), "manifest records must be an array")
    expected_entries = {"certification-manifest.json": canonical_json(manifest)}
    for raw in records:
        require(isinstance(raw, dict), "manifest record must be an object")
        if raw.get("visibility") != "public-file":
            continue
        archive_path = raw.get("archive_path")
        require(isinstance(archive_path, str), "public manifest record requires archive_path")
        expected_entries[archive_path] = b""

    with zipfile.ZipFile(bundle, "r") as archive:
        names = archive.namelist()
        require(names == sorted(names), "bundle entries must be sorted")
        require(set(names) == set(expected_entries), "bundle entry set disagrees with manifest")
        require(archive.read("certification-manifest.json") == canonical_json(manifest), "internal manifest differs from published manifest")
        public_by_path = {
            str(raw["archive_path"]): raw
            for raw in records
            if isinstance(raw, dict) and raw.get("visibility") == "public-file"
        }
        for name, raw in public_by_path.items():
            data = archive.read(name)
            require(raw.get("sha256") == sha256_bytes(data), f"digest mismatch for {name}")
            require(raw.get("size") == len(data), f"size mismatch for {name}")
    return {
        "status": "pass",
        "bundle": bundle.name,
        "bundle_sha256": actual_digest,
        "candidate_sha": manifest.get("candidate_sha"),
        "tag": manifest.get("tag"),
        "records": len(records),
    }


def fixture_git(root: Path, *args: str) -> str:
    completed = subprocess.run(
        ["git", "-C", str(root), "-c", "user.name=Fixture", "-c", "user.email=fixture@example.invalid", *args],
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def run_self_test() -> int:
    failures: list[str] = []
    try:
        with tempfile.TemporaryDirectory(prefix="release-certification-") as raw:
            base = Path(raw)
            root = base / "repo"
            evidence = base / "evidence"
            root.mkdir()
            evidence.mkdir()
            fixture_git(root, "init", "-b", "main")
            (root / ".github").mkdir()
            (root / "VERSION").write_text("1.2.3\n", encoding="utf-8")
            (root / ".github/repository-profile.json").write_text(
                json.dumps({"repository": "owner/repo"}) + "\n", encoding="utf-8"
            )
            (root / "tools").mkdir()
            (root / "tools/release_certification.py").write_text("fixture\n", encoding="utf-8")
            fixture_git(root, "add", "--all")
            fixture_git(root, "commit", "-m", "fixture")
            candidate = fixture_git(root, "rev-parse", "HEAD")
            records = []
            for role in sorted(REQUIRED_ROLES):
                filename = role + ".json"
                (evidence / filename).write_text(json.dumps({"role": role}) + "\n", encoding="utf-8")
                records.append(
                    {
                        "id": role,
                        "role": role,
                        "visibility": "public-file",
                        "path": filename,
                        "public_name": filename,
                        "redaction_reviewed": True,
                    }
                )
            records.append(
                {
                    "id": "restricted-provider-record",
                    "role": "restricted-provider-record",
                    "visibility": "restricted-reference",
                    "sha256": "a" * 64,
                    "reference": "provider-record:123",
                    "reason": "provider access restriction",
                }
            )
            specification = base / "certification-input.json"
            specification.write_text(
                json.dumps(
                    {
                        "schema_version": 1,
                        "repository": "owner/repo",
                        "version": "1.2.3",
                        "tag": "v1.2.3",
                        "candidate_sha": candidate,
                        "records": records,
                    },
                    indent=2,
                    sort_keys=True,
                )
                + "\n",
                encoding="utf-8",
            )
            first = base / "out-1"
            second = base / "out-2"
            result_one = build_bundle(root, specification, evidence, first)
            result_two = build_bundle(root, specification, evidence, second)
            first_bundle = Path(result_one["bundle"])
            second_bundle = Path(result_two["bundle"])
            require(first_bundle.read_bytes() == second_bundle.read_bytes(), "bundle is not deterministic")
            verified = verify_bundle(
                first_bundle,
                Path(result_one["manifest"]),
                Path(result_one["digest"]),
            )
            require(verified["candidate_sha"] == candidate, "verification lost candidate binding")

            tampered = base / "tampered.zip"
            tampered.write_bytes(first_bundle.read_bytes() + b"x")
            try:
                verify_bundle(tampered, Path(result_one["manifest"]), Path(result_one["digest"]))
                failures.append("tampered bundle was accepted")
            except CertificationError:
                pass

            bad_spec = json.loads(specification.read_text(encoding="utf-8"))
            bad_spec["records"] = [row for row in bad_spec["records"] if row["role"] != "excel-evidence"]
            missing = base / "missing.json"
            missing.write_text(json.dumps(bad_spec), encoding="utf-8")
            try:
                build_bundle(root, missing, evidence, base / "out-missing")
                failures.append("missing required role was accepted")
            except CertificationError:
                pass

            secret_name = "release-evidence.json"
            (evidence / secret_name).write_text("-----BEGIN PRIVATE KEY-----\n", encoding="utf-8")
            try:
                build_bundle(root, specification, evidence, base / "out-secret")
                failures.append("secret-like public evidence was accepted")
            except CertificationError:
                pass
    except (CertificationError, OSError, subprocess.SubprocessError, zipfile.BadZipFile) as error:
        failures.append(str(error))

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print("SELF-TEST PASS: deterministic build, exact-SHA binding, verification, required-role, tamper, restricted-reference, and secret-leakage controls passed.")
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--build", type=Path, metavar="SPECIFICATION")
    parser.add_argument("--evidence-dir", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--verify-bundle", type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--digest", type=Path)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        if options.self_test:
            return run_self_test()
        if options.build is not None:
            require(options.evidence_dir is not None, "--build requires --evidence-dir")
            require(options.output_dir is not None, "--build requires --output-dir")
            require(options.verify_bundle is None and options.manifest is None and options.digest is None, "build and verify modes are mutually exclusive")
            result = build_bundle(options.root, options.build, options.evidence_dir, options.output_dir)
        elif options.verify_bundle is not None:
            require(options.manifest is not None and options.digest is not None, "--verify-bundle requires --manifest and --digest")
            require(options.evidence_dir is None and options.output_dir is None, "verify mode does not accept build directories")
            result = verify_bundle(options.verify_bundle, options.manifest, options.digest)
        else:
            raise CertificationError("choose --build, --verify-bundle, or --self-test")
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (CertificationError, OSError, UnicodeError, zipfile.BadZipFile) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
