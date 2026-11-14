#!/usr/bin/env python3
"""Assemble release evidence and the certification bundle for one exact candidate.

``prepare`` runs before the tag. From a clean checkout of the candidate and the
bundle tools/Certify-Excel.ps1 wrote for it, it:

- validates the Excel host record;
- reads the hosted runs at the candidate through ``gh api``, and for the
  canonical template also the live rulesets, the external-link observation and
  the published Wiki;
- runs the local release gates;
- writes ``release-evidence.json`` with the ``excel-host-evidence`` binding
  computed;
- runs ``check_release.py`` before the tag.

``bundle`` runs after the signed tag exists. It runs the post-tag
``check_release.py``, writes the certification specification with one unique
role per record, builds the bundle with ``release_certification.py``, verifies
it, and rebuilds it to prove the build is deterministic. It signs nothing,
publishes nothing and never pushes: tag and ZIP signatures and the GitHub
Release stay with the maintainer.

Everything is written under ``--work-dir``, outside the candidate tree.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any

import check_excel_evidence as excel

STATIC_WORKFLOW = ".github/workflows/static-checks.yml"
MATRIX_WORKFLOW = ".github/workflows/generated-profile-matrix.yml"
LINKS_WORKFLOW = "external-links.yml"
TEMPLATE_GATES = (
    ("check_release_semantics.py", "--self-test"),
    ("check_release_semantics.py", "--root", "."),
    ("check_repo.py", "--root", ".", "--self-test"),
    ("check_repo.py", "--root", "."),
    ("check_release.py", "--root", ".", "--self-test"),
    ("checker_development.py", "--root", ".", "--self-test"),
    ("check_policy_coverage.py", "--root", ".", "--self-test"),
    ("release_certification.py", "--self-test"),
    ("_release_closeout.py", "--self-test"),
)
GENERATED_GATES = (
    ("check_release_semantics.py", "--root", "."),
    ("check_repo.py", "--root", "."),
    ("check_vba_public_api.py", "--root", "."),
)
# Files of the Certify-Excel bundle keep their names, so the copied host.json still
# resolves its logs; the workbook itself is not evidence and stays out.
EXCEL_FILES = ("host.json", "session.json", "suites.json", "certification.md")
# Fixed bundle roles; any other evidence file gets a role equal to its own id.
ROLES = {
    "host.json": ("excel-host-record", "excel-evidence"),
    "external-links.json": ("external-links-report", "external-links"),
    "gate-evidence.md": ("gate-summary", "gate-evidence"),
    "live-governance.json": ("live-governance", "live-governance"),
    "release-evidence.json": ("release-evidence", "release-evidence"),
    "release-integrity-posttag.json": ("release-integrity-posttag", "release-integrity"),
    "release-integrity-pretag.json": ("release-integrity-pretag", "release-integrity-pretag"),
    "wiki-publication.json": ("wiki-publication", "wiki-publication"),
}


class EvidenceError(RuntimeError):
    """A prerequisite or check failed; nothing after it is trustworthy."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def run(command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
    """The one boundary to git, gh and the sibling tools; tests replace it."""
    return subprocess.run(command, cwd=cwd, capture_output=True, check=False)


def checked(command: list[str], cwd: Path, what: str) -> bytes:
    completed = run(command, cwd)
    if completed.returncode != 0:
        detail = completed.stderr.decode("utf-8", "replace").strip().splitlines()
        raise EvidenceError(f"{what} failed: {detail[-1] if detail else 'exit ' + str(completed.returncode)}")
    return completed.stdout


def gh_json(root: Path, endpoint: str) -> Any:
    return json.loads(checked(["gh", "api", endpoint], root, f"gh api {endpoint}"))


def tool(root: Path, name: str, *arguments: str) -> list[str]:
    """A sibling tool of the candidate itself, never of another checkout."""
    return [sys.executable, str(root / "tools" / name), *arguments]


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def candidate_identity(root: Path, tag: str, sha: str) -> dict[str, Any]:
    head = checked(["git", "rev-parse", "HEAD"], root, "git rev-parse").decode().strip()
    require(head == sha, f"checkout is at {head}, not the candidate {sha}")
    status = checked(["git", "status", "--porcelain"], root, "git status").decode().strip()
    require(not status, "the checkout has uncommitted changes")
    version = (root / "VERSION").read_text(encoding="utf-8").strip()
    require(tag == f"v{version}", f"tag {tag} does not match VERSION {version}")
    profile = json.loads((root / ".github/repository-profile.json").read_text(encoding="utf-8"))
    template = profile.get("mode") == "template"
    return {"repository": profile["repository"], "version": version, "tag": tag, "sha": sha,
            "template": template, "profile": "template" if template else profile["profile"]}


def successful_run(runs: list[dict[str, Any]], path: str, sha: str) -> dict[str, Any]:
    """The latest completed push run of one workflow at the candidate; it must have succeeded."""
    matching = [item for item in runs if item.get("path") == path and item.get("head_sha") == sha
                and item.get("event") == "push" and item.get("status") == "completed"]
    require(bool(matching), f"no completed push run of {path} at the candidate")
    latest = max(matching, key=lambda item: (str(item.get("created_at", "")), int(item.get("id", 0))))
    require(latest.get("conclusion") == "success",
            f"{path} at the candidate concluded {latest.get('conclusion')}: {latest.get('html_url')}")
    return latest


def hosted_runs(root: Path, identity: dict[str, Any]) -> list[dict[str, Any]]:
    data = gh_json(root, f"repos/{identity['repository']}/actions/runs?event=push"
                         f"&head_sha={identity['sha']}&per_page=100")
    runs = data.get("workflow_runs", [])
    successful_run(runs, STATIC_WORKFLOW, identity["sha"])
    if identity["template"]:
        successful_run(runs, MATRIX_WORKFLOW, identity["sha"])
    return [item for item in runs if item.get("event") == "push"]


def governance(rulesets: list[dict[str, Any]]) -> dict[str, Any]:
    """Judge the live branch and tag rulesets against the release baseline."""
    def rules(ruleset: dict[str, Any]) -> dict[str, Any]:
        return {rule["type"]: rule.get("parameters") or {} for rule in ruleset.get("rules", [])}

    def active(target: str, ref: str) -> dict[str, Any]:
        # As in the provisioner, a ruleset with exclusions cannot prove coverage.
        scope = [(item, item.get("conditions", {}).get("ref_name", {})) for item in rulesets]
        found = [item for item, names in scope if item.get("target") == target
                 and item.get("enforcement") == "active"
                 and ref in names.get("include", []) and not names.get("exclude")]
        require(bool(found), f"no active {target} ruleset covers {ref} without exclusions")
        return found[0]

    branch, tag = active("branch", "~DEFAULT_BRANCH"), active("tag", "refs/tags/v*")
    branch_rules, tag_rules = rules(branch), rules(tag)
    status = branch_rules.get("required_status_checks", {})
    contexts = [item.get("context") for item in status.get("required_status_checks", [])]
    checks = {
        "branch: deletion blocked": "deletion" in branch_rules,
        "branch: force-push blocked": "non_fast_forward" in branch_rules,
        "branch: pull request required": "pull_request" in branch_rules,
        "branch: 'Repository integrity' required, strict": (
            "Repository integrity" in contexts and status.get("strict_required_status_checks_policy") is True),
        "tag: deletion blocked": "deletion" in tag_rules,
        "tag: force-update blocked": "non_fast_forward" in tag_rules,
        "tag: update blocked (no fetch-and-merge)": (
            "update" in tag_rules and not tag_rules["update"].get("update_allows_fetch_and_merge", False)),
        "no bypass actors": not branch.get("bypass_actors") and not tag.get("bypass_actors"),
    }
    failed = [name for name, ok in checks.items() if not ok]
    require(not failed, "live governance is below the baseline: " + "; ".join(failed))
    return {"rulesets": {"branch": {"id": branch["id"], "name": branch["name"]},
                         "tag": {"id": tag["id"], "name": tag["name"]}},
            "checks": {name: "PASS" for name in checks}, "raw": {"branch": branch, "tag": tag}}


def live_governance(root: Path, identity: dict[str, Any]) -> dict[str, Any]:
    listing = gh_json(root, f"repos/{identity['repository']}/rulesets")
    full = [gh_json(root, f"repos/{identity['repository']}/rulesets/{item['id']}") for item in listing]
    record = governance(full)
    return {"schema_version": 1, "role": "live-governance", "repository": identity["repository"],
            "method": "GitHub REST read-back: GET /repos/{repo}/rulesets and /rulesets/{id}", **record}


def link_counts(report: dict[str, Any]) -> dict[str, Any]:
    counts: dict[str, Any] = report.get("counts", {})
    require(report.get("status") == "pass", f"external-link observation status is {report.get('status')}")
    require(counts.get("deterministic_public_defects") == 0,
            f"external-link observation found {counts.get('deterministic_public_defects')} defects")
    return counts


def external_links(root: Path, identity: dict[str, Any], evidence: Path) -> dict[str, Any]:
    repo, sha = identity["repository"], identity["sha"]
    runs = gh_json(root, f"repos/{repo}/actions/workflows/{LINKS_WORKFLOW}/runs?head_sha={sha}&per_page=20")
    done = [item for item in runs.get("workflow_runs", [])
            if item.get("status") == "completed" and item.get("conclusion") == "success"]
    require(bool(done), f"no successful {LINKS_WORKFLOW} run at the candidate: dispatch it on main first")
    latest = max(done, key=lambda item: int(item["id"]))
    artifacts = gh_json(root, f"repos/{repo}/actions/runs/{latest['id']}/artifacts").get("artifacts", [])
    require(bool(artifacts), f"run {latest['id']} retained no artifact")
    raw = checked(["gh", "api", f"repos/{repo}/actions/artifacts/{artifacts[0]['id']}/zip"], root,
                  "external-link artifact download")
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        report_bytes = archive.read("external-links.json")
    counts = link_counts(json.loads(report_bytes))
    (evidence / "external-links.json").write_bytes(report_bytes)
    return {"run": latest, "artifact_id": artifacts[0]["id"],
            "artifact_sha256": hashlib.sha256(raw).hexdigest(), "counts": counts}


def wiki_publication(root: Path, identity: dict[str, Any], work: Path, reviewed: bool) -> dict[str, Any]:
    require(reviewed, "confirm the browser review of the published Wiki with --wiki-browser-reviewed")
    clone = work / "wiki-readback"
    if clone.exists():
        shutil.rmtree(clone)
    checked(["git", "clone", "--quiet", f"https://github.com/{identity['repository']}.wiki.git", str(clone)],
            root, "Wiki clone")
    checked(tool(root, "check_wiki.py", "--root", ".", "--published-dir", str(clone)), root,
            "Wiki read-back comparison with the candidate")
    log = checked(["git", "-C", str(clone), "log", "-2", "--format=%H"], root, "Wiki log").decode().split()
    files = {path.name: hashlib.sha256(path.read_bytes()).hexdigest()
             for path in sorted(clone.iterdir()) if path.is_file() and path.suffix in (".md", ".json")}
    return {"schema_version": 1, "role": "wiki-publication", "repository": identity["repository"],
            "source_sha": identity["sha"], "wiki_repository": f"{identity['repository']}.wiki",
            "wiki_commit": log[0], "wiki_parent": log[1] if len(log) > 1 else None,
            "readback": "fresh clone; tools/check_wiki.py --published-dir", "readback_result": "PASS",
            "files": files, "browser_review": "PASS: confirmed by the maintainer with --wiki-browser-reviewed"}


def local_gates(root: Path, template: bool) -> list[str]:
    results = []
    for gate in TEMPLATE_GATES if template else GENERATED_GATES:
        checked(tool(root, *gate), root, " ".join(gate))
        results.append(" ".join(gate))
    return results


def excel_record(root: Path, sha: str, bundle: Path, evidence: Path) -> dict[str, Any]:
    host = bundle / "host.json"
    require(host.is_file(), f"{bundle} has no host.json: run tools/Certify-Excel.ps1 on the candidate")
    report = excel.evaluate(root, sha, host)
    require(report["status"] == "pass", "the Excel host record does not validate: " + ", ".join(report["outcomes"]))
    raw = host.read_bytes()
    for path in sorted(bundle.iterdir()):
        if path.is_file() and (path.name in EXCEL_FILES or path.suffix == ".log"):
            shutil.copyfile(path, evidence / path.name)
    copied = excel.evaluate(root, sha, evidence / "host.json")
    require(copied["status"] == "pass", "the copied Excel record no longer validates")
    return {"record": json.loads(raw), "sha256": hashlib.sha256(raw).hexdigest()}


def release_checks(identity: dict[str, Any], host: dict[str, Any], runs: list[dict[str, Any]],
                   policy: dict[str, Any], extra: dict[str, str], facts: dict[str, Any]) -> dict[str, Any]:
    sha, record = identity["sha"], host["record"]
    environment, harness = excel.environment_summary(record), record["harness"]

    def check(detail: str, **fields: Any) -> dict[str, Any]:
        return {"status": "PASS", "candidate_sha": sha, "detail": detail, **fields}

    static = successful_run(runs, STATIC_WORKFLOW, sha)
    checks = {
        "repository-integrity": check("Static repository checks on push at the candidate succeeded; "
                                      "the local release gates passed (gate-evidence.md).",
                                      run_url=static["html_url"]),
        "vba-compile": check("tools/Certify-Excel.ps1 imported every configured component and "
                             "compiled with no error.", environment=environment),
        "regression": check(f"{harness['entry_point']}: RESULT=PASS; completeness={harness['completeness']}; "
                            f"cases={harness['cases']}; assertions={harness['assertions']}; failures=0.",
                            entry_point=harness["entry_point"], environment=environment,
                            cases=harness["cases"], assertions=harness["assertions"],
                            failures=harness["failures"], completeness=harness["completeness"],
                            cleanup="PASS"),
        "excel-host-evidence": check("Structured host record from tools/Certify-Excel.ps1; validates as PASS "
                                     "with check_excel_evidence.py.",
                                     sha256=host["sha256"], execution=record["execution"]),
    }
    if identity["template"]:
        matrix = successful_run(runs, MATRIX_WORKFLOW, sha)
        checks["generated-profile-pilots"] = check(
            f"Generated profile CI matrix at the candidate succeeded: {matrix['html_url']}."
            + (f" {facts['pilot']}" if facts.get("pilot") else ""))
        checks["live-governance"] = check(
            "Live ruleset read-back meets the baseline with no bypass actors (live-governance.json).")
    for name, detail in extra.items():
        checks[name] = check(detail)
    required = list(policy["core_checks"]) + list(policy["profiles"][identity["profile"]]["required_checks"])
    missing = [name for name in required if name not in checks]
    require(not missing, "supply the profile checks with --check NAME=DETAIL: " + ", ".join(missing))
    return checks


def gate_summary(identity: dict[str, Any], runs: list[dict[str, Any]], gates: list[str],
                 links: dict[str, Any] | None) -> str:
    lines = [f"# {identity['tag']} gate evidence", "", f"Candidate: `{identity['sha']}`", "",
             "## Hosted runs on push at the candidate", "", "| Workflow | Run | Result |", "| --- | --- | --- |"]
    for item in sorted(runs, key=lambda value: str(value.get("name", ""))):
        lines.append(f"| {item.get('name')} | {item.get('html_url')} | {item.get('conclusion')} |")
    lines += ["", "## Local gates at the exact candidate (clean tree)", "", "| Gate | Result |", "| --- | --- |"]
    lines += [f"| `{gate}` | PASS |" for gate in gates]
    if links:
        counts = links["counts"]
        lines += ["", "## External-link observation", "",
                  f"Run {links['run']['html_url']} at the candidate: 0 deterministic defects, "
                  f"{counts.get('restricted_historical', 0)} restricted historical, "
                  f"{counts.get('pending_publication', 0)} pending publication, "
                  f"{counts.get('access_restricted', 0)} access restricted, "
                  f"{counts.get('transient_failures', 0)} transient. Artifact {links['artifact_id']}, "
                  f"SHA-256 {links['artifact_sha256']}."]
    return "\n".join(lines) + "\n"


def parse_checks(values: list[str]) -> dict[str, str]:
    result = {}
    for value in values:
        name, separator, detail = value.partition("=")
        require(bool(separator and name and detail), f"--check needs NAME=DETAIL, got {value!r}")
        result[name] = detail
    return result


def prepare(options: argparse.Namespace) -> int:
    root, work = options.root.resolve(), options.work_dir.resolve()
    extra = parse_checks(options.check)
    identity = candidate_identity(root, options.tag, options.candidate_sha)
    evidence = work / "evidence"
    require(not evidence.exists(), f"{evidence} already exists: use a fresh --work-dir")
    evidence.mkdir(parents=True)
    host = excel_record(root, identity["sha"], options.excel_bundle.resolve(), evidence)
    runs = hosted_runs(root, identity)
    links = None
    if identity["template"]:
        write_json(evidence / "live-governance.json", live_governance(root, identity))
        links = external_links(root, identity, evidence)
        write_json(evidence / "wiki-publication.json",
                   wiki_publication(root, identity, work, options.wiki_browser_reviewed))
    if options.pilot_record:
        shutil.copyfile(options.pilot_record, evidence / options.pilot_record.name)
    gates = local_gates(root, identity["template"])
    (evidence / "gate-evidence.md").write_text(gate_summary(identity, runs, gates, links), encoding="utf-8")
    policy = json.loads((root / ".github/release-policy.json").read_text(encoding="utf-8"))
    facts = {"pilot": options.pilot_summary}
    checks = release_checks(identity, host, runs, policy, extra, facts)
    release = {"schema_version": 1, "version": identity["version"], "tag": identity["tag"],
               "candidate_sha": identity["sha"], "profile": identity["profile"],
               "distribution": "source-only", "checks": checks, "assets": []}
    write_json(work / "release-evidence.json", release)
    shutil.copyfile(work / "release-evidence.json", evidence / "release-evidence.json")
    checked(tool(root, "check_release.py", "--root", ".", "--tag", identity["tag"], "--candidate-sha", identity["sha"],
                 "--evidence", str(work / "release-evidence.json"),
                 "--excel-evidence", str(evidence / "host.json"),
                 "--output", str(evidence / "release-integrity-pretag.json")), root, "pre-tag check_release.py")
    print(f"PREPARED: {work / 'release-evidence.json'}; pre-tag integrity PASS. Review every file under "
          f"{evidence} for private data before running bundle --redaction-reviewed.")
    return 0


def bundle_spec(identity: dict[str, Any], evidence: Path) -> dict[str, Any]:
    records = []
    for path in sorted(evidence.iterdir()):
        if not path.is_file():
            continue
        slug = "".join(char if char.isalnum() else "-" for char in path.stem.lower()).strip("-")
        if path.name not in ROLES and (path.name in EXCEL_FILES or path.suffix == ".log"):
            slug = f"excel-{slug}"
        record_id, role = ROLES.get(path.name, (slug, slug))
        records.append({"id": record_id, "role": role, "visibility": "public-file", "path": path.name,
                        "public_name": path.name, "redaction_reviewed": True})
    roles = [record["role"] for record in records]
    require(len(roles) == len(set(roles)), "two evidence files map to the same bundle role")
    return {"schema_version": 1, "repository": identity["repository"], "version": identity["version"],
            "tag": identity["tag"], "candidate_sha": identity["sha"], "records": records}


def bundle(options: argparse.Namespace) -> int:
    root, work = options.root.resolve(), options.work_dir.resolve()
    require(options.redaction_reviewed, "review the evidence files for private data, then pass --redaction-reviewed")
    identity = candidate_identity(root, options.tag, options.candidate_sha)
    require(identity["template"], "the certification bundle applies to the canonical template only")
    evidence = work / "evidence"
    require((work / "release-evidence.json").is_file(), f"run prepare first: {work} has no release-evidence.json")
    complete = work / "bundle-complete.json"
    complete.unlink(missing_ok=True)
    checked(tool(root, "check_release.py", "--root", ".", "--tag", identity["tag"], "--candidate-sha", identity["sha"],
                 "--evidence", str(work / "release-evidence.json"),
                 "--excel-evidence", str(evidence / "host.json"), "--require-tag-ref",
                 "--output", str(evidence / "release-integrity-posttag.json")), root, "post-tag check_release.py")
    write_json(work / "certification-input.json", bundle_spec(identity, evidence))
    outputs = []
    for name in ("out", "out-rebuild"):
        target = work / name
        if target.exists():
            shutil.rmtree(target)
        checked(tool(root, "release_certification.py", "--root", ".", "--build", str(work / "certification-input.json"),
                     "--evidence-dir", str(evidence), "--output-dir", str(target)), root, f"bundle build into {name}")
        outputs.append(target / f"certification-{identity['tag']}.zip")
    out = work / "out"
    checked(tool(root, "release_certification.py", "--verify-bundle", str(outputs[0]),
                 "--manifest", str(out / f"certification-{identity['tag']}.manifest.json"),
                 "--digest", str(out / f"certification-{identity['tag']}.sha256")), root, "bundle verification")
    require(outputs[0].read_bytes() == outputs[1].read_bytes(), "the bundle rebuild is not byte-identical")
    # Written last: its presence means build, verification and rebuild all passed.
    write_json(complete, {"tag": identity["tag"], "candidate_sha": identity["sha"],
                          "bundle_sha256": hashlib.sha256(outputs[0].read_bytes()).hexdigest()})
    print(f"BUNDLED: {outputs[0]} SHA-256 {hashlib.sha256(outputs[0].read_bytes()).hexdigest()}; "
          "verified and deterministic. Sign it next.")
    return 0


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    commands = result.add_subparsers(dest="command", required=True)
    for name in ("prepare", "bundle"):
        command = commands.add_parser(name)
        command.add_argument("--root", type=Path, default=Path("."))
        command.add_argument("--tag", required=True)
        command.add_argument("--candidate-sha", required=True)
        command.add_argument("--work-dir", type=Path, required=True)
    prepare_command = commands.choices["prepare"]
    prepare_command.add_argument("--excel-bundle", type=Path, required=True)
    prepare_command.add_argument("--wiki-browser-reviewed", action="store_true")
    prepare_command.add_argument("--pilot-record", type=Path)
    prepare_command.add_argument("--pilot-summary", default="")
    prepare_command.add_argument("--check", action="append", default=[])
    commands.choices["bundle"].add_argument("--redaction-reviewed", action="store_true")
    return result


def main(arguments: list[str] | None = None) -> int:
    options = parser().parse_args(arguments)
    try:
        return prepare(options) if options.command == "prepare" else bundle(options)
    except EvidenceError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
