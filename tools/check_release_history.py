#!/usr/bin/env python3
"""Validate canonical-template release history against the squash-by-default policy."""

from __future__ import annotations

import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from typing import Any

from _gatelib import git_text, parse_report_args as parse_args, run_gate

POLICY_PATH = ".github/release-history-policy.json"
PROFILE_PATH = ".github/repository-profile.json"
TOOL_NAME = "Release history convention"
SUPPORTED_FINDINGS = frozenset({"merge-commit", "duplicate-subject"})
SHA_RE = re.compile(r"[0-9a-f]{40}")
TAG_RE = re.compile(r"v(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:[-+][0-9A-Za-z.-]+)?")
REVIEW_RE = re.compile(
    r"https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+/(?:issues|pull)/[1-9]\d*"
)


class HistoryError(RuntimeError):
    """The history gate could not evaluate the requested repository."""


def _git_output(root: Path, *arguments: str) -> str:
    completed = git_text(root, *arguments)
    if completed.returncode:
        detail = completed.stderr.strip() or "git command failed"
        raise HistoryError(detail)
    return completed.stdout.strip()


def _candidate_json(root: Path, candidate: str, path: str) -> dict[str, Any]:
    raw = _git_output(root, "show", f"{candidate}:{path}")
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as error:
        raise HistoryError(f"{path}: invalid JSON: {error.msg}") from error
    if not isinstance(value, dict):
        raise HistoryError(f"{path}: root must be an object")
    return value


def _nonempty(value: object) -> bool:
    return isinstance(value, str) and bool(value.strip())


def _validate_exception(value: object) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise HistoryError("release-history exception must be an object")
    expected = {"base_tag", "commit", "findings", "review_ref", "reason"}
    if set(value) != expected:
        raise HistoryError("release-history exception has invalid fields")
    base_tag = value["base_tag"]
    commit = value["commit"]
    findings = value["findings"]
    review_ref = value["review_ref"]
    reason = value["reason"]
    if not isinstance(base_tag, str) or TAG_RE.fullmatch(base_tag) is None:
        raise HistoryError("release-history exception base_tag is invalid")
    if not isinstance(commit, str) or SHA_RE.fullmatch(commit) is None:
        raise HistoryError("release-history exception commit is invalid")
    if (
        not isinstance(findings, list)
        or not findings
        or any(item not in SUPPORTED_FINDINGS for item in findings)
        or len(findings) != len(set(findings))
    ):
        raise HistoryError("release-history exception findings are invalid")
    if not isinstance(review_ref, str) or REVIEW_RE.fullmatch(review_ref) is None:
        raise HistoryError("release-history exception review_ref must be a GitHub issue or PR URL")
    if not _nonempty(reason):
        raise HistoryError("release-history exception reason must be non-empty")
    return value


def _validate_historical(value: object) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise HistoryError("historical release-history record must be an object")
    expected = {"release", "commit", "pull_request", "reason"}
    if set(value) != expected:
        raise HistoryError("historical release-history record has invalid fields")
    release = value["release"]
    commit = value["commit"]
    pull_request = value["pull_request"]
    reason = value["reason"]
    if not isinstance(release, str) or TAG_RE.fullmatch(release) is None:
        raise HistoryError("historical release tag is invalid")
    if not isinstance(commit, str) or SHA_RE.fullmatch(commit) is None:
        raise HistoryError("historical release commit is invalid")
    if not isinstance(pull_request, int) or isinstance(pull_request, bool) or pull_request <= 0:
        raise HistoryError("historical pull_request must be a positive integer")
    if not _nonempty(reason):
        raise HistoryError("historical release reason must be non-empty")
    return value


def _load_policy(root: Path, candidate: str) -> dict[str, Any]:
    policy = _candidate_json(root, candidate, POLICY_PATH)
    expected = {"schema_version", "rules", "exceptions", "historical_records"}
    if set(policy) != expected:
        raise HistoryError(f"{POLICY_PATH}: invalid top-level fields")
    if policy.get("schema_version") != 1:
        raise HistoryError(f"{POLICY_PATH}: unsupported schema version")
    rules = policy.get("rules")
    if not isinstance(rules, dict) or rules != {
        "merge_commits": "block-unless-excepted",
        "duplicate_subjects": "block-unless-excepted",
    }:
        raise HistoryError(f"{POLICY_PATH}: unsupported rule set")
    exceptions = policy.get("exceptions")
    historical = policy.get("historical_records")
    if not isinstance(exceptions, list):
        raise HistoryError(f"{POLICY_PATH}: exceptions must be an array")
    if not isinstance(historical, list):
        raise HistoryError(f"{POLICY_PATH}: historical_records must be an array")
    policy["exceptions"] = [_validate_exception(item) for item in exceptions]
    policy["historical_records"] = [_validate_historical(item) for item in historical]
    return policy


def _template_mode(root: Path, candidate: str) -> bool:
    profile = _candidate_json(root, candidate, PROFILE_PATH)
    return profile.get("mode") == "template" and profile.get("profile") is None


def _previous_release_tag(root: Path, candidate: str) -> str:
    parent = _git_output(root, "rev-parse", f"{candidate}^")
    tags = _git_output(
        root,
        "tag",
        "--merged",
        parent,
        "--list",
        "v[0-9]*",
        "--sort=-version:refname",
    ).splitlines()
    for tag in tags:
        if TAG_RE.fullmatch(tag):
            return tag
    raise HistoryError("no previous semantic release tag is reachable from candidate parent")


def _history_rows(root: Path, base_tag: str, candidate: str) -> list[dict[str, Any]]:
    raw = _git_output(
        root,
        "log",
        "--reverse",
        "--format=%H%x1f%P%x1f%s%x1e",
        f"{base_tag}..{candidate}",
    )
    rows: list[dict[str, Any]] = []
    for record in raw.split("\x1e"):
        if not record.strip():
            continue
        fields = record.strip().split("\x1f")
        if len(fields) != 3:
            raise HistoryError("could not parse Git history record")
        sha, parents, subject = fields
        rows.append(
            {
                "commit": sha,
                "parents": parents.split() if parents else [],
                "subject": subject.strip(),
            }
        )
    return rows


def _active_exceptions(policy: dict[str, Any], base_tag: str) -> dict[str, set[str]]:
    active: dict[str, set[str]] = {}
    for item in policy["exceptions"]:
        if item["base_tag"] != base_tag:
            continue
        active.setdefault(item["commit"], set()).update(item["findings"])
    return active


def _observed_conditions(rows: list[dict[str, Any]]) -> dict[str, set[str]]:
    observed: dict[str, set[str]] = {}
    subjects: dict[str, str] = {}
    for row in rows:
        commit = row["commit"]
        if len(row["parents"]) > 1:
            observed.setdefault(commit, set()).add("merge-commit")
        normalized = row["subject"].strip().casefold()
        if not normalized:
            continue
        first = subjects.get(normalized)
        if first is None:
            subjects[normalized] = commit
        else:
            observed.setdefault(commit, set()).add("duplicate-subject")
    return observed


def _findings(
    rows: list[dict[str, Any]],
    active: dict[str, set[str]],
) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    observed = _observed_conditions(rows)
    findings: list[dict[str, str]] = []
    usage: list[dict[str, Any]] = []
    row_by_sha = {row["commit"]: row for row in rows}
    for commit, conditions in sorted(observed.items()):
        allowed = active.get(commit, set())
        for condition in sorted(conditions):
            if condition in allowed:
                usage.append({"commit": commit, "finding": condition, "status": "used"})
                continue
            row = row_by_sha[commit]
            if condition == "merge-commit":
                message = "multi-parent commit is not covered by a reviewed history exception"
                code = "unapproved-merge-commit"
            else:
                message = f"duplicate commit subject is not excepted: {row['subject']!r}"
                code = "duplicate-commit-subject"
            findings.append({"code": code, "path": commit, "message": message})
    for commit, allowed in sorted(active.items()):
        actual = observed.get(commit, set())
        if commit not in row_by_sha:
            findings.append(
                {
                    "code": "stale-history-exception",
                    "path": commit,
                    "message": "active exception commit is not in the inspected release range",
                }
            )
            continue
        for condition in sorted(allowed - actual):
            findings.append(
                {
                    "code": "overbroad-history-exception",
                    "path": commit,
                    "message": f"exception permits {condition!r} but that condition is not present",
                }
            )
    return findings, usage


def build_report(root: Path) -> dict[str, Any]:
    root = root.resolve()
    candidate = _git_output(root, "rev-parse", "HEAD")
    if not _template_mode(root, candidate):
        return {
            "schema_version": 1,
            "tool": TOOL_NAME,
            "status": "pass",
            "applicable": False,
            "candidate_sha": candidate,
            "previous_tag": None,
            "commits": [],
            "exception_usage": [],
            "historical_records": [],
            "findings": [],
            "scope_note": "Template-maintainer policy is not applied to generated repositories.",
        }
    policy = _load_policy(root, candidate)
    base_tag = _previous_release_tag(root, candidate)
    rows = _history_rows(root, base_tag, candidate)
    active = _active_exceptions(policy, base_tag)
    findings, usage = _findings(rows, active)
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "applicable": True,
        "candidate_sha": candidate,
        "previous_tag": base_tag,
        "commits": rows,
        "exception_usage": usage,
        "historical_records": policy["historical_records"],
        "findings": findings,
        "scope_note": (
            "Inspects only the canonical template's previous-release-to-candidate Git range; "
            "generated repositories do not inherit this history policy."
        ),
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "## Release history convention",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Applicable:** {'yes' if report['applicable'] else 'no'}",
        f"- **Candidate:** `{report['candidate_sha']}`",
        f"- **Previous release:** `{report['previous_tag'] or 'not applicable'}`",
        f"- **Commits inspected:** {len(report['commits'])}",
        f"- **Exceptions used:** {len(report['exception_usage'])}",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["findings"]:
        lines.extend(["", "### Findings", ""])
        for item in report["findings"]:
            lines.append(f"- `{item['code']}` at `{item['path']}` — {item['message']}")
    if report["exception_usage"]:
        lines.extend(["", "### Reviewed exceptions used", ""])
        for item in report["exception_usage"]:
            lines.append(f"- `{item['commit']}` — `{item['finding']}`")
    lines.extend(["", f"> {report['scope_note']}", ""])
    return "\n".join(lines)


def _command(root: Path, *arguments: str) -> str:
    completed = subprocess.run(
        ["git", "-C", str(root), "-c", "user.name=History Fixture",
         "-c", "user.email=history@example.invalid", *arguments],
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def _base_fixture(root: Path) -> None:
    (root / ".github").mkdir(parents=True)
    (root / PROFILE_PATH).write_text(
        json.dumps({"mode": "template", "profile": None}) + "\n", encoding="utf-8"
    )
    policy = {
        "schema_version": 1,
        "rules": {
            "merge_commits": "block-unless-excepted",
            "duplicate_subjects": "block-unless-excepted",
        },
        "exceptions": [],
        "historical_records": [],
    }
    (root / POLICY_PATH).write_text(json.dumps(policy, indent=2) + "\n", encoding="utf-8")
    (root / "fixture.txt").write_text("base\n", encoding="utf-8")
    _command(root, "init", "-b", "main")
    _command(root, "add", "--all")
    _command(root, "commit", "-m", "Base release")
    _command(root, "tag", "-a", "v1.0.0", "-m", "v1.0.0")


def _commit_file(root: Path, text: str, subject: str) -> str:
    (root / "fixture.txt").write_text(text + "\n", encoding="utf-8")
    _command(root, "add", "fixture.txt")
    _command(root, "commit", "-m", subject)
    return _command(root, "rev-parse", "HEAD")


def _write_exception(root: Path, merge_sha: str) -> None:
    policy = json.loads((root / POLICY_PATH).read_text(encoding="utf-8"))
    policy["exceptions"] = [
        {
            "base_tag": "v1.0.0",
            "commit": merge_sha,
            "findings": ["merge-commit"],
            "review_ref": "https://github.com/example/repo/pull/1",
            "reason": "Fixture-reviewed ancestry-preserving merge.",
        }
    ]
    (root / POLICY_PATH).write_text(json.dumps(policy, indent=2) + "\n", encoding="utf-8")
    _command(root, "add", POLICY_PATH)
    _command(root, "commit", "-m", "Register reviewed merge exception")


def _merge_fixture(root: Path, approved: bool) -> dict[str, Any]:
    _base_fixture(root)
    _command(root, "switch", "-c", "feature")
    _commit_file(root, "feature", "Feature change")
    _command(root, "switch", "main")
    _commit_file(root, "main", "Mainline preparation")
    _command(root, "merge", "--no-ff", "feature", "-m", "Reviewed feature merge")
    merge_sha = _command(root, "rev-parse", "HEAD")
    if approved:
        _write_exception(root, merge_sha)
    return build_report(root)


def run_self_test() -> int:
    results: list[tuple[str, bool]] = []
    with tempfile.TemporaryDirectory(prefix="release-history-") as temporary:
        area = Path(temporary)

        compliant = area / "compliant"
        compliant.mkdir()
        _base_fixture(compliant)
        _commit_file(compliant, "one", "Focused squash result")
        report = build_report(compliant)
        results.append(("compliant-squash-history", report["status"] == "pass"))

        approved = area / "approved"
        approved.mkdir()
        report = _merge_fixture(approved, approved=True)
        results.append(
            (
                "approved-merge-exception",
                report["status"] == "pass" and len(report["exception_usage"]) == 1,
            )
        )

        unapproved = area / "unapproved"
        unapproved.mkdir()
        report = _merge_fixture(unapproved, approved=False)
        codes = {item["code"] for item in report["findings"]}
        results.append(("unapproved-merge-commit", "unapproved-merge-commit" in codes))

        duplicate = area / "duplicate"
        duplicate.mkdir()
        _base_fixture(duplicate)
        _commit_file(duplicate, "one", "Repeated subject")
        _commit_file(duplicate, "two", "Repeated subject")
        report = build_report(duplicate)
        codes = {item["code"] for item in report["findings"]}
        results.append(("duplicate-subject-history", "duplicate-commit-subject" in codes))

    failures = [name for name, passed in results if not passed]
    for name, passed in results:
        print(f"[{'PASS' if passed else 'FAIL'}] {name}")
    if failures:
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(f"SELF-TEST PASS: {len(results)} release-history fixtures passed.")
    return 0


def main(argv: list[str] | None = None) -> int:
    options = parse_args(
        sys.argv[1:] if argv is None else argv,
        description=__doc__,
    )
    return run_gate(
        options,
        build=lambda: build_report(options.root),
        markdown=markdown_report,
        errors=(HistoryError, OSError, UnicodeError, json.JSONDecodeError, subprocess.SubprocessError),
        self_test=run_self_test,
        self_test_error_prefix="SELF-TEST ERROR",
    )


if __name__ == "__main__":
    raise SystemExit(main())
