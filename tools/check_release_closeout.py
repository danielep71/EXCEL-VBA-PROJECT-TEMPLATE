#!/usr/bin/env python3
"""Validate post-release closeout from an exact Git candidate and captured provider facts."""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from _gatelib import git_text, run_gate

PROFILE_PATH = ".github/repository-profile.json"
RELEASE_POLICY_PATH = ".github/release-policy.json"
VERSION_PATH = "VERSION"
CHANGELOG_PATH = "CHANGELOG.md"
DEFAULT_WORKFLOW = ".github/workflows/static-checks.yml"
VERSION_PATTERN = re.compile(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")


class CloseoutError(RuntimeError):
    """Raised when captured closeout evidence cannot be evaluated."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CloseoutError(message)


def unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=unique_object)
    except FileNotFoundError as error:
        raise CloseoutError(f"evidence file is missing: {path}") from error
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise CloseoutError(f"cannot read {path}: {error}") from error


def git_show(root: Path, candidate: str, relative: str) -> str:
    completed = git_text(root, "show", f"{candidate}:{relative}")
    if completed.returncode:
        raise CloseoutError(
            f"cannot read {relative} from {candidate}: "
            f"{completed.stderr.strip() or 'git show failed'}"
        )
    return completed.stdout


def candidate_json(root: Path, candidate: str, relative: str) -> dict[str, Any]:
    try:
        value = json.loads(git_show(root, candidate, relative), object_pairs_hook=unique_object)
    except json.JSONDecodeError as error:
        raise CloseoutError(f"{relative} at {candidate} is invalid JSON: {error}") from error
    require(isinstance(value, dict), f"{relative} at {candidate} must be an object")
    return value


def finding(category: str, control: str, message: str) -> dict[str, str]:
    return {"category": category, "control": control, "message": message}


def workflow_runs(value: Any) -> tuple[list[dict[str, Any]], int | None]:
    pages = value if isinstance(value, list) else [value]
    rows: list[dict[str, Any]] = []
    total: int | None = None
    for page in pages:
        require(isinstance(page, dict), "workflow_runs must be an object or slurped objects")
        current = page.get("workflow_runs")
        require(isinstance(current, list), "workflow_runs payload lacks workflow_runs")
        rows.extend(row for row in current if isinstance(row, dict))
        count = page.get("total_count")
        if isinstance(count, int):
            total = count if total is None else max(total, count)
    return rows, total


def milestone_items(value: Any) -> list[dict[str, Any]]:
    require(isinstance(value, list), "milestone_items must be an array or slurped arrays")
    if value and all(isinstance(page, list) for page in value):
        return [row for page in value for row in page if isinstance(row, dict)]
    return [row for row in value if isinstance(row, dict)]


def source_identity(
    root: Path,
    candidate: str,
    tag: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    version = git_show(root, candidate, VERSION_PATH).strip()
    if VERSION_PATTERN.fullmatch(version) is None:
        findings.append(
            finding("deterministic", "source-version", f"candidate VERSION is invalid: {version!r}")
        )
    if tag != f"v{version}":
        findings.append(
            finding(
                "deterministic",
                "source-version",
                f"tag {tag!r} disagrees with candidate VERSION {version!r}",
            )
        )

    profile_data = candidate_json(root, candidate, PROFILE_PATH)
    repository = profile_data.get("repository")
    require(isinstance(repository, str) and "/" in repository, "candidate repository is invalid")
    mode = profile_data.get("mode")
    profile = "template" if mode == "template" else profile_data.get("profile")
    require(isinstance(profile, str) and profile, "candidate release profile is missing")

    policy = candidate_json(root, candidate, RELEASE_POLICY_PATH)
    profiles = policy.get("profiles")
    require(isinstance(profiles, dict), "release policy profiles are missing")
    profile_policy = profiles.get(profile)
    require(isinstance(profile_policy, dict), f"release policy profile {profile!r} is missing")
    allowed = profile_policy.get("allowed_asset_globs")
    require(
        isinstance(allowed, list) and all(isinstance(item, str) and item for item in allowed),
        "allowed_asset_globs must be an array of non-empty strings",
    )

    changelog = git_show(root, candidate, CHANGELOG_PATH)
    heading = re.search(
        rf"^## \[{re.escape(version)}\] - (?P<date>\d{{4}}-\d{{2}}-\d{{2}})\s*$",
        changelog,
        re.MULTILINE,
    )
    if heading is None:
        findings.append(
            finding(
                "deterministic",
                "changelog",
                f"candidate changelog has no released section for {version}",
            )
        )

    release_link = re.search(
        rf"^\[{re.escape(version)}\]:\s+https://github\.com/{re.escape(repository)}/compare/"
        rf"(?P<base>v[^\s]+)\.\.\.{re.escape(tag)}\s*$",
        changelog,
        re.MULTILINE,
    )
    previous = release_link.group("base") if release_link else None
    if release_link is None:
        findings.append(
            finding(
                "deterministic",
                "comparison-link",
                f"candidate comparison link for {version} does not target {tag}",
            )
        )
    if re.search(
        rf"^\[Unreleased\]:\s+https://github\.com/{re.escape(repository)}/compare/"
        rf"{re.escape(tag)}\.\.\.HEAD\s*$",
        changelog,
        re.MULTILINE,
    ) is None:
        findings.append(
            finding(
                "deterministic",
                "comparison-link",
                f"Unreleased comparison link does not start from {tag}",
            )
        )

    return {
        "repository": repository,
        "mode": mode,
        "profile": profile,
        "version": version,
        "changelog_date": heading.group("date") if heading else None,
        "previous_tag": previous,
        "allowed_asset_globs": allowed,
    }


def check_tag(
    snapshot: dict[str, Any],
    tag: str,
    candidate: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    ref = snapshot.get("tag_ref")
    require(isinstance(ref, dict), "snapshot.tag_ref must be an object")
    if ref.get("ref") != f"refs/tags/{tag}":
        findings.append(finding("deterministic", "tag", "provider tag ref does not match release tag"))
    obj = ref.get("object")
    require(isinstance(obj, dict), "snapshot.tag_ref lacks object")
    ref_type = obj.get("type")
    target: str | None = None
    tag_object_sha: str | None = None
    if ref_type == "commit":
        target = obj.get("sha") if isinstance(obj.get("sha"), str) else None
        findings.append(finding("deterministic", "tag", f"{tag} is lightweight; annotated tag required"))
    elif ref_type == "tag":
        tag_object_sha = obj.get("sha") if isinstance(obj.get("sha"), str) else None
        annotated = snapshot.get("tag_object")
        require(isinstance(annotated, dict), "annotated tag requires snapshot.tag_object")
        if annotated.get("sha") != tag_object_sha:
            findings.append(finding("deterministic", "tag", "annotated tag object SHA mismatches tag ref"))
        if annotated.get("tag") != tag:
            findings.append(finding("deterministic", "tag", "annotated tag object names another tag"))
        annotated_target = annotated.get("object")
        require(isinstance(annotated_target, dict), "snapshot.tag_object lacks target object")
        if annotated_target.get("type") != "commit":
            findings.append(finding("deterministic", "tag", "annotated tag does not target a commit"))
        target = (
            annotated_target.get("sha")
            if isinstance(annotated_target.get("sha"), str)
            else None
        )
    else:
        findings.append(finding("deterministic", "tag", f"unsupported tag object type: {ref_type!r}"))
    if target != candidate:
        findings.append(
            finding(
                "deterministic",
                "tag",
                f"tag resolves to {target!r}; expected certified candidate {candidate}",
            )
        )
    return {"ref_type": ref_type, "tag_object_sha": tag_object_sha, "target_sha": target}


def check_workflow(
    snapshot: dict[str, Any],
    path: str,
    tag: str,
    candidate: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    rows, total = workflow_runs(snapshot.get("workflow_runs"))
    if total is not None and total > len(rows):
        findings.append(
            finding(
                "deterministic",
                "tag-ci",
                f"workflow capture is incomplete: {len(rows)} of {total} runs",
            )
        )
    matches = [
        row
        for row in rows
        if row.get("path") == path
        and row.get("head_branch") == tag
        and row.get("head_sha") == candidate
        and row.get("event") == "push"
    ]
    if not matches:
        findings.append(
            finding("deterministic", "tag-ci", f"no tag-triggered {path} run matches candidate")
        )
        return {"matches": 0, "run_id": None, "status": None, "conclusion": None}
    selected = max(
        matches,
        key=lambda row: (
            int(row.get("run_attempt") or 0),
            int(row.get("run_number") or 0),
            int(row.get("id") or 0),
        ),
    )
    if selected.get("status") != "completed" or selected.get("conclusion") != "success":
        findings.append(
            finding(
                "deterministic",
                "tag-ci",
                f"tag CI is {selected.get('status')!r}/{selected.get('conclusion')!r}",
            )
        )
    return {
        "matches": len(matches),
        "run_id": selected.get("id"),
        "status": selected.get("status"),
        "conclusion": selected.get("conclusion"),
    }


def check_release(
    snapshot: dict[str, Any],
    identity: dict[str, Any],
    expect_prerelease: bool,
    expect_latest: bool,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    release = snapshot.get("release")
    latest = snapshot.get("latest_release")
    require(isinstance(release, dict), "snapshot.release must be an object")
    require(isinstance(latest, dict), "snapshot.latest_release must be an object")
    tag = f"v{identity['version']}"
    if release.get("tag_name") != tag:
        findings.append(finding("deterministic", "release", "GitHub Release uses another tag"))
    if release.get("draft") is not False or not release.get("published_at"):
        findings.append(finding("deterministic", "release", "GitHub Release is draft/unpublished"))
    if bool(release.get("prerelease")) != expect_prerelease:
        findings.append(
            finding(
                "deterministic",
                "release",
                f"prerelease flag is {bool(release.get('prerelease'))}; expected {expect_prerelease}",
            )
        )
    latest_matches = release.get("id") == latest.get("id")
    if latest_matches != expect_latest:
        findings.append(
            finding(
                "deterministic",
                "release",
                f"latest-release expectation is {expect_latest}; observed {latest_matches}",
            )
        )

    assets = release.get("assets")
    require(isinstance(assets, list), "snapshot.release.assets must be an array")
    names = [
        asset["name"]
        for asset in assets
        if isinstance(asset, dict) and isinstance(asset.get("name"), str)
    ]
    if len(names) != len(assets) or len(names) != len(set(names)):
        findings.append(finding("deterministic", "assets", "uploaded asset names are invalid/duplicate"))
    allowed = identity["allowed_asset_globs"]
    unexpected = [
        name
        for name in names
        if not any(fnmatch.fnmatchcase(name, pattern) for pattern in allowed)
    ]
    if not allowed and names:
        findings.append(
            finding(
                "deterministic",
                "assets",
                "source-only release profile contains uploaded assets",
            )
        )
    if unexpected:
        findings.append(
            finding(
                "deterministic",
                "assets",
                "uploaded assets outside allowed_asset_globs: " + ", ".join(sorted(unexpected)),
            )
        )

    zip_exposed = isinstance(release.get("zipball_url"), str) and bool(release["zipball_url"])
    tar_exposed = isinstance(release.get("tarball_url"), str) and bool(release["tarball_url"])
    if not zip_exposed or not tar_exposed:
        findings.append(
            finding(
                "deterministic",
                "source-archives",
                "GitHub-generated source archive URLs are incomplete",
            )
        )
    archive_observation = snapshot.get("source_archives")
    require(isinstance(archive_observation, dict), "snapshot.source_archives must be an object")
    for kind in ("zip", "tar"):
        if archive_observation.get(kind) != "pass":
            findings.append(
                finding(
                    "observation",
                    "source-archives",
                    f"{kind} source archive retrieval is {archive_observation.get(kind)!r}",
                )
            )
    return {
        "release_id": release.get("id"),
        "published_at": release.get("published_at"),
        "draft": release.get("draft"),
        "prerelease": bool(release.get("prerelease")),
        "latest_matches": latest_matches,
        "asset_mode": "source-only" if not allowed else "binary-capable",
        "uploaded_assets": sorted(names),
        "unexpected_assets": sorted(unexpected),
        "source_zip_exposed": zip_exposed,
        "source_tar_exposed": tar_exposed,
    }


def check_compare(
    snapshot: dict[str, Any],
    identity: dict[str, Any],
    candidate: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    previous = identity.get("previous_tag")
    compare = snapshot.get("compare")
    require(isinstance(compare, dict), "snapshot.compare must be an object")
    if previous is None:
        return {"previous_tag": None, "status": None, "candidate_seen": False}
    expected = f"/compare/{previous}...v{identity['version']}"
    if not isinstance(compare.get("html_url"), str) or not compare["html_url"].endswith(expected):
        findings.append(
            finding("deterministic", "comparison-link", "provider comparison range mismatches changelog")
        )
    status = compare.get("status")
    if status not in {"ahead", "identical"}:
        findings.append(
            finding("deterministic", "comparison-link", f"provider comparison status is {status!r}")
        )
    commits = compare.get("commits")
    require(isinstance(commits, list), "snapshot.compare.commits must be an array")
    candidate_seen = any(
        isinstance(row, dict) and row.get("sha") == candidate for row in commits
    )
    if status == "ahead" and not candidate_seen:
        findings.append(
            finding(
                "deterministic",
                "comparison-link",
                "provider comparison does not contain the certified candidate",
            )
        )
    return {"previous_tag": previous, "status": status, "candidate_seen": candidate_seen}


def check_milestone(
    snapshot: dict[str, Any],
    milestone_number: int,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    milestone = snapshot.get("milestone")
    require(isinstance(milestone, dict), "snapshot.milestone must be an object")
    items = milestone_items(snapshot.get("milestone_items"))
    if milestone.get("number") != milestone_number:
        findings.append(finding("deterministic", "milestone", "milestone number mismatches input"))
    if any(
        not isinstance(row.get("milestone"), dict)
        or row["milestone"].get("number") != milestone_number
        for row in items
    ):
        findings.append(
            finding("deterministic", "milestone", "captured item belongs to another milestone")
        )
    open_items = sorted(
        row["number"]
        for row in items
        if row.get("state") == "open" and isinstance(row.get("number"), int)
    )
    closed = [row for row in items if row.get("state") == "closed"]
    if any(row.get("state") not in {"open", "closed"} for row in items):
        findings.append(finding("deterministic", "milestone", "unsupported milestone item state"))
    if open_items:
        findings.append(
            finding(
                "deterministic",
                "milestone",
                "release milestone has open items: " + ", ".join(f"#{value}" for value in open_items),
            )
        )
    if milestone.get("open_issues") != len(open_items) or milestone.get("closed_issues") != len(closed):
        findings.append(
            finding(
                "deterministic",
                "milestone",
                "provider milestone counters disagree with actual captured membership/state",
            )
        )
    return {
        "number": milestone_number,
        "title": milestone.get("title"),
        "state": milestone.get("state"),
        "open_items": open_items,
        "closed_items": len(closed),
        "membership_count": len(items),
    }


def check_wiki(
    snapshot: dict[str, Any],
    identity: dict[str, Any],
    candidate: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    if identity["mode"] != "template":
        return {"applicable": False, "status": "not-applicable", "source_sha": None}
    wiki = snapshot.get("wiki")
    require(isinstance(wiki, dict), "template closeout requires snapshot.wiki")
    if wiki.get("status") != "pass":
        findings.append(
            finding("observation", "wiki", f"authoritative Wiki report is {wiki.get('status')!r}")
        )
    if wiki.get("source_sha") != candidate:
        findings.append(
            finding(
                "observation",
                "wiki",
                f"Wiki source SHA is {wiki.get('source_sha')!r}; expected {candidate}",
            )
        )
    browser = snapshot.get("ui_observations")
    require(isinstance(browser, dict), "snapshot.ui_observations must be an object")
    if browser.get("wiki_browser_review") != "pass":
        findings.append(
            finding(
                "observation",
                "wiki-browser-review",
                f"Wiki browser review is {browser.get('wiki_browser_review')!r}",
            )
        )
    return {
        "applicable": True,
        "status": wiki.get("status"),
        "source_sha": wiki.get("source_sha"),
        "browser_review": browser.get("wiki_browser_review"),
    }


def build_report(options: argparse.Namespace) -> dict[str, Any]:
    root = options.root.resolve()
    candidate = options.candidate_sha
    require(re.fullmatch(r"[0-9a-f]{40}", candidate) is not None, "candidate SHA must be full lowercase hex")
    require(
        git_text(root, "cat-file", "-e", f"{candidate}^{{commit}}").returncode == 0,
        f"candidate commit is unavailable locally: {candidate}",
    )
    snapshot = load_json(options.snapshot)
    require(isinstance(snapshot, dict), "snapshot root must be an object")
    require(snapshot.get("schema_version") == 1, "unsupported closeout snapshot schema")

    findings: list[dict[str, str]] = []
    identity = source_identity(root, candidate, options.tag, findings)
    if snapshot.get("repository") != identity["repository"]:
        findings.append(
            finding("deterministic", "provider-scope", "snapshot repository mismatches candidate")
        )
    tag_state = check_tag(snapshot, options.tag, candidate, findings)
    workflow = check_workflow(
        snapshot, options.workflow_path, options.tag, candidate, findings
    )
    release = check_release(
        snapshot, identity, options.expect_prerelease, options.expect_latest, findings
    )
    comparison = check_compare(snapshot, identity, candidate, findings)
    milestone = check_milestone(snapshot, options.milestone_number, findings)
    wiki = check_wiki(snapshot, identity, candidate, findings)

    deterministic = [row for row in findings if row["category"] == "deterministic"]
    observations = [row for row in findings if row["category"] == "observation"]
    return {
        "schema_version": 1,
        "gate": "release-closeout",
        "status": "pass" if not findings else "fail",
        "deterministic_status": "pass" if not deterministic else "fail",
        "observation_status": "pass" if not observations else "incomplete",
        "repository": identity["repository"],
        "candidate_sha": candidate,
        "version": identity["version"],
        "tag": options.tag,
        "profile": identity["profile"],
        "snapshot_sha256": hashlib.sha256(options.snapshot.read_bytes()).hexdigest(),
        "source": {
            "version": identity["version"],
            "changelog_date": identity["changelog_date"],
            "previous_tag": identity["previous_tag"],
        },
        "tag_state": tag_state,
        "tag_workflow": workflow,
        "release": release,
        "comparison": comparison,
        "milestone": milestone,
        "wiki": wiki,
        "findings": findings,
        "scope_note": (
            "Read-only closeout validation over one exact Git candidate and captured GitHub REST "
            "facts. Uploaded assets are distinct from GitHub-generated source archives. Wiki "
            "publication bytes remain owned by check_wiki.py; browser and archive retrieval are "
            "explicit observations rather than inferred repository facts."
        ),
    }


def markdown_report(report: dict[str, Any]) -> str:
    controls = [
        (
            "Annotated tag / candidate",
            report["tag_state"]["ref_type"] == "tag"
            and report["tag_state"]["target_sha"] == report["candidate_sha"],
            report["tag_state"]["target_sha"],
        ),
        (
            "Tag-triggered workflow",
            report["tag_workflow"]["conclusion"] == "success",
            f"run {report['tag_workflow']['run_id']}",
        ),
        (
            "GitHub Release/latest",
            not report["release"]["draft"] and report["release"]["latest_matches"],
            f"id {report['release']['release_id']}",
        ),
        (
            "Uploaded assets",
            not report["release"]["unexpected_assets"],
            f"{report['release']['asset_mode']}; {len(report['release']['uploaded_assets'])} uploaded",
        ),
        (
            "Comparison link",
            report["comparison"]["status"] in {"ahead", "identical"},
            f"{report['comparison']['previous_tag']} → {report['tag']}",
        ),
        (
            "Milestone membership",
            not report["milestone"]["open_items"],
            f"{report['milestone']['membership_count']} items",
        ),
    ]
    if report["wiki"]["applicable"]:
        controls.append(
            (
                "Wiki source/read-back",
                report["wiki"]["status"] == "pass"
                and report["wiki"]["source_sha"] == report["candidate_sha"],
                report["wiki"]["source_sha"],
            )
        )
    lines = [
        "# Release closeout",
        "",
        f"- **Status:** {report['status'].upper()}",
        f"- **Deterministic checks:** {report['deterministic_status'].upper()}",
        f"- **Provider/UI observations:** {report['observation_status'].upper()}",
        f"- **Release:** `{report['tag']}` at `{report['candidate_sha']}`",
        f"- **Profile:** `{report['profile']}`",
        f"- **Snapshot SHA-256:** `{report['snapshot_sha256']}`",
        "",
        "| Control | Result | Evidence |",
        "| --- | --- | --- |",
    ]
    lines.extend(
        f"| {name} | {'PASS' if ok else 'FAIL'} | {detail or 'n/a'} |"
        for name, ok, detail in controls
    )
    lines.extend(["", f"**Findings:** {len(report['findings'])}"])
    if report["findings"]:
        lines.append("")
        lines.extend(
            f"- **{row['category']} / {row['control']}:** {row['message']}"
            for row in report["findings"]
        )
    lines.extend(["", report["scope_note"], ""])
    return "\n".join(lines)


def git(root: Path, *arguments: str) -> str:
    completed = subprocess.run(
        [
            "git",
            "-C",
            str(root),
            "-c",
            "user.name=Closeout Fixture",
            "-c",
            "user.email=fixture@example.invalid",
            *arguments,
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def fixture(base: Path) -> tuple[Path, Path, str]:
    root = base / "repo"
    root.mkdir()
    git(root, "init", "-b", "main")
    (root / ".github").mkdir()
    repository = "owner/repo"
    version = "1.2.3"
    tag = f"v{version}"
    (root / VERSION_PATH).write_text(version + "\n", encoding="utf-8")
    (root / CHANGELOG_PATH).write_text(
        "# Changelog\n\n"
        "## [Unreleased]\n\n"
        f"## [{version}] - 2026-09-12\n\n"
        f"[Unreleased]: https://github.com/{repository}/compare/{tag}...HEAD\n"
        f"[{version}]: https://github.com/{repository}/compare/v1.2.2...{tag}\n",
        encoding="utf-8",
    )
    write_json(
        root / PROFILE_PATH,
        {"schema_version": 1, "mode": "template", "profile": None, "repository": repository},
    )
    write_json(
        root / RELEASE_POLICY_PATH,
        {"schema_version": 1, "profiles": {"template": {"allowed_asset_globs": []}}},
    )
    git(root, "add", "--all")
    git(root, "commit", "-m", "Fixture release")
    candidate = git(root, "rev-parse", "HEAD")
    tag_object_sha = "b" * 40
    release = {
        "id": 55,
        "tag_name": tag,
        "draft": False,
        "prerelease": False,
        "published_at": "2026-09-12T10:00:00Z",
        "assets": [],
        "zipball_url": f"https://api.github.com/repos/{repository}/zipball/{tag}",
        "tarball_url": f"https://api.github.com/repos/{repository}/tarball/{tag}",
    }
    snapshot = {
        "schema_version": 1,
        "repository": repository,
        "tag_ref": {
            "ref": f"refs/tags/{tag}",
            "object": {"type": "tag", "sha": tag_object_sha},
        },
        "tag_object": {
            "sha": tag_object_sha,
            "tag": tag,
            "object": {"type": "commit", "sha": candidate},
        },
        "workflow_runs": {
            "total_count": 1,
            "workflow_runs": [
                {
                    "id": 123,
                    "run_number": 7,
                    "run_attempt": 1,
                    "path": DEFAULT_WORKFLOW,
                    "head_branch": tag,
                    "head_sha": candidate,
                    "event": "push",
                    "status": "completed",
                    "conclusion": "success",
                }
            ],
        },
        "release": release,
        "latest_release": dict(release),
        "compare": {
            "html_url": f"https://github.com/{repository}/compare/v1.2.2...{tag}",
            "status": "ahead",
            "commits": [{"sha": candidate}],
        },
        "milestone": {
            "number": 4,
            "title": "v1.2.3",
            "state": "closed",
            "open_issues": 0,
            "closed_issues": 2,
        },
        "milestone_items": [
            {"number": 1, "state": "closed", "milestone": {"number": 4}},
            {"number": 2, "state": "closed", "milestone": {"number": 4}},
        ],
        "wiki": {"status": "pass", "source_sha": candidate, "findings": []},
        "source_archives": {"zip": "pass", "tar": "pass"},
        "ui_observations": {"wiki_browser_review": "pass"},
    }
    snapshot_path = base / "snapshot.json"
    write_json(snapshot_path, snapshot)
    return root, snapshot_path, candidate


def options(root: Path, snapshot: Path, candidate: str) -> argparse.Namespace:
    return argparse.Namespace(
        root=root,
        snapshot=snapshot,
        tag="v1.2.3",
        candidate_sha=candidate,
        milestone_number=4,
        workflow_path=DEFAULT_WORKFLOW,
        expect_prerelease=False,
        expect_latest=True,
        output=None,
        summary=None,
        self_test=False,
    )


def run_self_test() -> int:
    failures: list[str] = []

    def run_case(name: str, mutate: Any = None, expected: str | None = None) -> None:
        with tempfile.TemporaryDirectory(prefix="release-closeout-") as raw:
            root, snapshot_path, candidate = fixture(Path(raw))
            if mutate is not None:
                value = load_json(snapshot_path)
                mutate(value, candidate)
                write_json(snapshot_path, value)
            report = build_report(options(root, snapshot_path, candidate))
            controls = {row["control"] for row in report["findings"]}
            ok = report["status"] == ("pass" if expected is None else "fail")
            ok = ok and (expected is None or expected in controls)
            if not ok:
                failures.append(f"{name}: unexpected report {report['findings']!r}")

    run_case("valid-closeout")
    run_case(
        "lightweight-tag",
        lambda value, candidate: value.update(
            {"tag_ref": {"ref": "refs/tags/v1.2.3", "object": {"type": "commit", "sha": candidate}}}
        ),
        "tag",
    )

    def moved(value: dict[str, Any], _candidate: str) -> None:
        value["tag_object"]["object"]["sha"] = "c" * 40

    run_case("moved-tag", moved, "tag")

    def failed_ci(value: dict[str, Any], _candidate: str) -> None:
        value["workflow_runs"]["workflow_runs"][0]["conclusion"] = "failure"

    run_case("failed-tag-ci", failed_ci, "tag-ci")

    def draft(value: dict[str, Any], _candidate: str) -> None:
        value["release"]["draft"] = True

    run_case("draft-release", draft, "release")

    def prerelease(value: dict[str, Any], _candidate: str) -> None:
        value["release"]["prerelease"] = True

    run_case("unexpected-prerelease", prerelease, "release")

    def latest(value: dict[str, Any], _candidate: str) -> None:
        value["latest_release"]["id"] = 99

    run_case("not-latest", latest, "release")

    def asset(value: dict[str, Any], _candidate: str) -> None:
        value["release"]["assets"] = [{"name": "dist/unexpected.zip"}]

    run_case("unexpected-asset", asset, "assets")

    def archive_url(value: dict[str, Any], _candidate: str) -> None:
        value["release"]["zipball_url"] = None

    run_case("missing-archive-url", archive_url, "source-archives")

    def compare(value: dict[str, Any], _candidate: str) -> None:
        value["compare"]["status"] = "diverged"

    run_case("unresolved-compare", compare, "comparison-link")

    def open_item(value: dict[str, Any], _candidate: str) -> None:
        value["milestone_items"][0]["state"] = "open"
        value["milestone"]["open_issues"] = 1
        value["milestone"]["closed_issues"] = 1

    run_case("open-milestone", open_item, "milestone")

    def stale_count(value: dict[str, Any], _candidate: str) -> None:
        value["milestone"]["closed_issues"] = 99

    run_case("milestone-counter-drift", stale_count, "milestone")

    def wiki(value: dict[str, Any], _candidate: str) -> None:
        value["wiki"]["status"] = "fail"

    run_case("wiki-drift", wiki, "wiki")

    def archive(value: dict[str, Any], _candidate: str) -> None:
        value["source_archives"]["zip"] = "fail"

    run_case("archive-retrieval", archive, "source-archives")

    if failures:
        print("SELF-TEST FAIL:", file=sys.stderr)
        for failure in failures:
            print(f"  - {failure}", file=sys.stderr)
        return 1
    print(
        "SELF-TEST PASS: annotated/lightweight/moved tags, tag CI, release state, "
        "asset policy, comparison resolution, milestone membership/counters, Wiki binding "
        "and source-archive observations passed."
    )
    return 0


def parse_arguments(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--snapshot", type=Path)
    parser.add_argument("--tag")
    parser.add_argument("--candidate-sha")
    parser.add_argument("--milestone-number", type=int)
    parser.add_argument("--workflow-path", default=DEFAULT_WORKFLOW)
    parser.add_argument("--expect-prerelease", action="store_true")
    parser.add_argument("--allow-not-latest", dest="expect_latest", action="store_false")
    parser.set_defaults(expect_latest=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    result = parser.parse_args(argv)
    if not result.self_test:
        missing = [
            name
            for name in ("snapshot", "tag", "candidate_sha", "milestone_number")
            if getattr(result, name) in (None, "")
        ]
        if missing:
            parser.error(
                "operational closeout requires "
                + ", ".join("--" + name.replace("_", "-") for name in missing)
            )
    return result


def main(argv: list[str] | None = None) -> int:
    args = parse_arguments(sys.argv[1:] if argv is None else argv)
    return run_gate(
        args,
        build=lambda: build_report(args),
        markdown=markdown_report,
        errors=(CloseoutError, OSError, ValueError, TypeError),
        self_test=run_self_test,
        self_test_error_prefix="SELF-TEST ERROR",
    )


if __name__ == "__main__":
    raise SystemExit(main())
