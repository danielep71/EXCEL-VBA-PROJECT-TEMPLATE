#!/usr/bin/env python3
"""Sign, publish and close out a canonical release after ``prepare_release_evidence.py prepare``.

Run it on the maintainer's machine, from a clean checkout of the candidate. It
needs ``gh`` signed in with write access, ``ssh-keygen`` and the SSH signing key
registered on GitHub. In order, it:

1. signs the annotated tag, runs the post-tag ``check_release.py`` (which
   verifies the signature), asks for confirmation and pushes the tag;
2. builds, verifies and rebuilds the certification bundle;
3. signs the bundle and verifies that signature;
4. writes the release notes from the CHANGELOG section and the evidence, and
   asks for confirmation;
5. creates the GitHub Release from the existing tag with the four
   certification files;
6. closes the release milestone;
7. dispatches *Release closeout* and waits for its verdict.

Each step is skipped when its result already exists, so a run that stopped can
be started again with the same arguments. A tag that fails its post-tag check is
deleted locally and never pushed. Before a step that publishes something, the
script asks for confirmation; ``--yes`` answers yes to every question.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any

import prepare_release_evidence as evidence

NAMESPACE = "excel-vba-release-certification"
CLOSEOUT = "release-closeout.yml"
POLL_SECONDS = 15
POLL_LIMIT = 80


def confirm(question: str, assume_yes: bool) -> None:
    if assume_yes:
        print(f"{question} yes (--yes)")
        return
    answer = input(f"{question} [y/N] ").strip().lower()
    evidence.require(answer in ("y", "yes"), "stopped at the maintainer's request; nothing further was published")


def succeeds(root: Path, command: list[str]) -> bool:
    return evidence.run(command, root).returncode == 0


def signing_key(value: Path | None) -> Path:
    raw = value or (Path(os.environ["RELEASE_SIGNING_KEY"]) if os.environ.get("RELEASE_SIGNING_KEY") else None)
    evidence.require(raw is not None, "give --signing-key or set RELEASE_SIGNING_KEY to the private signing key")
    assert raw is not None
    evidence.require(raw.suffix != ".pub", "--signing-key must be the private key, not the .pub file")
    evidence.require(raw.is_file(), f"signing key {raw} does not exist")
    return raw


def publish_tag(root: Path, identity: dict[str, Any], work: Path, key: Path, assume_yes: bool) -> None:
    tag, sha = identity["tag"], identity["sha"]
    remote = evidence.checked(["git", "ls-remote", "--tags", "origin", f"refs/tags/{tag}"], root, "git ls-remote")
    if remote.strip():
        print(f"SKIP: {tag} is already published")
        return
    if not succeeds(root, ["git", "rev-parse", "--verify", "--quiet", f"refs/tags/{tag}"]):
        message = f"{identity['repository'].split('/')[1]} {identity['version']}"
        evidence.checked(["git", "-c", "gpg.format=ssh", "-c", f"user.signingkey={key}", "tag", "-s", tag, sha,
                          "-m", message], root, "tag signing")
    check = evidence.tool(root, "check_release.py", "--root", ".", "--tag", tag, "--candidate-sha", sha,
                          "--evidence", str(work / "release-evidence.json"),
                          "--excel-evidence", str(work / "evidence" / "host.json"), "--require-tag-ref")
    if not succeeds(root, check):
        evidence.run(["git", "tag", "-d", tag], root)
        raise evidence.EvidenceError(f"post-tag check failed; the local {tag} was deleted and nothing was pushed")
    confirm(f"Push the signed tag {tag} for {sha}?", assume_yes)
    evidence.checked(["git", "push", "origin", f"refs/tags/{tag}:refs/tags/{tag}"], root, "tag push")


def sign_bundle(root: Path, identity: dict[str, Any], work: Path, key: Path, reviewed: bool) -> Path:
    bundle = work / "out" / f"certification-{identity['tag']}.zip"
    if not bundle.is_file():
        evidence.bundle(argparse.Namespace(root=root, work_dir=work, tag=identity["tag"],
                                           candidate_sha=identity["sha"], redaction_reviewed=reviewed))
    signature = Path(f"{bundle}.sig")
    if not signature.is_file():
        evidence.checked(["ssh-keygen", "-Y", "sign", "-f", str(key), "-n", NAMESPACE, str(bundle)], root,
                         "bundle signing")
    evidence.checked(evidence.tool(root, "release_certification.py", "--root", ".", "--verify-signature",
                                   str(signature), "--bundle", str(bundle), "--candidate-sha", identity["sha"]),
                     root, "bundle signature verification")
    return bundle


def changelog_section(text: str, version: str) -> tuple[str, str | None]:
    match = re.search(rf"^## \[{re.escape(version)}\] - \d{{4}}-\d{{2}}-\d{{2}}\s*$(.*?)(?=^## \[|\Z)",
                      text, re.MULTILINE | re.DOTALL)
    evidence.require(match is not None, f"CHANGELOG.md has no released section for {version}")
    assert match is not None
    body = re.sub(r"^### ", "#### ", match.group(1).strip(), flags=re.MULTILINE)
    link = re.search(rf"^\[{re.escape(version)}\]:\s+(\S+)\s*$", text, re.MULTILINE)
    return body, link.group(1) if link else None


def release_notes(root: Path, identity: dict[str, Any], work: Path, bundle: Path) -> str:
    version, tag = identity["version"], identity["tag"]
    body, comparison = changelog_section((root / "CHANGELOG.md").read_text(encoding="utf-8"), version)
    suites = json.loads((work / "evidence" / "suites.json").read_text(encoding="utf-8"))
    rows = [f"| `{item['entry_point']}` | {item['status']} | {item['detail']} |"
            for item in [*suites.get("suites", []), *suites.get("examples", [])]]
    out = bundle.parent
    assets = []
    for name in (bundle.name, f"{bundle.name}.sig", f"certification-{tag}.manifest.json",
                 f"certification-{tag}.sha256"):
        digest = hashlib.sha256((out / name).read_bytes()).hexdigest()
        assets.append(f"- `{name}`: SHA-256 `{digest}`")
    lines = [f"## {tag}", "", "### 📝 Changes", "", body, "", "### 🔏 Certification", "",
             f"Tag `{tag}` is SSH-signed and points to certified commit `{identity['sha']}`.", "",
             "| Suite or example | Status | Detail |", "| --- | --- | --- |", *rows, "",
             f"Release assets, the certification bundle and its detached SSH signature (namespace "
             f"`{NAMESPACE}`), manifest and digest file:", "", *assets, "", "### 📦 Distribution", "",
             "Source-only: no workbook or add-in binary; the only release assets are the certification "
             "files above. Start a new project with **Use this template**.", ""]
    if comparison:
        lines += [f"[Comparison with the previous release]({comparison})", ""]
    return "\n".join(lines)


def milestone(root: Path, identity: dict[str, Any]) -> dict[str, Any] | None:
    items = evidence.gh_json(root, f"repos/{identity['repository']}/milestones?state=all&per_page=100")
    found = [item for item in items if str(item.get("title", "")).split(" ")[0] == identity["tag"]]
    return found[0] if found else None


def publish_release(root: Path, identity: dict[str, Any], work: Path, bundle: Path, assume_yes: bool) -> None:
    tag, repo = identity["tag"], identity["repository"]
    if succeeds(root, ["gh", "api", f"repos/{repo}/releases/tags/{tag}"]):
        print(f"SKIP: the GitHub Release {tag} already exists")
        return
    notes = work / "release-notes.md"
    if not notes.is_file():
        notes.write_text(release_notes(root, identity, work, bundle), encoding="utf-8")
    found = milestone(root, identity)
    title = str(found["title"]) if found else tag
    confirm(f"Publish the GitHub Release '{title}' with the notes in {notes}?", assume_yes)
    out = bundle.parent
    files = [str(bundle), f"{bundle}.sig", str(out / f"certification-{tag}.manifest.json"),
             str(out / f"certification-{tag}.sha256")]
    evidence.checked(["gh", "release", "create", tag, *files, "--repo", repo, "--verify-tag", "--title", title,
                      "--notes-file", str(notes), "--latest"], root, "GitHub Release creation")


def close_out(root: Path, identity: dict[str, Any], wiki_reviewed: bool) -> str:
    repo, tag = identity["repository"], identity["tag"]
    found = milestone(root, identity)
    evidence.require(found is not None, f"no milestone titled {tag}…: the closeout needs one")
    assert found is not None
    if found.get("state") != "closed":
        evidence.checked(["gh", "api", "-X", "PATCH", f"repos/{repo}/milestones/{found['number']}",
                          "-f", "state=closed"], root, "milestone closing")
    runs = f"repos/{repo}/actions/workflows/{CLOSEOUT}/runs?per_page=1"
    before = evidence.gh_json(root, runs).get("workflow_runs", [])
    last = before[0]["id"] if before else 0
    evidence.checked(["gh", "workflow", "run", CLOSEOUT, "--repo", repo, "--ref", "main",
                      "-f", f"tag={tag}", "-f", f"candidate_sha={identity['sha']}",
                      "-f", f"milestone_number={found['number']}",
                      "-f", f"wiki_browser_reviewed={'true' if wiki_reviewed else 'false'}",
                      "-f", "expect_prerelease=false", "-f", "allow_not_latest=false"], root, "closeout dispatch")
    for _ in range(POLL_LIMIT):
        time.sleep(POLL_SECONDS)
        latest = evidence.gh_json(root, runs).get("workflow_runs", [])
        if latest and latest[0]["id"] != last and latest[0].get("status") == "completed":
            run = latest[0]
            evidence.require(run.get("conclusion") == "success",
                             f"Release closeout concluded {run.get('conclusion')}: {run.get('html_url')}")
            return str(run.get("html_url"))
    raise evidence.EvidenceError("Release closeout did not finish in time; check the Actions tab")


def publish(options: argparse.Namespace) -> int:
    root, work = options.root.resolve(), options.work_dir.resolve()
    identity = evidence.candidate_identity(root, options.tag, options.candidate_sha)
    evidence.require(identity["template"], "publish_release.py applies to the canonical template only")
    evidence.require((work / "release-evidence.json").is_file(),
                     f"run prepare_release_evidence.py prepare first: {work} has no release-evidence.json")
    key = signing_key(options.signing_key)
    publish_tag(root, identity, work, key, options.yes)
    bundle = sign_bundle(root, identity, work, key, options.redaction_reviewed)
    publish_release(root, identity, work, bundle, options.yes)
    url = close_out(root, identity, options.wiki_browser_reviewed)
    print(f"PUBLISHED: {identity['tag']} is released and its closeout passed: {url}")
    return 0


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    result.add_argument("--root", type=Path, default=Path("."))
    result.add_argument("--tag", required=True)
    result.add_argument("--candidate-sha", required=True)
    result.add_argument("--work-dir", type=Path, required=True)
    result.add_argument("--signing-key", type=Path)
    result.add_argument("--redaction-reviewed", action="store_true")
    result.add_argument("--wiki-browser-reviewed", action="store_true")
    result.add_argument("--yes", action="store_true", help="answer yes to every confirmation")
    return result


def main(arguments: list[str] | None = None) -> int:
    options = parser().parse_args(arguments)
    try:
        return publish(options)
    except evidence.EvidenceError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
