#!/usr/bin/env python3
"""Cut a release candidate: VERSION, the dated CHANGELOG section and its links, on a reviewable branch.

From a clean checkout of the default branch, it:

1. moves the Unreleased entries into ``## [X.Y.Z] - DATE``, keeping any
   Unreleased guidance comment and recreating an empty Unreleased section;
2. points ``[Unreleased]`` at ``vX.Y.Z...HEAD`` and adds ``[X.Y.Z]``, a
   comparison with the previous release or, for the first release, a link to
   its tag;
3. writes VERSION;
4. commits ``Cut vX.Y.Z release candidate`` on ``chore/cut-vX.Y.Z``;
5. runs ``check_release_semantics.py`` on that commit.

With ``--push`` it also pushes the branch and opens the pull request with
``gh``. The comparison links of the new version need no classification: the
external-link check reports them as pending until the tag is published.
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path

import prepare_release_evidence as evidence

SEMVER = re.compile(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)")
EMPTY_MARKER = "No unreleased changes recorded."


def key(version: str) -> tuple[int, int, int]:
    match = SEMVER.fullmatch(version)
    evidence.require(match is not None, f"{version!r} is not a MAJOR.MINOR.PATCH version")
    assert match is not None
    return int(match[1]), int(match[2]), int(match[3])


def cut_changelog(text: str, version: str, day: str, repository: str) -> str:
    """Return the CHANGELOG with the Unreleased entries moved under the new version."""
    heading = re.search(r"^## \[Unreleased\][ \t]*\n", text, re.MULTILINE)
    evidence.require(heading is not None, "CHANGELOG.md has no ## [Unreleased] section")
    assert heading is not None
    following = re.search(r"^## \[|^---[ \t]*$", text[heading.end():], re.MULTILINE)
    end = heading.end() + (following.start() if following else len(text) - heading.end())
    body = text[heading.end():end]
    comments = "".join(re.findall(r"<!--.*?-->[ \t]*\n?", body, re.DOTALL))
    entries = re.sub(r"<!--.*?-->[ \t]*\n?", "", body, flags=re.DOTALL).replace(EMPTY_MARKER, "").strip()
    evidence.require(bool(entries), "the Unreleased section has no entries to release")
    released = re.findall(r"^## \[(\d+\.\d+\.\d+)\] - ", text, re.MULTILINE)
    evidence.require(version not in released, f"CHANGELOG.md already has a section for {version}")
    previous = max(released, key=key) if released else None
    if previous is not None:
        evidence.require(key(version) > key(previous), f"{version} is not newer than {previous}")
    unreleased = "## [Unreleased]\n\n" + (comments.strip() + "\n\n" if comments.strip() else "")
    section = f"## [{version}] - {day}\n\n{entries}\n\n"
    text = text[:heading.start()] + unreleased + section + text[end:].lstrip("\n")
    base = f"https://github.com/{repository}"
    target = f"{base}/compare/v{previous}...v{version}" if previous else f"{base}/releases/tag/v{version}"
    links = f"[Unreleased]: {base}/compare/v{version}...HEAD\n[{version}]: {target}\n"
    existing = re.search(r"^\[Unreleased\]:[^\n]*\n?", text, re.MULTILINE)
    if existing is not None:
        return text[:existing.start()] + links + text[existing.end():]
    return text.rstrip("\n") + "\n\n" + links


def cut(options: argparse.Namespace) -> int:
    root, version = options.root.resolve(), options.version
    key(version)
    status = evidence.checked(["git", "status", "--porcelain"], root, "git status").decode().strip()
    evidence.require(not status, "the checkout has uncommitted changes")
    profile = (root / ".github/repository-profile.json").read_text(encoding="utf-8")
    repository = re.search(r'"repository":\s*"([^"]+)"', profile)
    evidence.require(repository is not None, "the repository profile names no repository")
    assert repository is not None
    changelog = root / "CHANGELOG.md"
    updated = cut_changelog(changelog.read_text(encoding="utf-8"), version, options.date, repository[1])
    branch = f"chore/cut-v{version}"
    evidence.checked(["git", "switch", "-c", branch], root, f"creating {branch}")
    changelog.write_text(updated, encoding="utf-8")
    (root / "VERSION").write_text(version + "\n", encoding="utf-8")
    evidence.checked(["git", "add", "CHANGELOG.md", "VERSION"], root, "staging the cut")
    evidence.checked(["git", "commit", "--quiet", "-m", f"Cut v{version} release candidate"], root, "commit")
    evidence.checked(evidence.tool(root, "check_release_semantics.py", "--root", "."), root,
                     f"release semantics on {branch} (the branch is kept for inspection)")
    if options.push:
        evidence.checked(["git", "push", "--set-upstream", "origin", branch], root, "branch push")
        evidence.checked(["gh", "pr", "create", "--base", options.base, "--head", branch,
                          "--title", f"Cut v{version} release candidate",
                          "--body", f"Release cut for v{version}: VERSION, the {options.date} CHANGELOG section and "
                                    "its comparison links. Release semantics pass on this commit."],
                         root, "pull request")
        print(f"CUT: {branch} pushed and its pull request opened")
    else:
        print(f"CUT: {branch} committed locally and release semantics pass. After review, publish it with:\n"
              f"  git push --set-upstream origin {branch}\n"
              f"  gh pr create --base {options.base} --head {branch} --title \"Cut v{version} release candidate\" "
              "--fill")
    return 0


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    result.add_argument("--root", type=Path, default=Path("."))
    result.add_argument("--version", required=True)
    result.add_argument("--date", default=date.today().isoformat(), help="cut/freeze date, default today")
    result.add_argument("--base", default="main")
    result.add_argument("--push", action="store_true", help="push the branch and open the pull request")
    return result


def main(arguments: list[str] | None = None) -> int:
    options = parser().parse_args(arguments)
    try:
        return cut(options)
    except evidence.EvidenceError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
