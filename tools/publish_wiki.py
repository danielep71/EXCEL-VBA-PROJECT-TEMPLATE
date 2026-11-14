#!/usr/bin/env python3
"""Publish the Wiki from one exact candidate and read the publication back.

Run it on the maintainer's machine, from a clean checkout of the candidate, with
Git credentials that can push to the ``OWNER/REPO.wiki`` repository. It:

1. exports the Wiki bundle with ``check_wiki.py --export-dir``;
2. clones the published Wiki, and refuses to delete a published page that the
   bundle no longer has unless ``--remove-page NAME`` names it as reviewed;
3. replaces the published files and checks the result with
   ``check_wiki.py --published-dir`` before committing anything;
4. commits ``Publish wiki from <candidate>`` and, after confirmation, pushes
   without force;
5. clones the Wiki again, compares it byte for byte with the candidate and
   writes ``wiki-publication.json`` with the source, Wiki commit, parent and
   the SHA-256 of every file.

When the published Wiki already matches the candidate, nothing is committed or
pushed, and the read-back still runs and records the existing publication.
Everything is written under ``--work-dir``, outside the candidate tree.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from pathlib import Path
from typing import Any

import prepare_release_evidence as evidence
from publish_release import confirm

PUBLISHED = (".md", ".json")


def published_files(directory: Path) -> set[str]:
    return {path.name for path in directory.iterdir() if path.is_file() and path.suffix in PUBLISHED}


def candidate(root: Path, sha: str) -> str:
    head = evidence.checked(["git", "rev-parse", "HEAD"], root, "git rev-parse").decode().strip()
    evidence.require(head == sha, f"checkout is at {head}, not the candidate {sha}")
    status = evidence.checked(["git", "status", "--porcelain"], root, "git status").decode().strip()
    evidence.require(not status, "the checkout has uncommitted changes")
    profile = json.loads((root / ".github/repository-profile.json").read_text(encoding="utf-8"))
    evidence.require(profile.get("mode") == "template", "the Wiki is published from the canonical template only")
    return str(profile["repository"])


def fresh_clone(root: Path, repository: str, target: Path) -> None:
    if target.exists():
        shutil.rmtree(target)
    evidence.checked(["git", "clone", "--quiet", f"https://github.com/{repository}.wiki.git", str(target)],
                     root, "Wiki clone")


def compare(root: Path, published: Path, what: str) -> None:
    evidence.checked(evidence.tool(root, "check_wiki.py", "--root", ".", "--published-dir", str(published)),
                     root, what)


def stage(root: Path, export: Path, checkout: Path, removable: list[str]) -> list[str]:
    """Replace the published files with the bundle; return the pages it removes."""
    exported, current = published_files(export), published_files(checkout)
    removed = sorted(current - exported)
    unreviewed = [name for name in removed if name not in removable]
    evidence.require(not unreviewed, "the bundle drops published pages; review them and pass --remove-page for "
                     "each: " + ", ".join(unreviewed))
    for name in removed:
        (checkout / name).unlink()
    for name in sorted(exported):
        shutil.copyfile(export / name, checkout / name)
    return removed


def record(repository: str, sha: str, readback: Path, root: Path) -> dict[str, Any]:
    log = evidence.checked(["git", "-C", str(readback), "log", "-2", "--format=%H"], root, "Wiki log")
    commits = log.decode().split()
    files = {name: hashlib.sha256((readback / name).read_bytes()).hexdigest()
             for name in sorted(published_files(readback))}
    return {"schema_version": 1, "role": "wiki-publication", "repository": repository, "source_sha": sha,
            "wiki_repository": f"{repository}.wiki", "wiki_commit": commits[0],
            "wiki_parent": commits[1] if len(commits) > 1 else None,
            "readback": "fresh clone; tools/check_wiki.py --published-dir", "readback_result": "PASS",
            "files": files}


def publish(options: argparse.Namespace) -> int:
    root, work, sha = options.root.resolve(), options.work_dir.resolve(), options.candidate_sha
    repository = candidate(root, sha)
    work.mkdir(parents=True, exist_ok=True)
    export, checkout, readback = work / "wiki-export", work / "wiki-checkout", work / "wiki-readback"
    if export.exists():
        shutil.rmtree(export)
    evidence.checked(evidence.tool(root, "check_wiki.py", "--root", ".", "--export-dir", str(export)), root,
                     "Wiki export")
    fresh_clone(root, repository, checkout)
    removed = stage(root, export, checkout, options.remove_page)
    compare(root, checkout, "comparison of the staged Wiki with the candidate")
    evidence.checked(["git", "-C", str(checkout), "add", "--all"], root, "Wiki staging")
    changed = evidence.run(["git", "-C", str(checkout), "diff", "--cached", "--quiet"], root).returncode != 0
    if changed:
        evidence.checked(["git", "-C", str(checkout), "commit", "--quiet", "-m", f"Publish wiki from {sha}"],
                         root, "Wiki commit")
        removal = f"; it removes {', '.join(removed)}" if removed else ""
        confirm(f"Push the Wiki published from {sha} to {repository}.wiki{removal}?", options.yes)
        evidence.checked(["git", "-C", str(checkout), "push", "origin", "HEAD"], root, "Wiki push")
    else:
        print(f"SKIP: the published Wiki already matches {sha}")
    fresh_clone(root, repository, readback)
    compare(root, readback, "read-back of the published Wiki")
    output = work / "wiki-publication.json"
    evidence.write_json(output, record(repository, sha, readback, root))
    print(f"WIKI PUBLISHED AND VERIFIED: {output}")
    return 0


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    result.add_argument("--root", type=Path, default=Path("."))
    result.add_argument("--candidate-sha", required=True)
    result.add_argument("--work-dir", type=Path, required=True)
    result.add_argument("--remove-page", action="append", default=[],
                        help="a published page the bundle no longer has, reviewed for removal")
    result.add_argument("--yes", action="store_true", help="push without asking for confirmation")
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
