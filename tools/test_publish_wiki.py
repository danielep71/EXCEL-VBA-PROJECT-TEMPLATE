#!/usr/bin/env python3
"""Offline fixtures for publish_wiki.py: git and check_wiki.py are simulated."""

from __future__ import annotations

import contextlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))

import prepare_release_evidence as evidence  # noqa: E402
import publish_wiki as tool  # noqa: E402

SHA = "a" * 40
REPO = "owner/template"


class Fake:
    """A simulated Wiki remote; clones receive its current files."""

    def __init__(self, remote: dict[str, str], export: dict[str, str], *, fail: str = "", head: str = SHA,
                 dirty: str = "") -> None:
        self.remote, self.export, self.fail, self.head, self.dirty = remote, export, fail, head, dirty
        self.pushed = False
        self.commands: list[list[str]] = []

    def __call__(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
        self.commands.append(command)
        out, code = b"", 0
        if self.fail and self.fail in " ".join(command):
            code = 1
        elif command[:2] == ["git", "rev-parse"]:
            out = self.head.encode()
        elif command[:2] == ["git", "status"]:
            out = self.dirty.encode()
        elif command[:2] == ["git", "clone"]:
            target = Path(command[-1])
            target.mkdir(parents=True)
            for name, text in self.remote.items():
                (target / name).write_text(text)
        elif "--export-dir" in command:
            target = Path(command[-1])
            target.mkdir(parents=True)
            for name, text in self.export.items():
                (target / name).write_text(text)
        elif "diff" in command:
            checkout = Path(command[2])
            current = {path.name: path.read_text() for path in checkout.iterdir() if path.is_file()}
            code = 0 if current == self.remote else 1
        elif "push" in command:
            checkout = Path(command[2])
            self.remote = {path.name: path.read_text() for path in checkout.iterdir() if path.is_file()}
            self.pushed = True
        elif "log" in command:
            out = b"b" * 40 + b"\n" + b"c" * 40 + b"\n"
        return subprocess.CompletedProcess(command, code, out, b"simulated failure\n" if code else b"")


class PublishWikiTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        base = Path(self.temporary.name)
        self.root, self.work = base / "repo", base / "work"
        (self.root / ".github").mkdir(parents=True)
        (self.root / ".github/repository-profile.json").write_text(json.dumps({"mode": "template", "repository": REPO}))
        self.export = {"Home.md": "new home\n", "_Sidebar.md": "sidebar\n", "Wiki-Source.json": "{}\n"}

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def main(self, fake: Fake, *extra: str, answer: str = "y") -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        arguments = ["--root", str(self.root), "--candidate-sha", SHA, "--work-dir", str(self.work), *extra]
        with patch.object(evidence, "run", fake), patch("builtins.input", return_value=answer), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = tool.main(arguments)
        return code, out.getvalue(), err.getvalue()

    def test_publishes_reads_back_and_records(self) -> None:
        fake = Fake({"Home.md": "old home\n", "_Sidebar.md": "sidebar\n", "Wiki-Source.json": "{}\n"}, self.export)
        code, out, err = self.main(fake)
        self.assertEqual(code, 0, err)
        self.assertTrue(fake.pushed)
        self.assertEqual(fake.remote, self.export)
        commit = next(command for command in fake.commands if "commit" in command)
        self.assertEqual(commit[-1], f"Publish wiki from {SHA}")
        self.assertFalse(any("--force" in command for command in fake.commands))
        record = json.loads((self.work / "wiki-publication.json").read_text())
        self.assertEqual((record["source_sha"], record["wiki_commit"], record["wiki_parent"]), (SHA, "b" * 40, "c" * 40))
        self.assertEqual(sorted(record["files"]), sorted(self.export))
        compares = [command for command in fake.commands if "--published-dir" in command]
        self.assertEqual([Path(command[-1]).name for command in compares], ["wiki-checkout", "wiki-readback"])
        # A second run finds the Wiki current: nothing is committed or pushed, the read-back still runs.
        again = Fake(dict(fake.remote), self.export)
        code, out, err = self.main(again)
        self.assertEqual(code, 0, err)
        self.assertIn("SKIP: the published Wiki already matches", out)
        self.assertFalse(again.pushed or any("commit" in command for command in again.commands))

    def test_work_dir_overlapping_the_checkout_is_refused_before_any_deletion(self) -> None:
        base = self.root.parent
        source = base / "wiki-export"
        self.root.rename(source)
        self.root = source
        for work in (base, source, source / "scratch"):
            with self.subTest(work=work):
                self.work = work
                fake = Fake({"Home.md": "old\n"}, self.export)
                code, _, err = self.main(fake)
                self.assertEqual(code, 1)
                self.assertIn("overlaps the candidate checkout", err)
                self.assertTrue((source / ".github/repository-profile.json").is_file())
                self.assertFalse(any(command[:2] == ["git", "clone"] for command in fake.commands))

    def test_clones_keep_published_line_endings(self) -> None:
        fake = Fake({"Home.md": "old\n"}, self.export)
        self.assertEqual(self.main(fake)[0], 0)
        clones = [command for command in fake.commands if command[:2] == ["git", "clone"]]
        self.assertEqual(len(clones), 2)
        for command in clones:
            self.assertEqual(command[3:7], ["-c", "core.autocrlf=false", "-c", "core.eol=lf"])

    def test_removed_pages_need_review(self) -> None:
        remote = {**self.export, "Old-page.md": "gone\n"}
        code, _, err = self.main(Fake(dict(remote), self.export))
        self.assertEqual(code, 1)
        self.assertIn("--remove-page for each: Old-page.md", err)
        fake = Fake(dict(remote), self.export)
        code, out, err = self.main(fake, "--remove-page", "Old-page.md", "--yes")
        self.assertEqual(code, 0, err)
        self.assertNotIn("Old-page.md", fake.remote)
        self.assertIn("yes (--yes)", out)

    def test_declined_push_publishes_nothing(self) -> None:
        fake = Fake({"Home.md": "old\n"}, self.export)
        code, _, err = self.main(fake, answer="n")
        self.assertEqual((code, "maintainer's request" in err), (1, True))
        self.assertFalse(fake.pushed)

    def test_prerequisite_and_check_failures(self) -> None:
        remote = {"Home.md": "old\n"}
        cases = ((Fake(remote, self.export, head="b" * 40), "not the candidate"),
                 (Fake(remote, self.export, dirty=" M x"), "uncommitted changes"),
                 (Fake(remote, self.export, fail="--published-dir "), "comparison of the staged Wiki"),
                 (Fake(remote, self.export, fail="wiki-readback"), "Wiki clone failed"),
                 (Fake(remote, self.export, fail="push"), "Wiki push failed"))
        for fake, message in cases:
            with self.subTest(message=message):
                code, _, err = self.main(fake)
                self.assertEqual(code, 1)
                self.assertIn(message, err)
        (self.root / ".github/repository-profile.json").write_text(json.dumps(
            {"mode": "generated", "repository": REPO, "profile": "library"}))
        self.assertIn("canonical template only", self.main(Fake(remote, self.export))[2])


if __name__ == "__main__":
    unittest.main()
