#!/usr/bin/env python3
"""Offline fixtures for publish_release.py: git, gh, ssh-keygen and sibling tools are simulated."""

from __future__ import annotations

import contextlib
import io
import json
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from typing import Any
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))

import prepare_release_evidence as evidence  # noqa: E402
import publish_release as tool  # noqa: E402

SHA = "a" * 40
REPO = "owner/EXCEL-VBA-PROJECT-TEMPLATE"
CHANGELOG = """# Changelog

## [Unreleased]

## [1.3.0] - 2026-11-14

### Added

- Something useful.

## [1.2.1] - 2026-09-28

### Fixed

- Older.

[1.3.0]: https://github.com/owner/EXCEL-VBA-PROJECT-TEMPLATE/compare/v1.2.1...v1.3.0
"""


class Fake:
    """Simulated external commands with a small amount of remote state."""

    def __init__(self, *, remote_tag: bool = False, release: bool = False, fail: str = "",
                 milestone_state: str = "open", closeout: str = "success", milestones: bool = True,
                 draft: bool = False, assets: int = 4, other_run: bool = False, verified_tag: str = "v1.3.0") -> None:
        self.remote_tag, self.release, self.fail = remote_tag, release, fail
        self.draft, self.assets, self.other_run, self.verified_tag = draft, assets, other_run, verified_tag
        self.local_tag = False
        self.milestone_state, self.closeout, self.milestones = milestone_state, closeout, milestones
        self.dispatched = False
        self.commands: list[list[str]] = []

    def done(self, command: list[str], code: int = 0, out: Any = b"") -> subprocess.CompletedProcess[bytes]:
        data = out if isinstance(out, bytes) else json.dumps(out).encode()
        return subprocess.CompletedProcess(command, code, data, b"simulated failure\n" if code else b"")

    def __call__(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
        self.commands.append(command)
        if self.fail and self.fail in " ".join(command):
            return self.done(command, 1)
        if command[0] == "git":
            return self.git(command)
        if command[:2] == ["gh", "api"]:
            return self.gh(command)
        if "--verify-bundle" in command:
            return self.done(command, out={"status": "pass", "tag": self.verified_tag, "candidate_sha": SHA})
        self.effects(command)
        return self.done(command)

    def git(self, command: list[str]) -> subprocess.CompletedProcess[bytes]:
        verb = command[1]
        if verb == "rev-parse":
            if "--verify" in command:
                return self.done(command, 0 if self.local_tag else 1)
            return self.done(command, out=SHA.encode())
        if verb == "ls-remote":
            return self.done(command, out=f"{'b' * 40}\trefs/tags/v1.3.0\n".encode() if self.remote_tag else b"")
        if "-s" in command:
            self.local_tag = True
        elif command[-2:-1] == ["-d"]:
            self.local_tag = False
        elif verb == "push":
            self.remote_tag = True
        return self.done(command)

    def effects(self, command: list[str]) -> None:
        """Files and remote state that the simulated tool or gh command leaves behind."""
        if command[0] == "ssh-keygen":
            Path(command[-1] + ".sig").write_text("signature")
        elif "--build" in command:
            target = Path(command[command.index("--output-dir") + 1])
            target.mkdir(parents=True)
            (target / "certification-v1.3.0.zip").write_bytes(b"zip")
            (target / "certification-v1.3.0.manifest.json").write_text("{}")
            (target / "certification-v1.3.0.sha256").write_text("digest")
        elif "check_release.py" in " ".join(command) and "--output" in command:
            Path(command[command.index("--output") + 1]).write_text("{}")
        action = tuple(command[1:3])
        if action == ("release", "create"):
            self.release, self.draft, self.assets = True, False, 4
        elif action == ("release", "upload"):
            self.assets = 4
        elif action == ("release", "edit"):
            self.draft = False
        elif action == ("workflow", "run"):
            self.dispatched = True

    def gh(self, command: list[str]) -> subprocess.CompletedProcess[bytes]:
        endpoint = command[-1] if command[2] != "-X" else command[4]
        if "/releases?per_page" in endpoint:
            names = ["certification-v1.3.0.zip", "certification-v1.3.0.zip.sig",
                     "certification-v1.3.0.manifest.json", "certification-v1.3.0.sha256"][:self.assets]
            release = {"tag_name": "v1.3.0", "draft": self.draft, "assets": [{"name": name} for name in names]}
            return self.done(command, out=[{"tag_name": "v1.2.1", "draft": False, "assets": []}]
                             + ([release] if self.release else []))
        if "milestones?state=all" in endpoint:
            items = [{"number": 5, "title": "v1.3.0 — Reference", "state": self.milestone_state}]
            return self.done(command, out=items if self.milestones else [])
        if command[2] == "-X":
            self.milestone_state = "closed"
            return self.done(command, out={})
        if f"/{tool.CLOSEOUT}/runs" in endpoint:
            old = {"id": 1, "status": "completed", "display_title": f"Release closeout — v1.3.0 @ {SHA}"}
            if not self.dispatched:
                return self.done(command, out={"workflow_runs": [old]})
            runs = [{"id": 2, "status": "completed", "conclusion": self.closeout,
                     "display_title": f"Release closeout — v1.3.0 @ {SHA}",
                     "html_url": f"https://github.com/{REPO}/actions/runs/2"}, old]
            if self.other_run:
                runs.insert(0, {"id": 3, "status": "completed", "conclusion": "success",
                                "display_title": f"Release closeout — v9.9.9 @ {'c' * 40}"})
            return self.done(command, out={"workflow_runs": runs})
        raise AssertionError(f"unexpected endpoint {endpoint}")


class PublishReleaseTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        base = Path(self.temporary.name)
        self.root, self.work = base / "repo", base / "work"
        (self.root / ".github").mkdir(parents=True)
        (self.root / "VERSION").write_text("1.3.0\n")
        (self.root / "CHANGELOG.md").write_text(CHANGELOG)
        (self.root / ".github/repository-profile.json").write_text(json.dumps({"mode": "template", "repository": REPO}))
        (self.work / "evidence").mkdir(parents=True)
        (self.work / "release-evidence.json").write_text("{}")
        (self.work / "evidence/host.json").write_text("{}")
        (self.work / "evidence/suites.json").write_text(json.dumps({
            "suites": [{"entry_point": "ProjectTests.RunProjectTests", "status": "PASS", "detail": "cases=4"}],
            "examples": [{"entry_point": "ProjectExample.RunProjectExample", "status": "PASS",
                          "detail": "ran without error"}]}))
        self.key = base / "release_signing"
        self.key.write_text("private")
        sleep = patch.object(time, "sleep", return_value=None)
        sleep.start()
        self.addCleanup(sleep.stop)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def main(self, fake: Fake, *extra: str, answer: str = "y") -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        arguments = ["--root", str(self.root), "--tag", "v1.3.0", "--candidate-sha", SHA,
                     "--work-dir", str(self.work), *extra]
        with patch.object(evidence, "run", fake), patch("builtins.input", return_value=answer), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = tool.main(arguments)
        return code, out.getvalue(), err.getvalue()

    def test_full_publication_then_resume_skips_every_step(self) -> None:
        fake = Fake()
        code, out, err = self.main(fake, "--signing-key", str(self.key), "--redaction-reviewed",
                                   "--wiki-browser-reviewed")
        self.assertEqual(code, 0, err)
        self.assertIn("PUBLISHED", out)
        texts = [" ".join(command) for command in fake.commands]
        tag = next(text for text in texts if " tag -s " in text)
        self.assertIn("-m EXCEL-VBA-PROJECT-TEMPLATE 1.3.0", tag)
        self.assertIn(f"user.signingkey={self.key}", tag)
        self.assertTrue(any(text.startswith(f"ssh-keygen -Y sign -f {self.key} -n {tool.NAMESPACE}") for text in texts))
        create = next(command for command in fake.commands if command[:3] == ["gh", "release", "create"])
        self.assertEqual(sum(item.startswith(str(self.work / "out")) for item in create), 4)
        self.assertIn("v1.3.0 — Reference", create)
        self.assertIn("wiki_browser_reviewed=true", " ".join(texts))
        notes = (self.work / "release-notes.md").read_text()
        self.assertIn("#### Added", notes)
        self.assertNotIn("Older.", notes)
        self.assertIn("| `ProjectTests.RunProjectTests` | PASS | cases=4 |", notes)
        self.assertIn("compare/v1.2.1...v1.3.0", notes)
        self.assertEqual(fake.milestone_state, "closed")

        resumed = Fake(remote_tag=True, release=True, milestone_state="closed")
        code, out, err = self.main(resumed, "--signing-key", str(self.key), "--redaction-reviewed")
        self.assertEqual(code, 0, err)
        self.assertIn("SKIP: v1.3.0 is already published", out)
        self.assertIn("SKIP: the GitHub Release v1.3.0 is published with all four files", out)
        self.assertIn("SKIP: the certification bundle is complete", out)
        self.assertFalse(any(command[0] == "ssh-keygen" for command in resumed.commands))
        self.assertFalse(any("--build" in command for command in resumed.commands))

    def test_interrupted_bundle_is_rebuilt_not_trusted(self) -> None:
        fake = Fake(remote_tag=True, release=True, milestone_state="closed")
        self.assertEqual(self.main(fake, "--signing-key", str(self.key), "--redaction-reviewed")[0], 0)
        marker = self.work / "bundle-complete.json"
        for damage in ("missing-marker", "changed-zip", "wrong-tag"):
            with self.subTest(damage=damage):
                again = Fake(remote_tag=True, release=True, milestone_state="closed",
                             verified_tag="v0.0.1" if damage == "wrong-tag" else "v1.3.0")
                if damage == "missing-marker":
                    marker.unlink()
                if damage == "changed-zip":
                    (self.work / "out/certification-v1.3.0.zip").write_bytes(b"tampered")
                code, out, err = self.main(again, "--signing-key", str(self.key), "--redaction-reviewed")
                self.assertTrue(any("--build" in command for command in again.commands), out + err)
                self.assertEqual(code, 0, err)
                self.assertTrue(marker.is_file())

    def test_interrupted_draft_release_is_completed(self) -> None:
        fake = Fake(remote_tag=True, release=True, draft=True, assets=2)
        code, out, err = self.main(fake, "--signing-key", str(self.key), "--redaction-reviewed", "--yes")
        self.assertEqual(code, 0, err)
        upload = next(command for command in fake.commands if command[:3] == ["gh", "release", "upload"])
        self.assertEqual([Path(item).name for item in upload[4:6]],
                         ["certification-v1.3.0.manifest.json", "certification-v1.3.0.sha256"])
        self.assertTrue(any(command[:3] == ["gh", "release", "edit"] for command in fake.commands))
        self.assertFalse(any(command[:3] == ["gh", "release", "create"] for command in fake.commands))
        published = Fake(remote_tag=True, release=True, draft=True, assets=4)
        self.assertEqual(self.main(published, "--signing-key", str(self.key), "--yes")[0], 0)
        self.assertFalse(any(command[:3] == ["gh", "release", "upload"] for command in published.commands))

    def test_closeout_ignores_another_release_run(self) -> None:
        fake = Fake(remote_tag=True, release=True, closeout="failure", other_run=True)
        code, _, err = self.main(fake, "--signing-key", str(self.key), "--redaction-reviewed", "--yes")
        self.assertEqual(code, 1)
        self.assertIn("concluded failure", err)

    def test_failed_post_tag_check_deletes_the_local_tag(self) -> None:
        fake = Fake(fail="--require-tag-ref")
        code, _, err = self.main(fake, "--signing-key", str(self.key))
        self.assertEqual(code, 1)
        self.assertIn("nothing was pushed", err)
        self.assertFalse(fake.local_tag)
        self.assertFalse(any(command[:2] == ["git", "push"] for command in fake.commands))

    def test_declined_confirmations_stop_publication(self) -> None:
        fake = Fake()
        code, _, err = self.main(fake, "--signing-key", str(self.key), answer="n")
        self.assertEqual((code, "maintainer's request" in err), (1, True))
        self.assertFalse(fake.remote_tag)
        later = Fake(remote_tag=True)
        (self.work / "out").mkdir()
        for name in ("certification-v1.3.0.zip", "certification-v1.3.0.zip.sig",
                     "certification-v1.3.0.manifest.json", "certification-v1.3.0.sha256"):
            (self.work / "out" / name).write_text("x")
        code, _, err = self.main(later, "--signing-key", str(self.key), "--redaction-reviewed", answer="no")
        self.assertEqual((code, "maintainer's request" in err), (1, True))
        self.assertFalse(later.release)
        code, out, _ = self.main(Fake(remote_tag=True), "--signing-key", str(self.key), "--redaction-reviewed", "--yes")
        self.assertIn("yes (--yes)", out)

    def test_signing_key_and_prerequisite_failures(self) -> None:
        cases = (((), "RELEASE_SIGNING_KEY"),
                 (("--signing-key", str(self.key) + ".pub"), "not the .pub"),
                 (("--signing-key", str(self.key) + "-missing"), "does not exist"))
        for extra, message in cases:
            with self.subTest(message=message), patch.dict("os.environ", {}, clear=True):
                code, _, err = self.main(Fake(), *extra)
                self.assertEqual(code, 1)
                self.assertIn(message, err)
        with patch.dict("os.environ", {"RELEASE_SIGNING_KEY": str(self.key)}):
            self.assertEqual(self.main(Fake(fail="ssh-keygen"))[0], 1)
        (self.work / "release-evidence.json").unlink()
        self.assertIn("run prepare_release_evidence.py prepare first", self.main(Fake(), "--signing-key", str(self.key))[2])
        (self.root / ".github/repository-profile.json").write_text(json.dumps(
            {"mode": "generated", "repository": REPO, "profile": "library"}))
        self.assertIn("canonical template only", self.main(Fake(), "--signing-key", str(self.key))[2])

    def test_closeout_failures(self) -> None:
        cases = ((Fake(closeout="failure"), "concluded failure"),
                 (Fake(milestones=False), "no milestone titled"))
        for fake, message in cases:
            with self.subTest(message=message):
                self.work = self.work  # same prepared work directory; each run resumes
                code, _, err = self.main(fake, "--signing-key", str(self.key), "--redaction-reviewed", "--yes")
                self.assertEqual(code, 1)
                self.assertIn(message, err)
        with patch.object(tool, "POLL_LIMIT", 0):
            code, _, err = self.main(Fake(remote_tag=True, release=True), "--signing-key", str(self.key))
            self.assertIn("did not finish in time", err)

    def test_changelog_section_must_exist(self) -> None:
        with self.assertRaisesRegex(evidence.EvidenceError, "no released section for 9.9.9"):
            tool.changelog_section(CHANGELOG, "9.9.9")
        body, link = tool.changelog_section("## [1.0.0] - 2026-01-01\n\nText.\n", "1.0.0")
        self.assertEqual((body, link), ("Text.", None))


if __name__ == "__main__":
    unittest.main()
