#!/usr/bin/env python3
"""Offline fixtures for run_clean_room_pilot.py: GitHub, git, Excel and sibling tools are simulated."""

from __future__ import annotations

import contextlib
import io
import json
import os
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
import publish_release  # noqa: E402
import run_clean_room_pilot as tool  # noqa: E402

SHA = "a" * 40
INIT = "b" * 40
CANDIDATE = "c" * 40
TEMPLATE = "owner/template"
PILOT = "owner/pilot-140"
SUITES = {"suites": [{"module": "TextTests", "status": "PASS", "detail": "cases=29; assertions=135"}],
          "examples": [{"module": "TextExample", "status": "PASS", "detail": "ran without error"}]}


class Fake:
    """Simulated external world. ``fail`` makes the first command containing it fail."""

    def __init__(self, base: Path, fail: str = "", checks: str = "success", settings: str = "pass",
                 blocked: bool = False, rerun: str = "no-op") -> None:
        self.base, self.fail, self.checks, self.settings = base, fail, checks, settings
        self.blocked, self.rerun = blocked, rerun
        self.head = INIT
        self.commands: list[list[str]] = []
        self.streamed: list[list[str]] = []

    def done(self, command: list[str], out: Any = b"", code: int = 0) -> subprocess.CompletedProcess[bytes]:
        data = out if isinstance(out, bytes) else json.dumps(out).encode()
        return subprocess.CompletedProcess(command, code, data, b"simulated failure\n" if code else b"")

    def __call__(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
        self.commands.append(command)
        if self.fail and self.fail in " ".join(command):
            return self.done(command, code=1)
        if command[0] == "git":
            return self.git(command, cwd)
        if command[:2] == ["gh", "api"]:
            return self.gh(command)
        if command[0] == "gh":
            if command[1:3] == ["pr", "merge"]:
                self.head = CANDIDATE if self.head != INIT or any("cut_release.py" in " ".join(item)
                                                                    for item in self.commands) else INIT
            if command[1:3] == ["auth", "token"]:
                return self.done(command, b"token\n")
            return self.done(command)
        return self.tool(command)

    def git(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
        if command[1] == "rev-parse":
            return self.done(command, (SHA if cwd.name == "template" else self.head).encode())
        if command[1] == "clone":
            target = Path(command[-1])
            (target / ".github").mkdir(parents=True)
            (target / ".github/initialization.json").write_text(json.dumps({"template_contract": {"version": "1.2.0"}}))
            (target / "CHANGELOG.md").write_text("## [Unreleased]\n\nNo unreleased changes recorded.\n")
        return self.done(command)

    def gh(self, command: list[str]) -> subprocess.CompletedProcess[bytes]:
        endpoint = command[-1] if command[2] != "-X" else command[4]
        if endpoint == f"repos/{TEMPLATE}/commits/main":
            return self.done(command, {"sha": SHA})
        if "/pulls?head=" in endpoint:
            return self.done(command, [{"number": 7, "head": {"sha": "d" * 40}}])
        if "/check-runs" in endpoint:
            return self.done(command, {"check_runs": [{"name": "Repository integrity", "status": "completed",
                                                       "conclusion": self.checks}]})
        return self.done(command, {})

    def tool(self, command: list[str]) -> subprocess.CompletedProcess[bytes]:
        script = Path(command[1]).name
        if script == "initialize_repository.py" and "--apply" not in command:
            return self.done(command, {"status": self.rerun})
        if script == "provision_repository.py":
            output = Path(command[command.index("--output") + 1])
            plan = {"plan_sha256": "e" * 64, "actions": [1, 2], "blocked": ["ruleset"] if self.blocked else []}
            output.write_text(json.dumps({"plan": plan, "verification": {"remaining_actions": [], "blocked": []}}))
        if script == "check_repository_settings.py":
            Path(command[command.index("--output") + 1]).write_text(json.dumps({"status": self.settings}))
        return self.done(command)

    def stream(self, command: list[str], cwd: Path) -> int:
        self.streamed.append(command)
        bundle = Path(command[command.index("-BundleDir") + 1])
        bundle.mkdir(parents=True)
        (bundle / "suites.json").write_text(json.dumps(SUITES))
        return 1 if self.fail == "Certify-Excel" else 0


class CleanRoomPilotTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        base = Path(self.temporary.name)
        self.template, self.work = base / "template", base / "work"
        (self.template / ".github").mkdir(parents=True)
        (self.template / ".github/repository-profile.json").write_text(json.dumps({"repository": TEMPLATE}))
        for item in (patch.object(time, "sleep", return_value=None), patch.object(tool, "on_windows", return_value=True),
                     patch.dict(os.environ, {}, clear=False)):
            item.start()
            self.addCleanup(item.stop)
        os.environ.pop("GH_TOKEN", None)
        self.closeout = patch.object(publish_release, "close_out", return_value="https://example.invalid/run/9")
        self.closeout.start()
        self.addCleanup(self.closeout.stop)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def main(self, fake: Fake, *extra: str, answer: str = "y", fresh: bool = True) -> tuple[int, str, str]:
        if fresh:
            self.runs = getattr(self, "runs", 0) + 1
            self.work = self.work.with_name(f"work{self.runs}")
        out, err = io.StringIO(), io.StringIO()
        arguments = ["--template-root", str(self.template), "--template-sha", SHA, "--owner", "owner",
                     "--name", "pilot-140", "--work-dir", str(self.work), "--maintainer", "Maintainer", *extra]
        with patch.object(evidence, "run", fake), patch.object(tool, "stream", fake.stream), \
                patch("builtins.input", return_value=answer), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = tool.main(arguments)
        return code, out.getvalue(), err.getvalue()

    def record(self) -> dict[str, Any]:
        value: dict[str, Any] = json.loads((self.work / "clean-room-journey.json").read_text())
        return value

    def test_full_journey_records_every_step_and_deletes_the_pilot(self) -> None:
        fake = Fake(self.work)
        code, out, err = self.main(fake, "--yes")
        self.assertEqual(code, 0, err)
        self.assertIn("JOURNEY PASS", out)
        steps = self.record()["steps"]
        self.assertEqual([row["result"] for row in steps], ["PASS"] * 8)
        self.assertEqual(steps[-1]["name"], "Clean up (delete)")
        texts = [" ".join(command) for command in fake.commands]
        self.assertIn(f"gh repo create {PILOT} --public --template {TEMPLATE}", texts)
        self.assertTrue(any("--approve-plan " + "e" * 64 in text for text in texts))
        self.assertTrue(any(text.startswith("git tag -a v0.1.0 " + CANDIDATE) for text in texts))
        self.assertIn(f"gh repo delete {PILOT} --yes", texts)
        self.assertEqual(sum("actions/permissions" in text and "-X PUT" in text for text in texts), 4)
        prepare = next(text for text in texts if "prepare_release_evidence.py prepare" in text)
        self.assertIn("--check public-api=", prepare)
        self.assertIn("-BundleDir", fake.streamed[0])
        markdown = (self.work / "clean-room-journey.md").read_text()
        self.assertIn("Verdict: **PASS**; 8 steps", markdown)
        self.assertEqual(os.environ["GH_TOKEN"], "token")
        detach = texts.index("git switch --detach")
        entry = next(index for index, text in enumerate(texts) if "Record the initial release entry" in text)
        cut = next(index for index, text in enumerate(texts) if "cut_release.py" in text)
        self.assertLess(detach, entry)
        self.assertLess(entry, cut)

    def test_failure_is_recorded_and_the_failed_pilot_is_kept(self) -> None:
        cases = (
            (Fake(self.work, checks="failure"), "checks failed", 2),
            (Fake(self.work, blocked=True), "plan is blocked", 3),
            (Fake(self.work, settings="fail"), "settings check reports fail", 4),
            (Fake(self.work, fail="Certify-Excel"), "Certify-Excel did not pass", 6),
            (Fake(self.work, rerun="changed"), "not a no-op", 2),
        )
        for fake, message, step in cases:
            with self.subTest(message=message):
                code, out, err = self.main(fake, "--yes", answer="n")
                self.assertEqual(code, 1)
                self.assertIn(message, err)
                steps = self.record()["steps"]
                self.assertEqual((steps[-1]["number"], steps[-1]["result"]), (step, "FAIL"))
                self.assertIn(f"FAIL at step {step}", (self.work / "clean-room-journey.md").read_text())
                self.assertIn("KEPT", out)
                self.assertFalse(any(command[:3] == ["gh", "repo", "delete"] for command in fake.commands))

    def test_archive_keep_and_nothing_to_clean(self) -> None:
        fake = Fake(self.work)
        self.assertEqual(self.main(fake, "--cleanup", "archive", "--yes")[0], 0)
        self.assertTrue(any("archived=true" in command for command in fake.commands))
        kept = Fake(self.work)
        code, _, _ = self.main(kept, "--cleanup", "keep")
        self.assertEqual(code, 0)
        self.assertFalse(any("delete" in command or "archived=true" in command for command in kept.commands))
        early = Fake(self.work, fail="repo create")
        self.assertEqual(self.main(early, "--yes")[0], 1)
        self.assertFalse(any(command[:3] == ["gh", "repo", "delete"] for command in early.commands))

    def test_prerequisites(self) -> None:
        with patch.object(tool, "on_windows", return_value=False):
            code, _, err = self.main(Fake(self.work), "--yes")
            self.assertEqual(code, 1)
            self.assertIn("needs Windows", err)
        code, _, err = self.main(Fake(self.work), "--profile", "application")
        self.assertIn("--check NAME=DETAIL", err)

        class Moved(Fake):
            def gh(self, command: list[str]) -> subprocess.CompletedProcess[bytes]:
                if command[-1].endswith("/commits/main"):
                    return self.done(command, {"sha": "f" * 40})
                return super().gh(command)

        self.assertIn("Use this template copies the default branch", self.main(Moved(self.work), "--yes")[2])

        class Elsewhere(Fake):
            def git(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
                return self.done(command, b"f" * 40) if command[1] == "rev-parse" else super().git(command, cwd)

        self.assertIn("the template checkout is at", self.main(Elsewhere(self.work))[2])

    def test_reused_work_dir_is_refused(self) -> None:
        self.assertEqual(self.main(Fake(self.work), "--cleanup", "keep")[0], 0)
        kept = self.work / "release-work" / "evidence" / "keep-me.json"
        kept.parent.mkdir(parents=True, exist_ok=True)
        kept.write_text("{}")
        code, _, err = self.main(Fake(self.work), "--cleanup", "keep", fresh=False)
        self.assertEqual((code, "fresh --work-dir" in err), (1, True))
        self.assertTrue(kept.is_file())

    def test_library_checks_need_passing_reference_results(self) -> None:
        bundle = self.work / "bundle"
        bundle.mkdir(parents=True)
        failing = {**SUITES, "examples": [{"module": "TextExample", "status": "FAIL", "detail": "error"}]}
        (bundle / "suites.json").write_text(json.dumps(failing))
        with self.assertRaisesRegex(evidence.EvidenceError, "did not pass in Excel"):
            tool.library_checks(bundle)

    def test_wait_for_times_out_and_stream_runs(self) -> None:
        with patch.object(tool, "POLL_LIMIT", 2), self.assertRaisesRegex(evidence.EvidenceError, "timed out"):
            tool.wait_for(lambda: False, "nothing")
        self.assertEqual(tool.stream([sys.executable, "-c", "pass"], Path(".")), 0)


if __name__ == "__main__":
    unittest.main()
