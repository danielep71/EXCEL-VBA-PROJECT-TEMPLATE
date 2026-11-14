#!/usr/bin/env python3
"""Offline fixtures for release_template.py: every tool, git, gh and Excel call is simulated."""

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

import check_excel_evidence as excel  # noqa: E402
import prepare_release_evidence as evidence  # noqa: E402
import release_template as tool  # noqa: E402
import run_clean_room_pilot as journey  # noqa: E402

BASE = "a" * 40
CANDIDATE = "c" * 40
PROVIDER = "e" * 40
REPO = "owner/template"


class World:
    """The checkout, the remote and the release tools, simulated together."""

    def __init__(self, root: Path, *, workflow_changed: bool = False, links_green: bool = True) -> None:
        self.root, self.workflow_changed, self.links_green = root, workflow_changed, links_green
        self.head, self.branch = BASE, "main"
        self.dispatched = False
        self.commands: list[list[str]] = []
        self.streamed: list[str] = []

    def done(self, command: list[str], out: Any = b"", code: int = 0) -> subprocess.CompletedProcess[bytes]:
        data = out if isinstance(out, bytes) else json.dumps(out).encode()
        return subprocess.CompletedProcess(command, code, data, b"simulated failure\n" if code else b"")

    def run(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
        self.commands.append(command)
        if command[:2] == ["git", "rev-parse"]:
            return self.done(command, (BASE if command[-1] == "origin/main" else self.head).encode())
        if command[:2] == ["git", "branch"]:
            return self.done(command, self.branch.encode())
        if command[:2] == ["git", "diff"]:
            return self.done(command, code=1 if self.workflow_changed else 0)
        if command[:2] == ["gh", "api"] and "external-links.yml/runs" in command[-1]:
            ready = self.links_green and self.dispatched
            runs = [{"status": "completed", "conclusion": "success", "html_url": "https://example.invalid/r/1"}]
            return self.done(command, {"workflow_runs": runs if ready else []})
        if command[1:3] == ["workflow", "run"]:
            self.dispatched = True
        if command[1:3] == ["pr", "merge"]:
            self.head, self.branch = CANDIDATE, "main"
        if command[:2] == ["gh", "api"] and "/pulls?head=" in command[-1]:
            return self.done(command, [{"number": 3, "head": {"sha": "d" * 40}}])
        if command[:2] == ["gh", "api"] and "/check-runs" in command[-1]:
            return self.done(command, {"check_runs": [{"name": "Repository integrity", "status": "completed",
                                                       "conclusion": "success"}]})
        return self.done(command)

    def stream(self, command: list[str], cwd: Path) -> int:
        name = Path(command[1]).name if command[0] == sys.executable else Path(command[5]).name
        self.streamed.append(name)
        if name == "cut_release.py":
            (self.root / "VERSION").write_text("1.4.0\n")
            self.branch = "chore/cut-v1.4.0"
        elif name == "Certify-Excel.ps1":
            bundle = Path(command[command.index("-BundleDir") + 1])
            bundle.mkdir(parents=True)
            (bundle / "host.json").write_text("{}")
        elif name == "run_clean_room_pilot.py":
            work = Path(command[command.index("--work-dir") + 1])
            work.mkdir(parents=True)
            (work / "clean-room-journey.md").write_text("# journey\n")
            (work / "clean-room-journey.json").write_text(json.dumps({
                "template_sha": CANDIDATE, "pilot": "owner/vba-cleanroom-pilot-140", "profile": "library",
                "steps": [{"result": "PASS"}] * 8}))
        elif name == "prepare_release_evidence.py":
            work = Path(command[command.index("--work-dir") + 1])
            (work / "evidence").mkdir(parents=True)
            (work / "evidence/host.json").write_text("{}")
            (work / "release-evidence.json").write_text("{}")
        return 0


class ReleaseTemplateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        base = Path(self.temporary.name)
        self.root, self.work = base / "template", base / "work"
        (self.root / ".github").mkdir(parents=True)
        (self.root / "docs").mkdir()
        (self.root / ".github/repository-profile.json").write_text(json.dumps({"mode": "template", "repository": REPO}))
        (self.root / "VERSION").write_text("1.3.0\n")
        (self.root / "docs/REUSABLE_WORKFLOWS.md").write_text(
            f"| `1.0.0` | `{'0' * 40}` | first |\n| `1.0.1` | `{PROVIDER}` | second |\n")
        self.key = base / "release_signing"
        self.key.write_text("private")
        for item in (patch.object(time, "sleep", return_value=None),
                     patch.object(journey, "on_windows", return_value=True),
                     patch.object(excel, "evaluate", return_value={"status": "pass", "outcomes": ["PASS"]})):
            item.start()
            self.addCleanup(item.stop)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def main(self, world: World, *extra: str, answer: str = "y") -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        arguments = ["--root", str(self.root), "--version", "1.4.0", "--work-dir", str(self.work),
                     "--maintainer", "Maintainer", "--signing-key", str(self.key), *extra]
        with patch.object(evidence, "run", world.run), patch.object(journey, "stream", world.stream), \
                patch("builtins.input", return_value=answer), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = tool.main(arguments)
        return code, out.getvalue(), err.getvalue()

    def state(self) -> dict[str, Any]:
        value: dict[str, Any] = json.loads((self.work / "release-state.json").read_text())
        return value

    def test_full_release_runs_every_phase_in_order_then_resumes(self) -> None:
        world = World(self.root)
        code, out, err = self.main(world)
        self.assertEqual(code, 0, err)
        self.assertIn(f"RELEASED: v1.4.0 at {CANDIDATE}", out)
        self.assertEqual(world.streamed, ["cut_release.py", "Certify-Excel.ps1", "publish_wiki.py",
                                          "run_clean_room_pilot.py", "prepare_release_evidence.py",
                                          "publish_release.py"])
        state = self.state()
        self.assertEqual(state["done"], list(tool.PHASES))
        self.assertEqual(state["candidate"], CANDIDATE)
        self.assertTrue(state["wiki_browser_reviewed"] and state["redaction_reviewed"])
        self.assertIn("8 steps", state["pilot_summary"])
        self.assertTrue(world.dispatched)
        rerun = World(self.root)
        rerun.head = CANDIDATE
        code, out, _ = self.main(rerun)
        self.assertEqual(code, 0)
        self.assertEqual(out.count("SKIP"), len(tool.PHASES))
        self.assertEqual(rerun.streamed, [])

    def test_changed_called_workflow_stops_before_excel(self) -> None:
        world = World(self.root, workflow_changed=True)
        code, _, err = self.main(world)
        self.assertEqual(code, 1)
        self.assertIn("changed since interface 1.0.1", err)
        self.assertEqual(self.state()["done"], ["cut"])
        self.assertNotIn("Certify-Excel.ps1", world.streamed)

    def test_declined_attestations_and_moved_checkout(self) -> None:
        code, _, err = self.main(World(self.root), "--yes", answer="n")
        self.assertEqual(code, 1)
        self.assertIn("maintainer's request", err)
        self.assertEqual(self.state()["done"], ["cut", "interface", "excel"])
        moved = World(self.root)
        moved.head = "f" * 40
        self.assertIn("checkout moved", self.main(moved)[2])

    def test_existing_pilot_record_and_dirty_or_stale_start(self) -> None:
        record = self.work.parent / "journey" / "clean-room-journey.md"
        record.parent.mkdir()
        record.write_text("# journey\n")
        record.with_suffix(".json").write_text(json.dumps({"template_sha": "9" * 40, "pilot": "p",
                                                          "profile": "library", "steps": [{"result": "PASS"}]}))
        (self.root / "VERSION").write_text("1.4.0\n")
        world = World(self.root)
        code, _, err = self.main(world, "--pilot-record", str(record))
        self.assertEqual(code, 1)
        self.assertIn("belongs to another candidate", err)
        self.assertNotIn("run_clean_room_pilot.py", world.streamed)
        self.assertNotIn("cut_release.py", world.streamed)
        record.with_suffix(".json").write_text(json.dumps({"template_sha": BASE, "steps": [{"result": "FAIL"}]}))
        world = World(self.root)
        self.assertIn("failed or missing step", self.main(world, "--pilot-record", str(record))[2])

        (self.root / "VERSION").write_text("1.3.0\n")
        self.work = self.work.with_name("fresh")

        class Dirty(World):
            def run(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
                if command[:2] == ["git", "status"]:
                    return self.done(command, b" M x")
                return super().run(command, cwd)

        self.assertIn("uncommitted changes", self.main(Dirty(self.root))[2])
        stale = World(self.root)
        stale.head = "9" * 40
        self.work = self.work.with_name("fresh2")
        self.assertIn("not at origin/main", self.main(stale)[2])

    def test_state_guards_and_failing_tools(self) -> None:
        self.work.mkdir()
        (self.work / "release-state.json").write_text(json.dumps({"version": "9.9.9", "done": []}))
        self.assertIn("belongs to v9.9.9", self.main(World(self.root))[2])
        (self.root / ".github/repository-profile.json").write_text(json.dumps({"mode": "generated", "repository": REPO}))
        self.assertIn("canonical template only", self.main(World(self.root))[2])
        (self.root / ".github/repository-profile.json").write_text(json.dumps({"mode": "template", "repository": REPO}))
        self.work = self.work.with_name("fresh")

        class Failing(World):
            def stream(self, command: list[str], cwd: Path) -> int:
                super().stream(command, cwd)
                return 1 if "publish_wiki.py" in " ".join(command) else 0

        code, _, err = self.main(Failing(self.root))
        self.assertEqual(code, 1)
        self.assertIn("Wiki publication stopped (exit 1)", err)
        (self.root / "docs/REUSABLE_WORKFLOWS.md").write_text("no pins\n")
        self.work = self.work.with_name("fresh2")
        self.assertIn("no published interface pin", self.main(World(self.root))[2])

    def test_unavailable_prerequisites(self) -> None:
        with patch.object(journey, "on_windows", return_value=False):
            self.assertIn("needs Windows", self.main(World(self.root))[2])
        self.work = self.work.with_name("invalid")
        (self.root / "VERSION").write_text("1.3.0\n")
        with patch.object(excel, "evaluate", return_value={"status": "fail", "outcomes": ["EVIDENCE_INVALID"]}):
            self.assertIn("EVIDENCE_INVALID", self.main(World(self.root))[2])
        self.work = self.work.with_name("links")
        (self.root / "VERSION").write_text("1.3.0\n")
        with patch.object(journey, "POLL_LIMIT", 1):
            self.assertIn("timed out", self.main(World(self.root, links_green=False))[2])

        class Fetching(World):
            def run(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
                if command[:2] == ["git", "cat-file"]:
                    return self.done(command, code=1)
                return super().run(command, cwd)

        self.work = self.work.with_name("fetch")
        (self.root / "VERSION").write_text("1.3.0\n")
        fetching = Fetching(self.root)
        self.assertEqual(self.main(fetching)[0], 0)
        self.assertIn(["git", "fetch", "origin", PROVIDER], fetching.commands)


if __name__ == "__main__":
    unittest.main()
