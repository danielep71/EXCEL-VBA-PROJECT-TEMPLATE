#!/usr/bin/env python3
"""Fixtures for cut_release.py: CHANGELOG rewriting offline, the branch flow in a real temporary repository."""

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

import cut_release as tool  # noqa: E402
import prepare_release_evidence as evidence  # noqa: E402

REPO = "owner/repo"
TEMPLATE = """# Changelog

<a id="unreleased"></a>

## [Unreleased]

### Added

- New thing.

## [1.3.0] - 2026-11-14

### Fixed

- Old thing.

[Unreleased]: https://github.com/owner/repo/compare/v1.3.0...HEAD
[1.3.0]: https://github.com/owner/repo/compare/v1.2.1...v1.3.0
"""
GENERATED = """# Changelog

## [Unreleased]

<!-- Add only user-visible changes made in this generated project. -->

### Added

- First feature.

---

<!--
Release procedure: keep this comment.
-->
"""


class ChangelogTests(unittest.TestCase):
    def test_template_cut_moves_entries_and_links(self) -> None:
        text = tool.cut_changelog(TEMPLATE, "1.4.0", "2027-02-01", REPO)
        self.assertIn("## [Unreleased]\n\n## [1.4.0] - 2027-02-01\n\n### Added\n\n- New thing.\n\n## [1.3.0]", text)
        self.assertIn("[Unreleased]: https://github.com/owner/repo/compare/v1.4.0...HEAD\n"
                      "[1.4.0]: https://github.com/owner/repo/compare/v1.3.0...v1.4.0\n"
                      "[1.3.0]: https://github.com/owner/repo/compare/v1.2.1...v1.3.0\n", text)

    def test_first_generated_release_links_its_tag_and_keeps_guidance(self) -> None:
        text = tool.cut_changelog(GENERATED, "0.1.0", "2026-11-14", REPO)
        self.assertIn("## [Unreleased]\n\n<!-- Add only user-visible changes made in this generated project. -->\n\n"
                      "## [0.1.0] - 2026-11-14\n\n### Added\n\n- First feature.\n\n---", text)
        self.assertIn("Release procedure: keep this comment.", text)
        self.assertTrue(text.endswith("[Unreleased]: https://github.com/owner/repo/compare/v0.1.0...HEAD\n"
                                      "[0.1.0]: https://github.com/owner/repo/releases/tag/v0.1.0\n"))

    def test_rejections(self) -> None:
        empty = GENERATED.replace("### Added\n\n- First feature.\n", "No unreleased changes recorded.\n")
        cases = ((TEMPLATE, "1.3.0", "already has a section"), (TEMPLATE, "1.2.9", "not newer than 1.3.0"),
                 (empty, "0.1.0", "no entries to release"), ("# Changelog\n", "0.1.0", r"no ## \[Unreleased\]"),
                 (TEMPLATE, "1.4", "not a MAJOR.MINOR.PATCH"))
        for text, version, message in cases:
            with self.subTest(message=message), self.assertRaisesRegex(evidence.EvidenceError, message):
                tool.cut_changelog(text, version, "2027-02-01", REPO)


class CutCommandTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        (self.root / ".github").mkdir()
        (self.root / "tools").mkdir()
        (self.root / ".github/repository-profile.json").write_text(json.dumps({"repository": REPO}))
        (self.root / "CHANGELOG.md").write_text(TEMPLATE)
        (self.root / "VERSION").write_text("1.3.0\n")
        semantics = "import sys\nsys.exit(0 if open('VERSION').read().strip() != '9.9.9' else 1)\n"
        (self.root / "tools/check_release_semantics.py").write_text(semantics)
        for command in (["init", "-q", "-b", "main"], ["add", "--all"],
                        ["-c", "user.name=Fixture", "-c", "user.email=fixture@example.invalid", "commit", "-qm", "base"]):
            subprocess.run(["git", "-C", str(self.root), *command], check=True)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def main(self, *arguments: str) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        environment = {"GIT_AUTHOR_NAME": "Fixture", "GIT_AUTHOR_EMAIL": "fixture@example.invalid",
                       "GIT_COMMITTER_NAME": "Fixture", "GIT_COMMITTER_EMAIL": "fixture@example.invalid"}
        with patch.dict("os.environ", environment), contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = tool.main(["--root", str(self.root), *arguments])
        return code, out.getvalue(), err.getvalue()

    def git(self, *arguments: str) -> str:
        return subprocess.run(["git", "-C", str(self.root), *arguments], check=True, capture_output=True,
                              text=True).stdout.strip()

    def test_cut_commits_on_a_branch_and_checks_semantics(self) -> None:
        code, out, err = self.main("--version", "1.4.0", "--date", "2027-02-01")
        self.assertEqual(code, 0, err)
        self.assertIn("committed locally", out)
        self.assertIn("git push --set-upstream origin chore/cut-v1.4.0", out)
        self.assertIn("gh pr create --base main --head chore/cut-v1.4.0", out)
        self.assertEqual(self.git("branch", "--show-current"), "chore/cut-v1.4.0")
        self.assertEqual(self.git("log", "-1", "--format=%s"), "Cut v1.4.0 release candidate")
        self.assertEqual(self.git("show", "HEAD:VERSION"), "1.4.0")
        self.assertEqual(self.git("status", "--porcelain"), "")

    def test_dirty_tree_and_failing_semantics(self) -> None:
        (self.root / "VERSION").write_text("dirty\n")
        self.assertIn("uncommitted changes", self.main("--version", "1.4.0")[2])
        self.git("checkout", "--", "VERSION")
        code, _, err = self.main("--version", "9.9.9")
        self.assertEqual(code, 1)
        self.assertIn("branch is kept for inspection", err)
        (self.root / ".github/repository-profile.json").write_text("{}")
        self.git("checkout", "-q", "main")
        self.git("-c", "user.name=F", "-c", "user.email=f@example.invalid", "commit", "-qam", "drop repository")
        self.assertIn("names no repository", self.main("--version", "1.4.0")[2])

    def test_push_opens_the_pull_request(self) -> None:
        calls: list[list[str]] = []
        real = evidence.run

        def recording(command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
            if command[0] == "gh" or command[:2] == ["git", "push"]:
                calls.append(command)
                return subprocess.CompletedProcess(command, 0, b"", b"")
            return real(command, cwd)

        with patch.object(evidence, "run", recording):
            code, out, err = self.main("--version", "1.4.0", "--date", "2027-02-01", "--push")
        self.assertEqual(code, 0, err)
        self.assertEqual(calls[0], ["git", "push", "--set-upstream", "origin", "chore/cut-v1.4.0"])
        self.assertEqual(calls[1][:7], ["gh", "pr", "create", "--base", "main", "--head", "chore/cut-v1.4.0"])
        self.assertIn("pull request opened", out)


if __name__ == "__main__":
    unittest.main()
