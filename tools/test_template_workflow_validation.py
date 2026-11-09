#!/usr/bin/env python3
"""Template-maintainer unit tests for workflow contracts and consumer fixtures.

Covers the reusable-workflow consumer generator, CodeQL and Scorecard
eligibility guards, CodeQL language selection, and the Scorecard publication
contract implemented by the retained ``test_workflow_validation`` module, which
generated projects run with the pinned actionlint. Run with
``python3 tools/test_template_workflow_validation.py`` or ``unittest``.
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path
from unittest.mock import patch

import create_reusable_workflow_fixture as reusable_fixture
import initialize_repository as initializer
import test_workflow_validation as retained_workflow

WORKFLOW_SHA = "a" * 40


class ReusableWorkflowFixtureTests(unittest.TestCase):
    """Offline behavior suite for disposable reusable-workflow consumer generation."""

    def test_invalid_sha_is_rejected_before_destination_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary) / "consumer"
            with patch.object(initializer, "copy_fixture") as copy_fixture:
                with self.assertRaisesRegex(
                    ValueError, "workflow-sha must be a full lowercase commit SHA"
                ):
                    reusable_fixture.create_fixture(
                        Path("source"), destination, "library", "A" * 40
                    )
            copy_fixture.assert_not_called()
            self.assertFalse(destination.exists())

    def test_existing_destination_is_rejected_without_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary) / "consumer"
            destination.mkdir()
            marker = destination / "marker.txt"
            marker.write_text("preserve\n", encoding="utf-8")
            with patch.object(initializer, "copy_fixture") as copy_fixture:
                with self.assertRaisesRegex(ValueError, "destination must not exist"):
                    reusable_fixture.create_fixture(
                        Path("source"), destination, "library", WORKFLOW_SHA
                    )
            copy_fixture.assert_not_called()
            self.assertEqual(marker.read_text(encoding="utf-8"), "preserve\n")

    def test_each_profile_writes_exact_sha_consumer_and_commits(self) -> None:
        for profile in initializer.SUPPORTED_PROFILES:
            with self.subTest(profile=profile), tempfile.TemporaryDirectory() as temporary:
                source = Path(temporary) / "source"
                destination = Path(temporary) / "consumer"

                def copy_fixture(_source: Path, target: Path) -> None:
                    (target / ".github/workflows").mkdir(parents=True)

                scalars = ["PROJECT_NAME=Fixture"]
                repeatable = ["KNOWN_LIMITATION=Fixture"]
                changes = {"README.md": b"fixture\n"}

                with (
                    patch.object(
                        initializer,
                        "copy_fixture",
                        side_effect=copy_fixture,
                    ) as copy_mock,
                    patch.object(
                        initializer,
                        "fixture_arguments",
                        return_value=(scalars, repeatable),
                    ) as arguments_mock,
                    patch.object(
                        initializer,
                        "build_changes",
                        return_value=(changes, {"status": "ready"}),
                    ) as build_mock,
                    patch.object(
                        initializer, "apply_changes"
                    ) as apply_mock,
                    patch.object(reusable_fixture, "git_text") as git_mock,
                ):
                    reusable_fixture.create_fixture(
                        source, destination, profile, WORKFLOW_SHA
                    )

                copy_mock.assert_called_once_with(source, destination)
                arguments_mock.assert_called_once_with(profile)
                build_mock.assert_called_once_with(destination, profile, scalars, repeatable)
                apply_mock.assert_called_once_with(destination, changes)
                caller = (destination / ".github/workflows/static-checks.yml").read_text(
                    encoding="utf-8"
                )
                canonical_repo = "".join(("danielep71/EXCEL-VBA-", "PROJECT-TEMPLATE"))
                self.assertIn(
                    f"uses: {canonical_repo}/.github/workflows/static-checks.yml@{WORKFLOW_SHA} # v1.0.0",
                    caller,
                )
                self.assertIn(f"expected-profile: {profile}", caller)
                self.assertIn("CHECKED_SHA", caller)
                self.assertIn("test -f tests/modules/ProjectTests.bas", caller)
                self.assertEqual(git_mock.call_args_list[-2].args, (destination, "add", "--all"))
                self.assertEqual(
                    git_mock.call_args_list[-1].args,
                    (
                        destination,
                        "commit",
                        "-m",
                        f"Create {profile} reusable workflow consumer",
                    ),
                )

    def test_cli_resolves_paths_and_delegates(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "source"
            destination = Path(temporary) / "consumer"
            argv = [
                "create_reusable_workflow_fixture.py",
                "--root",
                str(root),
                "--destination",
                str(destination),
                "--profile",
                "library",
                "--workflow-sha",
                WORKFLOW_SHA,
            ]
            with patch.object(sys, "argv", argv), patch.object(
                reusable_fixture, "create_fixture"
            ) as create_mock:
                self.assertEqual(reusable_fixture.main(), 0)
            create_mock.assert_called_once_with(
                root.resolve(), destination.resolve(), "library", WORKFLOW_SHA
            )



class SupplyChainEligibilityTests(unittest.TestCase):
    """Keep private/public analyzer eligibility explicit and least-privileged."""

    def setUp(self) -> None:
        root = Path(__file__).resolve().parents[1]
        self.codeql = (root / ".github/workflows/codeql.yml").read_text(encoding="utf-8")
        self.scorecard = (root / ".github/workflows/scorecard.yml").read_text(encoding="utf-8")

    @staticmethod
    def scorecard_private_guarded(text: str) -> bool:
        return (
            text.count("github.event.repository.private == false") >= 2
            and "publish_results: true" in text
            and "verify-publication:" in text
            and "needs: published" in text
            and "pull_request:" not in text
        )

    @staticmethod
    def codeql_eligibility_guarded(text: str) -> bool:
        return (
            "github.event.repository.private == false || vars.ENABLE_PRIVATE_CODEQL == 'true'" in text
            and "needs: eligibility" in text
            and "upload: never" in text
            and "upload-database: false" in text
            and "pull_request_target:" not in text
        )

    def test_current_visibility_and_capability_contract(self) -> None:
        self.assertTrue(self.scorecard_private_guarded(self.scorecard))
        self.assertTrue(self.codeql_eligibility_guarded(self.codeql))

    def test_private_scorecard_publication_guard_removal_is_detected(self) -> None:
        changed = self.scorecard.replace(
            "      github.event.repository.private == false &&\n",
            "",
            1,
        )
        self.assertFalse(self.scorecard_private_guarded(changed))

    def test_private_codeql_opt_in_guard_removal_is_detected(self) -> None:
        changed = self.codeql.replace(
            "    if: github.event.repository.private == false || vars.ENABLE_PRIVATE_CODEQL == 'true'\n",
            "    if: always()\n",
            1,
        )
        self.assertFalse(self.codeql_eligibility_guarded(changed))

    def test_pr_codeql_privilege_expansion_is_detected(self) -> None:
        changed = self.codeql.replace(
            "      contents: read\n",
            "      contents: read\n      security-events: write\n",
            1,
        ).replace("          upload: never", "          upload: always", 1)
        self.assertFalse(self.codeql_eligibility_guarded(changed))


class CodeqlLanguageSelectionTests(unittest.TestCase):
    """CodeQL analyzes exactly the languages the checked-out tree tracks.

    CodeQL fails a language with no source, so a fixed matrix broke every
    minimal-tier project once the tier removed the label scripts.
    """

    def setUp(self) -> None:
        root = Path(__file__).resolve().parents[1]
        self.codeql = (root / ".github/workflows/codeql.yml").read_text(encoding="utf-8")

    def selection_script(self) -> str:
        match = re.search(r"        id: languages\n        shell: bash\n        run: \|\n((?:          .*\n|\n)+)", self.codeql)
        self.assertIsNotNone(match, "codeql.yml must select languages in a step with id 'languages'")
        assert match is not None
        return textwrap.dedent(match.group(1))

    def selected(self, files: tuple[str, ...]) -> str:
        with tempfile.TemporaryDirectory() as temporary:
            tree = Path(temporary)
            subprocess.run(["git", "init", "-q", str(tree)], check=True)
            for name in files:
                (tree / name).parent.mkdir(parents=True, exist_ok=True)
                (tree / name).write_text("x = 1\n", encoding="utf-8")
            subprocess.run(["git", "-C", str(tree), "add", "--all"], check=True)
            output = tree / "github-output"
            subprocess.run(
                ["bash", "-c", self.selection_script()],
                cwd=tree,
                env={**os.environ, "GITHUB_OUTPUT": str(output)},
                check=True,
                capture_output=True,
            )
            return output.read_text(encoding="utf-8").strip()

    def test_languages_follow_tracked_sources(self) -> None:
        self.assertEqual(self.selected(("tools/a.py", ".github/scripts/b.mjs")),
                         'languages=["python","javascript-typescript"]')
        self.assertEqual(self.selected(("tools/a.py",)), 'languages=["python"]')
        self.assertEqual(self.selected(("README.md",)), "languages=[]")

    def test_both_analysis_jobs_use_the_selection(self) -> None:
        self.assertEqual(self.codeql.count("language: ${{ fromJSON(needs.eligibility.outputs.languages) }}"), 2)
        self.assertEqual(self.codeql.count("needs.eligibility.outputs.languages != '[]'"), 2)
        self.assertNotIn("          - javascript-typescript", self.codeql)


class ScorecardReviewRegressionTests(unittest.TestCase):
    """Hold the retained Scorecard publication contract: job environment and
    defaults, the file-mode input binding, and the workflow's SVG badge script."""

    def setUp(self) -> None:
        self.workflow = (Path(__file__).resolve().parents[1] / retained_workflow.SCORECARD_WORKFLOW).read_text(
            encoding="utf-8"
        )

    def check_contract(self, workflow: str) -> tuple[str, list[str]]:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            retained_workflow.write_fixture_text(root, retained_workflow.SCORECARD_WORKFLOW, workflow)
            return retained_workflow.scorecard_publication_contract(root)

    def test_current_workflow(self) -> None:
        self.assertEqual(self.check_contract(self.workflow), ("PASS", []))

    def test_flow_environment_and_defaults_are_rejected(self) -> None:
        for key in (
            "env", "'env'", '"env"', "defaults", "'defaults'", '"defaults"',
            r'"e\u006ev"', r'"defa\u0075lts"', r'"\x65nv"',
        ):
            for indent, boundary in (("", "jobs:\n"), ("    ", "  published:\n")):
                with self.subTest(key=key, indent=indent):
                    entry = f"{indent}{key}: {{FOO: bar}}\n"
                    changed = self.workflow.replace(
                        boundary, entry + boundary if not indent else boundary + entry, 1
                    )
                    self.assertEqual(self.check_contract(changed)[0], "FAIL")

    def test_file_mode_is_bound_to_scorecard_input(self) -> None:
        for replacement in (
            "          file_mode: archive # file_mode: git",
            "          # file_mode: git",
            "          other: git # file_mode: git",
            "        env:\n          file_mode: git",
        ):
            with self.subTest(replacement=replacement):
                changed = self.workflow.replace("          file_mode: git", replacement)
                self.assertEqual(self.check_contract(changed)[0], "FAIL")
        changed = self.workflow.replace("          file_mode: git", "          file_mode: archive")
        changed = changed.replace("          fetch-depth: 0", "          fetch-depth: 0\n          file_mode: git")
        self.assertEqual(self.check_contract(changed)[0], "FAIL")
        for value in ("git # documented", "'git'", '"git"'):
            with self.subTest(value=value):
                changed = self.workflow.replace("file_mode: git", f"file_mode: {value}")
                self.assertEqual(self.check_contract(changed), ("PASS", []))

    def test_actual_badge_script_requires_safe_svg_document(self) -> None:
        scripts = self.workflow.split("if python3 - <<'PY'\n")[1:]
        script = next(part.split("\n          PY", 1)[0] for part in scripts
                      if 'badge = Path("public-scorecard-badge.svg")' in part)
        code = textwrap.dedent(script)
        cases = (
            ('<svg xmlns="http://www.w3.org/2000/svg"><text>8.0</text></svg>', True),
            ('<?xml version="1.0"?><svg/>', True),
            ('<html><body><svg/></body></html>', False),
            ('<svg>', False),
            ('', False),
            ('<!DOCTYPE svg [<!ENTITY x "test">]><svg>&x;</svg>', False),
            ('<svg><text>invalid repo path</text></svg>', False),
        )
        for badge, expected in cases:
            with self.subTest(badge=badge), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                (root / "public-scorecard-badge.svg").write_text(badge, encoding="utf-8")
                result = subprocess.run(
                    [sys.executable, "-c", code], cwd=root,
                    capture_output=True, text=True, check=False,
                )
                self.assertEqual(result.returncode == 0, expected, result.stderr)


class LabelReconcileBranchTests(unittest.TestCase):
    """Live label reconciliation runs only outside pull requests and only on the
    default branch, so a dispatch on another branch cannot apply its catalogue."""

    def setUp(self) -> None:
        self.workflow = (Path(__file__).resolve().parents[1] / retained_workflow.LABEL_SYNC_WORKFLOW).read_text(
            encoding="utf-8"
        )

    def check_contract(self, workflow: str | None) -> tuple[str, list[str]]:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            if workflow is not None:
                retained_workflow.write_fixture_text(root, retained_workflow.LABEL_SYNC_WORKFLOW, workflow)
            return retained_workflow.label_reconcile_contract(root)

    def with_condition(self, condition: str) -> str:
        changed, count = re.subn(r"(?m)^(  reconcile:\n(?:    .*\n)*?)    if: .*$", rf"\g<1>    if: {condition}",
                                 self.workflow, count=1)
        self.assertEqual(count, 1)
        return changed

    def test_current_workflow(self) -> None:
        self.assertEqual(self.check_contract(self.workflow), ("PASS", []))

    def test_any_branch_is_rejected(self) -> None:
        status, failures = self.check_contract(self.with_condition("github.event_name != 'pull_request'"))
        self.assertEqual(status, "FAIL")
        self.assertIn("default_branch", failures[0])

    def test_weakened_conditions_are_rejected(self) -> None:
        default = "github.ref == format('refs/heads/{0}', github.event.repository.default_branch)"
        for condition in (
            default,
            f"github.event_name != 'pull_request' || {default}",
            "github.event_name != 'pull_request' && github.ref == 'refs/heads/main'",
            "github.event_name != 'pull_request' && always()",
        ):
            with self.subTest(condition=condition):
                self.assertEqual(self.check_contract(self.with_condition(condition))[0], "FAIL")

    def test_equivalent_forms_are_accepted(self) -> None:
        default = "github.ref == format('refs/heads/{0}', github.event.repository.default_branch)"
        for condition in (
            f"${{{{ github.event_name != 'pull_request' && {default} }}}}",
            f"{default} && github.event_name != 'pull_request'",
        ):
            with self.subTest(condition=condition):
                self.assertEqual(self.check_contract(self.with_condition(condition)), ("PASS", []))

    def test_missing_job_or_workflow(self) -> None:
        self.assertEqual(self.check_contract(None), ("N/A", []))
        renamed = self.workflow.replace("  reconcile:\n", "  apply:\n")
        self.assertEqual(self.check_contract(renamed)[0], "FAIL")


if __name__ == "__main__":
    unittest.main()
