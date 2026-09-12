"""Offline behavioral tests for disposable reusable-workflow consumer generation."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import create_reusable_workflow_fixture as fixture

SHA = "a" * 40


class ReusableWorkflowFixtureTests(unittest.TestCase):
    def test_invalid_sha_is_rejected_before_destination_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary) / "consumer"
            with patch.object(fixture.initializer, "_copy_fixture") as copy_fixture:
                with self.assertRaisesRegex(
                    ValueError, "workflow-sha must be a full lowercase commit SHA"
                ):
                    fixture.create_fixture(Path("source"), destination, "library", "A" * 40)
            copy_fixture.assert_not_called()
            self.assertFalse(destination.exists())

    def test_existing_destination_is_rejected_without_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary) / "consumer"
            destination.mkdir()
            marker = destination / "marker.txt"
            marker.write_text("preserve\n", encoding="utf-8")
            with patch.object(fixture.initializer, "_copy_fixture") as copy_fixture:
                with self.assertRaisesRegex(ValueError, "destination must not exist"):
                    fixture.create_fixture(Path("source"), destination, "library", SHA)
            copy_fixture.assert_not_called()
            self.assertEqual(marker.read_text(encoding="utf-8"), "preserve\n")

    def test_each_profile_writes_exact_sha_consumer_and_commits(self) -> None:
        for profile in fixture.initializer.SUPPORTED_PROFILES:
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
                        fixture.initializer, "_copy_fixture", side_effect=copy_fixture
                    ) as copy_mock,
                    patch.object(
                        fixture.initializer,
                        "_fixture_arguments",
                        return_value=(scalars, repeatable),
                    ) as arguments_mock,
                    patch.object(
                        fixture.initializer,
                        "_build_changes",
                        return_value=(changes, {"status": "ready"}),
                    ) as build_mock,
                    patch.object(fixture.initializer, "_apply_changes") as apply_mock,
                    patch.object(fixture.initializer, "_git") as git_mock,
                ):
                    fixture.create_fixture(source, destination, profile, SHA)

                copy_mock.assert_called_once_with(source, destination)
                arguments_mock.assert_called_once_with(profile)
                build_mock.assert_called_once_with(destination, profile, scalars, repeatable)
                apply_mock.assert_called_once_with(destination, changes)
                caller = (destination / ".github/workflows/static-checks.yml").read_text(
                    encoding="utf-8"
                )
                canonical_repo = "".join(("danielep71/EXCEL-VBA-", "PROJECT-TEMPLATE"))
                self.assertIn(
                    f"uses: {canonical_repo}/.github/workflows/static-checks.yml@{SHA} # v1.0.0",
                    caller,
                )
                self.assertIn(f"expected-profile: {profile}", caller)
                self.assertIn("CHECKED_SHA", caller)
                self.assertIn("test -f tests/modules/ProjectTests.bas", caller)
                self.assertEqual(
                    git_mock.call_args_list[-2].args,
                    (destination, "add", "--all"),
                )
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
                SHA,
            ]
            with patch.object(sys, "argv", argv), patch.object(
                fixture, "create_fixture"
            ) as create_mock:
                self.assertEqual(fixture.main(), 0)
            create_mock.assert_called_once_with(
                root.resolve(), destination.resolve(), "library", SHA
            )


if __name__ == "__main__":
    unittest.main()
