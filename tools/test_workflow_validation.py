#!/usr/bin/env python3
"""Exercise the pinned authoritative GitHub Actions validator.

The portable repository checker remains dependency-free. This hosted companion
proves that the selected actionlint binary accepts the tracked workflows and
rejects syntax, schema, duplicate-key, and local-action defects that a small
portable YAML subset parser cannot authoritatively decide.
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import tempfile
import textwrap
import unittest
from dataclasses import dataclass
from pathlib import Path
from unittest.mock import patch

import create_reusable_workflow_fixture as reusable_fixture

EXPECTED_ACTIONLINT_VERSION = "1.7.12"
WORKFLOW_SHA = "a" * 40
SCORECARD_WORKFLOW = ".github/workflows/scorecard.yml"
SCORECARD_APPROVED_ACTIONS = {
    "actions/checkout",
    "actions/upload-artifact",
    "github/codeql-action/upload-sarif",
    "ossf/scorecard-action",
    "step-security/harden-runner",
}


@dataclass(frozen=True)
class Fixture:
    name: str
    workflow: str
    expected_pattern: str | None
    files: tuple[tuple[str, str], ...] = ()


FIXTURES = (
    Fixture(
        "valid-local-action",
        """name: Valid local action
on: push
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: ./.github/actions/example
""",
        None,
        (
            (
                ".github/actions/example/action.yml",
                """name: Example
description: Valid composite fixture
runs:
  using: composite
  steps:
    - shell: bash
      run: echo fixture
""",
            ),
        ),
    ),
    Fixture(
        "invalid-yaml",
        """name: Invalid YAML
on: [push
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - run: echo fixture
""",
        r"could not parse as YAML",
    ),
    Fixture(
        "duplicate-job",
        """name: Duplicate job
on: push
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - run: echo first
  test:
    runs-on: ubuntu-latest
    steps:
      - run: echo second
""",
        r"duplicated",
    ),
    Fixture(
        "invalid-job-structure",
        """name: Invalid job structure
on: push
jobs:
  test:
    runs-on: ubuntu-latest
    unsupported-key: true
    steps:
      - run: echo fixture
""",
        r"unexpected key \"unsupported-key\"",
    ),
    Fixture(
        "malformed-local-action",
        """name: Malformed local action
on: push
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: ./.github/actions/example
""",
        r"runs\.using.*missing",
        (
            (
                ".github/actions/example/action.yml",
                """name: Example
description: Missing runs metadata
""",
            ),
        ),
    ),
    Fixture(
        "missing-local-entrypoint",
        """name: Missing local entrypoint
on: push
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: ./.github/actions/example
""",
        r"file \"dist/index\.js\" does not exist",
        (
            (
                ".github/actions/example/action.yml",
                """name: Example
description: Missing entrypoint fixture
runs:
  using: node24
  main: dist/index.js
""",
            ),
        ),
    ),
)


class ReusableWorkflowFixtureTests(unittest.TestCase):
    """Offline behavior suite for disposable reusable-workflow consumer generation."""

    def test_invalid_sha_is_rejected_before_destination_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary) / "consumer"
            with patch.object(reusable_fixture.initializer, "_copy_fixture") as copy_fixture:
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
            with patch.object(reusable_fixture.initializer, "_copy_fixture") as copy_fixture:
                with self.assertRaisesRegex(ValueError, "destination must not exist"):
                    reusable_fixture.create_fixture(
                        Path("source"), destination, "library", WORKFLOW_SHA
                    )
            copy_fixture.assert_not_called()
            self.assertEqual(marker.read_text(encoding="utf-8"), "preserve\n")

    def test_each_profile_writes_exact_sha_consumer_and_commits(self) -> None:
        for profile in reusable_fixture.initializer.SUPPORTED_PROFILES:
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
                        reusable_fixture.initializer,
                        "_copy_fixture",
                        side_effect=copy_fixture,
                    ) as copy_mock,
                    patch.object(
                        reusable_fixture.initializer,
                        "_fixture_arguments",
                        return_value=(scalars, repeatable),
                    ) as arguments_mock,
                    patch.object(
                        reusable_fixture.initializer,
                        "_build_changes",
                        return_value=(changes, {"status": "ready"}),
                    ) as build_mock,
                    patch.object(
                        reusable_fixture.initializer, "_apply_changes"
                    ) as apply_mock,
                    patch.object(reusable_fixture.initializer, "_git") as git_mock,
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


class ScorecardReviewRegressionTests(unittest.TestCase):
    """Exercise the three late review findings against actual retained policy."""

    def setUp(self) -> None:
        self.workflow = (Path(__file__).resolve().parents[1] / SCORECARD_WORKFLOW).read_text(
            encoding="utf-8"
        )

    def check_contract(self, workflow: str) -> tuple[str, list[str]]:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            write_fixture_text(root, SCORECARD_WORKFLOW, workflow)
            return scorecard_publication_contract(root)

    def test_current_workflow(self) -> None:
        self.assertEqual(self.check_contract(self.workflow), ("PASS", []))

    def test_flow_environment_and_defaults_are_rejected(self) -> None:
        for key in ("env", "'env'", '"env"', "defaults", "'defaults'", '"defaults"'):
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


def run_actionlint(
    executable: Path, root: Path, workflows: list[Path]
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [str(executable), "-no-color", *(str(path) for path in workflows)],
        cwd=root,
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )


def write_fixture_text(root: Path, relative: str, content: str) -> Path:
    path = root / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")
    return path


def scorecard_git_input(published: str) -> bool:
    """Bind the input to the Scorecard step's block-style with mapping.

    Unsupported layouts fail closed; actionlint remains the YAML syntax owner.
    Comments and inputs on other steps cannot satisfy this policy.
    """
    steps = re.split(r"(?m)^      - ", published)
    scorecard_steps = [
        step for step in steps
        if re.search(r"(?m)^        uses: ossf/scorecard-action@[^\s]+(?:\s+#.*)?$", step)
    ]
    if len(scorecard_steps) != 1:
        return False
    mappings = re.findall(
        r"(?ms)^        with:[ \t]*(?:#[^\n]*)?\n(.*?)(?=^        \S|\Z)",
        scorecard_steps[0],
    )
    if len(mappings) != 1:
        return False
    values = re.findall(
        r"(?m)^          (?:file_mode|'file_mode'|\"file_mode\"):[ \t]*([^\n]*)$",
        mappings[0],
    )
    return len(values) == 1 and re.fullmatch(
        r"(?:git|'git'|\"git\")[ \t]*(?:#.*)?", values[0]
    ) is not None


def scorecard_publication_contract(root: Path) -> tuple[str, list[str]]:
    path = root / SCORECARD_WORKFLOW
    if not path.is_file():
        return "N/A", []
    text = path.read_text(encoding="utf-8")
    if "publish_results: true" not in text:
        return "N/A", []

    failures: list[str] = []
    if re.search(r"(?m)^(?:env|'env'|\"env\")[ \t]*:", text):
        failures.append("Scorecard publication workflow must not define top-level env")
    if re.search(r"(?m)^(?:defaults|'defaults'|\"defaults\")[ \t]*:", text):
        failures.append("Scorecard publication workflow must not define top-level defaults")

    match = re.search(
        r"(?ms)^  published:\n(?P<body>.*?)(?=^  [A-Za-z0-9_-]+:\n|\Z)",
        text,
    )
    if match is None:
        failures.append("Scorecard publication workflow has no published job")
        return "FAIL", failures

    published = match.group("body")
    if re.search(r"(?m)^    (?:env|'env'|\"env\")[ \t]*:", published):
        failures.append("Scorecard published job must not define job-level env")
    if re.search(r"(?m)^    (?:defaults|'defaults'|\"defaults\")[ \t]*:", published):
        failures.append("Scorecard published job must not define job-level defaults")
    if re.search(r"(?m)^        run:", published):
        failures.append("Scorecard published job may use only OpenSSF-approved actions")

    for action in re.findall(r"(?m)^        uses:\s*([^\s#]+)", published):
        owner_action = action.split("@", 1)[0]
        if owner_action not in SCORECARD_APPROVED_ACTIONS:
            failures.append(
                f"Scorecard published job uses unsupported action {owner_action}"
            )

    if text.count("id-token: write") != 1 or "id-token: write" not in published:
        failures.append(
            "Scorecard publication must grant id-token: write only to the published job"
        )
    if not scorecard_git_input(published):
        failures.append(
            "Scorecard published job must enumerate the checked repository through git"
        )
    if "verify-publication:" not in text:
        failures.append("Scorecard workflow must verify the public result after publication")
    expected_url = (
        "https://api.scorecard.dev/projects/github.com/"
        + "${GITHUB_REPOSITORY}?commit=${GITHUB_SHA}"
    )
    if expected_url not in text:
        failures.append("Scorecard public verification must bind to the exact GitHub SHA")
    badge_url = (
        "https://api.scorecard.dev/projects/github.com/"
        + "${GITHUB_REPOSITORY}/badge"
    )
    if badge_url not in text or "invalid repo path" not in text:
        failures.append("Scorecard public verification must validate the rendered badge surface")

    return ("PASS" if not failures else "FAIL"), failures


def build_report(root: Path, executable: Path) -> tuple[str, list[str]]:
    failures: list[str] = []
    version = subprocess.run(
        [str(executable), "-version"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    observed_version = version.stdout.splitlines()[0].strip() if version.stdout else ""
    if version.returncode != 0 or observed_version != EXPECTED_ACTIONLINT_VERSION:
        failures.append(
            f"validator version must be {EXPECTED_ACTIONLINT_VERSION}; "
            f"observed {observed_version or 'unavailable'}"
        )

    workflow_directory = root / ".github/workflows"
    workflows = sorted(
        path
        for path in workflow_directory.iterdir()
        if path.is_file() and path.suffix.casefold() in {".yml", ".yaml"}
    )
    current = run_actionlint(executable, root, workflows)
    if current.returncode != 0:
        failures.append("tracked workflows failed authoritative validation:\n" + current.stdout.strip())

    scorecard_status, scorecard_failures = scorecard_publication_contract(root)
    failures.extend(scorecard_failures)

    fixture_rows: list[tuple[str, str, str]] = []
    with tempfile.TemporaryDirectory(prefix="workflow-validation-") as temporary:
        fixture_root = Path(temporary)
        for fixture in FIXTURES:
            case_root = fixture_root / fixture.name
            case_root.mkdir(parents=True)
            initialized = subprocess.run(
                ["git", "init", "--quiet"],
                cwd=case_root,
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
            if initialized.returncode != 0:
                failures.append(
                    f"fixture {fixture.name!r} Git initialization failed:\n"
                    + initialized.stdout.strip()
                )
            workflow = write_fixture_text(
                case_root,
                f".github/workflows/{fixture.name}.yml",
                fixture.workflow,
            )
            for relative, content in fixture.files:
                write_fixture_text(case_root, relative, content)
            completed = run_actionlint(executable, case_root, [workflow])
            if fixture.expected_pattern is None:
                passed = completed.returncode == 0
                expectation = "accepted"
            else:
                passed = completed.returncode != 0 and re.search(
                    fixture.expected_pattern,
                    completed.stdout,
                    re.IGNORECASE | re.DOTALL,
                ) is not None
                expectation = "rejected"
            fixture_rows.append((fixture.name, expectation, "PASS" if passed else "FAIL"))
            if not passed:
                failures.append(
                    f"fixture {fixture.name!r} was not {expectation} as expected:\n"
                    + completed.stdout.strip()
                )

    lines = [
        "# Authoritative workflow validation",
        "",
        f"- Validator: **actionlint {observed_version or 'unavailable'}**",
        f"- Tracked workflows: **{len(workflows)}**",
        f"- Fixtures: **{len(FIXTURES)}**",
        "",
        "| Check | Expected | Result |",
        "| --- | --- | --- |",
        (
            f"| Current tracked workflows | accepted | "
            f"{'PASS' if current.returncode == 0 else 'FAIL'} |"
        ),
        f"| Scorecard publication contract | accepted | {scorecard_status} |",
    ]
    lines.extend(
        f"| {name} | {expectation} | {result} |" for name, expectation, result in fixture_rows
    )
    lines.extend(
        [
            "",
            (
                "PASS: authoritative workflow and local-action fixtures satisfied."
                if not failures
                else f"FAIL: {len(failures)} authoritative validation failure(s)."
            ),
            "",
        ]
    )
    return "\n".join(lines), failures


def parse_arguments(arguments: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--actionlint", type=Path, required=True)
    parser.add_argument("--summary", type=Path)
    return parser.parse_args(arguments)


def main(arguments: list[str] | None = None) -> int:
    options = parse_arguments(sys.argv[1:] if arguments is None else arguments)
    root = options.root.resolve()
    executable = options.actionlint.resolve()
    if not executable.is_file() or not os.access(executable, os.X_OK):
        print(f"ERROR: actionlint executable is unavailable: {executable}", file=sys.stderr)
        return 2
    report, failures = build_report(root, executable)
    print(report, end="")
    if options.summary:
        options.summary.parent.mkdir(parents=True, exist_ok=True)
        options.summary.write_text(report, encoding="utf-8", newline="\n")
    if failures:
        for message in failures:
            print(f"[FAIL] {message}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
