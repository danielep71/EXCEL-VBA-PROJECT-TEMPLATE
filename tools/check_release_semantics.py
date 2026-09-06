#!/usr/bin/env python3
"""Validate strict SemVer and complete changelog release semantics."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import json
from pathlib import Path
import re
import sys
from typing import Any

from _gatelib import parse_report_args as parse_args, run_gate

CONFIG_PATH = ".github/repository-profile.json"
SEMVER_RE = re.compile(
    r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)"
    r"(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?"
    r"(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?$"
)
HEADING_RE = re.compile(r"^## \[([^\]]+)\] - (\d{4}-\d{2}-\d{2})\s*$")
LINK_RE = re.compile(r"^\[([^\]]+)\]:\s*(\S+)\s*$")
TOOL_NAME = "Release semantics"


@dataclass(frozen=True)
class SemVer:
    text: str
    major: int
    minor: int
    patch: int
    prerelease: tuple[str, ...]
    build: tuple[str, ...]


def parse_semver(value: str) -> SemVer:
    match = SEMVER_RE.fullmatch(value)
    if not match:
        raise ValueError(f"invalid SemVer: {value!r}")
    prerelease = tuple(match.group(4).split(".")) if match.group(4) else ()
    build = tuple(match.group(5).split(".")) if match.group(5) else ()
    for identifier in prerelease:
        if identifier.isdigit() and len(identifier) > 1 and identifier.startswith("0"):
            raise ValueError(
                f"numeric pre-release identifier must not contain leading zeros: {identifier!r}"
            )
    return SemVer(
        value,
        int(match.group(1)),
        int(match.group(2)),
        int(match.group(3)),
        prerelease,
        build,
    )


def compare(left: SemVer, right: SemVer) -> int:
    core_left = (left.major, left.minor, left.patch)
    core_right = (right.major, right.minor, right.patch)
    if core_left != core_right:
        return 1 if core_left > core_right else -1
    if not left.prerelease and not right.prerelease:
        return 0
    if not left.prerelease:
        return 1
    if not right.prerelease:
        return -1
    for left_item, right_item in zip(left.prerelease, right.prerelease):
        if left_item == right_item:
            continue
        left_numeric = left_item.isdigit()
        right_numeric = right_item.isdigit()
        if left_numeric and right_numeric:
            return 1 if int(left_item) > int(right_item) else -1
        if left_numeric != right_numeric:
            return -1 if left_numeric else 1
        return 1 if left_item > right_item else -1
    if len(left.prerelease) == len(right.prerelease):
        return 0
    return 1 if len(left.prerelease) > len(right.prerelease) else -1


def valid_date(value: str) -> bool:
    try:
        date.fromisoformat(value)
    except ValueError:
        return False
    return True


def _parse_releases(
    changelog: str, findings: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    releases: list[dict[str, Any]] = []
    seen: dict[str, int] = {}
    for number, raw in enumerate(changelog.splitlines(), start=1):
        match = HEADING_RE.match(raw)
        if not match:
            continue
        value, date_text = match.groups()
        try:
            parsed = parse_semver(value)
        except ValueError as error:
            findings.append(
                {"path": "CHANGELOG.md", "line": number, "message": str(error)}
            )
            continue
        if not valid_date(date_text):
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "line": number,
                    "message": f"release date is not a real Gregorian date: {date_text}",
                }
            )
        if value in seen:
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "line": number,
                    "message": (
                        f"duplicate release version {value!r}; first declared at line {seen[value]}"
                    ),
                }
            )
        else:
            seen[value] = number
        releases.append(
            {"version": value, "semver": parsed, "date": date_text, "line": number}
        )
    return releases


def _validate_release_order(
    releases: list[dict[str, Any]],
    version: str,
    version_semver: SemVer | None,
    findings: list[dict[str, Any]],
) -> None:
    for current, older in zip(releases, releases[1:]):
        if compare(current["semver"], older["semver"]) > 0:
            continue
        findings.append(
            {
                "path": "CHANGELOG.md",
                "line": current["line"],
                "message": (
                    "released versions must be strictly descending by SemVer precedence; "
                    f"{current['version']} is not newer than {older['version']}"
                ),
            }
        )
    if (
        releases
        and version_semver is not None
        and version != "0.0.0"
        and releases[0]["version"] != version
    ):
        findings.append(
            {
                "path": "VERSION",
                "message": (
                    f"VERSION {version!r} must match the newest dated changelog release "
                    f"{releases[0]['version']!r}."
                ),
            }
        )


def _parse_links(
    changelog: str, findings: list[dict[str, Any]]
) -> dict[str, tuple[str, int]]:
    links: dict[str, tuple[str, int]] = {}
    for number, raw in enumerate(changelog.splitlines(), start=1):
        match = LINK_RE.match(raw)
        if not match:
            continue
        name, url = match.groups()
        if name in links:
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "line": number,
                    "message": f"duplicate comparison-link definition for [{name}]",
                }
            )
        else:
            links[name] = (url, number)
    return links


def _validate_expected_link(
    links: dict[str, tuple[str, int]],
    name: str,
    expected: str,
    missing_message: str,
    findings: list[dict[str, Any]],
) -> None:
    actual = links.get(name)
    if actual is None:
        findings.append({"path": "CHANGELOG.md", "message": missing_message})
    elif actual[0] != expected:
        findings.append(
            {
                "path": "CHANGELOG.md",
                "line": actual[1],
                "message": f"[{name}] link must be {expected}; observed {actual[0]}",
            }
        )


def _validate_links(
    releases: list[dict[str, Any]],
    links: dict[str, tuple[str, int]],
    repository: str,
    findings: list[dict[str, Any]],
) -> None:
    if not releases:
        return
    latest = releases[0]["version"]
    expected_unreleased = f"https://github.com/{repository}/compare/v{latest}...HEAD"
    _validate_expected_link(
        links,
        "Unreleased",
        expected_unreleased,
        f"missing [Unreleased] comparison link; expected {expected_unreleased}",
        findings,
    )
    for index, release in enumerate(releases):
        value = release["version"]
        if index + 1 < len(releases):
            older = releases[index + 1]["version"]
            expected = f"https://github.com/{repository}/compare/v{older}...v{value}"
        else:
            expected = f"https://github.com/{repository}/releases/tag/v{value}"
        _validate_expected_link(
            links,
            value,
            expected,
            f"missing [{value}] release comparison link; expected {expected}",
            findings,
        )


def analyze(version: str, changelog: str, repository: str) -> dict[str, Any]:
    findings: list[dict[str, Any]] = []
    try:
        version_semver = parse_semver(version)
    except ValueError as error:
        version_semver = None
        findings.append({"path": "VERSION", "message": str(error)})

    unreleased_count = sum(
        raw.strip() == "## [Unreleased]" for raw in changelog.splitlines()
    )
    if unreleased_count != 1:
        findings.append(
            {
                "path": "CHANGELOG.md",
                "message": (
                    "Changelog requires exactly one [Unreleased] heading; "
                    f"observed {unreleased_count}."
                ),
            }
        )

    releases = _parse_releases(changelog, findings)
    _validate_release_order(releases, version, version_semver, findings)
    links = _parse_links(changelog, findings)
    _validate_links(releases, links, repository, findings)

    release_evidence = [
        {"version": item["version"], "date": item["date"], "line": item["line"]}
        for item in releases
    ]
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "version": version,
        "repository": repository,
        "releases": release_evidence,
        "link_policy": {
            "unreleased": "latest-tag...HEAD",
            "initial_release": "release-tag",
            "later_release": "preceding-tag...release-tag",
        },
        "findings": findings,
    }


def run_check(root: Path) -> dict[str, Any]:
    version = (root / "VERSION").read_text(encoding="utf-8").strip()
    changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    return analyze(version, changelog, config["repository"])


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "## Release semantics",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **VERSION:** `{report['version']}`",
        f"- **Released headings:** {len(report['releases'])}",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["releases"]:
        lines.extend(["", "| Version | Date | Line |", "| --- | --- | ---: |"])
        for item in report["releases"]:
            lines.append(f"| `{item['version']}` | {item['date']} | {item['line']} |")
    if report["findings"]:
        lines.extend(["", "### Findings", ""])
        for item in report["findings"]:
            location = item.get("path", ".")
            if item.get("line"):
                location += f":{item['line']}"
            lines.append(f"- `{location}` — {item['message']}")
    return "\n".join(lines) + "\n"


def fixture(
    version: str, headings: list[tuple[str, str]], links: dict[str, str]
) -> dict[str, Any]:
    repository = "example/repo"
    lines = ["# Changelog", "", "## [Unreleased]", "", "No unreleased changes.", ""]
    for value, date_text in headings:
        lines.extend([f"## [{value}] - {date_text}", "", "- Change.", ""])
    for name, url in links.items():
        lines.append(f"[{name}]: {url}")
    return analyze(version, "\n".join(lines) + "\n", repository)


def canonical_links(versions: list[str]) -> dict[str, str]:
    repository = "example/repo"
    result = {
        "Unreleased": f"https://github.com/{repository}/compare/v{versions[0]}...HEAD"
    }
    for index, value in enumerate(versions):
        if index + 1 < len(versions):
            result[value] = (
                f"https://github.com/{repository}/compare/v{versions[index + 1]}...v{value}"
            )
        else:
            result[value] = f"https://github.com/{repository}/releases/tag/v{value}"
    return result


def run_self_test() -> int:
    cases: list[tuple[str, str, dict[str, Any]]] = []
    stable_versions = ["1.1.0", "1.0.0"]
    cases.append(
        (
            "valid-stable",
            "pass",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")],
                canonical_links(stable_versions),
            ),
        )
    )
    pre_versions = ["1.1.0-rc.1", "1.0.0"]
    cases.append(
        (
            "valid-prerelease",
            "pass",
            fixture(
                "1.1.0-rc.1",
                [("1.1.0-rc.1", "2026-09-05"), ("1.0.0", "2026-09-04")],
                canonical_links(pre_versions),
            ),
        )
    )
    cases.append(
        (
            "leading-zero-prerelease",
            "fail",
            fixture(
                "1.1.0-01",
                [("1.1.0-01", "2026-09-05"), ("1.0.0", "2026-09-04")],
                canonical_links(["1.1.0-01", "1.0.0"]),
            ),
        )
    )
    cases.append(
        (
            "out-of-order",
            "fail",
            fixture(
                "1.0.0",
                [("1.0.0", "2026-09-05"), ("1.1.0", "2026-09-04")],
                canonical_links(["1.0.0", "1.1.0"]),
            ),
        )
    )
    cases.append(
        (
            "duplicate-version",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.1.0", "2026-09-04")],
                canonical_links(["1.1.0", "1.1.0"]),
            ),
        )
    )
    cases.append(
        (
            "impossible-date",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-02-30"), ("1.0.0", "2026-09-04")],
                canonical_links(stable_versions),
            ),
        )
    )
    bad_links = canonical_links(stable_versions)
    bad_links["Unreleased"] = "https://github.com/example/repo/compare/v0.9.0...HEAD"
    cases.append(
        (
            "wrong-unreleased-link",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")],
                bad_links,
            ),
        )
    )
    missing_link = canonical_links(stable_versions)
    del missing_link["1.1.0"]
    cases.append(
        (
            "missing-release-link",
            "fail",
            fixture(
                "1.1.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")],
                missing_link,
            ),
        )
    )
    cases.append(
        (
            "version-heading-mismatch",
            "fail",
            fixture(
                "1.2.0",
                [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")],
                canonical_links(stable_versions),
            ),
        )
    )

    failures: list[str] = []
    for name, expected, report in cases:
        if report["status"] != expected:
            failures.append(
                f"{name}: expected {expected}, got {report['status']} ({report['findings']})"
            )
    if compare(parse_semver("1.0.0-alpha.2"), parse_semver("1.0.0-alpha.10")) >= 0:
        failures.append("SemVer numeric prerelease precedence is incorrect")
    if compare(parse_semver("1.0.0"), parse_semver("1.0.0-rc.1")) <= 0:
        failures.append("stable release must outrank prerelease")

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: stable/prerelease SemVer, numeric identifier rules, precedence, "
        "duplicates, ordering, Gregorian dates, VERSION agreement, and comparison-link "
        "policy passed."
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    return run_gate(
        options,
        build=lambda: run_check(options.root),
        markdown=markdown_report,
        errors=(OSError, UnicodeError, ValueError, json.JSONDecodeError),
        self_test=run_self_test,
        self_test_error_prefix="SELF-TEST ERROR",
    )

if __name__ == "__main__":
    raise SystemExit(main())
