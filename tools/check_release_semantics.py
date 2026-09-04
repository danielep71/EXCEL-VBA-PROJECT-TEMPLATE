#!/usr/bin/env python3
"""Validate strict SemVer and complete changelog release semantics."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date
import json
from pathlib import Path
import re
import sys

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
        value, int(match.group(1)), int(match.group(2)), int(match.group(3)), prerelease, build
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
    for l_item, r_item in zip(left.prerelease, right.prerelease):
        if l_item == r_item:
            continue
        l_numeric = l_item.isdigit()
        r_numeric = r_item.isdigit()
        if l_numeric and r_numeric:
            return 1 if int(l_item) > int(r_item) else -1
        if l_numeric != r_numeric:
            return -1 if l_numeric else 1
        return 1 if l_item > r_item else -1
    if len(left.prerelease) == len(right.prerelease):
        return 0
    return 1 if len(left.prerelease) > len(right.prerelease) else -1


def valid_date(value: str) -> bool:
    try:
        date.fromisoformat(value)
    except ValueError:
        return False
    return True


def analyze(version: str, changelog: str, repository: str) -> dict[str, object]:
    findings: list[dict[str, object]] = []
    try:
        version_semver = parse_semver(version)
    except ValueError as error:
        version_semver = None
        findings.append({"path": "VERSION", "message": str(error)})

    unreleased_lines = [
        number for number, raw in enumerate(changelog.splitlines(), start=1)
        if raw.strip() == "## [Unreleased]"
    ]
    if len(unreleased_lines) != 1:
        findings.append({
            "path": "CHANGELOG.md",
            "message": f"Changelog requires exactly one [Unreleased] heading; observed {len(unreleased_lines)}.",
        })

    releases: list[dict[str, object]] = []
    seen: dict[str, int] = {}
    for number, raw in enumerate(changelog.splitlines(), start=1):
        match = HEADING_RE.match(raw)
        if not match:
            continue
        value, date_text = match.groups()
        try:
            parsed = parse_semver(value)
        except ValueError as error:
            findings.append({"path": "CHANGELOG.md", "line": number, "message": str(error)})
            continue
        if not valid_date(date_text):
            findings.append({
                "path": "CHANGELOG.md", "line": number,
                "message": f"release date is not a real Gregorian date: {date_text}",
            })
        if value in seen:
            findings.append({
                "path": "CHANGELOG.md", "line": number,
                "message": f"duplicate release version {value!r}; first declared at line {seen[value]}",
            })
        else:
            seen[value] = number
        releases.append({"version": value, "semver": parsed, "date": date_text, "line": number})

    for current, older in zip(releases, releases[1:]):
        precedence = compare(current["semver"], older["semver"])
        if precedence <= 0:
            findings.append({
                "path": "CHANGELOG.md", "line": current["line"],
                "message": (
                    "released versions must be strictly descending by SemVer precedence; "
                    f"{current['version']} is not newer than {older['version']}"
                ),
            })

    if releases and version_semver is not None and version != "0.0.0":
        if releases[0]["version"] != version:
            findings.append({
                "path": "VERSION",
                "message": (
                    f"VERSION {version!r} must match the newest dated changelog release "
                    f"{releases[0]['version']!r}."
                ),
            })

    links: dict[str, tuple[str, int]] = {}
    for number, raw in enumerate(changelog.splitlines(), start=1):
        match = LINK_RE.match(raw)
        if not match:
            continue
        name, url = match.groups()
        if name in links:
            findings.append({
                "path": "CHANGELOG.md", "line": number,
                "message": f"duplicate comparison-link definition for [{name}]",
            })
        else:
            links[name] = (url, number)

    if releases:
        latest = releases[0]["version"]
        expected_unreleased = f"https://github.com/{repository}/compare/v{latest}...HEAD"
        actual_unreleased = links.get("Unreleased")
        if actual_unreleased is None:
            findings.append({
                "path": "CHANGELOG.md",
                "message": f"missing [Unreleased] comparison link; expected {expected_unreleased}",
            })
        elif actual_unreleased[0] != expected_unreleased:
            findings.append({
                "path": "CHANGELOG.md", "line": actual_unreleased[1],
                "message": (
                    f"[Unreleased] comparison link must be {expected_unreleased}; "
                    f"observed {actual_unreleased[0]}"
                ),
            })

        for index, release in enumerate(releases):
            value = release["version"]
            if index + 1 < len(releases):
                older = releases[index + 1]["version"]
                expected = f"https://github.com/{repository}/compare/v{older}...v{value}"
            else:
                expected = f"https://github.com/{repository}/releases/tag/v{value}"
            actual = links.get(value)
            if actual is None:
                findings.append({
                    "path": "CHANGELOG.md",
                    "message": f"missing [{value}] release comparison link; expected {expected}",
                })
            elif actual[0] != expected:
                findings.append({
                    "path": "CHANGELOG.md", "line": actual[1],
                    "message": f"[{value}] link must be {expected}; observed {actual[0]}",
                })

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


def run_check(root: Path) -> dict[str, object]:
    version = (root / "VERSION").read_text(encoding="utf-8").strip()
    changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    return analyze(version, changelog, config["repository"])


def markdown_report(report: dict[str, object]) -> str:
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


def fixture(version: str, headings: list[tuple[str, str]], links: dict[str, str]) -> dict[str, object]:
    repository = "example/repo"
    lines = ["# Changelog", "", "## [Unreleased]", "", "No unreleased changes.", ""]
    for value, date_text in headings:
        lines.extend([f"## [{value}] - {date_text}", "", "- Change.", ""])
    for name, url in links.items():
        lines.append(f"[{name}]: {url}")
    return analyze(version, "\n".join(lines) + "\n", repository)


def canonical_links(versions: list[str]) -> dict[str, str]:
    repository = "example/repo"
    result = {"Unreleased": f"https://github.com/{repository}/compare/v{versions[0]}...HEAD"}
    for index, value in enumerate(versions):
        if index + 1 < len(versions):
            result[value] = f"https://github.com/{repository}/compare/v{versions[index + 1]}...v{value}"
        else:
            result[value] = f"https://github.com/{repository}/releases/tag/v{value}"
    return result


def run_self_test() -> int:
    cases: list[tuple[str, str, dict[str, object]]] = []
    stable_versions = ["1.1.0", "1.0.0"]
    cases.append(("valid-stable", "pass", fixture("1.1.0", [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")], canonical_links(stable_versions))))
    pre_versions = ["1.1.0-rc.1", "1.0.0"]
    cases.append(("valid-prerelease", "pass", fixture("1.1.0-rc.1", [("1.1.0-rc.1", "2026-09-05"), ("1.0.0", "2026-09-04")], canonical_links(pre_versions))))
    cases.append(("leading-zero-prerelease", "fail", fixture("1.1.0-01", [("1.1.0-01", "2026-09-05"), ("1.0.0", "2026-09-04")], canonical_links(["1.1.0-01", "1.0.0"]))))
    cases.append(("out-of-order", "fail", fixture("1.0.0", [("1.0.0", "2026-09-05"), ("1.1.0", "2026-09-04")], canonical_links(["1.0.0", "1.1.0"]))))
    cases.append(("duplicate-version", "fail", fixture("1.1.0", [("1.1.0", "2026-09-05"), ("1.1.0", "2026-09-04")], canonical_links(["1.1.0", "1.1.0"]))))
    cases.append(("impossible-date", "fail", fixture("1.1.0", [("1.1.0", "2026-02-30"), ("1.0.0", "2026-09-04")], canonical_links(stable_versions))))
    bad_links = canonical_links(stable_versions)
    bad_links["Unreleased"] = "https://github.com/example/repo/compare/v0.9.0...HEAD"
    cases.append(("wrong-unreleased-link", "fail", fixture("1.1.0", [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")], bad_links)))
    missing_link = canonical_links(stable_versions)
    del missing_link["1.1.0"]
    cases.append(("missing-release-link", "fail", fixture("1.1.0", [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")], missing_link)))
    cases.append(("version-heading-mismatch", "fail", fixture("1.2.0", [("1.1.0", "2026-09-05"), ("1.0.0", "2026-09-04")], canonical_links(stable_versions))))

    failures: list[str] = []
    for name, expected, report in cases:
        if report["status"] != expected:
            failures.append(f"{name}: expected {expected}, got {report['status']} ({report['findings']})")
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
        "SELF-TEST PASS: stable/prerelease SemVer, numeric identifier rules, precedence, duplicates, "
        "ordering, Gregorian dates, VERSION agreement, and comparison-link policy passed."
    )
    return 0


def write_text(path: Path | None, text: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    if options.self_test:
        try:
            return run_self_test()
        except (OSError, UnicodeError, ValueError, json.JSONDecodeError) as error:
            print(f"SELF-TEST ERROR: {error}", file=sys.stderr)
            return 2
    try:
        report = run_check(options.root)
        write_text(options.output, json.dumps(report, indent=2, sort_keys=True) + "\n")
        write_text(options.summary, markdown_report(report))
        print(markdown_report(report).rstrip())
        return 0 if report["status"] == "pass" else 1
    except (OSError, UnicodeError, ValueError, json.JSONDecodeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
