#!/usr/bin/env python3
"""Prove blocking policy-branch fixture coverage for the repository gates."""

from __future__ import annotations

import argparse
import ast
import importlib.util
import inspect
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from types import ModuleType
from typing import Callable

TOOL_NAME = "Policy branch coverage"
CORE_TOOL = "tools/check_repo.py"


class CoverageError(RuntimeError):
    pass


def load_module(path: Path, name: str) -> ModuleType:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise CoverageError(f"Cannot load Python module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def production_finding_sites(source_path: Path) -> dict[str, dict[str, object]]:
    source = source_path.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(source_path))
    fixture_boundary = None
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == "_write_fixture":
            fixture_boundary = node.lineno
            break
    if fixture_boundary is None:
        raise CoverageError("Cannot locate the canonical fixture boundary in check_repo.py")

    sites: dict[str, dict[str, object]] = {}

    class Visitor(ast.NodeVisitor):
        def __init__(self) -> None:
            self.functions: list[str] = []

        def visit_FunctionDef(self, node: ast.FunctionDef) -> None:
            self.functions.append(node.name)
            self.generic_visit(node)
            self.functions.pop()

        visit_AsyncFunctionDef = visit_FunctionDef

        def visit_Call(self, node: ast.Call) -> None:
            if (
                node.lineno < fixture_boundary
                and isinstance(node.func, ast.Name)
                and node.func.id == "finding"
            ):
                function = self.functions[-1] if self.functions else "<module>"
                key = f"{function}:{node.lineno}"
                expression = ast.get_source_segment(source, node) or "finding(...)"
                sites[key] = {
                    "id": key,
                    "function": function,
                    "line": node.lineno,
                    "expression": re.sub(r"\s+", " ", expression).strip(),
                }
            self.generic_visit(node)

    Visitor().visit(tree)
    return dict(sorted(sites.items(), key=lambda item: (item[1]["line"], item[0])))


def rule_by_id(report: dict[str, object], rule_id: str) -> dict[str, object] | None:
    for result in report.get("rules", []):
        if result.get("id") == rule_id:
            return result
    return None


def write_tracked(module: ModuleType, root: Path, relative: str, content: str | bytes, *, crlf: bool = False) -> None:
    module._write_fixture(root / relative, content, crlf=crlf)


def add_force(module: ModuleType, root: Path, relative: str) -> None:
    module._run_git(root, "add", "-f", relative)


def rewrite_vba(module: ModuleType, root: Path, relative: str, text: str) -> None:
    module._write_fixture(root / relative, text, crlf=True)


def extra_cases(module: ModuleType) -> list[tuple[str, str, str, Callable[[Path], None]]]:
    cases: list[tuple[str, str, str, Callable[[Path], None]]] = []

    def case(name: str, rule: str, pattern: str):
        def register(function: Callable[[Path], None]) -> Callable[[Path], None]:
            cases.append((name, rule, pattern, function))
            return function
        return register

    @case("text-nul", "text-integrity", "NUL byte")
    def _(root: Path) -> None:
        (root / "README.md").write_bytes(b"# Fixture\n\x00\n")

    @case("text-invalid-encoding", "text-integrity", "invalid encoding")
    def _(root: Path) -> None:
        (root / "README.md").write_bytes(b"# Fixture\n\xff\n")

    @case("text-private-key", "text-integrity", "private-key material")
    def _(root: Path) -> None:
        write_tracked(module, root, "README.md", "# Fixture\n\n-----BEGIN PRIVATE KEY-----\n")

    @case("text-github-token", "text-integrity", "GitHub token material")
    def _(root: Path) -> None:
        write_tracked(module, root, "README.md", "# Fixture\n\nghp_fixturetoken\n")

    @case("text-aws-key", "text-integrity", "AWS access key")
    def _(root: Path) -> None:
        write_tracked(module, root, "README.md", "# Fixture\n\nAKIAABCDEFGHIJKLMNOP\n")

    @case("line-bom", "line-endings", "BOM")
    def _(root: Path) -> None:
        (root / "README.md").write_bytes(b"\xef\xbb\xbf# Fixture\n")

    @case("line-missing-final-newline", "line-endings", "end with a newline")
    def _(root: Path) -> None:
        (root / "README.md").write_bytes(b"# Fixture")

    @case("line-cross-platform-crlf", "line-endings", "Cross-platform text must use LF")
    def _(root: Path) -> None:
        write_tracked(module, root, "README.md", "# Fixture\n\nCRLF\n", crlf=True)

    @case("line-vba-lf", "line-endings", "Windows/VBA source must use CRLF")
    def _(root: Path) -> None:
        path = root / "src/modules/Quality.bas"
        text = path.read_bytes().decode("cp1252").replace("\r\n", "\n")
        path.write_bytes(text.encode("cp1252"))

    @case("artifact-office-lock", "forbidden-artifacts", "Office lock file")
    def _(root: Path) -> None:
        write_tracked(module, root, "~$fixture.xlsx", b"lock")
        add_force(module, root, "~$fixture.xlsx")

    @case("artifact-office-binary", "forbidden-artifacts", "Office binary is not permitted")
    def _(root: Path) -> None:
        write_tracked(module, root, "fixture.xlsx", b"office")
        add_force(module, root, "fixture.xlsx")

    @case("artifact-env", "forbidden-artifacts", "Local environment or secret file")
    def _(root: Path) -> None:
        write_tracked(module, root, ".env", "SECRET=fixture\n")
        add_force(module, root, ".env")

    @case("artifact-private-directory", "forbidden-artifacts", "Private review material")
    def _(root: Path) -> None:
        write_tracked(module, root, "private/note.txt", "private\n")
        add_force(module, root, "private/note.txt")

    @case("artifact-private-token", "forbidden-artifacts", "Private review material")
    def _(root: Path) -> None:
        write_tracked(module, root, "docs/confidential-review.txt", "private\n")
        add_force(module, root, "docs/confidential-review.txt")

    @case("changelog-missing-unreleased", "version-changelog", "Unreleased")
    def _(root: Path) -> None:
        write_tracked(module, root, "CHANGELOG.md", "# Changelog\n\nFixture.\n")

    @case("vba-elseif-without-if", "vba-structure", "#ElseIf without #If")
    def _(root: Path) -> None:
        rewrite_vba(module, root, "src/modules/Quality.bas", '''Attribute VB_Name = "Quality"\nOption Explicit\n#ElseIf VBA7 Then\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n''')

    @case("vba-else-without-if", "vba-structure", "#Else without #If")
    def _(root: Path) -> None:
        rewrite_vba(module, root, "src/modules/Quality.bas", '''Attribute VB_Name = "Quality"\nOption Explicit\n#Else\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n''')

    @case("vba-endif-without-if", "vba-structure", "#End If without #If")
    def _(root: Path) -> None:
        rewrite_vba(module, root, "src/modules/Quality.bas", '''Attribute VB_Name = "Quality"\nOption Explicit\n#End If\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n''')

    @case("vba-unclosed-if", "vba-structure", "conditional-compilation block")
    def _(root: Path) -> None:
        rewrite_vba(module, root, "src/modules/Quality.bas", '''Attribute VB_Name = "Quality"\nOption Explicit\n#If VBA7 Then\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n''')

    @case("vba-reachable-declare", "vba-structure", "must include PtrSafe")
    def _(root: Path) -> None:
        rewrite_vba(module, root, "src/modules/Quality.bas", '''Attribute VB_Name = "Quality"\nOption Explicit\n#If VBA7 Then\nPrivate Declare Function Tick Lib "kernel32" () As Long\n#End If\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n''')

    @case("vba-end-without-opener", "vba-structure", "has no opener")
    def _(root: Path) -> None:
        rewrite_vba(module, root, "src/modules/Quality.bas", '''Attribute VB_Name = "Quality"\nOption Explicit\nEnd Function\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n''')

    @case("vba-missing-jump", "vba-structure", "Jump target is not defined")
    def _(root: Path) -> None:
        rewrite_vba(module, root, "src/modules/Quality.bas", '''Attribute VB_Name = "Quality"\nOption Explicit\nPublic Function Echo(ByVal value As String) As String\n    GoTo Missing\n    Echo = value\nEnd Function\n''')

    return cases


def run_core_coverage(root: Path) -> dict[str, object]:
    module = load_module(root / CORE_TOOL, "coverage_check_repo")
    source_sites = production_finding_sites(root / CORE_TOOL)
    executed: dict[str, set[str]] = {key: set() for key in source_sites}
    unexpected_sites: dict[str, set[str]] = {}
    current_case = "bootstrap"
    original_finding = module.finding

    def recording_finding(*args, **kwargs):
        frame = inspect.currentframe()
        caller = frame.f_back if frame is not None else None
        if caller is not None and Path(caller.f_code.co_filename).resolve() == (root / CORE_TOOL).resolve():
            key = f"{caller.f_code.co_name}:{caller.f_lineno}"
            if key in executed:
                executed[key].add(current_case)
            else:
                unexpected_sites.setdefault(key, set()).add(current_case)
        return original_finding(*args, **kwargs)

    module.finding = recording_finding
    case_results: list[dict[str, object]] = []

    def run_case(name: str, expected_rule: str, mutate: Callable[[Path], None], pattern: str | None = None) -> None:
        nonlocal current_case
        current_case = name
        with tempfile.TemporaryDirectory(prefix=f"policy-coverage-{name}-") as temporary:
            fixture_root = Path(temporary)
            module._initialize_fixture(fixture_root)
            mutate(fixture_root)
            report = module.build_report(fixture_root)
            rule = rule_by_id(report, expected_rule)
            status = "pass"
            detail = ""
            if rule is None:
                status = "fail"
                detail = f"expected rule {expected_rule!r} did not run"
            elif rule.get("status") != "fail":
                status = "fail"
                detail = f"expected rule {expected_rule!r} did not fail"
            elif pattern is not None and not any(
                pattern.casefold() in str(item.get("message", "")).casefold()
                for item in rule.get("findings", [])
            ):
                status = "fail"
                detail = f"expected finding pattern was absent: {pattern!r}"
            case_results.append(
                {
                    "id": name,
                    "owner": expected_rule,
                    "status": status,
                    "detail": detail,
                }
            )

    for expected_rule, mutate in module.SELF_TEST_CASES:
        run_case(f"rule-{expected_rule}", expected_rule, mutate)
    for entry in module.BRANCH_SELF_TEST_CASES:
        case_name, expected_rule, mutate = entry
        run_case(f"branch-{case_name}", expected_rule, mutate)
    for name, expected_rule, pattern, mutate in extra_cases(module):
        run_case(name, expected_rule, mutate, pattern)

    current_case = "operational-rule-crash"
    with tempfile.TemporaryDirectory(prefix="policy-coverage-operational-") as temporary:
        fixture_root = Path(temporary)
        module._initialize_fixture(fixture_root)
        original_checks = module.CHECKS

        def crash(repo, config):
            del repo, config
            raise RuntimeError("fixture operational failure")

        crash.__name__ = "check_fixture_operational_failure"
        module.CHECKS = (crash,) + original_checks
        report = module.build_report(fixture_root)
        module.CHECKS = original_checks
        rule = rule_by_id(report, "fixture-operational-failure")
        ok = (
            rule is not None
            and rule.get("status") == "fail"
            and any("Rule could not complete" in str(item.get("message", "")) for item in rule.get("findings", []))
        )
        case_results.append(
            {
                "id": "operational-rule-crash",
                "owner": "build-report",
                "status": "pass" if ok else "fail",
                "detail": "" if ok else "rule exception was not mapped to an explicit policy finding",
            }
        )

    current_case = "complete"
    uncovered = [
        {**source_sites[key], "cases": []}
        for key in source_sites
        if not executed[key]
    ]
    covered = [
        {**source_sites[key], "cases": sorted(executed[key])}
        for key in source_sites
        if executed[key]
    ]
    return {
        "source": CORE_TOOL,
        "sites": len(source_sites),
        "covered": covered,
        "uncovered": uncovered,
        "unexpected_sites": [
            {"id": key, "cases": sorted(value)}
            for key, value in sorted(unexpected_sites.items())
        ],
        "cases": case_results,
    }


def run_tool_selftest(root: Path, relative: str) -> dict[str, object]:
    completed = subprocess.run(
        [sys.executable, str(root / relative), "--root", str(root), "--self-test"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return {
        "tool": relative,
        "status": "pass" if completed.returncode == 0 else "fail",
        "exit_code": completed.returncode,
        "summary": (completed.stdout + completed.stderr).strip().splitlines()[-1]
        if (completed.stdout + completed.stderr).strip()
        else "no output",
    }


def operational_exit_fixture(root: Path) -> dict[str, object]:
    with tempfile.TemporaryDirectory(prefix="policy-coverage-not-git-") as temporary:
        completed = subprocess.run(
            [sys.executable, str(root / CORE_TOOL), "--root", temporary],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    return {
        "id": "top-level-operational-error",
        "status": "pass" if completed.returncode == 2 else "fail",
        "exit_code": completed.returncode,
        "expected_exit_code": 2,
    }


def delegated_workflow_fixtures(root: Path) -> dict[str, object]:
    workflow_source = (root / "tools/test_workflow_validation.py").read_text(encoding="utf-8")
    workflow = (root / ".github/workflows/static-checks.yml").read_text(encoding="utf-8")
    fixture_probes = (
        "valid",
        "invalid-yaml",
        "duplicate-jobs",
        "invalid-job-structure",
        "invalid-local-action-metadata",
        "missing-local-entrypoint",
    )
    missing = [probe for probe in fixture_probes if probe not in workflow_source]
    wired = (
        "Validate workflows and authoritative fixtures" in workflow
        and "AUTHORITATIVE_WORKFLOWS_OUTCOME" in workflow
    )
    return {
        "owner": "tools/test_workflow_validation.py",
        "execution": "delegated-authoritative-workflow-step",
        "fixtures": list(fixture_probes),
        "status": "pass" if not missing and wired else "fail",
        "missing": missing,
        "terminally_required": wired,
    }


def build_report(root: Path) -> dict[str, object]:
    core = run_core_coverage(root)
    focused_tools = (
        "tools/check_committed_whitespace.py",
        "tools/check_vba_jumps.py",
        "tools/check_vba_conditionals.py",
        "tools/check_vba_public_api.py",
        "tools/check_local_actions.py",
        "tools/check_release_semantics.py",
        "tools/check_release.py",
    )
    focused = [run_tool_selftest(root, path) for path in focused_tools]
    operational = operational_exit_fixture(root)
    delegated = delegated_workflow_fixtures(root)

    failures = []
    if core["uncovered"]:
        failures.append(f"{len(core['uncovered'])} canonical finding site(s) are not exercised")
    if core["unexpected_sites"]:
        failures.append(f"{len(core['unexpected_sites'])} runtime finding site(s) were not in the static inventory")
    bad_cases = [case for case in core["cases"] if case["status"] != "pass"]
    if bad_cases:
        failures.append(f"{len(bad_cases)} canonical branch fixture(s) failed")
    bad_tools = [tool for tool in focused if tool["status"] != "pass"]
    if bad_tools:
        failures.append(f"{len(bad_tools)} focused gate self-test(s) failed")
    if operational["status"] != "pass":
        failures.append("top-level operational-error exit mapping failed")
    if delegated["status"] != "pass":
        failures.append("authoritative workflow fixture delegation is incomplete")

    matrix = [
        {
            "id": item["id"],
            "owner": item["function"],
            "source_line": item["line"],
            "fixtures": item["cases"],
            "status": "covered",
        }
        for item in core["covered"]
    ]
    matrix.extend(
        {
            "id": f"focused:{Path(tool['tool']).stem}",
            "owner": tool["tool"],
            "source_line": None,
            "fixtures": ["--self-test"],
            "status": "covered" if tool["status"] == "pass" else "failed",
        }
        for tool in focused
    )
    matrix.append(
        {
            "id": "delegated:authoritative-workflows",
            "owner": delegated["owner"],
            "source_line": None,
            "fixtures": delegated["fixtures"],
            "status": "covered" if delegated["status"] == "pass" else "failed",
        }
    )
    matrix.append(
        {
            "id": "operational:top-level-exit-2",
            "owner": CORE_TOOL,
            "source_line": None,
            "fixtures": [operational["id"]],
            "status": "covered" if operational["status"] == "pass" else "failed",
        }
    )

    exclusions = [
        {
            "id": "line-coverage-percentage",
            "reason": "Numeric line coverage is deliberately excluded because certification depends on asserted policy outcomes, not executed lines.",
            "blocking": False,
        },
        {
            "id": "unreachable-hosted-service-failures",
            "reason": "External service/network behavior is not synthesized inside the dependency-free checker; the hosted terminal workflow fails closed when required authoritative steps do not complete.",
            "blocking": False,
        },
    ]
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not failures else "fail",
        "counts": {
            "canonical_sites": core["sites"],
            "canonical_covered": len(core["covered"]),
            "canonical_uncovered": len(core["uncovered"]),
            "canonical_cases": len(core["cases"]),
            "focused_selftests": len(focused),
            "matrix_rows": len(matrix),
            "exclusions": len(exclusions),
        },
        "matrix": matrix,
        "uncovered": core["uncovered"],
        "canonical_cases": core["cases"],
        "focused": focused,
        "delegated": delegated,
        "operational": operational,
        "exclusions": exclusions,
        "failures": failures,
    }


def markdown_report(report: dict[str, object]) -> str:
    counts = report["counts"]
    lines = [
        "## Policy branch coverage",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Canonical finding sites:** {counts['canonical_covered']}/{counts['canonical_sites']} covered",
        f"- **Canonical synthetic cases:** {counts['canonical_cases']}",
        f"- **Focused gate self-tests:** {counts['focused_selftests']}",
        f"- **Matrix rows:** {counts['matrix_rows']}",
        f"- **Reviewed non-blocking exclusions:** {counts['exclusions']}",
    ]
    if report["uncovered"]:
        lines.extend(["", "### Uncovered canonical finding sites", ""])
        for item in report["uncovered"]:
            lines.append(
                f"- `{item['id']}` — `{item['expression']}`"
            )
    bad_cases = [case for case in report["canonical_cases"] if case["status"] != "pass"]
    if bad_cases:
        lines.extend(["", "### Failed fixture assertions", ""])
        for case in bad_cases:
            lines.append(f"- `{case['id']}` — {case['detail']}")
    if report["failures"]:
        lines.extend(["", "### Coverage failures", ""])
        lines.extend(f"- {message}" for message in report["failures"])
    lines.extend(["", "### Reviewed exclusions", ""])
    for item in report["exclusions"]:
        lines.append(f"- `{item['id']}` — {item['reason']}")
    return "\n".join(lines) + "\n"


def write_text(path: Path | None, text: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def deterministic_selftest(root: Path) -> tuple[dict[str, object], list[str]]:
    first = build_report(root)
    second = build_report(root)
    failures: list[str] = []
    first_text = json.dumps(first, indent=2, sort_keys=True) + "\n"
    second_text = json.dumps(second, indent=2, sort_keys=True) + "\n"
    if first_text != second_text:
        failures.append("coverage JSON differs across identical runs")
    if markdown_report(first) != markdown_report(second):
        failures.append("coverage Markdown differs across identical runs")
    if first["status"] != "pass":
        failures.append("policy coverage report is not complete")
    return first, failures


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    root = options.root.resolve()
    try:
        if options.self_test:
            report, failures = deterministic_selftest(root)
            if failures:
                for message in failures:
                    print(f"[FAIL] {message}")
                print(markdown_report(report).rstrip())
                return 1
            print(
                "SELF-TEST PASS: all canonical blocking finding sites map to asserted fixtures; "
                "focused gate self-tests, delegated authoritative fixtures, operational exit mapping, "
                "and deterministic JSON/Markdown passed."
            )
            return 0
        report = build_report(root)
        write_text(options.output, json.dumps(report, indent=2, sort_keys=True) + "\n")
        write_text(options.summary, markdown_report(report))
        print(markdown_report(report).rstrip())
        return 0 if report["status"] == "pass" else 1
    except (OSError, UnicodeError, CoverageError, RuntimeError, subprocess.SubprocessError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
