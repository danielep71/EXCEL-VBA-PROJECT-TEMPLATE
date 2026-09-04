#!/usr/bin/env python3
"""Verify the development contract of the canonical portable repository checker."""

from __future__ import annotations

import argparse
import ast
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
from types import ModuleType
from typing import Any

TOOL_NAME = "Checker development contract"
CHECKER_PATH = Path("tools/check_repo.py")
SECTION_STARTS = (
    ("runtime-core", "Repository"),
    ("configuration", "_same_keys"),
    ("repository-policy", "check_required_paths"),
    ("vba-policy", "_vba_paths"),
    ("reporting", "build_report"),
    ("fixtures", "_write_fixture"),
    ("cli", "parse_arguments"),
)
EXPECTED_CHECK_FUNCTIONS = (
    "check_required_paths",
    "check_placeholders",
    "check_identity",
    "check_dotfile_policy",
    "check_structured_data",
    "check_markdown_links",
    "check_text_integrity",
    "check_forbidden_artifacts",
    "check_line_endings",
    "check_label_manifest",
    "check_issue_forms",
    "check_workflow_actions",
    "check_version_changelog",
    "check_git_diff",
    "check_vba_option_explicit",
    "check_vba_export_header",
    "check_vba_structure",
    "check_vba_visibility",
    "check_generated_vba_contract",
    "check_vba_public_api",
)
ALLOWED_IMPORT_ROOTS = {
    "argparse",
    "fnmatch",
    "hashlib",
    "json",
    "pathlib",
    "re",
    "subprocess",
    "sys",
    "tempfile",
    "typing",
    "urllib",
    "xml",
    "__future__",
}


class ContractError(RuntimeError):
    """Raised when checker development assumptions cannot be evaluated."""


def load_checker(path: Path) -> ModuleType:
    spec = importlib.util.spec_from_file_location("checker_development_target", path)
    if spec is None or spec.loader is None:
        raise ContractError(f"Cannot load checker module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def definition_nodes(tree: ast.Module) -> list[ast.AST]:
    return [
        node
        for node in tree.body
        if isinstance(node, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef))
    ]


def node_name(node: ast.AST) -> str:
    if isinstance(node, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
        return node.name
    raise ContractError("Unexpected non-definition node.")


def section_report(source: str, tree: ast.Module) -> tuple[list[dict[str, Any]], list[str]]:
    del source
    definitions = definition_nodes(tree)
    positions = {node_name(node): index for index, node in enumerate(definitions)}
    failures: list[str] = []
    starts: list[tuple[str, int]] = []
    for section, sentinel in SECTION_STARTS:
        index = positions.get(sentinel)
        if index is None:
            failures.append(f"section {section!r} sentinel is missing: {sentinel}")
            continue
        starts.append((section, index))
    if len(starts) == len(SECTION_STARTS):
        indexes = [index for _, index in starts]
        if indexes != sorted(indexes) or len(indexes) != len(set(indexes)):
            failures.append("section sentinels are not strictly ordered")
    rows: list[dict[str, Any]] = []
    if failures:
        return rows, failures
    for offset, (section, start) in enumerate(starts):
        end = starts[offset + 1][1] if offset + 1 < len(starts) else len(definitions)
        owned = definitions[start:end]
        if not owned:
            failures.append(f"section {section!r} owns no definitions")
            continue
        rows.append(
            {
                "id": section,
                "start": node_name(owned[0]),
                "end": node_name(owned[-1]),
                "definitions": len(owned),
                "start_line": getattr(owned[0], "lineno", None),
                "end_line": getattr(
                    owned[-1], "end_lineno", getattr(owned[-1], "lineno", None)
                ),
            }
        )
    covered = sum(int(row["definitions"]) for row in rows)
    if covered != len(definitions):
        failures.append(
            f"section ownership covers {covered} of {len(definitions)} top-level definitions"
        )
    return rows, failures


def import_report(tree: ast.Module) -> tuple[list[str], list[str]]:
    observed: set[str] = set()
    failures: list[str] = []
    for node in tree.body:
        if isinstance(node, ast.Import):
            observed.update(alias.name.split(".", 1)[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            if node.level:
                failures.append("runtime checker must not use relative imports")
                continue
            if node.module:
                observed.add(node.module.split(".", 1)[0])
    unsupported = sorted(observed - ALLOWED_IMPORT_ROOTS)
    if unsupported:
        failures.append(
            "runtime checker imports non-approved package roots: " + ", ".join(unsupported)
        )
    return sorted(observed), failures


def parser_tests(module: ModuleType) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []

    def record(name: str, ok: bool, detail: str = "") -> None:
        results.append({"id": name, "status": "pass" if ok else "fail", "detail": detail})

    yaml_errors = module.validate_yaml_subset("name: x\non: [push\n")
    record("yaml-invalid-flow", bool(yaml_errors), "malformed flow collection must fail")

    valid_yaml = module.validate_yaml_subset(
        "name: x\njobs:\n  test:\n    runs-on: ubuntu-latest\n"
    )
    record("yaml-basic-valid", not valid_yaml, f"unexpected errors: {valid_yaml!r}")

    slugs = module._github_slugs("# Alpha\n# Alpha\n<a id=\"stable-anchor\"></a>\n")
    record(
        "markdown-slugs",
        {"alpha", "alpha-1", "stable-anchor"}.issubset(slugs),
        f"observed slugs: {sorted(slugs)!r}",
    )

    sections = module._editorconfig_sections(
        "root = true\n\n[*.{bas,cls,frm}]\nend_of_line = crlf\n"
    )
    record(
        "editorconfig-parser",
        sections.get("", {}).get("root") == "true"
        and sections.get("*.{bas,cls,frm}", {}).get("end_of_line") == "crlf",
        f"observed sections: {sections!r}",
    )

    stripped = module._strip_vba_line(
        'Debug.Print "apostrophe stays"; value \' real comment'
    )
    record(
        "vba-line-strip",
        "real comment" not in stripped and "Debug.Print" in stripped,
        f"observed: {stripped!r}",
    )

    return results


def reporter_tests(module: ModuleType) -> list[dict[str, Any]]:
    finding = module.finding("README.md", "Pipe | must be escaped", 7)
    failed = module.rule_result("sample", "Sample", [finding], "")
    report = {
        "schema_version": 1,
        "tool": "fixture",
        "repository": "example/repo",
        "commit": "a" * 40,
        "mode": "generated",
        "profile": "library",
        "scope_note": "fixture scope",
        "status": "fail",
        "counts": {"rules": 1, "passed": 0, "failed": 1, "findings": 1},
        "rules": [failed],
    }
    first_markdown = module.markdown_report(report)
    second_markdown = module.markdown_report(report)
    first_console = module.console_report(report)
    second_console = module.console_report(report)
    return [
        {
            "id": "markdown-determinism",
            "status": "pass" if first_markdown == second_markdown else "fail",
            "detail": "",
        },
        {
            "id": "console-determinism",
            "status": "pass" if first_console == second_console else "fail",
            "detail": "",
        },
        {
            "id": "markdown-escaping",
            "status": "pass" if r"Pipe \| must be escaped" in first_markdown else "fail",
            "detail": first_markdown if r"Pipe \| must be escaped" not in first_markdown else "",
        },
        {
            "id": "finding-location",
            "status": "pass" if "README.md:7" in first_console else "fail",
            "detail": first_console if "README.md:7" not in first_console else "",
        },
    ]


def cli_tests(module: ModuleType, checker: Path) -> list[dict[str, Any]]:
    parsed = module.parse_arguments(
        ["--root", ".", "--output", "out.json", "--summary", "out.md"]
    )
    flags_ok = (
        parsed.root == Path(".")
        and parsed.output == Path("out.json")
        and parsed.summary == Path("out.md")
        and parsed.self_test is False
    )
    with tempfile.TemporaryDirectory(prefix="checker-contract-not-git-") as temporary:
        completed = subprocess.run(
            [sys.executable, str(checker), "--root", temporary],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    return [
        {
            "id": "cli-flags",
            "status": "pass" if flags_ok else "fail",
            "detail": repr(parsed) if not flags_ok else "",
        },
        {
            "id": "operational-exit-2",
            "status": "pass" if completed.returncode == 2 else "fail",
            "detail": f"observed exit code {completed.returncode}",
        },
    ]


def check_ids(module: ModuleType) -> tuple[list[str], list[str]]:
    functions = [check.__name__ for check in module.CHECKS]
    failures = []
    if tuple(functions) != EXPECTED_CHECK_FUNCTIONS:
        failures.append("canonical CHECKS order drifted: " + ", ".join(functions))
    return functions, failures


def build_report(root: Path) -> dict[str, Any]:
    checker = (root / CHECKER_PATH).resolve()
    source = checker.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(checker))
    module = load_checker(checker)

    sections, failures = section_report(source, tree)
    imports, import_failures = import_report(tree)
    failures.extend(import_failures)
    ids, id_failures = check_ids(module)
    failures.extend(id_failures)

    parser_results = parser_tests(module)
    reporter_results = reporter_tests(module)
    cli_results = cli_tests(module, checker)
    all_unit_results = [*parser_results, *reporter_results, *cli_results]
    failed_units = [item for item in all_unit_results if item["status"] != "pass"]
    if failed_units:
        failures.append(f"{len(failed_units)} independent parser/reporter/CLI test(s) failed")

    digest = hashlib.sha256(checker.read_bytes()).hexdigest()
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not failures else "fail",
        "artifact": {
            "path": CHECKER_PATH.as_posix(),
            "development_model": "reviewed-source-is-distributable",
            "build_transform": "none",
            "sha256": digest,
            "runtime_dependencies": "python-standard-library-only",
        },
        "sections": sections,
        "imports": imports,
        "canonical_checks": ids,
        "unit_tests": all_unit_results,
        "failures": failures,
    }


def markdown_report(report: dict[str, Any]) -> str:
    artifact = report["artifact"]
    lines = [
        "## Checker development contract",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Runtime artifact:** `{artifact['path']}`",
        f"- **Development model:** `{artifact['development_model']}`",
        f"- **Build transform:** `{artifact['build_transform']}`",
        f"- **SHA-256:** `{artifact['sha256']}`",
        f"- **Runtime dependencies:** `{artifact['runtime_dependencies']}`",
        f"- **Internal sections:** {len(report['sections'])}",
        f"- **Canonical policy checks:** {len(report['canonical_checks'])}",
        f"- **Independent unit tests:** {len(report['unit_tests'])}",
        "",
        "| Section | Start | End | Definitions |",
        "| --- | --- | --- | ---: |",
    ]
    for section in report["sections"]:
        lines.append(
            f"| `{section['id']}` | `{section['start']}` | `{section['end']}` | "
            f"{section['definitions']} |"
        )
    lines.extend(["", "### Independent tests", "", "| Test | Result |", "| --- | --- |"])
    for item in report["unit_tests"]:
        lines.append(f"| `{item['id']}` | {str(item['status']).upper()} |")
    if report["failures"]:
        lines.extend(["", "### Failures", ""])
        lines.extend(f"- {item}" for item in report["failures"])
    return "\n".join(lines).rstrip() + "\n"


def write_text(path: Path | None, text: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def run_self_test(root: Path) -> int:
    first = build_report(root)
    second = build_report(root)
    failures: list[str] = []
    if first["status"] != "pass":
        failures.append("development contract is not satisfied")
    if json.dumps(first, sort_keys=True) != json.dumps(second, sort_keys=True):
        failures.append("JSON evidence is not deterministic")
    if markdown_report(first) != markdown_report(second):
        failures.append("Markdown evidence is not deterministic")
    if failures:
        for item in failures:
            print(f"[FAIL] {item}")
        print(markdown_report(first).rstrip())
        return 1
    print(
        "SELF-TEST PASS: internal boundaries, parser/reporter units, CLI contract, "
        "canonical check order, artifact identity, and standard-library-only runtime passed."
    )
    return 0


def parse_arguments(arguments: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(arguments)


def main(arguments: list[str] | None = None) -> int:
    options = parse_arguments(sys.argv[1:] if arguments is None else arguments)
    try:
        if options.self_test:
            return run_self_test(options.root)
        report = build_report(options.root)
        write_text(options.output, json.dumps(report, indent=2, sort_keys=True) + "\n")
        write_text(options.summary, markdown_report(report))
        print(markdown_report(report).rstrip())
        return 0 if report["status"] == "pass" else 1
    except (
        OSError,
        UnicodeError,
        SyntaxError,
        ContractError,
        subprocess.SubprocessError,
    ) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
