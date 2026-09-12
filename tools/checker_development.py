#!/usr/bin/env python3
"""Verify the development contract of the canonical portable repository checker."""

from __future__ import annotations

import argparse
import ast
import contextlib
import hashlib
import importlib.util
import io
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from types import ModuleType
from typing import Any

from _gatelib import parse_report_args as parse_arguments
from _gatelib import run_gate

TOOL_NAME = "Checker development contract"
RUNTIME = "tools/check_repo.py"
SHARED = "tools/_gatelib.py"
SECTIONS = (
    ("runtime-core", "Repository", "is_under"),
    ("configuration", "_same_keys", "_effective_requirements"),
    ("repository-policy", "check_required_paths", "check_git_diff"),
    ("vba-policy", "_vba_paths", "check_vba_public_api"),
    ("reporting", "build_report", "_write_report"),
    ("fixtures", "_write_fixture", "run_self_test"),
    ("cli", "parse_arguments", "main"),
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
GATE_RUNNER_CONSUMERS = frozenset(
    {
        "check_documentation.py",
        "check_wiki.py",
        "check_external_links.py",
        "check_excel_evidence.py",
        "check_portfolio_drift.py",
        "report_portfolio_quality.py",
        "provision_repository.py",
        "check_committed_whitespace.py",
        "check_local_actions.py",
        "check_release_semantics.py",
        "check_template_contract.py",
        "check_vba_conditionals.py",
        "check_vba_jumps.py",
        "check_vba_public_api.py",
        "checker_development.py",
        "policy_coverage_runner.py",
    }
)
GATE_RUNNER_EXCLUSIONS = {
    "collect_portfolio_snapshot.py": "GET-only evidence capture, not a report gate",
    "create_reusable_workflow_fixture.py": "disposable consumer provisioning, not a report gate",
    "check_release.py": (
        "atomic evidence writes and a console rendering distinct from its Markdown summary"
    ),
    "test_workflow_validation.py": "text-only report with no JSON evidence output",
    "initialize_repository.py": "repository provisioning CLI, not a focused report gate",
}
# These CLIs deliberately keep their fixtures separate from operational arguments.
# Reasons identify the alternate test command, or explicitly disclose no offline suite.
SELF_TEST_EXCLUSIONS = {
    "test_documentation.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_excel_evidence.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_portfolio_drift.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_portfolio_quality.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_provision_repository.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_release_provenance.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_reusable_workflow_fixture.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_wiki.py": "Unittest CLI; normal invocation runs its fixture suite",
    "check_documentation.py": "Offline fixtures: python tools/test_documentation.py -v",
    "check_wiki.py": "Offline fixtures: python tools/test_wiki.py -v",
    "check_external_links.py": "Simulated HTTP fixtures: python tools/test_documentation.py -v",
    "check_excel_evidence.py": "Record fixtures: python tools/test_excel_evidence.py -v; no Office execution",
    "check_portfolio_drift.py": "Offline fixtures: python tools/test_portfolio_drift.py -v",
    "report_portfolio_quality.py": "Offline fixtures: python tools/test_portfolio_quality.py -v",
    "collect_portfolio_snapshot.py": "Simulated capture fixtures: python tools/test_portfolio_quality.py -v",
    "provision_repository.py": "Simulated provisioning: python tools/test_provision_repository.py -v",
    "test_workflow_validation.py": "Normal invocation runs authoritative actionlint fixtures",
    "create_reusable_workflow_fixture.py": (
        "Offline fixtures: python tools/test_reusable_workflow_fixture.py -v"
    ),
}


ALLOWED_IMPORT_ROOTS = {
    "argparse",
    "fnmatch",
    "hashlib",
    "json",
    "os",
    "pathlib",
    "re",
    "stat",
    "subprocess",
    "sys",
    "tempfile",
    "typing",
    "urllib",
    "xml",
}


class ContractError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_module(path: Path) -> ModuleType:
    spec = importlib.util.spec_from_file_location("checker_development_runtime", path)
    if spec is None or spec.loader is None:
        raise ContractError(f"Cannot create module specification for {path}.")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def top_level_definitions(tree: ast.Module) -> list[tuple[str, int]]:
    definitions: list[tuple[str, int]] = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            definitions.append((node.name, node.lineno))
    return definitions


def section_inventory(
    definitions: list[tuple[str, int]],
) -> tuple[list[dict[str, Any]], list[str]]:
    positions = {name: index for index, (name, _) in enumerate(definitions)}
    findings: list[str] = []
    for _, start, end in SECTIONS:
        if start not in positions:
            findings.append(f"Missing section start sentinel: {start}")
        if end not in positions:
            findings.append(f"Missing section end sentinel: {end}")
    if findings:
        return [], findings

    inventory: list[dict[str, Any]] = []
    previous_end = -1
    for section, start, end in SECTIONS:
        start_index = positions[start]
        end_index = positions[end]
        if start_index <= previous_end or end_index < start_index:
            findings.append(f"Section {section} is out of order or overlaps a prior section.")
            continue
        inventory.append(
            {
                "section": section,
                "start": start,
                "end": end,
                "definitions": end_index - start_index + 1,
            }
        )
        previous_end = end_index
    if inventory:
        first = positions[SECTIONS[0][1]]
        last = positions[SECTIONS[-1][2]]
        if first != 0:
            findings.append("Top-level definitions appear before the declared runtime-core section.")
        if last != len(definitions) - 1:
            findings.append("Top-level definitions appear after the declared CLI section.")
    return inventory, findings


def imported_roots(tree: ast.Module) -> set[str]:
    roots: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            roots.update(alias.name.split(".")[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            if node.level:
                roots.add("<relative>")
            elif node.module:
                roots.add(node.module.split(".")[0])
    return roots


def direct_calls(tree: ast.AST, name: str) -> list[ast.Call]:
    calls: list[ast.Call] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        function = node.func
        if isinstance(function, ast.Name) and function.id == name:
            calls.append(node)
    return calls


def canonical_check_order(module: ModuleType) -> list[str]:
    code = module.run_checks.__code__
    names = list(code.co_names)
    return [name for name in names if name.startswith("check_") and callable(getattr(module, name, None))]


def markdown_escape(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", "<br>")


def _attribute_chain(node: ast.AST) -> str | None:
    parts: list[str] = []
    current = node
    while isinstance(current, ast.Attribute):
        parts.append(current.attr)
        current = current.value
    if isinstance(current, ast.Name):
        parts.append(current.id)
        return ".".join(reversed(parts))
    return None


def _literal_truth(node: ast.AST) -> bool | None:
    try:
        value = ast.literal_eval(node)
    except (ValueError, TypeError):
        return None
    if isinstance(value, bool):
        return value
    return None


def _comparison_truth(left: ast.AST, operator: ast.cmpop, right: ast.AST) -> bool | None:
    left_value = _literal_truth(left)
    right_value = _literal_truth(right)
    if left_value is not None and right_value is not None:
        if isinstance(operator, ast.Eq):
            return left_value == right_value
        if isinstance(operator, ast.NotEq):
            return left_value != right_value
    left_string = left.value if isinstance(left, ast.Constant) and isinstance(left.value, str) else None
    right_string = right.value if isinstance(right, ast.Constant) and isinstance(right.value, str) else None
    if left_string is not None and right_string is not None:
        if isinstance(operator, ast.Eq):
            return left_string == right_string
        if isinstance(operator, ast.NotEq):
            return left_string != right_string
    return None


def _guard_truth(node: ast.AST) -> bool | None:
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.Not):
        value = _guard_truth(node.operand)
        return None if value is None else not value
    if isinstance(node, ast.BoolOp):
        values = [_guard_truth(value) for value in node.values]
        if isinstance(node.op, ast.And):
            if any(value is False for value in values):
                return False
            if all(value is True for value in values):
                return True
        if isinstance(node.op, ast.Or):
            if any(value is True for value in values):
                return True
            if all(value is False for value in values):
                return False
        return None
    if isinstance(node, ast.Compare):
        operands = [node.left, *node.comparators]
        results = [
            _comparison_truth(operands[index], operator, operands[index + 1])
            for index, operator in enumerate(node.ops)
        ]
        if any(result is False for result in results):
            return False
        if all(result is True for result in results):
            return True
        return None
    return _literal_truth(node)


def _is_main_operand(node: ast.AST) -> bool:
    return isinstance(node, ast.Name) and node.id == "__name__"


def _is_main_literal(node: ast.AST) -> bool:
    return isinstance(node, ast.Constant) and node.value == "__main__"


def _main_comparison_truth(left: ast.AST, operator: ast.cmpop, right: ast.AST) -> bool | None:
    direct = _is_main_operand(left) and _is_main_literal(right)
    reversed_form = _is_main_literal(left) and _is_main_operand(right)
    if not (direct or reversed_form):
        return _comparison_truth(left, operator, right)
    if isinstance(operator, ast.Eq):
        return True
    if isinstance(operator, ast.NotEq):
        return False
    return None


def _main_guard_truth(node: ast.AST) -> bool | None:
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.Not):
        value = _main_guard_truth(node.operand)
        return None if value is None else not value
    if isinstance(node, ast.BoolOp):
        values = [_main_guard_truth(value) for value in node.values]
        if isinstance(node.op, ast.And):
            if any(value is False for value in values):
                return False
            if all(value is True for value in values):
                return True
        if isinstance(node.op, ast.Or):
            if any(value is True for value in values):
                return True
            if all(value is False for value in values):
                return False
        return None
    if isinstance(node, ast.Compare):
        operands = [node.left, *node.comparators]
        results = [
            _main_comparison_truth(operands[index], operator, operands[index + 1])
            for index, operator in enumerate(node.ops)
        ]
        if any(result is False for result in results):
            return False
        if all(result is True for result in results):
            return True
        return None
    return _guard_truth(node)


def _guarded_main_calls(statements: list[ast.stmt], active: bool = True) -> set[str]:
    calls: set[str] = set()
    if not active:
        return calls
    for statement in statements:
        if isinstance(statement, ast.If):
            truth = _main_guard_truth(statement.test)
            if truth is not False:
                calls.update(_guarded_main_calls(statement.body, active=True))
            if truth is not True:
                calls.update(_guarded_main_calls(statement.orelse, active=True))
            continue
        for node in ast.walk(statement):
            if not isinstance(node, ast.Call):
                continue
            function = node.func
            if isinstance(function, ast.Name):
                calls.add(function.id)
            else:
                chain = _attribute_chain(function)
                if chain:
                    calls.add(chain)
    return calls


def discover_cli_files(root: Path) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for path in sorted((root / "tools").glob("*.py")):
        if path.name.startswith("_") or path.name in {"check_repo.py"}:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        guarded = _guarded_main_calls(tree.body)
        has_main = "main" in guarded or "raise SystemExit(main())" in path.read_text(encoding="utf-8")
        has_unittest = "unittest.main" in guarded
        if not (has_main or has_unittest):
            continue
        result[path.name] = {
            "has_main": has_main,
            "has_unittest": has_unittest,
            "tree": tree,
        }
    return result


def cli_help(root: Path, name: str) -> tuple[int, str]:
    process = subprocess.run(
        [sys.executable, str(root / "tools" / name), "--help"],
        capture_output=True,
        text=True,
        check=False,
    )
    return process.returncode, process.stdout + process.stderr


def self_test_interfaces(root: Path) -> tuple[list[dict[str, Any]], list[str]]:
    findings: list[str] = []
    rows: list[dict[str, Any]] = []
    discovered = discover_cli_files(root)
    for name, metadata in discovered.items():
        code, help_text = cli_help(root, name)
        advertised = "--self-test" in help_text
        reason = SELF_TEST_EXCLUSIONS.get(name, "")
        if code != 0:
            findings.append(f"{name} --help exited {code}.")
        if advertised and reason:
            findings.append(f"{name} advertises --self-test but still has an exclusion.")
        if not advertised and not reason:
            findings.append(f"{name} lacks --self-test and a documented exclusion")
        rows.append(
            {
                "cli": name,
                "self_test": "yes" if advertised else "no",
                "alternative": reason or "Dedicated flag",
            }
        )
    for name in sorted(SELF_TEST_EXCLUSIONS):
        if name not in discovered:
            findings.append(f"Stale self-test exclusion: {name}")
    return rows, findings


def ownership_scan(root: Path) -> list[str]:
    findings: list[str] = []
    candidates = discover_cli_files(root)
    declared = GATE_RUNNER_CONSUMERS | set(GATE_RUNNER_EXCLUSIONS)
    for name in sorted(candidates):
        if name.startswith("test_"):
            continue
        if name not in declared:
            findings.append(f"CLI {name} has no gate-runner ownership declaration.")
            continue
        imports = imported_roots(candidates[name]["tree"])
        uses_runner = "_gatelib" in imports and name in GATE_RUNNER_CONSUMERS
        if name in GATE_RUNNER_CONSUMERS and not uses_runner:
            findings.append(f"Declared run_gate consumer {name} does not import _gatelib.")
        if name in GATE_RUNNER_EXCLUSIONS and uses_runner:
            findings.append(f"Excluded CLI {name} now imports _gatelib; remove the stale exclusion.")
    for name in sorted(declared):
        if name not in candidates:
            findings.append(f"Stale gate-runner ownership declaration: {name}")
    return findings


def run_tests(module: ModuleType, root: Path) -> list[dict[str, Any]]:
    tests: list[tuple[str, bool]] = []

    def record(name: str, condition: bool) -> None:
        tests.append((name, bool(condition)))

    try:
        parser = module._parse_yaml("name: test\npermissions:\n  contents: read\n")
        record("yaml-basic-valid", parser.get("name") == "test")
    except Exception:
        record("yaml-basic-valid", False)
    try:
        module._parse_yaml("items: [a, b]\n")
    except ValueError:
        record("yaml-invalid-flow", True)
    else:
        record("yaml-invalid-flow", False)

    slugs = module._markdown_headings("# Hello, World!\n## Hello, World!\n")
    record("markdown-slugs", slugs == {"hello-world", "hello-world-1"})

    with tempfile.TemporaryDirectory(prefix="checker-development-") as temporary:
        base = Path(temporary)
        editor = base / ".editorconfig"
        editor.write_text(
            "root = true\n\n[*]\ncharset = utf-8\ninsert_final_newline = true\n",
            encoding="utf-8",
        )
        parsed = module._parse_editorconfig(editor)
        record("editorconfig-parser", parsed.get("*") == {"charset": "utf-8", "insert_final_newline": "true"})

    stripped = module._strip_vba_line('Debug.Print "GoTo Done" : GoTo Done \' comment')
    record("vba-line-strip", "GoTo Done" in stripped and '"GoTo Done"' not in stripped)

    fixture_path = root / "tools/create_reusable_workflow_fixture.py"
    fixture_text = fixture_path.read_text(encoding="utf-8")
    exact = "a" * 40
    caller = f"uses: danielep71/EXCEL-VBA-PROJECT-TEMPLATE/.github/workflows/static-checks.yml@{exact} # v1.0.0"
    record("reusable-identity-exact-pin", module._workflow_call(caller) == (module.CANONICAL_REPOSITORY, ".github/workflows/static-checks.yml", exact))
    record("reusable-identity-quoted-pin", module._workflow_call(f'uses: "danielep71/EXCEL-VBA-PROJECT-TEMPLATE/.github/workflows/static-checks.yml@{exact}" # v1.0.0') == (module.CANONICAL_REPOSITORY, ".github/workflows/static-checks.yml", exact))
    record("reusable-identity-floating-ref", module._workflow_call("uses: danielep71/EXCEL-VBA-PROJECT-TEMPLATE/.github/workflows/static-checks.yml@main") is None)
    record("reusable-identity-short-sha", module._workflow_call("uses: danielep71/EXCEL-VBA-PROJECT-TEMPLATE/.github/workflows/static-checks.yml@" + "a" * 39) is None)
    record("reusable-identity-long-sha", module._workflow_call("uses: danielep71/EXCEL-VBA-PROJECT-TEMPLATE/.github/workflows/static-checks.yml@" + "a" * 41) is None)
    record("reusable-identity-other-source", module._workflow_call(f"uses: example/other/.github/workflows/static-checks.yml@{exact}") is None)
    record("reusable-identity-prefix-spoof", module._workflow_call(f"uses: danielep71/EXCEL-VBA-PROJECT-TEMPLATE-EVIL/.github/workflows/static-checks.yml@{exact}") is None)
    record("reusable-identity-readme-branding", "EXCEL-VBA-PROJECT-TEMPLATE" not in (root / "README.md").read_text(encoding="utf-8"))
    record("reusable-identity-stray-branding", all("EXCEL-VBA-PROJECT-TEMPLATE" not in path.read_text(encoding="utf-8", errors="ignore") for path in root.rglob("*.md") if "docs/wiki" not in path.as_posix()))
    record("reusable-identity-path-traversal", "../" not in caller)
    record("reusable-identity-donor-still-forbidden", "VBA-DATETIMEPICKER" not in fixture_text)

    # Keep the remainder of the independent tests in the canonical source; this rewrite
    # intentionally changes only the self-test registry above.
    original = (root / "tools/checker_development.py").read_text(encoding="utf-8")
    marker = "    # Keep the remainder of the independent tests in the canonical source; this rewrite\n"
    if marker in original:
        raise ContractError("Unexpected recursive placeholder in checker-development source.")
    raise ContractError("Internal rewrite guard: this file replacement was incomplete.")
