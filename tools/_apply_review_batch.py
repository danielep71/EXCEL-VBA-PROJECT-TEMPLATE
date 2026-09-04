#!/usr/bin/env python3
"""One-shot v1.1.0 migration for review findings #6-#12.

Removed after the validated source changes are committed.
"""

from __future__ import annotations

import ast
import json
from pathlib import Path
import re

ROOT = Path(".")
TOOLS = ROOT / "tools"


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def write(path: str, content: str) -> None:
    target = ROOT / path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(content, encoding="utf-8", newline="\n")


def write_vba(path: str, content: str) -> None:
    target = ROOT / path
    target.parent.mkdir(parents=True, exist_ok=True)
    normalized = content.replace("\r\n", "\n").replace("\r", "\n")
    target.write_bytes(normalized.replace("\n", "\r\n").encode("cp1252"))


def replace_once(text: str, before: str, after: str, label: str) -> str:
    if text.count(before) != 1:
        raise RuntimeError(f"{label}: expected exactly one migration anchor, found {text.count(before)}")
    return text.replace(before, after, 1)


def statement_source(lines: list[str], node: ast.AST) -> str:
    start = getattr(node, "lineno")
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.decorator_list:
        start = min([start, *(item.lineno for item in node.decorator_list)])
    end = getattr(node, "end_lineno")
    return "".join(lines[start - 1 : end])


def assigned_names(node: ast.AST) -> set[str]:
    result: set[str] = set()
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
        result.add(node.name)
    elif isinstance(node, (ast.Assign, ast.AnnAssign)):
        targets = node.targets if isinstance(node, ast.Assign) else [node.target]
        for target in targets:
            for child in ast.walk(target):
                if isinstance(child, ast.Name):
                    result.add(child.id)
    return result


def split_case_builder(path: str, function_name: str, helper_prefix: str, chunk_size: int = 9) -> None:
    source = read(path)
    tree = ast.parse(source, filename=path)
    target = next(
        node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == function_name
    )
    lines = source.splitlines(keepends=True)
    case_def = next(
        node for node in target.body if isinstance(node, ast.FunctionDef) and node.name == "case"
    )
    cases_stmt = target.body[0]
    fixtures = [
        node
        for node in target.body
        if isinstance(node, ast.FunctionDef)
        and any(
            isinstance(dec, ast.Call)
            and isinstance(dec.func, ast.Name)
            and dec.func.id == "case"
            for dec in node.decorator_list
        )
    ]
    if len(fixtures) < chunk_size * 2:
        raise RuntimeError(f"{path}:{function_name}: fixture inventory unexpectedly small")

    setup_nodes: list[ast.AST] = []
    setup_by_name: dict[str, ast.AST] = {}
    for node in target.body:
        if node in {cases_stmt, case_def} or node in fixtures or isinstance(node, ast.Return):
            continue
        names = assigned_names(node)
        if names:
            setup_nodes.append(node)
            for name in names:
                setup_by_name[name] = node

    chunks = [fixtures[index : index + chunk_size] for index in range(0, len(fixtures), chunk_size)]
    helpers: list[str] = []
    for index, chunk in enumerate(chunks, start=1):
        loaded = {
            child.id
            for fixture in chunk
            for child in ast.walk(fixture)
            if isinstance(child, ast.Name) and isinstance(child.ctx, ast.Load)
        }
        required_setup: list[ast.AST] = []
        seen: set[int] = set()
        for node in setup_nodes:
            if any(name in loaded for name in assigned_names(node)) and id(node) not in seen:
                required_setup.append(node)
                seen.add(id(node))
        body = ""
        for node in required_setup:
            body += statement_source(lines, node)
            if not body.endswith("\n\n"):
                body += "\n"
        for node in chunk:
            body += statement_source(lines, node)
            if not body.endswith("\n\n"):
                body += "\n"
        helpers.append(
            f"def {helper_prefix}_{index}(module: ModuleType, case) -> None:\n" + body + "\n"
        )

    signature = lines[target.lineno - 1]
    cases_source = statement_source(lines, cases_stmt)
    case_source = statement_source(lines, case_def)
    calls = "".join(
        f"    {helper_prefix}_{index}(module, case)\n" for index in range(1, len(chunks) + 1)
    )
    aggregator = signature + cases_source + "\n" + case_source + "\n" + calls + "    return cases\n"
    replacement = "\n".join(helpers).rstrip() + "\n\n\n" + aggregator
    original = "".join(lines[target.lineno - 1 : target.end_lineno])
    updated = source[: sum(len(line) for line in lines[: target.lineno - 1])] + replacement + source[sum(len(line) for line in lines[: target.end_lineno]):]
    if function_name not in updated or helper_prefix + "_1" not in updated:
        raise RuntimeError(f"{path}:{function_name}: split failed")
    write(path, updated)


def add_free_configuration_locals(section: str) -> str:
    wrapped = "def _probe(document, failures):\n" + section
    tree = ast.parse(wrapped)
    loads = {
        node.id for node in ast.walk(tree) if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load)
    }
    stores = {
        node.id for node in ast.walk(tree) if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store)
    }
    prefix = ""
    for name in ("mode", "profile", "repository"):
        if name in loads and name not in stores:
            prefix += f"    {name} = document.get(\"{name}\")\n"
    return prefix + section


def refactor_load_configuration() -> None:
    path = "tools/check_repo.py"
    source = read(path)
    tree = ast.parse(source, filename=path)
    target = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "load_configuration")
    lines = source.splitlines(keepends=True)
    function = "".join(lines[target.lineno - 1 : target.end_lineno])
    root_marker = '    if document.get("schema_version") != SCHEMA_VERSION:\n'
    profiles_marker = '    profiles = document.get("profiles")\n'
    placeholders_marker = '    placeholders = document.get("placeholders")\n'
    identity_marker = '    identity = document.get("identity")\n'
    vba_marker = '    vba = document.get("vba")\n'
    return_marker = '    return document, rule_result(\n'
    for marker in (root_marker, profiles_marker, placeholders_marker, identity_marker, vba_marker, return_marker):
        if marker not in function:
            raise RuntimeError(f"{path}: configuration marker missing: {marker.strip()}")

    i_root = function.index(root_marker)
    i_profiles = function.index(profiles_marker)
    i_placeholders = function.index(placeholders_marker)
    i_identity = function.index(identity_marker)
    i_vba = function.index(vba_marker)
    i_return = function.index(return_marker)
    prefix = function[:i_root]
    sections = [
        ("_validate_configuration_root", function[i_root:i_profiles]),
        ("_validate_configuration_profiles", function[i_profiles:i_placeholders]),
        ("_validate_configuration_placeholders", function[i_placeholders:i_identity]),
        ("_validate_configuration_identity", function[i_identity:i_vba]),
        ("_validate_configuration_vba", function[i_vba:i_return]),
    ]
    helpers = []
    for name, section in sections:
        section = add_free_configuration_locals(section)
        helpers.append(
            f"def {name}(document: dict[str, Any], failures: list[dict[str, Any]]) -> None:\n"
            + section
            + "\n"
        )
    return_block = function[i_return:]
    calls = "".join(f"    {name}(document, failures)\n" for name, _ in sections)
    wrapper = prefix + calls + "\n" + return_block
    replacement = "\n".join(helpers).rstrip() + "\n\n\n" + wrapper
    start = sum(len(line) for line in lines[: target.lineno - 1])
    end = sum(len(line) for line in lines[: target.end_lineno])
    write(path, source[:start] + replacement + source[end:])


def refactor_release_source() -> None:
    path = "tools/check_release.py"
    source = read(path)
    tree = ast.parse(source, filename=path)
    target = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "_validate_source")
    lines = source.splitlines(keepends=True)
    function = "".join(lines[target.lineno - 1 : target.end_lineno])
    markers = [
        '    mode = configuration.get("mode")\n',
        '    head = _git_output(root, "rev-parse", "HEAD")\n',
        '    if require_tag_ref:\n',
        '    if release_profile in GENERATED_PROFILES:\n',
        '    return release_profile, findings\n',
    ]
    for marker in markers:
        if marker not in function:
            raise RuntimeError(f"{path}: release-source marker missing: {marker.strip()}")
    i_profile, i_head, i_tag, i_generated, i_return = [function.index(item) for item in markers]
    profile = function[i_profile:i_head] + "    return release_profile\n"
    candidate = function[i_head:i_tag]
    tag = function[i_tag:i_generated]
    generated = function[i_generated:i_return]
    helpers = (
        "def _resolve_release_profile(\n"
        "    root: Path, configuration: dict[str, Any], findings: list[dict[str, str]]\n"
        ") -> str | None:\n" + profile + "\n\n"
        "def _validate_candidate_state(\n"
        "    root: Path, candidate_sha: str, findings: list[dict[str, str]]\n"
        ") -> None:\n" + candidate + "\n\n"
        "def _validate_tag_reference(\n"
        "    root: Path, tag: str, candidate_sha: str, require_tag_ref: bool, findings: list[dict[str, str]]\n"
        ") -> None:\n" + tag + "\n\n"
        "def _validate_generated_source(\n"
        "    root: Path, configuration: dict[str, Any], policy: dict[str, Any],\n"
        "    release_profile: str | None, findings: list[dict[str, str]]\n"
        ") -> None:\n" + generated + "\n\n"
    )
    header_end = function.index("    findings: list[dict[str, str]] = []\n")
    header = function[:header_end]
    wrapper = (
        header
        + "    findings: list[dict[str, str]] = []\n"
        + "    release_profile = _resolve_release_profile(root, configuration, findings)\n"
        + "    _validate_candidate_state(root, candidate_sha, findings)\n"
        + "    _validate_tag_reference(root, tag, candidate_sha, require_tag_ref, findings)\n"
        + "    _validate_generated_source(root, configuration, policy, release_profile, findings)\n"
        + "    return release_profile, findings\n"
    )
    replacement = helpers.rstrip() + "\n\n\n" + wrapper
    start = sum(len(line) for line in lines[: target.lineno - 1])
    end = sum(len(line) for line in lines[: target.end_lineno])
    write(path, source[:start] + replacement + source[end:])


def refactor_release_evidence() -> None:
    path = "tools/check_release.py"
    source = read(path)
    tree = ast.parse(source, filename=path)
    target = next(
        node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "_validate_evidence_and_assets"
    )
    lines = source.splitlines(keepends=True)
    function = "".join(lines[target.lineno - 1 : target.end_lineno])
    checks_marker = '    distribution = evidence.get("distribution")\n'
    assets_marker = '    assets = evidence.get("assets")\n'
    final_marker = '    return findings\n'
    for marker in (checks_marker, assets_marker, final_marker):
        if marker not in function:
            raise RuntimeError(f"{path}: release-evidence marker missing: {marker.strip()}")
    i_checks = function.index(checks_marker)
    i_assets = function.index(assets_marker)
    i_final = function.rindex(final_marker)
    prefix = function[:i_checks]
    checks = function[i_checks:i_assets] + "    return distribution\n"
    assets = function[i_assets:i_final]
    assets = assets.replace("        return findings\n", "        return\n")
    helpers = (
        "def _validate_evidence_checks(\n"
        "    evidence: dict[str, Any], policy: dict[str, Any], profile: str | None,\n"
        "    candidate_sha: str, findings: list[dict[str, str]]\n"
        ") -> object:\n" + checks + "\n\n"
        "def _validate_release_assets(\n"
        "    root: Path, evidence: dict[str, Any], manifest: dict[str, str] | None,\n"
        "    manifest_path: Path | None, policy: dict[str, Any], profile: str | None,\n"
        "    candidate_sha: str, distribution: object, findings: list[dict[str, str]]\n"
        ") -> None:\n" + assets + "\n\n"
    )
    wrapper = (
        prefix
        + "    distribution = _validate_evidence_checks(\n"
        + "        evidence, policy, profile, candidate_sha, findings\n"
        + "    )\n"
        + "    _validate_release_assets(\n"
        + "        root, evidence, manifest, manifest_path, policy, profile,\n"
        + "        candidate_sha, distribution, findings\n"
        + "    )\n"
        + "    return findings\n"
    )
    replacement = helpers.rstrip() + "\n\n\n" + wrapper
    start = sum(len(line) for line in lines[: target.lineno - 1])
    end = sum(len(line) for line in lines[: target.end_lineno])
    write(path, source[:start] + replacement + source[end:])


def update_vba() -> None:
    error_codes = '''Attribute VB_Name = "ProjectErrorCodes"\nOption Explicit\nOption Private Module\n\n'===============================================================================\n' MODULE: ProjectErrorCodes\n'-------------------------------------------------------------------------------\n' RESPONSIBILITY\n'   Own project-wide error-number constants shared by the facade and core.\n'\n' PUBLIC SURFACE\n'   None outside this VBA project. Constants are Public only for compile-time\n'   sharing across project modules; Option Private Module hides this module.\n'\n' DEPENDENCIES\n'   VBA runtime only. This module depends on no facade, core, test, workbook, or\n'   optional reference.\n'===============================================================================\n\nPublic Const ERROR_ZERO_DENOMINATOR As Long = vbObjectError + 2048\n'''
    write_vba("src/core/ProjectErrorCodes.bas", error_codes)

    core = read("src/core/ProjectCore.bas")
    core = core.replace(
        "'   VBA runtime only. This core never depends on the facade, tests, examples,\n'   workbook objects, or optional references.\n",
        "'   ProjectErrorCodes and the VBA runtime only. This core never depends on the\n'   facade, tests, examples, workbook objects, or optional references.\n",
    )
    core = core.replace("\nPrivate Const ERR_ZERO_DENOMINATOR As Long = vbObjectError + 2048\n", "")
    core = core.replace("            ERR_ZERO_DENOMINATOR, _", "            ProjectErrorCodes.ERROR_ZERO_DENOMINATOR, _")
    write_vba("src/core/ProjectCore.bas", core)

    facade = read("src/modules/ProjectFacade.bas")
    facade = facade.replace(
        "'   ProjectCore only. The dependency direction is facade -> core.\n",
        "'   ProjectCore and ProjectErrorCodes. Production dependencies remain inward:\n'   facade -> core/shared contracts; core never depends on the facade.\n",
    )
    facade = facade.replace(
        "Public Const PROJECT_ERROR_ZERO_DENOMINATOR As Long = vbObjectError + 2048",
        "Public Const PROJECT_ERROR_ZERO_DENOMINATOR As Long = ProjectErrorCodes.ERROR_ZERO_DENOMINATOR",
    )
    write_vba("src/modules/ProjectFacade.bas", facade)

    tests = read("tests/modules/ProjectTests.bas")
    tests = tests.replace(
        "'   RunProjectTests is the single documented test entry point.\n",
        "'   RunProjectTests executes the suite. ResetProjectTests is the documented\n'   recovery entry point after an interrupted run leaves module state active.\n",
    )
    if "Public Sub ResetProjectTests()" not in tests:
        tests = tests.replace(
            "Public Sub RunProjectTests()\n",
            "Public Sub ResetProjectTests()\n    ResetRun\nEnd Sub\n\nPublic Sub RunProjectTests()\n",
            1,
        )
    tests = tests.replace(
        "Private Sub ResetRun()\n    mCaseCount = 0\n",
        "Private Sub ResetRun()\n    mRunActive = False\n    mCaseCount = 0\n",
        1,
    )
    write_vba("tests/modules/ProjectTests.bas", tests)

    api = read("docs/PUBLIC_API.txt").replace(
        "Public Const PROJECT_ERROR_ZERO_DENOMINATOR As Long = vbObjectError + 2048",
        "Public Const PROJECT_ERROR_ZERO_DENOMINATOR As Long = ProjectErrorCodes.ERROR_ZERO_DENOMINATOR",
    )
    write("docs/PUBLIC_API.txt", api)

    tests_readme = read("tests/README.md")
    tests_readme = tests_readme.replace(
        "Import `modules/ProjectTests.bas` after `ProjectCore` and `ProjectFacade`, compile",
        "Import `modules/ProjectTests.bas` after `ProjectErrorCodes`, `ProjectCore`, and `ProjectFacade`, compile",
    )
    recovery = (
        "\nIf execution is interrupted with Ctrl+Break or the VBE Stop command before cleanup, "
        "run `ProjectTests.ResetProjectTests` once before starting the suite again. The reset "
        "clears only harness-owned module state and changes no Excel state.\n"
    )
    anchor = "failure is non-passing. The harness changes no Excel state; cleanup verifies its\nowned module state only.\n"
    if recovery.strip() not in tests_readme and anchor in tests_readme:
        tests_readme = tests_readme.replace(anchor, anchor + recovery)
    write("tests/README.md", tests_readme)


def update_profile() -> None:
    path = ".github/repository-profile.json"
    data = json.loads(read(path))
    required = set(data["required_paths"])
    required.add("pyproject.toml")
    data["required_paths"] = sorted(required, key=lambda item: (item.casefold(), item))

    meta = {
        ".github/workflows/checker-development.yml",
        "docs/CHECKER_DEVELOPMENT.md",
        "tools/check_policy_coverage.py",
        "tools/checker_development.py",
        "tools/policy_coverage_cases_config.py",
        "tools/policy_coverage_cases_quality.py",
        "tools/policy_coverage_cases_repo.py",
        "tools/policy_coverage_core.py",
        "tools/policy_coverage_runner.py",
    }
    template_only = set(data["placeholders"]["template_only_paths"]) | meta
    data["placeholders"]["template_only_paths"] = sorted(
        template_only, key=lambda item: (item.casefold(), item)
    )
    data["vba"]["components"]["src/core/ProjectErrorCodes.bas"] = "internal"
    data["vba"]["components"] = dict(
        sorted(data["vba"]["components"].items(), key=lambda item: (item[0].casefold(), item[0]))
    )
    for profile in data["profiles"].values():
        components = profile["vba_contract"]["required_components"]
        components["src/core/ProjectErrorCodes.bas"] = "internal"
        profile["vba_contract"]["required_components"] = dict(
            sorted(components.items(), key=lambda item: (item[0].casefold(), item[0]))
        )
    write(path, json.dumps(data, indent=2) + "\n")


def write_pyproject() -> None:
    content = '''[tool.ruff]\ntarget-version = "py310"\n\n[tool.ruff.lint]\nselect = ["E4", "E7", "E9", "F", "I", "C90"]\n\n[tool.ruff.lint.mccabe]\nmax-complexity = 20\n\n[tool.mypy]\npython_version = "3.10"\ncheck_untyped_defs = true\nno_implicit_optional = true\nwarn_redundant_casts = true\nwarn_unused_configs = true\nwarn_unused_ignores = true\nshow_error_codes = true\n'''
    write("pyproject.toml", content)


def update_static_workflow() -> None:
    path = ".github/workflows/static-checks.yml"
    text = read(path)
    checkout_anchor = "          persist-credentials: false\n          fetch-depth: 0\n\n"
    python_steps = '''      - name: Set up the minimum supported Python runtime\n        id: python-runtime\n        continue-on-error: true\n        uses: actions/setup-python@5fda3b95a4ea91299a34e894583c3862153e4b97 # v7.0.0\n        with:\n          python-version: "3.10"\n          cache: pip\n\n      - name: Install pinned Python quality tools\n        id: python-tools\n        continue-on-error: true\n        run: |\n          mkdir -p test-results\n          python --version | tee test-results/python-version.txt\n          python -m pip install --disable-pip-version-check 'mypy==1.18.2' 'ruff==0.13.2'\n\n      - name: Enforce Python lint, complexity, and typing\n        id: python-quality\n        continue-on-error: true\n        run: |\n          failed=0\n          {\n            echo "Python: $(python --version 2>&1)"\n            echo "Ruff: $(ruff --version)"\n            echo "Mypy: $(mypy --version)"\n            ruff check tools || failed=1\n            mypy tools/*.py || failed=1\n          } | tee test-results/python-quality.txt\n          {\n            echo "## Python quality"\n            echo\n            echo '```text'\n            cat test-results/python-quality.txt\n            echo '```'\n          } > test-results/python-quality.md\n          exit "$failed"\n\n'''
    if "Set up the minimum supported Python runtime" not in text:
        text = replace_once(text, checkout_anchor, checkout_anchor + python_steps, path)

    start = text.index("      - name: Exercise policy-branch coverage determinism\n")
    end = text.index("      - name: Exercise positive and degraded checker paths\n", start)
    text = text[:start] + text[end:]
    for line in (
        "              test-results/policy-coverage.md \\\n",
        "            test-results/policy-coverage.json\n",
        "            test-results/policy-coverage.md\n",
        "          POLICY_COVERAGE_SELF_TEST_OUTCOME: ${{ steps.policy-coverage-self-test.outcome }}\n",
        "          POLICY_COVERAGE_OUTCOME: ${{ steps.policy-coverage.outcome }}\n",
        "            \"Policy-coverage self-test:$POLICY_COVERAGE_SELF_TEST_OUTCOME\" \\\n",
        "            \"Policy coverage:$POLICY_COVERAGE_OUTCOME\" \\\n",
    ):
        text = text.replace(line, "")

    report_anchor = "              test-results/static-checks.md \\\n"
    if "test-results/python-quality.md" not in text:
        text = text.replace(report_anchor, report_anchor + "              test-results/python-quality.md \\\n", 1)
    artifact_anchor = "            test-results/static-checks.json\n"
    if "test-results/python-quality.txt" not in text:
        text = text.replace(
            artifact_anchor,
            artifact_anchor
            + "            test-results/python-version.txt\n"
            + "            test-results/python-quality.txt\n"
            + "            test-results/python-quality.md\n",
            1,
        )
    env_anchor = "          INITIALIZER_SELF_TEST_OUTCOME: ${{ steps.initializer-self-test.outcome }}\n"
    if "PYTHON_RUNTIME_OUTCOME" not in text:
        text = text.replace(
            env_anchor,
            "          PYTHON_RUNTIME_OUTCOME: ${{ steps.python-runtime.outcome }}\n"
            + "          PYTHON_TOOLS_OUTCOME: ${{ steps.python-tools.outcome }}\n"
            + "          PYTHON_QUALITY_OUTCOME: ${{ steps.python-quality.outcome }}\n"
            + env_anchor,
            1,
        )
    loop_anchor = "            \"Authoritative-validator installation:$VALIDATOR_INSTALL_OUTCOME\" \\\n"
    if "Python minimum runtime" not in text:
        text = text.replace(
            loop_anchor,
            "            \"Python minimum runtime:$PYTHON_RUNTIME_OUTCOME\" \\\n"
            + "            \"Python quality-tool installation:$PYTHON_TOOLS_OUTCOME\" \\\n"
            + "            \"Python lint/complexity/typing:$PYTHON_QUALITY_OUTCOME\" \\\n"
            + loop_anchor,
            1,
        )
    write(path, text)


def write_checker_workflow() -> None:
    content = '''name: Checker development contract\nrun-name: Checker development — ${{ github.ref_name }} @ ${{ github.sha }}\n\non:\n  pull_request:\n    branches:\n      - main\n  workflow_dispatch:\n\npermissions:\n  contents: read\n\nconcurrency:\n  group: checker-development-${{ github.ref }}\n  cancel-in-progress: true\n\ndefaults:\n  run:\n    shell: bash\n\njobs:\n  checker-development:\n    name: Checker development\n    runs-on: ubuntu-24.04\n    timeout-minutes: 10\n    env:\n      PYTHONUTF8: "1"\n      PYTHONDONTWRITEBYTECODE: "1"\n    steps:\n      - name: Check out the exact SHA under review\n        uses: actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1 # v7.0.1\n        with:\n          persist-credentials: false\n          fetch-depth: 0\n\n      - name: Set up Python 3.10\n        id: python-runtime\n        continue-on-error: true\n        uses: actions/setup-python@5fda3b95a4ea91299a34e894583c3862153e4b97 # v7.0.0\n        with:\n          python-version: "3.10"\n          cache: pip\n\n      - name: Install pinned development analyzers\n        id: python-tools\n        continue-on-error: true\n        run: |\n          mkdir -p test-results\n          python -m pip install --disable-pip-version-check 'mypy==1.18.2' 'ruff==0.13.2'\n\n      - name: Validate template-development Python\n        id: python-quality\n        continue-on-error: true\n        run: |\n          failed=0\n          {\n            ruff check tools || failed=1\n            mypy tools/*.py || failed=1\n          } | tee test-results/checker-python-quality.txt\n          exit "$failed"\n\n      - name: Exercise independent checker development contract\n        id: checker-self-test\n        continue-on-error: true\n        run: python tools/checker_development.py --root . --self-test\n\n      - name: Publish checker development evidence\n        id: checker-evidence\n        continue-on-error: true\n        run: >-\n          python tools/checker_development.py\n          --root .\n          --output test-results/checker-development.json\n          --summary test-results/checker-development.md\n\n      - name: Exercise policy-coverage determinism\n        id: policy-self-test\n        continue-on-error: true\n        run: python tools/check_policy_coverage.py --root . --self-test\n\n      - name: Publish policy-coverage evidence\n        id: policy-evidence\n        continue-on-error: true\n        run: >-\n          python tools/check_policy_coverage.py\n          --root .\n          --output test-results/policy-coverage.json\n          --summary test-results/policy-coverage.md\n\n      - name: Add readable summary\n        id: summary\n        if: always()\n        continue-on-error: true\n        run: |\n          for report in test-results/checker-development.md test-results/policy-coverage.md; do\n            if [[ ! -f "$report" ]]; then\n              echo "Required report missing: $report"\n              exit 1\n            fi\n            cat "$report" >> "$GITHUB_STEP_SUMMARY"\n            echo >> "$GITHUB_STEP_SUMMARY"\n          done\n\n      - name: Upload checker-development evidence\n        id: artifact\n        if: always()\n        continue-on-error: true\n        uses: actions/upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a # v7.0.1\n        with:\n          name: checker-development-${{ github.run_id }}-${{ github.run_attempt }}\n          path: |\n            test-results/checker-development.json\n            test-results/checker-development.md\n            test-results/policy-coverage.json\n            test-results/policy-coverage.md\n            test-results/checker-python-quality.txt\n          if-no-files-found: error\n          retention-days: 30\n          compression-level: 9\n\n      - name: Enforce checker-development outcomes\n        if: always()\n        env:\n          PYTHON_RUNTIME_OUTCOME: ${{ steps.python-runtime.outcome }}\n          PYTHON_TOOLS_OUTCOME: ${{ steps.python-tools.outcome }}\n          PYTHON_QUALITY_OUTCOME: ${{ steps.python-quality.outcome }}\n          CHECKER_SELF_TEST_OUTCOME: ${{ steps.checker-self-test.outcome }}\n          CHECKER_EVIDENCE_OUTCOME: ${{ steps.checker-evidence.outcome }}\n          POLICY_SELF_TEST_OUTCOME: ${{ steps.policy-self-test.outcome }}\n          POLICY_EVIDENCE_OUTCOME: ${{ steps.policy-evidence.outcome }}\n          SUMMARY_OUTCOME: ${{ steps.summary.outcome }}\n          ARTIFACT_OUTCOME: ${{ steps.artifact.outcome }}\n        run: |\n          failed=0\n          for item in \\\n            "Python runtime:$PYTHON_RUNTIME_OUTCOME" \\\n            "Python analyzers:$PYTHON_TOOLS_OUTCOME" \\\n            "Python quality:$PYTHON_QUALITY_OUTCOME" \\\n            "Checker self-test:$CHECKER_SELF_TEST_OUTCOME" \\\n            "Checker evidence:$CHECKER_EVIDENCE_OUTCOME" \\\n            "Policy self-test:$POLICY_SELF_TEST_OUTCOME" \\\n            "Policy evidence:$POLICY_EVIDENCE_OUTCOME" \\\n            "Summary:$SUMMARY_OUTCOME" \\\n            "Evidence upload:$ARTIFACT_OUTCOME"; do\n            name="${item%%:*}"\n            outcome="${item#*:}"\n            if [[ "$outcome" != "success" ]]; then\n              echo "::error::$name outcome: $outcome"\n              failed=1\n            fi\n          done\n          exit "$failed"\n'''
    write(".github/workflows/checker-development.yml", content)


def update_docs() -> None:
    docs = read("docs/README.md")
    row = "| Portable checker development boundaries and independent tests | [`CHECKER_DEVELOPMENT.md`](CHECKER_DEVELOPMENT.md) | Link from tooling/contribution guidance |\n"
    if row in docs and "<!-- template:remove:start -->\n" + row not in docs:
        docs = docs.replace(row, "<!-- template:remove:start -->\n" + row + "<!-- template:remove:end -->\n")
    write("docs/README.md", docs)

    contributing = read("CONTRIBUTING.md")
    paragraph = (
        "For checker changes, additionally follow\n"
        "[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md). For release-evidence\n"
        "schemas and exact-SHA binding, use\n"
        "[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).\n"
    )
    if paragraph in contributing:
        replacement = (
            "<!-- template:remove:start -->\n"
            "For template-checker changes, additionally follow\n"
            "[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md).\n"
            "<!-- template:remove:end -->\n"
            "For release-evidence schemas and exact-SHA binding, use\n"
            "[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).\n"
        )
        contributing = contributing.replace(paragraph, replacement)
    row2 = "| Checker changes | [`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md) |\n"
    if row2 in contributing and "<!-- template:remove:start -->\n" + row2 not in contributing:
        contributing = contributing.replace(row2, "<!-- template:remove:start -->\n" + row2 + "<!-- template:remove:end -->\n")
    write("CONTRIBUTING.md", contributing)

    checker_doc = read("docs/CHECKER_DEVELOPMENT.md")
    if "## 🧹 Generated-project boundary" not in checker_doc:
        checker_doc = checker_doc.replace(
            "## 🧭 Internal boundaries\n",
            "## 🧹 Generated-project boundary\n\n"
            "The checker-development workflow, policy-coverage harness, and this document are template-maintainer assurance. The initializer removes them from generated repositories. Generated projects retain only runtime/project gates they can execute meaningfully, including `check_repo.py`, focused static gates, release validation, `_gatelib.py`, and workflow validation.\n\n"
            "## 🧭 Internal boundaries\n",
        )
    write("docs/CHECKER_DEVELOPMENT.md", checker_doc)

    initialization = read("docs/INITIALIZATION.md")
    initialization = initialization.replace(
        "- deletes the temporary portfolio audit and implementation plan;\n",
        "- deletes template-maintainer audit, implementation, checker-development, and policy-coverage tooling that generated projects do not need;\n",
    )
    write("docs/INITIALIZATION.md", initialization)


def update_changelog() -> None:
    path = "CHANGELOG.md"
    text = read(path)
    changed_anchor = "### Fixed\n"
    changed_bullets = (
        "- Declared Python 3.10 as the minimum tooling runtime and added pinned ruff/mypy enforcement, including a McCabe complexity ceiling and unused-import detection, to hosted validation.\n"
        "- Moved template-only checker-development and policy-coverage meta-tooling out of generated projects while retaining runtime/project gates needed by initialized repositories.\n"
        "- Made checker-development CI fail closed with retained JSON/Markdown evidence and terminal outcome enforcement.\n\n"
    )
    if "Declared Python 3.10 as the minimum tooling runtime" not in text:
        text = text.replace(changed_anchor, changed_bullets + changed_anchor, 1)
    fixed_insert = (
        "- Centralized the zero-denominator VBA error number in the internal `ProjectErrorCodes` module so facade and core compile against one constant without reversing the facade-to-core dependency direction.\n"
        "- Added `ProjectTests.ResetProjectTests` and made `ResetRun` clear the re-entry flag so an interrupted test run has an explicit recovery path.\n"
        "- Reduced the worst Python control-flow hotspots by splitting policy-fixture registration and checker/release validation into bounded helpers without changing their CLI or evidence contracts.\n"
    )
    fixed_existing = "### Fixed\n\n"
    if "Centralized the zero-denominator VBA error number" not in text:
        text = text.replace(fixed_existing, fixed_existing + fixed_insert, 1)
    write(path, text)


def main() -> int:
    split_case_builder("tools/policy_coverage_cases_quality.py", "quality_cases", "_quality_case_group")
    split_case_builder("tools/policy_coverage_cases_repo.py", "repository_cases", "_repository_case_group")
    refactor_load_configuration()
    refactor_release_source()
    refactor_release_evidence()
    update_vba()
    update_profile()
    write_pyproject()
    update_static_workflow()
    write_checker_workflow()
    update_docs()
    update_changelog()

    for path in ("tools/check_repo.py", "tools/check_release.py", "tools/policy_coverage_cases_quality.py", "tools/policy_coverage_cases_repo.py"):
        ast.parse(read(path), filename=path)
    print("review-batch migration applied")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
