#!/usr/bin/env python3
"""Temporary one-shot refactor that extracts shared focused-gate primitives."""

from __future__ import annotations

import ast
import json
from pathlib import Path
import re

TOOLS = Path("tools")

GATELIB = '''#!/usr/bin/env python3
"""Private standard-library primitives shared by focused repository gates.

The canonical portable checker, ``check_repo.py``, intentionally does not import
this module: that file remains a self-contained distributable artifact.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import subprocess
from typing import Any


def git_bytes(
    root: Path, *args: str, check: bool = False
) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(
        ["git", "-C", str(root), *args],
        check=check,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def git_text(
    root: Path, *args: str, check: bool = False
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", "-C", str(root), *args],
        check=check,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )


def tracked_files(root: Path) -> set[str]:
    completed = git_bytes(root, "ls-files", "-z")
    if completed.returncode:
        raise RuntimeError(
            completed.stderr.decode("utf-8", errors="replace").strip()
        )
    return {
        item.decode("utf-8", errors="surrogateescape")
        for item in completed.stdout.split(b"\\0")
        if item
    }


def write_text(path: Path | None, text: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\\n")


def write_json(path: Path | None, value: object) -> None:
    write_text(path, json.dumps(value, indent=2, sort_keys=True) + "\\n")


def parse_report_args(
    argv: list[str], *, description: str | None = None
) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)
'''

TARGETS: dict[str, tuple[set[str], str]] = {
    "check_committed_whitespace.py": (
        {"git", "write_text"},
        "from _gatelib import git_text as git, write_text",
    ),
    "check_local_actions.py": (
        {"git", "tracked_files", "write_text", "parse_args"},
        "from _gatelib import git_bytes as git, parse_report_args as parse_args, tracked_files, write_text",
    ),
    "check_release_semantics.py": (
        {"write_text", "parse_args"},
        "from _gatelib import parse_report_args as parse_args, write_text",
    ),
    "check_vba_conditionals.py": (
        {"git", "write_text", "parse_args"},
        "from _gatelib import git_bytes as git, parse_report_args as parse_args, write_text",
    ),
    "check_vba_jumps.py": (
        {"git", "write_text", "parse_args"},
        "from _gatelib import git_bytes as git, parse_report_args as parse_args, write_text",
    ),
    "check_vba_public_api.py": (
        {"git", "write_text", "parse_args"},
        "from _gatelib import git_bytes as git, parse_report_args as parse_args, write_text",
    ),
    "checker_development.py": (
        {"write_text", "parse_arguments"},
        "from _gatelib import parse_report_args as parse_arguments, write_text",
    ),
    "policy_coverage_runner.py": (
        {"write_text", "parse_args"},
        "from _gatelib import parse_report_args as parse_args, write_text",
    ),
}

COMMON_PARSER_CONSUMERS = {
    "check_local_actions.py",
    "check_release_semantics.py",
    "check_vba_conditionals.py",
    "check_vba_jumps.py",
    "check_vba_public_api.py",
    "checker_development.py",
    "policy_coverage_runner.py",
}


def remove_top_level_functions(text: str, names: set[str]) -> str:
    tree = ast.parse(text)
    spans: list[tuple[int, int]] = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in names:
            if node.end_lineno is None:
                raise RuntimeError(f"missing end line for {node.name}")
            spans.append((node.lineno, node.end_lineno))
    missing = names - {
        node.name
        for node in tree.body
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }
    if missing:
        raise RuntimeError(f"functions not found: {sorted(missing)}")
    lines = text.splitlines(keepends=True)
    for start, end in sorted(spans, reverse=True):
        del lines[start - 1 : end]
    return "".join(lines)


def add_import(text: str, statement: str) -> str:
    if statement in text:
        return text
    tree = ast.parse(text)
    import_nodes = [
        node for node in tree.body if isinstance(node, (ast.Import, ast.ImportFrom))
    ]
    if not import_nodes:
        raise RuntimeError("no import block found")
    last = import_nodes[-1]
    if last.end_lineno is None:
        raise RuntimeError("import end line unavailable")
    lines = text.splitlines(keepends=True)
    lines.insert(last.end_lineno, statement + "\n")
    return "".join(lines)


def remove_unused_argparse(text: str) -> str:
    if "argparse." not in text:
        text = text.replace("import argparse\n", "")
    return text


def refactor_tool(path: Path, names: set[str], import_statement: str) -> None:
    text = path.read_text(encoding="utf-8")
    text = remove_top_level_functions(text, names)
    text = add_import(text, import_statement)
    text = remove_unused_argparse(text)
    ast.parse(text, filename=str(path))
    path.write_text(text, encoding="utf-8", newline="\n")


def rename_fixture_writer(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    text = re.sub(r"(?<!\.)\bwrite_text\(", "write_fixture_text(", text)
    ast.parse(text, filename=str(path))
    path.write_text(text, encoding="utf-8", newline="\n")


def update_profile(path: Path) -> None:
    document = json.loads(path.read_text(encoding="utf-8"))
    required = document["required_paths"]
    if "tools/_gatelib.py" not in required:
        required.append("tools/_gatelib.py")
        required.sort(key=lambda item: (item.casefold(), item))
    path.write_text(json.dumps(document, indent=2, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")


def update_checker_contract(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    text = text.replace(
        'CHECKER_PATH = Path("tools/check_repo.py")\n',
        'CHECKER_PATH = Path("tools/check_repo.py")\nGATELIB_PATH = Path("tools/_gatelib.py")\n',
        1,
    )
    anchor = '''def build_report(root: Path) -> dict[str, Any]:\n'''
    shared = '''def shared_library_report(root: Path) -> tuple[dict[str, Any], list[str]]:\n    failures: list[str] = []\n    gatelib = (root / GATELIB_PATH).resolve()\n    source = gatelib.read_text(encoding="utf-8")\n    tree = ast.parse(source, filename=str(gatelib))\n    import_roots: set[str] = set()\n    for node in ast.walk(tree):\n        if isinstance(node, ast.Import):\n            import_roots.update(alias.name.split(".", 1)[0] for alias in node.names)\n        elif isinstance(node, ast.ImportFrom):\n            if node.module is not None:\n                import_roots.add(node.module.split(".", 1)[0])\n    non_stdlib = sorted(\n        root_name\n        for root_name in import_roots\n        if root_name != "__future__" and root_name not in sys.stdlib_module_names\n    )\n    if non_stdlib:\n        failures.append("_gatelib.py has non-stdlib imports: " + ", ".join(non_stdlib))\n\n    checker_tree = ast.parse((root / CHECKER_PATH).read_text(encoding="utf-8"))\n    for node in ast.walk(checker_tree):\n        if isinstance(node, ast.ImportFrom) and node.module == "_gatelib":\n            failures.append("check_repo.py must remain independent of _gatelib.py")\n        elif isinstance(node, ast.Import) and any(alias.name == "_gatelib" for alias in node.names):\n            failures.append("check_repo.py must remain independent of _gatelib.py")\n\n    forbidden_helpers = {"git", "write_text", "tracked_files"}\n    local_helper_owners: list[str] = []\n    parser_consumers = {\n        "check_local_actions.py",\n        "check_release_semantics.py",\n        "check_vba_conditionals.py",\n        "check_vba_jumps.py",\n        "check_vba_public_api.py",\n        "checker_development.py",\n        "policy_coverage_runner.py",\n    }\n    parser_imports: list[str] = []\n    for tool in sorted((root / "tools").glob("*.py")):\n        if tool.name in {CHECKER_PATH.name, GATELIB_PATH.name}:\n            continue\n        tool_tree = ast.parse(tool.read_text(encoding="utf-8"), filename=str(tool))\n        definitions = {\n            node.name\n            for node in tool_tree.body\n            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))\n        }\n        duplicates = sorted(definitions & forbidden_helpers)\n        if duplicates:\n            local_helper_owners.append(f"{tool.name}: {', '.join(duplicates)}")\n        if tool.name in parser_consumers:\n            imported = any(\n                isinstance(node, ast.ImportFrom)\n                and node.module == "_gatelib"\n                and any(alias.name == "parse_report_args" for alias in node.names)\n                for node in tool_tree.body\n            )\n            if imported:\n                parser_imports.append(tool.name)\n            else:\n                failures.append(f"{tool.name} does not consume _gatelib.parse_report_args")\n    if local_helper_owners:\n        failures.append("shared helpers redefined outside _gatelib.py: " + "; ".join(local_helper_owners))\n\n    evidence = {\n        "path": GATELIB_PATH.as_posix(),\n        "imports": sorted(import_roots),\n        "parser_consumers": parser_imports,\n        "check_repo_independent": not any("check_repo.py must remain" in item for item in failures),\n    }\n    return evidence, failures\n\n\n'''
    if anchor not in text:
        raise RuntimeError("checker-development build_report anchor missing")
    text = text.replace(anchor, shared + anchor, 1)
    text = text.replace(
        '    ids, id_failures = check_ids(module)\n    failures.extend(id_failures)\n',
        '    ids, id_failures = check_ids(module)\n    failures.extend(id_failures)\n    shared_library, shared_failures = shared_library_report(root)\n    failures.extend(shared_failures)\n',
        1,
    )
    text = text.replace(
        '        "canonical_checks": ids,\n',
        '        "canonical_checks": ids,\n        "shared_library": shared_library,\n',
        1,
    )
    text = text.replace(
        '        f"- **Canonical policy checks:** {len(report[\'canonical_checks\'])}",\n',
        '        f"- **Canonical policy checks:** {len(report[\'canonical_checks\'])}",\n        f"- **Shared focused-gate library:** `{report[\'shared_library\'][\'path\']}`",\n',
        1,
    )
    text = text.replace(
        '"canonical check order, artifact identity, and standard-library-only runtime passed."\n',
        '"canonical check order, artifact identity, shared-helper ownership, and standard-library-only runtime passed."\n',
        1,
    )
    ast.parse(text, filename=str(path))
    path.write_text(text, encoding="utf-8", newline="\n")


def update_docs(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    old = '`tools/check_repo.py` remains the canonical portable checker delivered to generated repositories. The development model intentionally avoids a package-manager or multi-file runtime dependency: the reviewed source **is** the distributable artifact, so there is no bundle transform that can drift from source.\n'
    new = '`tools/check_repo.py` remains the canonical portable checker delivered to generated repositories. **The single-file identity rule applies to that canonical checker only.** Its reviewed source is the distributable artifact, so there is no bundle transform that can drift from source. Focused sibling gates are development/runtime tools within the repository and may share private standard-library infrastructure.\n\n## 🔧 Shared focused-gate primitives\n\n`tools/_gatelib.py` owns the small cross-tool mechanics that are genuinely identical: Git subprocess wrappers, tracked-file enumeration, deterministic UTF-8/LF report writes, and the common `--root` / `--output` / `--summary` / `--self-test` parser. Focused gates import those primitives instead of maintaining copies. Tool-specific `main`, `run_check`, `build_report`, `run_self_test`, and Markdown renderers remain local because their behavior and evidence schemas differ.\n\n`tools/check_repo.py` must never import `_gatelib.py`. Generated repositories retain `_gatelib.py` for the focused gates, while the canonical checker remains independently copyable and executable as one standard-library-only file. `checker_development.py` enforces this ownership boundary.\n'
    if old not in text:
        raise RuntimeError("checker development intro anchor missing")
    text = text.replace(old, new, 1)
    text = text.replace(
        '1. Change `tools/check_repo.py` and any focused P2 gate required by the policy change.\n',
        '1. Change `tools/check_repo.py`, `_gatelib.py`, and/or the focused gate that owns the affected behavior. Keep `check_repo.py` independent of `_gatelib.py`.\n',
        1,
    )
    path.write_text(text, encoding="utf-8", newline="\n")


def update_tools_readme(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    anchor = "##"
    note = '''\n### Shared focused-gate infrastructure\n\n`_gatelib.py` is the private, standard-library-only owner of Git, report-output, tracked-file, and common focused-gate CLI primitives. `check_repo.py` deliberately does not import it: the canonical checker remains a self-contained distributable artifact.\n'''
    if "### Shared focused-gate infrastructure" not in text:
        index = text.find(anchor)
        if index < 0:
            text += note
        else:
            text = text[:index] + note + "\n" + text[index:]
    path.write_text(text, encoding="utf-8", newline="\n")


def update_changelog(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    anchor = "### Compatibility\n"
    entry = "- Consolidated duplicated focused-gate Git, report-output, tracked-file, and common CLI mechanics into the private standard-library-only `tools/_gatelib.py`, while keeping `tools/check_repo.py` explicitly self-contained.\n\n"
    if entry not in text:
        pos = text.find(anchor)
        if pos < 0:
            raise RuntimeError("changelog compatibility anchor missing")
        text = text[:pos] + entry + text[pos:]
    path.write_text(text, encoding="utf-8", newline="\n")


def main() -> int:
    (TOOLS / "_gatelib.py").write_text(GATELIB, encoding="utf-8", newline="\n")
    for filename, (names, statement) in TARGETS.items():
        refactor_tool(TOOLS / filename, names, statement)
    rename_fixture_writer(TOOLS / "test_workflow_validation.py")
    update_profile(Path(".github/repository-profile.json"))
    update_checker_contract(TOOLS / "checker_development.py")
    update_docs(Path("docs/CHECKER_DEVELOPMENT.md"))
    update_tools_readme(TOOLS / "README.md")
    update_changelog(Path("CHANGELOG.md"))
    for path in [*TOOLS.glob("*.py")]:
        ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    print("gatelib refactor applied")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
