#!/usr/bin/env python3
"""Temporary one-shot source migration for the v1.1.0 typing cleanup.

This file is removed after the generated source changes are committed.
"""

from __future__ import annotations

import ast
from pathlib import Path


TOOLS = Path("tools")
REPLACEMENTS = (
    ("dict[str, object]", "dict[str, Any]"),
)


def _has_any_import(text: str) -> bool:
    tree = ast.parse(text)
    return any(
        isinstance(node, ast.ImportFrom)
        and node.module == "typing"
        and any(alias.name == "Any" for alias in node.names)
        for node in tree.body
    )


def _add_any_import(text: str) -> str:
    if _has_any_import(text):
        return text
    lines = text.splitlines(keepends=True)
    for index, line in enumerate(lines):
        if line.startswith("from typing import ("):
            lines.insert(index + 1, "    Any,\n")
            return "".join(lines)
        if line.startswith("from typing import "):
            names = line.removeprefix("from typing import ").rstrip("\n").split(", ")
            names.append("Any")
            names = sorted(set(names))
            newline = "\n" if line.endswith("\n") else ""
            lines[index] = "from typing import " + ", ".join(names) + newline
            return "".join(lines)
    insert_at = 0
    for index, line in enumerate(lines):
        if line.startswith("import ") or line.startswith("from "):
            insert_at = index + 1
    lines.insert(insert_at, "from typing import Any\n")
    return "".join(lines)


def _replace(text: str, before: str, after: str, *, path: str) -> str:
    if before not in text:
        raise RuntimeError(f"{path}: expected migration anchor not found: {before[:80]!r}")
    return text.replace(before, after, 1)


def _specific_repairs(path: Path, text: str) -> str:
    name = path.name

    if name == "check_vba_public_api.py":
        text = _replace(
            text,
            "    tests = []\n",
            "    tests: list[tuple[str, dict[str, Any], str, str | None]] = []\n",
            path=name,
        )

    elif name == "check_vba_conditionals.py":
        text = _replace(
            text,
            "            value = int(token)\n            if value not in {-1, 0}:\n",
            "            integer_value = int(token)\n            if integer_value not in {-1, 0}:\n",
            path=name,
        )
        text = _replace(
            text,
            "            return value == -1\n",
            "            return integer_value == -1\n",
            path=name,
        )
        text = _replace(
            text,
            "    stacks = {name: [] for name in ENVIRONMENTS}\n",
            "    stacks: dict[str, list[Frame]] = {name: [] for name in ENVIRONMENTS}\n",
            path=name,
        )

    elif name == "check_release.py":
        start = text.index("    if release_profile in GENERATED_PROFILES:\n")
        end = text.index("        changelog_path = root / \"CHANGELOG.md\"\n", start)
        section = text[start:end]
        section = section.replace("template_tokens: list[str] = []", "source_template_tokens: list[str] = []")
        section = section.replace("template_tokens.extend(", "source_template_tokens.extend(")
        section = section.replace("for token in template_tokens:", "for token in source_template_tokens:")
        text = text[:start] + section + text[end:]
        text = _replace(
            text,
            "        if not _safe_relative(relative):\n",
            "        if not isinstance(relative, str) or not _safe_relative(relative):\n",
            path=name,
        )
        text = _replace(
            text,
            "        target = root / PurePosixPath(relative)\n",
            "        target = root / relative\n",
            path=name,
        )
        text = _replace(
            text,
            "            context = {\n",
            "            context: dict[str, Any] = {\n",
            path=name,
        )
        text = _replace(
            text,
            "        negative(\n            \"missing-asset-manifest\",\n            lambda c: (binary_mutation(c, approved=True, digest_matches=True), c.update(manifest_path=None)),\n            \"missing-asset-manifest\", profile=\"application\",\n        )\n",
            "        def missing_manifest_mutation(context) -> None:\n            binary_mutation(context, approved=True, digest_matches=True)\n            context.update(manifest_path=None)\n\n        negative(\n            \"missing-asset-manifest\",\n            missing_manifest_mutation,\n            \"missing-asset-manifest\", profile=\"application\",\n        )\n",
            path=name,
        )

    elif name == "policy_coverage_core.py":
        text = _replace(
            text,
            "    if fixture_boundary is None:\n        raise CoverageError(\"Cannot locate the canonical fixture boundary in check_repo.py\")\n\n    sites:",
            "    if fixture_boundary is None:\n        raise CoverageError(\"Cannot locate the canonical fixture boundary in check_repo.py\")\n    boundary = fixture_boundary\n\n    sites:",
            path=name,
        )
        text = _replace(
            text,
            "        def visit_FunctionDef(self, node: ast.FunctionDef) -> None:\n            self.functions.append(node.name)\n            self.generic_visit(node)\n            self.functions.pop()\n\n        visit_AsyncFunctionDef = visit_FunctionDef\n",
            "        def _visit_function(\n            self, node: ast.FunctionDef | ast.AsyncFunctionDef\n        ) -> None:\n            self.functions.append(node.name)\n            self.generic_visit(node)\n            self.functions.pop()\n\n        def visit_FunctionDef(self, node: ast.FunctionDef) -> None:\n            self._visit_function(node)\n\n        def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef) -> None:\n            self._visit_function(node)\n",
            path=name,
        )
        text = text.replace("node.lineno < fixture_boundary", "node.lineno < boundary")

    elif name == "check_repo.py":
        for before, after in (
            (
                "    else:\n        pattern = placeholders.get(\"pattern\")\n",
                "    else:\n        assert isinstance(placeholders, dict)\n        pattern = placeholders.get(\"pattern\")\n",
            ),
            (
                "    else:\n        forbidden_tokens = _string_list(\n            identity.get(\"forbidden_tokens\"),\n",
                "    else:\n        assert isinstance(identity, dict)\n        forbidden_tokens = _string_list(\n            identity.get(\"forbidden_tokens\"),\n",
            ),
            (
                "    else:\n        source_roots = _string_list(\n            vba.get(\"source_roots\"), \"vba.source_roots\", failures, paths=True\n",
                "    else:\n        assert isinstance(vba, dict)\n        source_roots = _string_list(\n            vba.get(\"source_roots\"), \"vba.source_roots\", failures, paths=True\n",
            ),
            (
                "    else:\n        profiles = overlays.get(\"profile\")\n",
                "    else:\n        assert isinstance(overlays, dict)\n        profiles = overlays.get(\"profile\")\n",
            ),
        ):
            text = _replace(text, before, after, path=name)
        text = _replace(
            text,
            "        match = reference.match(line)\n        if match:\n            yield number, match.group(1).strip()\n",
            "        reference_match = reference.match(line)\n        if reference_match:\n            yield number, reference_match.group(1).strip()\n",
            path=name,
        )
        text = _replace(
            text,
            "        match = aws_key.search(text)\n        if match:\n",
            "        aws_match = aws_key.search(text)\n        if aws_match:\n",
            path=name,
        )
        text = _replace(
            text,
            "                    line_number(text, match.start()),\n",
            "                    line_number(text, aws_match.start()),\n",
            path=name,
        )

    elif name == "policy_coverage_cases_repo.py":
        text = _replace(
            text,
            "    cases: list[tuple[str, str, str | None, Callable[[Path], None]]] = []\n\n",
            "    cases: list[tuple[str, str, str | None, Callable[[Path], None]]] = []\n\n    def template_mode(document: dict) -> None:\n        document.update(\n            mode=\"template\",\n            profile=None,\n            repository=\"example/TEMPLATE-IDENTITY\",\n        )\n\n",
            path=name,
        )
        text = text.replace(
            "mutate_config(module, root, lambda d: (d.__setitem__(\"mode\", \"template\"), d.__setitem__(\"profile\", None), d.__setitem__(\"repository\", \"example/TEMPLATE-IDENTITY\")))",
            "mutate_config(module, root, template_mode)",
        )

    elif name == "policy_coverage_cases_config.py":
        text = _replace(
            text,
            "    cases: list[tuple[str, str, str | None, Callable[[Path], None]]] = []\n\n",
            "    cases: list[tuple[str, str, str | None, Callable[[Path], None]]] = []\n\n    def assign(**values: object) -> Callable[[dict], None]:\n        def mutate(document: dict) -> None:\n            document.update(values)\n        return mutate\n\n",
            path=name,
        )
        text = text.replace(
            'lambda d: (d.__setitem__("mode", "template"), d.__setitem__("profile", "library"))',
            'assign(mode="template", profile="library")',
        )
        text = text.replace(
            'lambda d: (d.__setitem__("mode", "template"), d.__setitem__("profile", None), d.__setitem__("repository", "example/TEMPLATE-IDENTITY"), d.__setitem__("label_domains", ["domain"]))',
            'assign(mode="template", profile=None, repository="example/TEMPLATE-IDENTITY", label_domains=["domain"])',
        )
        text = text.replace(
            'lambda d: (d.__setitem__("mode", "template"), d.__setitem__("profile", None), d.__setitem__("repository", "example/plain"))',
            'assign(mode="template", profile=None, repository="example/plain")',
        )

    elif name == "initialize_repository.py":
        text = _replace(
            text,
            "            mode = stat.S_IMODE(originals[path][1]) if originals[path] is not None else 0o644\n",
            "            original = originals[path]\n            mode = stat.S_IMODE(original[1]) if original is not None else 0o644\n",
            path=name,
        )

    elif name == "policy_coverage_runner.py":
        text = _replace(
            text,
            "    module = load_module(root / CORE_TOOL, \"coverage_check_repo\")\n",
            "    module: Any = load_module(root / CORE_TOOL, \"coverage_check_repo\")\n",
            path=name,
        )

    return text


def main() -> int:
    changed: list[str] = []
    for path in sorted(TOOLS.glob("*.py")):
        if path.name == Path(__file__).name:
            continue
        original = path.read_text(encoding="utf-8")
        updated = original
        for before, after in REPLACEMENTS:
            updated = updated.replace(before, after)
        updated = _specific_repairs(path, updated)
        if updated == original:
            continue
        if "Any" in updated and not _has_any_import(updated):
            updated = _add_any_import(updated)
        ast.parse(updated)
        path.write_text(updated, encoding="utf-8", newline="\n")
        changed.append(path.as_posix())
    print(f"typing migration changed {len(changed)} file(s)")
    for changed_path in changed:
        print(changed_path)
    if not changed:
        raise SystemExit("typing migration produced no changes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
