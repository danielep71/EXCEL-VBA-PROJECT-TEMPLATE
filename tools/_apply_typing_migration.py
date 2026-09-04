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


def main() -> int:
    changed: list[str] = []
    for path in sorted(TOOLS.glob("*.py")):
        if path.name == Path(__file__).name:
            continue
        original = path.read_text(encoding="utf-8")
        updated = original
        for before, after in REPLACEMENTS:
            updated = updated.replace(before, after)
        if updated == original:
            continue
        updated = _add_any_import(updated)
        ast.parse(updated)
        path.write_text(updated, encoding="utf-8", newline="\n")
        changed.append(path.as_posix())
    print(f"typing migration changed {len(changed)} file(s)")
    for path in changed:
        print(path)
    if not changed:
        raise SystemExit("typing migration produced no changes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
