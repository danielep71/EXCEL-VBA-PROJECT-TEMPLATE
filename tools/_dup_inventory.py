#!/usr/bin/env python3
"""Temporary AST inventory for shared gate-helper extraction."""

from __future__ import annotations

import ast
import hashlib
from pathlib import Path

TARGETS = {
    "main", "markdown_report", "write_text", "run_self_test", "parse_args",
    "run_check", "git", "build_report", "parse_arguments", "_write_report",
    "tracked_files", "console_report", "_markdown_escape",
}


def signature(node: ast.FunctionDef | ast.AsyncFunctionDef) -> str:
    parts: list[str] = []
    for arg in [*node.args.posonlyargs, *node.args.args]:
        parts.append(arg.arg)
    if node.args.vararg:
        parts.append("*" + node.args.vararg.arg)
    for arg in node.args.kwonlyargs:
        parts.append(arg.arg)
    if node.args.kwarg:
        parts.append("**" + node.args.kwarg.arg)
    return "(" + ", ".join(parts) + ")"


def main() -> int:
    rows: list[tuple[str, str, str, str, str]] = []
    counts: dict[str, int] = {}
    for path in sorted(Path("tools").glob("*.py")):
        if path.name == Path(__file__).name:
            continue
        source = path.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(path))
        for node in tree.body:
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            counts[node.name] = counts.get(node.name, 0) + 1
            if node.name not in TARGETS:
                continue
            body_dump = ast.dump(ast.Module(body=node.body, type_ignores=[]), include_attributes=False)
            digest = hashlib.sha256(body_dump.encode()).hexdigest()[:12]
            segment = ast.get_source_segment(source, node) or ""
            excerpt = " ".join(segment.split())[:240]
            rows.append((node.name, path.name, signature(node), digest, excerpt))
    print("DUPLICATE TOP-LEVEL NAMES")
    for name, count in sorted(counts.items(), key=lambda item: (-item[1], item[0])):
        if count >= 3:
            print(f"{count:2d} {name}")
    print("\nTARGET DETAILS")
    for row in sorted(rows):
        print(" | ".join(row))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
