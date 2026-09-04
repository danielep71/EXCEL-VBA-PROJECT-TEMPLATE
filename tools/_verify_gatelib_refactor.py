#!/usr/bin/env python3
"""Temporary verification that shared gate helpers are actually consolidated."""

from __future__ import annotations

import ast
from collections import Counter
from pathlib import Path

EXCLUDED = {"check_repo.py", "_gatelib.py", Path(__file__).name, "_apply_gatelib_refactor.py", "_dup_inventory.py"}


def main() -> int:
    counts: Counter[str] = Counter()
    violations: list[str] = []
    for path in sorted(Path("tools").glob("*.py")):
        if path.name in EXCLUDED:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        names = {
            node.name
            for node in tree.body
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        }
        counts.update(names)
        duplicated_helpers = names & {"git", "write_text", "tracked_files"}
        if duplicated_helpers:
            violations.append(f"{path.name}: local shared helpers {sorted(duplicated_helpers)}")
    print("POST-REFACTOR TOP-LEVEL NAMES >=3")
    for name, count in sorted(counts.items(), key=lambda item: (-item[1], item[0])):
        if count >= 3:
            print(f"{count:2d} {name}")
    if violations:
        for item in violations:
            print(f"[FAIL] {item}")
        return 1
    print("SHARED HELPER OWNERSHIP PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
