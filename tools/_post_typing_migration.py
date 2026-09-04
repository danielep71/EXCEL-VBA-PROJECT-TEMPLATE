#!/usr/bin/env python3
"""Repair the two contextual match names after the one-shot typing migration."""

from __future__ import annotations

import ast
from pathlib import Path


PATH = Path("tools/check_repo.py")


def replace_once(text: str, before: str, after: str) -> str:
    if text.count(before) != 1:
        raise RuntimeError(f"expected exactly one migration anchor, found {text.count(before)}")
    return text.replace(before, after, 1)


def main() -> int:
    text = PATH.read_text(encoding="utf-8")
    text = replace_once(
        text,
        '                        f"Template placeholders are prohibited in executable or VBA files: {token}",\n'
        '                        line_number(text, aws_match.start()),\n',
        '                        f"Template placeholders are prohibited in executable or VBA files: {token}",\n'
        '                        line_number(text, match.start()),\n',
    )
    text = replace_once(
        text,
        '                    "Possible AWS access key is tracked.",\n'
        '                    line_number(text, match.start()),\n',
        '                    "Possible AWS access key is tracked.",\n'
        '                    line_number(text, aws_match.start()),\n',
    )
    ast.parse(text)
    PATH.write_text(text, encoding="utf-8", newline="\n")
    print("post-migration contextual match repair applied")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
