#!/usr/bin/env python3
"""Validate reachable VBA declarations across supported compilation environments."""

from __future__ import annotations

import argparse
import ast
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile

VBA_SUFFIXES = {".bas", ".cls", ".frm"}
TOOL_NAME = "VBA conditional compilation"
ENVIRONMENTS = (
    ("vba6-win32", {"VBA6": True, "VBA7": False, "WIN32": True, "WIN64": False}),
    ("vba7-win32", {"VBA6": False, "VBA7": True, "WIN32": True, "WIN64": False}),
    ("vba7-win64", {"VBA6": False, "VBA7": True, "WIN32": False, "WIN64": True}),
)
SUPPORTED_SYMBOLS = {"VBA6", "VBA7", "WIN32", "WIN64"}
DECLARE_RE = re.compile(
    r"^\s*(?:(?:Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\b",
    re.IGNORECASE,
)
PTRSAFE_RE = re.compile(r"\bPtrSafe\b", re.IGNORECASE)
IF_RE = re.compile(r"^\s*#If\s+(.+?)\s+Then\s*$", re.IGNORECASE)
ELSEIF_RE = re.compile(r"^\s*#ElseIf\s+(.+?)\s+Then\s*$", re.IGNORECASE)
ELSE_RE = re.compile(r"^\s*#Else\s*$", re.IGNORECASE)
END_RE = re.compile(r"^\s*#End\s+If\s*$", re.IGNORECASE)
CONST_RE = re.compile(r"^\s*#Const\b", re.IGNORECASE)


class ExpressionError(ValueError):
    pass


def tracked_vba(root: Path) -> list[str]:
    completed = subprocess.run(
        ["git", "-C", str(root), "ls-files", "-z"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    result: list[str] = []
    for item in completed.stdout.split(b"\0"):
        if not item:
            continue
        relative = item.decode("utf-8", errors="surrogateescape")
        if Path(relative).suffix.casefold() in VBA_SUFFIXES:
            result.append(relative)
    return sorted(result)


def strip_vba(raw: str) -> str:
    result: list[str] = []
    in_string = False
    index = 0
    while index < len(raw):
        char = raw[index]
        if char == '"':
            if in_string and index + 1 < len(raw) and raw[index + 1] == '"':
                result.extend(('"', '"'))
                index += 2
                continue
            in_string = not in_string
            result.append(char)
            index += 1
            continue
        if char == "'" and not in_string:
            break
        result.append(char)
        index += 1
    text = "".join(result)
    if re.match(r"^\s*Rem(?:\s|$)", text, re.IGNORECASE):
        return ""
    return text.rstrip()


def logical_units(lines: list[str]) -> list[tuple[int, int, str, str]]:
    units: list[tuple[int, int, str, str]] = []
    buffer: list[str] = []
    start = 0
    for number, raw in enumerate(lines, start=1):
        code = strip_vba(raw)
        stripped = code.strip()
        if not buffer and stripped.startswith("#"):
            units.append((number, number, "directive", stripped))
            continue
        if not buffer:
            start = number
        if re.search(r"\s_\s*$", code):
            buffer.append(re.sub(r"\s_\s*$", " ", code))
            continue
        buffer.append(code)
        units.append((start, number, "code", " ".join(part.strip() for part in buffer)))
        buffer.clear()
    if buffer:
        units.append((start, len(lines), "code", " ".join(part.strip() for part in buffer)))
    return units


def _python_expression(expression: str) -> str:
    translated = re.sub(r"\bAnd\b", " and ", expression, flags=re.IGNORECASE)
    translated = re.sub(r"\bOr\b", " or ", translated, flags=re.IGNORECASE)
    translated = re.sub(r"\bNot\b", " not ", translated, flags=re.IGNORECASE)
    translated = translated.replace("<>", "!=")
    translated = re.sub(r"(?<![<>=!])=(?!=)", "==", translated)
    return translated.strip()


def evaluate(expression: str, symbols: dict[str, bool]) -> bool:
    try:
        tree = ast.parse(_python_expression(expression), mode="eval")
    except SyntaxError as error:
        raise ExpressionError(f"invalid expression syntax: {error.msg}") from error

    def visit(node: ast.AST) -> bool | int:
        if isinstance(node, ast.Expression):
            return visit(node.body)
        if isinstance(node, ast.Name):
            key = node.id.upper()
            if key == "TRUE":
                return True
            if key == "FALSE":
                return False
            if key not in SUPPORTED_SYMBOLS:
                raise ExpressionError(
                    f"unsupported symbol {node.id!r}; supported symbols are VBA6, VBA7, Win32, Win64"
                )
            return symbols[key]
        if isinstance(node, ast.Constant) and isinstance(node.value, (bool, int)):
            return node.value
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.Not):
            return not bool(visit(node.operand))
        if isinstance(node, ast.BoolOp) and isinstance(node.op, (ast.And, ast.Or)):
            values = [bool(visit(value)) for value in node.values]
            return all(values) if isinstance(node.op, ast.And) else any(values)
        if isinstance(node, ast.Compare) and len(node.ops) == len(node.comparators) == 1:
            left = visit(node.left)
            right = visit(node.comparators[0])
            if isinstance(node.ops[0], ast.Eq):
                return left == right
            if isinstance(node.ops[0], ast.NotEq):
                return left != right
        raise ExpressionError(f"unsupported expression element: {type(node).__name__}")

    return bool(visit(tree))


def analyze_environment(
    path: str,
    units: list[tuple[int, int, str, str]],
    environment: str,
    symbols: dict[str, bool],
) -> list[dict[str, object]]:
    findings: list[dict[str, object]] = []
    stack: list[dict[str, object]] = []

    def active() -> bool:
        return bool(stack[-1]["active"]) if stack else True

    def evaluate_or_report(expression: str, line: int, directive: str) -> bool:
        try:
            return evaluate(expression, symbols)
        except ExpressionError as error:
            findings.append(
                {
                    "path": path,
                    "line": line,
                    "environment": environment,
                    "expression": expression,
                    "message": f"Unsupported or indeterminate {directive} expression: {error}.",
                }
            )
            return False

    for start_line, end_line, kind, code in units:
        if kind == "directive":
            if CONST_RE.match(code):
                findings.append(
                    {
                        "path": path,
                        "line": start_line,
                        "environment": environment,
                        "message": (
                            "Project-defined #Const symbols are outside the supported model; "
                            "conditional-compilation evaluation fails closed."
                        ),
                    }
                )
                continue

            match = IF_RE.match(code)
            if match:
                parent_active = active()
                condition = evaluate_or_report(match.group(1), start_line, "#If")
                stack.append(
                    {
                        "line": start_line,
                        "parent_active": parent_active,
                        "branch_taken": condition,
                        "active": parent_active and condition,
                        "else_seen": False,
                    }
                )
                continue

            match = ELSEIF_RE.match(code)
            if match:
                if not stack:
                    findings.append(
                        {"path": path, "line": start_line, "environment": environment,
                         "message": "#ElseIf has no matching #If."}
                    )
                    continue
                frame = stack[-1]
                if frame["else_seen"]:
                    findings.append(
                        {"path": path, "line": start_line, "environment": environment,
                         "message": "#ElseIf cannot follow #Else in the same block."}
                    )
                    frame["active"] = False
                    continue
                condition = evaluate_or_report(match.group(1), start_line, "#ElseIf")
                selected = not bool(frame["branch_taken"]) and condition
                frame["active"] = bool(frame["parent_active"]) and selected
                frame["branch_taken"] = bool(frame["branch_taken"]) or condition
                continue

            if ELSE_RE.match(code):
                if not stack:
                    findings.append(
                        {"path": path, "line": start_line, "environment": environment,
                         "message": "#Else has no matching #If."}
                    )
                    continue
                frame = stack[-1]
                if frame["else_seen"]:
                    findings.append(
                        {"path": path, "line": start_line, "environment": environment,
                         "message": "Conditional block contains more than one #Else."}
                    )
                    frame["active"] = False
                    continue
                frame["active"] = bool(frame["parent_active"]) and not bool(frame["branch_taken"])
                frame["branch_taken"] = True
                frame["else_seen"] = True
                continue

            if END_RE.match(code):
                if not stack:
                    findings.append(
                        {"path": path, "line": start_line, "environment": environment,
                         "message": "#End If has no matching #If."}
                    )
                else:
                    stack.pop()
                continue

            findings.append(
                {"path": path, "line": start_line, "environment": environment,
                 "message": f"Unsupported conditional-compilation directive: {code}"}
            )
            continue

        if active() and DECLARE_RE.match(code) and symbols["VBA7"] and not PTRSAFE_RE.search(code):
            findings.append(
                {
                    "path": path,
                    "line": start_line,
                    "end_line": end_line,
                    "environment": environment,
                    "message": "Declare reachable in a supported VBA7 environment must include PtrSafe.",
                }
            )

    for frame in stack:
        findings.append(
            {
                "path": path,
                "line": frame["line"],
                "environment": environment,
                "message": "#If block is not closed by #End If.",
            }
        )
    return findings


def run_check(root: Path) -> dict[str, object]:
    paths = tracked_vba(root)
    raw_findings: list[dict[str, object]] = []
    declare_count = 0
    for relative in paths:
        text = (root / relative).read_bytes().decode("cp1252")
        units = logical_units(text.splitlines())
        declare_count += sum(kind == "code" and bool(DECLARE_RE.match(code)) for _, _, kind, code in units)
        for environment, symbols in ENVIRONMENTS:
            raw_findings.extend(analyze_environment(relative, units, environment, symbols))

    grouped: dict[tuple[object, ...], dict[str, object]] = {}
    for item in raw_findings:
        key = (
            item.get("path"), item.get("line"), item.get("end_line"),
            item.get("expression"), item.get("message"),
        )
        if key not in grouped:
            grouped[key] = {k: v for k, v in item.items() if k != "environment"}
            grouped[key]["environments"] = []
        grouped[key]["environments"].append(item["environment"])

    findings = sorted(
        grouped.values(),
        key=lambda item: (str(item.get("path")), int(item.get("line", 0)), str(item.get("message"))),
    )
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "components": len(paths),
        "declare_statements": declare_count,
        "environments": [name for name, _ in ENVIRONMENTS],
        "findings": findings,
    }


def markdown_report(report: dict[str, object]) -> str:
    lines = [
        "## VBA conditional compilation",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Components:** {report['components']}",
        f"- **Declare statements:** {report['declare_statements']}",
        "- **Environments:** " + ", ".join(f"`{item}`" for item in report["environments"]),
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["findings"]:
        lines.extend(["", "| Path | Line | Environments | Finding |", "| --- | ---: | --- | --- |"])
        for item in report["findings"]:
            lines.append(
                "| {path} | {line} | {envs} | {message} |".format(
                    path=str(item.get("path", "")).replace("|", "\\|"),
                    line=item.get("line", "—"),
                    envs=", ".join(f"`{value}`" for value in item.get("environments", [])),
                    message=str(item.get("message", "")).replace("|", "\\|"),
                )
            )
    return "\n".join(lines) + "\n"


def write_text(path: Path | None, text: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def fixture_result(source: str) -> dict[str, object]:
    with tempfile.TemporaryDirectory(prefix="vba-conditional-") as temporary:
        root = Path(temporary)
        (root / "src").mkdir()
        (root / "src" / "Fixture.bas").write_bytes(source.replace("\n", "\r\n").encode("cp1252"))
        for command in (
            ("init", "-b", "main"),
            ("config", "user.name", "Conditional Self-Test"),
            ("config", "user.email", "conditional@example.invalid"),
            ("add", "src/Fixture.bas"),
        ):
            completed = subprocess.run(
                ["git", "-C", str(root), *command], stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            if completed.returncode != 0:
                raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
        return run_check(root)


def run_self_test() -> int:
    fixtures = {
        "valid-nested": ("pass", '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function Alpha Lib "kernel32" () As LongPtr
    #ElseIf Win32 Then
        Private Declare PtrSafe Function Alpha Lib "kernel32" () As Long
    #Else
        Private Declare PtrSafe Function Alpha Lib "kernel32" () As Long
    #End If
#Else
    Private Declare Function Alpha Lib "kernel32" () As Long
#End If
'''),
        "reachable-win64-bad": ("fail", '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 Then
    #If Win64 Then
        Private Declare Function Alpha Lib "kernel32" () As LongPtr
    #Else
        Private Declare PtrSafe Function Alpha Lib "kernel32" () As Long
    #End If
#Else
    Private Declare Function Alpha Lib "kernel32" () As Long
#End If
'''),
        "reachable-win32-bad": ("fail", '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 And Win32 Then
    Private Declare Function Alpha Lib "kernel32" () As Long
#ElseIf VBA7 And Win64 Then
    Private Declare PtrSafe Function Alpha Lib "kernel32" () As LongPtr
#Else
    Private Declare Function Alpha Lib "kernel32" () As Long
#End If
'''),
        "inactive-nested": ("pass", '''Attribute VB_Name = "Fixture"
Option Explicit
#If Not VBA7 Then
    #If Win64 Then
        Private Declare Function Alpha Lib "kernel32" () As Long
    #Else
        Private Declare Function Alpha Lib "kernel32" () As Long
    #End If
#Else
    Private Declare PtrSafe Function Alpha Lib "kernel32" () As LongPtr
#End If
'''),
        "unsupported-symbol": ("fail", '''Attribute VB_Name = "Fixture"
Option Explicit
#If Mac Then
    Private Declare PtrSafe Function Alpha Lib "kernel32" () As Long
#End If
'''),
        "unbalanced": ("fail", '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function Alpha Lib "kernel32" () As Long
'''),
        "continued-declare": ("pass", '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function Alpha Lib "kernel32" _
        () As Long
#Else
    Private Declare Function Alpha Lib "kernel32" () As Long
#End If
'''),
    }
    failures: list[str] = []
    for name, (expected, source) in fixtures.items():
        report = fixture_result(source)
        if report["status"] != expected:
            failures.append(f"{name}: expected {expected}, got {report['status']} ({report['findings']})")

    win64 = fixture_result(fixtures["reachable-win64-bad"][1])
    if not any(item.get("environments") == ["vba7-win64"] for item in win64["findings"]):
        failures.append("reachable-win64-bad: finding did not isolate vba7-win64")
    win32 = fixture_result(fixtures["reachable-win32-bad"][1])
    if not any(item.get("environments") == ["vba7-win32"] for item in win32["findings"]):
        failures.append("reachable-win32-bad: finding did not isolate vba7-win32")
    unsupported = fixture_result(fixtures["unsupported-symbol"][1])
    if not any("unsupported symbol" in item.get("message", "").casefold() for item in unsupported["findings"]):
        failures.append("unsupported-symbol: diagnostic did not name unsupported symbol")

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: nested VBA6/VBA7 and Win32/Win64 branches, ElseIf selection, "
        "inactive nesting, reachable PtrSafe failures, continued declares, unknown symbols, "
        "and unbalanced directives passed."
    )
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    if options.self_test:
        try:
            return run_self_test()
        except (OSError, UnicodeError, RuntimeError, subprocess.SubprocessError) as error:
            print(f"SELF-TEST ERROR: {error}", file=sys.stderr)
            return 2
    try:
        report = run_check(options.root)
        write_text(options.output, json.dumps(report, indent=2, sort_keys=True) + "\n")
        write_text(options.summary, markdown_report(report))
        print(markdown_report(report).rstrip())
        return 0 if report["status"] == "pass" else 1
    except (OSError, UnicodeError, RuntimeError, subprocess.SubprocessError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
