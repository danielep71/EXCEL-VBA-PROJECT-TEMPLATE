#!/usr/bin/env python3
"""Validate explicit VBA public API declarations and the checked-in manifest."""

from __future__ import annotations

import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from typing import Any

from _gatelib import git_bytes as git, parse_report_args as parse_args, write_text

CONFIG_PATH = ".github/repository-profile.json"
TOOL_NAME = "VBA public API"
VBA_SUFFIXES = {".bas", ".cls", ".frm"}
SIG_PREFIX = "# SIG\t"
PROPERTY_KINDS = {"property get", "property let", "property set"}

PROC = re.compile(
    r"^\s*(?:(?P<vis>Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z_]\w*)\b",
    re.I,
)
END_PROC = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
BLOCK = re.compile(
    r"^\s*Public\s+(?P<kind>Enum|Type)\s+(?P<name>[A-Za-z_]\w*)\b",
    re.I,
)
END_BLOCK = {
    "enum": re.compile(r"^\s*End\s+Enum\b", re.I),
    "type": re.compile(r"^\s*End\s+Type\b", re.I),
}
DECLARE = re.compile(
    r"^\s*(?:Public|Global)\s+Declare\s+(?:PtrSafe\s+)?"
    r"(?P<kind>Function|Sub)\s+(?P<name>[A-Za-z_]\w*)\b",
    re.I,
)
EVENT = re.compile(
    r"^\s*(?:Public|Global)\s+Event\s+(?P<name>[A-Za-z_]\w*)\b",
    re.I,
)
CONST = re.compile(
    r"^\s*(?:Public|Global)\s+Const\s+(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$",
    re.I,
)
VARIABLE = re.compile(
    r"^\s*(?:Public|Global)\s+(?P<withevents>WithEvents\s+)?"
    r"(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$",
    re.I,
)
IMPLICIT = re.compile(
    r"^\s*(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|"
    r"Enum|Type|Event|Declare\s+(?:PtrSafe\s+)?(?:Function|Sub))\b",
    re.I,
)


def tracked_vba(root: Path) -> list[str]:
    completed = git(root, "ls-files", "-z")
    if completed.returncode:
        raise RuntimeError(
            completed.stderr.decode("utf-8", errors="replace").strip()
        )
    return sorted(
        path.decode("utf-8", errors="surrogateescape")
        for path in completed.stdout.split(b"\0")
        if path
        and Path(path.decode("utf-8", errors="surrogateescape")).suffix.casefold()
        in VBA_SUFFIXES
    )


def strip_vba(raw: str) -> str:
    output: list[str] = []
    in_string = False
    index = 0
    while index < len(raw):
        character = raw[index]
        if character == '"':
            if in_string and index + 1 < len(raw) and raw[index + 1] == '"':
                output.extend(('"', '"'))
                index += 2
                continue
            in_string = not in_string
        elif character == "'" and not in_string:
            break
        output.append(character)
        index += 1
    text = "".join(output)
    return "" if re.match(r"^\s*Rem(?:\s|$)", text, re.I) else text.rstrip()


def logical(lines: list[str]) -> list[tuple[int, int, str]]:
    result: list[tuple[int, int, str]] = []
    buffer: list[str] = []
    start = 0
    for number, raw in enumerate(lines, 1):
        code = strip_vba(raw)
        if not buffer:
            start = number
        if re.search(r"\s_\s*$", code):
            buffer.append(re.sub(r"\s_\s*$", " ", code))
            continue
        buffer.append(code)
        result.append(
            (start, number, " ".join(part.strip() for part in buffer))
        )
        buffer.clear()
    if buffer:
        result.append(
            (start, len(lines), " ".join(part.strip() for part in buffer))
        )
    return result


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def key(component: str, kind: str, name: str) -> str:
    return f"{component}\t{kind}\t{name}"


def top_comma(text: str) -> bool:
    depth = 0
    in_string = False
    index = 0
    while index < len(text):
        character = text[index]
        if character == '"':
            if in_string and index + 1 < len(text) and text[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string:
            if character == "(":
                depth += 1
            elif character == ")" and depth:
                depth -= 1
            elif character == "," and depth == 0:
                return True
        index += 1
    return False


def record(
    output: list[dict[str, Any]],
    supported: bool,
    component: str,
    kind: str,
    name: str,
    signature: str,
    path: str,
    line: int,
) -> None:
    if supported:
        output.append(
            {
                "component": component,
                "kind": kind,
                "name": name,
                "signature": signature,
                "path": path,
                "line": line,
            }
        )


def parse_component(
    path: str,
    text: str,
    supported: bool,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    component = Path(path).stem
    declarations: list[dict[str, Any]] = []
    findings: list[dict[str, Any]] = []
    statements = logical(text.splitlines())
    inside_procedure = False
    index = 0

    while index < len(statements):
        line, _, code = statements[index]
        stripped = code.strip()
        if not stripped or stripped.startswith(("Attribute ", "Option ", "#")):
            index += 1
            continue

        match = PROC.match(code)
        if match and " declare " not in f" {code.casefold()} ":
            visibility = (match.group("vis") or "").casefold()
            if not visibility:
                findings.append(
                    {
                        "path": path,
                        "line": line,
                        "message": (
                            "Implicit public procedure is prohibited: "
                            f"{match.group('name')}."
                        ),
                    }
                )
            if visibility == "public":
                kind = " ".join(
                    part.capitalize() for part in match.group("kind").split()
                )
                record(
                    declarations,
                    supported,
                    component,
                    kind,
                    match.group("name"),
                    norm(code),
                    path,
                    line,
                )
            inside_procedure = True
            index += 1
            continue

        if END_PROC.match(code):
            inside_procedure = False
            index += 1
            continue
        if inside_procedure:
            index += 1
            continue

        if IMPLICIT.match(code):
            findings.append(
                {
                    "path": path,
                    "line": line,
                    "message": (
                        "Implicit public module-level declaration is prohibited: "
                        f"{norm(code)}"
                    ),
                }
            )
            index += 1
            continue

        match = BLOCK.match(code)
        if match:
            kind = match.group("kind").capitalize()
            name = match.group("name")
            body = [norm(code)]
            close = END_BLOCK[kind.casefold()]
            closing_index = index + 1
            while (
                closing_index < len(statements)
                and not close.match(statements[closing_index][2])
            ):
                if statements[closing_index][2].strip():
                    body.append(norm(statements[closing_index][2]))
                closing_index += 1
            if closing_index >= len(statements):
                findings.append(
                    {
                        "path": path,
                        "line": line,
                        "message": f"Public {kind} {name} is not closed.",
                    }
                )
                index += 1
                continue
            body.append(norm(statements[closing_index][2]))
            record(
                declarations,
                supported,
                component,
                kind,
                name,
                " | ".join(body),
                path,
                line,
            )
            index = closing_index + 1
            continue

        match = DECLARE.match(code)
        if match:
            kind = (
                "Declare Function"
                if match.group("kind").casefold() == "function"
                else "Declare Sub"
            )
            record(
                declarations,
                supported,
                component,
                kind,
                match.group("name"),
                norm(code),
                path,
                line,
            )
            index += 1
            continue

        match = EVENT.match(code)
        if match:
            record(
                declarations,
                supported,
                component,
                "Event",
                match.group("name"),
                norm(code),
                path,
                line,
            )
            index += 1
            continue

        match = CONST.match(code)
        if match:
            if top_comma(match.group("rest")):
                findings.append(
                    {
                        "path": path,
                        "line": line,
                        "message": (
                            "Public Const declarations must contain one identifier "
                            "per statement."
                        ),
                    }
                )
            else:
                record(
                    declarations,
                    supported,
                    component,
                    "Const",
                    match.group("name"),
                    norm(code),
                    path,
                    line,
                )
            index += 1
            continue

        match = VARIABLE.match(code)
        if match:
            if top_comma(match.group("rest")):
                findings.append(
                    {
                        "path": path,
                        "line": line,
                        "message": (
                            "Public variable declarations must contain one identifier "
                            "per statement."
                        ),
                    }
                )
            else:
                kind = (
                    "WithEvents Variable"
                    if match.group("withevents")
                    else "Variable"
                )
                record(
                    declarations,
                    supported,
                    component,
                    kind,
                    match.group("name"),
                    norm(code),
                    path,
                    line,
                )
            index += 1
            continue

        index += 1

    return declarations, findings


def read_manifest(
    root: Path,
    relative: str,
) -> tuple[set[str], dict[str, str], list[dict[str, Any]]]:
    path = root / relative
    if not path.is_file():
        return (
            set(),
            {},
            [{"path": relative, "message": "Configured public API manifest is missing."}],
        )

    rows: set[str] = set()
    signatures: dict[str, str] = {}
    findings: list[dict[str, Any]] = []
    for number, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not raw.strip():
            continue
        if raw.startswith(SIG_PREFIX):
            fields = raw[len(SIG_PREFIX) :].split("\t", 3)
            if len(fields) != 4:
                findings.append(
                    {"path": relative, "line": number, "message": "Malformed # SIG record."}
                )
                continue
            declaration_key = key(fields[0], fields[1], fields[2])
            if any(
                existing.casefold() == declaration_key.casefold()
                for existing in signatures
            ):
                findings.append(
                    {
                        "path": relative,
                        "line": number,
                        "message": f"Duplicate signature record: {declaration_key}",
                    }
                )
            signatures[declaration_key] = fields[3]
            continue
        if raw.lstrip().startswith("#"):
            continue

        fields = raw.split("\t")
        if len(fields) != 3 or any(not field for field in fields):
            findings.append(
                {
                    "path": relative,
                    "line": number,
                    "message": (
                        "Manifest declaration rows require component, kind, and name."
                    ),
                }
            )
            continue
        declaration_key = key(*fields)
        if any(
            existing.casefold() == declaration_key.casefold() for existing in rows
        ):
            findings.append(
                {
                    "path": relative,
                    "line": number,
                    "message": f"Duplicate manifest declaration: {declaration_key}",
                }
            )
        rows.add(declaration_key)
    return rows, signatures, findings


def is_property(kind: str) -> bool:
    return kind.casefold() in PROPERTY_KINDS


def property_pair(left: dict[str, Any], right: dict[str, Any]) -> bool:
    return (
        str(left["component"]).casefold() == str(right["component"]).casefold()
        and is_property(str(left["kind"]))
        and is_property(str(right["kind"]))
        and str(left["kind"]).casefold() != str(right["kind"]).casefold()
    )


def run_check(root: Path) -> dict[str, Any]:
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    vba = config["vba"]
    components: dict[str, str] = vba["components"]
    manifest = vba["public_api_manifest"]
    findings: