#!/usr/bin/env python3
"""Validate explicit VBA public API declarations and the checked-in manifest."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile

CONFIG_PATH = ".github/repository-profile.json"
TOOL_NAME = "VBA public API"
VBA_SUFFIXES = {".bas", ".cls", ".frm"}
VISIBILITY = r"(?:Public|Private|Friend|Global)"
PROC_OPEN = re.compile(
    rf"^\s*(?:(?P<visibility>{VISIBILITY})\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?P<name>[A-Za-z_]\w*)\b",
    re.IGNORECASE,
)
PROC_CLOSE = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
PUBLIC_SIMPLE = re.compile(
    r"^\s*(?P<visibility>Public|Global)\s+"
    r"(?P<kind>Const|Event|Declare\s+(?:PtrSafe\s+)?(?:Function|Sub))\s+"
    r"(?P<name>[A-Za-z_]\w*)\b",
    re.IGNORECASE,
)
PUBLIC_BLOCK = re.compile(
    r"^\s*Public\s+(?P<kind>Enum|Type)\s+(?P<name>[A-Za-z_]\w*)\b",
    re.IGNORECASE,
)
BLOCK_CLOSE = {
    "enum": re.compile(r"^\s*End\s+Enum\b", re.IGNORECASE),
    "type": re.compile(r"^\s*End\s+Type\b", re.IGNORECASE),
}
PUBLIC_VARIABLE = re.compile(
    r"^\s*(?P<visibility>Public|Global)\s+(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$",
    re.IGNORECASE,
)
IMPLICIT_PUBLIC = re.compile(
    r"^\s*(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set)|Enum|Type|Event|Declare\s+(?:Function|Sub))\b",
    re.IGNORECASE,
)
SIG_PREFIX = "# SIG\t"


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(
        ["git", "-C", str(root), *args], stdout=subprocess.PIPE, stderr=subprocess.PIPE
    )


def tracked_vba(root: Path) -> list[str]:
    completed = git(root, "ls-files", "-z")
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


def logical_statements(lines: list[str]) -> list[tuple[int, int, str]]:
    statements: list[tuple[int, int, str]] = []
    buffer: list[str] = []
    start = 0
    for number, raw in enumerate(lines, start=1):
        code = strip_vba(raw)
        if not buffer:
            start = number
        if re.search(r"\s_\s*$", code):
            buffer.append(re.sub(r"\s_\s*$", " ", code))
            continue
        buffer.append(code)
        statements.append((start, number, " ".join(part.strip() for part in buffer)))
        buffer.clear()
    if buffer:
        statements.append((start, len(lines), " ".join(part.strip() for part in buffer)))
    return statements


def normalize_signature(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def declaration_key(component: str, kind: str, name: str) -> str:
    return f"{component}\t{kind}\t{name}"


def parse_component(path: str, text: str, supported: bool) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    component = Path(path).stem
    declarations: list[dict[str, object]] = []
    findings: list[dict[str, object]] = []
    statements = logical_statements(text.splitlines())
    inside_procedure = False
    index = 0

    while index < len(statements):
        start_line, end_line, code = statements[index]
        stripped = code.strip()
        if not stripped or stripped.startswith("Attribute ") or stripped.startswith("Option ") or stripped.startswith("#"):
            index += 1
            continue

        proc = PROC_OPEN.match(code)
        if proc and " declare " not in f" {code.casefold()} ":
            visibility = (proc.group("visibility") or "").casefold()
            if not visibility:
                findings.append({
                    "path": path, "line": start_line,
                    "message": f"Implicit public procedure is prohibited: {proc.group('name')}.",
                })
            if supported and visibility == "public":
                kind = " ".join(part.capitalize() for part in proc.group("kind").split())
                declarations.append({
                    "component": component, "kind": kind, "name": proc.group("name"),
                    "signature": normalize_signature(code), "path": path, "line": start_line,
                })
            inside_procedure = True
            index += 1
            continue
        if PROC_CLOSE.match(code):
            inside_procedure = False
            index += 1
            continue
        if inside_procedure:
            index += 1
            continue

        if IMPLICIT_PUBLIC.match(code):
            findings.append({
                "path": path, "line": start_line,
                "message": f"Implicit public module-level declaration is prohibited: {normalize_signature(code)}",
            })
            index += 1
            continue

        block = PUBLIC_BLOCK.match(code)
        if block:
            kind = block.group("kind").capitalize()
            name = block.group("name")
            body = [normalize_signature(code)]
            closer = BLOCK_CLOSE[kind.casefold()]
            cursor = index + 1
            while cursor < len(statements) and not closer.match(statements[cursor][2]):
                if statements[cursor][2].strip():
                    body.append(normalize_signature(statements[cursor][2]))
                cursor += 1
            if cursor >= len(statements):
                findings.append({"path": path, "line": start_line, "message": f"Public {kind} {name} is not closed."})
                index += 1
                continue
            body.append(normalize_signature(statements[cursor][2]))
            if supported:
                declarations.append({
                    "component": component, "kind": kind, "name": name,
                    "signature": " | ".join(body), "path": path, "line": start_line,
                })
            index = cursor + 1
            continue

        simple = PUBLIC_SIMPLE.match(code)
        if simple:
            raw_kind = simple.group("kind")
            if raw_kind.casefold().startswith("declare"):
                kind = "Declare Function" if "function" in raw_kind.casefold() else "Declare Sub"
            else:
                kind = raw_kind.capitalize()
            if supported:
                declarations.append({
                    "component": component, "kind": kind, "name": simple.group("name"),
                    "signature": normalize_signature(code), "path": path, "line": start_line,
                })
            index += 1
            continue

        variable = PUBLIC_VARIABLE.match(code)
        if variable:
            rest = variable.group("rest")
            if "," in rest:
                findings.append({
                    "path": path, "line": start_line,
                    "message": "Public variable declarations must contain one identifier per statement.",
                })
            elif supported:
                declarations.append({
                    "component": component, "kind": "Variable", "name": variable.group("name"),
                    "signature": normalize_signature(code), "path": path, "line": start_line,
                })
        index += 1

    return declarations, findings


def read_manifest(root: Path, relative: str) -> tuple[set[str], dict[str, str], list[dict[str, object]]]:
    findings: list[dict[str, object]] = []
    path = root / relative
    if not path.is_file():
        return set(), {}, [{"path": relative, "message": "Configured public API manifest is missing."}]
    rows: set[str] = set()
    signatures: dict[str, str] = {}
    for number, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        if not raw.strip():
            continue
        if raw.startswith(SIG_PREFIX):
            fields = raw[len(SIG_PREFIX):].split("\t", 3)
            if len(fields) != 4:
                findings.append({"path": relative, "line": number, "message": "Malformed # SIG record."})
                continue
            component, kind, name, signature = fields
            key = declaration_key(component, kind, name)
            if key in signatures:
                findings.append({"path": relative, "line": number, "message": f"Duplicate signature record: {key}"})
            signatures[key] = signature
            continue
        if raw.lstrip().startswith("#"):
            continue
        fields = raw.split("\t")
        if len(fields) != 3 or any(not field for field in fields):
            findings.append({"path": relative, "line": number, "message": "Manifest declaration rows require component, kind, and name."})
            continue
        key = declaration_key(*fields)
        if key in rows:
            findings.append({"path": relative, "line": number, "message": f"Duplicate manifest declaration: {key}"})
        rows.add(key)
    return rows, signatures, findings


def run_check(root: Path) -> dict[str, object]:
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    vba = config["vba"]
    components: dict[str, str] = vba["components"]
    manifest = vba["public_api_manifest"]
    failures: list[dict[str, object]] = []

    if manifest not in config["required_paths"]:
        failures.append({"path": CONFIG_PATH, "message": "public_api_manifest must be a global required_path for every profile."})
    for profile, entry in config["profiles"].items():
        minimum = entry["vba_contract"]["minimum_roles"].get("public", 0)
        if minimum < 1:
            failures.append({"path": CONFIG_PATH, "message": f"Profile {profile!r} must require at least one public-role component."})

    declarations: list[dict[str, object]] = []
    tracked = set(tracked_vba(root))
    for path, role in sorted(components.items()):
        if path not in tracked or not (root / path).is_file():
            continue
        parsed, component_findings = parse_component(path, (root / path).read_bytes().decode("cp1252"), role == "public")
        declarations.extend(parsed)
        failures.extend(component_findings)

    keys: dict[str, dict[str, object]] = {}
    global_names: dict[str, dict[str, object]] = {}
    for declaration in declarations:
        key = declaration_key(declaration["component"], declaration["kind"], declaration["name"])
        if key in keys:
            failures.append({"path": declaration["path"], "line": declaration["line"], "message": f"Public declaration appears more than once: {key}"})
        keys[key] = declaration
        path = str(declaration["path"])
        if Path(path).suffix.casefold() == ".bas":
            name_key = str(declaration["name"]).casefold()
            previous = global_names.get(name_key)
            if previous is not None:
                failures.append({
                    "path": path, "line": declaration["line"],
                    "message": (
                        f"Public standard-module name {declaration['name']!r} collides with "
                        f"{previous['component']}.{previous['name']}."
                    ),
                })
            else:
                global_names[name_key] = declaration

    rows, signatures, manifest_findings = read_manifest(root, manifest)
    failures.extend(manifest_findings)
    actual_rows = set(keys)
    for missing in sorted(actual_rows - rows, key=str.casefold):
        failures.append({"path": manifest, "message": f"Public declaration is not recorded: {missing}"})
    for stale in sorted(rows - actual_rows, key=str.casefold):
        failures.append({"path": manifest, "message": f"Manifest declaration is not present in supported source: {stale}"})
    for key in sorted(actual_rows, key=str.casefold):
        expected = str(keys[key]["signature"])
        recorded = signatures.get(key)
        if recorded is None:
            failures.append({"path": manifest, "message": f"Normalized signature record is missing: {key}"})
        elif recorded != expected:
            failures.append({"path": manifest, "message": f"Signature mismatch for {key}: expected {expected!r}, recorded {recorded!r}"})
    for stale in sorted(set(signatures) - actual_rows, key=str.casefold):
        failures.append({"path": manifest, "message": f"Stale signature record: {stale}"})

    evidence = [
        {
            "component": item["component"], "kind": item["kind"], "name": item["name"],
            "signature": item["signature"], "path": item["path"], "line": item["line"],
        }
        for item in sorted(declarations, key=lambda item: (str(item["component"]).casefold(), str(item["name"]).casefold(), str(item["kind"])))
    ]
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not failures else "fail",
        "manifest": manifest,
        "policy": {
            "required_profiles": sorted(config["profiles"]),
            "required_from": "initialization",
            "implicit_visibility": "prohibited",
            "public_variable_form": "one identifier per statement",
        },
        "declarations": evidence,
        "findings": failures,
    }


def markdown_report(report: dict[str, object]) -> str:
    lines = [
        "## VBA public API",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Manifest:** `{report['manifest']}`",
        f"- **Declarations:** {len(report['declarations'])}",
        "- **Profiles:** " + ", ".join(f"`{item}`" for item in report["policy"]["required_profiles"]),
        "- **Manifest required from:** initialization",
        "- **Implicit public visibility:** prohibited",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["declarations"]:
        lines.extend(["", "| Component | Kind | Name | Signature |", "| --- | --- | --- | --- |"])
        for item in report["declarations"]:
            lines.append(
                "| {component} | {kind} | {name} | `{signature}` |".format(
                    component=item["component"], kind=item["kind"], name=item["name"],
                    signature=str(item["signature"]).replace("|", "\\|"),
                )
            )
    if report["findings"]:
        lines.extend(["", "### Findings", ""])
        for item in report["findings"]:
            location = str(item.get("path", "."))
            if item.get("line"):
                location += f":{item['line']}"
            lines.append(f"- `{location}` — {item['message']}")
    return "\n".join(lines) + "\n"


def write_text(path: Path | None, text: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def init_fixture(root: Path, facade: str, manifest_lines: list[str]) -> None:
    (root / ".github").mkdir()
    (root / "src" / "modules").mkdir(parents=True)
    (root / "src" / "core").mkdir(parents=True)
    (root / "tests" / "modules").mkdir(parents=True)
    (root / "docs").mkdir()
    config = {
        "required_paths": ["docs/PUBLIC_API.txt"],
        "profiles": {
            name: {"vba_contract": {"minimum_roles": {"public": 1}}}
            for name in ("application", "library", "ui-component")
        },
        "vba": {
            "components": {
                "src/core/Core.bas": "internal",
                "src/modules/Facade.bas": "public",
                "tests/modules/Tests.bas": "test",
            },
            "public_api_manifest": "docs/PUBLIC_API.txt",
        },
    }
    (root / CONFIG_PATH).write_text(json.dumps(config, indent=2) + "\n", encoding="utf-8")
    files = {
        "src/modules/Facade.bas": facade,
        "src/core/Core.bas": 'Attribute VB_Name = "Core"\nOption Explicit\nOption Private Module\nPublic Function InternalOnly() As Long\nEnd Function\n',
        "tests/modules/Tests.bas": 'Attribute VB_Name = "Tests"\nOption Explicit\nPublic Sub RunTests()\nEnd Sub\n',
    }
    for relative, text in files.items():
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(text.replace("\n", "\r\n").encode("cp1252"))
    (root / "docs/PUBLIC_API.txt").write_text("\n".join(manifest_lines) + "\n", encoding="utf-8")
    for command in (
        ("init", "-b", "main"),
        ("config", "user.name", "API Self-Test"),
        ("config", "user.email", "api@example.invalid"),
        ("add", "--all"),
    ):
        completed = git(root, *command)
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())


def fixture_report(facade: str, manifest_lines: list[str]) -> dict[str, object]:
    with tempfile.TemporaryDirectory(prefix="vba-api-") as temporary:
        root = Path(temporary)
        init_fixture(root, facade, manifest_lines)
        return run_check(root)


def run_self_test() -> int:
    facade = '''Attribute VB_Name = "Facade"
Option Explicit
Public Const C As Long = 7
Public Event Changed(ByVal value As Long)
Public Declare PtrSafe Function Tick Lib "kernel32" () As Long
Public Number As Long
Public Enum Mode
    ModeA = 1
End Enum
Public Type Pair
    Left As Long
    Right As Long
End Type
Public Function Echo( _
    ByVal value As Long) As Long
    Echo = value
End Function
Public Property Get Current() As Long
    Current = Number
End Property
'''
    manifest = [
        "# Supported VBA declarations: <component><TAB><kind><TAB><name>",
        "Facade\tConst\tC",
        "# SIG\tFacade\tConst\tC\tPublic Const C As Long = 7",
        "Facade\tEvent\tChanged",
        "# SIG\tFacade\tEvent\tChanged\tPublic Event Changed(ByVal value As Long)",
        "Facade\tDeclare Function\tTick",
        "# SIG\tFacade\tDeclare Function\tTick\tPublic Declare PtrSafe Function Tick Lib \"kernel32\" () As Long",
        "Facade\tVariable\tNumber",
        "# SIG\tFacade\tVariable\tNumber\tPublic Number As Long",
        "Facade\tEnum\tMode",
        "# SIG\tFacade\tEnum\tMode\tPublic Enum Mode | ModeA = 1 | End Enum",
        "Facade\tType\tPair",
        "# SIG\tFacade\tType\tPair\tPublic Type Pair | Left As Long | Right As Long | End Type",
        "Facade\tFunction\tEcho",
        "# SIG\tFacade\tFunction\tEcho\tPublic Function Echo( ByVal value As Long) As Long",
        "Facade\tProperty Get\tCurrent",
        "# SIG\tFacade\tProperty Get\tCurrent\tPublic Property Get Current() As Long",
    ]
    failures: list[str] = []
    positive = fixture_report(facade, manifest)
    if positive["status"] != "pass":
        failures.append(f"positive fixture failed: {positive['findings']}")

    implicit = fixture_report(facade.replace("Public Function Echo", "Function Echo"), manifest)
    if implicit["status"] != "fail" or not any("Implicit public procedure" in item["message"] for item in implicit["findings"]):
        failures.append("implicit public procedure was not rejected")

    continued_bad = fixture_report(facade, [line for line in manifest if not line.startswith("# SIG\tFacade\tFunction\tEcho")])
    if continued_bad["status"] != "fail" or not any("Normalized signature record is missing" in item["message"] for item in continued_bad["findings"]):
        failures.append("continued function signature drift was not detected")

    duplicate_facade = facade + "\nPublic Function Echo(ByVal value As String) As String\nEnd Function\n"
    duplicate = fixture_report(duplicate_facade, manifest)
    if duplicate["status"] != "fail" or not any("collides" in item["message"] or "more than once" in item["message"] for item in duplicate["findings"]):
        failures.append("public-name collision was not rejected")

    multi_variable = fixture_report(facade.replace("Public Number As Long", "Public Number As Long, Other As Long"), manifest)
    if multi_variable["status"] != "fail" or not any("one identifier per statement" in item["message"] for item in multi_variable["findings"]):
        failures.append("multi-variable public declaration was not rejected")

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: procedures, properties, constants, events, declares, variables, enums, "
        "types, continuations, implicit-public rejection, signature drift, collisions, and "
        "single-variable policy passed."
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
        except (OSError, UnicodeError, RuntimeError, subprocess.SubprocessError, json.JSONDecodeError) as error:
            print(f"SELF-TEST ERROR: {error}", file=sys.stderr)
            return 2
    try:
        report = run_check(options.root)
        write_text(options.output, json.dumps(report, indent=2, sort_keys=True) + "\n")
        write_text(options.summary, markdown_report(report))
        print(markdown_report(report).rstrip())
        return 0 if report["status"] == "pass" else 1
    except (OSError, UnicodeError, RuntimeError, subprocess.SubprocessError, json.JSONDecodeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
