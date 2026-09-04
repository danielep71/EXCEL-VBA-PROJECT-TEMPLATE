#!/usr/bin/env python3
'''Validate explicit VBA public API declarations and the checked-in manifest.'''

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from typing import Any

CONFIG_PATH = ".github/repository-profile.json"
TOOL_NAME = "VBA public API"
VBA_SUFFIXES = {".bas", ".cls", ".frm"}
SIG_PREFIX = "# SIG\t"
PROPERTY_KINDS = {"property get", "property let", "property set"}

PROC = re.compile(
    r"^\s*(?:(?P<vis>Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z_]\w*)\b", re.I
)
END_PROC = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
BLOCK = re.compile(r"^\s*Public\s+(?P<kind>Enum|Type)\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
END_BLOCK = {
    "enum": re.compile(r"^\s*End\s+Enum\b", re.I),
    "type": re.compile(r"^\s*End\s+Type\b", re.I),
}
DECLARE = re.compile(
    r"^\s*(?:Public|Global)\s+Declare\s+(?:PtrSafe\s+)?"
    r"(?P<kind>Function|Sub)\s+(?P<name>[A-Za-z_]\w*)\b", re.I
)
EVENT = re.compile(r"^\s*(?:Public|Global)\s+Event\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
CONST = re.compile(
    r"^\s*(?:Public|Global)\s+Const\s+(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$", re.I
)
VARIABLE = re.compile(
    r"^\s*(?:Public|Global)\s+(?P<withevents>WithEvents\s+)?"
    r"(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$", re.I
)
IMPLICIT = re.compile(
    r"^\s*(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|"
    r"Enum|Type|Event|Declare\s+(?:PtrSafe\s+)?(?:Function|Sub))\b", re.I
)


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(
        ["git", "-C", str(root), *args],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE
    )


def tracked_vba(root: Path) -> list[str]:
    cp = git(root, "ls-files", "-z")
    if cp.returncode:
        raise RuntimeError(cp.stderr.decode("utf-8", errors="replace").strip())
    return sorted(
        p.decode("utf-8", errors="surrogateescape")
        for p in cp.stdout.split(b"\0")
        if p and Path(p.decode("utf-8", errors="surrogateescape")).suffix.casefold() in VBA_SUFFIXES
    )


def strip_vba(raw: str) -> str:
    out: list[str] = []
    string = False
    i = 0
    while i < len(raw):
        ch = raw[i]
        if ch == '"':
            if string and i + 1 < len(raw) and raw[i + 1] == '"':
                out.extend(('"', '"')); i += 2; continue
            string = not string
        elif ch == "'" and not string:
            break
        out.append(ch); i += 1
    text = "".join(out)
    return "" if re.match(r"^\s*Rem(?:\s|$)", text, re.I) else text.rstrip()


def logical(lines: list[str]) -> list[tuple[int, int, str]]:
    result: list[tuple[int, int, str]] = []
    buf: list[str] = []
    start = 0
    for n, raw in enumerate(lines, 1):
        code = strip_vba(raw)
        if not buf: start = n
        if re.search(r"\s_\s*$", code):
            buf.append(re.sub(r"\s_\s*$", " ", code)); continue
        buf.append(code)
        result.append((start, n, " ".join(x.strip() for x in buf)))
        buf.clear()
    if buf:
        result.append((start, len(lines), " ".join(x.strip() for x in buf)))
    return result


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def key(component: str, kind: str, name: str) -> str:
    return f"{component}\t{kind}\t{name}"


def top_comma(text: str) -> bool:
    depth = 0
    string = False
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            if string and i + 1 < len(text) and text[i + 1] == '"':
                i += 2; continue
            string = not string
        elif not string:
            if ch == "(": depth += 1
            elif ch == ")" and depth: depth -= 1
            elif ch == "," and depth == 0: return True
        i += 1
    return False


def record(out: list[dict[str, Any]], supported: bool, component: str,
           kind: str, name: str, signature: str, path: str, line: int) -> None:
    if supported:
        out.append({
            "component": component, "kind": kind, "name": name,
            "signature": signature, "path": path, "line": line
        })


def parse_component(path: str, text: str, supported: bool):
    component = Path(path).stem
    decls: list[dict[str, Any]] = []
    findings: list[dict[str, Any]] = []
    statements = logical(text.splitlines())
    inside = False
    i = 0
    while i < len(statements):
        line, _, code = statements[i]
        s = code.strip()
        if not s or s.startswith(("Attribute ", "Option ", "#")):
            i += 1; continue

        m = PROC.match(code)
        if m and " declare " not in f" {code.casefold()} ":
            vis = (m.group("vis") or "").casefold()
            if not vis:
                findings.append({"path": path, "line": line,
                                 "message": f"Implicit public procedure is prohibited: {m.group('name')}."})
            if vis == "public":
                kind = " ".join(x.capitalize() for x in m.group("kind").split())
                record(decls, supported, component, kind, m.group("name"), norm(code), path, line)
            inside = True; i += 1; continue
        if END_PROC.match(code):
            inside = False; i += 1; continue
        if inside:
            i += 1; continue

        if IMPLICIT.match(code):
            findings.append({"path": path, "line": line,
                             "message": f"Implicit public module-level declaration is prohibited: {norm(code)}"})
            i += 1; continue

        m = BLOCK.match(code)
        if m:
            kind, name = m.group("kind").capitalize(), m.group("name")
            body = [norm(code)]
            close = END_BLOCK[kind.casefold()]
            j = i + 1
            while j < len(statements) and not close.match(statements[j][2]):
                if statements[j][2].strip(): body.append(norm(statements[j][2]))
                j += 1
            if j >= len(statements):
                findings.append({"path": path, "line": line,
                                 "message": f"Public {kind} {name} is not closed."})
                i += 1; continue
            body.append(norm(statements[j][2]))
            record(decls, supported, component, kind, name, " | ".join(body), path, line)
            i = j + 1; continue

        m = DECLARE.match(code)
        if m:
            kind = "Declare Function" if m.group("kind").casefold() == "function" else "Declare Sub"
            record(decls, supported, component, kind, m.group("name"), norm(code), path, line)
            i += 1; continue

        m = EVENT.match(code)
        if m:
            record(decls, supported, component, "Event", m.group("name"), norm(code), path, line)
            i += 1; continue

        m = CONST.match(code)
        if m:
            if top_comma(m.group("rest")):
                findings.append({"path": path, "line": line,
                                 "message": "Public Const declarations must contain one identifier per statement."})
            else:
                record(decls, supported, component, "Const", m.group("name"), norm(code), path, line)
            i += 1; continue

        m = VARIABLE.match(code)
        if m:
            if top_comma(m.group("rest")):
                findings.append({"path": path, "line": line,
                                 "message": "Public variable declarations must contain one identifier per statement."})
            else:
                kind = "WithEvents Variable" if m.group("withevents") else "Variable"
                record(decls, supported, component, kind, m.group("name"), norm(code), path, line)
            i += 1; continue

        i += 1
    return decls, findings


def read_manifest(root: Path, relative: str):
    p = root / relative
    if not p.is_file():
        return set(), {}, [{"path": relative, "message": "Configured public API manifest is missing."}]
    rows: set[str] = set()
    sigs: dict[str, str] = {}
    findings: list[dict[str, Any]] = []
    for n, raw in enumerate(p.read_text(encoding="utf-8").splitlines(), 1):
        if not raw.strip(): continue
        if raw.startswith(SIG_PREFIX):
            fields = raw[len(SIG_PREFIX):].split("\t", 3)
            if len(fields) != 4:
                findings.append({"path": relative, "line": n, "message": "Malformed # SIG record."}); continue
            k = key(fields[0], fields[1], fields[2])
            if any(x.casefold() == k.casefold() for x in sigs):
                findings.append({"path": relative, "line": n, "message": f"Duplicate signature record: {k}"})
            sigs[k] = fields[3]; continue
        if raw.lstrip().startswith("#"): continue
        fields = raw.split("\t")
        if len(fields) != 3 or any(not x for x in fields):
            findings.append({"path": relative, "line": n,
                             "message": "Manifest declaration rows require component, kind, and name."}); continue
        k = key(*fields)
        if any(x.casefold() == k.casefold() for x in rows):
            findings.append({"path": relative, "line": n, "message": f"Duplicate manifest declaration: {k}"})
        rows.add(k)
    return rows, sigs, findings


def is_property(kind: str) -> bool:
    return kind.casefold() in PROPERTY_KINDS


def property_pair(a: dict[str, Any], b: dict[str, Any]) -> bool:
    return (
        str(a["component"]).casefold() == str(b["component"]).casefold()
        and is_property(str(a["kind"])) and is_property(str(b["kind"]))
        and str(a["kind"]).casefold() != str(b["kind"]).casefold()
    )


def run_check(root: Path) -> dict[str, Any]:
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    vba = config["vba"]
    components: dict[str, str] = vba["components"]
    manifest = vba["public_api_manifest"]
    findings: list[dict[str, Any]] = []

    if manifest not in config["required_paths"]:
        findings.append({"path": CONFIG_PATH,
                         "message": "public_api_manifest must be a global required_path for every profile."})
    for profile, entry in config["profiles"].items():
        if entry["vba_contract"]["minimum_roles"].get("public", 0) < 1:
            findings.append({"path": CONFIG_PATH,
                             "message": f"Profile {profile!r} must require at least one public-role component."})

    decls: list[dict[str, Any]] = []
    tracked = set(tracked_vba(root))
    for path, role in sorted(components.items()):
        if path not in tracked or not (root / path).is_file(): continue
        parsed, errs = parse_component(path, (root / path).read_bytes().decode("cp1252"), role == "public")
        decls.extend(parsed); findings.extend(errs)

    keys: dict[str, dict[str, Any]] = {}
    names: dict[str, list[dict[str, Any]]] = {}
    for d in decls:
        k = key(str(d["component"]), str(d["kind"]), str(d["name"]))
        if any(existing.casefold() == k.casefold() for existing in keys):
            findings.append({"path": d["path"], "line": d["line"],
                             "message": f"Public declaration appears more than once: {k}"})
        else:
            keys[k] = d
        if Path(str(d["path"])).suffix.casefold() == ".bas":
            nk = str(d["name"]).casefold()
            prior = names.setdefault(nk, [])
            for p in prior:
                if not property_pair(p, d):
                    findings.append({"path": d["path"], "line": d["line"],
                                     "message": f"Public standard-module name {d['name']!r} collides with {p['component']}.{p['name']} ({p['kind']})."})
            prior.append(d)

    rows, sigs, errs = read_manifest(root, manifest)
    findings.extend(errs)
    actual = {k.casefold(): k for k in keys}
    rowmap = {k.casefold(): k for k in rows}
    sigmap = {k.casefold(): k for k in sigs}
    for folded, k in actual.items():
        if folded not in rowmap:
            findings.append({"path": manifest, "message": f"Public declaration is not recorded: {k}"})
        sk = sigmap.get(folded)
        if sk is None:
            findings.append({"path": manifest, "message": f"Normalized signature record is missing: {k}"})
        elif sigs[sk] != str(keys[k]["signature"]):
            findings.append({"path": manifest, "message": f"Signature mismatch for {k}: expected {keys[k]['signature']!r}, recorded {sigs[sk]!r}"})
    for folded, k in rowmap.items():
        if folded not in actual:
            findings.append({"path": manifest, "message": f"Manifest declaration is not present in supported source: {k}"})
    for folded, k in sigmap.items():
        if folded not in actual:
            findings.append({"path": manifest, "message": f"Stale signature record: {k}"})

    evidence = sorted(decls, key=lambda d: (
        str(d["component"]).casefold(), str(d["name"]).casefold(), str(d["kind"]).casefold()
    ))
    return {
        "schema_version": 1, "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "manifest": manifest,
        "policy": {
            "required_profiles": sorted(config["profiles"]),
            "required_from": "initialization",
            "implicit_visibility": "prohibited",
            "const_and_variable_form": "one identifier per statement",
            "property_accessors": "same-name Get/Let/Set allowed only within one component",
        },
        "declarations": evidence, "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "## VBA public API", "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Manifest:** `{report['manifest']}`",
        f"- **Declarations:** {len(report['declarations'])}",
        "- **Profiles:** " + ", ".join(f"`{x}`" for x in report["policy"]["required_profiles"]),
        "- **Manifest required from:** initialization",
        "- **Implicit public visibility:** prohibited",
        "- **Public Const/variable form:** one identifier per statement",
        "- **Property accessors:** same-name Get/Let/Set allowed within one component",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["declarations"]:
        lines += ["", "| Component | Kind | Name | Signature |", "| --- | --- | --- | --- |"]
        for d in report["declarations"]:
            signature = str(d["signature"]).replace("|", "\\|")
            lines.append(f"| {d['component']} | {d['kind']} | {d['name']} | `{signature}` |")
    if report["findings"]:
        lines += ["", "### Findings", ""]
        for item in report["findings"]:
            loc = str(item.get("path", "."))
            if item.get("line"): loc += f":{item['line']}"
            lines.append(f"- `{loc}` — {item['message']}")
    return "\n".join(lines) + "\n"


def write_text(path: Path | None, text: str) -> None:
    if path:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8", newline="\n")


def init_fixture(root: Path, facade: str, manifest: list[str], other: str | None = None) -> None:
    for p in (root/".github", root/"src/modules", root/"src/core", root/"tests/modules", root/"docs"):
        p.mkdir(parents=True, exist_ok=True)
    components = {
        "src/core/Core.bas": "internal",
        "src/modules/Facade.bas": "public",
        "tests/modules/Tests.bas": "test",
    }
    if other is not None: components["src/modules/Other.bas"] = "public"
    cfg = {
        "required_paths": ["docs/PUBLIC_API.txt"],
        "profiles": {x: {"vba_contract": {"minimum_roles": {"public": 1}}}
                     for x in ("application", "library", "ui-component")},
        "vba": {"components": components, "public_api_manifest": "docs/PUBLIC_API.txt"},
    }
    (root/CONFIG_PATH).write_text(json.dumps(cfg, indent=2)+"\n", encoding="utf-8")
    files = {
        "src/modules/Facade.bas": facade,
        "src/core/Core.bas": 'Attribute VB_Name = "Core"\nOption Explicit\nOption Private Module\nPublic Function InternalOnly() As Long\nEnd Function\n',
        "tests/modules/Tests.bas": 'Attribute VB_Name = "Tests"\nOption Explicit\nPublic Sub RunTests()\nEnd Sub\n',
    }
    if other is not None: files["src/modules/Other.bas"] = other
    for rel, text in files.items():
        (root/rel).write_bytes(text.replace("\n", "\r\n").encode("cp1252"))
    (root/"docs/PUBLIC_API.txt").write_text("\n".join(manifest)+"\n", encoding="utf-8")
    for cmd in (("init","-b","main"),("config","user.name","API Self-Test"),
                ("config","user.email","api@example.invalid"),("add","--all")):
        cp = git(root, *cmd)
        if cp.returncode: raise RuntimeError(cp.stderr.decode("utf-8", errors="replace").strip())


def fixture(facade: str, manifest: list[str], other: str | None = None) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="vba-api-") as tmp:
        root = Path(tmp); init_fixture(root, facade, manifest, other); return run_check(root)


def run_self_test() -> int:
    facade = '''Attribute VB_Name = "Facade"
Option Explicit
Public Const C As Long = 7
Public Event Changed(ByVal value As Long)
Public Declare PtrSafe Function Tick Lib "kernel32" () As Long
Public Number As Long
Public WithEvents Source As Object
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
Public Property Let Current(ByVal value As Long)
    Number = value
End Property
'''
    manifest = [
        "# Supported VBA declarations: <component><TAB><kind><TAB><name>",
        "Facade\tConst\tC", "# SIG\tFacade\tConst\tC\tPublic Const C As Long = 7",
        "Facade\tEvent\tChanged", "# SIG\tFacade\tEvent\tChanged\tPublic Event Changed(ByVal value As Long)",
        "Facade\tDeclare Function\tTick", '# SIG\tFacade\tDeclare Function\tTick\tPublic Declare PtrSafe Function Tick Lib "kernel32" () As Long',
        "Facade\tVariable\tNumber", "# SIG\tFacade\tVariable\tNumber\tPublic Number As Long",
        "Facade\tWithEvents Variable\tSource", "# SIG\tFacade\tWithEvents Variable\tSource\tPublic WithEvents Source As Object",
        "Facade\tEnum\tMode", "# SIG\tFacade\tEnum\tMode\tPublic Enum Mode | ModeA = 1 | End Enum",
        "Facade\tType\tPair", "# SIG\tFacade\tType\tPair\tPublic Type Pair | Left As Long | Right As Long | End Type",
        "Facade\tFunction\tEcho", "# SIG\tFacade\tFunction\tEcho\tPublic Function Echo( ByVal value As Long) As Long",
        "Facade\tProperty Get\tCurrent", "# SIG\tFacade\tProperty Get\tCurrent\tPublic Property Get Current() As Long",
        "Facade\tProperty Let\tCurrent", "# SIG\tFacade\tProperty Let\tCurrent\tPublic Property Let Current(ByVal value As Long)",
    ]
    failures: list[str] = []
    tests: list[tuple[str, dict[str, Any], str, str | None]] = []
    tests.append(("positive", fixture(facade, manifest), "pass", None))
    tests.append(("implicit", fixture(facade.replace("Public Function Echo","Function Echo"), manifest),
                  "fail", "Implicit public procedure"))
    tests.append(("missing-signature", fixture(facade, [x for x in manifest if not x.startswith("# SIG\tFacade\tFunction\tEcho")]),
                  "fail", "Normalized signature record is missing"))
    tests.append(("multi-var", fixture(facade.replace("Public Number As Long","Public Number As Long, Other As Long"), manifest),
                  "fail", "Public variable declarations must contain one identifier"))
    tests.append(("multi-const", fixture(facade.replace("Public Const C As Long = 7","Public Const C As Long = 7, D As Long = 8"), manifest),
                  "fail", "Public Const declarations must contain one identifier"))
    dup = facade + "\nPublic Function Echo(ByVal value As String) As String\nEnd Function\n"
    tests.append(("same-component-collision", fixture(dup, manifest), "fail", "collides"))
    other = '''Attribute VB_Name = "Other"
Option Explicit
Public Function Echo(ByVal value As Long) As Long
    Echo = value
End Function
'''
    tests.append(("cross-component-collision", fixture(facade, manifest, other), "fail", "collides"))
    for name, report, expected, needle in tests:
        if report["status"] != expected:
            failures.append(f"{name}: expected {expected}, got {report['status']}")
        if needle and not any(needle in item["message"] for item in report["findings"]):
            failures.append(f"{name}: missing diagnostic {needle!r}")
    if failures:
        for x in failures: print(f"[FAIL] {x}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s)."); return 1
    print("SELF-TEST PASS: procedures, paired properties, constants, events, declares, variables, WithEvents, enums, types, continuations, implicit-public rejection, signature drift, collisions, and one-identifier Const/variable policy passed.")
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--root", type=Path, default=Path.cwd())
    p.add_argument("--output", type=Path)
    p.add_argument("--summary", type=Path)
    p.add_argument("--self-test", action="store_true")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    o = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        if o.self_test: return run_self_test()
        report = run_check(o.root)
        write_text(o.output, json.dumps(report, indent=2, sort_keys=True)+"\n")
        write_text(o.summary, markdown_report(report))
        print(markdown_report(report).rstrip())
        return 0 if report["status"] == "pass" else 1
    except (OSError, UnicodeError, RuntimeError, subprocess.SubprocessError, json.JSONDecodeError) as e:
        print(f"ERROR: {e}", file=sys.stderr); return 2


if __name__ == "__main__":
    raise SystemExit(main())
