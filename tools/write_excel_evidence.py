#!/usr/bin/env python3
"""Write the host.json of an Excel evidence bundle from the logs a certification run left.

tools/Certify-Excel.ps1 imports the configured VBA into a new workbook, compiles it,
runs every suite and example in Excel and writes their logs, session.log and
session.json into the bundle. This script reads them, with the host policy and the
source inventory of the candidate, and writes host.json as docs/EXCEL_EVIDENCE.md
describes the record; it validates the bundle with check_excel_evidence.py, checks
every other suite's RESULT line against the counts its own source declares, and
writes certification.md and suites.json. It never runs Excel and records only what
the logs and session.json show: a suite without a RESULT line is TIMEOUT, and the
counts are the ones the RESULT lines print. A suite or example the committed profile
configures but session.json does not record is a failing "not run" row. The exit code
is 0 only when host.json validates as PASS, every configured suite passed with its
expected counts and every configured example ran.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any

import check_excel_evidence as evidence
from release_provenance import committed, decode, object_keys

DESCRIPTION = "Write host.json for an Excel evidence bundle, validate it and summarize every suite."
SESSION_KEYS = ("operator candidate_sha started_at finished_at excel_version excel_build "
                "office_bitness os os_architecture macro_policy vba_project_access references "
                "import compile suites examples workbook_saved workbook_closed")
RESULT_LINE = re.compile(r"^RESULT=(PASS|FAIL\w*); (?:completeness=(\w+); cases=(\d+); "
                         r"assertions=(\d+); failures=(\d+); )?cleanup=(\w+)\s*$", re.MULTILINE)
CONSTANT = re.compile(r"^\s*Private Const EXPECTED_(CASES|ASSERTIONS)\s+As Long\s*=\s*(\d+)", re.IGNORECASE)
DIRECTIVE = re.compile(r"^\s*#(If|ElseIf|Else|End If)\b(.*)$", re.IGNORECASE)


class BundleError(ValueError):
    """The bundle cannot be turned into a record; nothing is written."""


def reference(bundle: Path, name: str) -> dict[str, str]:
    return {"path": name, "sha256": hashlib.sha256((bundle / name).read_bytes()).hexdigest()}


def read_log(bundle: Path, name: str) -> str | None:
    path = bundle / name
    if not path.is_file():
        return None
    return path.read_bytes().decode("utf-8-sig", errors="replace").replace("\r\n", "\n")


def stage(status: str, detail: str, log: dict[str, str] | None) -> dict[str, Any]:
    return {"status": status, "detail": detail, "log": None if status == "NOT_RUN" else log}


def environment_line(log: str) -> dict[str, str]:
    """The fields of the one ENVIRONMENT= line a suite prints, or nothing."""
    values = re.findall(r"^ENVIRONMENT=(.*)$", log, re.MULTILINE)
    fields: dict[str, str] = {}
    for item in (values[0].strip().split("; ") if len(values) == 1 else []):
        key, _, value = item.partition("=")
        fields[key] = value
    return fields


def result_of(log: str) -> dict[str, Any] | None:
    """The one RESULT line of a log, parsed; None without exactly one such line."""
    found = RESULT_LINE.findall(log)
    if len(found) != 1:
        return None
    verdict, completeness, cases, assertions, failures, cleanup = found[0]
    return {"verdict": verdict, "completeness": completeness or "INCOMPLETE",
            "cases": int(cases) if cases else 0, "assertions": int(assertions) if assertions else 0,
            "failures": int(failures) if failures else 0, "cleanup": cleanup}


def expected_counts(source: str) -> dict[str, int]:
    """EXPECTED_CASES and EXPECTED_ASSERTIONS of a suite's source, taking the non-Mac branch."""
    counts: dict[str, int] = {}
    skipping = False
    for line in source.replace("\r\n", "\n").split("\n"):
        directive = DIRECTIVE.match(line)
        if directive:
            keyword = directive.group(1).casefold()
            condition = directive.group(2).casefold()
            if keyword == "if":
                skipping = "mac" in condition and "not mac" not in condition
            elif keyword == "elseif":
                skipping = "mac" in condition and "not mac" not in condition
            elif keyword == "else":
                skipping = not skipping
            else:
                skipping = False
            continue
        if skipping:
            continue
        match = CONSTANT.match(line)
        if match:
            counts[match.group(1).casefold()] = int(match.group(2))
    return counts


def lost_lines(closed: str) -> int | None:
    match = re.fullmatch(r"CLOSED; lost=(\d+)", closed or "")
    return int(match.group(1)) if match else None


def harness_record(policy: dict[str, Any], log: str | None) -> tuple[str, dict[str, Any] | None, str]:
    """Return the regression status the harness log supports, the harness object and its cleanup."""
    if log is None:
        return "NOT_RUN", None, "NOT_RUN"
    result = result_of(log)
    if result is None:
        return "TIMEOUT", None, "NOT_RUN"
    complete = (result["verdict"] == "PASS" and result["completeness"] == "COMPLETE"
                and result["failures"] == 0 and result["cleanup"] == "PASS")
    errors = [{"case": case, "status": "PASS" if complete else "NOT_RUN",
               "detail": "Implied by the complete harness without assertion failures" if complete
               else "Not inferred: the harness was incomplete, failed or did not clean up"}
              for case in policy["expected_error_cases"]]
    harness = {"entry_point": policy["entry_point"], "cases": result["cases"],
               "assertions": result["assertions"], "failures": result["failures"],
               "completeness": result["completeness"], "expected_errors": errors}
    return ("PASS" if complete else "FAIL"), harness, result["cleanup"]


def suite_summary(root: Path, sha: str, bundle: Path, item: dict[str, Any], harness: str) -> dict[str, Any]:
    """Judge one suite of session.json by its log and the counts its source declares."""
    entry = object_keys(item, "module entry_point log opened status error closed", "session suite")
    module = str(entry["module"])
    summary: dict[str, Any] = {"module": module, "entry_point": entry["entry_point"], "log": entry["log"],
                               "harness": module == harness, "status": "FAIL", "detail": ""}
    if entry["opened"] != "OPEN":
        summary["detail"] = f"log not opened: {entry['opened']}"
        return summary
    log = read_log(bundle, str(entry["log"]))
    lost = lost_lines(str(entry["closed"]))
    if log is None:
        summary["detail"] = "log missing from the bundle"
        return summary
    summary["sha256"] = hashlib.sha256((bundle / str(entry["log"])).read_bytes()).hexdigest()
    if lost is None or lost > 0:
        summary["detail"] = f"lines lost while writing the log: {entry['closed']}"
        return summary
    if entry["status"] != "RAN":
        summary["detail"] = f"the entry point raised: {entry['error']}"
        return summary
    result = result_of(log)
    if result is None:
        summary["status"] = "TIMEOUT"
        summary["detail"] = "no single RESULT line"
        return summary
    summary.update({key: result[key] for key in ("verdict", "completeness", "cases", "assertions",
                                                   "failures", "cleanup")})
    try:
        source = committed(root, sha, f"tests/modules/{module}.bas").decode("cp1252")
    except (OSError, ValueError, UnicodeDecodeError):
        summary["detail"] = "source of the suite is not in the candidate"
        return summary
    expected = expected_counts(source)
    summary["expected_cases"] = expected.get("cases")
    summary["expected_assertions"] = expected.get("assertions")
    problems = []
    if result["verdict"] != "PASS" or result["completeness"] != "COMPLETE" or result["failures"] != 0:
        problems.append(f"RESULT={result['verdict']}; completeness={result['completeness']}; "
                        f"failures={result['failures']}")
    if result["cleanup"] != "PASS":
        problems.append(f"cleanup={result['cleanup']}")
    for key in ("cases", "assertions"):
        if expected.get(key) is not None and expected[key] != result[key]:
            problems.append(f"{key}={result[key]}, source declares {expected[key]}")
    summary["status"] = "FAIL" if problems else "PASS"
    summary["detail"] = "; ".join(problems) if problems else (
        f"cases={result['cases']}; assertions={result['assertions']}; cleanup=PASS")
    return summary


def configured_modules(config: dict[str, Any], role: str, exclude: frozenset[str] | set[str] = frozenset()) -> list[str]:
    """The modules of one role in the committed profile, named as Certify-Excel.ps1 names them."""
    components = (config.get("vba") or {}).get("components") or {}
    return [Path(path).stem for path, value in sorted(components.items())
            if value == role and Path(path).stem not in exclude]


def build_record(root: Path, sha: str, bundle: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    """The evidence record the bundle supports, and the summary of every suite and example."""
    config = decode(committed(root, sha, ".github/repository-profile.json"))
    policy = evidence.load_policy(root, sha)
    harness_module = policy["entry_point"].split(".")[0]
    session = object_keys(json.loads((bundle / "session.json").read_text(encoding="utf-8-sig")),
                          SESSION_KEYS, "session.json")
    if session["candidate_sha"] != sha:
        raise BundleError(f"session.json names candidate {session['candidate_sha']}, not {sha}")
    session_log = reference(bundle, "session.log")
    imported = object_keys(session["import"], "status detail components", "session import")
    compiled = object_keys(session["compile"], "status detail", "session compile")
    ran = imported["status"] == "PASS" and compiled["status"] == "PASS"
    harness_log = read_log(bundle, "harness.log") if ran else None
    harness_status, harness, cleanup = harness_record(policy, harness_log)
    environment = environment_line(harness_log) if harness_log else {}
    if harness_status in ("PASS", "FAIL"):
        regression = stage(harness_status, "Certify-Excel: the harness run on its own into harness.log",
                           reference(bundle, "harness.log"))
    elif harness_status == "TIMEOUT":
        regression = stage("TIMEOUT", "The harness log has no single RESULT line", reference(bundle, "harness.log"))
    else:
        regression = stage("NOT_RUN", "Not run: import or compile did not pass", None)
    closed = session["workbook_closed"] is True
    cleanup_status = "PASS" if closed and cleanup in ("PASS", "NOT_RUN") else "FAIL"
    cleanup_detail = (f"Harness cleanup {cleanup}; " if cleanup != "NOT_RUN" else "") + \
        ("workbook closed without saving the run" if closed else "workbook not closed")
    record = {
        "schema_version": 1, "repository": config["repository"], "candidate_sha": sha,
        "template_contract": config["template_contract"], "execution": "manual",
        "availability_reason": None,
        "started_at": session["started_at"], "finished_at": session["finished_at"],
        "runner": {"class": "manual-interactive", "identity": session["operator"], "workflow": None},
        "environment": {
            "excel_version": environment.get("version", session["excel_version"]),
            "excel_build": session["excel_build"],
            "office_bitness": environment.get("office", session["office_bitness"]),
            "os": session["os"], "os_architecture": session["os_architecture"],
            "runtime": environment.get("runtime", "unknown"),
            "macro_policy": session["macro_policy"],
            "vba_project_access": session["vba_project_access"], "trust_changes": False},
        "sources": evidence.source_inventory(root, sha, config),
        "stages": {
            "import": stage(imported["status"], imported["detail"], session_log),
            "compile": stage(compiled["status"] if imported["status"] == "PASS" else "NOT_RUN",
                             compiled["detail"], session_log),
            "regression": regression,
            "cleanup": stage(cleanup_status, cleanup_detail, session_log)},
        "harness": harness if harness_status != "NOT_RUN" else None,
    }
    suites = [suite_summary(root, sha, bundle, item, harness_module) for item in session["suites"]]
    examples = []
    for item in session["examples"]:
        entry = object_keys(item, "module entry_point status error", "session example")
        examples.append({"module": entry["module"], "entry_point": entry["entry_point"],
                         "status": "PASS" if entry["status"] == "RAN" else "FAIL",
                         "detail": "ran without error" if entry["status"] == "RAN"
                         else f"{entry['status']}: {entry['error']}"})
    suites += [{"module": module, "entry_point": f"{module}.Run{module}", "log": None,
                "harness": module == harness_module, "status": "FAIL", "detail": "not run: missing from session.json"}
               for module in configured_modules(config, "test", exclude={"TestLog"})
               if module not in {suite["module"] for suite in suites}]
    examples += [{"module": module, "entry_point": f"{module}.Run{module}", "status": "FAIL",
                  "detail": "not run: missing from session.json"}
                 for module in configured_modules(config, "example")
                 if module not in {example["module"] for example in examples}]
    summary = {"candidate_sha": sha, "harness_module": harness_module, "suites": suites, "examples": examples}
    return record, summary


def certification_markdown(report: dict[str, Any], summary: dict[str, Any], record: dict[str, Any]) -> str:
    lines = [f"# Excel certification of {summary['candidate_sha']}", "",
             f"Host record: {report['status'].upper()} ({', '.join(report['outcomes'])})", ""]
    lines += [f"- {item}" for item in report["findings"]]
    environment = record["environment"]
    lines += ["", f"Environment: Excel {environment['excel_version']}, {environment['excel_build']}, "
              f"{environment['office_bitness']}; {environment['os']}; runtime {environment['runtime']}", "",
              "| Suite | Status | Detail |", "| --- | --- | --- |"]
    for suite in summary["suites"]:
        label = suite["entry_point"] + (" (harness)" if suite["harness"] else "")
        lines.append(f"| `{label}` | {suite['status']} | {suite['detail']} |")
    for example in summary["examples"]:
        lines.append(f"| `{example['entry_point']}` (example) | {example['status']} | {example['detail']} |")
    verdict = "PASS" if certified(report, summary) else "FAIL"
    lines += ["", f"RESULT={verdict}", ""]
    return "\n".join(lines)


def certified(report: dict[str, Any], summary: dict[str, Any]) -> bool:
    return (report["status"] == "pass"
            and all(suite["status"] == "PASS" for suite in summary["suites"])
            and all(example["status"] == "PASS" for example in summary["examples"]))


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=DESCRIPTION)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="the certified checkout")
    parser.add_argument("--candidate-sha", required=True, help="the full SHA of the checked-out commit")
    parser.add_argument("--bundle", type=Path, required=True,
                        help="the bundle folder holding session.json, session.log and the suite logs")
    options = parser.parse_args(arguments)
    target = options.bundle / "host.json"
    if target.exists():
        print(f"ERROR: {target} exists already; a record is never replaced", file=sys.stderr)
        return 2
    try:
        record, summary = build_record(options.root, options.candidate_sha, options.bundle)
    except (ValueError, KeyError, TypeError, OSError, json.JSONDecodeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    target.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    report = evidence.evaluate(options.root, options.candidate_sha, target)
    (options.bundle / "suites.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    text = certification_markdown(report, summary, record)
    (options.bundle / "certification.md").write_text(text, encoding="utf-8")
    print(text)
    return 0 if certified(report, summary) else 1


if __name__ == "__main__":
    raise SystemExit(main())
