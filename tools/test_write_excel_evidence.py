"""Bundle fixtures for write_excel_evidence.py; none executes Office."""

from __future__ import annotations

import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from typing import Any

import check_excel_evidence as host
import check_release as release
import write_excel_evidence as writer

ROOT = Path(__file__).resolve().parents[1]
HARNESS_SOURCE = (
    "Attribute VB_Name = \"ProjectTests\"\r\nOption Explicit\r\nOption Private Module\r\n"
    "        Private Const EXPECTED_ASSERTIONS      As Long = 6                       'Required\r\n"
    "        Private Const EXPECTED_CASES           As Long = 4                       'Required\r\n"
)
SUITE_SOURCE = (
    "Attribute VB_Name = \"TextTests\"\r\nOption Explicit\r\nOption Private Module\r\n"
    "        Private Const EXPECTED_CASES           As Long = 29\r\n"
    "#If Mac Then\r\n"
    "        Private Const EXPECTED_ASSERTIONS      As Long = 131\r\n"
    "#Else\r\n"
    "        Private Const EXPECTED_ASSERTIONS      As Long = 135\r\n"
    "#End If\r\n"
)


class WriteEvidenceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.area = Path(self.temporary.name)
        self.root = self.area / "candidate"
        self.bundle = self.area / "bundle"
        self.bundle.mkdir()
        release_policy = json.loads((ROOT / release.POLICY_PATH).read_text())
        release._fixture_repository(self.root, "library", release_policy)
        config = release._fixture_configuration("library")
        config["template_contract"] = {"version": "1.2.0", "source": "example/template"}
        config["vba"] = {"components": {
            "src/Project.bas": "public", "tests/modules/ProjectTests.bas": "test",
            "tests/modules/TextTests.bas": "test", "examples/modules/ProjectExample.bas": "example"}}
        (self.root / release.PROFILE_PATH).write_text(json.dumps(config))
        (self.root / "tests/modules").mkdir(parents=True)
        (self.root / "examples/modules").mkdir(parents=True)
        (self.root / "tests/modules/ProjectTests.bas").write_bytes(HARNESS_SOURCE.encode("cp1252"))
        (self.root / "tests/modules/TextTests.bas").write_bytes(SUITE_SOURCE.encode("cp1252"))
        (self.root / "examples/modules/ProjectExample.bas").write_bytes(b"Attribute VB_Name = \"ProjectExample\"\r\n")
        self.policy = json.loads((ROOT / host.POLICY).read_text())
        (self.root / host.POLICY).write_text(json.dumps(self.policy))
        release._git(self.root, "add", "--all")
        release._git(self.root, "commit", "-m", "Synthetic certification candidate")
        self.sha = release._git_output(self.root, "rev-parse", "HEAD") or ""
        self.write_bundle()

    def write_bundle(self, harness_result: str | None = None, suite_result: str | None = None) -> None:
        harness_result = harness_result or (
            "RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS")
        suite_result = suite_result or (
            "RESULT=PASS; completeness=COMPLETE; cases=29; assertions=135; failures=0; cleanup=PASS")
        harness = "\r\n".join([
            f"CANDIDATE={self.sha}", "PROJECT TESTS",
            "ENVIRONMENT=host=Microsoft Excel; version=16.0; os=Windows (64-bit) NT 10.00; office=64-bit; runtime=VBA7+",
            *("CASE=" + case for case in self.policy["cases"]),
            "CASES=4", "ASSERTIONS=6", "FAILURES=0", "CLEANUP=PASS; detail=state unchanged", harness_result, ""])
        (self.bundle / "harness.log").write_bytes(harness.encode())
        suite = "\r\n".join([f"CANDIDATE={self.sha}", "TEXT TESTS", "CASES=29", "ASSERTIONS=135",
                             "FAILURES=0", suite_result, ""])
        (self.bundle / "TextTests.log").write_bytes(suite.encode())
        (self.bundle / "session.log").write_text("09:00:00  Excel certification\n")
        self.session: dict[str, Any] = {
            "operator": "Synthetic operator", "candidate_sha": self.sha,
            "started_at": "2026-11-07T09:00:00+01:00", "finished_at": "2026-11-07T09:05:00+01:00",
            "excel_version": "16.0", "excel_build": "Version 2510 (build 19328.20000)",
            "office_bitness": "64-bit", "os": "Windows 11 Pro (version 10.0.26100, build 26100)",
            "os_architecture": "x64", "macro_policy": "Disable VBA macros with notification",
            "vba_project_access": "On", "references": ["Visual Basic For Applications"],
            "import": {"status": "PASS", "detail": "VBComponents.Import", "components": ["src/Project.bas"]},
            "compile": {"status": "PASS", "detail": "Debug > Compile VBAProject completed with no error"},
            "suites": [
                {"module": "ProjectTests", "entry_point": "ProjectTests.RunProjectTests", "log": "harness.log",
                 "opened": "OPEN", "status": "RAN", "error": None, "closed": "CLOSED; lost=0"},
                {"module": "TextTests", "entry_point": "TextTests.RunTextTests", "log": "TextTests.log",
                 "opened": "OPEN", "status": "RAN", "error": None, "closed": "CLOSED; lost=0"}],
            "examples": [{"module": "ProjectExample", "entry_point": "ProjectExample.RunProjectExample",
                          "status": "RAN", "error": None}],
            "workbook_saved": True, "workbook_closed": True}
        self.save_session()

    def save_session(self) -> None:
        (self.bundle / "session.json").write_text(json.dumps(self.session))

    def run_writer(self) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = writer.main(["--root", str(self.root), "--candidate-sha", self.sha,
                                "--bundle", str(self.bundle)])
        return code, out.getvalue(), err.getvalue()

    def summary(self) -> dict[str, Any]:
        loaded: dict[str, Any] = json.loads((self.bundle / "suites.json").read_text())
        return loaded

    def test_complete_run_passes_and_validates(self) -> None:
        code, out, _ = self.run_writer()
        self.assertEqual(code, 0, out)
        record = json.loads((self.bundle / "host.json").read_text())
        self.assertEqual(host.evaluate(self.root, self.sha, self.bundle / "host.json")["outcomes"], ["PASS"])
        self.assertEqual(record["environment"]["runtime"], "VBA7+")
        self.assertEqual(record["harness"]["expected_errors"][0]["status"], "PASS")
        self.assertIn("RESULT=PASS", (self.bundle / "certification.md").read_text())
        statuses = {suite["module"]: suite["status"] for suite in self.summary()["suites"]}
        self.assertEqual(statuses, {"ProjectTests": "PASS", "TextTests": "PASS"})
        self.assertEqual(self.summary()["suites"][1]["expected_assertions"], 135)

    def test_suite_counts_must_match_its_source(self) -> None:
        self.write_bundle(suite_result=(
            "RESULT=PASS; completeness=COMPLETE; cases=28; assertions=135; failures=0; cleanup=PASS"))
        code, _, _ = self.run_writer()
        self.assertEqual(code, 1)
        suite = self.summary()["suites"][1]
        self.assertEqual(suite["status"], "FAIL")
        self.assertIn("source declares 29", suite["detail"])

    def test_mac_branch_is_not_the_windows_expectation(self) -> None:
        self.assertEqual(writer.expected_counts(SUITE_SOURCE), {"cases": 29, "assertions": 135})
        self.assertEqual(writer.expected_counts(HARNESS_SOURCE), {"cases": 4, "assertions": 6})

    def test_missing_result_line_is_timeout(self) -> None:
        (self.bundle / "harness.log").write_bytes(f"CANDIDATE={self.sha}\r\nPROJECT TESTS\r\n".encode())
        code, _, _ = self.run_writer()
        self.assertEqual(code, 1)
        record = json.loads((self.bundle / "host.json").read_text())
        self.assertEqual(record["stages"]["regression"]["status"], "TIMEOUT")
        self.assertIsNone(record["harness"])
        report = host.evaluate(self.root, self.sha, self.bundle / "host.json")
        self.assertIn("EXECUTION_TIMEOUT", report["outcomes"])

    def test_failed_import_leaves_later_stages_not_run(self) -> None:
        self.session["import"] = {"status": "FAIL", "detail": "Import raised", "components": []}
        self.session["compile"] = {"status": "NOT_RUN", "detail": "Not run"}
        self.session["suites"] = []
        self.session["examples"] = []
        self.save_session()
        code, _, _ = self.run_writer()
        self.assertEqual(code, 1)
        record = json.loads((self.bundle / "host.json").read_text())
        self.assertEqual([record["stages"][name]["status"] for name in host.STAGES],
                         ["FAIL", "NOT_RUN", "NOT_RUN", "PASS"])
        self.assertIsNone(record["harness"])
        self.assertEqual(host.evaluate(self.root, self.sha, self.bundle / "host.json")["outcomes"],
                         ["IMPORT_FAILED", "INCOMPLETE"])

    def test_lost_lines_and_raised_entry_point_fail_the_suite(self) -> None:
        self.session["suites"][1]["closed"] = "CLOSED; lost=2"
        self.save_session()
        self.assertEqual(self.run_writer()[0], 1)
        self.assertIn("lines lost", self.summary()["suites"][1]["detail"])
        (self.bundle / "host.json").unlink()
        self.session["suites"][1].update(closed="CLOSED; lost=0", status="ERROR", error="Run-time error")
        self.save_session()
        self.assertEqual(self.run_writer()[0], 1)
        self.assertIn("raised", self.summary()["suites"][1]["detail"])

    def test_raised_entry_point_is_named_before_log_closing(self) -> None:
        self.session["suites"][1].update(status="ERROR", error="Regression failed", closed="NONE")
        self.save_session()
        self.assertEqual(self.run_writer()[0], 1)
        detail = self.summary()["suites"][1]["detail"]
        self.assertIn("the entry point raised: Regression failed", detail)
        self.assertNotIn("lines lost", detail)

    def test_example_error_fails_the_certification(self) -> None:
        self.session["examples"][0].update(status="ERROR", error="Subscript out of range")
        self.save_session()
        self.assertEqual(self.run_writer()[0], 1)
        self.assertEqual(self.summary()["examples"][0]["status"], "FAIL")
        self.assertEqual(host.evaluate(self.root, self.sha, self.bundle / "host.json")["outcomes"], ["PASS"])

    def test_configured_suite_or_example_missing_from_the_run_fails(self) -> None:
        self.session["suites"] = self.session["suites"][:1]
        self.save_session()
        code, out, _ = self.run_writer()
        self.assertEqual(code, 1, out)
        suite = self.summary()["suites"][1]
        self.assertEqual((suite["module"], suite["status"]), ("TextTests", "FAIL"))
        self.assertIn("not run", suite["detail"])
        self.assertIn("| `TextTests.RunTextTests` | FAIL |", (self.bundle / "certification.md").read_text())
        (self.bundle / "host.json").unlink()
        self.write_bundle()
        self.session["examples"] = []
        self.save_session()
        self.assertEqual(self.run_writer()[0], 1)
        example = self.summary()["examples"][0]
        self.assertEqual((example["module"], example["status"]), ("ProjectExample", "FAIL"))
        self.assertIn("RESULT=FAIL", (self.bundle / "certification.md").read_text())

    def test_record_is_never_replaced_and_candidate_must_match(self) -> None:
        (self.bundle / "host.json").write_text("{}")
        code, _, err = self.run_writer()
        self.assertEqual((code, "never replaced" in err), (2, True))
        (self.bundle / "host.json").unlink()
        self.session["candidate_sha"] = "0" * 40
        self.save_session()
        code, _, err = self.run_writer()
        self.assertEqual(code, 2)
        self.assertIn("names candidate", err)
        self.assertFalse((self.bundle / "host.json").exists())


if __name__ == "__main__":
    unittest.main()
