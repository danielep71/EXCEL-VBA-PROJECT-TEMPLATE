#!/usr/bin/env python3
"""Offline fixtures for prepare_release_evidence.py: every external call is simulated."""

from __future__ import annotations

import contextlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_excel_evidence as excel  # noqa: E402
import prepare_release_evidence as tool  # noqa: E402

SHA = "a" * 40
REPO = "owner/template"
HOST = {"execution": "manual", "harness": {"entry_point": "ProjectTests.RunProjectTests", "cases": 4,
                                           "assertions": 6, "failures": 0, "completeness": "COMPLETE"}}


def run_record(path: str, conclusion: str = "success", run_id: int = 1, sha: str = SHA) -> dict[str, Any]:
    return {"id": run_id, "name": Path(path).stem, "path": path, "head_sha": sha, "event": "push",
            "status": "completed", "conclusion": conclusion, "created_at": f"2026-11-14T10:0{run_id}:00Z",
            "html_url": f"https://github.com/{REPO}/actions/runs/{run_id}"}


def rulesets(bypass: list[Any] | None = None, strict: bool = True) -> list[dict[str, Any]]:
    branch = {"id": 1, "name": "Protect main", "target": "branch", "enforcement": "active",
              "bypass_actors": bypass or [], "conditions": {"ref_name": {"include": ["~DEFAULT_BRANCH"]}},
              "rules": [{"type": "deletion"}, {"type": "non_fast_forward"}, {"type": "pull_request"},
                        {"type": "required_status_checks", "parameters": {
                            "strict_required_status_checks_policy": strict,
                            "required_status_checks": [{"context": "Repository integrity"}]}}]}
    tag = {"id": 2, "name": "Protect releases", "target": "tag", "enforcement": "active", "bypass_actors": [],
           "conditions": {"ref_name": {"include": ["refs/tags/v*"]}},
           "rules": [{"type": "deletion"}, {"type": "non_fast_forward"}, {"type": "update"}]}
    return [branch, tag]


def links_zip(defects: int = 0) -> bytes:
    report = {"status": "pass" if defects == 0 else "fail", "counts": {
        "deterministic_public_defects": defects, "restricted_historical": 21, "pending_publication": 2,
        "access_restricted": 40, "transient_failures": 0}}
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("external-links.json", json.dumps(report))
    return buffer.getvalue()


class Fake:
    """Simulated git, gh and sibling tools; records every command."""

    def __init__(self, runs: list[dict[str, Any]], head: str = SHA, dirty: str = "",
                 fail: str = "", governance: list[dict[str, Any]] | None = None, defects: int = 0) -> None:
        self.runs, self.head, self.dirty, self.fail = runs, head, dirty, fail
        self.governance, self.defects = governance or rulesets(), defects
        self.commands: list[list[str]] = []

    def __call__(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
        self.commands.append(command)
        text = " ".join(command)
        if self.fail and self.fail in text:
            return subprocess.CompletedProcess(command, 1, b"", b"simulated failure\n")
        out: Any = b""
        if command[:2] == ["git", "rev-parse"]:
            out = self.head.encode()
        elif command[:2] == ["git", "status"]:
            out = self.dirty.encode()
        elif command[:2] == ["git", "clone"]:
            clone = Path(command[-1])
            clone.mkdir(parents=True)
            (clone / "Home.md").write_text("home\n")
            (clone / "Wiki-Source.json").write_text("{}\n")
        elif "log" in command and command[0] == "git":
            out = b"b" * 40 + b"\n" + b"c" * 40 + b"\n"
        elif command[:2] == ["gh", "api"]:
            out = self.gh(command[2])
        elif "release_certification.py" in text and "--build" in command:
            target = Path(command[command.index("--output-dir") + 1])
            target.mkdir(parents=True)
            (target / "certification-v1.3.0.zip").write_bytes(b"zip")
        elif "check_release.py" in text and "--output" in command:
            output = Path(command[command.index("--output") + 1])
            output.write_text(json.dumps({"status": "pass", "findings": []}))
        return subprocess.CompletedProcess(command, 0, out if isinstance(out, bytes) else json.dumps(out).encode(), b"")

    def gh(self, endpoint: str) -> Any:
        if "actions/runs?event=push" in endpoint:
            return {"workflow_runs": self.runs}
        if endpoint.endswith("/rulesets"):
            return [{"id": item["id"]} for item in self.governance]
        if "/rulesets/" in endpoint:
            return next(item for item in self.governance if str(item["id"]) == endpoint.rsplit("/", 1)[1])
        if "external-links.yml/runs" in endpoint:
            return {"workflow_runs": [run_record(".github/workflows/external-links.yml", run_id=7)]}
        if endpoint.endswith("/artifacts"):
            return {"artifacts": [{"id": 99}]}
        if endpoint.endswith("/zip"):
            return links_zip(self.defects)
        raise AssertionError(f"unexpected endpoint {endpoint}")


class PrepareReleaseEvidenceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        base = Path(self.temporary.name)
        self.root, self.work, self.bundle = base / "repo", base / "work", base / "excel"
        (self.root / ".github").mkdir(parents=True)
        self.bundle.mkdir()
        (self.root / "VERSION").write_text("1.3.0\n")
        self.profile({"mode": "template", "repository": REPO})
        (self.root / ".github/release-policy.json").write_text(json.dumps({
            "core_checks": ["repository-integrity", "vba-compile", "regression"],
            "profiles": {"template": {"required_checks": ["generated-profile-pilots", "live-governance"]},
                         "library": {"required_checks": ["public-api", "caller-contract"]}}}))
        (self.bundle / "host.json").write_text(json.dumps(HOST))
        (self.bundle / "harness.log").write_text("RESULT=PASS\n")
        (self.bundle / "certification.md").write_text("# summary\n")
        (self.bundle / "Certification.xlsm").write_bytes(b"workbook")
        self.pilot = base / "clean-room-journey.md"
        self.pilot.write_text("# journey\n")
        patches = (patch.object(excel, "evaluate", return_value={"status": "pass", "outcomes": ["PASS"]}),
                   patch.object(excel, "environment_summary", return_value="execution=manual; environment={}"))
        for item in patches:
            item.start()
            self.addCleanup(item.stop)
        self.runs = [run_record(tool.STATIC_WORKFLOW, run_id=1), run_record(tool.MATRIX_WORKFLOW, run_id=2)]

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def profile(self, value: dict[str, Any]) -> None:
        (self.root / ".github/repository-profile.json").write_text(json.dumps(value))

    def main(self, fake: Fake, *arguments: str) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with patch.object(tool, "run", fake), contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = tool.main(list(arguments))
        return code, out.getvalue(), err.getvalue()

    def prepare(self, fake: Fake, *extra: str) -> tuple[int, str, str]:
        return self.main(fake, "prepare", "--root", str(self.root), "--tag", "v1.3.0", "--candidate-sha", SHA,
                         "--work-dir", str(self.work), "--excel-bundle", str(self.bundle), *extra)

    def bundle_command(self, fake: Fake, *extra: str) -> tuple[int, str, str]:
        return self.main(fake, "bundle", "--root", str(self.root), "--tag", "v1.3.0", "--candidate-sha", SHA,
                         "--work-dir", str(self.work), *extra)

    def test_template_prepare_then_bundle(self) -> None:
        fake = Fake(self.runs)
        code, out, err = self.prepare(fake, "--wiki-browser-reviewed", "--pilot-record", str(self.pilot),
                                      "--pilot-summary", "Pilot released v0.1.0.")
        self.assertEqual(code, 0, err)
        self.assertIn("PREPARED", out)
        evidence = self.work / "evidence"
        self.assertEqual(sorted(path.name for path in evidence.iterdir()), [
            "certification.md", "clean-room-journey.md", "external-links.json", "gate-evidence.md",
            "harness.log", "host.json", "live-governance.json", "release-evidence.json",
            "release-integrity-pretag.json", "wiki-publication.json"])
        release = json.loads((self.work / "release-evidence.json").read_text())
        checks = release["checks"]
        self.assertEqual(sorted(checks), ["excel-host-evidence", "generated-profile-pilots", "live-governance",
                                          "regression", "repository-integrity", "vba-compile"])
        self.assertEqual(checks["repository-integrity"]["run_url"], f"https://github.com/{REPO}/actions/runs/1")
        self.assertEqual(checks["regression"]["cases"], 4)
        self.assertEqual(checks["excel-host-evidence"]["execution"], "manual")
        self.assertEqual(len(checks["excel-host-evidence"]["sha256"]), 64)
        self.assertIn("Pilot released v0.1.0.", checks["generated-profile-pilots"]["detail"])
        wiki = json.loads((evidence / "wiki-publication.json").read_text())
        self.assertEqual((wiki["wiki_commit"], wiki["wiki_parent"]), ("b" * 40, "c" * 40))
        self.assertIn("0 deterministic defects", (evidence / "gate-evidence.md").read_text())
        self.assertFalse((evidence / "Certification.xlsm").exists())

        self.assertEqual(self.bundle_command(fake)[0], 1)
        code, out, err = self.bundle_command(fake, "--redaction-reviewed")
        self.assertEqual(code, 0, err)
        self.assertIn("BUNDLED", out)
        spec = json.loads((self.work / "certification-input.json").read_text())
        roles = {record["path"]: record["role"] for record in spec["records"]}
        self.assertEqual(roles["host.json"], "excel-evidence")
        self.assertEqual(roles["harness.log"], "excel-harness")
        self.assertEqual(roles["certification.md"], "excel-certification")
        self.assertEqual(roles["clean-room-journey.md"], "clean-room-journey")
        self.assertEqual(roles["release-integrity-posttag.json"], "release-integrity")
        self.assertTrue(any("--require-tag-ref" in command for command in fake.commands))
        scripts = [command[1] for command in fake.commands if command[0] == sys.executable]
        self.assertTrue(scripts)
        self.assertTrue(all(script.startswith(str(self.root / "tools")) for script in scripts), scripts)
        # A second bundle run replaces the earlier outputs.
        self.assertEqual(self.bundle_command(fake, "--redaction-reviewed")[0], 0)

    def test_generated_profile_needs_explicit_profile_checks(self) -> None:
        self.profile({"mode": "generated", "repository": REPO, "profile": "library"})
        fake = Fake([run_record(tool.STATIC_WORKFLOW)])
        code, _, err = self.prepare(fake)
        self.assertEqual(code, 1)
        self.assertIn("public-api, caller-contract", err)
        self.work = self.work.with_name("work2")
        code, _, err = self.prepare(Fake([run_record(tool.STATIC_WORKFLOW)]),
                                    "--check", "public-api=check_vba_public_api.py PASS",
                                    "--check", "caller-contract=TextExample ran")
        self.assertEqual(code, 0, err)
        checks = json.loads((self.work / "release-evidence.json").read_text())["checks"]
        self.assertEqual(checks["public-api"]["detail"], "check_vba_public_api.py PASS")
        self.assertFalse((self.work / "evidence/live-governance.json").exists())
        code, _, err = self.bundle_command(Fake([]), "--redaction-reviewed")
        self.assertEqual((code, "canonical template only" in err), (1, True))

    def test_candidate_identity_failures(self) -> None:
        cases = ((Fake(self.runs, head="b" * 40), "not the candidate"),
                 (Fake(self.runs, dirty=" M VERSION"), "uncommitted changes"),
                 (Fake(self.runs, fail="git rev-parse"), "git rev-parse failed: simulated failure"))
        for fake, message in cases:
            with self.subTest(message=message):
                code, _, err = self.prepare(fake)
                self.assertEqual(code, 1)
                self.assertIn(message, err)
        (self.root / "VERSION").write_text("1.2.9\n")
        self.assertIn("does not match VERSION", self.prepare(Fake(self.runs))[2])

    def test_hosted_and_live_failures(self) -> None:
        cases = (
            (Fake([run_record(tool.STATIC_WORKFLOW, "failure")]), "concluded failure"),
            (Fake([run_record(tool.STATIC_WORKFLOW)]), "no completed push run of " + tool.MATRIX_WORKFLOW),
            (Fake(self.runs, governance=rulesets(bypass=[{"actor_id": 5}])), "no bypass actors"),
            (Fake(self.runs, governance=rulesets(strict=False)), "'Repository integrity' required, strict"),
            (Fake(self.runs, governance=rulesets()[:1]), "no active tag ruleset"),
            (Fake(self.runs, defects=2), "status is fail"),
            (Fake(self.runs, fail="check_wiki.py"), "Wiki read-back comparison"),
            (Fake(self.runs, fail="check_repo.py --root . --self-test"), "check_repo.py --root . --self-test failed"),
            (Fake(self.runs, fail="--output"), "pre-tag check_release.py failed"),
        )
        for index, (fake, message) in enumerate(cases):
            with self.subTest(message=message):
                self.work = self.work.with_name(f"case{index}")
                code, _, err = self.prepare(fake, "--wiki-browser-reviewed")
                self.assertEqual(code, 1)
                self.assertIn(message, err)
        self.work = self.work.with_name("unreviewed")
        self.assertIn("--wiki-browser-reviewed", self.prepare(Fake(self.runs))[2])

    def test_input_and_reuse_failures(self) -> None:
        (self.bundle / "host.json").unlink()
        self.assertIn("has no host.json", self.prepare(Fake(self.runs))[2])
        (self.bundle / "host.json").write_text(json.dumps(HOST))
        with patch.object(excel, "evaluate", return_value={"status": "fail", "outcomes": ["EVIDENCE_INVALID"]}):
            self.work = self.work.with_name("invalid")
            self.assertIn("EVIDENCE_INVALID", self.prepare(Fake(self.runs))[2])
        self.work.mkdir(exist_ok=True)
        (self.work / "evidence").mkdir(exist_ok=True)
        self.assertIn("fresh --work-dir", self.prepare(Fake(self.runs))[2])
        self.work = self.work.with_name("nothing-prepared")
        self.assertIn("run prepare first", self.bundle_command(Fake(self.runs), "--redaction-reviewed")[2])
        self.work = self.work.with_name("bad-check")
        self.assertIn("NAME=DETAIL", self.prepare(Fake(self.runs), "--check", "no-separator")[2])
        self.assertFalse(self.work.exists())

    def test_bundle_rejects_duplicate_roles_and_nondeterminism(self) -> None:
        evidence = self.work / "evidence"
        evidence.mkdir(parents=True)
        (evidence / "a_b.md").write_text("x")
        (evidence / "a-b.md").write_text("y")
        (evidence / "nested").mkdir()
        identity = {"repository": REPO, "version": "1.3.0", "tag": "v1.3.0", "sha": SHA}
        with self.assertRaisesRegex(tool.EvidenceError, "same bundle role"):
            tool.bundle_spec(identity, evidence)
        (evidence / "a-b.md").unlink()
        (self.work / "release-evidence.json").write_text("{}")

        class Drifting(Fake):
            builds = 0

            def __call__(self, command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
                result = super().__call__(command, cwd)
                if "--build" in command:
                    Drifting.builds += 1
                    target = Path(command[command.index("--output-dir") + 1])
                    (target / "certification-v1.3.0.zip").write_bytes(f"zip{Drifting.builds}".encode())
                return result

        code, _, err = self.bundle_command(Drifting(self.runs), "--redaction-reviewed")
        self.assertEqual((code, "not byte-identical" in err), (1, True))

    def test_run_boundary_and_failure_without_stderr(self) -> None:
        completed = tool.run([sys.executable, "-c", "print('ok')"], Path("."))
        self.assertEqual(completed.stdout.strip(), b"ok")

        def silent(command: list[str], cwd: Path) -> subprocess.CompletedProcess[bytes]:
            return subprocess.CompletedProcess(command, 3, b"", b"")

        with patch.object(tool, "run", silent), self.assertRaisesRegex(tool.EvidenceError, "exit 3"):
            tool.checked(["anything"], Path("."), "probe")


if __name__ == "__main__":
    unittest.main()
