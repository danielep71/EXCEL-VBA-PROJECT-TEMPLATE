"""Offline drift and simulated external HTTP failure boundaries."""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import tempfile
import threading
import time
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

import check_documentation as docs
import check_external_links as links

ROOT = Path(__file__).resolve().parents[1]
TODAY = date(2026, 9, 9)


def assert_readme_presentation(testcase: unittest.TestCase, root: Path) -> None:
    """Validate canonical-template or generated-project README ownership boundaries."""
    readme = (root / "README.md").read_text(encoding="utf-8")
    profile = json.loads(
        (root / ".github/repository-profile.json").read_text(encoding="utf-8")
    )

    if profile.get("mode") == "generated":
        record = json.loads(
            (root / ".github/initialization.json").read_text(encoding="utf-8")
        )
        testcase.assertEqual(
            [line for line in readme.splitlines() if line.startswith("# ")],
            [f"# ⚡ {record['values']['PROJECT_NAME']}"],
        )
        repository = profile["repository"]
        testcase.assertIn(
            f"https://github.com/{repository}/actions/workflows/static-checks.yml",
            readme,
        )
        testcase.assertNotIn("{" + "{", readme)
        template_marker = "<!-- " + "template:"
        testcase.assertNotIn(template_marker, readme)

        preview = record["values"].get("SOCIAL_PREVIEW_PATH")
        if preview:
            testcase.assertTrue((root / preview).is_file())
            testcase.assertIn(
                f"<!-- generated-social-preview: {preview} -->",
                readme,
            )
        else:
            testcase.assertNotIn('src="assets/social-preview.png"', readme)
            testcase.assertNotIn("<!-- generated-social-preview:", readme)
        return

    repository_token = "{" + "{REPOSITORY_PATH}" + "}"
    preview_token = "{" + "{SOCIAL_PREVIEW_PATH}" + "}"
    testcase.assertNotIn(f"https://github.com/{repository_token}", readme)
    testcase.assertNotIn(
        f"https://api.scorecard.dev/projects/github.com/{repository_token}",
        readme,
    )
    testcase.assertNotIn(
        f"https://scorecard.dev/viewer/?uri=github.com/{repository_token}",
        readme,
    )
    testcase.assertNotIn("securityscorecards.dev", readme)
    canonical_repository = "danielep71/" + "EXCEL-VBA-" + "PROJECT-TEMPLATE"
    testcase.assertIn(
        f"https://api.scorecard.dev/projects/github.com/{canonical_repository}/badge",
        readme,
    )
    testcase.assertIn('src="assets/social-preview.png"', readme)
    testcase.assertIn(
        f"<!-- generated-social-preview: {preview_token} -->",
        readme,
    )


@unittest.skipUnless(shutil.which("bash"), "Bash required for release command fixture")
class ReleaseCommandTests(unittest.TestCase):
    """Execute documented tag-publication blocks with offline command stubs."""

    def release_block(self, heading: str) -> str | None:
        text = (ROOT / "RELEASING.md").read_text(encoding="utf-8")
        if heading not in text:
            return None
        section = text.split(heading, 1)[1]
        return section.split("```bash\n", 1)[1].split("```", 1)[0]

    def exercise(self, heading: str, *, signed: bool) -> None:
        block = self.release_block(heading)
        if block is None:
            self.skipTest(f"{heading} is not applicable in this generated repository")
        stages = ["switch", "pull", "rev-parse", "version", "precheck", "tag", "postcheck", "push"]
        stubs = r'''
record() { printf '%s\n' "$1" >> calls.log; [ "$1" != "$fail_at" ]; }
git() {
  stage=""
  for arg in "$@"; do
    case "$arg" in
      switch|pull|rev-parse|tag|push) stage="$arg"; break ;;
    esac
  done
  record "$stage" || return 23
  case "$stage" in
    rev-parse) printf '%040d\n' 1 ;;
    tag)
      case " $* " in
        *" 0000000000000000000000000000000000000001 "*) ;;
        *) return 24 ;;
      esac
      ;;
    push)
      case " $* " in
        *" refs/tags/v1.0.0:refs/tags/v1.0.0 "*) ;;
        *) return 25 ;;
      esac
      ;;
  esac
}
tr() { record version || return 23; printf '1.0.0'; }
python3() {
  case " $* " in
    *" --require-tag-ref "*) record postcheck ;;
    *) record precheck ;;
  esac
}
'''
        for fail_at in ["none", *stages]:
            with self.subTest(heading=heading, fail_at=fail_at), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                (root / "VERSION").write_text("1.0.0\n", encoding="utf-8")
                env_prefix = "RELEASE_SIGNING_KEY=/tmp/offline-fixture-key\n" if signed else ""
                result = subprocess.run(
                    ["bash", "-c", env_prefix + f"fail_at={fail_at}\n" + stubs + block],
                    cwd=root,
                    capture_output=True,
                    text=True,
                    check=False,
                )
                expected = stages if fail_at == "none" else stages[: stages.index(fail_at) + 1]
                self.assertEqual((root / "calls.log").read_text().splitlines(), expected)
                self.assertEqual(result.returncode == 0, fail_at == "none", result.stderr)

    def test_generated_release_sequence_stops_at_every_failure(self) -> None:
        self.exercise("### Initialized generated project", signed=False)

    def test_canonical_signed_release_sequence_stops_at_every_failure(self) -> None:
        self.exercise("### Canonical-template SSH-signed tag", signed=True)

    def test_canonical_signing_key_is_required_before_validation(self) -> None:
        block = self.release_block("### Canonical-template SSH-signed tag")
        if block is None:
            self.skipTest("Canonical signed-tag block is not applicable in generated mode")
        stubs = r'''
git() {
  case "$1" in
    rev-parse) printf '%040d\n' 1 ;;
    *) return 0 ;;
  esac
}
tr() { printf '1.0.0'; }
python3() { printf 'unexpected validation\n' >> unexpected.log; return 0; }
'''
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "VERSION").write_text("1.0.0\n", encoding="utf-8")
            result = subprocess.run(
                ["bash", "-c", stubs + block],
                cwd=root,
                capture_output=True,
                text=True,
                check=False,
                env={},
            )
            self.assertNotEqual(result.returncode, 0)
            self.assertFalse((root / "unexpected.log").exists())



class DocumentationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / ".github/workflows").mkdir(parents=True)
        (self.root / "tools").mkdir()
        self.policy = json.loads((ROOT / docs.POLICY).read_text())
        self.policy["references"] = []
        self.policy["historical_documents"] = {}
        self.policy["network"]["domains"] = {"example.org": "Synthetic documentation service"}
        self.policy["network"]["classifications"] = []
        (self.root / "README.md").write_text("# Fixture\n\npython3 tools/fixture.py --root .\n")
        (self.root / "tools/fixture.py").write_text("import argparse\np=argparse.ArgumentParser()\np.add_argument('--root')\n")
        subprocess.run(["git", "init", "-q", str(self.root)], check=True)
        self.save()

    def save(self):
        (self.root / docs.POLICY).write_text(json.dumps(self.policy))
        subprocess.run(["git", "-C", str(self.root), "add", "--all"], check=True)

    def test_documented_command_passes_without_execution(self):
        with patch.object(subprocess, "Popen", wraps=subprocess.Popen) as popen:
            self.assertEqual(docs.build_report(self.root)["status"], "pass")
            self.assertTrue(all(call.args[0][0] == "git" for call in popen.call_args_list))

    def test_readme_presentation_uses_repository_identity_and_selected_assets(self):
        assert_readme_presentation(self, ROOT)

    def _generated_readme_fixture(
        self,
        *,
        preview: bool,
        badge: str = "status",
        repository: str = "example/generated-project",
    ) -> Path:
        fixture = self.root / f"generated-{'preview' if preview else 'no-preview'}-{badge}"
        (fixture / ".github").mkdir(parents=True)
        values = {
            "PROJECT_NAME": "Generated Project",
            "SOCIAL_PREVIEW_PATH": "assets/social-preview.png" if preview else None,
        }
        (fixture / ".github/repository-profile.json").write_text(
            json.dumps({
                "mode": "generated",
                "profile": "library",
                "repository": repository,
            }),
            encoding="utf-8",
        )
        (fixture / ".github/initialization.json").write_text(
            json.dumps({"values": values}),
            encoding="utf-8",
        )
        readme = (
            "# ⚡ Generated Project\n\n"
            f"[{badge}](https://github.com/{repository}/actions/workflows/static-checks.yml)\n"
        )
        if preview:
            (fixture / "assets").mkdir()
            (fixture / "assets/social-preview.png").write_bytes(b"fixture")
            readme += (
                '\n<img src="assets/social-preview.png" alt="preview">\n'
                "<!-- generated-social-preview: assets/social-preview.png -->\n"
            )
        (fixture / "README.md").write_text(readme, encoding="utf-8")
        return fixture

    def test_generated_readme_accepts_preview_present_and_absent(self):
        for preview in (True, False):
            with self.subTest(preview=preview):
                fixture = self._generated_readme_fixture(preview=preview)
                assert_readme_presentation(self, fixture)

    def test_generated_readme_allows_project_owned_badge_text(self):
        for badge in ("status", "build", "quality-gate"):
            with self.subTest(badge=badge):
                fixture = self._generated_readme_fixture(preview=False, badge=badge)
                assert_readme_presentation(self, fixture)

    def test_generated_readme_rejects_wrong_repository_identity(self):
        fixture = self._generated_readme_fixture(preview=False)
        readme = (fixture / "README.md").read_text(encoding="utf-8")
        (fixture / "README.md").write_text(
            readme.replace("example/generated-project", "example/wrong-project"),
            encoding="utf-8",
        )
        with self.assertRaises(AssertionError):
            assert_readme_presentation(self, fixture)

    def test_generated_readme_rejects_residual_template_markers(self):
        fixture = self._generated_readme_fixture(preview=False)
        with (fixture / "README.md").open("a", encoding="utf-8") as handle:
            marker = "<!-- " + "template:optional:SOCIAL_PREVIEW_PATH -->"
            handle.write("\n" + marker + "\n" + "{" + "{PROJECT_NAME}" + "}\n")
        with self.assertRaises(AssertionError):
            assert_readme_presentation(self, fixture)

    def test_generated_readme_rejects_missing_selected_preview(self):
        fixture = self._generated_readme_fixture(preview=True)
        (fixture / "assets/social-preview.png").unlink()
        with self.assertRaises(AssertionError):
            assert_readme_presentation(self, fixture)

    def test_generated_readme_rejects_canonical_only_presentation_assertion(self):
        fixture = self._generated_readme_fixture(preview=False)
        readme = (fixture / "README.md").read_text(encoding="utf-8")
        canonical_repository = "danielep71/" + "EXCEL-VBA-" + "PROJECT-TEMPLATE"
        self.assertNotIn(
            f"https://api.scorecard.dev/projects/github.com/{canonical_repository}/badge",
            readme,
        )
        self.assertNotIn('src="assets/social-preview.png"', readme)


    def test_retained_documentation_semantics_match_repository_mode(self):
        profile = json.loads(
            (ROOT / ".github/repository-profile.json").read_text(encoding="utf-8")
        )
        tools_readme = (ROOT / "tools/README.md").read_text(encoding="utf-8")
        initialization = (ROOT / "docs/INITIALIZATION.md").read_text(encoding="utf-8")
        house_style = (ROOT / "docs/VBA_HOUSE_STYLE.md").read_text(encoding="utf-8")
        contract = (ROOT / "docs/TEMPLATE_CONTRACT.md").read_text(encoding="utf-8")
        docs_index = (ROOT / "docs/README.md").read_text(encoding="utf-8")
        releasing = (ROOT / "RELEASING.md").read_text(encoding="utf-8")

        self.assertNotIn("Exercise all three profile fixtures with:", tools_readme)
        self.assertNotIn("initializer self-test for all three profiles", house_style)
        self.assertNotIn("pending 1.2.0 contract", contract)
        self.assertIn("External provenance-record SSH signatures", contract)
        self.assertIn("canonical Git-tag and certification-bundle signatures", contract)
        self.assertIn("SUPPLY_CHAIN_ASSURANCE.md", docs_index)
        self.assertIn("RELEASE_CLOSEOUT.md", docs_index)
        self.assertNotIn("protected release path", releasing)
        self.assertNotIn("protected annotated tag", releasing)
        self.assertIn("verify the required branch/tag protection", releasing)

        if profile.get("mode") == "generated":
            self.assertIn(
                "In an initialized generated repository it\nvalidates only the recorded selected profile",
                tools_readme,
            )
            self.assertIn(
                "In an initialized\ngenerated repository, the same command validates only the recorded selected",
                initialization,
            )
            for text in (tools_readme, initialization, house_style):
                self.assertNotRegex(
                    text,
                    r"(?m)^\s*python3 tools/(?:checker_development|check_policy_coverage)\.py",
                )

    def test_utf8_repository_reads_do_not_depend_on_locale(self):
        workflow = self.root / ".github/workflows/fixture.yml"
        workflow.write_text(
            "name: UTF-8 workflow 🔐\njobs:\n  check:\n    name: UTF-8 context\n    runs-on: ubuntu-24.04\n",
            encoding="utf-8",
        )
        (self.root / "README.md").write_text(
            "# Fixture 🔐\n\nUTF-8 workflow 🔐\n\npython3 tools/fixture.py --root .\n",
            encoding="utf-8",
        )
        self.policy["references"] = [{
            "document": "README.md",
            "target": ".github/workflows/fixture.yml",
            "token": "UTF-8 workflow 🔐",
            "kind": "workflow-name",
        }]
        self.save()
        original = Path.read_text

        def require_explicit_utf8(path, *args, **kwargs):
            encoding = kwargs.get("encoding")
            if encoding is None and args:
                encoding = args[0]
            if encoding is None:
                raise UnicodeDecodeError(
                    "charmap", b"\x8f", 0, 1, "character maps to <undefined>"
                )
            self.assertEqual(encoding, "utf-8")
            return original(path, *args, **kwargs)

        with patch.object(Path, "read_text", require_explicit_utf8):
            self.assertEqual(docs.build_report(self.root)["status"], "pass")

    def test_renamed_command_detected(self):
        (self.root / "tools/fixture.py").rename(self.root / "tools/renamed.py")
        self.save()
        self.assertEqual(docs.build_report(self.root)["status"], "fail")

    def test_removed_cli_option_detected(self):
        (self.root / "tools/fixture.py").write_text("import argparse\n")
        self.assertEqual(docs.build_report(self.root)["status"], "fail")

    def test_shared_runner_does_not_grant_parser_flags(self):
        (self.root / "tools/fixture.py").write_text("from _gatelib import run_gate\n")
        (self.root / "tools/_gatelib.py").write_text("p.add_argument('--self-test')\n")
        self.assertNotIn("--self-test", docs.command_flags(self.root, "tools/fixture.py"))
        (self.root / "tools/fixture.py").write_text("from _gatelib import parse_report_args\n")
        self.assertIn("--self-test", docs.command_flags(self.root, "tools/fixture.py"))

    def test_multiline_and_inline_commands(self):
        text = "`python3 tools/a.py --root .`\npython3 tools/b.py \\\n  --output out.json\n"
        self.assertEqual(docs.commands(text), [("tools/a.py", {"--root"}), ("tools/b.py", {"--output"})])

    def test_workflow_and_context_renames_detected(self):
        workflow = self.root / ".github/workflows/fixture.yml"
        workflow.write_text("name: Example workflow\njobs:\n  check:\n    name: Example context\n    runs-on: ubuntu-24.04\n")
        (self.root / "README.md").write_text("Example workflow; Example context\n")
        self.policy["references"] = [
            {"document": "README.md", "target": ".github/workflows/fixture.yml", "token": "Example workflow", "kind": "workflow-name"},
            {"document": "README.md", "target": ".github/workflows/fixture.yml", "token": "Example context", "kind": "job-name", "job": "check"}]
        self.save()
        self.assertEqual(docs.build_report(self.root)["status"], "pass")
        workflow.write_text(workflow.read_text().replace("Example context", "Changed context"))
        self.assertEqual(docs.build_report(self.root)["status"], "fail")
        workflow.write_text(workflow.read_text().replace("Example workflow", "Changed workflow"))
        self.assertEqual(len(docs.build_report(self.root)["findings"]), 2)

    def test_filename_and_policy_value_drift(self):
        (self.root / "value.json").write_text('{"required": true}')
        (self.root / "README.md").write_text("Required value.json\n")
        self.policy["references"] = [{"document": "README.md", "target": "value.json", "token": "Required",
                                     "kind": "json-value", "pointer": "/required", "value": True}]
        self.save()
        self.assertEqual(docs.build_report(self.root)["status"], "pass")
        (self.root / "value.json").write_text('{"required": false}')
        self.assertEqual(docs.build_report(self.root)["status"], "fail")
        (self.root / "value.json").unlink()
        self.save()
        self.assertEqual(docs.build_report(self.root)["status"], "fail")

    def probe(self, responses):
        calls = []
        sleeps = []
        def transport(url, timeout):
            calls.append((url, timeout))
            value = responses[len(calls) - 1]
            if isinstance(value, Exception):
                raise value
            return value
        report = links.probe("https://example.org/page", self.policy["network"], transport, sleeps.append)
        return report, calls, sleeps

    def test_consistent_404_retried_and_reported(self):
        report, calls, sleeps = self.probe([(404, None)] * 3)
        self.assertEqual(report["status"], "PERMANENT_FAILURE")
        self.assertEqual(len(calls), 3)
        self.assertEqual(sleeps, [1, 2])

    def test_public_404_remains_deterministic_public_defect(self):
        url = "https://example.org/missing"
        (self.root / "README.md").write_text(f"[missing]({url})\n")
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: (404, None), lambda *args: None)
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["links"][0]["status"], "PERMANENT_FAILURE")
        self.assertEqual(report["counts"]["deterministic_public_defects"], 1)
        self.assertEqual(report["counts"]["restricted_historical"], 0)

    def test_restricted_historical_is_non_green_and_not_probed(self):
        url = "https://example.org/private-history"
        identifier = hashlib.sha256(url.encode()).hexdigest()
        (self.root / "README.md").write_text(f"[history]({url})\n")
        self.policy["network"]["classifications"] = [{
            "id": identifier,
            "kind": "restricted-historical",
            "reason": "Authenticated historical evidence",
            "expires": "2026-09-10",
        }]
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("classified target must not be probed"))
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["links"][0]["status"], "RESTRICTED_HISTORICAL")
        self.assertEqual(report["counts"]["restricted_historical"], 1)
        self.assertEqual(report["counts"]["deterministic_public_defects"], 0)
        self.assertNotIn("private-history", json.dumps(report) + links.markdown(report))

    def test_pending_publication_is_distinct_non_green_classification(self):
        url = "https://example.org/compare/v1.0.0...v1.1.0"
        identifier = hashlib.sha256(url.encode()).hexdigest()
        (self.root / "README.md").write_text(f"[future-tag]({url})\n")
        self.policy["network"]["classifications"] = [{
            "id": identifier,
            "kind": "pending-publication",
            "reason": "Reviewed candidate link depends on a tag not published yet",
            "expires": "2026-09-10",
        }]
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("pending target must not be probed"))
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["links"][0]["status"], "PENDING_PUBLICATION")
        self.assertEqual(report["counts"]["pending_publication"], 1)
        self.assertEqual(report["counts"]["deterministic_public_defects"], 0)

    def test_classification_cannot_override_local_url_policy(self):
        cases = (
            ("http://example.org/private-history", "POLICY_BLOCKED"),
            ("https://example.org/private-history?token=fixture", "ACCESS_RESTRICTED"),
        )
        for url, expected in cases:
            with self.subTest(url=url):
                identifier = hashlib.sha256(url.encode()).hexdigest()
                (self.root / "README.md").write_text(f"[classified]({url})\n")
                self.policy["network"]["classifications"] = [{
                    "id": identifier,
                    "kind": "restricted-historical",
                    "reason": "Reviewed historical classification",
                    "expires": "2026-09-10",
                }]
                self.save()
                report = links.build_report(
                    self.root,
                    TODAY,
                    lambda *args: self.fail("policy-rejected classified target must not be probed"),
                )
                self.assertEqual(report["links"][0]["status"], expected)
                self.assertEqual(report["links"][0]["attempts"], 0)
                self.assertEqual(report["status"], "fail")

    def test_markdown_exposes_every_json_count_category(self):
        rows = [
            {"status": "PERMANENT_FAILURE"},
            {"status": "RESTRICTED_HISTORICAL"},
            {"status": "PENDING_PUBLICATION"},
            {"status": "ACCESS_RESTRICTED"},
            {"status": "TRANSIENT_FAILURE"},
        ]
        counts = links._counts(rows)
        self.assertEqual(set(counts), set(links.COUNT_LABELS))
        report = {
            "status": "fail",
            "discovered": len(rows),
            "limit_exceeded": False,
            "counts": counts,
            "links": [],
            "scope_note": "fixture scope",
        }
        rendered = links.markdown(report)
        for key, label in links.COUNT_LABELS.items():
            self.assertIn(f"{label}: {counts[key]}", rendered)

    def test_transient_then_recovery(self):
        report, calls, _ = self.probe([(503, None), (200, None)])
        self.assertEqual(report["status"], "OK")
        self.assertEqual(len(calls), 2)

    def test_mixed_missing_transient_not_permanent(self):
        report, _, _ = self.probe([(404, None), TimeoutError("secret URL"), (404, None)])
        self.assertEqual(report["status"], "TRANSIENT_FAILURE")
        self.assertNotIn("secret", json.dumps(report))

    def test_access_denied_not_retried(self):
        report, calls, _ = self.probe([(403, None)])
        self.assertEqual(report["status"], "ACCESS_RESTRICTED")
        self.assertEqual(len(calls), 1)

    def test_redirect_checked_and_loop_bounded(self):
        calls = []
        def redirect(url, timeout):
            calls.append(url)
            return 302, "https://127.0.0.1/private"
        self.assertEqual(links.probe("https://example.org", self.policy["network"], redirect)["status"], "POLICY_BLOCKED")
        self.assertEqual(len(calls), 1)
        report = links.probe("https://example.org/page", self.policy["network"], lambda *args: (302, "/page"))
        self.assertEqual(report["status"], "REDIRECT_FAILURE")

    def test_redirect_queries_not_transmitted(self):
        calls = []
        def transport(url, timeout):
            calls.append(url)
            return 302, "/login?token=secret"
        report = links.probe("https://example.org", self.policy["network"], transport)
        self.assertEqual(report["status"], "ACCESS_RESTRICTED")
        self.assertEqual(len(calls), 1)
        self.assertNotIn("secret", json.dumps(report))

    def test_query_credentials_scheme_domain_not_requested(self):
        for url in ("https://example.org/?token=secret", "https://user:secret@example.org/",
                    "http://example.org", "https://other.example/page", "ftp://example.org"):
            with self.subTest(url=url), patch.object(links, "request") as transport:
                report = links.probe(url, self.policy["network"], transport)
                transport.assert_not_called()
                self.assertNotEqual(report["status"], "OK")

    def test_private_dns_never_connects(self):
        with patch.object(links.socket, "getaddrinfo", return_value=[(2, 1, 6, "", ("127.0.0.1", 443))]), \
                patch.object(links, "PinnedHTTPS") as connection:
            with self.assertRaises(ValueError):
                links.request("https://example.org", 2)
            connection.assert_not_called()

    def test_tls_connect_uses_checked_ip(self):
        with patch.object(links.socket, "create_connection") as connect:
            client = links.PinnedHTTPS("example.org", "93.184.215.14", 2)
            with patch.object(client.tls_context, "wrap_socket") as wrap:
                client.connect()
                connect.assert_called_once_with(("93.184.215.14", 443), 2)
                wrap.assert_called_once_with(connect.return_value, server_hostname="example.org")

    def test_exception_expiry_reason_and_limits(self):
        network = self.policy["network"]
        network["exceptions"] = [{"id": "a" * 64, "reason": "Temporary service outage", "expires": "2026-09-10"}]
        links.validate_policy(network, TODAY)
        with self.assertRaises(ValueError):
            links.validate_policy(network, date(2026, 9, 11))
        network["exceptions"] = []
        network["attempts"] = 100
        with self.assertRaises(ValueError):
            links.validate_policy(network, TODAY)

    def test_classification_expiry_kind_and_identity_fail_closed(self):
        network = self.policy["network"]
        network["classifications"] = [{
            "id": "b" * 64,
            "kind": "restricted-historical",
            "reason": "Reviewed private evidence",
            "expires": "2026-09-08",
        }]
        with self.assertRaises(ValueError):
            links.validate_policy(network, TODAY)
        network["classifications"][0]["expires"] = "2026-09-10"
        network["classifications"][0]["kind"] = "private-maybe"
        with self.assertRaises(ValueError):
            links.validate_policy(network, TODAY)
        network["classifications"][0]["kind"] = "restricted-historical"
        network["exceptions"] = [{
            "id": "b" * 64,
            "reason": "Conflicting exception",
            "expires": "2026-09-10",
        }]
        with self.assertRaises(ValueError):
            links.validate_policy(network, TODAY)

    def test_reports_redact_urls_and_deduplicate(self):
        url = "https://example.org/private?token=TOP_SECRET"
        (self.root / "README.md").write_text(f"[one]({url})\n[two]({url})\n")
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("must not request"))
        self.assertEqual(report["discovered"], 1)
        self.assertEqual(len(report["links"][0]["locations"]), 2)
        rendered = json.dumps(report) + links.markdown(report)
        self.assertNotIn("TOP_SECRET", rendered)
        self.assertNotIn("/private", rendered)

    def test_active_exception_is_explicit(self):
        url = "https://example.org/page"
        (self.root / "README.md").write_text(f"[page]({url})")
        self.policy["network"]["exceptions"] = [{"id": hashlib.sha256(url.encode()).hexdigest(), "reason": "Reviewed temporary gap", "expires": "2026-09-10"}]
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("excepted"))
        self.assertEqual(report["links"][0]["status"], "EXCEPTED")

    def test_total_limit_cannot_claim_complete(self):
        (self.root / "README.md").write_text("[a](https://example.org/a)\n[b](https://example.org/b)")
        self.policy["network"]["max_links"] = 1
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: (200, None))
        self.assertTrue(report["limit_exceeded"])
        self.assertEqual(report["status"], "fail")

    def test_worker_concurrency_bounded(self):
        (self.root / "README.md").write_text("\n".join(f"[page](https://example.org/{i})" for i in range(12)))
        self.policy["network"]["concurrency"] = 2
        self.save()
        active = peak = 0
        lock = threading.Lock()
        def transport(*args):
            nonlocal active, peak
            with lock:
                active += 1
                peak = max(peak, active)
            time.sleep(0.01)
            with lock:
                active -= 1
            return 200, None
        self.assertEqual(links.build_report(self.root, TODAY, transport)["status"], "pass")
        self.assertLessEqual(peak, 2)


if __name__ == "__main__":
    unittest.main()
