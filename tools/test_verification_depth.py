#!/usr/bin/env python3
"""Focused behavioral tests for release-critical Python verification depth."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import _release_closeout as closeout
import check_release as release
import check_template_contract as contract
import check_wiki as wiki
import initialize_repository as initializer


class InitializerDepthTests(unittest.TestCase):
    def test_load_config_failure_shapes(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            with self.assertRaises(initializer.InitializationError):
                initializer._load_config(root)
            (root / ".github").mkdir()
            config = root / initializer.CONFIG_PATH

            config.write_text("{", encoding="utf-8")
            with self.assertRaises(initializer.InitializationError):
                initializer._load_config(root)

            config.write_text("[]", encoding="utf-8")
            with self.assertRaises(initializer.InitializationError):
                initializer._load_config(root)

            for document in (
                {},
                {"placeholders": []},
                {"placeholders": {"pattern": "(", "catalogue": {}, "block_markers": {},
                                  "template_only_paths": [], "exclude_paths": []}},
                {"placeholders": {"pattern": "(A)(B)", "catalogue": {}, "block_markers": {},
                                  "template_only_paths": [], "exclude_paths": []}},
                {"placeholders": {"pattern": "(A)", "catalogue": [], "block_markers": {},
                                  "template_only_paths": [], "exclude_paths": []}},
                {"placeholders": {"pattern": "(A)", "catalogue": {}, "block_markers": {},
                                  "template_only_paths": "x", "exclude_paths": []}},
            ):
                config.write_text(json.dumps(document), encoding="utf-8")
                with self.assertRaises(initializer.InitializationError):
                    initializer._load_config(root)

    def test_parse_assignments_rejects_bad_values(self) -> None:
        self.assertEqual(
            initializer._parse_assignments(["NAME=value", "NAME=second"], "--set"),
            {"NAME": ["value", "second"]},
        )
        bad = ("NOVALUE", "bad-name=x", "NAME=", "NAME=a\nb", "NAME={{TOKEN}}")
        for entry in bad:
            with self.subTest(entry=entry):
                with self.assertRaises(initializer.InitializationError):
                    initializer._parse_assignments([entry], "--set")

    def test_validate_values_rejects_schema_and_value_errors(self) -> None:
        catalogue = {
            "REQ": {"category": "required"},
            "OPT": {"category": "optional"},
            "REP": {"category": "repeatable"},
            "PROF": {"category": "profile-specific"},
            "REPOSITORY_PATH": {"category": "required"},
            "SUPPORT_CONTACT": {"category": "required"},
            "COPYRIGHT_YEAR": {"category": "required"},
            "MAINTAINER_NAME": {"category": "required"},
            "PROJECT_NAME": {"category": "required"},
            "PROJECT_TAGLINE": {"category": "required"},
            "PROJECT_DESCRIPTION": {"category": "required"},
            "SOCIAL_PREVIEW_PATH": {"category": "optional"},
        }
        base = [
            "REQ=x",
            "REPOSITORY_PATH=owner/repo",
            "SUPPORT_CONTACT=security@example.invalid",
            "COPYRIGHT_YEAR=2026",
            "MAINTAINER_NAME=M",
            "PROJECT_NAME=P",
            "PROJECT_TAGLINE=T",
            "PROJECT_DESCRIPTION=D",
        ]
        root = Path(".")
        with self.assertRaisesRegex(initializer.InitializationError, "Unknown substitutions"):
            initializer._validate_values(root, set(), catalogue, base + ["UNKNOWN=x"], [])
        with self.assertRaisesRegex(initializer.InitializationError, "more than once"):
            initializer._validate_values(root, set(), catalogue, base + ["REQ=y"], [])
        with self.assertRaisesRegex(initializer.InitializationError, "cannot be supplied with --set"):
            initializer._validate_values(root, set(), catalogue, base + ["REP=x"], [])
        with self.assertRaisesRegex(initializer.InitializationError, "cannot be supplied with --add"):
            initializer._validate_values(root, set(), catalogue, base, ["OPT=x"])
        with self.assertRaisesRegex(initializer.InitializationError, "Missing required"):
            initializer._validate_values(root, set(), catalogue, base[1:], [])
        for replacement, expected in (
            ("REPOSITORY_PATH=bad", "owner/name"),
            ("SUPPORT_CONTACT=bad", "email address or HTTPS"),
            ("COPYRIGHT_YEAR=1999", "four-digit year"),
        ):
            values = [
                replacement if item.split("=")[0] == replacement.split("=")[0] else item
                for item in base
            ]
            with self.assertRaisesRegex(initializer.InitializationError, expected):
                initializer._validate_values(root, set(), catalogue, values, [])
        values = [item for item in base if not item.startswith("PROJECT_NAME=")]
        values.append("PROJECT_NAME=" + "x" * 101)
        with self.assertRaisesRegex(initializer.InitializationError, "100-character"):
            initializer._validate_values(root, set(), catalogue, values, [])
        with self.assertRaisesRegex(initializer.InitializationError, "tracked repository-relative"):
            initializer._validate_values(
                root, set(), catalogue, base + ["SOCIAL_PREVIEW_PATH=assets/x.png"], []
            )
        with self.assertRaisesRegex(initializer.InitializationError, "repository-relative"):
            initializer._validate_values(
                root, set(), catalogue, base + ["SOCIAL_PREVIEW_PATH=../x.png"], [],
                require_preview_file=False,
            )

    def test_render_blocks_behaviors_and_failures(self) -> None:
        catalogue = {
            "OPT": {"category": "optional"},
            "REP": {"category": "repeatable"},
            "BAD": {"category": "required"},
        }
        text = (
            "a\n"
            "<!-- template:remove:start -->\nremove\n<!-- template:remove:end -->\n"
            "<!-- template:profile:library:start -->\nlib\n<!-- template:profile:library:end -->\n"
            "<!-- template:optional:OPT:start -->\nopt\n<!-- template:optional:OPT:end -->\n"
            "<!-- template:repeatable:REP:start -->\nrep\n<!-- template:repeatable:REP:end -->\n"
        )
        rendered = initializer._render_blocks(
            "README.md", text, "library", {"OPT": "yes"}, {"REP": ["x"]}, catalogue
        )
        self.assertEqual(rendered, "a\nlib\nopt\nrep\n")
        cases = (
            ("bad <!-- template:oops -->\n", "invalid template block marker"),
            ("<!-- template:remove:start -->\n<!-- template:remove:start -->\n", "may not nest"),
            ("<!-- template:optional:BAD:start -->\n", "is not optional"),
            ("<!-- template:repeatable:BAD:start -->\n", "is not repeatable"),
            ("<!-- template:remove:end -->\n", "unmatched template block end"),
            ("<!-- template:remove:start -->\n", "unclosed template block"),
        )
        for text, message in cases:
            with self.subTest(message=message):
                with self.assertRaisesRegex(initializer.InitializationError, message):
                    initializer._render_blocks("README.md", text, "library", {}, {}, catalogue)

    def test_replacement_badges_changelog_and_helpers(self) -> None:
        catalogue = {
            "PROFILE": {"category": "profile-specific", "values": {"library": "lib"}},
            "REP": {"category": "repeatable", "item_format": "- {value}"},
        }
        self.assertEqual(
            initializer._replacement_values("library", {"A": "x"}, {"REP": ["one", "two"]}, catalogue),
            {"A": "x", "PROFILE": "lib", "REP": "- one\n- two"},
        )
        readme = (
            "https://github.com/owner/template/actions "
            "https://img.shields.io/github/v/release/owner/template?x "
            "https://img.shields.io/github/issues/owner/template?x"
        )
        retargeted = initializer._render_readme_badges(
            "README.md", readme, "owner/template", "owner/product"
        )
        self.assertNotIn("owner/template", retargeted)
        self.assertEqual(
            initializer._render_readme_badges("OTHER.md", readme, "owner/template", "owner/product"),
            readme,
        )
        changelog = "# C\n\n## [Unreleased]\n\nold\n\n---\nrest\n"
        reset = initializer._reset_changelog(changelog)
        self.assertIn("No unreleased changes recorded.", reset)
        for bad in ("# C\n", "## [Unreleased]\nno boundary"):
            with self.assertRaises(initializer.InitializationError):
                initializer._reset_changelog(bad)
        self.assertIsNone(initializer._sha256(None))
        self.assertEqual(len(initializer._sha256(b"x") or ""), 64)

    def test_already_initialized_and_placeholder_protection(self) -> None:
        config = {"mode": "template"}
        self.assertFalse(initializer._already_initialized(Path("."), config, "library", {}, {}))
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            generated = {
                "mode": "generated",
                "profile": "library",
                "repository": "owner/repo",
                "template_contract": {"version": "1.2.0", "source": "owner/template"},
            }
            scalars = {"REPOSITORY_PATH": "owner/repo"}
            with self.assertRaisesRegex(initializer.InitializationError, "missing"):
                initializer._already_initialized(root, generated, "library", scalars, {})
            (root / ".github").mkdir()
            (root / initializer.RECORD_PATH).write_bytes(
                initializer._record("library", scalars, {}, generated["template_contract"])
            )
            self.assertTrue(initializer._already_initialized(root, generated, "library", scalars, {}))
            with self.assertRaisesRegex(initializer.InitializationError, "different profile"):
                initializer._already_initialized(root, generated, "application", scalars, {})
            (root / initializer.RECORD_PATH).write_bytes(b"{}")
            with self.assertRaisesRegex(initializer.InitializationError, "different substitution"):
                initializer._already_initialized(root, generated, "library", scalars, {})
        with self.assertRaises(initializer.InitializationError):
            initializer._reject_executable_placeholders("tool.py", [object()])
        initializer._reject_executable_placeholders("README.md", [object()])

    def test_strip_workflow_plan_and_apply_rollback(self) -> None:
        workflow = (
            "before\n"
            "      - name: Exercise policy-branch coverage determinism\n"
            "        run: hidden\n"
            "      - name: Exercise positive and degraded checker paths\n"
            "keep\n"
            "test-results/policy-coverage.json\n"
            "POLICY_COVERAGE_OUTCOME=x\n"
            "\"Policy coverage:$X\"\n"
        )
        stripped = initializer._strip_template_maintenance_workflow(
            ".github/workflows/static-checks.yml", workflow
        )
        self.assertEqual(stripped, "before\n      - name: Exercise positive and degraded checker paths\nkeep\n")
        self.assertEqual(initializer._strip_template_maintenance_workflow("x.yml", workflow), workflow)
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / "old.txt").write_bytes(b"old")
            plan = initializer._plan(
                root, "library", {"old.txt": b"new", "new.txt": b"x", "gone.txt": None}
            )
            self.assertEqual([row["action"] for row in plan["changes"]], ["update", "create", "delete"])
            initializer._apply_changes(root, {"old.txt": b"new", "new.txt": b"x"})
            self.assertEqual((root / "old.txt").read_bytes(), b"new")
            self.assertEqual((root / "new.txt").read_bytes(), b"x")
            before = (root / "old.txt").read_bytes()
            with patch.object(initializer.os, "replace", side_effect=OSError("boom")):
                with self.assertRaisesRegex(initializer.InitializationError, "original files were restored"):
                    initializer._apply_changes(root, {"old.txt": b"broken"})
            self.assertEqual((root / "old.txt").read_bytes(), before)


class ReleaseDepthTests(unittest.TestCase):
    def base_policy(self) -> dict:
        profile = {"required_checks": ["vba-compile", "regression"], "allowed_asset_globs": []}
        return {
            "schema_version": 1,
            "evidence_schema_version": 1,
            "provenance_signature_mode": "none",
            "core_checks": ["repository-integrity"],
            "profiles": {
                "application": dict(profile),
                "library": dict(profile),
                "template": dict(profile),
                "ui-component": dict(profile),
            },
            "source_scan_exclude_paths": ["README.md"],
            "template_construction_markers": ["template construction"],
        }

    def test_json_configuration_and_policy_failures(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / ".github").mkdir()
            config_path = root / release.PROFILE_PATH
            value, findings = release._read_json(config_path, "repository-profile")
            self.assertIsNone(value)
            self.assertTrue(findings)
            config_path.write_text("{", encoding="utf-8")
            self.assertTrue(release._read_json(config_path, "repository-profile")[1])
            config_path.write_text("[]", encoding="utf-8")
            self.assertIsNone(release._load_configuration(root)[0])
            config_path.write_text("{}", encoding="utf-8")
            self.assertEqual(release._load_configuration(root), ({}, []))

            policy_path = root / release.POLICY_PATH
            good = self.base_policy()
            policy_path.write_text(json.dumps(good), encoding="utf-8")
            self.assertIsNotNone(release._load_policy(root)[0])
            variants = []
            value = dict(good); value.pop("schema_version"); variants.append(value)
            value = dict(good); value["schema_version"] = 2; variants.append(value)
            value = dict(good); value["provenance_signature_mode"] = "bad"; variants.append(value)
            value = dict(good); value["core_checks"] = []; variants.append(value)
            value = dict(good); value["core_checks"] = ["x", "x"]; variants.append(value)
            value = dict(good); value["profiles"] = {}; variants.append(value)
            value = json.loads(json.dumps(good)); value["profiles"]["library"]["extra"] = 1; variants.append(value)
            value = json.loads(json.dumps(good)); value["profiles"]["library"]["required_checks"] = []; variants.append(value)
            value = json.loads(json.dumps(good)); value["profiles"]["library"]["required_checks"] = ["repository-integrity"]; variants.append(value)
            value = json.loads(json.dumps(good)); value["profiles"]["library"]["allowed_asset_globs"] = ["../bad"]; variants.append(value)
            value = dict(good); value["source_scan_exclude_paths"] = ["../bad"]; variants.append(value)
            value = dict(good); value["template_construction_markers"] = []; variants.append(value)
            for variant in variants:
                with self.subTest(variant=variant):
                    policy_path.write_text(json.dumps(variant), encoding="utf-8")
                    self.assertIsNone(release._load_policy(root)[0])

    def test_safe_relative_and_git_operational_paths(self) -> None:
        for value, expected in (
            ("a/b", True), ("", False), ("/a", False), ("../a", False),
            ("a\\b", False), ("a\0b", False), (1, False),
        ):
            self.assertEqual(release._safe_relative(value), expected)
        with patch.object(release.subprocess, "run", side_effect=FileNotFoundError()):
            with self.assertRaises(release.OperationalError):
                release._git(Path("."), "status")
        with patch.object(release, "_git_output", return_value=None):
            with self.assertRaises(release.OperationalError):
                release._tracked_files(Path("."))

    def test_resolve_release_profile_modes(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / ".github").mkdir()
            generated = {"mode": "generated", "profile": "library", "repository": "owner/product"}
            profile, findings = release._resolve_release_profile(root, generated)
            self.assertEqual(profile, "library")
            self.assertTrue(findings)

            record = {"profile": "application", "values": {"REPOSITORY_PATH": "owner/other"}}
            (root / release.INITIALIZATION_PATH).write_text(json.dumps(record), encoding="utf-8")
            _, findings = release._resolve_release_profile(root, generated)
            self.assertGreaterEqual(len(findings), 2)

            template = {
                "mode": "template", "profile": None, "repository": "owner/product",
                "identity": {"template_tokens": ["TEMPLATE"]},
            }
            _, findings = release._resolve_release_profile(root, template)
            self.assertGreaterEqual(len(findings), 2)

            profile, findings = release._resolve_release_profile(root, {"mode": "other"})
            self.assertIsNone(profile)
            self.assertTrue(findings)

    def test_candidate_git_state_variants(self) -> None:
        with patch.object(release, "_git_output", side_effect=["a" * 40, " M x"]):
            findings = release._validate_candidate_git_state(Path("."), "b" * 40, "v1.0.0", False)
            self.assertEqual({row["code"] for row in findings}, {"candidate-sha-mismatch", "dirty-candidate"})
        with patch.object(release, "_git_output", side_effect=["a" * 40, "", None, None]):
            findings = release._validate_candidate_git_state(Path("."), "a" * 40, "v1.0.0", True)
            self.assertEqual(findings[0]["code"], "missing-tag-ref")
        with patch.object(release, "_git_output", side_effect=["a" * 40, "", "commit", "b" * 40]):
            codes = {row["code"] for row in release._validate_candidate_git_state(Path("."), "a" * 40, "v1.0.0", True)}
            self.assertEqual(codes, {"lightweight-tag", "tag-target-mismatch"})

    def test_version_changelog_failures(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            version, findings = release._validate_version_and_changelog(root, "v1.0.0")
            self.assertIsNone(version)
            self.assertTrue(findings)
            (root / "VERSION").write_text("0.0.0\n", encoding="utf-8")
            (root / "CHANGELOG.md").write_text("## [0.0.0] - 2026-02-30\n", encoding="utf-8")
            _, findings = release._validate_version_and_changelog(root, "v9.9.9")
            codes = {row["code"] for row in findings}
            self.assertTrue({"zero-version", "tag-version-mismatch", "invalid-changelog-date"} <= codes)
            (root / "VERSION").write_text("1.0.0 extra\n", encoding="utf-8")
            (root / "CHANGELOG.md").write_text("# no release\n", encoding="utf-8")
            _, findings = release._validate_version_and_changelog(root, "v1.0.0")
            codes = {row["code"] for row in findings}
            self.assertTrue({"invalid-version", "tag-version-mismatch", "missing-changelog-release"} <= codes)

    def test_validate_check_field_failures(self) -> None:
        sha = "a" * 40
        self.assertEqual(release._validate_check("x", None, sha)[0]["code"], "invalid-evidence-check")
        self.assertTrue(release._validate_check("repository-integrity", {}, sha))
        base = {"status": "FAIL", "candidate_sha": "b" * 40, "detail": ""}
        findings = release._validate_check("x", base, sha)
        self.assertEqual({row["code"] for row in findings}, {"failed-evidence-check", "evidence-sha-mismatch", "invalid-evidence-check"})
        repo = {**base, "status": "PASS", "candidate_sha": sha, "detail": "x", "run_url": "http://x"}
        self.assertTrue(release._validate_check("repository-integrity", repo, sha))
        compile_check = {**base, "status": "PASS", "candidate_sha": sha, "detail": "x", "environment": ""}
        self.assertTrue(release._validate_check("vba-compile", compile_check, sha))
        regression = {
            "status": "PASS", "candidate_sha": sha, "detail": "x", "entry_point": "",
            "environment": "", "cases": True, "assertions": 0, "failures": 1,
            "completeness": "PARTIAL", "cleanup": "FAIL",
        }
        self.assertGreaterEqual(len(release._validate_check("regression", regression, sha)), 7)

    def test_manifest_paths_and_asset_validation(self) -> None:
        import hashlib
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            self.assertTrue(release._parse_manifest(root / "missing.txt")[1])
            manifest = root / "manifest.txt"
            manifest.write_text(
                "bad\n" + "a" * 64 + "  z.bin\n" + "b" * 64 + "  a.bin\n" + "c" * 64 + "  a.bin\n",
                encoding="utf-8",
            )
            _, findings = release._parse_manifest(manifest)
            self.assertGreaterEqual(len(findings), 3)
            asset_file = root / "dist.bin"
            asset_file.write_bytes(b"x")
            sha = "a" * 40
            self.assertTrue(release._validate_asset_record(root, {}, "asset", ["*.bin"], sha, {}))
            malformed = {"path": "../x", "sha256": "x", "candidate_sha": sha, "package_test": "PASS"}
            self.assertTrue(release._validate_asset_record(root, malformed, "asset", ["*.bin"], sha, {}))
            digest = hashlib.sha256(b"wrong").hexdigest()
            record = {"path": "dist.bin", "sha256": digest, "candidate_sha": "b" * 40, "package_test": "FAIL"}
            findings = release._validate_asset_record(root, record, "asset", ["*.zip"], sha, {})
            codes = {row["code"] for row in findings}
            self.assertTrue({"asset-sha-binding-mismatch", "failed-package-test", "unapproved-binary", "asset-digest-mismatch"} <= codes)

    def test_evidence_metadata_checks_and_binary_policy(self) -> None:
        policy = self.base_policy()
        sha = "a" * 40
        evidence = {
            "schema_version": 9, "version": "9.9.9", "tag": "bad", "candidate_sha": "b" * 40,
            "profile": "other", "distribution": "bad", "checks": {}, "assets": [], "extra": 1,
        }
        findings = release._validate_evidence_metadata(evidence, Path("e.json"), policy, "library", "1.0.0", "v1.0.0", sha)
        self.assertGreaterEqual(len(findings), 6)
        self.assertTrue(release._validate_evidence_checks({"checks": []}, policy, "library", sha))
        self.assertTrue(release._validate_evidence_checks({"checks": {"BAD ID": {}}}, policy, "library", sha))
        findings = release._validate_binary_assets(Path("."), [], None, None, policy, "library", sha)
        codes = {row["code"] for row in findings}
        self.assertTrue({"unapproved-binary", "missing-release-assets", "missing-asset-manifest"} <= codes)

    def test_reports_and_atomic_write(self) -> None:
        report = {
            "status": "fail", "tag": "v1", "candidate_sha": "a" * 40, "profile": "library",
            "counts": {"findings": 1},
            "findings": [{"code": "x", "path": "a", "message": "pipe | here"}],
            "scope_note": "scope",
        }
        self.assertIn("[FAIL]", release.console_report(report))
        self.assertIn("\\|", release.markdown_report(report))
        with tempfile.TemporaryDirectory() as name:
            target = Path(name) / "nested" / "report.txt"
            release._write_atomic(target, "hello")
            self.assertEqual(target.read_text(encoding="utf-8"), "hello")


class CloseoutDepthTests(unittest.TestCase):
    def test_type_helpers_and_json_errors(self) -> None:
        with self.assertRaises(closeout.CloseoutError):
            closeout.require(False, "x")
        for func, value in ((closeout.as_object, []), (closeout.as_list, {}), (closeout.as_string, "")):
            with self.subTest(func=func.__name__):
                with self.assertRaises(closeout.CloseoutError):
                    func(value, "x")
        self.assertEqual(closeout.unique_object([("a", 1)]), {"a": 1})
        with self.assertRaises(closeout.CloseoutError):
            closeout.unique_object([("a", 1), ("a", 2)])
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            with self.assertRaises(closeout.CloseoutError):
                closeout.load_json(root / "missing.json")
            path = root / "x.json"
            path.write_text('{"a":1,"a":2}', encoding="utf-8")
            with self.assertRaises(closeout.CloseoutError):
                closeout.load_json(path)
            path.write_text("{", encoding="utf-8")
            with self.assertRaises(closeout.CloseoutError):
                closeout.load_json(path)

    def test_workflow_runs_and_milestone_items(self) -> None:
        rows, total = closeout.workflow_runs([
            {"workflow_runs": [{"id": 1}], "total_count": 2},
            {"workflow_runs": [{"id": 2}], "total_count": 3},
        ])
        self.assertEqual((len(rows), total), (2, 3))
        self.assertEqual(closeout.milestone_items([[{"number": 1}], [{"number": 2}], ["bad"]]), [{"number": 1}, {"number": 2}])
        self.assertEqual(closeout.milestone_items([{"number": 1}, "bad"]), [{"number": 1}])

    def test_tag_variants(self) -> None:
        candidate = "a" * 40
        findings: list[dict[str, str]] = []
        state = closeout.check_tag({"tag_ref": {"ref": "bad", "object": {"type": "commit", "sha": "b" * 40}}}, "v1.0.0", candidate, findings)
        self.assertEqual(state["ref_type"], "commit")
        self.assertGreaterEqual(len(findings), 3)
        findings = []
        closeout.check_tag({
            "tag_ref": {"ref": "refs/tags/v1.0.0", "object": {"type": "tag", "sha": "t"}},
            "tag_object": {"sha": "wrong", "tag": "v2", "object": {"type": "blob", "sha": candidate}},
        }, "v1.0.0", candidate, findings)
        self.assertGreaterEqual(len(findings), 3)
        findings = []
        closeout.check_tag({"tag_ref": {"ref": "refs/tags/v1.0.0", "object": {"type": "tree"}}}, "v1.0.0", candidate, findings)
        self.assertTrue(findings)

    def test_workflow_selection_and_failure(self) -> None:
        candidate = "a" * 40
        findings: list[dict[str, str]] = []
        result = closeout.check_workflow({"workflow_runs": {"workflow_runs": [], "total_count": 1}}, "x.yml", "v1", candidate, findings)
        self.assertEqual(result["matches"], 0)
        self.assertGreaterEqual(len(findings), 2)
        rows = [
            {"id": 1, "path": "x.yml", "head_branch": "v1", "head_sha": candidate, "event": "push", "run_attempt": 1, "run_number": 1, "status": "completed", "conclusion": "failure"},
            {"id": 2, "path": "x.yml", "head_branch": "v1", "head_sha": candidate, "event": "push", "run_attempt": 2, "run_number": 2, "status": "completed", "conclusion": "success"},
        ]
        findings = []
        result = closeout.check_workflow({"workflow_runs": {"workflow_runs": rows, "total_count": 2}}, "x.yml", "v1", candidate, findings)
        self.assertEqual(result["run_id"], 2)
        self.assertEqual(findings, [])

    def test_release_asset_archive_and_latest_failures(self) -> None:
        identity = {"version": "1.0.0", "allowed_asset_globs": []}
        snapshot = {
            "release": {"id": 1, "tag_name": "bad", "draft": True, "published_at": None, "prerelease": True,
                        "assets": [{"name": "x.bin"}, {"name": "x.bin"}, "bad"], "zipball_url": None, "tarball_url": ""},
            "latest_release": {"id": 2},
            "source_archives": {"zip": "fail", "tar": "missing"},
        }
        findings: list[dict[str, str]] = []
        result = closeout.check_release(snapshot, identity, False, True, findings)
        self.assertEqual(result["asset_mode"], "source-only")
        self.assertTrue({"release", "assets", "source-archives"} <= {row["control"] for row in findings})

    def test_compare_milestone_and_wiki_failures(self) -> None:
        findings: list[dict[str, str]] = []
        result = closeout.check_compare({"compare": {}}, {"version": "1.0.0", "previous_tag": None}, "a" * 40, findings)
        self.assertIsNone(result["previous_tag"])
        self.assertEqual(findings, [])
        findings = []
        closeout.check_compare({"compare": {"html_url": "bad", "status": "behind", "commits": []}}, {"version": "1.0.0", "previous_tag": "v0.9.0"}, "a" * 40, findings)
        self.assertGreaterEqual(len(findings), 2)
        findings = []
        result = closeout.check_milestone({
            "milestone": {"number": 9, "state": "open", "open_issues": 0, "closed_issues": 0, "title": "m"},
            "milestone_items": [
                {"number": 1, "state": "open", "milestone": {"number": 8}},
                {"number": 2, "state": "mystery", "milestone": {"number": 9}},
            ],
        }, 4, findings)
        self.assertEqual(result["open_items"], [1])
        self.assertGreaterEqual(len(findings), 5)
        self.assertFalse(closeout.check_wiki({}, {"mode": "generated"}, "a" * 40, [])["applicable"]
        findings = []
        result = closeout.check_wiki({
            "wiki": {"status": "fail", "source_sha": "b" * 40},
            "ui_observations": {"wiki_browser_review": "fail"},
        }, {"mode": "template"}, "a" * 40, findings)
        self.assertTrue(result["applicable"])
        self.assertEqual(len(findings), 3)


class ContractAndWikiDepthTests(unittest.TestCase):
    def test_contract_read_shape_notes_and_record_errors(self) -> None:
        findings: list[dict[str, str]] = []
        self.assertEqual(contract._check_shape([], findings), ("", ""))
        self.assertTrue(findings)
        findings = []
        version, source = contract._check_shape({"version": "1.2.0-rc.1", "source": "bad", "extra": 1}, findings)
        self.assertEqual((version, source), ("", ""))
        self.assertGreaterEqual(len(findings), 3)
        findings = []
        self.assertEqual(contract._check_supported("", findings), frozenset())
        self.assertEqual(contract._check_supported("9.9.9", findings), frozenset())
        self.assertTrue(findings)
        findings = []
        contract._check_mode({"mode": "template", "repository": "owner/a"}, "owner/b", findings)
        contract._check_mode({"mode": "generated", "repository": "owner/a"}, "owner/a", findings)
        self.assertEqual(len(findings), 2)
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            with self.assertRaises(contract.ContractError):
                contract._read_json(root, "missing.json")
            (root / "docs").mkdir()
            (root / contract.NOTES_PATH).write_text("### 1.0.0\n", encoding="utf-8")
            findings = []
            contract._check_notes(root, findings)
            self.assertGreaterEqual(len(findings), 2)
            (root / ".github").mkdir()
            (root / contract.RECORD_PATH).write_text("[]", encoding="utf-8")
            findings = []
            self.assertIsNone(contract._check_record(root, {"mode": "generated"}, "1.2.0", "owner/template", findings))
            self.assertTrue(findings)

    def test_wiki_helpers_failure_boundaries(self) -> None:
        self.assertEqual(wiki.directories({"a/b/c.txt"}), {"a", "a/b"})
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            path = root / "x.json"
            path.write_text('{"a":1,"a":2}', encoding="utf-8")
            with self.assertRaises(ValueError):
                wiki.read_json(path)
            path.write_text("[]", encoding="utf-8")
            with self.assertRaises(ValueError):
                wiki.read_json(path)
            with self.assertRaises(ValueError):
                wiki.bundle(root, {"pages": []}, "bad")
            with self.assertRaises(ValueError):
                wiki.compare_bundle({}, root / "missing")
            outside = Path(name).parent / (Path(name).name + "-export")
            try:
                wiki.write_bundle({"x": b"1"}, outside, root)
                self.assertEqual((outside / "x").read_bytes(), b"1")
                with self.assertRaises(ValueError):
                    wiki.write_bundle({}, outside, root)
            finally:
                if outside.exists():
                    for item in outside.iterdir():
                        item.unlink()
                    outside.rmdir()

    def test_wiki_render_and_compare(self) -> None:
        text = "[remote](https://example.com) [anchor](#x) [page](Other.md#y) [file](../../README.md)"
        rendered = wiki.render_page(text, "Home.md", {"Home", "Other"}, "owner/repo", "a" * 40)
        self.assertIn("](Other#y)", rendered)
        self.assertIn("/blob/" + "a" * 40, rendered)
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / "x").write_bytes(b"bad")
            differences = wiki.compare_bundle({"x": b"good", "y": b"new"}, root)
            self.assertEqual(len(differences), 2)


if __name__ == "__main__":
    unittest.main()
