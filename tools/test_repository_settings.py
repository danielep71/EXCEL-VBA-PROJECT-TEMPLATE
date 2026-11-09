"""Read-only settings check against a simulated GitHub API; no live request or credential."""
from __future__ import annotations

import contextlib
import copy
import hashlib
import io
import json
import tempfile
import unittest
from datetime import datetime, timezone
from email.message import Message
from pathlib import Path
from typing import Any
from unittest.mock import patch
from urllib.error import HTTPError

import check_repository_settings as settings

ROOT = Path(__file__).resolve().parents[1]
REPOSITORY = "example/generated"
NOW = datetime(2026, 11, 9, 12, 0, tzinfo=timezone.utc)


class FakeSettings:
    """A compliant public repository; tests change one setting at a time."""

    def __init__(self, policy: dict[str, Any], labels: list[dict[str, Any]]) -> None:
        self.paths: list[str] = []
        self.unreadable: set[str] = set()
        self.metadata: dict[str, Any] = {
            **policy["features"], "full_name": REPOSITORY, "default_branch": "main",
            "description": "Example generated VBA project", "visibility": "public", "private": False,
            "security_and_analysis": {"secret_scanning": {"status": "enabled"},
                                      "secret_scanning_push_protection": {"status": "enabled"}}}
        self.topics = ["excel", "library", "vba"]
        self.labels = copy.deepcopy(labels)
        self.rulesets = [{"id": index + 1, **rule} for index, rule in enumerate(settings.baseline_rules(policy))]
        self.security: dict[str, Any] = {"private_vulnerability_reporting": True, "vulnerability_alerts": True,
                                         "automated_security_fixes": True, "paused": False}
        self.actions: dict[str, dict[str, Any]] = {
            "/actions/permissions": {"enabled": True, "allowed_actions": "all", "sha_pinning_required": True},
            "/actions/permissions/workflow": {"default_workflow_permissions": "read",
                                              "can_approve_pull_request_reviews": False},
            "/actions/permissions/fork-pr-contributor-approval": {"approval_policy": "all_external_contributors"},
            "/actions/permissions/artifact-and-log-retention": {"days": 90, "maximum_allowed_days": 90}}

    def get(self, path: str) -> tuple[int, Any]:
        self.paths.append(path)
        base = path.split("?", 1)[0]
        if base in self.unreadable:
            return 403, None
        if path == "":
            return 200, copy.deepcopy(self.metadata)
        if path == "/topics":
            return 200, {"names": list(self.topics)}
        if base == "/labels":
            return 200, copy.deepcopy(self.labels) if path.endswith("page=1") else []
        if base == "/rulesets":
            return 200, [{"id": row["id"], "name": row["name"]} for row in self.rulesets] if path.endswith("page=1") else []
        if base.startswith("/rulesets/"):
            return 200, copy.deepcopy(next(row for row in self.rulesets if str(row["id"]) == base.rsplit("/", 1)[1]))
        if path == "/private-vulnerability-reporting":
            return 200, {"enabled": self.security["private_vulnerability_reporting"]}
        if path == "/vulnerability-alerts":
            return (204 if self.security["vulnerability_alerts"] else 404), None
        if path == "/automated-security-fixes":
            return 200, {"enabled": self.security["automated_security_fixes"], "paused": self.security["paused"]}
        if path in self.actions:
            return 200, copy.deepcopy(self.actions[path])
        raise AssertionError(path)


class RepositorySettingsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        (self.root / ".github").mkdir()
        self.policy = json.loads((ROOT / settings.POLICY).read_text(encoding="utf-8"))
        manifest = json.loads((ROOT / settings.LABELS).read_text(encoding="utf-8"))
        self.config = {"mode": "generated", "profile": "library", "repository": REPOSITORY, "label_domains": []}
        self.write(settings.CONFIG, self.config)
        self.write(settings.LABELS, manifest)
        self.write(settings.RECORD, {"values": {"PROJECT_DESCRIPTION": "Example generated VBA project"}})
        self.write(settings.POLICY, self.policy)
        self.api = FakeSettings(self.policy, settings.selected_labels(manifest, self.config))

    def write(self, relative: str, value: Any) -> None:
        (self.root / relative).write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")

    def check(self) -> dict[str, Any]:
        report: dict[str, Any] = settings.check(self.root, self.api.get, REPOSITORY, NOW)
        return report

    def status_of(self, report: dict[str, Any], area: str, item: str) -> str:
        status: str = next(row["status"] for row in report["rows"] if row["area"] == area and row["item"] == item)
        return status

    def test_compliant_repository_passes(self) -> None:
        report = self.check()
        self.assertEqual(report["status"], "pass", settings.markdown(report))
        self.assertEqual(report["counts"], {"PASS": len(report["rows"]), "DRIFT": 0, "UNVERIFIED": 0, "NOT_APPLICABLE": 0})
        areas = {row["area"] for row in report["rows"]}
        self.assertEqual(areas, {"repository", "labels", "rulesets", "security", "actions"})
        self.assertEqual(report["checked_at"], "2026-11-09T12:00:00Z")
        self.assertEqual(report["policy_sha256"], hashlib.sha256((self.root / settings.POLICY).read_bytes()).hexdigest())

    def test_each_drift_kind_is_reported(self) -> None:
        label = self.api.labels[0]["name"]
        cases = (
            ("repository", "has_wiki", lambda: self.api.metadata.update(has_wiki=True)),
            ("repository", "description", lambda: self.api.metadata.update(description="Other")),
            ("repository", "topics", lambda: self.api.topics.remove("library")),
            ("labels", label, lambda: self.api.labels[0].update(color="000000")),
            ("labels", label, lambda: self.api.labels.pop(0)),
            ("rulesets", "branch baseline", lambda: self.api.rulesets[0].update(bypass_actors=[{"actor_id": 1}])),
            ("rulesets", "tag baseline", lambda: self.api.rulesets.pop()),
            ("security", "vulnerability_alerts", lambda: self.api.security.update(vulnerability_alerts=False)),
            ("security", "automated_security_fixes", lambda: self.api.security.update(paused=True)),
            ("security", "secret_scanning_push_protection", lambda: self.api.metadata["security_and_analysis"].update(
                secret_scanning_push_protection={"status": "disabled"})),
            ("actions", "default_workflow_permissions",
             lambda: self.api.actions["/actions/permissions/workflow"].update(default_workflow_permissions="write")),
            ("actions", "retention_days",
             lambda: self.api.actions["/actions/permissions/artifact-and-log-retention"].update(days=30)),
            ("actions", "sha_pinning_required",
             lambda: self.api.actions["/actions/permissions"].update(sha_pinning_required=False)),
        )
        for area, item, change in cases:
            with self.subTest(area=area, item=item):
                self.setUp()
                change()
                report = self.check()
                self.assertEqual(report["status"], "fail")
                self.assertEqual(self.status_of(report, area, item), "DRIFT")
                self.assertEqual(report["counts"]["DRIFT"], 1, settings.markdown(report))

    def test_narrower_or_stronger_settings_pass(self) -> None:
        self.api.metadata["allow_rebase_merge"] = False
        self.policy["security"]["private_vulnerability_reporting"] = False
        self.write(settings.POLICY, self.policy)
        report = self.check()
        self.assertEqual(report["status"], "pass", settings.markdown(report))

    def test_unreadable_settings_are_unverified_never_pass(self) -> None:
        cases = (
            ("security", "private_vulnerability_reporting", "/private-vulnerability-reporting"),
            ("security", "vulnerability_alerts", "/vulnerability-alerts"),
            ("security", "automated_security_fixes", "/automated-security-fixes"),
            ("actions", "retention_days", "/actions/permissions/artifact-and-log-retention"),
            ("actions", "default_workflow_permissions", "/actions/permissions/workflow"),
            ("rulesets", "branch baseline", "/rulesets"),
            ("labels", "selected labels", "/labels"),
            ("repository", "topics", "/topics"),
        )
        for area, item, path in cases:
            with self.subTest(path=path):
                self.setUp()
                self.api.unreadable.add(path)
                report = self.check()
                self.assertEqual((report["status"], self.status_of(report, area, item)), ("fail", "UNVERIFIED"))
        self.setUp()
        del self.api.metadata["security_and_analysis"]
        report = self.check()
        self.assertEqual(self.status_of(report, "security", "secret_scanning"), "UNVERIFIED")

    def test_private_repository_marks_public_only_settings_not_applicable(self) -> None:
        self.api.metadata.update(visibility="private", private=True)
        del self.api.metadata["security_and_analysis"]
        self.api.unreadable |= {"/private-vulnerability-reporting", "/actions/permissions/fork-pr-contributor-approval"}
        report = self.check()
        self.assertEqual(report["status"], "pass", settings.markdown(report))
        self.assertEqual(report["counts"]["NOT_APPLICABLE"], 4)
        self.assertNotIn("/actions/permissions/fork-pr-contributor-approval", self.api.paths)

    def test_policy_without_security_or_actions_checks_the_rest(self) -> None:
        for key in ("security", "actions"):
            self.policy.pop(key)
        self.write(settings.POLICY, self.policy)
        report = self.check()
        self.assertEqual(report["status"], "pass")
        self.assertFalse({"security", "actions"} & {row["area"] for row in report["rows"]})

    def test_identity_and_unreadable_repository_are_operational_errors(self) -> None:
        self.api.metadata["full_name"] = "other/repository"
        with self.assertRaisesRegex(ValueError, "identity"):
            self.check()
        self.api.unreadable.add("")
        with self.assertRaisesRegex(ValueError, "unreadable"):
            self.check()

    def test_cli_writes_dated_evidence_and_exits_one_on_drift(self) -> None:
        output, summary = self.root / "settings.json", self.root / "settings.md"
        arguments = ["--root", str(self.root), "--output", str(output), "--summary", str(summary)]
        with patch.object(settings, "live_getter", return_value=self.api.get), contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(settings.main(arguments), 0)
        record = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(record["repository"], REPOSITORY)
        self.assertRegex(record["checked_at"], r"^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$")
        self.assertEqual(record["policy_sha256"], hashlib.sha256((self.root / settings.POLICY).read_bytes()).hexdigest())
        self.assertIn("| security | vulnerability_alerts | true | true | PASS |", summary.read_text(encoding="utf-8"))
        self.api.security["vulnerability_alerts"] = False
        output.unlink()
        with patch.object(settings, "live_getter", return_value=self.api.get), contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(settings.main(arguments), 1)

    def test_live_client_sends_get_only_and_maps_errors(self) -> None:
        methods: list[tuple[str, str]] = []

        class Response(io.BytesIO):
            status = 200

            def __enter__(self) -> "Response":
                return self

            def __exit__(self, *details: object) -> None:
                self.close()

        class Opener:
            def open(self, request: Any, timeout: int) -> Response:
                methods.append((request.get_method(), request.full_url))
                if request.full_url.endswith("/vulnerability-alerts"):
                    raise HTTPError(request.full_url, 404, "Not Found", Message(), None)
                return Response(b'{"ok": true}')

        with patch.object(settings, "build_opener", return_value=Opener()):
            get = settings.live_getter(REPOSITORY, "fixture-token", "https://ghe.example.com/api/v3/")
            self.assertEqual(get(""), (200, {"ok": True}))
            self.assertEqual(get("/vulnerability-alerts"), (404, None))
            with self.assertRaises(ValueError):
                get("/../other")
        self.assertEqual(methods, [("GET", "https://ghe.example.com/api/v3/repos/example/generated"),
                                   ("GET", "https://ghe.example.com/api/v3/repos/example/generated/vulnerability-alerts")])
        self.assertFalse(hasattr(settings, "post") or hasattr(settings, "put") or hasattr(settings, "patch"))


if __name__ == "__main__":
    unittest.main()
