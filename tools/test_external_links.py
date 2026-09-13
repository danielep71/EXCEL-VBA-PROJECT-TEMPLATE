#!/usr/bin/env python3
"""Focused regression tests for external-link policy/classification precedence."""

from __future__ import annotations

import hashlib
import unittest
from pathlib import Path
from typing import Any
from unittest.mock import patch

import check_external_links as external_links


class ExternalLinkPolicyPrecedenceTests(unittest.TestCase):
    @staticmethod
    def policy(url: str, kind: str = "restricted-historical") -> dict[str, Any]:
        return {
            "attempts": 3,
            "timeout_seconds": 1,
            "concurrency": 1,
            "redirects": 1,
            "max_links": 10,
            "domains": {"example.com": "fixture"},
            "exceptions": [],
            "classifications": [
                {
                    "id": hashlib.sha256(url.encode()).hexdigest(),
                    "kind": kind,
                    "reason": "fixture",
                    "expires": "2027-01-01",
                }
            ],
        }

    def report(self, url: str, kind: str = "restricted-historical") -> dict[str, Any]:
        digest = hashlib.sha256(url.encode()).hexdigest()
        policy = self.policy(url, kind)
        links = {digest: {"url": url, "locations": ["README.md:1"]}}

        def no_network(_url: str, _timeout: int) -> tuple[int, str | None]:
            raise AssertionError("classified links must not be probed")

        with (
            patch.object(external_links, "load_policy", return_value={"network": policy}),
            patch.object(external_links, "collect", return_value=links),
        ):
            return external_links.build_report(
                Path("."),
                external_links.date(2026, 9, 13),
                transport=no_network,
                pause=lambda _seconds: None,
            )

    def test_policy_blocked_precedes_classification_without_network(self) -> None:
        report = self.report("http://example.com/private")
        row = report["links"][0]
        self.assertEqual(row["status"], "POLICY_BLOCKED")
        self.assertEqual(row["attempts"], 0)
        self.assertEqual(row["classification"], "restricted-historical")

    def test_access_restricted_precedes_classification_without_network(self) -> None:
        report = self.report("https://example.com/private?token=fixture")
        row = report["links"][0]
        self.assertEqual(row["status"], "ACCESS_RESTRICTED")
        self.assertEqual(row["attempts"], 0)
        self.assertEqual(report["counts"]["access_restricted"], 1)

    def test_policy_clean_classification_remains_classified_and_offline(self) -> None:
        report = self.report("https://example.com/historical")
        row = report["links"][0]
        self.assertEqual(row["status"], "RESTRICTED_HISTORICAL")
        self.assertEqual(row["attempts"], 0)
        self.assertEqual(report["status"], "fail")

        pending = self.report("https://example.com/release", "pending-publication")
        self.assertEqual(pending["links"][0]["status"], "PENDING_PUBLICATION")

    def test_markdown_renders_every_count_owned_by_counts(self) -> None:
        rows = [
            {"status": "PERMANENT_FAILURE"},
            {"status": "RESTRICTED_HISTORICAL"},
            {"status": "PENDING_PUBLICATION"},
            {"status": "ACCESS_RESTRICTED"},
            {"status": "TRANSIENT_FAILURE"},
        ]
        counts = external_links._counts(rows)
        self.assertEqual(set(counts), set(external_links.COUNT_LABELS))

        report = {
            "status": "fail",
            "discovered": len(rows),
            "limit_exceeded": False,
            "counts": counts,
            "links": [],
            "scope_note": "fixture scope",
        }
        rendered = external_links.markdown(report)
        for key, label in external_links.COUNT_LABELS.items():
            self.assertIn(f"{label}: {counts[key]}", rendered)


if __name__ == "__main__":
    unittest.main()
