#!/usr/bin/env python3
"""Capture GitHub evidence using GET only; emit private snapshot JSON to stdout.

Optional GH_TOKEN grants read access only. No inspected repository is modified.
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import sys
from typing import Any
import urllib.error
import urllib.parse
import urllib.request

from check_portfolio_drift import CONFIG, PATHS


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req: Any, fp: Any, code: int, msg: str,
                         headers: Any, newurl: str) -> None:
        return None


def get(path: str) -> Any:
    headers = {"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"}
    token = os.environ.get("GH_TOKEN")
    if token:
        headers["Authorization"] = "Bearer " + token
    request = urllib.request.Request("https://api.github.com/repos/" + path, headers=headers, method="GET")
    with urllib.request.build_opener(NoRedirect).open(request, timeout=30) as response:
        data = response.read(10_000_001)
    if len(data) > 10_000_000:
        raise ValueError("Response exceeds snapshot size limit")
    return json.loads(data)


def pages(path: str) -> list[Any]:
    result = []
    for page in range(1, 101):
        data = get(f"{path}?per_page=100&page={page}")
        if not isinstance(data, list):
            raise ValueError("Expected paginated collection")
        result.extend(data)
        if len(data) < 100:
            return result
    raise ValueError("Pagination limit reached; refusing partial evidence")


def collect(repository: str) -> dict[str, Any]:
    if not re.fullmatch(r"[\w.-]+/[\w.-]+", repository):
        raise ValueError("Repository must be owner/name")
    metadata = get(repository)
    branch = urllib.parse.quote(metadata["default_branch"], safe="")
    commit = get(f"{repository}/commits/{branch}")["sha"]
    tree = get(f"{repository}/git/trees/{commit}?recursive=1")
    unavailable = {}
    blobs = {item["path"]: item["sha"] for item in tree["tree"] if item["type"] == "blob"}
    if tree.get("truncated"):
        unavailable["paths"] = "Recursive tree truncated"
    required = {path for paths in PATHS.values() for path in paths}
    required.update(path for path in blobs if path.startswith(".github/workflows/"))
    required.add(CONFIG)
    files = {}
    for path in sorted(required & set(blobs)):
        try:
            blob = get(f"{repository}/git/blobs/{blobs[path]}")
            files[path] = base64.b64decode(blob["content"]).decode("utf-8")
        except (urllib.error.HTTPError, UnicodeError, ValueError) as error:
            unavailable[path] = type(error).__name__
    live: dict[str, Any] = {}
    for resource in ("labels", "rulesets"):
        try:
            values = pages(f"{repository}/{resource}")
            live[resource] = ([get(f"{repository}/rulesets/{item['id']}") for item in values]
                              if resource == "rulesets" else values)
        except urllib.error.HTTPError as error:
            live[resource] = None
            unavailable[resource] = f"HTTP {error.code}; not evidence of absence"
    return {"repository": repository, "commit": commit,
            "paths": None if tree.get("truncated") else sorted(blobs), "files": files,
            "metadata": {key: metadata.get(key) for key in
                         ("default_branch", "description", "has_issues", "allow_auto_merge")},
            **live, "unavailable": unavailable, "decisions": []}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("repositories", nargs="+")
    parser.add_argument("--contract-commit", required=True, help="Exact template evaluator commit")
    options = parser.parse_args()
    try:
        if len(set(options.repositories)) != len(options.repositories):
            raise ValueError("Duplicate repository")
        if not re.fullmatch(r"[0-9a-f]{40}", options.contract_commit):
            raise ValueError("Invalid evaluator commit")
        data = {"schema_version": 1, "contract_source": "danielep71/EXCEL-VBA-PROJECT-TEMPLATE",
                "contract_commit": options.contract_commit,
                "repositories": [collect(name) for name in sorted(options.repositories)]}
        print(json.dumps(data, indent=2, sort_keys=True))
        return 0
    except (OSError, ValueError, KeyError) as error:
        print(f"Capture failed: {type(error).__name__}; no complete snapshot emitted", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
