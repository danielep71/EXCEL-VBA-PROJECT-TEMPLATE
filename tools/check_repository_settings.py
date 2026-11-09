#!/usr/bin/env python3
"""Compare a repository's live GitHub settings with its committed policy, read-only.

The provisioner applies `.github/provisioning-policy.json` once, from a template
checkout. This checker re-reads the live repository at any later time and reports
every setting the policy covers as PASS, DRIFT, UNVERIFIED or NOT_APPLICABLE:
repository features and description, topics, the selected labels, the branch and
tag baselines, the security features and the GitHub Actions settings. It issues
GET requests only. Unreadable settings are UNVERIFIED and never pass; the exit
code is 1 on any drift or unverified setting. The JSON report carries the
observation time and the SHA-256 of the policy it compared against.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable
from urllib.error import HTTPError
from urllib.request import HTTPRedirectHandler, Request, build_opener

from _gatelib import run_gate

POLICY = ".github/provisioning-policy.json"
CONFIG = ".github/repository-profile.json"
RECORD = ".github/initialization.json"
LABELS = ".github/labels.json"
FEATURES = ("has_issues", "has_wiki", "has_projects", "has_discussions", "is_template",
            "allow_merge_commit", "allow_squash_merge", "allow_rebase_merge", "allow_auto_merge",
            "delete_branch_on_merge", "allow_update_branch")
METHODS = {"allow_merge_commit", "allow_squash_merge", "allow_rebase_merge"}
OPTIONAL_FEATURES = {"has_wiki", "has_projects", "has_discussions", "is_template"}
POLICY_KEYS = {"schema_version", "contract_version", "description", "topics", "features",
               "required_checks", "exceptions"}
OPTIONAL_POLICY_KEYS = {"security", "actions"}

# Security features in write order: Dependabot alerts before the security updates
# that need them. GitHub offers private vulnerability reporting and, without a
# paid licence, secret scanning and push protection on public repositories only.
SECURITY_FEATURES = ("private_vulnerability_reporting", "vulnerability_alerts", "automated_security_fixes",
                     "secret_scanning", "secret_scanning_push_protection")
PUBLIC_ONLY_SECURITY = {"private_vulnerability_reporting", "secret_scanning", "secret_scanning_push_protection"}
ANALYSIS_FEATURES = ("secret_scanning", "secret_scanning_push_protection")

# Each Actions setting: the endpoint under actions/permissions that reports it,
# its field there, and whether GitHub offers it on public repositories only.
ACTIONS_SETTINGS = {
    "sha_pinning_required": ("", "sha_pinning_required", False),
    "default_workflow_permissions": ("/workflow", "default_workflow_permissions", False),
    "can_approve_pull_request_reviews": ("/workflow", "can_approve_pull_request_reviews", False),
    "fork_pr_approval_policy": ("/fork-pr-contributor-approval", "approval_policy", True),
    "retention_days": ("/artifact-and-log-retention", "days", False),
}
FORK_POLICIES = ("first_time_contributors_new_to_github", "first_time_contributors", "all_external_contributors")
WORKFLOW_PERMISSIONS = ("read", "write")

UNIVERSAL_CHECK = "Repository integrity"

# Baseline ruleset names by target: the name created now, then legacy names from
# earlier contract adapters. Coverage is judged by content, so a legacy-named
# ruleset that still covers the baseline is simply kept; a baseline-named ruleset
# with insufficient controls, under any of these names, blocks for review.
BASELINE_RULESET_NAMES = {
    "branch": ("Default branch protection", "Template default branch"),
    "tag": ("Version tag protection", "Template version tags"),
}

Getter = Callable[[str], tuple[int, Any]]
STATUSES = ("PASS", "DRIFT", "UNVERIFIED", "NOT_APPLICABLE")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def is_universal_check(name: str) -> bool:
    """The universal job, run directly or as the job a reusable-workflow caller names."""
    return name == UNIVERSAL_CHECK or name.endswith(" / " + UNIVERSAL_CHECK)


def baseline_rules(policy: dict[str, Any]) -> list[dict[str, Any]]:
    common = [{"type": "deletion"}, {"type": "non_fast_forward"}]
    branch = {"name": BASELINE_RULESET_NAMES["branch"][0], "target": "branch", "enforcement": "active", "bypass_actors": [],
              "conditions": {"ref_name": {"include": ["~DEFAULT_BRANCH"], "exclude": []}},
              "rules": common + [{"type": "pull_request", "parameters": {"required_approving_review_count": 0,
                  "dismiss_stale_reviews_on_push": False, "require_code_owner_review": False,
                  "require_last_push_approval": False, "required_review_thread_resolution": False}},
                  {"type": "required_status_checks", "parameters": {"strict_required_status_checks_policy": True,
                      "required_status_checks": [{"context": name} for name in sorted(set(policy["required_checks"]))]}}]}
    tag = {"name": BASELINE_RULESET_NAMES["tag"][0], "target": "tag", "enforcement": "active", "bypass_actors": [],
           "conditions": {"ref_name": {"include": ["refs/tags/v*"], "exclude": []}},
           "rules": common + [{"type": "update", "parameters": {"update_allows_fetch_and_merge": False}}]}
    return [branch, tag]


def rule_covers(actual: dict[str, Any], desired: dict[str, Any]) -> bool:
    if actual.get("type") != desired["type"]:
        return False
    expected, params = desired.get("parameters", {}), actual.get("parameters", {})
    if desired["type"] == "required_status_checks":
        names = {row["context"] for row in params.get("required_status_checks", [])}
        return params.get("strict_required_status_checks_policy") is True and \
            {row["context"] for row in expected["required_status_checks"]} <= names
    if desired["type"] == "pull_request":
        return bool(params.get("required_approving_review_count", -1) >= expected["required_approving_review_count"])
    if desired["type"] == "update" and expected == {"update_allows_fetch_and_merge": False}:
        # GitHub can omit the false update parameter when reading a ruleset back.
        return params.get("update_allows_fetch_and_merge", False) is False
    return all(params.get(key) == value for key, value in expected.items())


def covered(existing: list[dict[str, Any]], desired: dict[str, Any], branch: str) -> bool:
    accepted = {"~ALL", "~DEFAULT_BRANCH", f"refs/heads/{branch}"} if desired["target"] == "branch" else {"~ALL", "refs/tags/v*"}
    rules = []
    for item in existing:
        scope = (item.get("conditions") or {}).get("ref_name", {})
        if item.get("target") == desired["target"] and item.get("enforcement") == "active" and \
                not item.get("bypass_actors") and not scope.get("exclude") and accepted.intersection(scope.get("include", [])):
            rules += item.get("rules", [])
    return all(any(rule_covers(actual, rule) for actual in rules) for rule in desired["rules"])


def validate_extensions(policy: dict[str, Any]) -> None:
    """Check the optional security and Actions sections of a provisioning policy."""
    require(set(policy) <= POLICY_KEYS | OPTIONAL_POLICY_KEYS and POLICY_KEYS <= set(policy),
            "Unsupported provisioning policy fields")
    if "security" in policy:
        security = policy["security"]
        require(isinstance(security, dict) and set(security) == set(SECURITY_FEATURES)
                and all(type(value) is bool for value in security.values()),
                "security must set each of " + ", ".join(SECURITY_FEATURES) + " to true or false")
        require(not security["secret_scanning_push_protection"] or security["secret_scanning"],
                "Push protection requires secret scanning")
        require(not security["automated_security_fixes"] or security["vulnerability_alerts"],
                "Dependabot security updates require Dependabot alerts")
    if "actions" in policy:
        actions = policy["actions"]
        require(isinstance(actions, dict) and set(actions) == set(ACTIONS_SETTINGS),
                "actions must set exactly " + ", ".join(ACTIONS_SETTINGS))
        days = actions["retention_days"]
        require(type(actions["sha_pinning_required"]) is bool
                and actions["default_workflow_permissions"] in WORKFLOW_PERMISSIONS
                and type(actions["can_approve_pull_request_reviews"]) is bool
                and actions["fork_pr_approval_policy"] in FORK_POLICIES
                and type(days) is int and 1 <= days <= 400,
                "Invalid Actions settings: sha_pinning_required and can_approve_pull_request_reviews are true or "
                "false, default_workflow_permissions is read or write, fork_pr_approval_policy is one of "
                + ", ".join(FORK_POLICIES) + ", and retention_days is 1 to 400")


def is_public(metadata: dict[str, Any]) -> bool:
    return metadata.get("visibility") == "public" or metadata.get("private") is False


def security_states(get: Getter, metadata: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Each security feature as observed: True, False, or None when GitHub does not say."""
    public = is_public(metadata)
    states: dict[str, dict[str, Any]] = {}
    status, body = get("/private-vulnerability-reporting")
    states["private_vulnerability_reporting"] = (
        {"observed": bool(body.get("enabled")), "source": "GET private-vulnerability-reporting"}
        if status == 200 and isinstance(body, dict) and isinstance(body.get("enabled"), bool) else
        {"observed": None, "source": f"GET private-vulnerability-reporting (HTTP {status}: unreadable)"})
    status, body = get("/automated-security-fixes")
    if status == 200 and isinstance(body, dict) and isinstance(body.get("enabled"), bool):
        paused = body["enabled"] and body.get("paused") is True
        fixes = {"observed": body["enabled"], "paused": paused,
                 "source": "GET automated-security-fixes" + (" (paused)" if paused else "")}
    elif status == 404:
        fixes = {"observed": False, "paused": False, "source": "GET automated-security-fixes (HTTP 404: off)"}
    else:
        fixes = {"observed": None, "paused": False, "source": f"GET automated-security-fixes (HTTP {status}: unreadable)"}
    # GitHub answers 204 when Dependabot alerts are on and 404 when they are off.
    status, _ = get("/vulnerability-alerts")
    if status == 204:
        states["vulnerability_alerts"] = {"observed": True, "source": "GET vulnerability-alerts (HTTP 204: on)"}
    elif status == 404:
        states["vulnerability_alerts"] = {"observed": False, "source": "GET vulnerability-alerts (HTTP 404: off)"}
    else:
        states["vulnerability_alerts"] = {"observed": None, "source": f"GET vulnerability-alerts (HTTP {status}: unreadable)"}
    states["automated_security_fixes"] = fixes
    analysis = metadata.get("security_and_analysis")
    for key in ANALYSIS_FEATURES:
        value = (analysis or {}).get(key, {}).get("status") if isinstance(analysis, dict) else None
        states[key] = ({"observed": value == "enabled", "source": f"security_and_analysis.{key}"}
                       if value in ("enabled", "disabled") else
                       {"observed": None, "source": f"security_and_analysis.{key} (not reported; administrator access needed)"})
    for key in SECURITY_FEATURES:
        states[key]["applies"] = public or key not in PUBLIC_ONLY_SECURITY
    return {key: states[key] for key in SECURITY_FEATURES}


def actions_states(get: Getter, metadata: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Each Actions setting as observed, or None with the reason it is not known."""
    public = is_public(metadata)
    answers: dict[str, tuple[int, Any]] = {}
    states: dict[str, dict[str, Any]] = {}
    for setting, (suffix, field, public_only) in ACTIONS_SETTINGS.items():
        endpoint = "actions/permissions" + suffix
        if public_only and not public:
            states[setting] = {"observed": None, "applies": False, "source": f"GET {endpoint} (public repositories only)"}
            continue
        if endpoint not in answers:
            answers[endpoint] = get("/" + endpoint)
        status, body = answers[endpoint]
        if status != 200 or not isinstance(body, dict) or field not in body:
            states[setting] = {"observed": None, "applies": True,
                               "source": f"GET {endpoint} (HTTP {status}: no {field})"}
        else:
            states[setting] = {"observed": body[field], "applies": True, "source": f"GET {endpoint} {field}"}
    return states


def collection(get: Getter, path: str) -> list[Any] | None:
    """Every row of a paginated collection, or None when it cannot be read whole."""
    rows: list[Any] = []
    for page in range(1, 101):
        status, values = get(f"{path}?per_page=100&page={page}")
        if status != 200 or not isinstance(values, list):
            return None
        rows.extend(values)
        if len(values) < 100:
            return rows
    return None


def selected_labels(manifest: dict[str, Any], config: dict[str, Any]) -> list[dict[str, Any]]:
    overlays = manifest.get("overlays", {})
    labels = list(manifest.get("core", []))
    if config.get("profile"):
        labels += overlays.get("profile", {}).get(config["profile"], [])
    for domain in config.get("label_domains", []):
        labels += overlays.get("domain", {}).get(domain, [])
    return labels


def row(area: str, item: str, expected: Any, observed: Any, status: str, source: str) -> dict[str, Any]:
    return {"area": area, "item": item, "expected": expected, "observed": observed, "status": status, "source": source}


def compare_features(policy: dict[str, Any], metadata: dict[str, Any], description: str | None) -> list[dict[str, Any]]:
    rows = []
    exceptions = policy.get("exceptions", {})
    if description is None:
        rows.append(row("repository", "description", "resolved project description", metadata.get("description"),
                        "UNVERIFIED", "policy description placeholder could not be resolved"))
    else:
        rows.append(row("repository", "description", description, metadata.get("description"),
                        "PASS" if metadata.get("description") == description else "DRIFT", "GET repository"))
    for key in FEATURES:
        want, have = policy["features"][key], metadata.get(key)
        if have is None:
            status, note = "UNVERIFIED", "GET repository (not reported)"
        elif have == want:
            status, note = "PASS", "GET repository"
        elif key in METHODS and want is True and have is False:
            status, note = "PASS", "GET repository (narrower merge method retained)"
        elif key in OPTIONAL_FEATURES and want is False and have is True and "feature:" + key in exceptions:
            status, note = "PASS", f"GET repository (recorded feature:{key} exception)"
        else:
            status, note = "DRIFT", "GET repository"
        rows.append(row("repository", key, want, have, status, note))
    return rows


def compare_security(policy: dict[str, Any], states: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []
    for key in SECURITY_FEATURES:
        want, state = policy["security"][key], states[key]
        have = state["observed"]
        if not state["applies"]:
            status = "NOT_APPLICABLE"
        elif have is None:
            status = "UNVERIFIED"
        elif state.get("paused") and want:
            status, have = "DRIFT", "paused"
        elif have == want or (have is True and want is False):
            status = "PASS"
        else:
            status = "DRIFT"
        rows.append(row("security", key, want, have, status, state["source"]))
    return rows


def compare_actions(policy: dict[str, Any], states: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []
    for key in ACTIONS_SETTINGS:
        want, state = policy["actions"][key], states[key]
        have = state["observed"]
        status = ("NOT_APPLICABLE" if not state["applies"] else "UNVERIFIED" if have is None
                  else "PASS" if have == want else "DRIFT")
        rows.append(row("actions", key, want, have, status, state["source"]))
    return rows


def compare_labels(get: Getter, manifest: dict[str, Any], config: dict[str, Any]) -> list[dict[str, Any]]:
    live = collection(get, "/labels")
    wanted = selected_labels(manifest, config)
    if live is None:
        return [row("labels", "selected labels", len(wanted), None, "UNVERIFIED", "GET labels (unreadable)")]
    observed = {str(label.get("name", "")).lower(): label for label in live if isinstance(label, dict)}
    rows = []
    for label in wanted:
        have = observed.get(label["name"].lower())
        shown = None if have is None else {"name": have.get("name"), "color": str(have.get("color", "")).upper(),
                                            "description": have.get("description") or ""}
        rows.append(row("labels", label["name"], {key: label[key] for key in ("name", "color", "description")}, shown,
                        "PASS" if shown == {key: label[key] for key in ("name", "color", "description")} else "DRIFT",
                        "GET labels"))
    return rows


def compare_rulesets(get: Getter, policy: dict[str, Any], branch: str) -> list[dict[str, Any]]:
    summaries = collection(get, "/rulesets")
    details = []
    for summary in summaries or []:
        status, detail = get(f"/rulesets/{summary.get('id')}")
        if status != 200 or not isinstance(detail, dict) or not isinstance(detail.get("bypass_actors"), list):
            summaries = None
            break
        details.append(detail)
    rows = []
    for desired in baseline_rules(policy):
        label = f"{desired['target']} baseline"
        if summaries is None:
            rows.append(row("rulesets", label, desired["name"], None, "UNVERIFIED",
                            "GET rulesets (unreadable, or bypass actors not reported)"))
        else:
            ok = covered(details, desired, branch)
            rows.append(row("rulesets", label, desired["name"], "covered" if ok else "not covered",
                            "PASS" if ok else "DRIFT", "GET rulesets (coverage by content)"))
    return rows


def check(root: Path, get: Getter, repository: str | None = None, now: datetime | None = None) -> dict[str, Any]:
    policy_bytes = (root / POLICY).read_bytes()
    policy = json.loads(policy_bytes)
    validate_extensions(policy)
    config = json.loads((root / CONFIG).read_text(encoding="utf-8"))
    manifest = json.loads((root / LABELS).read_text(encoding="utf-8"))
    repository = repository or config.get("repository")
    require(isinstance(repository, str) and bool(re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository)),
            "Name the repository with --repository OWNER/NAME")
    description = policy["description"]
    if description == "@initialization:PROJECT_DESCRIPTION":
        record = root / RECORD
        description = (json.loads(record.read_text(encoding="utf-8")).get("values", {}).get("PROJECT_DESCRIPTION")
                       if record.is_file() else None)
    status, metadata = get("")
    require(status == 200 and isinstance(metadata, dict), f"Repository {repository} is unreadable (HTTP {status})")
    require(str(metadata.get("full_name", "")).lower() == repository.lower(), "Repository identity differs")
    rows = compare_features(policy, metadata, description)
    status, topics = get("/topics")
    required = sorted(set(policy["topics"]) | {"excel", "vba"} | ({config["profile"]} if config.get("profile") else set()))
    if status != 200 or not isinstance(topics, dict) or not isinstance(topics.get("names"), list):
        rows.append(row("repository", "topics", required, None, "UNVERIFIED", f"GET topics (HTTP {status})"))
    else:
        missing = [name for name in required if name not in topics["names"]]
        rows.append(row("repository", "topics", required, sorted(topics["names"]),
                        "DRIFT" if missing else "PASS", "GET topics" + (f" (missing {', '.join(missing)})" if missing else "")))
    rows += compare_labels(get, manifest, config)
    rows += compare_rulesets(get, policy, str(metadata.get("default_branch", "")))
    if "security" in policy:
        rows += compare_security(policy, security_states(get, metadata))
    if "actions" in policy:
        rows += compare_actions(policy, actions_states(get, metadata))
    counts = {status: sum(item["status"] == status for item in rows) for status in STATUSES}
    stamp = (now or datetime.now(timezone.utc)).astimezone(timezone.utc).replace(microsecond=0)
    return {
        "status": "fail" if counts["DRIFT"] or counts["UNVERIFIED"] else "pass",
        "repository": repository, "checked_at": stamp.isoformat().replace("+00:00", "Z"),
        "policy_path": POLICY, "policy_sha256": hashlib.sha256(policy_bytes).hexdigest(),
        "counts": counts, "rows": rows,
        "scope": "Read-only comparison of live settings with the committed policy; GET requests only.",
    }


def cell(value: Any) -> str:
    text = value if isinstance(value, str) else json.dumps(value, sort_keys=True)
    return text.replace("\r", " ").replace("\n", " ").replace("|", "\\|")


def markdown(report: dict[str, Any]) -> str:
    counts = report["counts"]
    lines = ["# Repository settings check", "", f"Result: {report['status'].upper()}", "",
             f"Repository: {report['repository']} · Checked at: {report['checked_at']}", "",
             f"Policy: `{report['policy_path']}` (SHA-256 `{report['policy_sha256']}`)", "",
             " · ".join(f"{key}: {counts[key]}" for key in STATUSES), "", report["scope"], "",
             "| Area | Setting | Expected | Observed | Status | Source |", "| --- | --- | --- | --- | --- | --- |"]
    lines += [f"| {item['area']} | {cell(item['item'])} | {cell(item['expected'])} | {cell(item['observed'])} | "
              f"{item['status']} | {cell(item['source'])} |" for item in report["rows"]]
    return "\n".join(lines) + "\n"


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, *args: Any, **kwargs: Any) -> None:
        return None


def live_getter(repository: str, token: str | None, base: str | None = None) -> Getter:
    """A GET-only client: it has no way to send any other method."""
    root = (base or os.environ.get("GITHUB_API_URL") or "https://api.github.com").rstrip("/")

    def get(path: str) -> tuple[int, Any]:
        require(path == "" or path.startswith("/") and not path.startswith("//") and ".." not in path,
                "Invalid repository-relative API path")
        headers = {"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"}
        if token:
            headers["Authorization"] = "Bearer " + token
        request = Request(f"{root}/repos/{repository}{path}", headers=headers, method="GET")
        try:
            with build_opener(NoRedirect).open(request, timeout=30) as response:
                data = response.read(10_000_001)
                require(len(data) <= 10_000_000, "Response size limit")
                return int(response.status), json.loads(data) if data else None
        except HTTPError as error:
            return int(error.code), None

    return get


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="checkout holding the committed policy")
    parser.add_argument("--repository", help="OWNER/NAME; defaults to the repository recorded in the profile")
    parser.add_argument("--output", type=Path, help="JSON evidence record")
    parser.add_argument("--summary", type=Path, help="Markdown evidence record")
    options = parser.parse_args(arguments)
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")

    def build() -> dict[str, Any]:
        config = json.loads((options.root / CONFIG).read_text(encoding="utf-8"))
        repository = options.repository or config.get("repository")
        return check(options.root, live_getter(str(repository), token), repository)

    return run_gate(options, build=build, markdown=markdown, errors=(OSError, ValueError, KeyError, TypeError))


if __name__ == "__main__":
    raise SystemExit(main())
