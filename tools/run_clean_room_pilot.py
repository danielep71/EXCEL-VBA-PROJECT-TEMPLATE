#!/usr/bin/env python3
"""Run the clean-room maintainer journey of a template candidate automatically.

Run it on the maintainer's Windows machine with desktop Excel, from a clean checkout
of the template candidate, with ``gh`` signed in. The token needs ``repo`` and
``workflow``, plus ``delete_repo`` for ``--cleanup delete``. It follows the Wiki
guide end to end on a new public repository and times every step:

1. creates the repository with **Use this template** (``gh repo create --template``)
   and turns off Wikis and Projects;
2. clones it, initializes the profile, proves the second run is a no-op, and
   merges the setup pull request once its checks pass;
3. provisions it: plans with the template's provisioner, then applies exactly
   that plan digest;
4. sets the GitHub Actions settings that the provisioner leaves to the
   maintainer, then requires ``check_repository_settings.py`` to report every
   setting as PASS;
5. cuts ``v0.1.0`` and merges the cut once its checks pass;
6. certifies the release candidate in Excel with the pilot's own
   ``Certify-Excel.ps1``;
7. assembles the release evidence, then creates and checks the annotated tag
   before and after tagging, and pushes it;
8. publishes the GitHub Release, closes its milestone and requires a green
   *Release closeout*.

It writes ``clean-room-journey.md`` and ``clean-room-journey.json`` under
``--work-dir``. A failed step is recorded with its error before the run stops,
so the record always shows how far the journey got. Finally the pilot is
deleted, archived or kept, as ``--cleanup`` says; the record keeps the evidence.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Callable

import prepare_release_evidence as evidence
import publish_release

# The mandatory journey, in order; a record is complete only when every step passed.
JOURNEY_STEPS = (
    "Create the repository from the template",
    "Initialize the profile and merge the setup pull request",
    "Provision: plan, then apply the approved digest",
    "Set the Actions settings and pass the live settings check",
    "Cut v0.1.0 and merge the cut",
    "Certify the release candidate in Excel",
    "Evidence, annotated tag, GitHub Release and closeout",
)
POLL_SECONDS = 15
POLL_LIMIT = 80
ACTIONS_SETTINGS = (
    ("actions/permissions", {"enabled": True, "allowed_actions": "all", "sha_pinning_required": True}),
    ("actions/permissions/workflow", {"default_workflow_permissions": "read",
                                      "can_approve_pull_request_reviews": False}),
    ("actions/permissions/fork-pr-contributor-approval", {"approval_policy": "all_external_contributors"}),
    ("actions/permissions/artifact-and-log-retention", {"days": 90}),
)


def stream(command: list[str], cwd: Path) -> int:
    """Run with the console attached, for Excel certification, which may ask the operator."""
    return subprocess.run(command, cwd=cwd, check=False).returncode


def on_windows() -> bool:
    return os.name == "nt"


def now() -> datetime:
    return datetime.now(timezone.utc)


class Journey:
    """Numbered, timed steps; written to disk after every step so a failure keeps its record."""

    def __init__(self, work: Path, header: dict[str, Any]) -> None:
        self.work, self.header = work, header
        self.steps: list[dict[str, Any]] = []

    def step(self, name: str, action: Callable[[], str]) -> str:
        started = now()
        row: dict[str, Any] = {"number": len(self.steps) + 1, "name": name,
                               "started_utc": started.strftime("%H:%M:%S")}
        self.steps.append(row)
        try:
            row["evidence"] = action()
            row["result"] = "PASS"
            return str(row["evidence"])
        except evidence.EvidenceError as error:
            row.update(result="FAIL", evidence=str(error))
            raise
        finally:
            row["seconds"] = round((now() - started).total_seconds())
            self.write()

    def write(self) -> None:
        evidence.write_json(self.work / "clean-room-journey.json", {**self.header, "steps": self.steps})
        total = sum(int(row.get("seconds", 0)) for row in self.steps)
        lines = [f"# Clean-room maintainer journey of {self.header['template_sha'][:7]}", "",
                 *[f"- **{key.replace('_', ' ').capitalize()}:** `{value}`" for key, value in self.header.items()],
                 "", "| # | Step | Start (UTC) | Seconds | Result | Evidence |", "| --- | --- | --- | ---: | --- | --- |"]
        for row in self.steps:
            text = str(row.get("evidence", "")).replace("|", "\\|").replace("\n", " ")
            lines.append(f"| {row['number']} | {row['name']} | {row['started_utc']} | {row.get('seconds', '')} | "
                         f"{row.get('result', 'RUNNING')} | {text} |")
        failed = [row for row in self.steps if row.get("result") == "FAIL"]
        verdict = f"FAIL at step {failed[0]['number']}" if failed else "PASS"
        lines += ["", f"Verdict: **{verdict}**; {len(self.steps)} steps in {total} s.", ""]
        (self.work / "clean-room-journey.md").write_text("\n".join(lines), encoding="utf-8")


def gh_api(root: Path, method: str, endpoint: str, body: dict[str, Any] | None = None) -> bytes:
    command = ["gh", "api", "-X", method, endpoint]
    for name, value in (body or {}).items():
        command += ["-F" if isinstance(value, (bool, int)) else "-f",
                    f"{name}={str(value).lower() if isinstance(value, bool) else value}"]
    return evidence.checked(command, root, f"{method} {endpoint}")


def wait_for(probe: Callable[[], bool], what: str) -> None:
    for _ in range(POLL_LIMIT):
        if probe():
            return
        time.sleep(POLL_SECONDS)
    raise evidence.EvidenceError(f"timed out waiting for {what}")


def checks_green(root: Path, repository: str, sha: str) -> Callable[[], bool]:
    """A probe that is true once every check run at the commit has completed, failing on any failure."""
    def probe() -> bool:
        runs = evidence.gh_json(root, f"repos/{repository}/commits/{sha}/check-runs?per_page=100")["check_runs"]
        failed = [run["name"] for run in runs if run.get("conclusion") in ("failure", "cancelled", "timed_out")]
        evidence.require(not failed, f"checks failed at {sha[:7]}: {', '.join(failed)}")
        names = {run["name"] for run in runs}
        return "Repository integrity" in names and all(run.get("status") == "completed" for run in runs)
    return probe


def merge_pull_request(root: Path, pilot: Path, repository: str, branch: str) -> str:
    pr = evidence.gh_json(root, f"repos/{repository}/pulls?head={repository.split('/')[0]}:{branch}&state=open")
    evidence.require(bool(pr), f"no open pull request for {branch}")
    head = str(pr[0]["head"]["sha"])
    wait_for(checks_green(root, repository, head), f"the checks of pull request #{pr[0]['number']}")
    evidence.checked(["gh", "pr", "merge", str(pr[0]["number"]), "--repo", repository, "--squash",
                      "--delete-branch"], pilot, "pull request merge")
    evidence.checked(["git", "switch", "main"], pilot, "switch to main")
    evidence.checked(["git", "pull", "--ff-only"], pilot, "pull main")
    merged = evidence.checked(["git", "rev-parse", "HEAD"], pilot, "rev-parse").decode().strip()
    return f"PR #{pr[0]['number']} merged as {merged}"


def library_checks(bundle: Path) -> list[str]:
    """Profile checks of the library reference, read from the Excel results the bundle recorded."""
    suites = json.loads((bundle / "suites.json").read_text(encoding="utf-8"))
    rows = {item["module"]: item for item in [*suites["suites"], *suites["examples"]]}
    text, example = rows.get("TextTests", {}), rows.get("TextExample", {})
    evidence.require(text.get("status") == "PASS" and example.get("status") == "PASS",
                     "the library reference suite or example did not pass in Excel")
    return ["--check", f"public-api=check_vba_public_api.py passed at the candidate; TextTests.RunTextTests: {text['detail']}",
            "--check", f"caller-contract=TextExample.RunTextExample, which calls only TextFacade, {example['detail']}"]


class Pilot:
    def __init__(self, options: argparse.Namespace) -> None:
        self.options = options
        self.template = options.template_root.resolve()
        self.work = options.work_dir.resolve()
        self.repository = f"{options.owner}/{options.name}"
        self.pilot = self.work / options.name
        self.sha = options.template_sha
        self.initialized = self.candidate = ""
        self.created = False

    def python(self, script: Path, *arguments: str) -> list[str]:
        return [sys.executable, str(script), *arguments]

    def create(self) -> str:
        source = json.loads((self.template / ".github/repository-profile.json").read_text(encoding="utf-8"))
        template = str(source["repository"])
        head = evidence.gh_json(self.template, f"repos/{template}/commits/main")["sha"]
        evidence.require(head == self.sha, f"{template} main is {head}, not the candidate {self.sha}: "
                         "Use this template copies the default branch")
        evidence.checked(["gh", "repo", "create", self.repository, "--public", "--template", template],
                         self.template, "repository creation")
        self.created = True

        def populated() -> bool:
            return evidence.run(["gh", "api", f"repos/{self.repository}/commits?per_page=1"], self.template).returncode == 0
        wait_for(populated, "the template copy")
        gh_api(self.template, "PATCH", f"repos/{self.repository}", {"has_wiki": False, "has_projects": False})
        return f"https://github.com/{self.repository} from {template}@{self.sha[:7]}; Wikis and Projects off"

    def initialize(self) -> str:
        evidence.disposable(self.template, self.pilot)
        evidence.checked(["git", "clone", "--quiet", f"https://github.com/{self.repository}.git", str(self.pilot)],
                         self.work, "clone")
        start = evidence.checked(["git", "rev-parse", "HEAD"], self.pilot, "rev-parse").decode().strip()
        evidence.checked(["git", "switch", "-c", "setup/initialize"], self.pilot, "setup branch")
        values = {"PROJECT_NAME": self.options.project_name, "PROJECT_TAGLINE": "Clean-room maintainer pilot",
                  "PROJECT_DESCRIPTION": f"Clean-room maintainer pilot of template {self.sha[:7]}.",
                  "REPOSITORY_PATH": self.repository, "MAINTAINER_NAME": self.options.maintainer,
                  "SUPPORT_CONTACT": f"https://github.com/{self.repository}/security/advisories/new",
                  "COPYRIGHT_YEAR": str(date.today().year)}
        arguments = ["--profile", self.options.profile]
        for name, value in values.items():
            arguments += ["--set", f"{name}={value}"]
        initializer = self.pilot / "tools/initialize_repository.py"
        evidence.checked(self.python(initializer, *arguments, "--apply"), self.pilot, "initialization")
        evidence.checked(["git", "add", "--all"], self.pilot, "staging")
        for gate in ("check_repo.py", "check_vba_public_api.py"):
            evidence.checked(self.python(self.pilot / "tools" / gate, "--root", "."), self.pilot, gate)
        evidence.checked(["git", "commit", "--quiet", "-m", "Initialize the selected project profile"], self.pilot,
                         "setup commit")
        rerun = evidence.checked(self.python(initializer, *arguments), self.pilot, "idempotence check")
        evidence.require(json.loads(rerun).get("status") == "no-op", "the second initializer run is not a no-op")
        evidence.checked(["git", "push", "--set-upstream", "origin", "setup/initialize"], self.pilot, "push")
        evidence.checked(["gh", "pr", "create", "--repo", self.repository, "--base", "main", "--head",
                          "setup/initialize", "--title", "Initialize the selected project profile", "--fill"],
                         self.pilot, "setup pull request")
        merged = merge_pull_request(self.template, self.pilot, self.repository, "setup/initialize")
        self.initialized = merged.rsplit(" ", 1)[1]
        return f"start {start[:7]}; {self.options.profile} initialized, second run no-op; {merged}"

    def provision(self) -> str:
        contract = json.loads((self.pilot / ".github/initialization.json").read_text(encoding="utf-8"))
        version = contract["template_contract"]["version"]
        tool = self.template / "tools/provision_repository.py"
        common = ["--repository", self.repository, "--profile", self.options.profile,
                  "--contract-version", version, "--source-sha", self.initialized]
        plan_path = self.work / "provision-plan.json"
        evidence.checked(self.python(tool, *common, "--output", str(plan_path)), self.template, "provisioning plan")
        plan = json.loads(plan_path.read_text(encoding="utf-8"))["plan"]
        evidence.require(not plan["blocked"], "provisioning plan is blocked: " + "; ".join(plan["blocked"]))
        apply_path = self.work / "provision-apply.json"
        evidence.checked(self.python(tool, *common, "--apply", "--approve-plan", plan["plan_sha256"],
                                     "--journal", str(self.work / "provision-journal.json"),
                                     "--output", str(apply_path)), self.template, "provisioning apply")
        verification = json.loads(apply_path.read_text(encoding="utf-8"))["verification"]
        evidence.require(not verification["remaining_actions"] and not verification["blocked"],
                         "provisioning left actions or blockers after apply")
        return f"plan {plan['plan_sha256'][:12]}: {len(plan['actions'])} actions applied and verified"

    def settings(self) -> str:
        for endpoint, body in ACTIONS_SETTINGS:
            gh_api(self.template, "PUT", f"repos/{self.repository}/{endpoint}", body)
        report = self.work / "settings-check.json"
        evidence.checked(self.python(self.pilot / "tools/check_repository_settings.py", "--root", ".",
                                     "--output", str(report)), self.pilot, "live settings check")
        status = json.loads(report.read_text(encoding="utf-8")).get("status")
        evidence.require(status == "pass", f"the live settings check reports {status}")
        return "Actions settings set; live settings check PASS with no drift or unverified setting"

    def cut(self) -> str:
        changelog = self.pilot / "CHANGELOG.md"
        text = changelog.read_text(encoding="utf-8")
        entry = (f"### Added\n\n- Initial `{self.options.profile}` release of the clean-room pilot of template "
                 f"`{self.sha[:7]}`.\n")
        # Detached, so local main stays at origin/main: the cut branch starts here and its squash
        # merge would otherwise leave main diverged and the later fast-forward pull failing.
        evidence.checked(["git", "switch", "--detach"], self.pilot, "detach from main")
        changelog.write_text(text.replace("No unreleased changes recorded.\n", entry, 1), encoding="utf-8")
        evidence.checked(["git", "commit", "--quiet", "-am", "Record the initial release entry"], self.pilot,
                         "release entry")
        evidence.checked(self.python(self.template / "tools/cut_release.py", "--root", ".", "--version", "0.1.0",
                                     "--push"), self.pilot, "release cut")
        merged = merge_pull_request(self.template, self.pilot, self.repository, "chore/cut-v0.1.0")
        self.candidate = merged.rsplit(" ", 1)[1]
        return merged

    def certify(self) -> str:
        evidence.require(on_windows(), "Excel certification needs Windows with desktop Excel")
        bundle = self.work / f"excel-evidence-{self.candidate[:7]}"
        code = stream(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                       str(self.pilot / "tools/Certify-Excel.ps1"), "-RepoRoot", str(self.pilot),
                       "-BundleDir", str(bundle)], self.pilot)
        evidence.require(code == 0, f"Certify-Excel did not pass (exit {code}); see {bundle}")
        return f"{bundle.name}: host record PASS"

    def release(self) -> str:
        bundle = self.work / f"excel-evidence-{self.candidate[:7]}"
        release_work = self.work / "release-work"
        wait_for(checks_green(self.template, self.repository, self.candidate), "the checks of the candidate")
        checks = library_checks(bundle) if self.options.profile == "library" else []
        for value in self.options.check:
            checks += ["--check", value]
        evidence.checked(self.python(self.template / "tools/prepare_release_evidence.py", "prepare", "--root", ".",
                                     "--tag", "v0.1.0", "--candidate-sha", self.candidate, "--work-dir",
                                     str(release_work), "--excel-bundle", str(bundle), *checks),
                         self.pilot, "release evidence")
        check = ["--tag", "v0.1.0", "--candidate-sha", self.candidate, "--evidence",
                 str(release_work / "release-evidence.json"), "--excel-evidence", str(release_work / "evidence/host.json")]
        evidence.checked(["git", "tag", "-a", "v0.1.0", self.candidate, "-m", f"{self.options.project_name} 0.1.0"],
                         self.pilot, "annotated tag")
        evidence.checked(self.python(self.pilot / "tools/check_release.py", "--root", ".", *check,
                                     "--require-tag-ref"), self.pilot, "post-tag check")
        evidence.checked(["git", "push", "origin", "refs/tags/v0.1.0"], self.pilot, "tag push")
        gh_api(self.template, "POST", f"repos/{self.repository}/milestones", {"title": "v0.1.0"})
        evidence.checked(["gh", "release", "create", "v0.1.0", "--repo", self.repository, "--verify-tag",
                          "--title", f"{self.options.project_name} 0.1.0", "--notes",
                          "Initial release of the clean-room maintainer pilot.", "--latest"], self.pilot,
                         "GitHub Release")
        identity = {"repository": self.repository, "tag": "v0.1.0", "sha": self.candidate}
        url = publish_release.close_out(self.template, identity, False)
        return f"v0.1.0 at {self.candidate[:7]} released; closeout {url}"

    def cleanup(self) -> str:
        mode = self.options.cleanup
        if mode == "delete":
            evidence.checked(["gh", "repo", "delete", self.repository, "--yes"], self.template, "repository deletion")
        elif mode == "archive":
            gh_api(self.template, "PATCH", f"repos/{self.repository}", {"archived": True})
        return f"pilot {mode}d" if mode != "keep" else "pilot kept"


def run(options: argparse.Namespace) -> int:
    pilot = Pilot(options)
    head = evidence.checked(["git", "rev-parse", "HEAD"], pilot.template, "rev-parse").decode().strip()
    evidence.require(head == pilot.sha, f"the template checkout is at {head}, not {pilot.sha}")
    evidence.require(options.profile == "library" or bool(options.check),
                     "profiles other than library need their release checks as --check NAME=DETAIL")
    # Freshness first: nothing under an existing work directory is ever deleted.
    evidence.require(not pilot.work.exists() or not any(pilot.work.iterdir()),
                     f"{pilot.work} is not empty: every journey needs a fresh --work-dir")
    evidence.disposable(pilot.template, pilot.work / "release-work")
    pilot.work.mkdir(parents=True, exist_ok=True)
    if not os.environ.get("GH_TOKEN"):
        # Provisioning and the settings check read the token from the environment, never from arguments.
        os.environ["GH_TOKEN"] = evidence.checked(["gh", "auth", "token"], pilot.template, "gh auth token").decode().strip()
    journey = Journey(pilot.work, {"template_sha": pilot.sha, "pilot": pilot.repository,
                                   "profile": options.profile, "governance_tier": "full",
                                   "started_utc": now().isoformat(timespec="seconds"), "cleanup": options.cleanup})
    steps = tuple(zip(JOURNEY_STEPS, (pilot.create, pilot.initialize, pilot.provision, pilot.settings,
                                      pilot.cut, pilot.certify, pilot.release)))
    failed = True
    try:
        for name, action in steps:
            journey.step(name, action)
        failed = False
    finally:
        clean_up(pilot, journey, failed)
    print(f"JOURNEY PASS: record in {pilot.work / 'clean-room-journey.md'}")
    return 0


def clean_up(pilot: Pilot, journey: Journey, failed: bool) -> None:
    """Delete or archive a pilot the run created; declining keeps it and is not a journey failure.

    After a failure the pilot stays for inspection unless the maintainer confirms; --yes never
    removes a failed pilot.
    """
    mode = pilot.options.cleanup
    if not pilot.created or mode == "keep":
        return
    try:
        publish_release.confirm(f"{mode.capitalize()} the pilot {pilot.repository} now?",
                                pilot.options.yes and not failed)
    except evidence.EvidenceError:
        print(f"KEPT: {pilot.repository} was not {mode}d")
        return
    journey.step(f"Clean up ({mode})", pilot.cleanup)


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    result.add_argument("--template-root", type=Path, default=Path("."))
    result.add_argument("--template-sha", required=True)
    result.add_argument("--owner", required=True)
    result.add_argument("--name", required=True, help="new pilot repository name, for example vba-cleanroom-pilot-140")
    result.add_argument("--work-dir", type=Path, required=True)
    result.add_argument("--maintainer", required=True)
    result.add_argument("--project-name", default="Clean-room Pilot")
    result.add_argument("--profile", default="library", choices=("library", "ui-component", "application"))
    result.add_argument("--check", action="append", default=[], help="profile release check NAME=DETAIL")
    result.add_argument("--cleanup", default="delete", choices=("delete", "archive", "keep"))
    result.add_argument("--yes", action="store_true", help="answer yes to the clean-up confirmation")
    return result


def main(arguments: list[str] | None = None) -> int:
    options = parser().parse_args(arguments)
    try:
        return run(options)
    except evidence.EvidenceError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
