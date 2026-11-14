#!/usr/bin/env python3
"""Release the canonical template with one command, resuming where a previous run stopped.

Run it on the maintainer's Windows machine with desktop Excel, from a clean
checkout of the template's ``main``. It needs ``gh`` signed in, Wiki push
credentials and the SSH signing key. It chains the release tools and records
progress in ``<work-dir>/release-state.json``:

1. ``cut``: ``cut_release.py --push``, then merges the cut pull request once its
   checks pass. The merged commit is the candidate.
2. ``interface``: stops if the called *Static repository checks* workflow differs
   from the last published reusable-workflow pin. A new interface revision is
   published first; see RELEASING.md.
3. ``excel``: ``Certify-Excel.ps1`` on the candidate.
4. ``wiki``: ``publish_wiki.py``, then the maintainer confirms the browser
   review.
5. ``pilot``: ``run_clean_room_pilot.py``, unless ``--pilot-record`` names a
   finished journey.
6. ``links``: dispatches *External documentation links* and waits for a green
   run at the candidate.
7. ``evidence``: ``prepare_release_evidence.py prepare``, then the maintainer
   confirms that the evidence was reviewed for private data.
8. ``publish``: ``publish_release.py``, which signs, publishes and closes out.

A finished phase is skipped on the next run, provided the checkout is still at
the recorded candidate. The interactive tools keep their own confirmations, so
the maintainer approves every push and publication. Nothing is signed or
published without one.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

import check_excel_evidence as excel
import prepare_release_evidence as evidence
import publish_release
import run_clean_room_pilot as journey_tool

STATIC = ".github/workflows/static-checks.yml"
PHASES = ("cut", "interface", "excel", "wiki", "pilot", "links", "evidence", "publish")


class Release:
    def __init__(self, options: argparse.Namespace) -> None:
        self.options = options
        self.root = options.root.resolve()
        self.work = options.work_dir.resolve()
        self.version = options.version
        self.tag = f"v{options.version}"
        self.state_path = self.work / "release-state.json"
        profile = json.loads((self.root / ".github/repository-profile.json").read_text(encoding="utf-8"))
        evidence.require(profile.get("mode") == "template", "release_template.py releases the canonical template only")
        self.repository = str(profile["repository"])
        self.state: dict[str, Any] = {"version": self.version, "candidate": None, "done": []}
        if self.state_path.is_file():
            self.state = json.loads(self.state_path.read_text(encoding="utf-8"))
            evidence.require(self.state.get("version") == self.version,
                             f"{self.work} belongs to v{self.state.get('version')}: use another --work-dir")

    def save(self) -> None:
        evidence.write_json(self.state_path, self.state)

    def head(self) -> str:
        return evidence.checked(["git", "rev-parse", "HEAD"], self.root, "git rev-parse").decode().strip()

    def python(self, script: str, *arguments: str) -> list[str]:
        return [sys.executable, str(self.root / "tools" / script), *arguments]

    def interactive(self, command: list[str], what: str) -> None:
        code = journey_tool.stream(command, self.root)
        evidence.require(code == 0, f"{what} stopped (exit {code}); rerun the same command after fixing it")

    def phase_cut(self) -> str:
        version = (self.root / "VERSION").read_text(encoding="utf-8").strip()
        if version != self.version:
            status = evidence.checked(["git", "status", "--porcelain"], self.root, "git status").decode().strip()
            evidence.require(not status, "the checkout has uncommitted changes")
            evidence.checked(["git", "fetch", "origin", "main"], self.root, "fetch main")
            origin = evidence.checked(["git", "rev-parse", "origin/main"], self.root, "rev-parse").decode().strip()
            evidence.require(self.head() == origin, "start from the current main: the checkout is not at origin/main")
            self.interactive(self.python("cut_release.py", "--root", ".", "--version", self.version, "--push"),
                             "the release cut")
        branch = f"chore/cut-v{self.version}"
        if evidence.checked(["git", "branch", "--show-current"], self.root, "branch").decode().strip() == branch:
            publish_release.confirm(f"Merge the {branch} pull request into main once its checks pass?",
                                    self.options.yes)
            journey_tool.merge_pull_request(self.root, self.root, self.repository, branch)
        candidate = self.head()
        self.state["candidate"] = candidate
        return f"candidate {candidate}"

    def phase_interface(self) -> str:
        text = (self.root / "docs/REUSABLE_WORKFLOWS.md").read_text(encoding="utf-8")
        pins = re.findall(r"^\| `(\d+\.\d+\.\d+)` \| `([0-9a-f]{40})` \|", text, re.MULTILINE)
        evidence.require(bool(pins), "docs/REUSABLE_WORKFLOWS.md lists no published interface pin")
        revision, provider = pins[-1]
        if evidence.run(["git", "cat-file", "-e", f"{provider}^{{commit}}"], self.root).returncode != 0:
            evidence.checked(["git", "fetch", "origin", provider], self.root, f"fetch the {revision} provider")
        same = evidence.run(["git", "diff", "--quiet", provider, self.state["candidate"], "--", STATIC], self.root)
        evidence.require(same.returncode == 0,
                         f"{STATIC} changed since interface {revision} ({provider[:7]}): publish the next "
                         "interface revision with its hosted fixtures first (RELEASING.md, canonical-template "
                         "release certification), then rerun")
        return f"called workflow identical to interface {revision} provider {provider[:7]}"

    def phase_excel(self) -> str:
        bundle = self.work / f"excel-evidence-{self.state['candidate'][:7]}"
        host = bundle / "host.json"
        if not host.is_file():
            evidence.require(journey_tool.on_windows(), "Excel certification needs Windows with desktop Excel")
            self.interactive(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                              str(self.root / "tools/Certify-Excel.ps1"), "-RepoRoot", str(self.root),
                              "-BundleDir", str(bundle)], "Certify-Excel")
        report = excel.evaluate(self.root, self.state["candidate"], host)
        evidence.require(report["status"] == "pass", "the Excel record does not validate: " + ", ".join(report["outcomes"]))
        self.state["excel_bundle"] = str(bundle)
        return f"{bundle.name}: host record PASS"

    def phase_wiki(self) -> str:
        self.interactive(self.python("publish_wiki.py", "--root", ".", "--candidate-sha", self.state["candidate"],
                                     "--work-dir", str(self.work / "wiki-work")), "Wiki publication")
        publish_release.confirm("Open the Wiki in a browser now. Do Home, the sidebar and every page link render "
                                "and navigate as intended?", False)
        self.state["wiki_browser_reviewed"] = True
        return "Wiki published, read back and reviewed in a browser"

    def phase_pilot(self) -> str:
        record = self.options.pilot_record
        if record is None:
            work = self.work / "pilot-work"
            name = f"vba-cleanroom-pilot-{self.version.replace('.', '')}"
            self.interactive(self.python("run_clean_room_pilot.py", "--template-root", ".", "--template-sha",
                                         self.state["candidate"], "--owner", self.repository.split("/")[0],
                                         "--name", name, "--work-dir", str(work),
                                         "--maintainer", self.options.maintainer), "the clean-room pilot")
            record = work / "clean-room-journey.md"
        data = json.loads(record.with_suffix(".json").read_text(encoding="utf-8"))
        evidence.require(data.get("template_sha") == self.state["candidate"],
                         "the clean-room record belongs to another candidate")
        evidence.require(all(row.get("result") == "PASS" for row in data.get("steps", [])) and data.get("steps"),
                         "the clean-room record has a failed or missing step")
        self.state["pilot_record"] = str(record)
        self.state["pilot_summary"] = (f"Live clean-room pilot {data['pilot']} ({data['profile']}) completed "
                                       f"{len(data['steps'])} steps from the template to a closed-out v0.1.0 release; "
                                       f"record retained as {record.name}.")
        return str(self.state["pilot_summary"])

    def phase_links(self) -> str:
        endpoint = (f"repos/{self.repository}/actions/workflows/external-links.yml/runs"
                    f"?head_sha={self.state['candidate']}&per_page=20")

        def green() -> list[dict[str, Any]]:
            return [run for run in evidence.gh_json(self.root, endpoint).get("workflow_runs", [])
                    if run.get("status") == "completed" and run.get("conclusion") == "success"]

        if not green():
            evidence.checked(["gh", "workflow", "run", "external-links.yml", "--repo", self.repository, "--ref",
                              "main"], self.root, "external-link dispatch")
            journey_tool.wait_for(lambda: bool(green()), "a green external-link run at the candidate")
        return f"external-link run {green()[0]['html_url']} PASS"

    def phase_evidence(self) -> str:
        work = self.work / "release-work"
        if not (work / "release-evidence.json").is_file():
            self.interactive(self.python("prepare_release_evidence.py", "prepare", "--root", ".", "--tag", self.tag,
                                         "--candidate-sha", self.state["candidate"], "--work-dir", str(work),
                                         "--excel-bundle", self.state["excel_bundle"], "--wiki-browser-reviewed",
                                         "--pilot-record", self.state["pilot_record"],
                                         "--pilot-summary", self.state["pilot_summary"]), "release evidence")
        files = sorted(path.name for path in (work / "evidence").iterdir() if path.is_file())
        print("Evidence files for the public certification bundle:\n  " + "\n  ".join(files))
        publish_release.confirm(f"Have you reviewed every file under {work / 'evidence'} for private data?", False)
        self.state["redaction_reviewed"] = True
        return f"{len(files)} evidence files prepared and reviewed for private data"

    def phase_publish(self) -> str:
        key = publish_release.signing_key(self.options.signing_key)
        self.interactive(self.python("publish_release.py", "--root", ".", "--tag", self.tag, "--candidate-sha",
                                     self.state["candidate"], "--work-dir", str(self.work / "release-work"),
                                     "--signing-key", str(key), "--redaction-reviewed", "--wiki-browser-reviewed"),
                         "publication")
        return f"{self.tag} signed, published and closed out"

    def run(self) -> int:
        self.work.mkdir(parents=True, exist_ok=True)
        for phase in PHASES:
            if phase in self.state["done"]:
                evidence.require(self.head() == self.state["candidate"],
                                 f"the checkout moved from the recorded candidate {self.state['candidate']}")
                print(f"SKIP {phase}: done")
                continue
            print(f"== {phase}")
            result = getattr(self, f"phase_{phase}")()
            self.state["done"].append(phase)
            self.state.setdefault("results", {})[phase] = result
            self.save()
            print(f"DONE {phase}: {result}")
        print(f"RELEASED: {self.tag} at {self.state['candidate']}")
        return 0


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    result.add_argument("--root", type=Path, default=Path("."))
    result.add_argument("--version", required=True)
    result.add_argument("--work-dir", type=Path, required=True)
    result.add_argument("--maintainer", required=True, help="maintainer name for the clean-room pilot")
    result.add_argument("--signing-key", type=Path, help="SSH signing key; default RELEASE_SIGNING_KEY")
    result.add_argument("--pilot-record", type=Path, help="a finished clean-room-journey.md to reuse")
    result.add_argument("--yes", action="store_true", help="merge the cut without asking")
    return result


def main(arguments: list[str] | None = None) -> int:
    options = parser().parse_args(arguments)
    try:
        return Release(options).run()
    except evidence.EvidenceError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
