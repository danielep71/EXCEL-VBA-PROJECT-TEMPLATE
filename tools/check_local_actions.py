#!/usr/bin/env python3
"""Validate repository-local GitHub Action references and tracked entrypoints."""

from __future__ import annotations

import argparse
import json
from pathlib import Path, PurePosixPath
import re
import subprocess
import sys
import tempfile

TOOL_NAME = "Repository-local GitHub Actions"
WORKFLOW_SUFFIXES = {".yml", ".yaml"}
METADATA_NAMES = ("action.yml", "action.yaml")
USES_RE = re.compile(r"^\s*(?:-\s*)?uses:\s*([^\s#]+)")
TOP_SCALAR_RE = re.compile(r"^([A-Za-z][A-Za-z0-9_-]*):\s*(.*?)\s*$")


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(
        ["git", "-C", str(root), *args], stdout=subprocess.PIPE, stderr=subprocess.PIPE
    )


def tracked_files(root: Path) -> set[str]:
    completed = git(root, "ls-files", "-z")
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    return {
        item.decode("utf-8", errors="surrogateescape")
        for item in completed.stdout.split(b"\0")
        if item
    }


def workflow_paths(files: set[str]) -> list[str]:
    return sorted(
        path for path in files
        if path.startswith(".github/workflows/") and PurePosixPath(path).suffix.casefold() in WORKFLOW_SUFFIXES
    )


def safe_relative(reference: str) -> str | None:
    if not reference.startswith("./"):
        return None
    raw = reference[2:]
    path = PurePosixPath(raw)
    if not raw or path.is_absolute() or any(part in {"", ".", ".."} for part in path.parts):
        return None
    if path.as_posix() != raw.rstrip("/"):
        return None
    return path.as_posix()


def parse_runs_metadata(text: str) -> tuple[str | None, dict[str, str], bool]:
    using: str | None = None
    entrypoints: dict[str, str] = {}
    has_steps = False
    in_runs = False
    runs_indent = 0
    for raw in text.splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        indent = len(raw) - len(raw.lstrip(" "))
        if re.match(r"^runs:\s*$", raw):
            in_runs = True
            runs_indent = indent
            continue
        if in_runs and indent <= runs_indent:
            in_runs = False
        if not in_runs:
            continue
        stripped = raw.strip()
        match = TOP_SCALAR_RE.match(stripped)
        if match:
            key, value = match.groups()
            key = key.casefold()
            value = value.strip('"\'')
            if key == "using":
                using = value
            elif key in {"main", "pre", "post", "image"}:
                entrypoints[key] = value
            elif key == "steps":
                has_steps = True
        elif stripped.startswith("steps:"):
            has_steps = True
    return using, entrypoints, has_steps


def validate_action(
    root: Path,
    files: set[str],
    workflow: str,
    line: int,
    reference: str,
) -> list[dict[str, object]]:
    findings: list[dict[str, object]] = []
    relative = safe_relative(reference)
    if relative is None:
        return [{
            "path": workflow, "line": line, "reference": reference,
            "message": "Local action reference must stay inside the repository and contain no traversal segments.",
        }]
    action_dir = root / PurePosixPath(relative)
    if not action_dir.is_dir():
        return [{
            "path": workflow, "line": line, "reference": reference,
            "message": f"Local action directory does not exist: {relative}",
        }]

    candidates = [f"{relative}/{name}" for name in METADATA_NAMES if (action_dir / name).is_file()]
    if len(candidates) != 1:
        findings.append({
            "path": workflow, "line": line, "reference": reference,
            "message": f"Local action must contain exactly one of action.yml or action.yaml; observed {len(candidates)}.",
        })
        return findings
    metadata = candidates[0]
    if metadata not in files:
        findings.append({
            "path": workflow, "line": line, "reference": reference,
            "message": f"Local action metadata is not tracked: {metadata}",
        })
        return findings

    text = (root / metadata).read_text(encoding="utf-8")
    if not re.search(r"(?m)^name:\s*\S", text):
        findings.append({"path": metadata, "message": "Local action metadata requires a non-empty name."})
    if not re.search(r"(?m)^description:\s*\S", text):
        findings.append({"path": metadata, "message": "Local action metadata requires a non-empty description."})
    using, entrypoints, has_steps = parse_runs_metadata(text)
    if using is None:
        findings.append({"path": metadata, "message": "Local action metadata requires runs.using."})
        return findings

    normalized_using = using.casefold()
    if normalized_using == "composite":
        if not has_steps:
            findings.append({"path": metadata, "message": "Composite local action requires runs.steps."})
    elif normalized_using in {"node20", "node24"}:
        if "main" not in entrypoints:
            findings.append({"path": metadata, "message": f"{using} local action requires runs.main."})
        for key in ("main", "pre", "post"):
            value = entrypoints.get(key)
            if value is None:
                continue
            entry = PurePosixPath(value)
            if entry.is_absolute() or any(part in {"", ".", ".."} for part in entry.parts):
                findings.append({"path": metadata, "message": f"runs.{key} must stay inside the action directory: {value}"})
                continue
            relative_entry = f"{relative}/{entry.as_posix()}"
            if not (root / relative_entry).is_file():
                findings.append({"path": metadata, "message": f"runs.{key} entrypoint does not exist: {relative_entry}"})
            elif relative_entry not in files:
                findings.append({"path": metadata, "message": f"runs.{key} entrypoint is not tracked: {relative_entry}"})
    elif normalized_using == "docker":
        image = entrypoints.get("image")
        if not image:
            findings.append({"path": metadata, "message": "Docker local action requires runs.image."})
        elif image.casefold() != "dockerfile":
            findings.append({"path": metadata, "message": "Reusable baseline supports local Docker actions only with runs.image: Dockerfile."})
        else:
            dockerfile = f"{relative}/Dockerfile"
            if not (root / dockerfile).is_file():
                findings.append({"path": metadata, "message": f"Dockerfile does not exist: {dockerfile}"})
            elif dockerfile not in files:
                findings.append({"path": metadata, "message": f"Dockerfile is not tracked: {dockerfile}"})
    else:
        findings.append({"path": metadata, "message": f"Unsupported local action runs.using value: {using}"})
    return findings


def run_check(root: Path) -> dict[str, object]:
    files = tracked_files(root)
    findings: list[dict[str, object]] = []
    references: list[dict[str, object]] = []
    workflows = workflow_paths(files)
    for workflow in workflows:
        text = (root / workflow).read_text(encoding="utf-8")
        for number, raw in enumerate(text.splitlines(), start=1):
            match = USES_RE.match(raw)
            if not match:
                continue
            reference = match.group(1)
            if not reference.startswith("./"):
                continue
            references.append({"workflow": workflow, "line": number, "reference": reference})
            findings.extend(validate_action(root, files, workflow, number, reference))
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "workflows": len(workflows),
        "local_references": references,
        "findings": findings,
    }


def markdown_report(report: dict[str, object]) -> str:
    lines = [
        "## Repository-local GitHub Actions",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Workflows:** {report['workflows']}",
        f"- **Local references:** {len(report['local_references'])}",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["findings"]:
        lines.extend(["", "### Findings", ""])
        for item in report["findings"]:
            location = str(item.get("path", "."))
            if item.get("line"):
                location += f":{item['line']}"
            lines.append(f"- `{location}` — {item['message']}")
    return "\n".join(lines) + "\n"


def write_text(path: Path | None, text: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def fixture_report(
    workflow_reference: str,
    action_files: dict[str, str] | None,
    tracked: tuple[str, ...],
) -> dict[str, object]:
    with tempfile.TemporaryDirectory(prefix="local-action-") as temporary:
        root = Path(temporary)
        workflow = root / ".github/workflows/test.yml"
        workflow.parent.mkdir(parents=True)
        workflow.write_text(
            "name: Test\non: push\njobs:\n  test:\n    runs-on: ubuntu-latest\n    steps:\n"
            f"      - uses: {workflow_reference}\n",
            encoding="utf-8",
        )
        if action_files:
            for relative, content in action_files.items():
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(content, encoding="utf-8")
        for command in (
            ("init", "-b", "main"),
            ("config", "user.name", "Local Action Self-Test"),
            ("config", "user.email", "local-action@example.invalid"),
        ):
            completed = git(root, *command)
            if completed.returncode != 0:
                raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
        for relative in tracked:
            completed = git(root, "add", relative)
            if completed.returncode != 0:
                raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
        return run_check(root)


def run_self_test() -> int:
    workflow_path = ".github/workflows/test.yml"
    composite = {
        ".github/actions/example/action.yml": "name: Example\ndescription: Composite\nruns:\n  using: composite\n  steps:\n    - shell: bash\n      run: echo ok\n"
    }
    node = {
        ".github/actions/example/action.yml": "name: Example\ndescription: Node\nruns:\n  using: node24\n  main: dist/index.js\n",
        ".github/actions/example/dist/index.js": "console.log('ok')\n",
    }
    cases = {
        "valid-composite": ("pass", "./.github/actions/example", composite, (workflow_path, ".github/actions/example/action.yml")),
        "valid-node": ("pass", "./.github/actions/example", node, (workflow_path, ".github/actions/example/action.yml", ".github/actions/example/dist/index.js")),
        "missing-path": ("fail", "./.github/actions/missing", None, (workflow_path,)),
        "traversal": ("fail", "./.github/actions/../outside", None, (workflow_path,)),
        "untracked-action": ("fail", "./.github/actions/example", composite, (workflow_path,)),
        "dual-metadata": (
            "fail", "./.github/actions/example",
            {**composite, ".github/actions/example/action.yaml": composite[".github/actions/example/action.yml"]},
            (workflow_path, ".github/actions/example/action.yml", ".github/actions/example/action.yaml"),
        ),
        "malformed-metadata": (
            "fail", "./.github/actions/example",
            {".github/actions/example/action.yml": "name: Example\ndescription: Missing runs\n"},
            (workflow_path, ".github/actions/example/action.yml"),
        ),
        "missing-entrypoint": (
            "fail", "./.github/actions/example",
            {".github/actions/example/action.yml": node[".github/actions/example/action.yml"]},
            (workflow_path, ".github/actions/example/action.yml"),
        ),
        "untracked-entrypoint": (
            "fail", "./.github/actions/example", node,
            (workflow_path, ".github/actions/example/action.yml"),
        ),
        "entrypoint-traversal": (
            "fail", "./.github/actions/example",
            {".github/actions/example/action.yml": "name: Example\ndescription: Node\nruns:\n  using: node24\n  main: ../index.js\n"},
            (workflow_path, ".github/actions/example/action.yml"),
        ),
    }
    failures: list[str] = []
    for name, (expected, reference, files, tracked) in cases.items():
        report = fixture_report(reference, files, tracked)
        if report["status"] != expected:
            failures.append(f"{name}: expected {expected}, got {report['status']} ({report['findings']})")
    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: valid composite/node actions, missing paths, traversal, untracked metadata, "
        "dual metadata, malformed metadata, missing/untracked entrypoints, and entrypoint traversal passed."
    )
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    if options.self_test:
        try:
            return run_self_test()
        except (OSError, UnicodeError, RuntimeError, subprocess.SubprocessError) as error:
            print(f"SELF-TEST ERROR: {error}", file=sys.stderr)
            return 2
    try:
        report = run_check(options.root)
        write_text(options.output, json.dumps(report, indent=2, sort_keys=True) + "\n")
        write_text(options.summary, markdown_report(report))
        print(markdown_report(report).rstrip())
        return 0 if report["status"] == "pass" else 1
    except (OSError, UnicodeError, RuntimeError, subprocess.SubprocessError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
