#!/usr/bin/env python3
'''Validate repository-local GitHub Action references and tracked entrypoints.'''

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
USES_RE = re.compile(
    r"^\s*(?:-\s*)?uses:\s*(?:\"([^\"]+)\"|'([^']+)'|([^\s#]+))"
)
TOP_SCALAR_RE = re.compile(r"^([A-Za-z][A-Za-z0-9_-]*):\s*(.*?)\s*$")


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(
        ["git", "-C", str(root), *args],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def tracked_files(root: Path) -> set[str]:
    completed = git(root, "ls-files", "-z")
    if completed.returncode:
        raise RuntimeError(
            completed.stderr.decode("utf-8", errors="replace").strip()
        )
    return {
        item.decode("utf-8", errors="surrogateescape")
        for item in completed.stdout.split(b"\0")
        if item
    }


def workflow_paths(files: set[str]) -> list[str]:
    return sorted(
        path
        for path in files
        if path.startswith(".github/workflows/")
        and PurePosixPath(path).suffix.casefold() in WORKFLOW_SUFFIXES
    )


def uses_reference(raw: str) -> str | None:
    match = USES_RE.match(raw)
    if not match:
        return None
    return next((value for value in match.groups() if value is not None), None)


def safe_relative(reference: str) -> str | None:
    if not reference.startswith("./"):
        return None
    raw = reference[2:]
    path = PurePosixPath(raw)
    if (
        not raw
        or path.is_absolute()
        or any(part in {"", ".", ".."} for part in path.parts)
    ):
        return None
    if path.as_posix() != raw.rstrip("/"):
        return None
    return path.as_posix()


def unquote_scalar(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        value = value[1:-1]
    return value.strip()


def top_scalar(text: str, key: str) -> str | None:
    pattern = re.compile(rf"(?m)^{re.escape(key)}:\s*(.*?)\s*$")
    match = pattern.search(text)
    if not match:
        return None
    value = unquote_scalar(match.group(1))
    return value or None


def parse_runs_metadata(
    text: str,
) -> tuple[str | None, dict[str, str], bool]:
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
            value = unquote_scalar(value)
            if key == "using":
                using = value
            elif key in {"main", "pre", "post", "image"}:
                entrypoints[key] = value
            elif key == "steps":
                has_steps = True
        elif stripped.startswith("steps:"):
            has_steps = True
    return using, entrypoints, has_steps


def finding(
    path: str,
    message: str,
    *,
    line: int | None = None,
    reference: str | None = None,
) -> dict[str, object]:
    item: dict[str, object] = {"path": path, "message": message}
    if line is not None:
        item["line"] = line
    if reference is not None:
        item["reference"] = reference
    return item


def validate_entrypoint(
    root: Path,
    files: set[str],
    metadata: str,
    action_relative: str,
    key: str,
    value: str,
) -> list[dict[str, object]]:
    entry = PurePosixPath(value)
    if (
        not value
        or entry.is_absolute()
        or any(part in {"", ".", ".."} for part in entry.parts)
        or entry.as_posix() != value
    ):
        return [
            finding(
                metadata,
                f"runs.{key} must stay inside the action directory: {value}",
            )
        ]
    relative_entry = f"{action_relative}/{entry.as_posix()}"
    if not (root / relative_entry).is_file():
        return [
            finding(
                metadata,
                f"runs.{key} entrypoint does not exist: {relative_entry}",
            )
        ]
    if relative_entry not in files:
        return [
            finding(
                metadata,
                f"runs.{key} entrypoint is not tracked: {relative_entry}",
            )
        ]
    return []


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
        return [
            finding(
                workflow,
                "Local action reference must stay inside the repository and "
                "contain no traversal segments.",
                line=line,
                reference=reference,
            )
        ]

    action_dir = root / PurePosixPath(relative)
    if not action_dir.is_dir():
        return [
            finding(
                workflow,
                f"Local action directory does not exist: {relative}",
                line=line,
                reference=reference,
            )
        ]

    candidates = [
        f"{relative}/{name}"
        for name in METADATA_NAMES
        if (action_dir / name).is_file()
    ]
    if len(candidates) != 1:
        return [
            finding(
                workflow,
                "Local action must contain exactly one of action.yml or "
                f"action.yaml; observed {len(candidates)}.",
                line=line,
                reference=reference,
            )
        ]

    metadata = candidates[0]
    if metadata not in files:
        return [
            finding(
                workflow,
                f"Local action metadata is not tracked: {metadata}",
                line=line,
                reference=reference,
            )
        ]

    text = (root / metadata).read_text(encoding="utf-8")
    if top_scalar(text, "name") is None:
        findings.append(
            finding(metadata, "Local action metadata requires a non-empty name.")
        )
    if top_scalar(text, "description") is None:
        findings.append(
            finding(
                metadata,
                "Local action metadata requires a non-empty description.",
            )
        )

    using, entrypoints, has_steps = parse_runs_metadata(text)
    if not using:
        findings.append(
            finding(metadata, "Local action metadata requires runs.using.")
        )
        return findings

    normalized = using.casefold()
    if normalized == "composite":
        if not has_steps:
            findings.append(
                finding(metadata, "Composite local action requires runs.steps.")
            )
    elif normalized in {"node20", "node24"}:
        if not entrypoints.get("main"):
            findings.append(
                finding(metadata, f"{using} local action requires runs.main.")
            )
        for key in ("main", "pre", "post"):
            value = entrypoints.get(key)
            if value is not None:
                findings.extend(
                    validate_entrypoint(
                        root, files, metadata, relative, key, value
                    )
                )
    elif normalized == "docker":
        image = entrypoints.get("image")
        if not image:
            findings.append(
                finding(metadata, "Docker local action requires runs.image.")
            )
        elif image.casefold() != "dockerfile":
            findings.append(
                finding(
                    metadata,
                    "Reusable baseline supports local Docker actions only with "
                    "runs.image: Dockerfile.",
                )
            )
        else:
            dockerfile = f"{relative}/Dockerfile"
            if not (root / dockerfile).is_file():
                findings.append(
                    finding(
                        metadata,
                        f"Dockerfile does not exist: {dockerfile}",
                    )
                )
            elif dockerfile not in files:
                findings.append(
                    finding(
                        metadata,
                        f"Dockerfile is not tracked: {dockerfile}",
                    )
                )
    else:
        findings.append(
            finding(
                metadata,
                f"Unsupported local action runs.using value: {using}",
            )
        )
    return findings


def run_check(root: Path) -> dict[str, object]:
    files = tracked_files(root)
    findings: list[dict[str, object]] = []
    references: list[dict[str, object]] = []
    workflows = workflow_paths(files)
    for workflow in workflows:
        text = (root / workflow).read_text(encoding="utf-8")
        for number, raw in enumerate(text.splitlines(), start=1):
            reference = uses_reference(raw)
            if reference is None or not reference.startswith("./"):
                continue
            references.append(
                {
                    "workflow": workflow,
                    "line": number,
                    "reference": reference,
                }
            )
            findings.extend(
                validate_action(
                    root, files, workflow, number, reference
                )
            )
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
            "name: Test\non: push\njobs:\n  test:\n"
            "    runs-on: ubuntu-latest\n    steps:\n"
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
            (
                "config",
                "user.email",
                "local-action@example.invalid",
            ),
        ):
            completed = git(root, *command)
            if completed.returncode:
                raise RuntimeError(
                    completed.stderr.decode(
                        "utf-8", errors="replace"
                    ).strip()
                )
        for relative in tracked:
            completed = git(root, "add", relative)
            if completed.returncode:
                raise RuntimeError(
                    completed.stderr.decode(
                        "utf-8", errors="replace"
                    ).strip()
                )
        return run_check(root)


def run_self_test() -> int:
    workflow_path = ".github/workflows/test.yml"
    composite = {
        ".github/actions/example/action.yml": (
            "name: Example\n"
            "description: Composite\n"
            "runs:\n"
            "  using: composite\n"
            "  steps:\n"
            "    - shell: bash\n"
            "      run: echo ok\n"
        )
    }
    node = {
        ".github/actions/example/action.yml": (
            "name: Example\n"
            "description: Node\n"
            "runs:\n"
            "  using: node24\n"
            "  main: dist/index.js\n"
        ),
        ".github/actions/example/dist/index.js": (
            "console.log('ok')\n"
        ),
    }
    tracked_composite = (
        workflow_path,
        ".github/actions/example/action.yml",
    )
    tracked_node = (
        workflow_path,
        ".github/actions/example/action.yml",
        ".github/actions/example/dist/index.js",
    )
    cases = {
        "valid-composite": (
            "pass",
            "./.github/actions/example",
            composite,
            tracked_composite,
        ),
        "valid-double-quoted": (
            "pass",
            '"./.github/actions/example"',
            composite,
            tracked_composite,
        ),
        "valid-single-quoted": (
            "pass",
            "'./.github/actions/example'",
            composite,
            tracked_composite,
        ),
        "valid-node": (
            "pass",
            "./.github/actions/example",
            node,
            tracked_node,
        ),
        "missing-path": (
            "fail",
            "./.github/actions/missing",
            None,
            (workflow_path,),
        ),
        "traversal": (
            "fail",
            "./.github/actions/../outside",
            None,
            (workflow_path,),
        ),
        "untracked-action": (
            "fail",
            "./.github/actions/example",
            composite,
            (workflow_path,),
        ),
        "dual-metadata": (
            "fail",
            "./.github/actions/example",
            {
                **composite,
                ".github/actions/example/action.yaml": composite[
                    ".github/actions/example/action.yml"
                ],
            },
            (
                workflow_path,
                ".github/actions/example/action.yml",
                ".github/actions/example/action.yaml",
            ),
        ),
        "empty-name": (
            "fail",
            "./.github/actions/example",
            {
                ".github/actions/example/action.yml": (
                    'name: ""\n'
                    "description: Composite\n"
                    "runs:\n  using: composite\n  steps:\n"
                    "    - shell: bash\n      run: echo ok\n"
                )
            },
            tracked_composite,
        ),
        "empty-description": (
            "fail",
            "./.github/actions/example",
            {
                ".github/actions/example/action.yml": (
                    "name: Example\n"
                    "description: ''\n"
                    "runs:\n  using: composite\n  steps:\n"
                    "    - shell: bash\n      run: echo ok\n"
                )
            },
            tracked_composite,
        ),
        "malformed-metadata": (
            "fail",
            "./.github/actions/example",
            {
                ".github/actions/example/action.yml": (
                    "name: Example\n"
                    "description: Missing runs\n"
                )
            },
            tracked_composite,
        ),
        "missing-entrypoint": (
            "fail",
            "./.github/actions/example",
            {
                ".github/actions/example/action.yml": node[
                    ".github/actions/example/action.yml"
                ]
            },
            tracked_composite,
        ),
        "untracked-entrypoint": (
            "fail",
            "./.github/actions/example",
            node,
            tracked_composite,
        ),
        "entrypoint-traversal": (
            "fail",
            "./.github/actions/example",
            {
                ".github/actions/example/action.yml": (
                    "name: Example\n"
                    "description: Node\n"
                    "runs:\n"
                    "  using: node24\n"
                    "  main: ../index.js\n"
                )
            },
            tracked_composite,
        ),
    }

    failures: list[str] = []
    for name, (expected, reference, files, tracked) in cases.items():
        report = fixture_report(reference, files, tracked)
        if report["status"] != expected:
            failures.append(
                f"{name}: expected {expected}, got "
                f"{report['status']} ({report['findings']})"
            )
    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1

    print(
        "SELF-TEST PASS: quoted/unquoted local references, valid "
        "composite/node actions, missing paths, traversal, tracked "
        "metadata, dual metadata, empty metadata scalars, missing/untracked "
        "entrypoints, and entrypoint traversal passed."
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
    try:
        if options.self_test:
            return run_self_test()
        report = run_check(options.root)
        write_text(
            options.output,
            json.dumps(report, indent=2, sort_keys=True) + "\n",
        )
        write_text(options.summary, markdown_report(report))
        print(markdown_report(report).rstrip())
        return 0 if report["status"] == "pass" else 1
    except (
        OSError,
        UnicodeError,
        RuntimeError,
        subprocess.SubprocessError,
    ) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
