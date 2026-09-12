from pathlib import Path
import textwrap

path = Path("tools/check_repo.py")
text = path.read_text(encoding="utf-8")
start = text.index("def check_workflow_actions(\n")
end = text.index("\ndef check_version_changelog(\n", start)
replacement = textwrap.dedent(r'''
WORKFLOW_WRITE_PERMISSION_RE = re.compile(
    r"(?:actions|attestations|checks|contents|deployments|discussions|id-token|issues|models|packages|pages|pull-requests|repository-projects|security-events|statuses):\s*write(?:\s*#.*)?"
)


def _workflow_has_event(lines: list[str], event: str) -> bool:
    block = re.compile(rf"^  {re.escape(event)}:\s*(?:#.*)?$")
    flow = re.compile(rf"^on:\s*\[[^]]*\b{re.escape(event)}\b[^]]*\]\s*$")
    return any(block.match(line) or flow.match(line) for line in lines)


def _check_pull_request_target(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> None:
    block = re.compile(r"^  pull_request_target:\s*(?:#.*)?$")
    flow = re.compile(r"^on:\s*\[[^]]*\bpull_request_target\b[^]]*\]\s*$")
    for number, line in enumerate(lines, start=1):
        if block.match(line) or flow.match(line):
            failures.append(
                finding(
                    path,
                    "pull_request_target is prohibited; untrusted pull requests must not execute privileged repository code.",
                    number,
                )
            )


def _check_workflow_level_pr_permissions(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> None:
    in_permissions = False
    for number, line in enumerate(lines, start=1):
        indent = len(line) - len(line.lstrip(" "))
        stripped = line.strip()
        if indent == 0 and re.fullmatch(
            r"permissions:\s*write-all(?:\s*#.*)?", stripped
        ):
            failures.append(
                finding(
                    path,
                    "Workflow triggered by pull_request must not request write-capable token permissions.",
                    number,
                )
            )
        if indent == 0 and re.fullmatch(r"permissions:\s*(?:#.*)?", stripped):
            in_permissions = True
            continue
        if in_permissions and stripped and not stripped.startswith("#") and indent == 0:
            in_permissions = False
        if in_permissions and WORKFLOW_WRITE_PERMISSION_RE.fullmatch(stripped):
            failures.append(
                finding(
                    path,
                    "Workflow triggered by pull_request must not request workflow-level write permissions.",
                    number,
                )
            )


def _workflow_job_blocks(lines: list[str]) -> list[tuple[int, list[str]]]:
    jobs_index = next(
        (
            index
            for index, line in enumerate(lines)
            if re.fullmatch(r"jobs:\s*(?:#.*)?", line.strip())
            and not line.startswith(" ")
        ),
        None,
    )
    if jobs_index is None:
        return []
    starts = [
        index
        for index in range(jobs_index + 1, len(lines))
        if re.match(r"^  [A-Za-z0-9_-]+:\s*(?:#.*)?$", lines[index])
    ]
    return [
        (
            start,
            lines[
                start : starts[position + 1]
                if position + 1 < len(starts)
                else len(lines)
            ],
        )
        for position, start in enumerate(starts)
    ]


def _job_excludes_pull_request(block: list[str]) -> bool:
    exclusion = re.compile(
        r"^    if:\s*github\.event_name\s*!=\s*(['\"])pull_request\1\s*(?:#.*)?$"
    )
    return any(exclusion.match(line) for line in block)


def _check_job_pr_permissions(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> None:
    for start, block in _workflow_job_blocks(lines):
        if _job_excludes_pull_request(block):
            continue
        in_permissions = False
        for offset, item in enumerate(block):
            number = start + offset + 1
            indent = len(item) - len(item.lstrip(" "))
            stripped = item.strip()
            if indent == 4 and re.fullmatch(
                r"permissions:\s*write-all(?:\s*#.*)?", stripped
            ):
                failures.append(
                    finding(
                        path,
                        "Job reachable from pull_request must not request write-all permissions.",
                        number,
                    )
                )
            if indent == 4 and re.fullmatch(r"permissions:\s*(?:#.*)?", stripped):
                in_permissions = True
                continue
            if in_permissions and stripped and not stripped.startswith("#") and indent <= 4:
                in_permissions = False
            if in_permissions and WORKFLOW_WRITE_PERMISSION_RE.fullmatch(stripped):
                failures.append(
                    finding(
                        path,
                        "Job reachable from pull_request must not request write-capable token permissions.",
                        number,
                    )
                )


def _check_pr_workflow_permissions(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> None:
    _check_pull_request_target(path, lines, failures)
    if not _workflow_has_event(lines, "pull_request"):
        return
    _check_workflow_level_pr_permissions(path, lines, failures)
    _check_job_pr_permissions(path, lines, failures)


def _check_external_action_references(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> int:
    uses_line = re.compile(
        r"^\s*(?:-\s*)?uses:\s*([^\s#]+)(?:\s+#\s*(.+?))?\s*$"
    )
    full_sha = re.compile(r"^[0-9a-f]{40}$")
    version_comment = re.compile(r"^v\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$")
    checked = 0
    for number, line in enumerate(lines, start=1):
        if "uses:" not in line:
            continue
        match = uses_line.match(line)
        if not match:
            failures.append(finding(path, "Action reference cannot be parsed.", number))
            continue
        reference, comment = match.groups()
        if reference.startswith("./"):
            continue
        checked += 1
        if "@" not in reference:
            failures.append(
                finding(path, "External action must include a revision.", number)
            )
            continue
        action, revision = reference.rsplit("@", 1)
        if action.startswith("docker://") or not full_sha.fullmatch(revision):
            failures.append(
                finding(
                    path,
                    "External action must be pinned to a full lowercase 40-character commit SHA.",
                    number,
                )
            )
        if not comment or not version_comment.fullmatch(comment.strip()):
            failures.append(
                finding(
                    path,
                    "Pinned action must include an audited semantic-version comment.",
                    number,
                )
            )
    return checked


def check_workflow_actions(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    checked = 0
    for path in repo.files:
        pure = PurePosixPath(path)
        if (
            not path.startswith(".github/workflows/")
            or pure.suffix.casefold() not in {".yml", ".yaml"}
        ):
            continue
        try:
            lines = repo.text(path).splitlines()
        except (OSError, UnicodeError):
            continue
        _check_pr_workflow_permissions(path, lines, failures)
        checked += _check_external_action_references(path, lines, failures)
    return rule_result(
        "workflow-actions",
        "Immutable workflow actions",
        failures,
        f"Validated {checked} external action references",
    )
''').lstrip()
path.write_text(text[:start] + replacement + text[end:], encoding="utf-8")
