#!/usr/bin/env python3
"""Post-transform refactor for residual review-batch complexity hotspots.

Temporary migration helper; removed after validated permanent changes land.
"""

from __future__ import annotations

import ast
from pathlib import Path

ROOT = Path(".")


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def write(path: str, text: str) -> None:
    (ROOT / path).write_text(text, encoding="utf-8", newline="\n")


def replace_once(text: str, before: str, after: str, label: str) -> str:
    count = text.count(before)
    if count != 1:
        raise RuntimeError(f"{label}: expected one anchor, found {count}")
    return text.replace(before, after, 1)


def insert_before(text: str, marker: str, addition: str, label: str) -> str:
    if marker not in text:
        raise RuntimeError(f"{label}: insertion marker missing")
    return text.replace(marker, addition.rstrip() + "\n\n\n" + marker, 1)


def refactor_release_assets() -> None:
    path = "tools/check_release.py"
    text = read(path)
    helper = '''def _validate_release_asset_record(
    root: Path,
    asset: object,
    index: int,
    allowed: list[str],
    candidate_sha: str,
    evidence_entries: dict[str, str],
    findings: list[dict[str, str]],
) -> None:
    item_path = f"evidence.assets[{index}]"
    if not isinstance(asset, dict) or set(asset) != ASSET_KEYS:
        findings.append(
            _finding("invalid-release-asset", item_path, "asset record has an invalid shape")
        )
        return
    relative = asset.get("path")
    digest = asset.get("sha256")
    if not isinstance(relative, str) or not _safe_relative(relative):
        findings.append(
            _finding(
                "invalid-release-asset",
                item_path,
                "path must be safe and repository-relative",
            )
        )
        return
    if relative in evidence_entries:
        findings.append(
            _finding("invalid-release-asset", item_path, f"duplicate asset {relative}")
        )
        return
    if not isinstance(digest, str) or SHA256_PATTERN.fullmatch(digest) is None:
        findings.append(
            _finding(
                "invalid-release-asset",
                item_path,
                "sha256 must be 64 lowercase hexadecimal characters",
            )
        )
        return

    evidence_entries[relative] = digest
    if asset.get("candidate_sha") != candidate_sha:
        findings.append(
            _finding(
                "asset-sha-binding-mismatch",
                item_path,
                "candidate_sha does not match the release candidate",
            )
        )
    if asset.get("package_test") != "PASS":
        findings.append(
            _finding("failed-package-test", item_path, "package_test must be PASS")
        )
    if not any(PurePosixPath(relative).match(pattern) for pattern in allowed):
        findings.append(
            _finding(
                "unapproved-binary",
                relative,
                "asset path is not approved for this profile",
            )
        )

    target = root / relative
    try:
        resolved = target.resolve(strict=True)
        resolved.relative_to(root.resolve())
    except (OSError, ValueError):
        findings.append(
            _finding(
                "missing-release-asset",
                relative,
                "asset is missing or escapes the repository root",
            )
        )
        return
    if not resolved.is_file():
        findings.append(
            _finding("missing-release-asset", relative, "asset is not a regular file")
        )
        return
    actual = _sha256(resolved)
    if actual != digest:
        findings.append(
            _finding(
                "asset-digest-mismatch",
                relative,
                f"actual SHA-256 is {actual}",
            )
        )
'''
    text = insert_before(text, "def _validate_release_assets(\n", helper, path)
    start = text.index("    for index, asset in enumerate(assets):\n", text.index("def _validate_release_assets("))
    end = text.index("    if manifest != evidence_entries:\n", start)
    replacement = '''    for index, asset in enumerate(assets):
        _validate_release_asset_record(
            root,
            asset,
            index,
            allowed,
            candidate_sha,
            evidence_entries,
            findings,
        )
'''
    text = text[:start] + replacement + text[end:]
    write(path, text)


def refactor_release_semantics() -> None:
    path = "tools/check_release_semantics.py"
    text = read(path)
    helper = '''def _validate_comparison_links(
    changelog: str,
    repository: str,
    releases: list[dict[str, Any]],
    findings: list[dict[str, Any]],
) -> None:
    links: dict[str, tuple[str, int]] = {}
    for number, raw in enumerate(changelog.splitlines(), start=1):
        match = LINK_RE.match(raw)
        if not match:
            continue
        name, url = match.groups()
        if name in links:
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "line": number,
                    "message": f"duplicate comparison-link definition for [{name}]",
                }
            )
        else:
            links[name] = (url, number)

    if not releases:
        return

    latest = releases[0]["version"]
    expected_unreleased = f"https://github.com/{repository}/compare/v{latest}...HEAD"
    actual_unreleased = links.get("Unreleased")
    if actual_unreleased is None:
        findings.append(
            {
                "path": "CHANGELOG.md",
                "message": (
                    f"missing [Unreleased] comparison link; expected {expected_unreleased}"
                ),
            }
        )
    elif actual_unreleased[0] != expected_unreleased:
        findings.append(
            {
                "path": "CHANGELOG.md",
                "line": actual_unreleased[1],
                "message": (
                    f"[Unreleased] comparison link must be {expected_unreleased}; "
                    f"observed {actual_unreleased[0]}"
                ),
            }
        )

    for index, release in enumerate(releases):
        value = release["version"]
        if index + 1 < len(releases):
            older = releases[index + 1]["version"]
            expected = f"https://github.com/{repository}/compare/v{older}...v{value}"
        else:
            expected = f"https://github.com/{repository}/releases/tag/v{value}"
        actual = links.get(value)
        if actual is None:
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "message": (
                        f"missing [{value}] release comparison link; expected {expected}"
                    ),
                }
            )
        elif actual[0] != expected:
            findings.append(
                {
                    "path": "CHANGELOG.md",
                    "line": actual[1],
                    "message": f"[{value}] link must be {expected}; observed {actual[0]}",
                }
            )
'''
    text = insert_before(text, "def analyze(version: str, changelog: str, repository: str) -> dict[str, Any]:\n", helper, path)
    start = text.index("    links: dict[str, tuple[str, int]] = {}\n", text.index("def analyze("))
    end = text.index("    release_evidence = [\n", start)
    text = text[:start] + "    _validate_comparison_links(changelog, repository, releases, findings)\n\n" + text[end:]
    write(path, text)


def refactor_placeholder_catalogue() -> None:
    path = "tools/check_repo.py"
    text = read(path)
    target_marker = "def _validate_configuration_placeholders(document: dict[str, Any], failures: list[dict[str, Any]]) -> None:\n"
    helper = '''def _validate_placeholder_catalogue(
    placeholders: dict[str, Any], failures: list[dict[str, Any]]
) -> None:
    catalogue = placeholders.get("catalogue")
    categories_seen: set[str] = set()
    if not isinstance(catalogue, dict) or not catalogue:
        failures.append(
            finding(
                CONFIG_PATH,
                "placeholders.catalogue must be a non-empty object.",
            )
        )
        return

    names = list(catalogue)
    if names != sorted(names, key=lambda item: (item.casefold(), item)):
        failures.append(
            finding(
                CONFIG_PATH,
                "placeholders.catalogue keys must be sorted case-insensitively.",
            )
        )
    for name, specification in catalogue.items():
        field = f"placeholders.catalogue.{name}"
        if not isinstance(name, str) or not PLACEHOLDER_NAME_PATTERN.fullmatch(name):
            failures.append(finding(CONFIG_PATH, f"{field} is not a canonical placeholder name."))
            continue
        if not isinstance(specification, dict):
            failures.append(finding(CONFIG_PATH, f"{field} must be an object."))
            continue
        category = specification.get("category")
        description = specification.get("description")
        if category not in PLACEHOLDER_CATEGORIES:
            failures.append(
                finding(
                    CONFIG_PATH,
                    f"{field}.category must be optional, profile-specific, repeatable, or required.",
                )
            )
            continue
        categories_seen.add(category)
        if not isinstance(description, str) or not description.strip():
            failures.append(finding(CONFIG_PATH, f"{field}.description must be non-empty."))
        expected_keys = {"category", "description"}
        if category == "profile-specific":
            expected_keys.add("values")
            values = specification.get("values")
            if not isinstance(values, dict) or set(values) != set(SUPPORTED_PROFILES):
                failures.append(
                    finding(
                        CONFIG_PATH,
                        f"{field}.values must cover exactly all supported profiles.",
                    )
                )
            elif any(
                not isinstance(value, str) or not value.strip()
                for value in values.values()
            ):
                failures.append(
                    finding(CONFIG_PATH, f"{field}.values must all be non-empty strings.")
                )
        elif category == "repeatable":
            expected_keys.add("item_format")
            item_format = specification.get("item_format")
            if not isinstance(item_format, str) or item_format.count("{value}") != 1:
                failures.append(
                    finding(
                        CONFIG_PATH,
                        f"{field}.item_format must contain one {{value}} field.",
                    )
                )
        if set(specification) != expected_keys:
            failures.append(
                finding(
                    CONFIG_PATH,
                    f"{field} has fields inconsistent with its category.",
                )
            )

    missing_categories = set(PLACEHOLDER_CATEGORIES) - categories_seen
    if missing_categories:
        failures.append(
            finding(
                CONFIG_PATH,
                "placeholders.catalogue does not exercise categories: "
                + ", ".join(sorted(missing_categories)),
            )
        )
'''
    text = insert_before(text, target_marker, helper, path)
    fn_start = text.index(target_marker)
    start = text.index("    catalogue = placeholders.get(\"catalogue\")\n", fn_start)
    end = text.index("    block_markers = placeholders.get(\"block_markers\")\n", start)
    text = text[:start] + "    _validate_placeholder_catalogue(placeholders, failures)\n" + text[end:]
    write(path, text)


def refactor_issue_forms() -> None:
    path = "tools/check_repo.py"
    text = read(path)
    helper = '''def _validate_issue_form(
    repo: Repository,
    filename: str,
    specification: dict[str, Any],
    label_names: set[object],
    failures: list[dict[str, Any]],
) -> None:
    path = f"{ISSUE_TEMPLATE_DIRECTORY}/{filename}"
    try:
        text = repo.text(path)
    except (OSError, UnicodeError) as error:
        failures.append(finding(path, f"Cannot read canonical issue form: {error}"))
        return

    for key in ("name", "description"):
        value = _yaml_header_scalar(text, key)
        if value is None or not value.strip():
            failures.append(finding(path, f"Top-level {key} must be non-empty."))
    if _yaml_header_scalar(text, "title") != specification["title"]:
        failures.append(
            finding(path, f"Top-level title must be {specification['title']!r}.")
        )

    labels = _yaml_flow_array(text, "labels")
    expected_label = specification["label"]
    if labels != [expected_label]:
        failures.append(
            finding(
                path,
                f"Top-level labels must be the JSON flow array [{expected_label!r}].",
            )
        )
    elif expected_label not in label_names:
        failures.append(
            finding(
                path,
                f"Issue form label is absent from {LABEL_MANIFEST_PATH}: {expected_label}",
            )
        )
    if _yaml_flow_array(text, "assignees") != []:
        failures.append(
            finding(path, "Top-level assignees must be an empty JSON flow array.")
        )

    blocks = _issue_form_blocks(text)
    if not blocks or len(blocks) > 10:
        failures.append(
            finding(path, "Issue form body must contain between 1 and 10 elements.")
        )
    invalid_types = sorted(
        {kind for kind, _ in blocks}
        - {"checkboxes", "dropdown", "input", "markdown", "textarea"}
    )
    if invalid_types:
        failures.append(
            finding(
                path,
                "Unsupported issue-form element types: " + ", ".join(invalid_types),
            )
        )
    identifiers = [identifier for _, identifier in blocks if identifier is not None]
    if len(identifiers) != len(set(identifiers)):
        failures.append(finding(path, "Issue-form element IDs must be unique."))
    missing = sorted(set(specification["required_ids"]) - set(identifiers))
    if missing:
        failures.append(
            finding(
                path,
                "Required issue-form element IDs are missing: " + ", ".join(missing),
            )
        )
    if "SECURITY.md" not in text or "private" not in text.casefold():
        failures.append(
            finding(
                path,
                "The opening guidance must route vulnerability details to SECURITY.md privately.",
            )
        )
    for identifier in set(specification["required_ids"]):
        if identifier == "acknowledgements":
            continue
        pattern = rf"(?ms)^    id:\s*{re.escape(identifier)}\s*$.*?(?=^  - type:|\Z)"
        match = re.search(pattern, text)
        if match and not re.search(
            r"(?m)^      required:\s*true\s*$", match.group(0)
        ):
            failures.append(
                finding(
                    path,
                    f"Required evidence field {identifier!r} must be mandatory.",
                )
            )
'''
    text = insert_before(text, "def check_issue_forms(\n", helper, path)
    fn_start = text.index("def check_issue_forms(\n")
    start = text.index("    for filename, specification in ISSUE_FORM_SPECS.items():\n", fn_start)
    end = text.index("    config_path = f\"{ISSUE_TEMPLATE_DIRECTORY}/config.yml\"\n", start)
    replacement = '''    for filename, specification in ISSUE_FORM_SPECS.items():
        _validate_issue_form(repo, filename, specification, label_names, failures)

'''
    text = text[:start] + replacement + text[end:]
    write(path, text)


def refactor_vba_structure() -> None:
    path = "tools/check_repo.py"
    text = read(path)
    helper = '''def _handle_vba_structure_directive(
    path: str,
    number: int,
    upper: str,
    directives: list[dict[str, bool]],
    failures: list[dict[str, Any]],
) -> bool:
    if upper.startswith("#IF "):
        condition = upper[4:]
        requires = "VBA7" in condition and "NOT VBA7" not in condition
        directives.append({"vba7": requires, "active": requires})
        return True
    if upper.startswith("#ELSEIF "):
        if not directives:
            failures.append(finding(path, "#ElseIf without #If.", number))
        else:
            condition = upper[8:]
            directives[-1]["active"] = (
                "VBA7" in condition and "NOT VBA7" not in condition
            )
        return True
    if upper.startswith("#ELSE"):
        if not directives:
            failures.append(finding(path, "#Else without #If.", number))
        else:
            directives[-1]["active"] = not directives[-1]["vba7"]
        return True
    if upper.startswith("#END IF"):
        if not directives:
            failures.append(finding(path, "#End If without #If.", number))
        else:
            directives.pop()
        return True
    return False
'''
    text = insert_before(text, "def check_vba_structure(\n", helper, path)
    fn_start = text.index("def check_vba_structure(\n")
    start = text.index('            if upper.startswith("#IF "):\n', fn_start)
    end_marker = '''                continue

            code = _strip_vba_line(raw)
'''
    end = text.index(end_marker, start) + len("                continue\n\n")
    replacement = '''            if _handle_vba_structure_directive(
                path, number, upper, directives, failures
            ):
                continue

'''
    text = text[:start] + replacement + text[end:]
    write(path, text)


def refactor_conditionals() -> None:
    path = "tools/check_vba_conditionals.py"
    text = read(path)
    helper = '''def _handle_conditional_directive(
    path: str,
    start_line: int,
    code: str,
    stacks: dict[str, list[Frame]],
    depth: int,
    findings: list[dict[str, Any]],
) -> tuple[bool, int]:
    if CONST_RE.match(code):
        findings.append(
            {
                "path": path,
                "line": start_line,
                "message": (
                    "Project-defined #Const symbols are outside the supported model; "
                    "conditional-compilation evaluation fails closed."
                ),
            }
        )
        return True, depth

    directive = DIRECTIVE_RE.match(code)
    if directive is None:
        findings.append(
            {
                "path": path,
                "line": start_line,
                "message": f"Unsupported conditional-compilation directive: {code}",
            }
        )
        return True, depth

    kind = " ".join(directive.group(1).split()).casefold()
    remainder = directive.group(2)
    if kind == "if":
        try:
            expression = parse_condition("If", remainder)
            values = {
                name: evaluate(expression, symbols)
                for name, symbols in ENVIRONMENTS.items()
            }
        except ExpressionError as error:
            findings.append(
                {
                    "path": path,
                    "line": start_line,
                    "message": f"Indeterminate #If directive: {error}.",
                }
            )
            values = {name: False for name in ENVIRONMENTS}
        for name, stack in stacks.items():
            parent = active(stack)
            selected = parent and values[name]
            stack.append(Frame(parent, selected, selected))
        return True, depth + 1

    if kind == "elseif":
        if depth == 0 or any(not stack for stack in stacks.values()):
            findings.append(
                {"path": path, "line": start_line, "message": "#ElseIf without #If."}
            )
            return True, depth
        if any(stack[-1].else_seen for stack in stacks.values()):
            findings.append(
                {"path": path, "line": start_line, "message": "#ElseIf after #Else."}
            )
            return True, depth
        try:
            expression = parse_condition("ElseIf", remainder)
            values = {
                name: evaluate(expression, symbols)
                for name, symbols in ENVIRONMENTS.items()
            }
        except ExpressionError as error:
            findings.append(
                {
                    "path": path,
                    "line": start_line,
                    "message": f"Indeterminate #ElseIf directive: {error}.",
                }
            )
            values = {name: False for name in ENVIRONMENTS}
        for name, stack in stacks.items():
            frame = stack[-1]
            selected = frame.parent_active and not frame.branch_taken and values[name]
            frame.current_active = selected
            frame.branch_taken = frame.branch_taken or selected
        return True, depth

    if kind == "else":
        if remainder.strip():
            findings.append(
                {
                    "path": path,
                    "line": start_line,
                    "message": "#Else must not contain trailing text.",
                }
            )
            return True, depth
        if depth == 0 or any(not stack for stack in stacks.values()):
            findings.append(
                {"path": path, "line": start_line, "message": "#Else without #If."}
            )
            return True, depth
        if any(stack[-1].else_seen for stack in stacks.values()):
            findings.append(
                {"path": path, "line": start_line, "message": "Duplicate #Else."}
            )
            return True, depth
        for stack in stacks.values():
            frame = stack[-1]
            selected = frame.parent_active and not frame.branch_taken
            frame.current_active = selected
            frame.branch_taken = True
            frame.else_seen = True
        return True, depth

    if remainder.strip():
        findings.append(
            {
                "path": path,
                "line": start_line,
                "message": "#End If must not contain trailing text.",
            }
        )
        return True, depth
    if depth == 0 or any(not stack for stack in stacks.values()):
        findings.append(
            {"path": path, "line": start_line, "message": "#End If without #If."}
        )
        return True, depth
    for stack in stacks.values():
        stack.pop()
    return True, depth - 1
'''
    text = insert_before(text, "def analyze_component(path: str, text: str) -> list[dict[str, Any]]:\n", helper, path)
    fn_start = text.index("def analyze_component(path: str, text: str) -> list[dict[str, Any]]:\n")
    start = text.index('        if unit_kind == "directive":\n', fn_start)
    end = text.index("        if DECLARE_RE.match(code):\n", start)
    replacement = '''        if unit_kind == "directive":
            handled, depth = _handle_conditional_directive(
                path, start_line, code, stacks, depth, findings
            )
            if handled:
                continue

'''
    text = text[:start] + replacement + text[end:]
    write(path, text)


def refactor_initializer() -> None:
    path = "tools/initialize_repository.py"
    text = read(path)
    helper = '''def _render_tracked_changes(
    root: Path,
    tracked_files: tuple[str, ...],
    template_only: set[str],
    excluded: set[str],
    token_pattern: re.Pattern[str],
    profile: str,
    scalars: dict[str, str],
    repeatable: dict[str, list[str]],
    catalogue: dict[str, Any],
    values: dict[str, str],
    config: dict[str, Any],
    changes: dict[str, bytes | None],
    seen: set[str],
) -> None:
    for path in tracked_files:
        if path in template_only or path in excluded or not _is_text(path):
            continue
        source = (root / path).read_bytes()
        text = _decode(path, source)
        matches = list(token_pattern.finditer(text))
        if matches and PurePosixPath(path).suffix.casefold() in EXECUTABLE_SUFFIXES:
            raise InitializationError(
                f"Placeholders are prohibited in executable or VBA file {path}."
            )
        rendered = _render_blocks(
            path, text, profile, scalars, repeatable, catalogue
        )
        for match in token_pattern.finditer(rendered):
            name = match.group(1)
            if name not in catalogue:
                raise InitializationError(f"{path}: unknown placeholder {name}.")
            seen.add(name)
        for name, value in values.items():
            rendered = rendered.replace("{{" + name + "}}", value)
        if path == ".github/ISSUE_TEMPLATE/config.yml":
            template_security_url = (
                f"https://github.com/{config['repository']}/security/policy"
            )
            generated_security_url = (
                f"https://github.com/{scalars['REPOSITORY_PATH']}/security/policy"
            )
            if template_security_url not in rendered:
                raise InitializationError(
                    f"{path}: canonical template security URL is missing."
                )
            rendered = rendered.replace(template_security_url, generated_security_url)
        unresolved = sorted(
            {match.group(1) for match in token_pattern.finditer(rendered)}
        )
        if unresolved:
            raise InitializationError(
                f"{path}: unresolved placeholders: {', '.join(unresolved)}"
            )
        if path == "CHANGELOG.md":
            rendered = _reset_changelog(rendered)
        elif path == "VERSION":
            rendered = "0.0.0\n"
        output = _encode(path, rendered)
        if output != source:
            changes[path] = output
'''
    text = insert_before(text, "def _build_changes(\n", helper, path)
    fn_start = text.index("def _build_changes(\n")
    start = text.index("    for path in tracked_files:\n", fn_start)
    end = text.index("    supplied = set(scalars) | set(repeatable)\n", start)
    replacement = '''    _render_tracked_changes(
        root,
        tracked_files,
        template_only,
        excluded,
        token_pattern,
        profile,
        scalars,
        repeatable,
        catalogue,
        values,
        config,
        changes,
        seen,
    )

'''
    text = text[:start] + replacement + text[end:]
    write(path, text)


def main() -> int:
    refactor_release_assets()
    refactor_release_semantics()
    refactor_placeholder_catalogue()
    refactor_issue_forms()
    refactor_vba_structure()
    refactor_conditionals()
    refactor_initializer()
    for path in (
        "tools/check_release.py",
        "tools/check_release_semantics.py",
        "tools/check_repo.py",
        "tools/check_vba_conditionals.py",
        "tools/initialize_repository.py",
    ):
        ast.parse(read(path), filename=path)
    print("residual complexity refactor applied")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
