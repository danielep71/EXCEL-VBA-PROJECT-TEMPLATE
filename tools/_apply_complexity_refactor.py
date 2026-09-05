#!/usr/bin/env python3
"""One-shot guarded transformation for the v1.1.0 complexity hardening pass."""
from __future__ import annotations

from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parents[1]


def replace_between(path: str, start: str, end: str, replacement: str) -> None:
    target = ROOT / path
    text = target.read_text(encoding="utf-8")
    start_at = text.index(start)
    end_at = text.index(end, start_at)
    target.write_text(text[:start_at] + replacement.rstrip() + "\n\n\n" + text[end_at:], encoding="utf-8", newline="\n")


def replace_once(path: str, old: str, new: str) -> None:
    target = ROOT / path
    text = target.read_text(encoding="utf-8")
    if text.count(old) != 1:
        raise RuntimeError(f"{path}: expected one replacement anchor, found {text.count(old)}")
    target.write_text(text.replace(old, new, 1), encoding="utf-8", newline="\n")


# P3-13: the checker is normalized, so remove its temporary lint exception.
replace_once(
    "pyproject.toml",
    '[tool.ruff.lint.per-file-ignores]\n# Temporary ratchet for the legacy VBA public-API parser style. Remove these two\n# exceptions when the existing one-line/semicolon statements are normalized.\n"tools/check_vba_public_api.py" = ["E701", "E702"]\n\n',
    "",
)

# Mechanical E731 cleanup in the already-decomposed policy coverage case builders.
quality = ROOT / "tools/policy_coverage_cases_quality.py"
text = quality.read_text(encoding="utf-8")
anchor = '''def _register(cases: list[Case], name: str, rule: str, pattern: str | None = None):\n    def decorator(function: Callable[[Path], None]) -> Callable[[Path], None]:\n        cases.append((name, rule, pattern, function))\n        return function\n\n    return decorator\n'''
helper = anchor + '''\n\ndef _case(cases: list[Case]):\n    def register(name: str, rule: str, pattern: str | None = None):\n        return _register(cases, name, rule, pattern)\n\n    return register\n'''
if anchor not in text:
    raise RuntimeError("quality case registration anchor missing")
text = text.replace(anchor, helper, 1)
text = text.replace(
    "    case = lambda name, rule, pattern=None: _register(cases, name, rule, pattern)\n",
    "    case = _case(cases)\n",
)
quality.write_text(text, encoding="utf-8", newline="\n")

# Initializer: move one independent decision out of the 21-complexity renderer.
initializer = ROOT / "tools/initialize_repository.py"
text = initializer.read_text(encoding="utf-8")
marker = "def _build_changes(\n"
helper = '''def _reject_executable_placeholders(path: str, matches: list[object]) -> None:\n    if matches and PurePosixPath(path).suffix.casefold() in EXECUTABLE_SUFFIXES:\n        raise InitializationError(\n            f"Placeholders are prohibited in executable or VBA file {path}."\n        )\n\n\n'''
if helper not in text:
    text = text.replace(marker, helper + marker, 1)
old = '''        if (\n            matches\n            and PurePosixPath(path).suffix.casefold() in EXECUTABLE_SUFFIXES\n        ):\n            raise InitializationError(f"Placeholders are prohibited in executable or VBA file {path}.")\n'''
if old not in text:
    raise RuntimeError("initializer executable-placeholder block missing")
text = text.replace(old, "        _reject_executable_placeholders(path, matches)\n", 1)
initializer.write_text(text, encoding="utf-8", newline="\n")

# Release integrity: separate profile resolution, Git state, source scan, evidence checks, and assets.
release_source = r'''def _resolve_release_profile(
    root: Path, configuration: dict[str, Any]
) -> tuple[str | None, list[dict[str, str]]]:
    findings: list[dict[str, str]] = []
    mode = configuration.get("mode")
    configured_profile = configuration.get("profile")
    initialization = root / INITIALIZATION_PATH
    if mode == "generated" and configured_profile in GENERATED_PROFILES:
        release_profile = str(configured_profile)
        if not initialization.is_file():
            findings.append(_finding(
                "template-identity", INITIALIZATION_PATH,
                "generated release candidates require an initialization record",
            ))
            return release_profile, findings
        record, record_findings = _read_json(initialization, "initialization-record")
        findings.extend(record_findings)
        if isinstance(record, dict):
            if record.get("profile") != release_profile:
                findings.append(_finding(
                    "template-identity", INITIALIZATION_PATH,
                    "profile differs from repository policy",
                ))
            values = record.get("values")
            repository = configuration.get("repository")
            if not isinstance(values, dict) or values.get("REPOSITORY_PATH") != repository:
                findings.append(_finding(
                    "template-identity", INITIALIZATION_PATH,
                    "repository identity differs from repository policy",
                ))
        return release_profile, findings
    if mode == "template" and configured_profile is None:
        identity = configuration.get("identity")
        repository = configuration.get("repository")
        template_tokens = identity.get("template_tokens") if isinstance(identity, dict) else None
        contains_template_identity = (
            isinstance(repository, str)
            and isinstance(template_tokens, list)
            and any(
                isinstance(token, str) and token.casefold() in repository.casefold()
                for token in template_tokens
            )
        )
        if not contains_template_identity:
            findings.append(_finding(
                "template-identity", PROFILE_PATH,
                "template release candidates must retain a declared template identity token",
            ))
        if initialization.exists():
            findings.append(_finding(
                "template-identity", INITIALIZATION_PATH,
                "template release candidates must not contain a generated-project initialization record",
            ))
        return "template", findings
    findings.append(_finding(
        "template-identity", PROFILE_PATH,
        "release candidates must be an initialized generated profile or the canonical template profile",
    ))
    return None, findings


def _validate_candidate_git_state(
    root: Path, candidate_sha: str, tag: str, require_tag_ref: bool
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    head = _git_output(root, "rev-parse", "HEAD")
    if head is None:
        raise OperationalError("git could not resolve HEAD")
    if head != candidate_sha:
        findings.append(_finding("candidate-sha-mismatch", "HEAD", f"HEAD is {head}, expected {candidate_sha}"))
    dirty = _git_output(root, "status", "--porcelain", "--untracked-files=no")
    if dirty is None:
        raise OperationalError("git could not inspect the working tree")
    if dirty:
        findings.append(_finding("dirty-candidate", ".", "tracked candidate files differ from HEAD"))
    if not require_tag_ref:
        return findings
    reference = f"refs/tags/{tag}"
    object_type = _git_output(root, "cat-file", "-t", reference)
    target = _git_output(root, "rev-list", "-n", "1", reference)
    if object_type is None or target is None:
        findings.append(_finding("missing-tag-ref", tag, "annotated tag is not available in this clone"))
        return findings
    if object_type != "tag":
        findings.append(_finding("lightweight-tag", tag, "release tag must be annotated"))
    if target != candidate_sha:
        findings.append(_finding("tag-target-mismatch", tag, f"tag targets {target}, expected {candidate_sha}"))
    return findings


def _validate_generated_source(
    root: Path, configuration: dict[str, Any], policy: dict[str, Any]
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    tracked = _tracked_files(root)
    identity = configuration.get("identity")
    identity_excludes: set[str] = set()
    source_template_tokens: list[str] = []
    if isinstance(identity, dict):
        raw_excludes = identity.get("exclude_paths")
        raw_tokens = identity.get("template_tokens")
        if isinstance(raw_excludes, list):
            identity_excludes.update(item for item in raw_excludes if isinstance(item, str))
        if isinstance(raw_tokens, list):
            source_template_tokens.extend(item for item in raw_tokens if isinstance(item, str))
    excludes = set(policy["source_scan_exclude_paths"]) | identity_excludes
    placeholder_pattern = re.compile(r"\{\{[A-Z][A-Z0-9_]*\}\}")
    for relative in tracked:
        if relative in excludes or not _is_text(relative):
            continue
        path = root / relative
        try:
            text = path.read_text(encoding="utf-8", errors="strict")
        except (OSError, UnicodeError) as error:
            findings.append(_finding("unreadable-release-text", relative, str(error)))
            continue
        if placeholder_pattern.search(text) or "<!-- template:" in text:
            findings.append(_finding(
                "unresolved-template-token", relative,
                "unresolved placeholder or template block marker",
            ))
        folded = text.casefold()
        for token in source_template_tokens:
            if token.casefold() in folded:
                findings.append(_finding(
                    "template-identity", relative,
                    f"contains template identity token {token}",
                ))
                break
    try:
        changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
    except OSError as error:
        findings.append(_finding("missing-changelog", "CHANGELOG.md", str(error)))
        changelog = ""
    for marker in policy["template_construction_markers"]:
        if str(marker).casefold() in changelog.casefold():
            findings.append(_finding(
                "template-construction-history", "CHANGELOG.md",
                f"contains construction marker {marker!r}",
            ))
    return findings


def _validate_source(
    root: Path,
    configuration: dict[str, Any],
    policy: dict[str, Any],
    candidate_sha: str,
    tag: str,
    require_tag_ref: bool,
) -> tuple[str | None, list[dict[str, str]]]:
    release_profile, findings = _resolve_release_profile(root, configuration)
    findings.extend(_validate_candidate_git_state(root, candidate_sha, tag, require_tag_ref))
    if release_profile in GENERATED_PROFILES:
        findings.extend(_validate_generated_source(root, configuration, policy))
    return release_profile, findings
'''
replace_between(
    "tools/check_release.py",
    "def _validate_source(\n",
    "def _validate_version_and_changelog(\n",
    release_source,
)

release_evidence = r'''def _validate_evidence_metadata(
    evidence: dict[str, Any],
    evidence_path: Path,
    policy: dict[str, Any],
    profile: str | None,
    version: str | None,
    tag: str,
    candidate_sha: str,
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    if set(evidence) != TOP_LEVEL_EVIDENCE_KEYS:
        missing = sorted(TOP_LEVEL_EVIDENCE_KEYS - set(evidence))
        extra = sorted(set(evidence) - TOP_LEVEL_EVIDENCE_KEYS)
        detail = []
        if missing:
            detail.append("missing " + ", ".join(missing))
        if extra:
            detail.append("unknown " + ", ".join(extra))
        findings.append(_finding("invalid-release-evidence", str(evidence_path), "; ".join(detail)))
    expected = {
        "schema_version": policy["evidence_schema_version"],
        "version": version,
        "tag": tag,
        "candidate_sha": candidate_sha,
        "profile": profile,
    }
    for key, expected_value in expected.items():
        if evidence.get(key) != expected_value:
            code = "evidence-sha-mismatch" if key == "candidate_sha" else "evidence-metadata-mismatch"
            findings.append(_finding(code, f"evidence.{key}", f"expected {expected_value!r}"))
    if evidence.get("distribution") not in {"source-only", "binary"}:
        findings.append(_finding("invalid-release-evidence", "evidence.distribution", "must be source-only or binary"))
    return findings


def _validate_evidence_checks(
    evidence: dict[str, Any],
    policy: dict[str, Any],
    profile: str | None,
    candidate_sha: str,
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    checks = evidence.get("checks")
    required_checks: set[str] = set(policy["core_checks"])
    if profile in SUPPORTED_PROFILES:
        required_checks.update(policy["profiles"][profile]["required_checks"])
    if not isinstance(checks, dict):
        return [_finding("invalid-release-evidence", "evidence.checks", "must be an object")]
    invalid_ids = sorted(
        key for key in checks
        if not isinstance(key, str) or CHECK_ID_PATTERN.fullmatch(key) is None
    )
    if invalid_ids:
        findings.append(_finding("invalid-release-evidence", "evidence.checks", "invalid check identifiers"))
    missing_checks = sorted(required_checks - set(checks))
    if missing_checks:
        findings.append(_finding(
            "missing-profile-evidence", "evidence.checks",
            "missing checks: " + ", ".join(missing_checks),
        ))
    for check_id in sorted(checks):
        findings.extend(_validate_check(check_id, checks[check_id], candidate_sha))
    return findings


def _validate_asset_record(
    root: Path,
    asset: object,
    item_path: str,
    allowed: list[str],
    candidate_sha: str,
    evidence_entries: dict[str, str],
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    if not isinstance(asset, dict) or set(asset) != ASSET_KEYS:
        return [_finding("invalid-release-asset", item_path, "asset record has an invalid shape")]
    relative = asset.get("path")
    digest = asset.get("sha256")
    if not isinstance(relative, str) or not _safe_relative(relative):
        return [_finding("invalid-release-asset", item_path, "path must be safe and repository-relative")]
    if relative in evidence_entries:
        return [_finding("invalid-release-asset", item_path, f"duplicate asset {relative}")]
    if not isinstance(digest, str) or SHA256_PATTERN.fullmatch(digest) is None:
        return [_finding("invalid-release-asset", item_path, "sha256 must be 64 lowercase hexadecimal characters")]
    evidence_entries[relative] = digest
    if asset.get("candidate_sha") != candidate_sha:
        findings.append(_finding("asset-sha-binding-mismatch", item_path, "candidate_sha does not match the release candidate"))
    if asset.get("package_test") != "PASS":
        findings.append(_finding("failed-package-test", item_path, "package_test must be PASS"))
    if not any(PurePosixPath(relative).match(pattern) for pattern in allowed):
        findings.append(_finding("unapproved-binary", relative, "asset path is not approved for this profile"))
    target = root / relative
    try:
        resolved = target.resolve(strict=True)
        resolved.relative_to(root.resolve())
    except (OSError, ValueError):
        findings.append(_finding("missing-release-asset", relative, "asset is missing or escapes the repository root"))
        return findings
    if not resolved.is_file():
        findings.append(_finding("missing-release-asset", relative, "asset is not a regular file"))
        return findings
    actual = _sha256(resolved)
    if actual != digest:
        findings.append(_finding("asset-digest-mismatch", relative, f"actual SHA-256 is {actual}"))
    return findings


def _validate_binary_assets(
    root: Path,
    assets: list[object],
    manifest: dict[str, str] | None,
    manifest_path: Path | None,
    policy: dict[str, Any],
    profile: str | None,
    candidate_sha: str,
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    if profile == "library":
        findings.append(_finding("unapproved-binary", "evidence.distribution", "library profile is source-only by default"))
    if not assets:
        findings.append(_finding("missing-release-assets", "evidence.assets", "binary distribution requires at least one asset"))
    if manifest is None:
        findings.append(_finding("missing-asset-manifest", str(manifest_path or "<not supplied>"), "binary distribution requires a SHA-256 manifest"))
        manifest = {}
    allowed = policy["profiles"].get(profile, {}).get("allowed_asset_globs", []) if profile else []
    evidence_entries: dict[str, str] = {}
    for index, asset in enumerate(assets):
        findings.extend(_validate_asset_record(
            root, asset, f"evidence.assets[{index}]", allowed, candidate_sha, evidence_entries
        ))
    if manifest != evidence_entries:
        findings.append(_finding("asset-manifest-mismatch", str(manifest_path), "manifest entries must exactly match evidence assets"))
    return findings


def _validate_evidence_and_assets(
    root: Path,
    evidence_path: Path,
    manifest_path: Path | None,
    policy: dict[str, Any],
    profile: str | None,
    version: str | None,
    tag: str,
    candidate_sha: str,
) -> list[dict[str, str]]:
    evidence, findings = _read_json(evidence_path, "release-evidence")
    manifest, manifest_findings = _parse_manifest(manifest_path)
    findings.extend(manifest_findings)
    if not isinstance(evidence, dict):
        if evidence is not None:
            findings.append(_finding("invalid-release-evidence", str(evidence_path), "root must be an object"))
        return findings
    findings.extend(_validate_evidence_metadata(
        evidence, evidence_path, policy, profile, version, tag, candidate_sha
    ))
    findings.extend(_validate_evidence_checks(evidence, policy, profile, candidate_sha))
    assets = evidence.get("assets")
    if not isinstance(assets, list):
        findings.append(_finding("invalid-release-evidence", "evidence.assets", "must be an array"))
        return findings
    distribution = evidence.get("distribution")
    if distribution == "source-only":
        if assets:
            findings.append(_finding("source-only-has-assets", "evidence.assets", "source-only releases cannot declare binary assets"))
        if manifest:
            findings.append(_finding("source-only-has-assets", str(manifest_path), "source-only manifest must be empty or omitted"))
        return findings
    if distribution == "binary":
        findings.extend(_validate_binary_assets(
            root, assets, manifest, manifest_path, policy, profile, candidate_sha
        ))
    return findings
'''
replace_between(
    "tools/check_release.py",
    "def _validate_evidence_and_assets(\n",
    "def build_report(\n",
    release_evidence,
)

# Canonical checker: decompose configuration, issue forms, and VBA structure while remaining self-contained.
configuration = r'''def _validate_profile_contract(
    name: str, entry: object, failures: list[dict[str, Any]]
) -> None:
    if not _same_keys(entry, {"required_paths", "required_directories", "vba_contract"}):
        failures.append(finding(
            CONFIG_PATH,
            f"profiles.{name} must contain exactly required_paths, required_directories, and vba_contract.",
        ))
        return
    assert isinstance(entry, dict)
    _string_list(entry.get("required_paths"), f"profiles.{name}.required_paths", failures, paths=True)
    _string_list(entry.get("required_directories"), f"profiles.{name}.required_directories", failures, paths=True)
    contract = entry.get("vba_contract")
    field = f"profiles.{name}.vba_contract"
    if not _same_keys(contract, {"minimum_roles", "required_components"}):
        failures.append(finding(CONFIG_PATH, f"{field} must contain exactly minimum_roles and required_components."))
        return
    assert isinstance(contract, dict)
    minimum_roles = contract.get("minimum_roles")
    if not isinstance(minimum_roles, dict) or not minimum_roles:
        failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles must be a non-empty object."))
        minimum_roles = {}
    else:
        role_names = list(minimum_roles)
        if role_names != sorted(role_names, key=lambda item: (item.casefold(), item)):
            failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles keys must be sorted case-insensitively."))
        for role, minimum in minimum_roles.items():
            if role not in VBA_ROLES:
                failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles has invalid role {role!r}."))
            if isinstance(minimum, bool) or not isinstance(minimum, int) or minimum < 1:
                failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles.{role} must be a positive integer."))
    required_components = contract.get("required_components")
    if not isinstance(required_components, dict) or not required_components:
        failures.append(finding(CONFIG_PATH, f"{field}.required_components must be a non-empty object."))
        required_components = {}
    else:
        paths = list(required_components)
        if paths != sorted(paths, key=lambda item: (item.casefold(), item)):
            failures.append(finding(CONFIG_PATH, f"{field}.required_components keys must be sorted case-insensitively."))
        for path, role in required_components.items():
            if not isinstance(path, str) or not _valid_relative_path(path):
                failures.append(finding(CONFIG_PATH, f"{field}.required_components contains an invalid path: {path!r}."))
            if role not in VBA_ROLES:
                failures.append(finding(CONFIG_PATH, f"{field}.required_components.{path} has invalid role {role!r}."))
    component_roles = {role for role in required_components.values() if isinstance(role, str)}
    for role in VBA_BASELINE_ROLES:
        if minimum_roles.get(role, 0) < 1:
            failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles must require the baseline role {role!r}."))
        if role not in component_roles:
            failures.append(finding(CONFIG_PATH, f"{field}.required_components must name a component with role {role!r}."))


def _validate_profiles(document: dict[str, Any], failures: list[dict[str, Any]]) -> None:
    profiles = document.get("profiles")
    if not isinstance(profiles, dict) or set(profiles) != set(SUPPORTED_PROFILES):
        failures.append(finding(CONFIG_PATH, "profiles must contain exactly application, library, and ui-component."))
        return
    for name in SUPPORTED_PROFILES:
        _validate_profile_contract(name, profiles[name], failures)


def _validate_placeholder_spec(
    name: object,
    specification: object,
    categories_seen: set[str],
    failures: list[dict[str, Any]],
) -> None:
    field = f"placeholders.catalogue.{name}"
    if not isinstance(name, str) or not PLACEHOLDER_NAME_PATTERN.fullmatch(name):
        failures.append(finding(CONFIG_PATH, f"{field} is not a canonical placeholder name."))
        return
    if not isinstance(specification, dict):
        failures.append(finding(CONFIG_PATH, f"{field} must be an object."))
        return
    category = specification.get("category")
    description = specification.get("description")
    if category not in PLACEHOLDER_CATEGORIES:
        failures.append(finding(CONFIG_PATH, f"{field}.category must be optional, profile-specific, repeatable, or required."))
        return
    categories_seen.add(category)
    if not isinstance(description, str) or not description.strip():
        failures.append(finding(CONFIG_PATH, f"{field}.description must be non-empty."))
    expected_keys = {"category", "description"}
    if category == "profile-specific":
        expected_keys.add("values")
        values = specification.get("values")
        if not isinstance(values, dict) or set(values) != set(SUPPORTED_PROFILES):
            failures.append(finding(CONFIG_PATH, f"{field}.values must cover exactly all supported profiles."))
        elif any(not isinstance(value, str) or not value.strip() for value in values.values()):
            failures.append(finding(CONFIG_PATH, f"{field}.values must all be non-empty strings."))
    elif category == "repeatable":
        expected_keys.add("item_format")
        item_format = specification.get("item_format")
        if not isinstance(item_format, str) or item_format.count("{value}") != 1:
            failures.append(finding(CONFIG_PATH, f"{field}.item_format must contain one {{value}} field."))
    if set(specification) != expected_keys:
        failures.append(finding(CONFIG_PATH, f"{field} has fields inconsistent with its category."))


def _validate_placeholders(document: dict[str, Any], failures: list[dict[str, Any]]) -> None:
    placeholders = document.get("placeholders")
    keys = {"pattern", "catalogue", "block_markers", "template_only_paths", "exclude_paths"}
    if not _same_keys(placeholders, keys):
        failures.append(finding(CONFIG_PATH, "placeholders must contain exactly pattern, catalogue, block_markers, template_only_paths, and exclude_paths."))
        return
    assert isinstance(placeholders, dict)
    pattern = placeholders.get("pattern")
    if not isinstance(pattern, str):
        failures.append(finding(CONFIG_PATH, "placeholders.pattern must be a string."))
    else:
        try:
            compiled = re.compile(pattern)
        except re.error as error:
            failures.append(finding(CONFIG_PATH, f"placeholders.pattern is invalid: {error}."))
        else:
            if compiled.groups != 1:
                failures.append(finding(CONFIG_PATH, "placeholders.pattern must contain exactly one capture group for the token name."))
    catalogue = placeholders.get("catalogue")
    categories_seen: set[str] = set()
    if not isinstance(catalogue, dict) or not catalogue:
        failures.append(finding(CONFIG_PATH, "placeholders.catalogue must be a non-empty object."))
    else:
        names = list(catalogue)
        if names != sorted(names, key=lambda item: (item.casefold(), item)):
            failures.append(finding(CONFIG_PATH, "placeholders.catalogue keys must be sorted case-insensitively."))
        for name, specification in catalogue.items():
            _validate_placeholder_spec(name, specification, categories_seen, failures)
        missing_categories = set(PLACEHOLDER_CATEGORIES) - categories_seen
        if missing_categories:
            failures.append(finding(CONFIG_PATH, "placeholders.catalogue does not exercise categories: " + ", ".join(sorted(missing_categories))))
    expected_markers = {
        "template_only": "template:remove",
        "profile": "template:profile:{profile}",
        "optional": "template:optional:{token}",
        "repeatable": "template:repeatable:{token}",
    }
    if placeholders.get("block_markers") != expected_markers:
        failures.append(finding(CONFIG_PATH, "placeholders.block_markers must use the canonical marker grammar."))
    _string_list(placeholders.get("template_only_paths"), "placeholders.template_only_paths", failures, paths=True)
    _string_list(placeholders.get("exclude_paths"), "placeholders.exclude_paths", failures, paths=True)


def _validate_identity(
    document: dict[str, Any], mode: object, repository: object, failures: list[dict[str, Any]]
) -> None:
    identity = document.get("identity")
    keys = {"forbidden_tokens", "template_tokens", "exclude_paths"}
    if not _same_keys(identity, keys):
        failures.append(finding(CONFIG_PATH, "identity must contain exactly forbidden_tokens, template_tokens, and exclude_paths."))
        return
    assert isinstance(identity, dict)
    forbidden = _string_list(identity.get("forbidden_tokens"), "identity.forbidden_tokens", failures)
    template = _string_list(identity.get("template_tokens"), "identity.template_tokens", failures)
    _string_list(identity.get("exclude_paths"), "identity.exclude_paths", failures, paths=True)
    if not template:
        failures.append(finding(CONFIG_PATH, "identity.template_tokens must not be empty."))
    if not isinstance(repository, str):
        return
    folded = repository.casefold()
    for token in forbidden:
        if token.casefold() in folded:
            failures.append(finding(CONFIG_PATH, f"repository contains a forbidden donor token: {token}"))
    contains_template = any(token.casefold() in folded for token in template)
    if mode == "template" and template and not contains_template:
        failures.append(finding(CONFIG_PATH, "Template mode repository must contain a declared template identity token."))
    elif mode == "generated" and contains_template:
        failures.append(finding(CONFIG_PATH, "Generated mode repository still contains a template identity token."))


def _validate_vba_configuration(document: dict[str, Any], failures: list[dict[str, Any]]) -> None:
    vba = document.get("vba")
    keys = {"source_roots", "test_roots", "components", "public_api_manifest"}
    if not _same_keys(vba, keys):
        failures.append(finding(CONFIG_PATH, "vba must contain exactly source_roots, test_roots, components, and public_api_manifest."))
        return
    assert isinstance(vba, dict)
    source_roots = _string_list(vba.get("source_roots"), "vba.source_roots", failures, paths=True)
    test_roots = _string_list(vba.get("test_roots"), "vba.test_roots", failures, paths=True)
    if set(source_roots).intersection(test_roots):
        failures.append(finding(CONFIG_PATH, "VBA source and test roots must not overlap."))
    components = vba.get("components")
    if not isinstance(components, dict):
        failures.append(finding(CONFIG_PATH, "vba.components must be an object."))
    else:
        component_keys = list(components)
        if component_keys != sorted(component_keys, key=lambda item: (item.casefold(), item)):
            failures.append(finding(CONFIG_PATH, "vba.components keys must be sorted case-insensitively."))
        for path, role in components.items():
            if not isinstance(path, str) or not _valid_relative_path(path):
                failures.append(finding(CONFIG_PATH, f"Invalid VBA component path: {path!r}."))
            if role not in VBA_ROLES:
                failures.append(finding(CONFIG_PATH, f"VBA component {path!r} has unsupported role {role!r}."))
    api_manifest = vba.get("public_api_manifest")
    if api_manifest is not None and (not isinstance(api_manifest, str) or not _valid_relative_path(api_manifest)):
        failures.append(finding(CONFIG_PATH, "vba.public_api_manifest must be null or a valid relative path."))


def _validate_configuration_root(
    document: dict[str, Any], failures: list[dict[str, Any]]
) -> tuple[object, object]:
    if document.get("schema_version") != SCHEMA_VERSION:
        failures.append(finding(CONFIG_PATH, f"schema_version must be {SCHEMA_VERSION}."))
    mode = document.get("mode")
    profile = document.get("profile")
    if mode not in {"template", "generated"}:
        failures.append(finding(CONFIG_PATH, "mode must be template or generated."))
    elif mode == "template" and profile is not None:
        failures.append(finding(CONFIG_PATH, "Template mode requires profile to be null."))
    elif mode == "generated" and profile not in SUPPORTED_PROFILES:
        failures.append(finding(CONFIG_PATH, "Generated mode requires profile to be application, library, or ui-component."))
    repository = document.get("repository")
    if not isinstance(repository, str) or not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository):
        failures.append(finding(CONFIG_PATH, "repository must use the owner/name form."))
    label_domains = _string_list(document.get("label_domains"), "label_domains", failures)
    for domain in label_domains:
        if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", domain):
            failures.append(finding(CONFIG_PATH, f"label_domains contains a non-kebab-case name: {domain!r}."))
    if mode == "template" and label_domains:
        failures.append(finding(CONFIG_PATH, "Template mode requires label_domains to be empty."))
    _string_list(document.get("required_paths"), "required_paths", failures, paths=True)
    _string_list(document.get("required_directories"), "required_directories", failures, paths=True)
    _string_list(document.get("allowed_office_binary_globs"), "allowed_office_binary_globs", failures)
    return mode, repository


def load_configuration(
    repo: Repository,
) -> tuple[dict[str, Any] | None, dict[str, Any]]:
    failures: list[dict[str, Any]] = []
    try:
        document = json.loads(repo.text(CONFIG_PATH))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        failures.append(finding(CONFIG_PATH, f"Cannot load repository profile: {error}", getattr(error, "lineno", None)))
        return None, rule_result("configuration", "Repository profile configuration", failures, "")
    if not _same_keys(document, CONFIG_KEYS):
        failures.append(finding(CONFIG_PATH, "Root object must contain exactly the canonical configuration keys."))
    if not isinstance(document, dict):
        document = {}
    mode, repository = _validate_configuration_root(document, failures)
    _validate_profiles(document, failures)
    _validate_placeholders(document, failures)
    _validate_identity(document, mode, repository, failures)
    _validate_vba_configuration(document, failures)
    return (
        document if not failures else None,
        rule_result(
            "configuration",
            "Repository profile configuration",
            failures,
            "Versioned template/profile configuration is valid",
        ),
    )
'''
replace_between(
    "tools/check_repo.py",
    "def load_configuration(\n",
    "def _effective_requirements(\n",
    configuration,
)

issue_forms = r'''def _issue_label_names(repo: Repository) -> set[object]:
    try:
        manifest = json.loads(repo.text(LABEL_MANIFEST_PATH))
    except (OSError, UnicodeError, json.JSONDecodeError):
        manifest = {}
    return {
        label.get("name")
        for label in manifest.get("core", [])
        if isinstance(label, dict) and isinstance(label.get("name"), str)
    }


def _validate_issue_form(
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
        failures.append(finding(path, f"Top-level title must be {specification['title']!r}."))
    labels = _yaml_flow_array(text, "labels")
    expected_label = specification["label"]
    if labels != [expected_label]:
        failures.append(finding(path, f"Top-level labels must be the JSON flow array [{expected_label!r}]."))
    elif expected_label not in label_names:
        failures.append(finding(path, f"Issue form label is absent from {LABEL_MANIFEST_PATH}: {expected_label}"))
    if _yaml_flow_array(text, "assignees") != []:
        failures.append(finding(path, "Top-level assignees must be an empty JSON flow array."))
    blocks = _issue_form_blocks(text)
    if not blocks or len(blocks) > 10:
        failures.append(finding(path, "Issue form body must contain between 1 and 10 elements."))
    invalid_types = sorted({kind for kind, _ in blocks} - {"checkboxes", "dropdown", "input", "markdown", "textarea"})
    if invalid_types:
        failures.append(finding(path, "Unsupported issue-form element types: " + ", ".join(invalid_types)))
    identifiers = [identifier for _, identifier in blocks if identifier is not None]
    if len(identifiers) != len(set(identifiers)):
        failures.append(finding(path, "Issue-form element IDs must be unique."))
    missing = sorted(set(specification["required_ids"]) - set(identifiers))
    if missing:
        failures.append(finding(path, "Required issue-form element IDs are missing: " + ", ".join(missing)))
    if "SECURITY.md" not in text or "private" not in text.casefold():
        failures.append(finding(path, "The opening guidance must route vulnerability details to SECURITY.md privately."))
    for identifier in set(specification["required_ids"]):
        if identifier == "acknowledgements":
            continue
        pattern = rf"(?ms)^    id:\s*{re.escape(identifier)}\s*$.*?(?=^  - type:|\Z)"
        match = re.search(pattern, text)
        if match and not re.search(r"(?m)^      required:\s*true\s*$", match.group(0)):
            failures.append(finding(path, f"Required evidence field {identifier!r} must be mandatory."))


def _validate_issue_intake_config(
    repo: Repository, repository: str, failures: list[dict[str, Any]]
) -> None:
    path = f"{ISSUE_TEMPLATE_DIRECTORY}/config.yml"
    try:
        text = repo.text(path)
    except (OSError, UnicodeError) as error:
        failures.append(finding(path, f"Cannot read issue-template configuration: {error}"))
        return
    if not re.search(r"(?m)^blank_issues_enabled:\s*false\s*$", text):
        failures.append(finding(path, "Blank issues must be disabled."))
    expected_url = f"https://github.com/{repository}/security/policy"
    url_match = re.search(r'(?m)^\s+url:\s*["\']?([^"\'\s]+)["\']?\s*$', text)
    actual_url = url_match.group(1) if url_match else None
    if actual_url != expected_url:
        failures.append(finding(path, f"Private-security contact URL must be {expected_url!r}."))
    if "private" not in text.casefold():
        failures.append(finding(path, "Security contact guidance must require private reporting."))


def check_issue_forms(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    label_names = _issue_label_names(repo)
    for filename, specification in ISSUE_FORM_SPECS.items():
        _validate_issue_form(repo, filename, specification, label_names, failures)
    _validate_issue_intake_config(repo, config["repository"], failures)
    return rule_result(
        "issue-forms",
        "Structured issue forms and private security routing",
        failures,
        "Validated three canonical forms, labels, evidence fields, and private security routing",
    )
'''
replace_between(
    "tools/check_repo.py",
    "def check_issue_forms(\n",
    "def check_workflow_actions(\n",
    issue_forms,
)

vba_structure = r'''def _handle_vba_directive(
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
            directives[-1]["active"] = "VBA7" in condition and "NOT VBA7" not in condition
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


def _scan_vba_structure_component(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> tuple[list[dict[str, bool]], list[tuple[str, str, int]], set[str], list[tuple[int, str]]]:
    opener = re.compile(
        r"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
        r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b",
        re.IGNORECASE,
    )
    closer = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
    label_re = re.compile(r"^\s*([A-Za-z_]\w*|\d+):\s*$")
    declare_re = re.compile(r"^\s*(?:Public|Private)?\s*Declare\s+(?:Function|Sub)\b", re.IGNORECASE)
    directives: list[dict[str, bool]] = []
    procedures: list[tuple[str, str, int]] = []
    labels: set[str] = set()
    executable: list[tuple[int, str]] = []
    for number, raw in enumerate(lines, start=1):
        upper = raw.strip().upper()
        if _handle_vba_directive(path, number, upper, directives, failures):
            continue
        code = _strip_vba_line(raw)
        if not code.strip():
            continue
        executable.append((number, code))
        match = label_re.match(code)
        if match:
            labels.add(match.group(1).casefold())
        if declare_re.match(code) and any(item["active"] for item in directives) and not re.search(r"\bPtrSafe\b", code, re.IGNORECASE):
            failures.append(finding(path, "Declare in an active VBA7 branch must include PtrSafe.", number))
        match = opener.match(code)
        if match and " declare " not in f" {code.casefold()} ":
            if procedures:
                kind, name, start = procedures[-1]
                failures.append(finding(path, f"{kind} {name} opened at line {start} has no closing statement.", number))
                procedures.clear()
            procedures.append((match.group(1), match.group(2), number))
            continue
        match = closer.match(code)
        if match:
            if not procedures:
                failures.append(finding(path, f"{match.group(0).strip()} has no opener.", number))
            else:
                procedures.pop()
    return directives, procedures, labels, executable


def _validate_vba_structure_tail(
    path: str,
    directives: list[dict[str, bool]],
    procedures: list[tuple[str, str, int]],
    labels: set[str],
    executable: list[tuple[int, str]],
    failures: list[dict[str, Any]],
) -> None:
    if directives:
        failures.append(finding(path, f"{len(directives)} conditional-compilation block(s) are unclosed."))
    for kind, name, start in procedures:
        failures.append(finding(path, f"{kind} {name} opened at line {start} is unclosed."))
    jump_re = re.compile(r"\b(?:GoTo|Resume)\s+([A-Za-z_]\w*|\d+|-1)\b", re.IGNORECASE)
    for number, code in executable:
        for match in jump_re.finditer(code):
            target = match.group(1)
            if target.casefold() in {"next", "0", "-1"}:
                continue
            if target.casefold() not in labels:
                failures.append(finding(path, f"Jump target is not defined: {target}", number))


def check_vba_structure(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    paths = _vba_paths(repo)
    for path in paths:
        try:
            lines = repo.text(path).splitlines()
        except (OSError, UnicodeError):
            continue
        state = _scan_vba_structure_component(path, lines, failures)
        _validate_vba_structure_tail(path, *state, failures)
    return rule_result(
        "vba-structure",
        "VBA structural safety",
        failures,
        f"Validated procedure, directive, jump, and PtrSafe structure in {len(paths)} components",
    )
'''
replace_between(
    "tools/check_repo.py",
    "def check_vba_structure(\n",
    "def check_vba_visibility(\n",
    vba_structure,
)

# Restore the permanent workflow from the parent commit and remove this one-shot helper.
workflow = subprocess.run(
    ["git", "-C", str(ROOT), "show", "HEAD^:.github/workflows/checker-development.yml"],
    check=True,
    stdout=subprocess.PIPE,
    text=True,
).stdout
(ROOT / ".github/workflows/checker-development.yml").write_text(workflow, encoding="utf-8", newline="\n")
Path(__file__).unlink()
