from __future__ import annotations

from pathlib import Path
from types import ModuleType
from typing import Callable

from policy_coverage_core import mutate_config


def configuration_cases(module: ModuleType) -> list[tuple[str, str, str | None, Callable[[Path], None]]]:
    cases: list[tuple[str, str, str | None, Callable[[Path], None]]] = []

    def assign(**values: object) -> Callable[[dict], None]:
        def mutate(document: dict) -> None:
            document.update(values)
        return mutate

    def case(name: str, rule: str, pattern: str | None = None):
        def register(function: Callable[[Path], None]) -> Callable[[Path], None]:
            cases.append((name, rule, pattern, function))
            return function
        return register

    def ccase(name: str, mutation: Callable[[dict], None], pattern: str | None = None) -> None:
        @case(name, "configuration", pattern)
        def _(root: Path) -> None:
            mutate_config(module, root, mutation)

    # Configuration/schema branches.
    @case("config-invalid-json", "configuration", "Cannot load repository profile")
    def _(root: Path) -> None:
        module._write_fixture(root / module.CONFIG_PATH, "{\n")

    ccase("config-root-keys", lambda d: d.__setitem__("extra", True), "canonical configuration keys")
    ccase("config-mode-invalid", lambda d: d.__setitem__("mode", "invalid"), "mode must be")
    ccase("config-template-profile", assign(mode="template", profile="library"), "Template mode requires profile")
    ccase("config-generated-profile", lambda d: d.__setitem__("profile", "invalid"), "Generated mode requires profile")
    ccase("config-repository-form", lambda d: d.__setitem__("repository", "invalid"), "owner/name")
    ccase("config-string-list-type", lambda d: d.__setitem__("required_paths", None), "array of strings")
    ccase("config-string-list-duplicate-order", lambda d: d.__setitem__("required_paths", ["z", "a", "a"]), None)
    ccase("config-invalid-relative-path", lambda d: d.__setitem__("required_paths", ["../bad"]), "invalid relative path")
    ccase("config-label-domain-name", lambda d: d.__setitem__("label_domains", ["Bad_Name"]), "non-kebab-case")
    ccase("config-template-domains", assign(mode="template", profile=None, repository="example/TEMPLATE-IDENTITY", label_domains=["domain"]), "Template mode requires label_domains")
    ccase("config-profiles-shape", lambda d: d.__setitem__("profiles", {}), "profiles must contain exactly")
    ccase("config-profile-entry-shape", lambda d: d["profiles"]["library"].__setitem__("extra", True), "required_paths, required_directories")
    ccase("config-contract-shape", lambda d: d["profiles"]["library"].__setitem__("vba_contract", {}), "minimum_roles and required_components")
    ccase("config-minimum-roles-empty", lambda d: d["profiles"]["library"]["vba_contract"].__setitem__("minimum_roles", {}), "minimum_roles must be a non-empty")
    ccase("config-minimum-roles-order", lambda d: d["profiles"]["library"]["vba_contract"].__setitem__("minimum_roles", {"test": 1, "public": 1, "internal": 1}), "minimum_roles keys must be sorted")
    ccase("config-minimum-role-invalid", lambda d: d["profiles"]["library"]["vba_contract"]["minimum_roles"].__setitem__("wrong", 1), "invalid role")
    ccase("config-minimum-value-invalid", lambda d: d["profiles"]["library"]["vba_contract"]["minimum_roles"].__setitem__("internal", 0), "positive integer")
    ccase("config-required-components-empty", lambda d: d["profiles"]["library"]["vba_contract"].__setitem__("required_components", {}), "required_components must be a non-empty")
    ccase("config-required-components-order", lambda d: d["profiles"]["library"]["vba_contract"].__setitem__("required_components", {"tests/modules/QualityTests.bas": "test", "src/modules/Quality.bas": "public", "src/core/QualityCore.bas": "internal"}), "required_components keys must be sorted")
    ccase("config-required-component-path", lambda d: d["profiles"]["library"]["vba_contract"]["required_components"].__setitem__("../bad.bas", "public"), "invalid path")
    ccase("config-required-component-role", lambda d: d["profiles"]["library"]["vba_contract"]["required_components"].__setitem__("src/modules/Quality.bas", "wrong"), "invalid role")
    ccase("config-baseline-minimum-role", lambda d: d["profiles"]["library"]["vba_contract"]["minimum_roles"].pop("public"), "baseline role")
    ccase("config-baseline-component-role", lambda d: d["profiles"]["library"]["vba_contract"]["required_components"].pop("src/modules/Quality.bas"), "must name a component with role")
    ccase("config-placeholders-shape", lambda d: d.__setitem__("placeholders", {}), "placeholders must contain exactly")
    ccase("config-placeholder-pattern-type", lambda d: d["placeholders"].__setitem__("pattern", 7), "pattern must be a string")
    ccase("config-placeholder-pattern-invalid", lambda d: d["placeholders"].__setitem__("pattern", "("), "pattern is invalid")
    ccase("config-placeholder-pattern-groups", lambda d: d["placeholders"].__setitem__("pattern", r"\{\{[A-Z]+\}\}"), "exactly one capture group")
    ccase("config-placeholder-catalogue-empty", lambda d: d["placeholders"].__setitem__("catalogue", {}), "catalogue must be a non-empty")
    ccase("config-placeholder-catalogue-order", lambda d: d["placeholders"].__setitem__("catalogue", {"ZZZ": {"category": "required", "description": "z"}, **d["placeholders"]["catalogue"]}), "catalogue keys must be sorted")
    ccase("config-placeholder-name", lambda d: d["placeholders"]["catalogue"].__setitem__("bad-name", {"category": "required", "description": "x"}), "canonical placeholder name")
    ccase("config-placeholder-object", lambda d: d["placeholders"]["catalogue"].__setitem__("OPTIONAL_NOTE", "bad"), "must be an object")
    ccase("config-placeholder-category", lambda d: d["placeholders"]["catalogue"]["OPTIONAL_NOTE"].__setitem__("category", "bad"), "category must be")
    ccase("config-placeholder-description", lambda d: d["placeholders"]["catalogue"]["OPTIONAL_NOTE"].__setitem__("description", ""), "description must be non-empty")
    ccase("config-placeholder-profile-values-shape", lambda d: d["placeholders"]["catalogue"]["PROFILE_NOTE"].__setitem__("values", {}), "values must cover exactly")
    ccase("config-placeholder-profile-values-empty", lambda d: d["placeholders"]["catalogue"]["PROFILE_NOTE"]["values"].__setitem__("library", ""), "values must all be non-empty")
    ccase("config-placeholder-repeatable-format", lambda d: d["placeholders"]["catalogue"]["REPEATABLE_NOTE"].__setitem__("item_format", "bad"), "item_format must contain one")
    ccase("config-placeholder-extra-fields", lambda d: d["placeholders"]["catalogue"]["OPTIONAL_NOTE"].__setitem__("extra", True), "fields inconsistent")
    ccase("config-placeholder-missing-category", lambda d: d["placeholders"]["catalogue"].pop("OPTIONAL_NOTE"), "does not exercise categories")
    ccase("config-placeholder-markers", lambda d: d["placeholders"].__setitem__("block_markers", {}), "canonical marker grammar")
    ccase("config-identity-shape", lambda d: d.__setitem__("identity", {}), "identity must contain exactly")
    ccase("config-template-tokens-empty", lambda d: d["identity"].__setitem__("template_tokens", []), "template_tokens must not be empty")
    ccase("config-repository-forbidden-token", lambda d: d.__setitem__("repository", "example/DONOR-PROJECT"), "forbidden donor token")
    ccase("config-template-identity-missing", assign(mode="template", profile=None, repository="example/plain"), "Template mode repository")
    ccase("config-generated-template-identity", lambda d: d.__setitem__("repository", "example/TEMPLATE-IDENTITY"), "Generated mode repository")
    ccase("config-vba-shape", lambda d: d.__setitem__("vba", {}), "vba must contain exactly")
    ccase("config-vba-roots-overlap", lambda d: d["vba"].__setitem__("test_roots", ["src"]), "source and test roots must not overlap")
    ccase("config-vba-components-type", lambda d: d["vba"].__setitem__("components", []), "components must be an object")
    ccase("config-vba-components-order", lambda d: d["vba"].__setitem__("components", {"tests/modules/QualityTests.bas": "test", "src/modules/Quality.bas": "public", "src/core/QualityCore.bas": "internal"}), "components keys must be sorted")
    ccase("config-vba-component-path", lambda d: d["vba"]["components"].__setitem__("../bad.bas", "public"), "Invalid VBA component path")
    ccase("config-vba-component-role", lambda d: d["vba"]["components"].__setitem__("src/modules/Quality.bas", "wrong"), "unsupported role")
    ccase("config-vba-api-manifest", lambda d: d["vba"].__setitem__("public_api_manifest", "../bad"), "public_api_manifest")

    return cases
