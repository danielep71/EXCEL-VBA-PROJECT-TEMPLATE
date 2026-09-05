#!/usr/bin/env python3
"""One-shot guarded v1.1.0 XML and generated-payload hardening."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def replace_once(path: str, old: str, new: str) -> None:
    target = ROOT / path
    text = target.read_text(encoding="utf-8")
    count = text.count(old)
    if count != 1:
        raise RuntimeError(f"{path}: expected one replacement anchor, found {count}")
    target.write_text(text.replace(old, new, 1), encoding="utf-8", newline="\n")


# P3-14: retain check_repo.py's stdlib-only portability while bounding XML input
# and rejecting DTD/entity declarations before the narrow stdlib parse call.
replace_once(
    "tools/check_repo.py",
    'TOOL_NAME = "Canonical repository quality"\n',
    'TOOL_NAME = "Canonical repository quality"\n'
    'MAX_XML_BYTES = 1024 * 1024\n'
    'UNSAFE_XML_DECLARATION = re.compile(r"<!\\s*(?:DOCTYPE|ENTITY)\\b", re.IGNORECASE)\n',
)

xml_helper = '''def _validate_xml_text(path: str, text: str) -> dict[str, Any] | None:\n    try:\n        size = len(text.encode("utf-8"))\n    except UnicodeError as error:\n        return finding(path, f"Cannot encode XML as UTF-8: {error}")\n    if size > MAX_XML_BYTES:\n        return finding(\n            path,\n            f"XML input exceeds the {MAX_XML_BYTES}-byte structural-validation limit.",\n        )\n    match = UNSAFE_XML_DECLARATION.search(text)\n    if match is not None:\n        return finding(\n            path,\n            "XML DTD and entity declarations are prohibited by the portable checker.",\n            line_number(text, match.start()),\n        )\n    try:\n        ET.fromstring(text)  # noqa: S314 -- bounded input with DTD/entity declarations rejected above.\n    except ET.ParseError as error:\n        return finding(path, f"Invalid XML: {error}", error.position[0])\n    return None\n\n\n'''
replace_once(
    "tools/check_repo.py",
    "def check_structured_data(\n",
    xml_helper + "def check_structured_data(\n",
)
replace_once(
    "tools/check_repo.py",
    '''        elif suffix == ".xml":\n            counts["xml"] += 1\n            try:\n                ET.fromstring(repo.text(path))\n            except (OSError, UnicodeError, ET.ParseError) as error:\n                line = error.position[0] if isinstance(error, ET.ParseError) else None\n                failures.append(finding(path, f"Invalid XML: {error}", line))\n''',
    '''        elif suffix == ".xml":\n            counts["xml"] += 1\n            try:\n                text = repo.text(path)\n            except (OSError, UnicodeError) as error:\n                failures.append(finding(path, f"Cannot decode XML as UTF-8: {error}"))\n                continue\n            xml_failure = _validate_xml_text(path, text)\n            if xml_failure is not None:\n                failures.append(xml_failure)\n''',
)

xml_degrades = '''def _degrade_structured_xml_doctype(root: Path) -> None:\n    path = root / "docs/doctype.xml"\n    _write_fixture(\n        path,\n        '<!DOCTYPE repository [<!ENTITY sample "value">]><repository>&sample;</repository>\\n',\n    )\n    _run_git(root, "add", path.relative_to(root).as_posix())\n\n\ndef _degrade_structured_xml_oversize(root: Path) -> None:\n    path = root / "docs/oversize.xml"\n    padding = "x" * MAX_XML_BYTES\n    _write_fixture(path, f"<repository>{padding}</repository>\\n")\n    _run_git(root, "add", path.relative_to(root).as_posix())\n\n\n'''
replace_once(
    "tools/check_repo.py",
    "def _degrade_markdown_links(root: Path) -> None:\n",
    xml_degrades + "def _degrade_markdown_links(root: Path) -> None:\n",
)
replace_once(
    "tools/check_repo.py",
    '''BRANCH_SELF_TEST_CASES: tuple[\n    tuple[str, str, Callable[[Path], None]], ...\n] = (\n    ("structured-yaml", "structured-data", _degrade_structured_yaml),\n    ("structured-xml", "structured-data", _degrade_structured_xml),\n)\n''',
    '''BRANCH_SELF_TEST_CASES: tuple[\n    tuple[str, str, Callable[[Path], None]], ...\n] = (\n    ("structured-yaml", "structured-data", _degrade_structured_yaml),\n    ("structured-xml", "structured-data", _degrade_structured_xml),\n    ("structured-xml-doctype", "structured-data", _degrade_structured_xml_doctype),\n    ("structured-xml-oversize", "structured-data", _degrade_structured_xml_oversize),\n)\n''',
)
replace_once(
    "pyproject.toml",
    'select = ["E4", "E7", "E9", "F", "C90"]',
    'select = ["E4", "E7", "E9", "F", "C90", "S314"]',
)

# P2-12: generated projects retain operational gates but not template-maintainer
# checker-development / semantic policy-coverage machinery.
profile_path = ROOT / ".github/repository-profile.json"
profile = json.loads(profile_path.read_text(encoding="utf-8"))
template_only = set(profile["placeholders"]["template_only_paths"])
template_only.update(
    {
        ".github/workflows/checker-development.yml",
        "docs/CHECKER_DEVELOPMENT.md",
        "tools/check_policy_coverage.py",
        "tools/checker_development.py",
        "tools/policy_coverage_cases_config.py",
        "tools/policy_coverage_cases_quality.py",
        "tools/policy_coverage_cases_repo.py",
        "tools/policy_coverage_core.py",
        "tools/policy_coverage_runner.py",
    }
)
profile["placeholders"]["template_only_paths"] = sorted(template_only, key=lambda item: (item.casefold(), item))
profile_path.write_text(json.dumps(profile, indent=2, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")

replace_once(
    "docs/INITIALIZATION.md",
    '''- deletes every path declared under `placeholders.template_only_paths`, subject\n  only to the documented social-preview retention exception;\n''',
    '''- deletes every path declared under `placeholders.template_only_paths`, subject\n  only to the documented social-preview retention exception; template-maintainer\n  checker-development and semantic policy-coverage tooling is deliberately in\n  that set, while operational repository/release/VBA gates remain in generated\n  projects;\n''',
)
replace_once(
    "docs/INITIALIZATION.md",
    '''The initializer self-test exercises missing, unknown, and unused inputs; dry-run\nimmutability; application; second-run idempotence; template-only cleanup; and a\ngreen generated tree for every profile. For each profile it also proves that a\n''',
    '''The initializer self-test exercises missing, unknown, and unused inputs; dry-run\nimmutability; application; second-run idempotence; template-only cleanup; and a\ngreen generated tree for every profile. It also verifies that template-maintainer\nchecker-development/policy-coverage files are absent after generation while the\noperational gate set remains. For each profile it proves that a\n''',
)
replace_once(
    "docs/CHECKER_DEVELOPMENT.md",
    '''`tools/check_repo.py` must never import `_gatelib.py`. Generated repositories retain `_gatelib.py` for the focused gates, while the canonical checker remains independently copyable and executable as one standard-library-only file. `checker_development.py` enforces this ownership boundary.\n''',
    '''`tools/check_repo.py` must never import `_gatelib.py`. Generated repositories retain `_gatelib.py` for the focused operational gates, while the canonical checker remains independently copyable and executable as one standard-library-only file. `checker_development.py` enforces this ownership boundary in the canonical template. The checker-development workflow, this document, and the `policy_coverage_*` semantic-coverage harness are template-maintainer assets and are removed by initialization rather than shipped into generated projects.\n''',
)
replace_once(
    "tools/README.md",
    '''The first command exercises a passing fixture, one deliberately degraded\nfixture for each canonical rule, direct malformed-YAML and malformed-XML branch\nfixtures, deterministic JSON and Markdown rendering, and read-only execution.\n''',
    '''The first command exercises a passing fixture, one deliberately degraded\nfixture for each canonical rule, malformed YAML/XML, prohibited XML DTD/entity\ndeclarations, oversized XML, deterministic JSON and Markdown rendering, and\nread-only execution.\n''',
)
replace_once(
    "tools/README.md",
    '''`_gatelib.py` is the private, standard-library-only owner of Git, report-output, tracked-file, and common focused-gate CLI primitives. `check_repo.py` deliberately does not import it: the canonical checker remains a self-contained distributable artifact.\n''',
    '''`_gatelib.py` is the private, standard-library-only owner of Git, report-output, tracked-file, and common focused-gate CLI primitives. `check_repo.py` deliberately does not import it: the canonical checker remains a self-contained distributable artifact. The canonical template also carries checker-development and semantic policy-coverage harnesses; initialization strips those maintainer-only files while retaining the operational gates needed by generated projects.\n''',
)
replace_once(
    "CHANGELOG.md",
    '''- Removed residual unused imports exposed by the enforced Ruff baseline.\n''',
    '''- Removed residual unused imports exposed by the enforced Ruff baseline.\n- Reduced Python checker complexity under a permanently enforced McCabe ceiling\n  of 20 and normalized the VBA public-API checker so no style exception remains.\n- Hardened stdlib XML validation with a bounded input size and fail-closed\n  rejection of DTD/entity declarations before parsing.\n- Split template-maintainer checker-development and semantic policy-coverage\n  tooling from the operational tool payload retained by generated repositories.\n''',
)

Path(__file__).unlink()
