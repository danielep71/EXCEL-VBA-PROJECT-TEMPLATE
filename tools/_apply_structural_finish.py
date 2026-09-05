#!/usr/bin/env python3
from __future__ import annotations

import json
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parents[1]


def replace_once(path: str, old: str, new: str) -> None:
    target = ROOT / path
    text = target.read_text(encoding="utf-8")
    count = text.count(old)
    if count != 1:
        raise SystemExit(f"{path}: expected one replacement anchor, found {count}")
    target.write_text(text.replace(old, new, 1), encoding="utf-8", newline="\n")


# Harden stdlib XML parsing for tracked repository XML without introducing a dependency.
replace_once(
    "tools/check_repo.py",
    '''        elif suffix == ".xml":\n            counts["xml"] += 1\n            try:\n                ET.fromstring(repo.text(path))\n            except (OSError, UnicodeError, ET.ParseError) as error:\n                line = error.position[0] if isinstance(error, ET.ParseError) else None\n                failures.append(finding(path, f"Invalid XML: {error}", line))\n''',
    '''        elif suffix == ".xml":\n            counts["xml"] += 1\n            try:\n                text = repo.text(path)\n            except (OSError, UnicodeError) as error:\n                failures.append(finding(path, f"Invalid XML: {error}"))\n                continue\n            if len(text.encode("utf-8")) > 1_000_000:\n                failures.append(finding(path, "XML exceeds the 1,000,000-byte parser safety limit."))\n                continue\n            if re.search(r"<!\\s*(?:DOCTYPE|ENTITY)\\b", text, re.IGNORECASE):\n                failures.append(finding(path, "XML DTD/entity declarations are prohibited."))\n                continue\n            try:\n                ET.fromstring(text)\n            except ET.ParseError as error:\n                failures.append(finding(path, f"Invalid XML: {error}", error.position[0]))\n''',
)

# Teach the initializer to remove template-maintainer policy-coverage references from generated CI.
insert_anchor = '''def _reject_executable_placeholders(path: str, matches: list[Any]) -> None:\n    if matches and PurePosixPath(path).suffix.casefold() in EXECUTABLE_SUFFIXES:\n        raise InitializationError(\n            f"Placeholders are prohibited in executable or VBA file {path}."\n        )\n\n\n'''
insert_new = insert_anchor + '''def _strip_template_maintenance_workflow(text: str) -> str:\n    lines = text.splitlines(keepends=True)\n    output: list[str] = []\n    skipping_policy_steps = False\n    for line in lines:\n        if line.startswith("      - name: Exercise policy-branch coverage determinism"):\n            skipping_policy_steps = True\n            continue\n        if skipping_policy_steps and line.startswith("      - name: Exercise positive and degraded checker paths"):\n            skipping_policy_steps = False\n        if skipping_policy_steps:\n            continue\n        if "test-results/policy-coverage." in line:\n            continue\n        if "POLICY_COVERAGE_" in line:\n            continue\n        if '"Policy-coverage self-test:' in line or '"Policy coverage:' in line:\n            continue\n        output.append(line)\n    return "".join(output)\n\n\n'''
replace_once("tools/initialize_repository.py", insert_anchor, insert_new)
replace_once(
    "tools/initialize_repository.py",
    '''        if path == "CHANGELOG.md":\n            rendered = _reset_changelog(rendered)\n        elif path == "VERSION":\n            rendered = "0.0.0\\n"\n''',
    '''        if path == "CHANGELOG.md":\n            rendered = _reset_changelog(rendered)\n        elif path == "VERSION":\n            rendered = "0.0.0\\n"\n        elif path == ".github/workflows/static-checks.yml":\n            rendered = _strip_template_maintenance_workflow(rendered)\n''',
)

# Classify template-maintainer-only harnesses as non-deployable.
profile_path = ROOT / ".github/repository-profile.json"
profile = json.loads(profile_path.read_text(encoding="utf-8"))
template_only = set(profile["placeholders"]["template_only_paths"])
template_only.update(
    {
        ".github/workflows/checker-development.yml",
        "docs/CHECKER_DEVELOPMENT.md",
        "tools/checker_development.py",
        "tools/check_policy_coverage.py",
    }
)
template_only.update(path.relative_to(ROOT).as_posix() for path in (ROOT / "tools").glob("policy_coverage_*.py"))
profile["placeholders"]["template_only_paths"] = sorted(template_only, key=lambda item: (item.casefold(), item))
profile_path.write_text(json.dumps(profile, indent=2, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")

# Durable documentation.
replace_once(
    "docs/INITIALIZATION.md",
    "Initialization removes template-only evidence and assets declared by the repository profile.",
    "Initialization removes template-only evidence, maintainer-only checker-development/policy-coverage tooling, and assets declared by the repository profile; generated static CI is pruned so it never references removed maintenance tools.",
)
replace_once(
    "CHANGELOG.md",
    "- Enforced Python 3.10 tooling quality with pinned Ruff and mypy checks in hosted CI.",
    "- Enforced Python 3.10 tooling quality with pinned Ruff and mypy checks in hosted CI.\n- Hardened stdlib XML validation with a 1 MB input ceiling and explicit DTD/entity rejection.\n- Reduced generated-project payload by stripping template-maintainer checker-development and policy-coverage tooling while pruning their static-workflow references.",
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
