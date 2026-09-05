#!/usr/bin/env python3
"""One-shot generated static-workflow pruning for template-only policy coverage."""
from pathlib import Path

path = Path(__file__).resolve().parent / "initialize_repository.py"
text = path.read_text(encoding="utf-8")
anchor = '''def _reject_executable_placeholders(path: str, matches: list[Any]) -> None:\n    if matches and PurePosixPath(path).suffix.casefold() in EXECUTABLE_SUFFIXES:\n        raise InitializationError(\n            f"Placeholders are prohibited in executable or VBA file {path}."\n        )\n\n\n'''
helper = anchor + '''def _strip_template_maintenance_workflow(text: str) -> str:\n    lines = text.splitlines(keepends=True)\n    output: list[str] = []\n    skipping = False\n    for line in lines:\n        if line.startswith("      - name: Exercise policy-branch coverage determinism"):\n            skipping = True\n            continue\n        if skipping and line.startswith("      - name: Exercise positive and degraded checker paths"):\n            skipping = False\n        if skipping:\n            continue\n        if "test-results/policy-coverage." in line:\n            continue\n        if "POLICY_COVERAGE_" in line:\n            continue\n        if '"Policy-coverage self-test:' in line or '"Policy coverage:' in line:\n            continue\n        output.append(line)\n    return "".join(output)\n\n\n'''
if text.count(anchor) != 1:
    raise RuntimeError(f"initializer helper anchor count: {text.count(anchor)}")
text = text.replace(anchor, helper, 1)
old = '''        if path == "CHANGELOG.md":\n            rendered = _reset_changelog(rendered)\n        elif path == "VERSION":\n            rendered = "0.0.0\\n"\n'''
new = '''        if path == "CHANGELOG.md":\n            rendered = _reset_changelog(rendered)\n        elif path == "VERSION":\n            rendered = "0.0.0\\n"\n        elif path == ".github/workflows/static-checks.yml":\n            rendered = _strip_template_maintenance_workflow(rendered)\n'''
if text.count(old) != 1:
    raise RuntimeError(f"initializer render anchor count: {text.count(old)}")
text = text.replace(old, new, 1)
path.write_text(text, encoding="utf-8", newline="\n")
Path(__file__).unlink()