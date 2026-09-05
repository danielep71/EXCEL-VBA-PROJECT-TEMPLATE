#!/usr/bin/env python3
"""One-shot coverage correction for bounded stdlib XML validation."""
from pathlib import Path

path = Path(__file__).resolve().parent / "check_repo.py"
text = path.read_text(encoding="utf-8")
old = '''def _validate_xml_text(path: str, text: str) -> dict[str, Any] | None:\n    try:\n        size = len(text.encode("utf-8"))\n    except UnicodeError as error:\n        return finding(path, f"Cannot encode XML as UTF-8: {error}")\n    if size > MAX_XML_BYTES:\n'''
new = '''def _validate_xml_text(path: str, text: str) -> dict[str, Any] | None:\n    size = len(text.encode("utf-8"))\n    if size > MAX_XML_BYTES:\n'''
if text.count(old) != 1:
    raise RuntimeError(f"XML encode branch anchor count: {text.count(old)}")
text = text.replace(old, new, 1)

anchor = '''def _degrade_structured_xml_doctype(root: Path) -> None:\n'''
insert = '''def _degrade_structured_xml_encoding(root: Path) -> None:\n    path = root / "docs/invalid-encoding.xml"\n    _write_fixture(path, b"<repository>\\xff</repository>\\n")\n    _run_git(root, "add", path.relative_to(root).as_posix())\n\n\n'''
if text.count(anchor) != 1:
    raise RuntimeError(f"XML encoding fixture anchor count: {text.count(anchor)}")
text = text.replace(anchor, insert + anchor, 1)

old_cases = '''    ("structured-xml", "structured-data", _degrade_structured_xml),\n    ("structured-xml-doctype", "structured-data", _degrade_structured_xml_doctype),\n'''
new_cases = '''    ("structured-xml", "structured-data", _degrade_structured_xml),\n    ("structured-xml-encoding", "structured-data", _degrade_structured_xml_encoding),\n    ("structured-xml-doctype", "structured-data", _degrade_structured_xml_doctype),\n'''
if text.count(old_cases) != 1:
    raise RuntimeError(f"XML branch-case anchor count: {text.count(old_cases)}")
text = text.replace(old_cases, new_cases, 1)
path.write_text(text, encoding="utf-8", newline="\n")
Path(__file__).unlink()
