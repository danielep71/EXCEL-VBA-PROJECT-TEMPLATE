#!/usr/bin/env python3
"""One-shot assertion upgrade for generated template-only payload cleanup."""
from pathlib import Path

path = Path(__file__).resolve().parent / "initialize_repository.py"
text = path.read_text(encoding="utf-8")
old = '''            if any(\n                (fixture / path).exists()\n                for path in (\n                    "assets/social-preview.svg",\n                    "docs/IMPLEMENTATION_PLAN.md",\n                    "docs/PILOT_CERTIFICATION.md",\n                    "docs/PORTFOLIO_AUDIT.md",\n                )\n            ):\n                raise AssertionError(f"{profile} retained template-only files.")\n'''
new = '''            generated_config = json.loads(\n                (fixture / CONFIG_PATH).read_text(encoding="utf-8")\n            )\n            retained_template_only = {CANONICAL_SOCIAL_PREVIEW_PATH}\n            forbidden_template_only = set(\n                generated_config["placeholders"]["template_only_paths"]\n            ) - retained_template_only\n            retained = sorted(\n                path for path in forbidden_template_only if (fixture / path).exists()\n            )\n            if retained:\n                raise AssertionError(\n                    f"{profile} retained template-only files: {', '.join(retained)}"\n                )\n'''
if text.count(old) != 1:
    raise RuntimeError(f"initializer cleanup assertion anchor count: {text.count(old)}")
path.write_text(text.replace(old, new, 1), encoding="utf-8", newline="\n")
Path(__file__).unlink()
