#!/usr/bin/env python3
"""One-shot removal-boundary markers for template-maintainer documentation links."""
from pathlib import Path

root = Path(__file__).resolve().parents[1]


def replace_once(relative: str, old: str, new: str) -> None:
    path = root / relative
    text = path.read_text(encoding="utf-8")
    if text.count(old) != 1:
        raise RuntimeError(f"{relative}: expected one anchor, found {text.count(old)}")
    path.write_text(text.replace(old, new, 1), encoding="utf-8", newline="\n")


replace_once(
    "README.md",
    '''For checker maintenance, the independent\n[`CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md) contract protects the\nsingle-file, standard-library runtime and parser/reporter development boundaries.\n''',
    '''<!-- template:remove:start -->\nFor checker maintenance, the independent\n[`CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md) contract protects the\nsingle-file, standard-library runtime and parser/reporter development boundaries.\n<!-- template:remove:end -->\n''',
)
replace_once(
    "CONTRIBUTING.md",
    '''For checker changes, additionally follow\n[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md). For release-evidence\nschemas and exact-SHA binding, use\n''',
    '''<!-- template:remove:start -->\nFor checker changes, additionally follow\n[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md).\n<!-- template:remove:end -->\nFor release-evidence schemas and exact-SHA binding, use\n''',
)
replace_once(
    "CONTRIBUTING.md",
    '| Checker changes | [`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md) |\n',
    '<!-- template:remove:start -->\n| Checker changes | [`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md) |\n<!-- template:remove:end -->\n',
)
replace_once(
    "RELEASING.md",
    '''For changes to checker behavior, also run the checker-development and semantic\npolicy-coverage contracts documented in\n[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md).\n''',
    '''<!-- template:remove:start -->\nFor changes to checker behavior in the canonical template, also run the\nchecker-development and semantic policy-coverage contracts documented in\n[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md).\n<!-- template:remove:end -->\n''',
)
replace_once(
    "docs/README.md",
    '| Portable checker development boundaries and independent tests | [`CHECKER_DEVELOPMENT.md`](CHECKER_DEVELOPMENT.md) | Link from tooling/contribution guidance |\n',
    '<!-- template:remove:start -->\n| Portable checker development boundaries and independent tests | [`CHECKER_DEVELOPMENT.md`](CHECKER_DEVELOPMENT.md) | Link from tooling/contribution guidance |\n<!-- template:remove:end -->\n',
)

Path(__file__).unlink()
