# 🧩 Checker Development Contract

[![Runtime: single file](https://img.shields.io/badge/runtime-single--file-217346)](../tools/check_repo.py)
[![Dependencies: stdlib only](https://img.shields.io/badge/dependencies-stdlib%20only-success)](../tools/checker_development.py)
[![P2-08](https://img.shields.io/badge/P2--08-checker%20development-6F42C1)](IMPLEMENTATION_PLAN.md#p2-hardening)

`tools/check_repo.py` remains the canonical portable checker delivered to generated repositories. The development model intentionally avoids a package-manager or multi-file runtime dependency: the reviewed source **is** the distributable artifact, so there is no bundle transform that can drift from source.

## 🧭 Internal boundaries

`tools/checker_development.py` parses the checker with Python AST and requires the following ordered ownership boundaries:

| Section | Sentinel | Responsibility |
| --- | --- | --- |
| Runtime core | `Repository` | repository I/O, findings and common primitives |
| Configuration | `_same_keys` | repository-profile schema and effective requirements |
| Repository policy | `check_required_paths` | generic repository, document, workflow and metadata rules |
| VBA policy | `_vba_paths` | VBE exports, VBA structure, roles, generated contracts and public surface |
| Reporting | `build_report` | deterministic report model, Markdown and console serialization |
| Fixtures | `_write_fixture` | positive/degraded synthetic repository fixtures and self-test |
| CLI | `parse_arguments` | supported command-line surface and exit behavior |

The development check fails if a boundary disappears, changes order, leaves a top-level definition outside the ordered sections, or changes the canonical policy-check sequence without an explicit update to the contract.

## 🧪 Independent tests

Maintainers can exercise parser and reporter behavior without running the full synthetic repository matrix:

```bash
python3 tools/checker_development.py --root . --self-test
```

The contract directly tests representative YAML, GitHub-style Markdown anchors, EditorConfig parsing, VBA lexical stripping, Markdown/console serialization, CLI flags and operational exit-code mapping. The full `check_repo.py --self-test` and P2-07 policy-coverage matrix remain separate higher-level gates.

## 📦 Portability and artifact identity

The runtime checker may import only Python standard-library modules and must not use relative imports. No `pip`, package manager, virtual environment, generated package tree or network dependency is required in a generated VBA repository.

Every development-contract report records the SHA-256 of `tools/check_repo.py`. Because the reviewed source and shipped artifact are the same bytes, reproducibility is an identity operation (`build_transform = none`) rather than a hidden build step.

## 🔁 Change procedure

1. Change `tools/check_repo.py` and any focused P2 gate required by the policy change.
2. Run `python3 tools/checker_development.py --root . --self-test`.
3. Run `python3 tools/check_repo.py --root . --self-test`.
4. Run `python3 tools/check_policy_coverage.py --root . --self-test` so every canonical blocking finding remains exercised.
5. Run the normal repository gate and require the hosted `Repository integrity` terminal verdict.
6. If a deliberate internal boundary, CLI contract or canonical check order changes, update this document and `checker_development.py` in the same reviewed change.

## 🚫 Non-goals

This contract does not replace authoritative GitHub Actions validation, Excel/VBA compilation, runtime regression tests, numerical assurance, UI-state tests, release evidence or profile-specific specialist controls.
