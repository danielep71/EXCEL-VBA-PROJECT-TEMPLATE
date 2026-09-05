# 🧩 Checker Development Contract

[![Runtime: single file](https://img.shields.io/badge/runtime-single--file-217346)](../tools/check_repo.py)
[![Dependencies: stdlib only](https://img.shields.io/badge/dependencies-stdlib%20only-success)](../tools/checker_development.py)
[![Coverage: semantic policy](https://img.shields.io/badge/coverage-semantic%20policy-6F42C1)](../tools/check_policy_coverage.py)

`tools/check_repo.py` remains the canonical portable checker delivered to generated repositories. **The single-file identity rule applies to that canonical checker only.** Its reviewed source is the distributable artifact, so there is no bundle transform that can drift from source. Focused sibling gates are development/runtime tools within the repository and may share private standard-library infrastructure.

## 🔧 Shared focused-gate primitives

`tools/_gatelib.py` owns the small cross-tool mechanics that are genuinely identical: Git subprocess wrappers, tracked-file enumeration, deterministic UTF-8/LF report writes, the common `--root` / `--output` / `--summary` / `--self-test` parser, and `run_gate`, the typed runner that owns the shared console, evidence and exit-code contract. Focused gates import those primitives instead of maintaining copies. Tool-specific `run_check`, `build_report`, `run_self_test`, semantic rules, fixtures and Markdown renderers remain local because their behavior and evidence schemas differ.

### `run_gate` ownership

`run_gate` centralizes only the orchestration that was provably identical across gates: self-test dispatch, report construction, canonical JSON serialization, Markdown summary writing, console output, and the `0` / `1` / `2` exit mapping. It never widens a gate's exception handling. Each caller passes its own operational-exception tuple, so a programming error still surfaces as a traceback instead of being reported as exit code `2`.

Gates that historically evaluated `--self-test` outside their operational handler reported failures as `SELF-TEST ERROR`; gates that evaluated it inside reported `ERROR`. `run_gate` preserves both wordings through `self_test_error_prefix`.

| Gate | Uses `run_gate` | Reason when excluded |
| --- | :---: | --- |
| `check_committed_whitespace.py` | yes | — |
| `check_local_actions.py` | yes | — |
| `check_release_semantics.py` | yes | — |
| `check_vba_conditionals.py` | yes | — |
| `check_vba_jumps.py` | yes | — |
| `check_vba_public_api.py` | yes | — |
| `checker_development.py` | yes | — |
| `policy_coverage_runner.py` | yes | — |
| `check_release.py` | no | atomic evidence writes and a console rendering distinct from its Markdown summary |
| `test_workflow_validation.py` | no | text-only report with no JSON evidence output |
| `initialize_repository.py` | no | repository provisioning CLI, not a focused report gate |
| `check_repo.py` | no | self-contained distributable that must not import `_gatelib.py` |

`checker_development.py` enforces this table. Every tool defining a top-level `main` must be either a declared `run_gate` consumer or a documented exclusion; adding a gate without updating the declaration fails the contract, as does an excluded tool quietly adopting the runner. Sixteen independent unit tests exercise the runner's CLI flags, defaults and `--help`, self-test dispatch, both self-test diagnostic prefixes, pass/fail/operational exits, deterministic JSON and Markdown evidence, report-write failures, and the propagation of non-operational exceptions.

`tools/check_repo.py` must never import `_gatelib.py`. Generated repositories retain `_gatelib.py` for the focused operational gates, while the canonical checker remains independently copyable and executable as one standard-library-only file. `checker_development.py` enforces this ownership boundary in the canonical template. The checker-development workflow, this document, and the `policy_coverage_*` semantic-coverage harness are template-maintainer assets and are removed by initialization rather than shipped into generated projects.

The public API gate also uses `check_vba_conditionals.py` to evaluate the same three supported compilation environments. It checks name collisions only where declarations can coexist, and requires one manifest declaration row plus a `# SIG` record for every distinct reachable signature. Unknown conditions fail closed. These are static models, not evidence of Excel runtime certification.

### Supported invocation mode

The supported focused-tool interface is **path execution from the repository checkout**, for example `python3 tools/check_vba_public_api.py --root . --self-test`. Python places the executed script's directory on `sys.path`, which is the declared mechanism by which focused sibling gates resolve the private `_gatelib.py` module.

`python -m tools.<module>` is **not** part of the supported contract: `tools/` is not a public Python package and no package-installation interface is promised. If module-mode execution is added later, it must be introduced deliberately with package-aware imports and fixtures for both invocation modes rather than relying on incidental interpreter path behavior. The canonical `check_repo.py` remains unaffected because it has no sibling import.

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

The contract directly tests representative YAML, GitHub-style Markdown anchors, EditorConfig parsing, VBA lexical stripping, Markdown/console serialization, CLI flags and operational exit-code mapping. The full `check_repo.py --self-test` and semantic policy-coverage matrix remain separate higher-level gates. Hosted CI keeps both template-maintainer contracts together in `checker-development.yml`; the operational `static-checks.yml` deliberately remains free of template-only policy-coverage tooling so the generated workflow is self-contained.

## 📦 Portability and artifact identity

The runtime checker may import only Python standard-library modules and must not use relative imports. No `pip`, package manager, virtual environment, generated package tree or network dependency is required in a generated VBA repository.

Every development-contract report records the SHA-256 of `tools/check_repo.py`. Because the reviewed source and shipped artifact are the same bytes, reproducibility is an identity operation (`build_transform = none`) rather than a hidden build step.

## 🔁 Change procedure

1. Change `tools/check_repo.py`, `_gatelib.py`, and/or the focused gate that owns the affected behavior. Keep `check_repo.py` independent of `_gatelib.py`.
2. Run `python3 tools/checker_development.py --root . --self-test`.
3. Run `python3 tools/check_repo.py --root . --self-test`.
4. Run `python3 tools/check_policy_coverage.py --root . --self-test` so every canonical blocking finding remains exercised.
5. Run the normal repository gate and require successful hosted terminal verdicts from both `Checker development` and `Repository integrity`.
6. If a deliberate internal boundary, CLI contract or canonical check order changes, update this document and `checker_development.py` in the same reviewed change.

## 🚫 Non-goals

This contract does not replace authoritative GitHub Actions validation, Excel/VBA compilation, runtime regression tests, numerical assurance, UI-state tests, release evidence or profile-specific specialist controls.
