# Tests

`tests/` is the canonical home for all verification source and stable test data.

Use subdirectories only when they contain real material:

| Location | Contents |
| --- | --- |
| `tests/modules/` | Exported VBA regression modules and release-certification entry points |
| `tests/fixtures/` | Deterministic inputs, manifests, and reusable test workbooks |
| `tests/expected/` | Reviewed expected outputs or golden files |

## Rules

- Test modules are never part of the production import set.
- Name the complete regression and release-certification entry points in `CONTRIBUTING.md` and `RELEASING.md`.
- Keep fixtures synthetic, anonymized, or explicitly redistributable.
- Bind numerical or platform-sensitive evidence to the exact candidate commit and environment.
- Do not commit transient output, logs, caches, or locally generated workbooks; use ignored output directories or workflow artifacts.
- New repositories use `tests/`, not `test/`.

A legacy `test/` directory is legitimate only when an existing public path, build script, or release contract makes migration materially disruptive. Document that exception and never keep both `test/` and `tests/`.

Delete this README only if the real harness and equivalent test documentation make the directory's role equally explicit.
