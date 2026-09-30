# 🛠️ Tools and Validation Guide

[![Runtime: Python 3.10](https://img.shields.io/badge/runtime-Python%203.10-1D76DB)](#canonical-repository-quality-gate)
[![Governance: minimal](https://img.shields.io/badge/governance-minimal-0969da)](../docs/INITIALIZATION.md#governance-tiers)

`tools/` holds the four validation gates of the **minimal governance tier**. All
of them are standard-library Python 3.10 with no dependencies to install, and
all are deterministic: identical commits produce identical reports.

| Gate | Checks |
| --- | --- |
| `check_repo.py` | Repository integrity, exported VBA structure, pinned workflow actions, version and changelog shape |
| `check_vba_public_api.py` | Explicit public VBA surface against `docs/PUBLIC_API.txt` |
| `check_vba_jumps.py` | `GoTo`, `GoSub` and `Resume` targets resolve inside their own procedure |
| `check_release_semantics.py` | SemVer, changelog order and comparison links |

Two helpers support them: `_gatelib.py` owns the shared Git, report and
command-line mechanics, and `check_vba_conditionals.py` evaluates conditional
compilation for the public-API gate. `check_repo.py` imports neither, so it
stays a self-contained file.

Every gate exits `0` when all rules pass, `1` when it reports findings and `2`
when it could not complete. `--self-test` exercises each gate's own passing and
deliberately degraded fixtures.

Release evidence, provenance, host-evidence validation, documentation-drift and
external-link checks, local-action validation and the initializer belong to the
full tier. Adopt them with the documented
[upgrade to the full tier](../docs/INITIALIZATION.md#governance-tiers).

<a id="canonical-repository-quality-gate"></a>

## Run the gates

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root .
python3 tools/check_vba_public_api.py --root . --self-test
python3 tools/check_vba_public_api.py --root .
python3 tools/check_vba_jumps.py --root . --self-test
python3 tools/check_vba_jumps.py --root .
python3 tools/check_release_semantics.py --root . --self-test
python3 tools/check_release_semantics.py --root .
```

Add `--output <file>.json --summary <file>.md` to any gate for machine-readable
and Markdown evidence. The hosted *Static repository checks* workflow runs the
same commands on every pull request.

These gates validate repository evidence and exported VBA source. They do not
execute Excel, compile a VBA project or prove numerical accuracy; the regression
suite in `tests/` and a real Excel run keep those responsibilities.

## Rules

- Tools must fail clearly and return a non-zero status for a blocking result.
- Keep transient output such as `test-results/` out of version control.
- Never embed credentials, personal paths, private data, or workstation-specific assumptions.
- Do not place production VBA, regression modules, examples, or GitHub workflow definitions here.

Workflow orchestration belongs under `.github/workflows/`; `tools/` contains the logic those workflows call.
