# 🧪 Tests and Evidence

`tests/` is the canonical home for all verification source and stable test data.

Use subdirectories only when they contain real material:

| Location | Contents |
| --- | --- |
| `tests/modules/` | Exported VBA regression modules and release-certification entry points |
| `tests/fixtures/` | Deterministic inputs, manifests, and reusable test workbooks |
| `tests/expected/` | Reviewed expected outputs or golden files |

## Neutral regression harness

Import `modules/ProjectTests.bas` after `ProjectCore` and `ProjectFacade`, compile
the VBA project, and run `ProjectTests.RunProjectTests`. The baseline executes
four deterministic cases and six assertions covering exact equality, tolerance,
the public expected-error contract, and repeatability. It reports environment,
case/assertion/failure counts, completeness, and cleanup to the Immediate window.

Success ends with:

```text
RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS
```

<!-- template:profile:library:start -->
The library profile adds `modules/TextTests.bas`. Import it after `TextCore` and
`TextFacade`, then run `TextTests.RunTextTests`. Its 26 cases and 123 assertions
cover every supported `TextFacade` member: normal and boundary behavior, each
documented invalid-argument rule with its number, source and description, the
Windows and macOS path rules, and repeatability. Success ends with:

```text
RESULT=PASS; completeness=COMPLETE; cases=26; assertions=123; failures=0; cleanup=PASS
```
<!-- template:profile:library:end -->

<!-- template:profile:ui-component:start -->
The ui-component profile adds `modules/ProgressTests.bas`. Import it after
`ProgressCore` and `ProgressFacade`, then run `ProgressTests.RunProgressTests`.
Its 12 cases cover begin, update, caller-state restoration, re-entry refusal,
calls without a session, cancellation, every invalid argument, the elapsed-time
seam, repeated sessions and recovery. A negative control breaks a restored
property and requires the check to detect it. Cursor checks run on Windows
only, so success ends with:

```text
RESULT=PASS; completeness=COMPLETE; cases=12; assertions=79; failures=0; cleanup=PASS
```

On macOS the expected count is 75 assertions. The status bar flickers while the
suite runs; the suite restores the values it found before cleanup is verified.
<!-- template:profile:ui-component:end -->

Any assertion, unexpected error, dirty start, incomplete execution, or cleanup
failure is non-passing. The harness changes no Excel state; cleanup verifies its
owned run flag and checks that calculation, display alerts, events and screen
updating match their pre-run values.

<!-- template:tier:full:start -->
The optional [Windows/Excel evidence interface](../docs/EXCEL_EVIDENCE.md)
records this harness output, source identity and host environment with explicit
manual/automated execution. Its validator does not execute Excel.
<!-- template:tier:full:end -->

## Rules

- Test modules are never part of the production import set.
- Name the complete regression and release-certification entry points in `CONTRIBUTING.md` and `RELEASING.md`.
- Keep fixtures synthetic, anonymized, or explicitly redistributable.
- Bind numerical or platform-sensitive evidence to the exact candidate commit and environment.
- Do not commit transient output, logs, caches, or locally generated workbooks; use ignored output directories or workflow artifacts.
- New repositories use `tests/`, not `test/`.

A legacy `test/` directory is legitimate only when an existing public path, build script, or release contract makes migration materially disruptive. Document that exception and never keep both `test/` and `tests/`.

Delete this README only if the real harness and equivalent test documentation make the directory's role equally explicit.
