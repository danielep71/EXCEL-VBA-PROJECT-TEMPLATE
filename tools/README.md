# Tools

`tools/` contains deterministic maintainer tooling used to validate, package, or produce evidence for the repository.

Appropriate contents include:

- VBA static checks and exported-source validators;
- formatting and documentation-link checks;
- release-provenance and hashing utilities;
- fixture or report generators whose inputs and outputs are documented; and
- local wrappers that reproduce a CI gate.

## Canonical repository-quality gate

`check_repo.py` is the dependency-free baseline gate for all three supported
repository profiles. Its versioned policy lives in
`.github/repository-profile.json`; generated repositories set `mode` to
`generated`, select `application`, `library`, or `ui-component`, and update the
declared paths and VBA component roles instead of editing checker logic.

Run the same commands locally that the hosted workflow runs:

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root . \
  --output test-results/repository-quality.json \
  --summary test-results/repository-quality.md
```

The first command exercises a passing fixture, one deliberately degraded
fixture for each of the 19 rules, deterministic JSON and Markdown rendering,
and read-only execution. The second command validates the current tracked tree,
prints a readable result, and writes optional machine-readable evidence.

Exit status `0` means every applicable rule passed, `1` means policy findings
were reported, and `2` means the checker could not complete. Reports contain no
timestamps, so identical commits and configurations produce identical bytes.

This gate validates repository evidence and exported VBA structure. It does not
execute Excel, compile a VBA project, prove numerical accuracy, exercise UI
state, or certify a release package. Profile and project gates retain those
responsibilities.

## Rules

- Tools must fail clearly and return a non-zero status for a blocking result.
- Pin or document material runtime dependencies.
- Separate generated output from the script and keep transient output ignored.
- Record the exact command used by CI and release certification.
- Never embed credentials, personal paths, private data, or workstation-specific assumptions.
- Do not place production VBA, regression modules, examples, or GitHub workflow definitions here.

Workflow orchestration belongs under `.github/workflows/`; `tools/` contains the reusable logic those workflows call.

Delete this README only if real tools and equivalent maintainer documentation make the directory's role equally explicit.
