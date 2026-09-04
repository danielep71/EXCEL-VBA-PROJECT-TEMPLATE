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

Run the portable commands locally:

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root . \
  --output test-results/static-checks.json \
  --summary test-results/static-checks.md
```

The first command exercises a passing fixture, one deliberately degraded
fixture for each of the 21 rules, direct malformed-YAML and malformed-XML branch
fixtures, deterministic JSON and Markdown rendering, and read-only execution.
The second command validates the current tracked tree, prints a readable result,
and writes optional machine-readable evidence.

The `generated-vba-contract` rule resolves the applicable profile contract and
requires its registered, tracked façade, core, and test assets. Its JSON evidence
records the selected profile, role minima, observed role counts, and mandatory
component paths. In template mode it validates all three supported contracts.
The `issue-forms` rule validates the three canonical intake forms, their
manifest-backed labels, required evidence fields, empty reusable assignees,
blank-issue policy and repository-specific private-security route.

Exit status `0` means every applicable rule passed, `1` means policy findings
were reported, and `2` means the checker could not complete. Reports contain no
timestamps, so identical commits and configurations produce identical bytes.

This gate validates repository evidence and exported VBA structure. It does not
execute Excel, compile a VBA project, prove numerical accuracy, exercise UI
state, or certify a release package. Profile and project gates retain those
responsibilities.

## Authoritative workflow validation

The hosted gate complements the portable YAML subset check with
[actionlint 1.7.12](https://github.com/rhysd/actionlint/releases/tag/v1.7.12),
whose release tag resolves to upstream commit
[`914e7df21a07ef503a81201c76d2b11c789d3fca`](https://github.com/rhysd/actionlint/commit/914e7df21a07ef503a81201c76d2b11c789d3fca).
The workflow downloads only
`actionlint_1.7.12_linux_amd64.tar.gz` and verifies the upstream-published
SHA-256
`8aca8db96f1b94770f1b0d72b6dddcb1ebb8123cb3712530b08cc387b349a3d8`
before extraction or execution. The version and digest are explicit workflow
constants; a failed download, digest, version, current-workflow check, fixture,
report, or artifact fails the terminal verdict.

`test_workflow_validation.py` requires that exact actionlint version. It
accepts the tracked workflows and a valid local composite action, and requires
rejection of malformed YAML, duplicate job IDs, an invalid job key, missing
local-action metadata, and a missing local Node entry point. Run it locally with
an independently verified binary:

```bash
python3 tools/test_workflow_validation.py \
  --root . \
  --actionlint /verified/path/to/actionlint \
  --summary test-results/workflow-validation.md
```

The release archive and checksum catalogue are authoritative upstream assets;
review both the [release notes](https://github.com/rhysd/actionlint/releases/tag/v1.7.12)
and [published checksums](https://github.com/rhysd/actionlint/releases/download/v1.7.12/actionlint_1.7.12_checksums.txt)
before changing either pin.

## Canonical repository initializer

`initialize_repository.py` converts a clean generated repository from template
mode to one explicit profile. It validates the complete substitution set,
defaults to a deterministic dry-run, applies only with `--apply`, removes
non-applicable/template-only content, resets inherited changelog history, and
supports an idempotent second run.

See [`docs/INITIALIZATION.md`](../docs/INITIALIZATION.md) for the token catalogue,
profile commands, optional and repeatable values, and the transparent manual
fallback. Exercise all three profile fixtures with:

```bash
python3 tools/initialize_repository.py --root . --self-test
```

## Rules

- Tools must fail clearly and return a non-zero status for a blocking result.
- Pin or document material runtime dependencies.
- Separate generated output from the script and keep transient output ignored.
- Record the exact command used by CI and release certification.
- Never embed credentials, personal paths, private data, or workstation-specific assumptions.
- Do not place production VBA, regression modules, examples, or GitHub workflow definitions here.

Workflow orchestration belongs under `.github/workflows/`; `tools/` contains the reusable logic those workflows call.

Delete this README only if real tools and equivalent maintainer documentation make the directory's role equally explicit.
