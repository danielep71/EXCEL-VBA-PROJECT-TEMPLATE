# Tools

`tools/` contains deterministic maintainer tooling used to validate, package, or produce evidence for the repository.

Appropriate contents include:

- VBA static checks and exported-source validators;
- formatting and documentation-link checks;
- release-provenance and hashing utilities;
- fixture or report generators whose inputs and outputs are documented; and
- local wrappers that reproduce a CI gate.

## Rules

- Tools must fail clearly and return a non-zero status for a blocking result.
- Pin or document material runtime dependencies.
- Separate generated output from the script and keep transient output ignored.
- Record the exact command used by CI and release certification.
- Never embed credentials, personal paths, private data, or workstation-specific assumptions.
- Do not place production VBA, regression modules, examples, or GitHub workflow definitions here.

Workflow orchestration belongs under `.github/workflows/`; `tools/` contains the reusable logic those workflows call.

Delete this README only if real tools and equivalent maintainer documentation make the directory's role equally explicit.
