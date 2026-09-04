# Documentation

`docs/` contains durable project documentation that is too detailed or specialized for the root README.

Use it for:

- public API and behavioral contracts;
- architecture and design decisions;
- numerical conventions and validation methods;
- compatibility, migration, and operational notes; and
- implementation or release plans that are intentionally part of the public record.

Keep portfolio governance files at the repository root: `README.md`, `CHANGELOG.md`, `CONTRIBUTING.md`, `INSTALLATION.md`, `SECURITY.md`, `CODE_OF_CONDUCT.md`, and `RELEASING.md`.

## Template documents

| Document | Purpose |
| --- | --- |
| [`INITIALIZATION.md`](INITIALIZATION.md) | Canonical token schema, profile selection, dry-run/apply procedure, and manual fallback |
| [`POST_CREATION_CHECKLIST.md`](POST_CREATION_CHECKLIST.md) | Live GitHub settings, label reconciliation, protection, and read-back evidence after generation |
| [`PUBLIC_API.txt`](PUBLIC_API.txt) | Machine-checked declarations supported by public-role VBA components |
| [`RELEASE_EVIDENCE.md`](RELEASE_EVIDENCE.md) | External evidence JSON, profile checks, asset manifests, and exact-SHA release validation |
| [`REPOSITORY_STRUCTURE.md`](REPOSITORY_STRUCTURE.md) | Canonical directory ownership and permitted profile alternatives |
<!-- template:remove:start -->
| [`PORTFOLIO_AUDIT.md`](PORTFOLIO_AUDIT.md) | Evidence used to design this portfolio template; remove it from a generated project unless that project deliberately maintains the portfolio audit |
<!-- template:remove:end -->

Do not duplicate the same contract in the root README, Wiki, and `docs/`. Choose one authoritative location and link to it. A Wiki may provide navigation or extended guidance, but versioned behavior belongs in the repository.

Delete this README only if an equivalent documentation index replaces it.
