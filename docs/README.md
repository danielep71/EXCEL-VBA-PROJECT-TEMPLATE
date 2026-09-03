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
| [`REPOSITORY_STRUCTURE.md`](REPOSITORY_STRUCTURE.md) | Canonical directory ownership and permitted profile alternatives |
| [`PORTFOLIO_AUDIT.md`](PORTFOLIO_AUDIT.md) | Evidence used to design this portfolio template; remove it from a generated project unless that project deliberately maintains the portfolio audit |

Do not duplicate the same contract in the root README, Wiki, and `docs/`. Choose one authoritative location and link to it. A Wiki may provide navigation or extended guidance, but versioned behavior belongs in the repository.

Delete this README only if an equivalent documentation index replaces it.
