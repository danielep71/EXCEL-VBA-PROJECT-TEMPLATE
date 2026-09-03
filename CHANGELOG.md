<div align="center">

# 📜 Changelog

### Release history for {{PROJECT_NAME}}

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **{{PROJECT_NAME}}** are documented here.

This changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It records
released behavior and material unreleased changes; it is not a commit log, issue
tracker, or substitute for release evidence.

Define the versioned public surface—API, behavior, defaults, errors, data
formats, compatibility, and supported environments—in maintained project
documentation before the first functional release.

---

## 🧭 Maintenance policy

- Add material changes under **Unreleased** in the same pull request as the
  behavior or documentation they describe.
- Write from the user's perspective: describe the observable result, contract,
  compatibility impact, and migration need.
- Link the owning issue or pull request when it contains useful engineering
  detail.
- Keep entries concise; do not duplicate implementation notes already preserved
  in source, issues, or technical documentation.
- Record only validation actually performed. State skipped environments and
  known limitations plainly.
- Move Unreleased entries into a dated version section during release.
- Do not edit a published release entry except to correct a demonstrable factual
  or link error; annotate material corrections instead of rewriting history.
- Never claim that a tag, binary, workbook, hash, test run, or environment was
  certified unless the evidence binds it to the released source.

See [CONTRIBUTING.md](CONTRIBUTING.md) for change and evidence requirements and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

<a id="date-and-version-rules"></a>

### Date and version rules

| Rule | Standard |
|---|---|
| Version | `MAJOR.MINOR.PATCH`, without the leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |
| Date | Gregorian calendar date in ISO `YYYY-MM-DD` format |
| Ordering | Unreleased first; released versions newest to oldest |
| Comparison | Unreleased → latest tag; each release → preceding tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |
| Pre-release | State maturity and compatibility boundaries explicitly |

A repository may remain below `1.0.0` while its supported surface is still
forming. Pre-release status does not excuse undocumented breaking changes.

<details>
<summary><strong>Entry categories</strong></summary>

<br>

| Category | Use for |
|---|---|
| **Added** | New supported capabilities, APIs, files, or tests |
| **Changed** | Changes to existing behavior, contracts, tooling, or documentation |
| **Deprecated** | Supported behavior scheduled for removal |
| **Removed** | Removed capabilities or compatibility |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Documentation** | Material documentation-only changes |
| **Validation** | Evidence actually produced |
| **Compatibility** | Upgrade or migration effects |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories needed by a release.

</details>

---

<a id="unreleased"></a>

## [Unreleased]

<!--
Keep only categories that contain entries. Write observable outcomes, not commit
messages. Replace every registered template token before publishing.
-->

### Added

- Added a classified double-brace token schema and deterministic,
  dry-run-first repository initializer with atomic validation and all-profile
  fixtures.
- Added a profile-driven, dependency-free repository-quality gate with
  deterministic JSON and Markdown evidence plus positive and degraded self-tests.
- Added the canonical root README with profile selection, generated-repository
  initialization, structure, source policy and quality boundaries.
- Added a neutral, importable VBA façade/core starter, a fixed public API
  manifest, a state-safe example, and a deterministic four-case regression
  harness with equality, tolerance, expected-error, environment, completeness,
  and cleanup reporting.
- Added per-profile substantive VBA contracts and full-tree fixtures proving
  that README-only trees or missing façade/core/test components fail while
  optional example removal remains valid.

### Changed

- Replaced the progress-oriented implementation plan with an exact-snapshot code
  review, weighted template-readiness score, P1/P2/P3 finding register and
  findings-driven certification sequence.
- Rebuilt the canonical README from a seven-repository benchmark, combining a
  premium identity block, quick navigation, profile-aware initialization,
  architecture, assurance boundaries, recovery, security and release guidance.
- Upgraded the canonical static-check workflow with bounded execution,
  non-persistent checkout credentials, deterministic JSON and Markdown
  artifacts, rerun-safe evidence names and an explicit terminal verdict.

---

<!--
Release procedure:
1. Move applicable Unreleased entries under: ## [X.Y.Z] - YYYY-MM-DD
2. Remove empty categories.
3. Add comparison links after the first release:
   [Unreleased]: https://github.com/OWNER/REPOSITORY/compare/vX.Y.Z...HEAD
   [X.Y.Z]: https://github.com/OWNER/REPOSITORY/releases/tag/vX.Y.Z
4. Recreate an empty Unreleased section at the top.
-->
