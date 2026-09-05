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
| Ordering | Unreleased first; released versions newest to oldest by SemVer precedence |
| Comparison | Unreleased → latest tag; each later release → preceding tag; initial release → release tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |
| Pre-release | Strict SemVer identifiers; numeric identifiers have no leading zeros |

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

### Fixed

- Distinguish local reusable-workflow job calls from local action steps, and require tracked workflow files.
- Ignore quoted VBA text when checking jump targets, while retaining checks for executable jumps on the same line.
- Validate public API declarations across the supported conditional-compilation environments; accept mutually exclusive variants and require every distinct signature in the manifest.
- Restore bounded exponential retry delays in label synchronization and drift checks when `Retry-After` is absent or empty.

## [1.1.0] - 2026-09-05

### Added

- Added committed-candidate whitespace validation with explicit commit-range,
  root-commit, staged, unstaged, and committed-versus-working-tree fixtures.
- Added procedure-scoped VBA jump validation and nested conditional-compilation
  validation across the supported VBA6/VBA7 and Win32/Win64 environments.
- Added complete public-API extraction with normalized signature records,
  paired-property handling, collision detection, and strict explicit visibility.
- Added repository-local GitHub Action containment, tracked-state, metadata, and
  entrypoint validation alongside the pinned authoritative workflow parser.
- Added strict release-semantics validation for SemVer precedence, prerelease
  identifiers, changelog ordering, Gregorian dates, VERSION agreement, and
  comparison-link policy.
- Added deterministic semantic policy-branch assurance proving every canonical
  blocking finding site is exercised; the v1.1.0 branch baseline covers 175/175
  production finding sites.
- Added an independent checker-development contract that preserves
  `tools/check_repo.py` as a single-file, standard-library-only distributable
  while testing parser/reporting boundaries, CLI behavior, canonical check order,
  and artifact identity.
- Added read-only live issue-label drift detection with deterministic
  create/update/delete evidence, canonical-plan cross-checking, scheduled/manual
  monitoring, and retained JSON/Markdown evidence without a mutation path.
- Added an explicit Python 3.10 tooling baseline with pinned Ruff and mypy checks
  in the hosted repository-integrity workflow, including retained lint/type
  evidence and fail-closed terminal enforcement.

### Changed

- Extended the hosted repository-integrity gate so focused hardening checks have
  deterministic self-tests, exact-candidate evidence, artifact retention, and
  fail-closed terminal enforcement.
- Consolidated documentation around one authoritative owner per evolving
  contract. The root README is now a shorter first-use/navigation surface while
  installation, contribution, security, conduct, release, initialization,
  repository structure, checker development, and release evidence remain in
  their specialized maintained documents.
- Required generated repositories to retain the read-only label-drift control in
  addition to the existing trusted label-reconciliation workflow.
- Removed the completed temporary implementation plan and redirected durable
  governance references to maintained contracts and historical evidence.

### Fixed

- Disarmed the regression runner's procedure-level error handler before cleanup
  and summary reporting so a cleanup/reporting fault cannot re-enter `CleanExit`
  indefinitely through `RunFailed`.
- Made regression cleanup evidence substantive by comparing the relevant Excel
  application state with the pre-run snapshot rather than merely re-reading an
  immediately cleared re-entry flag.
- Consolidated duplicated focused-gate Git, report-output, tracked-file, and
  common CLI mechanics into the private standard-library-only `tools/_gatelib.py`,
  while keeping `tools/check_repo.py` explicitly self-contained.
- Removed residual unused imports exposed by the enforced Ruff baseline.
- Reduced Python checker complexity under a permanently enforced McCabe ceiling
  of 20 and normalized the VBA public-API checker so no style exception remains.
- Hardened stdlib XML validation with a bounded input size and fail-closed
  rejection of DTD/entity declarations before parsing.
- Split template-maintainer checker-development and semantic policy-coverage
  tooling from the operational tool payload retained by generated repositories.

### Compatibility

- The neutral starter VBA public API and profile architecture are unchanged from
  v1.0.0. The release hardens repository governance, validation, documentation,
  and automation without intentionally changing the starter consumer contract.

## [1.0.0] - 2026-09-04

### Added

- Added a dependency-free release-integrity gate, versioned profile policy,
  external evidence contract, SHA-256 asset-manifest validation, annotated-tag
  verification, and deterministic positive and negative fixtures for generated
  projects and the canonical template itself.
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
- Added structured bug, feature and documentation forms, private security
  routing, a post-creation provisioning checklist and static form fixtures.
- Added content-pinned authoritative GitHub Actions validation with positive,
  malformed-YAML, duplicate-schema and local-action fixtures, plus direct XML
  and YAML branches in the portable checker self-test.

### Changed

- Renamed the canonical repository to `EXCEL-VBA-PROJECT-TEMPLATE` and aligned
  its profile identity, security route, evidence links and template-identity
  rejection rule with the new URL.
- Replaced the progress-oriented implementation plan with an exact-snapshot code
  review, weighted template-readiness score, P1/P2/P3 finding register and
  findings-driven certification sequence.
- Rebuilt the canonical README from a seven-repository benchmark, combining a
  premium identity block, quick navigation, profile-aware initialization,
  architecture, assurance boundaries, recovery, security and release guidance.
- Upgraded the canonical static-check workflow with bounded execution,
  non-persistent checkout credentials, deterministic JSON and Markdown
  artifacts, rerun-safe evidence names and an explicit terminal verdict.
- Made profile and domain label selection a versioned repository policy that
  both the checker and trusted reconciliation workflow validate and consume.

[Unreleased]: https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/releases/tag/v1.0.0

---

<!--
Release procedure:
1. Move applicable Unreleased entries under: ## [X.Y.Z] - YYYY-MM-DD
2. Remove empty categories.
3. Set [Unreleased] to compare the new latest tag to HEAD.
4. For the initial release, link directly to its release tag. For every later
   release, compare the preceding tag to the new release tag.
5. Recreate an empty Unreleased section at the top.
-->
