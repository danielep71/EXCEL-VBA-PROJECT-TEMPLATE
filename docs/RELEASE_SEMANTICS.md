# 🔖 Release Semantics Contract

[![SemVer: strict](https://img.shields.io/badge/SemVer-strict-3f4551)](https://semver.org/spec/v2.0.0.html)
[![Changelog: ordered](https://img.shields.io/badge/changelog-ordered-0969da)](../CHANGELOG.md)
[![Dates: Gregorian](https://img.shields.io/badge/dates-Gregorian-217346)](../CHANGELOG.md#date-and-version-rules)
[![Gate: fail closed](https://img.shields.io/badge/gate-fail%20closed-success)](../tools/check_release_semantics.py)

This document defines release-only version and changelog semantics for the
canonical template and initialized repositories. The generic repository checker
keeps only broad structural checks; release correctness is enforced separately
by `tools/check_release_semantics.py` and `tools/check_release.py`.

## SemVer contract

`VERSION` contains one SemVer 2.0.0 value without a leading `v`.

- Major, minor, and patch numeric identifiers never contain leading zeros.
- Pre-release identifiers follow SemVer precedence exactly.
- Numeric pre-release identifiers never contain leading zeros (`rc.01` is
  invalid; `rc.1` is valid).
- Build metadata is accepted but does not affect precedence.
- A stable release has higher precedence than a pre-release with the same core
  version.

The tag is formed by adding the lower-case `v` prefix to `VERSION`.

## Changelog contract

`CHANGELOG.md` contains exactly one `## [Unreleased]` heading. Dated release
headings use exactly:

```text
## [MAJOR.MINOR.PATCH] - YYYY-MM-DD
```

Pre-release versions use the same heading form with their valid SemVer suffix.
Every date must be a real Gregorian calendar date.

Dated releases appear newest to oldest by **full SemVer precedence**, not by
lexical text or date alone. Duplicate release versions are invalid. When
`VERSION` is not the generated-project development sentinel `0.0.0`, it must
match the newest dated release heading during a release candidate.

## Comparison-link policy

When at least one release exists:

- `[Unreleased]` compares the latest release tag to `HEAD`:

  ```text
  [Unreleased]: https://github.com/OWNER/REPOSITORY/compare/vLATEST...HEAD
  ```

- the initial release links to its immutable release tag:

  ```text
  [1.0.0]: https://github.com/OWNER/REPOSITORY/releases/tag/v1.0.0
  ```

- every later release compares the immediately preceding release tag to the new
  release tag:

  ```text
  [1.1.0]: https://github.com/OWNER/REPOSITORY/compare/v1.0.0...v1.1.0
  ```

Missing, duplicated, stale, or mismatched comparison links fail release
semantics.

## Validation

Run the deterministic policy fixtures:

```bash
python3 tools/check_release_semantics.py --root . --self-test
```

Validate the current tree and retain evidence:

```bash
python3 tools/check_release_semantics.py \
  --root . \
  --output test-results/release-semantics.json \
  --summary test-results/release-semantics.md
```

The self-test covers valid stable and pre-release versions, numeric pre-release
leading zeros, SemVer precedence, duplicate and out-of-order releases,
impossible dates, `VERSION`/heading disagreement, and missing or incorrect
comparison links.

A release candidate is not eligible for tagging unless this gate, the executable
release-integrity gate, the repository gates, and all applicable runtime evidence
are green for the same candidate source.
