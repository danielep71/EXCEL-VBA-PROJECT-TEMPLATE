# Release evidence contract

This contract defines the machine-readable evidence consumed by
`tools/check_release.py`. It binds version, changelog, tag, source checks,
Excel evidence, profile-specific assurance, and any staged binary assets to one
full candidate commit SHA.

## Why evidence stays outside the candidate

The evidence JSON is deliberately supplied separately from the candidate Git
tree. A committed file cannot contain the hash of the commit that contains that
file: changing the recorded hash changes the commit again. Keep the final
evidence in the release review, a protected workflow artifact, or the draft
release record. The checker proves that every record names the exact candidate;
the retention mechanism preserves that checked file.

## Policy

`.github/release-policy.json` is the versioned policy. Every profile requires:

- `repository-integrity` with an HTTPS run URL;
- `vba-compile` with the tested Excel environment; and
- `regression` with its entry point, environment, positive case and assertion
  counts, zero failures, complete execution, and passing cleanup.

Additional required checks are profile-specific:

| Profile | Required evidence |
| --- | --- |
| `library` | Public API and caller-contract evidence |
| `ui-component` | UI state, cleanup, recovery, DPI/accessibility, and lifecycle evidence |
| `application` | Startup, shutdown, upgrade, recovery, packaging, and end-to-end smoke evidence |
| `template` | All three generated-profile pilots and live branch/tag governance evidence |

All checks use `status: "PASS"`, a non-empty `detail`, and the same
`candidate_sha`. Project-specific checks may be added with lower-case,
hyphen-separated identifiers and the same three base fields.

## Evidence JSON

The top-level object contains exactly these fields:

```json
{
  "schema_version": 1,
  "version": "1.0.0",
  "tag": "v1.0.0",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "profile": "template",
  "distribution": "source-only",
  "checks": {},
  "assets": []
}
```

The complete `checks` object depends on the selected profile. A regression
record has this mandatory shape in addition to `status`, `candidate_sha`, and
`detail`:

```json
{
  "entry_point": "ProjectTests.RunProjectTests",
  "environment": "host=Microsoft Excel; version=16.0; os=Windows (64-bit) NT 10.00; office=64-bit; runtime=VBA7+",
  "cases": 4,
  "assertions": 6,
  "failures": 0,
  "completeness": "COMPLETE",
  "cleanup": "PASS"
}
```

The record reports evidence; it does not manufacture it. Copy counts and
environment details from the exact Excel run, and retain its raw output beside
the checked JSON.

## Source-only and binary distributions

A library or template release is source-only by default. Set `distribution` to
`source-only`, keep `assets` empty, and omit the asset manifest. UI-component
and application profiles may do the same.

The template profile deliberately retains its registered placeholders,
template-only blocks, canonical repository identity, and construction history.
Its `repository-integrity`, `generated-profile-pilots`, and `live-governance`
evidence proves those controls are intentional and usable. Generated project
profiles remain subject to the stricter unresolved-token, template-identity,
and inherited-construction-history rejection rules.

Optional UI/application binaries must match an allowed `dist/` glob in the
policy. Each asset record contains exactly:

```json
{
  "path": "dist/example.xlsm",
  "sha256": "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "package_test": "PASS"
}
```

Binary distribution requires a separate UTF-8 manifest, sorted by path, with
two spaces between each lower-case digest and safe repository-relative path:

```text
aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa  dist/example.xlsm
```

The checker requires exact equality among the evidence asset list, manifest,
and bytes staged under the candidate root. A digest establishes download
identity. It does not prove that a manually built workbook was produced from
the exported source; compile, regression, package-test, and build-process
evidence carry that separate claim.

## Commands

First certify the checker itself:

```bash
python3 tools/check_release.py --root . --self-test \
  --summary test-results/release-self-test.md
```

Before tagging, validate an external evidence file and any staged assets:

```bash
candidate_sha="$(git rev-parse HEAD)"
python3 tools/check_release.py \
  --root . \
  --tag "v$(tr -d '\r\n' < VERSION)" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --output test-results/release-integrity.json \
  --summary test-results/release-integrity.md
```

Add `--asset-manifest ../release-assets.sha256` for a binary distribution.
After creating the annotated tag locally, repeat the command with
`--require-tag-ref`. That final mode requires the tag object to be annotated and
to resolve to the same candidate SHA. Any non-zero result blocks publication.

## What the gate rejects

The deterministic self-test covers all four release profiles and rejects:

- a version/tag mismatch or `0.0.0` sentinel;
- an invalid or missing dated changelog release;
- unresolved template syntax, template identity, or inherited construction
  history in a generated project release;
- a template release carrying a generated-project initialization record or
  missing its pilot/governance evidence;
- missing evidence or missing profile-specific checks;
- evidence or assets bound to another candidate SHA;
- unapproved library or profile binaries;
- an absent manifest or incorrect asset digest; and
- a lightweight or moved tag.

The live `v*` ruleset provides the complementary server-side control: release
tags cannot be deleted, updated, or recreated outside the protected policy.
