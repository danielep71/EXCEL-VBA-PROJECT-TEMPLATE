# 🔐 Supply-Chain Assurance

This document owns the repository's public supply-chain analysis layer.
It complements the deterministic repository gates and the controlled dependency
update policy; it does not replace either one.

## Controls

| Control | Purpose | Trusted execution boundary |
| --- | --- | --- |
| CodeQL | Static security analysis of maintained Python and JavaScript source | `push` to `main` / `release/**`, schedule, or explicit manual dispatch |
| Dependabot | Weekly discovery of GitHub Actions updates | Proposal-only pull requests; no automatic approval or merge |
| OpenSSF Scorecard | Public repository supply-chain posture signal and SARIF evidence | `push` to `main` / `release/**`, branch-protection changes, schedule, or explicit manual dispatch |

All external Actions remain pinned to full immutable commit SHAs with audited
semantic-version comments. Dependabot may propose a new Actions revision, but a
proposal is not trusted merely because a bot opened it: the source/tag must be
reviewed under [DEPENDENCY_UPDATES.md](DEPENDENCY_UPDATES.md), the resulting
workflow must retain immutable pins, and merge remains manual.

## Permission model

The CodeQL and Scorecard publication workflows do not use `pull_request` or
`pull_request_target`. Their write permissions are therefore unavailable to
untrusted pull-request code. CodeQL receives only `contents: read` plus
`security-events: write` for analysis publication. Scorecard receives
`contents: read`, `security-events: write`, and `id-token: write`; the latter is
used only for Scorecard result publication. The repository-quality gate rejects
`pull_request_target` and rejects write-capable permissions in any workflow that
runs on `pull_request`.

Checkout credentials are not persisted. No supply-chain workflow has a source,
release, label, issue, branch, or repository mutation step.

## Evidence boundary

A successful CodeQL run means the configured CodeQL queries completed against
the exact checked source. A successful Scorecard run means the configured
Scorecard analysis completed and its SARIF/public result was published. A
Dependabot proposal means only that GitHub detected a candidate dependency
update.

None of those outcomes proves VBA compilation, Excel runtime behavior,
numerical accuracy, UI cleanup, packaging integrity, or release certification.
Those claims remain owned by their existing specialist and release evidence.
Likewise, a public Scorecard score is not the repository's project-quality
score.

## Review and maintenance

Pull-request review uses the existing read-only deterministic gates; analyzer
publication is deferred until the reviewed change reaches a trusted branch.

Before merging a CodeQL, Scorecard, or dependency-workflow change:

1. resolve every proposed Action release tag to its upstream commit and record
   the immutable SHA;
2. review permissions, triggers, checkout behavior, new network activity, and
   publication surfaces;
3. run repository integrity, authoritative workflow validation, checker
   development, and the dedicated supply-chain workflow fixtures;
4. preserve manual dependency approval and rollback under
   [DEPENDENCY_UPDATES.md](DEPENDENCY_UPDATES.md);
5. after integration into a trusted branch, retain the successful CodeQL and
   Scorecard workflow runs as provider evidence for that exact revision.

If a security analyzer is unavailable, misconfigured, or denied publication
permissions, its workflow is non-green. Do not convert an unavailable result
into a pass or weaken deterministic gates to recover a public score.
