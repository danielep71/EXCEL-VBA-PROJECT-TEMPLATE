# 🔐 Supply-Chain Assurance

This document owns the repository's public supply-chain analysis layer.
It complements the deterministic repository gates and the controlled dependency
update policy; it does not replace either one.

## Controls

| Control | Purpose | Trusted execution boundary |
| --- | --- | --- |
| CodeQL | Static security analysis of maintained Python and JavaScript source | `push` to `main` / `release/**`, schedule, or explicit manual dispatch |
| Dependabot | Weekly discovery of GitHub Actions updates | Proposal-only pull requests; no automatic approval or merge |
| OpenSSF Scorecard | Exact-commit release-candidate analysis plus public default-branch posture publication | CLI scan on trusted non-default branch push/manual dispatch; published Action on `main`, branch-protection changes, schedule, or main-branch manual dispatch |

All external Actions remain pinned to full immutable commit SHAs with audited
semantic-version comments. Dependabot may propose a new Actions revision, but a
proposal is not trusted merely because a bot opened it: the source/tag must be
reviewed under [DEPENDENCY_UPDATES.md](DEPENDENCY_UPDATES.md), the resulting
workflow must retain immutable pins, and merge remains manual.

The published Scorecard path uses `ossf/scorecard-action` v2.4.4, resolved to its
reviewed commit. OpenSSF restricts Action publication on `push` and `schedule`
to the repository default branch. Release candidates therefore use the official
Scorecard CLI v5.5.0 with `--commit` bound to the exact candidate SHA. The Linux
amd64 release archive is verified against its recorded SHA-256 before execution;
the resulting JSON and CLI version are retained as workflow evidence.

## Permission model

The CodeQL and Scorecard workflows do not use `pull_request` or
`pull_request_target`. Their credentials are therefore unavailable to untrusted
pull-request code. CodeQL receives only `contents: read` plus
`security-events: write` for analysis publication. The Scorecard release-candidate
CLI job receives only `contents: read`. The default-branch publication job
receives `contents: read`, `security-events: write`, and `id-token: write`; the
last permission is used only for authenticated Scorecard result publication.
The repository-quality gate rejects `pull_request_target` and rejects
write-capable permissions in any workflow that runs on `pull_request`. Event
detection, workflow-level permission analysis, job-level permission analysis,
and immutable Action pin validation remain separate checker routines so each
finding is independently testable and the canonical gate stays within its
enforced complexity ceiling.

Checkout credentials are not persisted. No supply-chain workflow has a source,
release, label, issue, branch, or repository mutation step.

## Evidence boundary

A successful CodeQL run means the configured CodeQL queries completed against
the exact checked source. A successful release-candidate Scorecard CLI run means
the official scanner completed against the requested commit and its retained
JSON is candidate-bound; some Scorecard checks necessarily observe current
repository settings rather than historical settings. A successful default-branch
Scorecard Action run additionally means the configured SARIF/public result was
published. A Dependabot proposal means only that GitHub detected a candidate
dependency update.

None of those outcomes proves VBA compilation, Excel runtime behavior,
numerical accuracy, UI cleanup, packaging integrity, or release certification.
Those claims remain owned by their existing specialist and release evidence.
Likewise, a public Scorecard score is not the repository's project-quality
score.

## Review and maintenance

Pull-request review uses the existing read-only deterministic gates; analyzer
publication is deferred until the reviewed change reaches the default branch.
Release-branch Scorecard evidence is an exact-commit CLI scan, not a claim that
the public Scorecard service published that non-default candidate.

Before merging a CodeQL, Scorecard, or dependency-workflow change:

1. resolve every proposed Action release tag to its upstream commit and record
   the immutable SHA; verify downloaded CLI assets against recorded official
   release digests;
2. review permissions, triggers, checkout behavior, new network activity, and
   publication surfaces;
3. run repository integrity, authoritative workflow validation, checker
   development, and the dedicated supply-chain workflow fixtures;
4. preserve manual dependency approval and rollback under
   [DEPENDENCY_UPDATES.md](DEPENDENCY_UPDATES.md);
5. retain successful CodeQL and release-candidate Scorecard runs for the exact
   reviewed revision, and after default-branch integration retain the successful
   published Scorecard Action run separately.

If a required security analyzer is unavailable, misconfigured, or denied its
required permissions, its workflow is non-green. Do not convert an unavailable
result into a pass or weaken deterministic gates to recover a public score.
