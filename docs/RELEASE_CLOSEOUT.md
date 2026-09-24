# Post-release closeout

This procedure closes the gap between pre-tag certification and the provider state that exists only after a release is published. It is **read-only with respect to repository and release state**: the workflow captures provider facts, validates them, and retains evidence; it does not move tags, alter a GitHub Release, close issues, edit milestones, publish the Wiki, or upload release assets.

## Canonical procedure

After the annotated tag, tag-triggered static checks, GitHub Release, release milestone, and (for the canonical template) Wiki publication are complete, run the **Release closeout** workflow manually from the released repository. Supply:

- `tag`: the published tag, for example `v1.2.1`;
- `candidate_sha`: the exact 40-character SHA that was certified before tagging;
- `milestone_number`: the GitHub milestone number used for that release;
- `wiki_browser_reviewed`: `true` only after a human has opened the published Wiki and checked Home, sidebar/navigation, and representative links in a browser;
- `expect_prerelease`: normally `false`; set it only for an intentionally prerelease publication;
- `allow_not_latest`: normally `false`; set it only when the release is intentionally not expected to be GitHub's current latest release.

The workflow checks out the exact candidate, captures GitHub REST facts, partitions and independently verifies canonical certification evidence when applicable, performs the existing template-mode Wiki read-back comparison, tests GitHub-generated source archive retrieval, builds one retained snapshot, and evaluates that snapshot with `tools/_release_closeout.py`. The terminal workflow verdict is green only when applicable certification verification, closeout validation, readable summary, and retained evidence all succeed.

## Deterministic provider controls

The closeout validator binds all deterministic checks to the same candidate SHA. It verifies:

1. the remote tag ref is an **annotated tag object**, the tag object has the expected name, and it resolves to the certified commit rather than a moved or lightweight tag;
2. the canonical static-check workflow has a completed successful `push` run whose `head_branch` is the release tag and whose `head_sha` is the candidate;
3. the GitHub Release exists for that tag, is published rather than draft, has the expected prerelease flag, and matches the expected latest-release state;
4. the raw GitHub Release asset list is partitioned before closeout validation into **certification-evidence attachments** and filtered **product/runtime assets**; the latter are compared with the selected candidate profile's `allowed_asset_globs` from `.github/release-policy.json`;
5. a source-only profile has no product/runtime assets; binary-capable profiles may contain only product names allowed by their candidate-bound patterns. Canonical template releases separately require the durable certification attachments defined by [RELEASE_EVIDENCE.md](RELEASE_EVIDENCE.md);
6. `VERSION`, the released CHANGELOG heading, the released comparison link, and the `Unreleased` comparison link remain coherent with the tag;
7. the provider comparison range resolves to the release range and contains the certified candidate when the range is ahead;
8. milestone closure is evaluated from the actual captured milestone membership and item states, with the provider's open/closed counters reconciled to those items rather than trusted as a UI percentage.

An unexpected uploaded asset, lightweight or moved tag, failed/missing tag CI, draft or incorrectly classified Release, wrong VERSION/tag relationship, unresolved comparison, open milestone, or stale milestone counters is non-green.

## Release asset partitioning and source archives

Three distinct provider surfaces must not be conflated:

- **raw uploaded Release assets**: every entry returned by the GitHub Release API;
- **certification-evidence attachments**: for a canonical template release, the four durable files required by [RELEASE_EVIDENCE.md](RELEASE_EVIDENCE.md) — the certification ZIP, detached ZIP signature, manifest and SHA-256 file;
- **product/runtime assets**: the remaining uploaded assets after certification evidence is partitioned out, governed by the selected profile's candidate-bound `allowed_asset_globs`;
- **provider-generated source archives**: GitHub's `zipball_url` and `tarball_url`, which are not entries in the Release API `assets` array.

The closeout workflow plans this partition from the raw Release response before invoking the deterministic closeout helper. For the canonical template's source-only profile, the four certification attachments are required while the filtered product/runtime asset list must be empty. Generated projects record certification as not applicable unless they deliberately adopt an equivalent local certification policy; their uploaded assets remain product assets and are checked normally.

Canonical certification verification is independent of the product-asset check. The workflow downloads the planned certification attachments, verifies their hashes, detached signature, candidate identity and committed/current signer trust through `tools/release_certification.py`, and writes `certification-verification.json`. A missing, tampered, incorrectly signed or candidate-mismatched certification set is non-green even when the filtered product-asset check passes.

GitHub-generated ZIP/tar archive availability and retrieval remain separate network observations and never substitute for either durable certification evidence or product assets.

## Wiki and UI observations

Wiki publication remains owned by `tools/check_wiki.py`; the closeout helper does not implement a second Wiki policy. For **template-mode** releases, the workflow clones the published Wiki, invokes the existing checker against the exact candidate source, and includes that checker's result in the closeout snapshot. The human `wiki_browser_reviewed` confirmation is likewise a template-only release requirement.

For **generated-project** releases, Wiki verification is not inherited automatically. The workflow records the Wiki check as `not-applicable`; a generated project that deliberately adopts a Wiki must define and evidence its own publication contract rather than treating the canonical template requirement as passed.

Some evidence is necessarily an **observation**, not a repository fact. The report labels these separately from deterministic provider controls:

- successful retrieval of GitHub-generated ZIP/tar source archives;
- for template mode only, the explicit browser review of the Wiki Home/sidebar/navigation.

A missing or failed applicable observation remains non-green, but its category is preserved so the report does not imply that a browser review or network retrieval was derived from Git history.

## Evidence retained

The workflow retains its diagnostic reports as Actions artifacts for 90 days, including:

- `certification-plan.json`: the raw-asset partition into certification evidence and product/runtime assets, or explicit generated-mode not-applicable certification scope;
- `certification-verification.json`: the independent certification verification result, or explicit not-applicable scope;
- `snapshot.json`: the normalized input facts used by the closeout validator;
- the authoritative Wiki comparison JSON/Markdown when applicable, or the explicit generated-mode Wiki record;
- `release-closeout.json`: machine-readable closeout result;
- `release-closeout.md`: concise human-readable closeout result.

`release-closeout.json` records the SHA-256 of `snapshot.json`, the candidate/tag/profile identity, deterministic and observation status, tag/CI/Release/product-asset/comparison/milestone/Wiki state, certification status and categorized findings. Retaining the snapshot makes the conclusion replayable without relying on a maintainer workstation.

The 90-day Actions retention is **not** the durability contract for canonical certification. The four certification attachments must remain attached unchanged to the GitHub Release for the lifetime of that release as required by [RELEASE_EVIDENCE.md](RELEASE_EVIDENCE.md). An expiring workflow artifact does not replace that durable record.

## Offline fixture contract

The helper has a deterministic self-test:

```bash
python3 tools/_release_closeout.py --self-test
```

The fixture matrix covers the valid path plus VERSION/tag mismatch, lightweight and moved tags, failed tag CI, draft/unexpected-prerelease/not-latest Release state, unexpected source-only assets, missing source-archive exposure, unresolved comparison, open milestone membership/state, stale milestone counters, Wiki drift, and source-archive retrieval failure. These fixtures do not call GitHub and do not mutate live provider state.

## Scope boundary

Post-release closeout supplements, but does not replace, pre-tag `check_release.py` certification, provenance validation, Excel runtime evidence, or the Wiki publication checker. A green closeout report means the captured post-publication state is coherent with the certified candidate under this contract; it is not a new claim about Excel execution, numerical accuracy, UI runtime behavior, or release-binary provenance.
