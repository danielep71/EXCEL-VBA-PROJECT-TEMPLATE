# 🗺️ v1.2.0 Implementation Plan

[![Milestone: v1.2.0](https://img.shields.io/badge/milestone-v1.2.0-1D76DB)](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/milestone/3)
[![Stabilization: main](https://img.shields.io/badge/stabilization-main-217346)](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/tree/main)
[![Baseline: v1.1.0](https://img.shields.io/badge/baseline-v1.1.0-6f42c1)](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/releases/tag/v1.1.0)
[![Status: release certification](https://img.shields.io/badge/status-release%20certification-success)](#immediate-next-action)

**Milestone:** v1.2.0 · **Stabilization branch:** `main` · **Published baseline:** v1.1.0 at
`502a3836bec0eff888194f61ad5a7ba4701bf102`

**Plan revision:** 18 — implementation scope complete; release-time certification separated from issue acceptance

> **Temporary execution document.** Delete this file before the final v1.2.0 source
> capture. Durable contracts live in the documents named in
> [the authority map](README.md#authority-map); issue scope and implementation
> evidence live in the issues themselves.

---

## 1. 🎯 What this document owns

This plan now owns only the milestone-level transition from completed
implementation to release certification:

- the inventory of v1.2.0 implementation issues;
- the release-time checks that cannot truthfully be completed before the final
  release candidate exists;
- the ordering needed to avoid circular acceptance criteria;
- the temporary decisions that must disappear with this file before release.

It does not redefine the contracts implemented by the milestone issues.

---

## 2. 📋 Milestone implementation scope

The thirteen planned P3 issues are implemented. #48 was the additional baseline
integration prerequisite.

| Issue | Item | Theme | Implementation |
| --- | --- | --- | --- |
| [#24](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/24) | P3-01 — Version the template contract independently | Contract | Complete |
| [#25](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/25) | P3-02 — Build semantic portfolio drift detection | Portfolio | Complete |
| [#28](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/28) | P3-03 — Add a dry-run-first repository provisioner | Portfolio | Complete |
| [#23](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/23) | P3-04 — Publish versioned reusable workflows | Distribution | Complete |
| [#22](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/22) | P3-05 — Add controlled dependency-update policy | Supply chain | Complete |
| [#29](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/29) | P3-06 — Publish portfolio quality and conformance reporting | Portfolio | Complete |
| [#21](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/21) | P3-07 — Add profile-aware advanced release provenance | Release | Complete |
| [#27](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/27) | P3-08 — Check external links and documentation drift separately | Assurance | Complete |
| [#26](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/26) | P3-09 — Add an optional Windows and Excel evidence pattern | Assurance | Complete |
| [#44](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/44) | P3-10 — Re-audit Markdown documentation and source comments | Documentation | Complete |
| [#45](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/45) | P3-11 — Apply the VBA house style to every starter module | Source | Complete |
| [#46](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/46) | P3-12 — Publish the repository creation and complete file-reference wiki | Documentation | Complete; closure reconciliation |
| [#47](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/47) | P3-13 — Consolidate repeated focused-gate CLI orchestration | Tooling | Complete |

[#48](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/issues/48)
restored the post-v1.1.0 fixes into the milestone baseline and is complete.

---

## 3. 🔄 Closure boundary for #46

#46 delivered the versioned Wiki source, fourteen-page journey, generated
sidebar, exact file/directory reference, all three profile lifecycles,
publication manifest, deterministic source checks, exact-source export and a
public development-edition read-back. #44 completed the documentation audit,
and current `main` continues to pass the Wiki source workflow.

The earlier #46 wording also required a final v1.2.0 tree, a final-source Wiki
publication, a complete live creation-to-release pilot and release-time network
observations **before the issue could close**. That is circular: the final
release candidate is not available until implementation issues are closed and
stabilization is complete. Those checks are therefore retained below as
release-certification obligations rather than being marked as if they had
already occurred.

This distinction does not weaken the evidence contract. A development Wiki
publication proves only its recorded source edition; the final release must
still refresh and read back the Wiki from the final source SHA.

<a id="immediate-next-action"></a>

## 4. ➡️ Immediate next action

**Finish release certification on `main`; do not add new milestone functionality
unless a release check discovers a defect.**

The pre-reconciliation `main` head was
`9c9df14ffe4e59cc977a7be9419db3be60e71737`. It is twelve commits ahead of the
#44 audited Wiki source `f01f3ab253d3b636e6d80c2345d9f5c07f61567b`, with no
tracked-file additions or removals in that comparison. The changed files are
limited to `CHANGELOG.md`, `README.md`, `RELEASING.md`,
`docs/CHECKER_DEVELOPMENT.md`, `tools/checker_development.py`,
`tools/initialize_repository.py` and `tools/test_release_provenance.py`.

Wiki source checks passed on that exact pre-reconciliation head in hosted run
`34405755491`. The public development edition remains the exact-source read-back
at Wiki commit `73201060b838d378df9ce6a78c8c369c647e464a` for source
`f01f3ab253d3b636e6d80c2345d9f5c07f61567b`; it is not relabeled as a publication
of a later `main` commit.

---

## 5. 🏁 Milestone release exit gate

v1.2.0 is ready to tag and publish only when all of the following hold:

- [ ] Every implementation issue in section 2 is closed with truthful evidence.
- [ ] `Repository integrity` and `Checker development` pass on the exact final candidate.
- [ ] The initializer self-test passes for `application`, `library` and `ui-component` on that candidate.
- [ ] `VERSION`, the dated changelog heading and the candidate tag agree under the release-semantics gate.
- [ ] Excel compile and regression evidence is captured for the exact final candidate SHA.
- [ ] The external release-evidence bundle validates with zero findings.
- [ ] A live external-link observation is completed under #27 policy; deterministic defects are zero and restricted/transient outcomes are reported separately rather than converted to PASS.
- [ ] A clean-room maintainer journey records live GitHub creation/provisioning, Excel validation and a first release, including numbered steps, gaps/corrections and elapsed-step evidence.
- [ ] Browser review confirms the published Wiki Home, sidebar and complete page navigation render as intended.
- [ ] This temporary plan is deleted, the Wiki catalogue/reference is regenerated for the resulting final tracked tree, and deterministic Wiki/source checks pass.
- [ ] The complete Wiki is exported from that exact final source SHA, published, freshly cloned/read back, byte-compared with zero drift, and its final Wiki commit is recorded.
- [ ] The v1.2.0 tag and GitHub release are published only after the preceding evidence is complete.

A network restriction or transient failure remains an explicit non-success
observation; it is not rewritten as a deterministic documentation defect. If a
release-time check finds a real implementation defect, fix it through a small
reviewed PR and rerun the affected evidence.

---

## 6. ⚖️ Resolved milestone decisions

| ID | Decision | Resolution |
| --- | --- | --- |
| D-01 | Ship a 1.1.1 patch for #43 fixes? | **No.** They ship with v1.2.0. |
| D-02 | Split the thirteen P3 issues across releases? | **No.** One v1.2.0 release. |
| D-03 | Continue stabilization only on `release/1.2.0` after PR #49? | **No.** PR #49 moved stabilization to protected `main`; the earlier release-branch-only sequence is superseded. |
| D-04 | Make final-release-only evidence a prerequisite for closing #46? | **No.** Preserve it in the milestone release exit gate so the evidence is still mandatory without creating a circular issue/release dependency. |

---

## 7. 🛡️ Governance and evidence rules

- Stabilization proceeds through small reviewed PRs into protected `main`.
- Close implementation issues only with evidence appropriate to their delivered
  scope; never claim final-release or live-host evidence that has not occurred.
- Bind release certification to the exact final source SHA, not to an earlier
  development candidate.
- Preserve stronger specialist controls and immutable dependency pins.
- Keep scheduled/untrusted observation workflows read-only unless a separate
  trusted write path is explicitly reviewed.
- Treat restricted and transient network outcomes separately from deterministic
  defects.
- Delete this plan before the final Wiki inventory/export so the published
  reference describes the actual release tree exactly once.

---

**Execution principle:** implementation is complete; final certification must now
prove the release candidate without reclassifying missing evidence as success.
