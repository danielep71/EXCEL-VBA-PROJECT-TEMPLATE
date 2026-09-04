# 🗺️ EXCEL VBA Project Template — Implementation Plan

[![P1: certified](https://img.shields.io/badge/P1-certified-success)](#p1-certification-gate)
[![Release: v1.0.0](https://img.shields.io/badge/release-v1.0.0-217346)](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/releases/tag/v1.0.0)
[![v1.1.0: active](https://img.shields.io/badge/v1.1.0-active-1D76DB)](#v110-execution)
[![P2: 10/12](https://img.shields.io/badge/P2-10%2F12-6F42C1)](#p2-hardening)
[![Policy branches: 175/175](https://img.shields.io/badge/policy%20branches-175%2F175-success)](#p2-07)

**Owner:** Daniele Penza

**Repository:** `danielep71/EXCEL-VBA-PROJECT-TEMPLATE`

**Development branch:** `v1.1.0`

**Single merge path:** PR #41 — `v1.1.0` → `main`

**Certified baseline:** `v1.0.0`

**Plan revision:** 21 — P2-08 complete; v1.1.0 execution state synchronized

**Plan date:** 4 September 2026

> **Temporary execution document.** Delete this file after the global definition
> of done is satisfied. Preserve durable decisions in maintained documentation,
> release notes, the portfolio audit and the final conformance report.

<a id="v110-execution"></a>

## 1. 🎯 v1.1.0 Objective

`v1.0.0` is the certified, published baseline. `v1.1.0` is a post-certification
hardening release: strengthen repository correctness, parser behavior, release
semantics, checker assurance, documentation authority and live drift detection
without changing the template's fundamental architecture or weakening any
specialist control.

All development remains on branch `v1.1.0`. `main` stays at the certified
baseline until the entire milestone is complete. PR #41 is the only planned
merge back to `main`.

### Current milestone state

| Measure | State |
| --- | ---: |
| P2 findings | 12 |
| Complete | **10** |
| Open | **2** |
| Canonical repository rules | **21/21** |
| Canonical blocking finding sites covered | **175/175** |
| Final merge to `main` | Not yet |

### Remaining execution order

1. **#16 / P2-09** — consolidate duplicated documentation contracts.
2. **#20 / P2-12** — detect live label drift without untrusted writes.
3. Synchronize maintained documentation, changelog and release notes.
4. Run the complete `v1.1.0` certification gate at one exact branch head.
5. Merge PR #41 once into `main`.
6. Tag/publish `v1.1.0` only from the certified merged candidate.

## 2. 🏛️ Certified v1.0.0 Baseline

The following P1 capabilities are complete and are not reopened by this plan:

- deterministic template initialization for application, library and UI-component profiles;
- substantive public façade, internal/core and regression-test VBA contracts;
- structured GitHub issue intake and private security routing;
- dependency-free repository and release integrity gates;
- authoritative workflow validation with immutable dependency pins;
- post-creation provisioning contract and all-profile pilots;
- exact-SHA Excel compile/regression evidence;
- protected `main`, protected `v*` tags and published `v1.0.0`.

The exact v1.0.0 release candidate remains
`1144dd69112f6c238488c7888158d58b014fdc70`; the published protected annotated
tag peels to that candidate. Historical review observations remain evidence of
the state at their original snapshot and must not be rewritten as current state.

<a id="p1-certification-gate"></a>

### P1 certification gate — retained compatibility anchor

P1 is certified. Durable certification evidence is maintained in
[`PILOT_CERTIFICATION.md`](PILOT_CERTIFICATION.md), release documentation,
GitHub Actions evidence and the published v1.0.0 release.

<a id="p2-hardening"></a>

## 3. 🛠️ P2 Hardening

| ID | Issue | Status | Implemented / remaining contract |
| --- | --- | --- | --- |
| P2-01 | #9 — committed-whitespace validation | ✅ Complete | CI validates the committed candidate/range; local working-tree mode remains distinct. |
| P2-02 | #13 — procedure-scoped VBA jumps | ✅ Complete | `GoTo`/`Resume` targets resolve within their owning procedure; cross-procedure labels fail. |
| P2-03 | #14 — conditional compilation | ✅ Complete | Nested branch-stack semantics cover VBA6/VBA7 and Win32/Win64; Win32 remains true on Win64; reachable VBA7 declares require `PtrSafe`. |
| P2-04 | #12 — complete public API contract | ✅ Complete | Explicit supported public declarations, paired properties, signatures and manifest/collision policy are enforced. |
| P2-05 | #11 — local GitHub Actions | ✅ Complete | Local references are repository-contained, tracked, unambiguous and entry-point valid, including quoted references. |
| P2-06 | #10 — release semantics | ✅ Complete | Strict SemVer precedence, dates, release ordering, duplicates, VERSION agreement and comparison links are deterministic. |
| P2-07 | #18 — blocking policy-branch fixtures | ✅ Complete | AST inventory proves every canonical production finding site is exercised; 175/175 covered with deterministic machine-readable evidence. |
| P2-08 | #17 — checker development contract | ✅ Complete | `check_repo.py` remains the single-file standard-library runtime; seven internal ownership boundaries, independent parser/reporter/CLI tests, canonical check order and artifact SHA-256 are enforced by `checker_development.py` and a dedicated read-only workflow. |
| P2-09 | #16 — documentation authority/consolidation | 🔵 Active | Assign one authority per evolving contract, eliminate contradictory duplication and shorten the first-use path without deleting substantive controls. |
| P2-10 | #15 — metadata/public presentation | ✅ Complete | Specific description/topics, intentional features/merge settings, social preview and private vulnerability reporting are established. |
| P2-11 | #19 — operational label overlays | ✅ Complete | Versioned profile/domain selection drives deterministic label reconciliation and exact post-run evidence. |
| P2-12 | #20 — live label drift detection | ⏳ Next | Add read-only scheduled/explicit drift detection; writes remain restricted to trusted reconciliation events. |

**P2 completion:** **10 / 12 findings complete.**

### Evidence ledger

| Package | Closure evidence |
| --- | --- |
| P2-01–P2-06 | Dedicated focused gates and their deterministic self-tests pass on `v1.1.0`; each issue is closed with exact implementation/run evidence. |
| P2-07 | Closure head `1748b22285fb3e3e24cc8bc7cfa7d0720353ef82`; PR run `33915808719` passed 175/175 policy sites, all focused gates, initializer, 21/21 repository rules, artifact publication and terminal enforcement. |
| P2-08 | Head `678762f1986200b90ded002fd17de12547bb6eaf`; checker-development run `33917719222` and full repository run `33917719103` both passed. Runtime CLI and standard-library-only dependency contract remain unchanged. |
| P2-10 | Certified repository metadata/social-preview state retained from v1.0.0 work. |
| P2-11 | Certified versioned label-selection and exact reconciliation evidence retained from v1.0.0 work. |

## 4. 🔬 Current Checker Assurance Model

The repository-quality stack now has distinct responsibilities:

| Layer | Purpose |
| --- | --- |
| `tools/check_repo.py` | Canonical portable 21-rule repository gate. |
| `tools/checker_development.py` | Development boundaries, independent parser/reporter/CLI tests, canonical check order, runtime import policy and artifact identity. |
| `tools/check_policy_coverage.py` | Semantic coverage of every canonical production blocking finding site. |
| Focused P2 tools | Committed whitespace, VBA jumps, conditional compilation, public API, local actions and release semantics. |
| `tools/test_workflow_validation.py` | Authoritative hosted GitHub Actions/YAML validation and fixtures. |
| `tools/check_release.py` | Release candidate, evidence, tag and optional asset integrity. |

The checker-development model deliberately has **no bundle transform**:
`tools/check_repo.py` is both reviewed source and distributable runtime. Generated
repositories therefore need no package manager, installed dependency or hidden
assembly process. See [`CHECKER_DEVELOPMENT.md`](CHECKER_DEVELOPMENT.md).

<a id="p2-07"></a>

### P2-07 policy coverage

The semantic branch-coverage contract inventories production `finding(...)`
sites using Python AST, records the fixture that exercises each site and fails
if a production site becomes uncovered. Current certified branch evidence is
**175/175**. Numeric Python line coverage is deliberately not used as a proxy
for policy assurance.

## 5. 📚 P2-09 — Documentation Consolidation

Issue #16 now owns the next work package.

### Required outcome

- define one maintained authority for each evolving contract;
- convert duplicated normative text elsewhere into concise summaries plus links;
- preserve historical evidence as historical evidence rather than current instruction;
- keep first-use/navigation material short and task-oriented;
- retain exact security, release, initialization, contribution and certification protections;
- keep all internal Markdown links/anchors passing the repository gate.

### Acceptance gate

- no contradictory duplicate for any maintained operational contract;
- README/installation/contribution/release/security/checker-development navigation is coherent;
- specialized detail remains discoverable without copying the same policy into multiple files;
- all documentation links and anchors pass;
- initializer, policy coverage, repository checker and hosted terminal gate remain green.

## 6. 🔭 P2-12 — Live Label Drift Detection

After P2-09, issue #20 will add a read-only drift control around the already
trusted label reconciliation path.

Required properties:

- scheduled and/or explicit read-only comparison of live labels with versioned policy;
- no writes from untrusted pull-request or scheduled drift-check contexts;
- deterministic report of missing, extra or changed labels;
- explicit trusted reconciliation remains the only mutation path;
- terminal CI evidence distinguishes drift detection from synchronization.

## 7. 🌱 v1.2.0 / P3 Roadmap

P3 remains assigned to `v1.2.0` and is deliberately outside the v1.1.0 merge.

| ID | Improvement |
| --- | --- |
| P3-01 | Version the template contract independently from generated project versions. |
| P3-02 | Add semantic portfolio drift detection. |
| P3-03 | Add a dry-run-first repository provisioner. |
| P3-04 | Publish versioned reusable workflows. |
| P3-05 | Add controlled dependency-update policy. |
| P3-06 | Add portfolio quality/conformance reporting. |
| P3-07 | Add advanced release provenance where justified. |
| P3-08 | Add external-link and documentation-drift review. |
| P3-09 | Add an optional Windows/Excel reusable evidence pattern. |

No P3 work is required to merge or release v1.1.0.

## 8. ✅ v1.1.0 Exit Gate

The branch may be merged only when all statements below are true:

- [x] P2-01 through P2-08 are closed with exact evidence.
- [ ] P2-09 documentation authority/consolidation is closed.
- [x] P2-10 and P2-11 certified capabilities remain intact.
- [ ] P2-12 read-only label drift detection is closed.
- [ ] All open v1.1.0 issues are zero.
- [ ] `docs/`, root documentation, changelog and version/release notes describe the final state consistently.
- [ ] Checker-development contract passes independently.
- [ ] Policy-coverage gate remains complete with no uncovered production finding site.
- [ ] Committed-whitespace, focused P2 gates, initializer and canonical repository gate are green.
- [ ] Authoritative workflow validation is green.
- [ ] PR #41's final exact head is fully green with no unresolved blocking review finding.
- [ ] One final merge of PR #41 lands the complete milestone on `main`.
- [ ] The merged candidate is revalidated before protected `v1.1.0` publication.

## 9. ➡️ Immediate Next Action

Implement **#16 / P2-09**: consolidate duplicated documentation contracts while
preserving every substantive control and every historical evidence boundary.

---

**Execution principle:** close findings only with exact evidence, keep `main`
untouched until the milestone is complete, preserve stronger specialist controls,
and keep one unambiguous next action.
