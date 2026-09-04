# 🗺️ EXCEL VBA Project Template — Implementation Plan

[![v1.0.0: certified](https://img.shields.io/badge/v1.0.0-certified-success)](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/releases/tag/v1.0.0)
[![v1.1.0: active](https://img.shields.io/badge/v1.1.0-active-1D76DB)](#v110-execution)
[![P2: 9/12 complete](https://img.shields.io/badge/P2-9%2F12%20complete-217346)](#p2-status)
[![Policy branches: 175/175](https://img.shields.io/badge/policy%20branches-175%2F175-success)](#current-certification-evidence)
[![Repository rules: 21/21](https://img.shields.io/badge/repository%20rules-21%2F21-success)](#current-certification-evidence)

**Owner:** Daniele Penza

**Repository:** `danielep71/EXCEL-VBA-PROJECT-TEMPLATE`

**Development branch:** `v1.1.0`

**Single merge path:** PR #41 — `v1.1.0` → `main`

**Certified baseline:** `v1.0.0`

**Plan revision:** 20 — v1.1.0 execution state synchronized

**Plan date:** 4 September 2026

**Status:** Active — 9 of 12 P2 findings complete; 3 remain

> **Temporary execution document.** This file is the live work plan for the
> `v1.1.0` hardening branch. Historical certification detail belongs in release
> evidence, issue closure records, and durable documentation. Delete this plan
> once the milestone is merged, released, and its durable decisions have been
> transferred to maintained documentation.

<a id="executive-status"></a>

## 1. 🎯 Executive Status

`v1.0.0` is a **certified and published baseline**. The template has substantive
VBA starter assets, all three generated profiles, structured issue intake,
deterministic initialization, release-integrity validation, live governance,
three certified pilots, protected release tags, and a published source-only
release.

`v1.1.0` is therefore **not a reconstruction milestone**. It is a correctness,
assurance, and maintainability hardening release. Its objective is to remove the
remaining false-green paths and development friction without weakening the
portable, dependency-light contract that made `v1.0.0` reusable.

The milestone uses one long-lived branch:

```text
main (certified v1.0.0)
          ↑
          │ one final merge only
          │
       v1.1.0
```

No completed v1.1.0 issue is merged individually to `main`. Each issue may be
closed once its implementation and branch-level evidence are complete, while
all milestone changes remain accumulated on `v1.1.0` until final certification.

<a id="release-boundaries"></a>

## 2. 🧱 Release Boundaries

<a id="p1-certification-gate"></a>

### ✅ `v1.0.0` — frozen certified baseline

The following P1 capabilities are complete and must not regress:

- substantive public façade, internal/core module, example, and regression harness;
- generated-profile contracts for `library`, `ui-component`, and `application`;
- classified placeholder schema and dry-run-first deterministic initializer;
- structured issue forms and private security routing;
- dependency-free release-integrity gate and protected `v*` tags;
- authoritative workflow validation with pinned and checksum-verified actionlint;
- post-creation provisioning contract and live read-back;
- three certified profile pilots;
- exact-SHA Excel compile/regression evidence;
- strict `Repository integrity` protection on `main`;
- published protected `v1.0.0` release.

The `v1.0.0` historical review score and pre-certification gaps remain valid as
historical evidence only. They are no longer statements about the current
repository.

### 🛠️ `v1.1.0` — current hardening release

The release must preserve the supported public VBA starter/API while hardening:

1. committed-state correctness;
2. VBA parser correctness;
3. API-manifest completeness;
4. workflow/local-action dependency validation;
5. SemVer and changelog semantics;
6. exhaustive blocking-policy fixture coverage;
7. checker development architecture;
8. documentation authority boundaries; and
9. safe live label-drift detection.

### 🌱 `v1.2.0` — future automation and portfolio maturity

The P3 tranche remains outside the current branch. It covers contract versioning,
semantic portfolio drift, repository provisioning, reusable workflows,
dependency-update policy, conformance reporting, advanced provenance,
external-link drift, and optional Windows/Excel evidence automation.

<a id="v110-execution"></a>

## 3. 🚦 v1.1.0 Execution Model

The milestone follows five rules:

1. **All development stays on `v1.1.0`.** `main` remains the certified v1.0.0 baseline.
2. **Issues close one by one.** Closure requires implemented code/documentation plus exact branch-level evidence.
3. **No intermediate milestone merges.** PR #41 is the only planned `v1.1.0` → `main` merge.
4. **Every new gate fails closed.** A new validator is not complete until its own negative cases prove it can block the intended defect.
5. **Final certification is cumulative.** The last branch head must pass the entire repository, initializer, workflow, release, focused-hardening, and coverage stack together.

The current implementation evidence baseline before this plan update is:

- branch head `1748b22285fb3e3e24cc8bc7cfa7d0720353ef82`;
- PR #41 exact-head run `33915808719`;
- complete `Repository integrity` success;
- canonical repository gate **21/21**;
- semantic policy coverage **175/175** production finding sites;
- all three initializer profiles green;
- all focused P2-01 through P2-07 gates green;
- evidence upload and terminal enforcement green.

<a id="p2-status"></a>

## 4. 🛠️ P2 Status — v1.1.0

| ID | Issue | Status | Current result / remaining objective |
| --- | ---: | --- | --- |
| **P2-01** | #9 | ✅ **Complete** | CI now validates committed candidate whitespace rather than a clean working tree against itself; local working-tree feedback remains separate. |
| **P2-02** | #13 | ✅ **Complete** | `GoTo`, `GoSub`, and `Resume` targets are resolved within the owning VBA procedure; cross-procedure false resolution is blocked. |
| **P2-03** | #14 | ✅ **Complete** | Nested conditional compilation is modelled with explicit environments and full branch state; Win64 correctly models both `Win32=True` and `Win64=True`. |
| **P2-04** | #12 | ✅ **Complete** | Public API extraction covers the supported declaration families, continuations, signatures, paired properties, collisions, and a mandatory all-profile manifest policy. |
| **P2-05** | #11 | ✅ **Complete** | Repository-local actions enforce containment, tracked metadata/entrypoints, one metadata file, runtime shape, quoted references, and non-empty metadata scalars. |
| **P2-06** | #10 | ✅ **Complete** | Strict SemVer precedence, prerelease identifier rules, release ordering, dates, VERSION agreement, and comparison-link semantics are enforced. |
| **P2-07** | #18 | ✅ **Complete** | Semantic branch coverage inventories production `finding(...)` sites and proves **175/175** current canonical sites with deterministic fixtures and machine-readable evidence. |
| **P2-08** | #17 | 🟡 **Next / active** | Modularize checker development while preserving the single-command dependency-light distributable and CLI contract. |
| **P2-09** | #16 | ⏳ **Pending** | Consolidate duplicated documentation contracts after the checker architecture stabilizes; assign one authority per evolving rule. |
| **P2-10** | #15 | ✅ **Complete** | Repository metadata, topics, social preview, feature settings, merge settings, and private vulnerability reporting are aligned. |
| **P2-11** | #19 | ✅ **Complete** | Label profile/domain overlays resolve from versioned repository policy with exact, idempotent reconciliation evidence. |
| **P2-12** | #20 | ⏳ **Pending** | Detect live label drift safely without allowing untrusted or scheduled writes. |

**P2 completion:** **9 / 12 findings complete.**

**Remaining issues:** **#17 → #16 → #20**.

<a id="completed-hardening"></a>

## 5. ✅ Completed v1.1.0 Hardening

### P2-01 — committed whitespace

The hosted gate now separates two Git states deliberately:

- **committed mode** validates the exact candidate range used for CI/release evidence;
- **working-tree mode** remains local feedback for staged and unstaged edits.

The self-test includes root-commit, committed-defect, explicit-base,
staged/unstaged, and clean-candidate cases.

### P2-02 — procedure-scoped VBA jumps

The hardening gate owns procedure boundaries and local label resolution for
`GoTo`, `GoSub`, and `Resume`. Same-name labels in different procedures are
legal; duplicate labels inside one procedure and cross-procedure targets are
blocking findings.

### P2-03 — conditional compilation

The analyzer evaluates supported VBA environments explicitly and maintains a
nested conditional stack. It fails closed on malformed/indeterminate policy
constructs and requires `PtrSafe` for every declaration reachable in supported
VBA7 environments.

### P2-04 — public API manifest

The reusable policy now requires explicit public visibility and one checked-in
`docs/PUBLIC_API.txt` manifest for all generated profiles. The gate records
normalized signatures and rejects missing/stale declarations, signature drift,
unsupported ambiguous multi-declarations, and public-name collisions.

### P2-05 — local GitHub Actions

The local-action gate complements actionlint rather than replacing it. It proves
repository containment, tracked metadata, metadata-file uniqueness, supported
runtime shapes, and tracked entrypoints while external actions continue to be
controlled by immutable SHA pins and audited version comments.

### P2-06 — release semantics

Release-ledger semantics are now independent from candidate-provenance checks.
The gate enforces SemVer 2.0.0 precedence, invalid numeric prerelease rejection,
strict descending changelog releases, real Gregorian dates, VERSION/newest
release agreement, and canonical comparison links.

### P2-07 — semantic blocking-branch coverage

This is the principal new assurance layer in the milestone so far.

`tools/check_policy_coverage.py` and its modular fixture implementation:

- statically inventory every production canonical finding site;
- dynamically record which fixture executes each site;
- fail when a production site becomes uncovered;
- fail when runtime emits a site missing from the static inventory;
- cover configuration, structured data, Markdown, secrets, artifacts, line endings,
  labels, issue forms, workflow actions, VBE headers, VBA structure/visibility,
  generated contracts, and legacy API paths;
- delegate the specialized P2-01 through P2-06 gates to their own self-tests;
- distinguish policy findings from operational error mapping;
- produce deterministic JSON and Markdown evidence.

At branch head `1748b22285fb3e3e24cc8bc7cfa7d0720353ef82`, the gate covered
**175/175** canonical production finding sites and the complete repository job
passed terminal enforcement.

<a id="remaining-work"></a>

## 6. 🔧 Remaining v1.1.0 Work

### 🟡 P2-08 / #17 — Modularize Checker Development

**Objective:** reduce review and maintenance cost without sacrificing the
portable checker artifact.

**Required implementation:**

- define explicit internal boundaries for configuration, structured-data parsing,
  VBA parsing, policy rules, reporting, and fixtures;
- choose reviewed source modules plus deterministic bundling, or an equivalently
  explicit internal architecture;
- retain `python3 tools/check_repo.py --root .` as the supported dependency-light
  generated-repository command;
- add independent parser and reporter tests;
- make the distributable artifact reproducible from reviewed source;
- add drift detection so generated/bundled output cannot silently diverge;
- document how a maintainer changes, tests, bundles, and releases the checker.

**Acceptance gate:**

- parser/reporter units can run independently;
- the existing CLI and deterministic report contract remain compatible unless a
  separately documented behavior change is intentional;
- bundled/distributed output is reproducible and drift-checked;
- P2-07 semantic coverage remains 100% against the *current* production finding
  inventory after refactoring;
- the complete positive/degraded, initializer, workflow, release, and focused
  suites remain green;
- generated projects require no package manager or new runtime dependency.

### ⏳ P2-09 / #16 — Consolidate Documentation Contracts

**Objective:** reduce duplicated operational policy while retaining the premium
navigation and all substantive safeguards already present.

**Dependency:** execute after P2-08 so documentation points to the final checker
architecture rather than an intermediate one.

**Required implementation:**

- assign one durable authority for initialization, repository structure, testing,
  release evidence, security, contribution, and checker-development contracts;
- convert duplicated procedural text elsewhere into concise links and context;
- preserve historical certification evidence without presenting it as current work;
- retain premium badges/glyphs, clear status cues, and stable explicit anchors;
- validate every local Markdown path and anchor after consolidation.

**Acceptance gate:** no contradictory duplicate contract; first-use navigation is
short; specialized detail remains discoverable; all documentation-link checks pass.

### ⏳ P2-12 / #20 — Detect Live Label Drift Without Untrusted Writes

**Objective:** make manual live-label drift visible even when no versioned label
file changes.

**Required implementation:**

- add a safe scheduled or manually invocable **read/plan-only** drift check;
- compare live labels with the versioned resolved policy;
- produce explicit machine-readable and human-readable drift evidence;
- prohibit automatic mutation on untrusted/scheduled drift-detection events;
- reserve reconciliation writes for reviewed trusted events or explicit maintainer action.

**Acceptance gate:** an out-of-band live label change becomes non-green/visible
without the drift detector modifying repository labels; deliberate reconciliation
restores exact match through the existing trusted path.

<a id="execution-sequence"></a>

## 7. 🧭 Remaining Execution Sequence

| Order | Issue | Work package | Dependency | Closure evidence |
| ---: | ---: | --- | --- | --- |
| **1** | **#17** | P2-08 checker modularization | P2-01…P2-07 stabilized | Reproducible distributable, unit boundaries, drift check, full green branch gate |
| **2** | **#16** | P2-09 documentation consolidation | Final P2-08 architecture | Premium docs, one authority per contract, zero stale links/anchors |
| **3** | **#20** | P2-12 safe live label-drift detection | Stable policy/docs | Read-only drift detection, deterministic evidence, trusted write separation |
| **4** | — | v1.1.0 release-candidate certification | All 12 P2 findings complete | Full exact-SHA certification stack green |
| **5** | **PR #41** | Single final merge to `main` | Certified release candidate | Required review/check state green; no unresolved blocking conversation |
| **6** | — | Protected `v1.1.0` tag and GitHub release | Merge commit certified | Annotated immutable tag, release evidence, source archive verification |

No separate feature branch is required for the remaining milestone work. The
`v1.1.0` branch is the development integration branch and PR #41 remains the
single final merge path.

<a id="current-certification-evidence"></a>

## 8. 🧪 Current Certification Evidence

The latest fully certified *development* head before this plan update is
`1748b22285fb3e3e24cc8bc7cfa7d0720353ef82`.

Exact-head PR run `33915808719` passed:

- authoritative workflow/actionlint fixtures;
- 22 deterministic release-integrity fixtures;
- strict release-semantics self-test and current ledger validation;
- all three repository-initializer profiles;
- committed-whitespace fixtures and candidate-range validation;
- procedure-scoped VBA jump fixtures and current-source validation;
- conditional-compilation fixtures and current-source validation;
- complete public-API fixtures and manifest validation;
- repository-local action fixtures and current workflow validation;
- policy-branch coverage determinism;
- **175/175** current canonical production finding sites;
- canonical repository checker self-test;
- canonical repository gate **21/21**;
- readable job-summary publication;
- evidence artifact upload; and
- final fail-closed terminal enforcement.

This evidence is **branch-level development evidence**, not the final v1.1.0
release certification. Any later implementation changes invalidate the head SHA
and require a fresh complete run.

<a id="final-v110-gate"></a>

## 9. 🏁 Final v1.1.0 Certification Gate

The milestone is ready for its **single merge** only when all statements below
are true at one exact branch head:

- [ ] #17 / P2-08 is closed with reproducible checker-development architecture.
- [ ] #16 / P2-09 is closed with no contradictory or stale documentation contract.
- [ ] #20 / P2-12 is closed with read-only live drift detection and trusted write separation.
- [ ] All 12 P2 findings are closed as completed.
- [ ] `v1.1.0` contains no unresolved temporary development branch assumptions.
- [ ] The canonical repository-quality gate passes every current rule with zero findings.
- [ ] Semantic branch coverage is **100% of the current production finding inventory**; the final count may legitimately differ from 175 after P2-08 refactoring.
- [ ] Every focused P2 validator self-test passes.
- [ ] Authoritative workflow validation passes at the exact candidate SHA.
- [ ] All three initializer profiles pass at the exact candidate SHA.
- [ ] Release-integrity and strict release-semantics suites pass.
- [ ] `VERSION`, dated changelog release heading, candidate tag intent, and comparison links are synchronized for `1.1.0`.
- [ ] Required template-release evidence is regenerated/bound to the exact release candidate according to `.github/release-policy.json`.
- [ ] Excel compile/regression evidence is refreshed if the active release policy requires exact-candidate Excel evidence.
- [ ] PR #41 has no unresolved blocking review conversation and all required checks are green.
- [ ] The final merge target remains `main`, with no intermediate v1.1.0 merge having occurred.

After the merge:

- [ ] verify the merged `main` SHA and its hosted gates;
- [ ] create protected **annotated** tag `v1.1.0` targeting the certified release SHA;
- [ ] verify tag object type and peeled commit;
- [ ] publish the GitHub `v1.1.0` release;
- [ ] verify generated source archives;
- [ ] delete the now-unnecessary `v1.1.0` development branch if no longer needed;
- [ ] close the v1.1.0 milestone; and
- [ ] remove this temporary implementation plan after durable documentation is synchronized.

<a id="v120-roadmap"></a>

## 10. 🌱 v1.2.0 / P3 Roadmap

The following work remains deliberately outside v1.1.0:

| ID | Issue | Improvement | Intended outcome |
| --- | ---: | --- | --- |
| P3-01 | #24 | Version template contract independently | Generated repositories can identify the baseline contract they adopted. |
| P3-02 | #25 | Semantic portfolio drift detection | Detect missing controls and unauthorized divergence without byte-for-byte cloning. |
| P3-03 | #28 | Dry-run-first repository provisioner | Apply and verify metadata, labels, merge settings, and rulesets reproducibly. |
| P3-04 | #23 | Versioned reusable workflows | Consumers can pin stable workflow releases while specialist Excel jobs remain local where appropriate. |
| P3-05 | #22 | Controlled dependency-update policy | SHA/tool updates carry provenance, review, and rollback evidence. |
| P3-06 | #29 | Portfolio quality/conformance reporting | Surface template version, controls, workflow health, protection, and release status. |
| P3-07 | #21 | Advanced release provenance | Add checksums/signatures/attestations where justified without forcing binaries on libraries. |
| P3-08 | #27 | External-link/documentation drift review | Keep local links blocking and network-dependent checks separate and resilient. |
| P3-09 | #26 | Optional Windows/Excel evidence pattern | Eligible repositories can automate exact-SHA Excel evidence without pretending every project has the same runner boundary. |

P3 should start only after the v1.1.0 contracts are stable enough to become the
next portfolio baseline.

<a id="governance"></a>

## 11. 🧭 Governance and Evidence Rules

Throughout the milestone:

- preserve stronger specialist controls in downstream repositories;
- never weaken a gate merely to make the template generic;
- distinguish current implementation evidence from historical certification;
- keep external actions immutable and version-audited;
- keep untrusted/scheduled workflows read-only unless a separate trusted write path exists;
- use exact commit SHAs and hosted run IDs for issue closure evidence;
- keep PR #41 as the single merge path;
- close issues only when the implemented acceptance criteria are demonstrably green.

<a id="next-action"></a>

## 12. ➡️ Immediate Next Action

**Proceed with #17 / P2-08 — Modularize checker development while preserving portability.**

The first implementation step is to define the reviewed source-module boundaries
and reproducible distributable contract for `tools/check_repo.py`, then add
independent parser/reporter tests and a drift check **before** moving production
rules out of the current portable artifact.

---

**Execution principle:** close one finding at a time with exact evidence; accumulate every v1.1.0 change on the milestone branch; certify once; merge once.
