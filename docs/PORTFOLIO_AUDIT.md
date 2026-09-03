# Portfolio audit — canonical repository baseline

> **Status:** P1.1 frozen and passed; post-audit implementation delta maintained in Section 0
>
> **Evidence cut:** 2026-09-03 11:19 UTC
>
> **Scope:** seven portfolio repositories; `GITHUB-TEMPLATE` is the target and is not scored
>
> **Decision rule:** evidence quality and reusability take precedence over recency

## 0. Post-audit implementation delta — 2026-09-03

The evidence cut, original scores and donor decisions in Sections 1–9 remain frozen. They describe the portfolio before canonical construction began and must not be silently rewritten as repositories change. This section records implementation completed after that cut. Current-quality scores will be recalculated only after P1 certification, when the baseline is stable.

### 0.1 Completed after the evidence cut

| Area | Implemented delta | Audit consequence |
|---|---|---|
| Canonical root policy | The template's canonical `.editorconfig`, `.gitattributes` and `.gitignore` were adapted across all seven repositories. | Dotfile differences now reflect repository needs rather than unmanaged drift. |
| Root governance documents | `CODE_OF_CONDUCT.md`, `SECURITY.md`, `CONTRIBUTING.md`, `CHANGELOG.md`, `INSTALLATION.md`, `RELEASING.md`, `VERSION` and MIT `LICENSE` were created or standardized across the eight repositories. Reciprocal navigation and repository-specific statements were reviewed. | The original documentation and release scores remain historical; this work is credited in the implementation plan and the later conformance rescore. |
| Canonical structure | `GITHUB-TEMPLATE` now contains `src/`, `tests/`, `examples/`, `assets/`, `docs/` and `tools/`, each with an instructional README where needed. `docs/REPOSITORY_STRUCTURE.md` defines ownership, canonical separation, legitimate legacy alternatives and non-duplication rules. | P1.3 passes. Existing repositories are not renamed merely to imitate the new template. |
| Pull-request intake | A canonical `.github/PULL_REQUEST_TEMPLATE.md` was selected after portfolio comparison and installed in all eight repositories. The implementations score 9.6–9.9/10, share exact-source, test/evidence, risk, rollback, security, provenance, documentation and release checks, and retain justified specialist blocks. | The PR portion of P1.5 is complete. Issue forms remain separate open work. |
| Issue labels | `GITHUB-TEMPLATE` now declares the frozen 20-label core in `.github/labels.json`, with `schema_version`, explicit `prune`, and separate profile/domain overlay arrays. A pinned, least-privilege workflow performs read-only pull-request validation and trusted reconciliation; the dependency-free reconciler self-tests, creates/updates/prunes deterministically and verifies an exact post-run match without a hard-coded label count. | The label portion of P1.5 passes. The audit scores and donor decision remain frozen; this implementation becomes the baseline for the later conformance assessment. |
| Repository quality | `tools/check_repo.py` combines the reusable KPR and Excel UI patterns as 19 generic, profile-driven rules configured by `.github/repository-profile.json`. It emits readable console, deterministic JSON and Markdown evidence; a positive fixture, 19 degraded fixtures and read-only hashes self-certify the tool. The pinned least-privilege workflow runs both the self-test and repository gate. | P1.6 passes. Domain-specific Excel, UI, numerical and release certification remains additive. The live tree continues to expose the independently open P1.2 root-README link blocker. |
| Repository settings | Merge, squash and rebase remain enabled; auto-merge remains disabled; merged branches are deleted; Update branch is enabled. The live protection matrix was rechecked without weakening documented exceptions. | Provisioning remains separate from file-template inheritance and still requires a post-creation contract. |
| Execution control | `docs/IMPLEMENTATION_PLAN.md` now records completed work, remaining P1–P3 tasks, dependencies and acceptance gates. | The plan is temporary and will be deleted after the programme definition of done is satisfied. |

No production VBA, workbook runtime behavior, numerical contract or release history was changed as part of this foundation standardization.

### 0.2 Current construction boundary

Completed:

- P1.1 portfolio audit;
- P1.3 canonical directory structure;
- canonical dotfiles and most root governance/release documents;
- portfolio-wide pull-request template standardization;
- canonical 20-label manifest, overlays and idempotent synchronization in `GITHUB-TEMPLATE`;
- canonical profile-driven repository checker, fixture certification and required workflow;
- repository-settings normalization and live protection review.

Still required before P1 certification:

- the template root `README.md`;
- `docs/POST_CREATION_CHECKLIST.md`;
- one placeholder syntax, optional/profile categories and template-facing documentation;
- bug, feature and documentation issue forms plus issue-template configuration;
- executable release gates and release fixtures;
- neutral VBA façade/core modules, premium headers and deterministic test harness;
- a generated-repository pilot and the template v0.1.0 certification.

### 0.3 Blocking donor defects after standardization

| Audit ID | Current state | Required disposition |
|---|---|---|
| `AUD-P1-01` — Logistic Regression source/test identity | **Still open.** Production and test paths still resolve to the same blob and declare the test module. Root governance and PR intake are improved, but functional certification would still be false. | Restore the production export and independently validate source and harness before using either as template evidence. |
| `AUD-P1-02` — Progress Bar inherited identity | **Still open.** Tracked production, test and demo content still exposes `cPerformanceManager`/`M_cPM_*` identity and Performance Manager artifacts. Root governance and PR intake are improved, but runtime provenance remains unresolved. | Reconcile the source manifest and reconstruct the actual Progress Bar export before any functional or release certification. |
| `AUD-P1-03` — Probability accuracy evidence | **Unchanged as frozen evidence.** The red run at the audit cut remains relevant to what was certifiable at that SHA. | Reassess only against a later exact source SHA; do not retroactively rewrite the audit verdict. |

The two code-identity defects are P2 repository-repair blockers, not reasons to import contaminated donor code into P1. The template must continue to use neutral VBA assets.

### 0.4 Score interpretation

The comparison matrix below is the **pre-standardization donor-selection baseline**, not a current portfolio league table. Root-document and PR-template improvements would change several present-day scores, but rescoring now would mix a partially built template with repositories at different migration stages. The next formal score replaces nothing: it will appear in `docs/PORTFOLIO_CONFORMANCE.md` after P1 certification and P2 migration.

The frozen donor decisions remain unchanged:

- DateTimePicker for documentation coverage;
- KPR for generic workflows, static checks and labels-as-code;
- Performance Manager for intake and release provenance;
- Excel UI for metadata/protection and generic VBA modularity;
- Probability Distributions for optional numerical and Excel assurance.

---

## 1. Executive conclusion

The canonical baseline must be a composite. No repository is strong enough across all eight axes to be copied wholesale.

| Axis | Strongest current implementation | Canonical decision |
|---|---|---|
| Documentation | `VBA-DATETIMEPICKER` | Reuse its operational coverage model, with project identity and subsystem detail moved to profile overlays. |
| Workflows | `KPR` for the generic baseline; `VBA-PROBABILITY-DISTRIBUTIONS` for specialised assurance | Make static integrity and label validation mandatory; keep Excel execution and numerical accuracy as opt-in extensions. |
| Static checks | `KPR` | Extract its generic checks into a profile-driven checker; do not copy KPR date-layer rules. |
| Issue and pull-request templates | `VBA-PERFORMANCE_MANAGER` | Use concise structured forms and its evidence/risk/rollback PR shape; add profile-specific sections only when selected. |
| Labels | `KPR` | Use a versioned manifest, validation, self-tests and idempotent reconciliation; separate the core taxonomy from overlays. |
| Release process | `VBA-PERFORMANCE_MANAGER` | Adopt exact-source tagging, explicit certification, asset hashes and honest provenance boundaries. |
| Repository metadata and protection | `VBA-EXCEL_UI` | Require deletion and force-push protection, PR routing and the generic integrity check; add protected release tags. |
| VBA structure and test harnesses | `VBA-EXCEL_UI` overall; `KPR`, `VBA-DATETIMEPICKER` and `VBA-PROBABILITY-DISTRIBUTIONS` by profile | Freeze a public-facade/internal-core pattern, exported-component hygiene and deterministic harness contract. |

The portfolio-wide score leader is `KPR` at **9.2/10**, but it is not selected automatically. `VBA-DATETIMEPICKER` is the documentation donor, `VBA-PERFORMANCE_MANAGER` the release donor, `VBA-EXCEL_UI` the protection and generic modularity donor, and `VBA-PROBABILITY-DISTRIBUTIONS` the specialised numerical/Excel-assurance donor. This is direct evidence that the audit did not equate “newest” with “best.”

Four constraints govern every later template-construction task:

1. Reuse patterns, not donor identity: names, badges, versions, issue numbers, current measurements and domain contracts must not leak into generated repositories.
2. A generic core and specialised extensions remain separate. A clean `library` repository must not inherit WinAPI, forms, numerical benchmarks or application packaging that it does not use.
3. Checks must be deterministic, self-testing and honest about what they cannot prove. Hosted text checks do not certify Excel execution; a workbook hash does not prove how the workbook was built.
4. Repository settings, rulesets and live labels are provisioning concerns. They must be specified beside the file template because GitHub does not copy them from a template repository.

## 2. Method and scoring

Each repository was inspected at the exact default-branch commit listed in the snapshot register. The recursive Git tree, repository metadata, releases, workflow definitions, rulesets, root documentation, declarative labels, checkers, exported VBA component headers and test-harness entry points were reviewed.

Scores use this scale:

| Score | Meaning |
|---:|---|
| `0` | Absent or unusable. |
| `1–3` | Materially incomplete, stale or unsafe to reuse. |
| `4–5` | Partial implementation with substantial gaps. |
| `6–7` | Useful implementation requiring material hardening. |
| `8–9` | Strong implementation; reusable after decoupling and focused corrections. |
| `10` | Canonical-ready, generic, tested and evidenced without unresolved defects. |

Scores measure fitness as template evidence, not the intrinsic value or functional maturity of a project. A project-specific numerical gate can be excellent and still require separation from the generic baseline.

## 3. Frozen default-branch and settings register

### 3.1 Default-branch trees

Every tree link below is commit-pinned. The recursive manifest is the complete captured tree; the shape column summarises tracked blobs by top-level location.

| Repository | Default-branch commit | Tracked blobs | Top-level shape | Complete tree evidence |
|---|---|---:|---|---|
| `VBA-PERFORMANCE_MANAGER` | [`142db3a`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/commit/142db3a49d6509e9c4143a74860dab783b06f008) | 33 | root 10; `.github` 7; `demo` 3; `docs` 4; `images` 4; `src` 2; `test` 1; `tools` 2 | [recursive tree](https://api.github.com/repos/danielep71/VBA-PERFORMANCE_MANAGER/git/trees/142db3a49d6509e9c4143a74860dab783b06f008?recursive=1) |
| `VBA-EXCEL_UI` | [`985dfaa`](https://github.com/danielep71/VBA-EXCEL_UI/commit/985dfaabbe0cc3f3004263c4998de45c9cbca0f4) | 33 | root 10; `.github` 5; `demo` 2; `docs` 3; `images` 5; `src` 4; `test` 1; `tools` 3 | [recursive tree](https://api.github.com/repos/danielep71/VBA-EXCEL_UI/git/trees/985dfaabbe0cc3f3004263c4998de45c9cbca0f4?recursive=1) |
| `VBA-PROGRESS_BAR` | [`f61a33a`](https://github.com/danielep71/VBA-PROGRESS_BAR/commit/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be) | 16 | root 6; `demo` 4; `images` 3; `src` 2; `test` 1 | [recursive tree](https://api.github.com/repos/danielep71/VBA-PROGRESS_BAR/git/trees/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be?recursive=1) |
| `VBA-DATETIMEPICKER` | [`d285865`](https://github.com/danielep71/VBA-DATETIMEPICKER/commit/d28586577900e465f323d4bae6d673fd041dc02c) | 34 | root 10; `.github` 5; `assets` 7; `demo` 2; `dist` 1; `images` 2; `src` 6; `test` 1 | [recursive tree](https://api.github.com/repos/danielep71/VBA-DATETIMEPICKER/git/trees/d28586577900e465f323d4bae6d673fd041dc02c?recursive=1) |
| `EXCEL-VBA-LOGISTIC-REGRESSION` | [`d44fde5`](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/commit/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e) | 12 | root 6; `demo` 1; `images` 3; `src` 1; `test` 1 | [recursive tree](https://api.github.com/repos/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/git/trees/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e?recursive=1) |
| `VBA-PROBABILITY-DISTRIBUTIONS` | [`f888215`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/commit/f8882155f5a61051bf6eec50f8b8977437b4cde1) | 237 | root 8; `.github` 8; `assets` 3; `benchmark` 206; `ci` 1; `docs` 2; `examples` 2; `src` 6; `tests` 1 | [recursive tree](https://api.github.com/repos/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/git/trees/f8882155f5a61051bf6eec50f8b8977437b4cde1?recursive=1) |
| `KPR` | [`686018e`](https://github.com/danielep71/KPR/commit/686018e1ab60e336840401b1856933182f9540a9) | 30 | root 11; `.github` 8; `assets` 1; `docs` 3; `src` 5; `test` 1; `tools` 1 | [recursive tree](https://api.github.com/repos/danielep71/KPR/git/trees/686018e1ab60e336840401b1856933182f9540a9?recursive=1) |

### 3.2 Repository metadata

All seven use `main`, report VBA as the primary language, use MIT where a licence is detected, enable Issues, disable Discussions, allow merge/squash/rebase, disable auto-merge, delete merged branches and expose the Update branch option. Those common settings are already the suite convention.

| Repository | Visibility | Description/topics | Projects | Wiki | Protection at the evidence cut |
|---|---|---|---:|---:|---|
| `VBA-PERFORMANCE_MANAGER` | Public | Specific description; 8 relevant topics | On | On | Active default-branch deletion/non-fast-forward and PR rule; no required status check in the ruleset. [Rule](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/rules/20905954) |
| `VBA-EXCEL_UI` | Public | Specific description; 9 relevant topics | On | On | Strongest: deletion/non-fast-forward, PR rule and strict `Repository and module checks`; no bypass actor. [Rule](https://github.com/danielep71/VBA-EXCEL_UI/rules/21104519) |
| `VBA-PROGRESS_BAR` | Private | Specific description; 5 relevant topics | On | Off | Branch unprotected. Repository rulesets are unavailable for this private repository under the current GitHub plan. |
| `VBA-DATETIMEPICKER` | Public | Specific description; 16 relevant topics | On | On | Deletion/non-fast-forward and PR rule on `main`; a separate deletion/non-fast-forward rule protects `traffic-history`; no required quality check. [Main rule](https://github.com/danielep71/VBA-DATETIMEPICKER/rules/22186296) |
| `EXCEL-VBA-LOGISTIC-REGRESSION` | Private | Specific description; no topics | On | Off | Branch unprotected. Repository rulesets are unavailable for this private repository under the current GitHub plan. |
| `VBA-PROBABILITY-DISTRIBUTIONS` | Public | Specific description; 13 relevant topics; Wiki homepage | Off | On | Deliberate project exception: deletion/non-fast-forward only, preserving direct pushes to `main`; no required check. [Rule](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/rules/22186928) |
| `KPR` | Public | Specific description; 8 relevant topics | Off | Off | Deletion/non-fast-forward plus strict `Repository integrity`; administrator-role always-bypass. Separate `v*` tag creation/update/deletion protection. [Main rule](https://github.com/danielep71/KPR/rules/21826384) · [tag rule](https://github.com/danielep71/KPR/rules/21826445) |

The target repository itself was at [`64095e2`](https://github.com/danielep71/GITHUB-TEMPLATE/commit/64095e269688043257c0ef5e84a65f5082807b7d) with four files: `.editorconfig`, `.gitattributes`, `.gitignore` and `CODE_OF_CONDUCT.md`. It is marked as a template and has active deletion/non-fast-forward protection. Its deliberately minimal current state is why this audit specifies a composite future baseline rather than treating the target as an eighth donor.

## 4. Seven-repository comparison matrix

Abbreviations: **Docs** documentation; **WF** workflows; **Static** static checks; **Intake** issue/PR templates; **Repo** metadata and protection. Each score links to the principal evidence; secondary evidence is recorded in the axis assessments below.

| Repository | Docs | WF | Static | Intake | Labels | Release | Repo | VBA/test | Average |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| `KPR` | [8.6](https://github.com/danielep71/KPR/blob/686018e1ab60e336840401b1856933182f9540a9/README.md) | [9.3](https://github.com/danielep71/KPR/tree/686018e1ab60e336840401b1856933182f9540a9/.github/workflows) | [9.8](https://github.com/danielep71/KPR/blob/686018e1ab60e336840401b1856933182f9540a9/tools/check_repo.py) | [9.1](https://github.com/danielep71/KPR/tree/686018e1ab60e336840401b1856933182f9540a9/.github/ISSUE_TEMPLATE) | [9.8](https://github.com/danielep71/KPR/blob/686018e1ab60e336840401b1856933182f9540a9/.github/labels.json) | [8.3](https://github.com/danielep71/KPR/releases/tag/v0.0.1) | [9.1](https://github.com/danielep71/KPR/rules/21826384) | [9.5](https://github.com/danielep71/KPR/tree/686018e1ab60e336840401b1856933182f9540a9/src/modules) | **9.2** |
| `VBA-PERFORMANCE_MANAGER` | [9.2](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/142db3a49d6509e9c4143a74860dab783b06f008/RELEASING.md) | [7.8](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/tree/142db3a49d6509e9c4143a74860dab783b06f008/.github/workflows) | [8.0](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/142db3a49d6509e9c4143a74860dab783b06f008/tools/vba_lint.py) | [9.3](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/tree/142db3a49d6509e9c4143a74860dab783b06f008/.github/ISSUE_TEMPLATE) | [6.4](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/142db3a49d6509e9c4143a74860dab783b06f008/.github/workflows/sync-label-colors.yml) | [9.8](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/releases/tag/v1.4.0) | [8.4](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/rules/20905954) | [9.2](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/tree/142db3a49d6509e9c4143a74860dab783b06f008/src) | **8.5** |
| `VBA-EXCEL_UI` | [8.9](https://github.com/danielep71/VBA-EXCEL_UI/blob/985dfaabbe0cc3f3004263c4998de45c9cbca0f4/README.md) | [8.1](https://github.com/danielep71/VBA-EXCEL_UI/blob/985dfaabbe0cc3f3004263c4998de45c9cbca0f4/.github/workflows/static-checks.yml) | [8.8](https://github.com/danielep71/VBA-EXCEL_UI/blob/985dfaabbe0cc3f3004263c4998de45c9cbca0f4/tools/check_repo.py) | [8.2](https://github.com/danielep71/VBA-EXCEL_UI/tree/985dfaabbe0cc3f3004263c4998de45c9cbca0f4/.github/ISSUE_TEMPLATE) | [3.0](https://github.com/danielep71/VBA-EXCEL_UI/tree/985dfaabbe0cc3f3004263c4998de45c9cbca0f4/.github) | [7.4](https://github.com/danielep71/VBA-EXCEL_UI/releases/tag/v1.1.2) | [9.4](https://github.com/danielep71/VBA-EXCEL_UI/rules/21104519) | [9.7](https://github.com/danielep71/VBA-EXCEL_UI/tree/985dfaabbe0cc3f3004263c4998de45c9cbca0f4/src) | **7.9** |
| `VBA-PROBABILITY-DISTRIBUTIONS` | [8.4](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/f8882155f5a61051bf6eec50f8b8977437b4cde1/README.md) | [8.8](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/tree/f8882155f5a61051bf6eec50f8b8977437b4cde1/.github/workflows) | [8.9](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/tree/f8882155f5a61051bf6eec50f8b8977437b4cde1/benchmark) | [8.0](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/tree/f8882155f5a61051bf6eec50f8b8977437b4cde1/.github/ISSUE_TEMPLATE) | [9.0](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/f8882155f5a61051bf6eec50f8b8977437b4cde1/.github/labels.json) | [3.0](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/releases) | [7.6](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/rules/22186928) | [9.6](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/tree/f8882155f5a61051bf6eec50f8b8977437b4cde1/src) | **7.9** |
| `VBA-DATETIMEPICKER` | [9.6](https://github.com/danielep71/VBA-DATETIMEPICKER/blob/d28586577900e465f323d4bae6d673fd041dc02c/INSTALLATION.md) | [2.5](https://github.com/danielep71/VBA-DATETIMEPICKER/blob/d28586577900e465f323d4bae6d673fd041dc02c/.github/workflows/daily-traffic.yml) | [1.5](https://github.com/danielep71/VBA-DATETIMEPICKER/tree/d28586577900e465f323d4bae6d673fd041dc02c/.github/workflows) | [8.9](https://github.com/danielep71/VBA-DATETIMEPICKER/tree/d28586577900e465f323d4bae6d673fd041dc02c/.github/ISSUE_TEMPLATE) | [3.0](https://github.com/danielep71/VBA-DATETIMEPICKER/tree/d28586577900e465f323d4bae6d673fd041dc02c/.github) | [8.6](https://github.com/danielep71/VBA-DATETIMEPICKER/releases/tag/v1.2.1) | [8.6](https://github.com/danielep71/VBA-DATETIMEPICKER/rules/22186296) | [9.5](https://github.com/danielep71/VBA-DATETIMEPICKER/tree/d28586577900e465f323d4bae6d673fd041dc02c/src) | **6.5** |
| `EXCEL-VBA-LOGISTIC-REGRESSION` | [4.8](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/blob/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e/README.md) | [0.0](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/tree/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e) | [0.0](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/tree/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e) | [0.0](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/tree/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e) | [0.0](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/tree/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e) | [0.0](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/releases) | [4.2](https://api.github.com/repos/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION) | [1.0](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/blob/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e/src/modules/M_LOGIT_REGRESSION.bas) | **1.3** |
| `VBA-PROGRESS_BAR` | [3.2](https://github.com/danielep71/VBA-PROGRESS_BAR/blob/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be/README.md) | [0.0](https://github.com/danielep71/VBA-PROGRESS_BAR/tree/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be) | [0.0](https://github.com/danielep71/VBA-PROGRESS_BAR/tree/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be) | [0.0](https://github.com/danielep71/VBA-PROGRESS_BAR/tree/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be) | [0.0](https://github.com/danielep71/VBA-PROGRESS_BAR/tree/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be) | [0.0](https://github.com/danielep71/VBA-PROGRESS_BAR/releases) | [4.8](https://api.github.com/repos/danielep71/VBA-PROGRESS_BAR) | [2.0](https://github.com/danielep71/VBA-PROGRESS_BAR/tree/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be/src) | **1.3** |

## 5. Axis assessments and template decisions

### 5.1 Documentation

| Required field | Audit finding |
|---|---|
| Strongest implementation | `VBA-DATETIMEPICKER`: the README, installation guide, contribution guide and security policy describe deployment variants, exported components, lifecycle ownership, validation states, upgrades, removal and recovery. Its documentation explains operational consequences rather than only listing APIs. |
| Useful secondary implementations | `VBA-PERFORMANCE_MANAGER` supplies the strongest release guide and assurance-boundary language. `VBA-EXCEL_UI` clearly documents facade/runtime/snapshot/title-bar ownership. `KPR` supplies contract-first and source-first engineering language. `VBA-PROBABILITY-DISTRIBUTIONS` supplies the best explanation of numerical evidence, direct-tail behaviour and benchmark provenance. |
| Reusable elements | A root set of `README.md`, `CHANGELOG.md`, `CONTRIBUTING.md`, `INSTALLATION.md`, `SECURITY.md`, `CODE_OF_CONDUCT.md`, `RELEASING.md`, `VERSION` and `LICENSE`; repository map; installation/import order; public versus internal surface; supported environments; validation boundary; failure/recovery guidance; security reporting; versioning and release evidence. |
| Project-specific elements to exclude | Donor names and prefixes; badges tied to current repositories; current versions, issue numbers and assertion counts; Wiki links; financial conventions; distribution catalogues; WinAPI details; provider leases; Ribbon callbacks; timer backends; screenshots and release filenames. |
| Defects before reuse | Performance Manager has no `CONTRIBUTING.md`. KPR's installation guide is intentionally a pre-release placeholder. Probability Distributions lacks root `CHANGELOG.md`, `INSTALLATION.md`, `RELEASING.md` and `VERSION`. The two private repositories contain only a README and conduct/licence layer. DateTimePicker's documentation explicitly admits its release evidence is procedural rather than automated. |
| Final decision and rationale | Use DateTimePicker's operational completeness as the model, Performance Manager's release language as the release chapter, and KPR's contract-first language for engineering guidance. Keep common documents short enough to remain maintainable; profile overlays own domain-specific sections. A generated repository must contain every root document with valid generic content, not “coming later” text. |

### 5.2 Workflows

| Required field | Audit finding |
|---|---|
| Strongest implementation | `KPR` has the strongest reusable baseline: a hosted static gate and a separate label workflow, least-privilege permissions, concurrency control, exact-SHA checkout/upload actions, checker self-tests, machine-readable output and `if: always()` artifact publication. The exact head passed [run 68](https://github.com/danielep71/KPR/actions/runs/33748387781). |
| Useful secondary implementations | Probability Distributions separates a pure-Python accuracy gate from a self-hosted Windows/Excel regression workflow. Excel UI runs both its checker and a house-style normal-form check. Performance Manager always publishes lint evidence. |
| Reusable elements | `push`, `pull_request` and manual triggers; path filters only where they cannot skip a required gate; read-only default permissions; explicit write permission only in reconciliation jobs; concurrency cancellation; exact-SHA checkout; self-tests; machine-readable artifacts; an explicit terminal enforcement step after evidence upload. |
| Project-specific elements to exclude | DateTimePicker's traffic export and `traffic-history` branch; benchmark dependency installation; KPR date-layer checks; workbook names; self-hosted runner labels; donor-specific artifact names; label maps embedded in workflow source. |
| Defects before reuse | Excel UI, Performance Manager and Probability Distributions use floating major action tags such as `@v4`, `@v5` or `@v7`; DateTimePicker pins checkout by SHA but lacks an audited version comment. Probability Distributions is not a green donor at this cut: its most recent observed Accuracy Gate [failed](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/33745235318), its Excel run was queued on an earlier SHA, and the current head has no matching successful run. DateTimePicker has no software-quality workflow. The private repositories have no workflows. |
| Final decision and rationale | Require `static-checks.yml` and `labels-sync.yml` in all profiles, based on KPR. Pin every external action to a full commit SHA with an audited semantic-version comment. Add `excel-vba-regression.yml` only to profiles configured with an eligible runner; add `accuracy-gate.yml` only to numerical libraries. Operational analytics workflows never count as quality gates. |

### 5.3 Static checks

| Required field | Audit finding |
|---|---|
| Strongest implementation | `KPR/tools/check_repo.py`: 24 named rules, standard-library-only execution, structured JSON, Markdown job summary, positive/degraded self-tests and checks for required files, stale identity, structured formats, links, text integrity, forbidden artifacts, line endings, label schema, action pinning, Git whitespace and exported VBA hygiene. |
| Useful secondary implementations | Excel UI adds a versioned public-API manifest, procedure/control-flow checks and a reformatter self-test. Probability Distributions contributes reference provenance, grid coverage, holdouts, degradation tests and source-claim consistency. Performance Manager contributes released-changelog freezing, version consistency and API declaration checks. |
| Reusable elements | Generic required-file manifest; unresolved-template-token and stale-identity scan; JSON/YAML/XML parse checks; internal Markdown links; secret/conflict-marker scan; forbidden artifact policy; LF/CRLF and encoding policy; label-manifest validation; immutable action references; `git diff --check`; `Option Explicit`; VBE export header/component-name validation; balanced procedures/directives; optional public-API manifest. |
| Project-specific elements to exclude | KPR's date system, locale parser, array engine, date-window constants and exact public members; Probability's tolerances, grids and reference engines; Performance Manager's historical version freeze and expected case count; Excel UI's module list and allowed labels. |
| Defects before reuse | KPR's checker hard-codes the KPR repository identity, 23-label count, module graph and implementation-plan references. Excel UI's checker is generic in technique but fixed to one module set. Performance Manager's lint covers only three known files. None is currently profile-configured. |
| Final decision and rationale | Fork KPR's checker architecture, not its rule constants. Drive required paths, component roles, profile and optional checks from a small versioned manifest. The template repository runs in `template` mode; generated repositories run in one of the three frozen profile modes. Both modes must self-test and produce the same JSON schema. |

### 5.4 Issue and pull-request templates

| Required field | Audit finding |
|---|---|
| Strongest implementation | `VBA-PERFORMANCE_MANAGER`: balanced bug and feature templates, the portfolio's only native documentation issue form, and a PR template covering scope, compatibility, production package, verification, regression, risk, rollback, documentation and reviewer focus. |
| Useful secondary implementations | KPR requires exact source identity, contract, reproduction, evidence and non-goals. DateTimePicker provides the deepest UI/application-state prompts. Excel UI is the best concise component-oriented alternative. Probability Distributions correctly asks for bit-exact arguments and independent numerical references. |
| Reusable elements | Required summary, impact, exact version/SHA, minimal reproduction, expected/actual behaviour, environment, evidence, recovery/workaround and sensitive-data warning for bugs; problem, user, acceptance criteria, non-goals, compatibility, alternatives and validation for features; scope, linked issue, SemVer/API impact, tests, environment, risks, rollback, docs, release impact and reviewer focus for PRs. |
| Project-specific elements to exclude | Financial-contract matrices, distribution regions, UI subsystem inventories, provider leases, specific functions and workbook names. KPR and DateTimePicker's long optional sections would burden a small library if copied unchanged. |
| Defects before reuse | Markdown issue templates cannot require fields and become very long when every specialised subsystem is embedded. Assignee handling is inconsistent. Template URLs and maintainers are donor-specific. The two private repositories have no intake templates. |
| Final decision and rationale | Provide `bug.yml`, `feature.yml` and `documentation.yml` as concise GitHub issue forms, plus `config.yml` and `PULL_REQUEST_TEMPLATE.md`. Use Performance Manager's field balance, then inject small profile-specific blocks. The generator substitutes the maintainer placeholder; no donor username is embedded in reusable source. |

### 5.5 Labels

| Required field | Audit finding |
|---|---|
| Strongest implementation | `KPR`: a 23-label versioned manifest plus a dedicated validator/reconciler with self-tests, PR validation, trusted reconciliation, least-privilege write permissions and exact-SHA action pins. |
| Useful secondary implementations | Probability Distributions has a declarative 22-label taxonomy, descriptions, priority labels, domain overlays and prune semantics. Performance Manager demonstrates useful domain overlays but synchronises colors only. |
| Reusable elements | A declarative schema; stable name/color/description triples; case-insensitive duplicate detection; length and color validation; deterministic sorting; validation on PRs; idempotent create/update reconciliation after trusted changes; optional pruning; a self-test that exercises rejected and changed states. |
| Project-specific elements to exclude | `winapi`, `wiki`, financial/distribution families, timing, measurement, statistics, state management and other domain-specific labels. |
| Defects before reuse | KPR and Probability use different root keys (`version` versus `schema_version`), casing (`ci` versus `CI`) and work labels (`tests` versus `testing`). KPR hard-codes exactly 23 labels. Performance Manager has no manifest, descriptions, creation or pruning. Repositories without a manifest cannot recreate their live taxonomy from source. |
| Final decision and rationale | Freeze a 20-label core: `behavior-change`, `blocked`, `bug`, `ci`, `documentation`, `duplicate`, `enhancement`, `good first issue`, `help wanted`, `invalid`, `P1`, `P2`, `P3`, `question`, `refactor`, `release`, `repository`, `security`, `tests`, `wontfix`. Profile/domain overlays are separate arrays. Use `schema_version` and explicit `prune`; never hard-code a total label count in executable logic. |

### 5.6 Release process

| Required field | Audit finding |
|---|---|
| Strongest implementation | `VBA-PERFORMANCE_MANAGER`: a complete release guide, scope freeze, version and changelog sync, pre/post-merge checks, exact release SHA, Excel certification, source-aligned packaging, tag verification, SHA-256 manifest, publication checks and recovery procedures. The [v1.4.0 release](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/releases/tag/v1.4.0) ships both workbook and `release-manifest.json`. |
| Useful secondary implementations | KPR supplies `VERSION`, a dated changelog, a clearly bounded setup pre-release and protected `v*` tags. DateTimePicker consistently publishes `.xlsm`/`.xlam` assets. Excel UI publishes a demo asset and maintains release-oriented changelog sections. |
| Reusable elements | `VERSION`; Keep-a-Changelog structure; SemVer; release checklist; exact-SHA/tag alignment; source as authority; asset hashes; environment and bitness record; manual-build disclosure; post-download verification; recovery rules; protected release tags. |
| Project-specific elements to exclude | Current versions, test counts, Excel builds, filenames, required donor modules, historic issue references and exact release notes. A binary asset is never part of a clean source template. |
| Defects before reuse | Performance Manager's provenance tool has a fixed donor file list and explicitly cannot prove workbook-from-source construction. KPR's only release is a setup pre-release with no distributable asset. DateTimePicker lacks a committed provenance manifest. Excel UI lacks `VERSION` and a release guide. Probability Distributions and the private repositories have no releases in the captured state. |
| Final decision and rationale | Make `VERSION`, `CHANGELOG.md` and `RELEASING.md` common. Initialise generated repositories at `0.0.0` with an `Unreleased` section. Parameterise the provenance tool by profile manifest. Library releases may remain source-only; UI/application binary releases require a manifest and explicit build-provenance boundary. Provision a no-delete/no-update/no-create-without-bypass rule for `refs/tags/v*`. |

### 5.7 Repository metadata and protection

| Required field | Audit finding |
|---|---|
| Strongest implementation | `VBA-EXCEL_UI`: descriptive public metadata, relevant topics, standard merge settings and the strongest active `main` rule—deletion/non-fast-forward protection, PR routing and a strict named status check without bypass actors. |
| Useful secondary implementations | KPR adds protected release tags and a required generic integrity status. DateTimePicker demonstrates protection for an automation-owned history branch. Performance Manager has PR routing. Probability Distributions documents a valid project-specific exception where direct pushes remain permitted. |
| Reusable elements | Specific description; maintained topic set; MIT detection; Issues enabled; Discussions off by default; merge/squash/rebase enabled; auto-merge disabled; merged-branch deletion and Update branch enabled; template flag only on `GITHUB-TEMPLATE`; deletion/non-fast-forward protection; PR rule; strict required status; protected `v*` tags. |
| Project-specific elements to exclude | Wiki homepage, traffic-history branch, project-specific check names, admin bypasses and Probability's direct-push exception. Those can be applied after generation as explicit deviations. |
| Defects before reuse | Rulesets and repository metadata are not inherited by GitHub template generation. The two private repositories cannot enable repository rulesets under the current plan. KPR permits administrator-role bypass and does not require PRs. DateTimePicker and Performance Manager require PR routing but no quality status. Probability protects history only, intentionally. |
| Final decision and rationale | The canonical provisioning specification combines Excel UI's `main` rule with KPR's tag rule. Require the canonical `Repository integrity` status, PR routing with zero mandatory approvals for a single-maintainer repository, stale-review dismissal, deletion and non-fast-forward protection, and no default bypass. Keep the Probability direct-push policy as a documented repository exception, not a baseline rule. |

### 5.8 VBA structure and test harnesses

| Required field | Audit finding |
|---|---|
| Strongest implementation | `VBA-EXCEL_UI`: one public facade plus private runtime/snapshot/title-bar modules, a versioned public-API manifest, exported-component checks, a house-style formatter and a release-certification harness that distinguishes complete/incomplete verdicts, verifies cleanup and emits machine-readable evidence. |
| Useful secondary implementations | KPR cleanly separates public facade, parsing, errors, dates and pure array engine. Probability Distributions separates core/special functions/families and couples a consolidated VBA harness to Python and Excel gates. DateTimePicker demonstrates modules/classes/forms/Ribbon layout and the strongest dirty-start/cleanup-state harness. Performance Manager demonstrates class/support-module separation and deterministic case/assertion accounting. |
| Reusable elements | One exported component per file; `Attribute VB_Name`/filename consistency; `Option Explicit`; `Option Private Module` for internal modules; public-facade/internal-core direction; explicit component dependency rules; a single test entry point; deterministic suite order; assertion and failure counts; named failures; dirty-start refusal where state matters; best-effort cleanup that preserves the original error; environment record; machine-readable evidence; compile as a mandatory external prerequisite. |
| Project-specific elements to exclude | All donor prefixes and APIs; exact case registries; numeric tolerances and grids; DatePicker forms/Ribbon; Excel UI WinAPI targeting; Performance Manager timer backends; KPR date rules; demo workbook binaries. |
| Defects before reuse | `EXCEL-VBA-LOGISTIC-REGRESSION` is blocked: [`src/modules/M_LOGIT_REGRESSION.bas`](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/blob/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e/src/modules/M_LOGIT_REGRESSION.bas) and [`test/M_LOGIT_REGRESSION_TESTS.bas`](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/blob/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e/test/M_LOGIT_REGRESSION_TESTS.bas) are the same blob (`0ea505e7…`) and both declare the test module. `VBA-PROGRESS_BAR` is also blocked as a donor: its tracked source/test names still describe Performance Manager, and two demo blobs are byte-identical to Performance Manager. Folder naming is inconsistent (`test` versus `tests`; flat versus `modules`). Several mature repositories rely on very large single exported modules. |
| Final decision and rationale | Standardise on `src/modules`, `src/classes`, `src/forms`, `src/ribbon` and `tests/modules`, with only profile-relevant component folders populated. Start from neutral facade/core/test placeholders, not donor production code. Use Excel UI's certification contract as the generic harness base, KPR's layering for libraries, DateTimePicker's lifecycle assertions for stateful UI and Probability's external-evidence extension for numerical libraries. |

## 6. Blocking defect and exclusion register

| ID | Severity | Repository/evidence | Finding | Reuse disposition |
|---|---|---|---|---|
| `AUD-P1-01` | P1 | [Logistic source](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/blob/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e/src/modules/M_LOGIT_REGRESSION.bas) · [test](https://github.com/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION/blob/d44fde5f7cfb5027bdbface4c1c38b9cdaf1659e/test/M_LOGIT_REGRESSION_TESTS.bas) | Production and test paths resolve to the same blob and test-module identity. | Exclude all code and harness content until the production export is restored and independently tested. README structure only may be consulted. |
| `AUD-P1-02` | P1 | [Progress Bar tree](https://github.com/danielep71/VBA-PROGRESS_BAR/tree/f61a33ac1959a92e5bf9a7cafdc35ed0edafe8be) | Source and test components retain `cPerformanceManager`/`M_cPM_*` identities; demo-builder and demo module blobs match Performance Manager. | Exclude code, demos and harness. Treat the repository as needing reconstruction, not standardisation. |
| `AUD-P1-03` | P1 | [Probability Accuracy Gate run 189](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/33745235318) | Latest observed accuracy gate failed; current snapshot head has no matching successful run. | Reuse architecture and check patterns only. Do not import the workflow as certified-green baseline evidence until the gate is green at the source SHA. |
| `AUD-P1-04` | P2 | [DateTimePicker workflows](https://github.com/danielep71/VBA-DATETIMEPICKER/tree/d28586577900e465f323d4bae6d673fd041dc02c/.github/workflows) | Only committed workflow is traffic collection, not a software-quality gate. | Reuse the UI source layout and harness concepts, never the operational workflow. |
| `AUD-P1-05` | P2 | [KPR workflow pins](https://github.com/danielep71/KPR/blob/686018e1ab60e336840401b1856933182f9540a9/.github/workflows/static-checks.yml) | Other donor workflows use floating action tags; KPR demonstrates the required immutable pattern. | Repin every canonical action to a full audited SHA. |
| `AUD-P1-06` | P2 | [KPR checker](https://github.com/danielep71/KPR/blob/686018e1ab60e336840401b1856933182f9540a9/tools/check_repo.py) | The best checker mixes generic repository rules with KPR identity, date architecture and exact label count. | Extract generic rules and move profile/project rules to data-driven configuration. |
| `AUD-P1-07` | P2 | [Private Logistic repository](https://api.github.com/repos/danielep71/EXCEL-VBA-LOGISTIC-REGRESSION) · [private Progress Bar repository](https://api.github.com/repos/danielep71/VBA-PROGRESS_BAR) | Current GitHub plan does not expose repository rulesets for these private repositories; both default branches are unprotected. | Document plan limitation. Do not weaken the public baseline to match it. |
| `AUD-P1-08` | P3 | [Portfolio tree register](#31-default-branch-trees) | `test`/`tests`, PR-template case, label schema keys and label casing are inconsistent. | Canonicalise paths and schemas in the template; migrate donors separately without blind renaming. |

## 7. Frozen repository profiles

Profiles are structural contracts for generated repositories. They are not quality tiers and do not classify a repository by size. All profiles inherit the common baseline; an extension may add requirements but may not remove common gates.

### 7.1 Common baseline — inherited by all profiles

| Area | Required contract |
|---|---|
| Root files | `.editorconfig`, `.gitattributes`, `.gitignore`, `README.md`, `CHANGELOG.md`, `CONTRIBUTING.md`, `INSTALLATION.md`, `SECURITY.md`, `CODE_OF_CONDUCT.md`, `RELEASING.md`, `VERSION`, `LICENSE`. |
| Standard tree | `src/`, `tests/`, `examples/`, `assets/`, `docs/`, `tools/`; empty supported directories contain an explanatory `README.md` or `.gitkeep` as appropriate. |
| GitHub intake | `bug.yml`, `feature.yml`, `documentation.yml`, `config.yml`, `PULL_REQUEST_TEMPLATE.md`, declarative `labels.json`. |
| Required automation | Generic static integrity; label-manifest validation/reconciliation. Both include self-tests, least-privilege permissions, immutable actions and machine-readable evidence. |
| VBA baseline | One neutral public facade, one internal core module and one deterministic test harness. All must compile after project-name substitution; internal modules are private. |
| Release baseline | `VERSION=0.0.0`, an `Unreleased` changelog, SemVer guidance, exact-source/tag procedure and profile-driven provenance tooling. No binary is tracked merely to make the template appear complete. |
| Placeholders | A single documented token vocabulary (`PROJECT_NAME`, `PROJECT_PREFIX`, `MAINTAINER`, `DESCRIPTION`, `PROFILE`). Template mode validates tokens; generated mode rejects unresolved tokens. Donor names are forbidden. |
| Provisioned settings | Standard merge policy; deletion/non-fast-forward protection; PR routing; required `Repository integrity`; protected `v*` tags. Settings are applied after generation and then verified. |

### 7.2 Profile manifest

| Contract | `library` | `ui-component` | `application` |
|---|---|---|---|
| Purpose | Reusable callable VBA with no owned end-user shell. | Embeddable interactive component that owns a bounded UI surface. | Complete workbook/add-in solution with startup, lifecycle and distributable package. |
| Required production layout | `src/modules/`; `src/classes/` only when the API requires stateful objects. | `src/modules/` and `src/classes/`; `src/forms/` and `src/ribbon/` when used. | `src/modules/`, `src/classes/`, plus applicable `src/forms/`, `src/ribbon/` and `src/workbook/` lifecycle exports. |
| Required tests | `tests/modules/PROJECT_TESTS.bas`; pure deterministic cases; public/error/boundary contract. | Library contract plus application-state snapshot, dirty-start detection, cleanup verification and UI entry-point coverage. | UI contract where applicable plus startup/shutdown, upgrade, recovery, packaging and end-to-end smoke tests. |
| Examples and demos | Source examples under `examples/`; no tracked macro-enabled binary. | Demo source under `examples/`; release demo/add-in may be an external release asset. | Reproducible sample/configuration under `examples/`; distributable workbook/add-in is a release asset with manifest. |
| Documentation overlay | API, parameterisation, numerical/functional limits. | Interaction model, host state, accessibility, WinAPI/Ribbon boundaries and recovery. | Architecture, lifecycle, configuration/data boundaries, operations, deployment, upgrade and rollback. |
| Workflow overlay | Optional pure-Python accuracy or oracle gate; optional Excel runner. | Excel/VBA runner when eligible; optional packaging validation. | Excel/VBA and package smoke gates; provenance and release-asset validation required. |
| Release default | Source-authoritative release; binary optional. | Source-authoritative release; binary convenience asset optional and hashed. | Versioned binary package expected; manifest, environment record and post-download smoke evidence required. |
| Explicit exclusions | Forms, Ribbon, WinAPI and application lifecycle unless the profile is changed. | Unrelated numerical grids, application deployment and operational analytics. | Donor-specific business data, credentials, generated local state and unreviewable workbook-only source. |
| Principal donors | KPR layering; Probability specialised assurance; Excel UI API manifest. | Excel UI modularity/certification; DateTimePicker component layout/lifecycle harness. | Composite only: DateTimePicker packaging, Performance Manager release provenance and the common baseline. No current repository qualifies as a complete application-profile donor. |

### 7.3 Profile invariants

1. A selected profile is recorded in one versioned configuration file and is validated by the generic checker.
2. The generated tree contains only folders and checks relevant to that profile; optional empty subsystems are not presented as implemented features.
3. Changing profile is a reviewed architecture change, not a silent directory addition.
4. `ui-component` and `application` state tests must distinguish `PASS`, functional failure, cleanup failure and dirty start.
5. Numerical assurance is an extension of `library`, not a requirement imposed on every library.
6. `application` is not shorthand for “large.” It means the repository owns deployment and lifecycle of a complete Excel solution.

## 8. Canonical donor map

| Canonical component | Primary donor | Secondary donor(s) | Reuse condition |
|---|---|---|---|
| Root documentation architecture | DateTimePicker | Performance Manager, Excel UI, KPR, Probability | Neutralise identity; split common/profile content; remove current evidence. |
| `RELEASING.md` and provenance | Performance Manager | KPR, DateTimePicker | Parameterise file list and profile; retain manual-build disclosure. |
| Generic `check_repo.py` architecture | KPR | Excel UI, Performance Manager | Extract generic rules; load paths/roles from profile config; keep self-tests and JSON output. |
| Public API manifest option | Excel UI | KPR required-member model | Make opt-in for pre-1.0, mandatory once public compatibility is declared. |
| Label manifest and reconciler | KPR | Probability | Adopt one schema and core/overlay split; remove exact-count and domain assumptions. |
| Issue/PR intake | Performance Manager | Excel UI, KPR, DateTimePicker, Probability | Convert core issue types to concise forms; inject profile blocks. |
| Generic hosted workflow | KPR | Excel UI | Keep immutable actions, artifact publication and terminal enforcement. |
| Excel/VBA execution extension | Probability | Excel UI/DateTimePicker harness contracts | Runner availability must be explicit; fork PRs must not receive secrets or unsafe self-hosted execution. |
| Library module layout | KPR | Probability, Excel UI | Replace donor code with compile-clean neutral facade/core placeholders. |
| UI component layout | DateTimePicker | Excel UI | Include only selected forms/classes/Ribbon; add lifecycle and cleanup contract. |
| Deterministic test harness | Excel UI | DateTimePicker, KPR, Probability, Performance Manager | Standardise verdicts, counts, cleanup and evidence output; no project-specific assertions. |
| Main/tag protection | Excel UI main rule | KPR tag rule | Provision after generation and verify separately from file checks. |

## 9. P1.1 acceptance gate

- [x] Seven repositories captured at exact default-branch commits.
- [x] Relevant metadata, merge settings and branch/ruleset state recorded.
- [x] Documentation, workflows, static checks, intake templates, labels, releases, repository controls and VBA/test structure compared.
- [x] Every axis records the strongest implementation, secondary implementations, reusable content, exclusions, defects and final decision.
- [x] Seven-repository scored matrix includes direct evidence links.
- [x] Blocking donor defects are explicit and prevent unsafe reuse.
- [x] `library`, `ui-component` and `application` profiles are frozen.
- [x] Selection is demonstrably evidence-led rather than recency-led.

**P1.1 verdict: PASS.** The audit is complete. This verdict certifies the decision record and evidence coverage; it does not certify the still-unbuilt canonical repository baseline, which belongs to the remaining P1 work.
