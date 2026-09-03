# GitHub Template and Portfolio Standardization — Implementation Plan

**Owner:** Daniele Penza  
**Portfolio:** Seven existing Excel/VBA repositories plus `GITHUB-TEMPLATE`  
**Plan revision:** 3  
**Plan status:** Active — P1 in progress  
**Baseline date:** 3 September 2026  

> **Temporary execution document.** Keep this file under `docs/` while the programme is active. Delete it after the global definition of done is satisfied; preserve durable decisions in the audit, conformance report, repository documentation and changelog.

## 1. Objective

Construct `GITHUB-TEMPLATE` as the canonical starting point for Daniele Penza's Excel/VBA repositories, using the strongest proven elements from the existing portfolio while preserving legitimate differences between:

- reusable VBA libraries;
- UI components and add-ins;
- analytical models and full applications.

After certifying the template, assess each existing repository against it and migrate improvements selectively. Standardization must improve quality and governance without replacing stronger project-specific controls or forcing artificial structural uniformity.

## 2. Repositories in scope

| Repository | Current role | Migration posture |
| --- | --- | --- |
| `VBA-PERFORMANCE_MANAGER` | Mature performance library | Preserve release provenance and measurement controls |
| `VBA-EXCEL_UI` | Mature UI library | Preserve public API and UI regression controls |
| `VBA-PROGRESS_BAR` | Private UI component | Bring governance and CI up to baseline |
| `VBA-DATETIMEPICKER` | UI component/add-in | Preserve distribution and traffic-specific structure |
| `EXCEL-VBA-LOGISTIC-REGRESSION` | Private analytical library | Bring governance and CI up to baseline |
| `VBA-PROBABILITY-DISTRIBUTIONS` | Numerical library | Preserve specialized gates; defer required status checks until PR routing is adopted |
| `KPR` | Financial analytics library | Use as the principal general-purpose scaffold donor |
| `GITHUB-TEMPLATE` | Canonical template | Build, certify and version as the portfolio baseline |

## 3. Current baseline

Completed work:

- canonical `.editorconfig`, `.gitattributes` and `.gitignore` created in `GITHUB-TEMPLATE`;
- project-specific adaptations propagated across the seven existing repositories;
- evidence-based eight-axis audit completed and frozen in `docs/PORTFOLIO_AUDIT.md`;
- canonical `CODE_OF_CONDUCT.md`, `SECURITY.md`, `CONTRIBUTING.md`, `CHANGELOG.md`, `INSTALLATION.md`, `RELEASING.md`, `VERSION` and MIT `LICENSE` added to the template and standardized across the portfolio;
- reciprocal root-document navigation reviewed and corrected;
- canonical `src/`, `tests/`, `examples/`, `assets/`, `docs/` and `tools/` structure created in the template with instructional READMEs;
- canonical-versus-legacy directory policy documented in `docs/REPOSITORY_STRUCTURE.md`;
- `.github/PULL_REQUEST_TEMPLATE.md` benchmarked, standardized and installed across all eight repositories while retaining repository-specific evidence sections;
- declarative issue-label manifest, profile/domain overlay model and idempotent synchronization workflow implemented in `GITHUB-TEMPLATE`;
- canonical profile-driven repository-quality checker, deterministic reports, 19-rule fixture suite and required workflow implemented in `GITHUB-TEMPLATE`;
- merge commits, squash merges and rebase merges enabled consistently;
- auto-merge disabled consistently;
- automatic deletion of merged branches and the **Update branch** capability enabled;
- `GITHUB-TEMPLATE` marked as a template repository;
- basic deletion and non-fast-forward protection applied to its default branch;
- current live protection matrix verified: Performance Manager, Excel UI and DateTimePicker require PR routing; Excel UI additionally requires `Repository and module checks`; Probability retains its documented pre-v1.0 direct-push exception with deletion/non-fast-forward protection only.

Known incomplete items:

- the template root `README.md` and post-creation settings checklist remain to be created;
- placeholder syntax/category standardization and issue forms remain to be implemented;
- the generic release gate, reusable VBA headers, neutral modules and regression harness remain to be built;
- a clean repository generated from the template has not yet been certified;
- no existing repository has yet been benchmarked against a certified template release.
- `VBA-PROGRESS_BAR` remains excluded as a code/test donor because tracked production, test and demo identities still belong to Performance Manager.
- `EXCEL-VBA-LOGISTIC-REGRESSION` remains excluded as a code/test donor because its production and test paths still resolve to the same test-module blob.

## 4. Design principles

1. **Evidence before adoption.** A component enters the template only after comparison across all relevant repositories.
2. **Generic core, specialized extensions.** The template supplies universally valid controls; domain-specific gates remain additive.
3. **No regression.** Migration may not weaken existing tests, release provenance, numerical accuracy, documentation or branch protection.
4. **No blind renaming.** Existing `test`/`tests`, `demo`/`examples` and `images`/`assets` layouts are changed only when the benefit exceeds the migration cost.
5. **Compile-clean VBA.** Reusable VBA examples and the test harness must import into the VBE without unresolved symbols.
6. **Explicit placeholders.** Every value requiring customization must be mechanically discoverable and documented.
7. **Deterministic checks.** Repository-quality and release gates must produce stable results and actionable failure messages.
8. **Settings are separate from files.** GitHub template generation does not inherit labels, rulesets or repository settings; these require a post-creation contract.

## 5. Frozen donor map

| Component | Primary donor | Canonical decision |
| --- | --- | --- |
| Documentation | `VBA-DATETIMEPICKER` | Reuse its operational coverage model, neutralized and split into common/profile content |
| Workflows | `KPR`; Probability for specialized assurance | Require generic static integrity and label validation; keep Excel/numerical assurance optional |
| Static checks | `KPR` | Extract a profile-driven checker and remove KPR-specific date and exact-count rules |
| Issue and PR intake | `VBA-PERFORMANCE_MANAGER` | Reuse concise structured intake plus evidence, risk and rollback expectations |
| Labels | `KPR` | Use one versioned manifest with a 20-label core and profile/domain overlays |
| Release process | `VBA-PERFORMANCE_MANAGER` | Adopt exact-source tagging, certification boundaries, hashes and recovery rules |
| Metadata and protection | `VBA-EXCEL_UI` | Combine Excel UI's `main` policy with KPR's protected-tag policy |
| VBA structure and harnesses | `VBA-EXCEL_UI` overall; profile-specific secondary donors | Use public-facade/internal-core direction and a deterministic, evidence-producing harness |

The evidence, scores, exclusions and rationale are frozen in `docs/PORTFOLIO_AUDIT.md`. Later work may refine an implementation, but it may not silently replace the audit decision record.

## Delivery status

| Work package | Status | Remaining boundary |
| --- | --- | --- |
| P1.1 Portfolio audit | **Complete** | Frozen audit is the authoritative donor decision record |
| P1.2 Identity and documentation | **In progress** | Root `README.md` and post-creation checklist |
| P1.3 Canonical structure | **Complete** | No retroactive renaming of existing repositories |
| P1.4 Placeholder governance | **In progress** | Scanner/catalogue exist; one syntax, optional/profile categories and documentation remain |
| P1.5 Collaboration files | **In progress** | Issue forms and issue-template configuration; label sub-gate passes |
| P1.6 Repository-quality gate | **Complete** | Live baseline remains red until the P1.2 root README blocker is resolved |
| P1.7 Release controls | **In progress** | Documentation exists; executable release gate and fixtures do not |
| P1.8 Reusable VBA assets | **Not started** | Neutral headers, façade/core samples and harness |
| P1.9 Governance policy | **In progress** | Live state verified; post-creation policy not yet written |
| P1.10 Pilot certification | **Not started** | Depends on all prior P1 packages |
| P2 Selective migration | **Early pre-alignment only** | Certified-template comparison remains pending |
| P3 Drift automation | **Not started** | Depends on stable P1 and completed P2 |

---

# P1 — Build and certify the canonical baseline

**Priority:** Critical  
**Outcome:** A clean repository generated from `GITHUB-TEMPLATE` is structurally complete, self-explanatory and green without inheriting project-specific material.

## P1.1 Complete the portfolio audit

- [x] Capture the default-branch tree and relevant repository settings for all seven repositories.
- [x] Compare the eight required axes:
  - documentation;
  - workflows;
  - static checks;
  - issue and pull-request templates;
  - labels;
  - release process;
  - repository metadata and protection;
  - VBA structure and test harnesses.
- [x] For every axis, record:
  - strongest implementation;
  - useful secondary implementations;
  - reusable elements;
  - project-specific elements to exclude;
  - defects that must be corrected before reuse;
  - final template decision and rationale.
- [x] Produce a seven-repository comparison matrix with scores and evidence links.
- [x] Freeze three supported repository profiles: `library`, `ui-component`, and `application`.

**Deliverable:** `docs/PORTFOLIO_AUDIT.md`

**Acceptance gate:** All seven repositories and all eight axes are covered; no component is selected solely because it is the newest.

**P1.1 verdict: PASS.** The audit is frozen at its evidence cut. Later repository changes do not rewrite the historical snapshot; material changes are evaluated during P2 conformance.

## P1.2 Complete template identity and documentation

- [x] Add an MIT `LICENSE`.
- [x] Add `VERSION` at the neutral generated-repository baseline `0.0.0`; advance the template release to `0.1.0` only at certification.
- [ ] Create a template-oriented `README.md` explaining profiles, generation and first-use steps.
- [x] Create `CHANGELOG.md` using Keep a Changelog conventions.
- [x] Create `CONTRIBUTING.md` with branch, commit, testing and pull-request expectations.
- [x] Create `SECURITY.md` with supported-version and private-reporting guidance.
- [x] Create `INSTALLATION.md` for module import, add-in/application variants and trust settings.
- [x] Create `CODE_OF_CONDUCT.md`.
- [x] Create `RELEASING.md` with version, changelog, tag, evidence and release steps.
- [ ] Add `docs/POST_CREATION_CHECKLIST.md` for settings that GitHub templates cannot inherit.
- [x] Add `docs/REPOSITORY_STRUCTURE.md` explaining the canonical structure, supported profiles and legitimate alternatives.

**Acceptance gate:** All links resolve, no donor repository names remain, and the generated README clearly tells a maintainer what must be replaced, retained or deleted.

## P1.3 Create the canonical repository structure

- [x] Add `src/`, `tests/`, `examples/`, `assets/`, `docs/` and `tools/`.
- [x] Place a short instructional `README.md` in directories that would otherwise be empty.
- [x] Define when profile-specific alternatives such as `demo/`, `dist/`, `images/` or `test/` are legitimate.
- [x] Avoid empty decorative directories and duplicate locations serving the same purpose.
- [x] Document the canonical separation between public modules, internal/core modules, UI forms/classes and test modules.

**Acceptance gate:** A repository generated from the template has no unexplained empty directory and no ambiguity over where source, tests, examples, documentation and tooling belong.

**P1.3 verdict: PASS.** This structure is canonical only for newly generated repositories; existing repositories are not renamed merely to imitate it.

## P1.4 Add explicit placeholder governance

- [ ] Define a single placeholder syntax, for example `{{PROJECT_NAME}}` and `{{REPOSITORY_NAME}}`.
- [x] Maintain the allowed placeholder catalogue in one machine-readable manifest.
- [ ] Distinguish required, optional and profile-specific substitutions.
- [ ] Mark optional blocks using consistent template comments.
- [ ] Prohibit placeholders inside VBA identifiers when they would make modules uncompilable.
- [x] Add a scanner that fails when unresolved required placeholders remain after initialization.
- [x] Add a donor-identity scan covering repository names, URLs, badges, module prefixes and release versions.

Suggested required placeholders:

- `{{PROJECT_NAME}}`
- `{{REPOSITORY_NAME}}`
- `{{PROJECT_DESCRIPTION}}`
- `{{MODULE_PREFIX}}`
- `{{REPOSITORY_PROFILE}}`
- `{{CURRENT_VERSION}}`
- `{{SUPPORT_CONTACT}}`

**Acceptance gate:** The scanner finds every intentional placeholder and rejects every unregistered one; the template itself documents why each placeholder exists.

## P1.5 Standardize GitHub collaboration files

- [ ] Add bug, feature and documentation issue forms.
- [ ] Add issue-template configuration with valid support and security links.
- [x] Add a pull-request template covering scope, tests, documentation, compatibility, evidence, risk, rollback and release impact.
- [x] Standardize the PR template across all seven existing repositories, retaining justified repository-specific sections.
- [x] Adopt a declarative `labels.json` catalogue and synchronization workflow based on KPR.
- [x] Freeze the 20-label core taxonomy and overlay model in the portfolio audit.
- [x] Ensure template labels contain descriptions and accessible colors.
- [ ] Document the default rule that portfolio issues are assigned to `danielep71` only after repository creation, not hard-coded into reusable issue forms.

**Acceptance gate:** Label synchronization is idempotent; issue forms create valid issues; no form points to a donor repository.

**P1.5 label sub-gate: PASS.** The 20-label core is declared in `.github/labels.json`; profile/domain overlays are explicit; validation and reconciliation are self-tested; pull requests are read-only; trusted runs reconcile with least privilege and verify an exact post-run match. The overall package remains open until the issue forms and configuration are complete.

## P1.6 Build the generic repository-quality gate

- [x] Merge the strongest generic checks from KPR and `VBA-EXCEL_UI` into `tools/check_repo.py`.
- [x] Drive required files and profile differences through a configuration file rather than hard-coded donor assumptions.
- [x] Check at minimum:
  - required files and directories;
  - unresolved placeholders;
  - donor identity leakage;
  - dotfile policy;
  - Markdown link integrity for local links;
  - `VERSION` format;
  - changelog structure;
  - VBA file encoding, exported-name consistency and required headers;
  - absence of prohibited generated, temporary or credential files;
  - workflow syntax and action pinning policy.
- [x] Add `.github/workflows/repository-quality.yml` for pushes and pull requests.
- [x] Keep domain-specific checks outside the generic gate.
- [x] Produce both readable console output and a machine-readable report.

**Acceptance gate:** Positive fixtures pass, negative fixtures fail for the intended reason, and two consecutive runs on the same tree are byte-for-byte equivalent apart from timestamps explicitly excluded from comparison.

**P1.6 verdict: PASS.** The dependency-free gate executes 19 named rules. One positive fixture and 19 rule-specific degraded fixtures pass their expected outcomes; repeated JSON and Markdown reports are byte-for-byte identical; and before/after tree hashes prove read-only execution. The current construction tree separately reports three broken links to the not-yet-created root `README.md`, preserving that P1.2 blocker instead of weakening link validation.

## P1.7 Add release controls

- [x] Add canonical `VERSION`, `CHANGELOG.md` and `RELEASING.md` foundations.
- [ ] Add a release gate validating semantic versioning.
- [ ] Require equality among `VERSION`, the changelog release heading and the release tag.
- [ ] Reject unresolved placeholders and uncommitted release evidence.
- [ ] Verify required documentation and licences before a release.
- [ ] Define the expected evidence bundle by repository profile.
- [ ] Separate generic release integrity from optional project certification evidence.
- [ ] Add release instructions for protected `v*` tags.
- [ ] Document that rulesets and required checks must be configured after template generation.

**Acceptance gate:** A valid synthetic release passes; mismatched version, missing changelog, mutable tag or incomplete evidence fixtures fail deterministically.

## P1.8 Create reusable VBA assets

- [ ] Define a premium standard module header covering purpose, public surface, dependencies, state, error policy and worksheet safety.
- [ ] Provide compile-clean sample modules for public façade and internal/core patterns.
- [ ] Create a neutral regression harness with:
  - test registration;
  - equality and tolerance assertions;
  - expected-error assertions;
  - deterministic counters and summaries;
  - guaranteed cleanup;
  - explicit environment reporting.
- [ ] Provide guidance for class modules and UserForms without embedding project-specific controls.
- [ ] Add a small importable example demonstrating one passing suite.
- [ ] Check module names, exported attributes and line endings through the quality gate.

**Acceptance gate:** The sample suite imports and runs without editing; the template contains no hidden dependency on Excel UI, KPR, Performance Manager or another donor.

## P1.9 Verify and document live governance policies

- [x] Verify that `VBA-PERFORMANCE_MANAGER`, `VBA-EXCEL_UI` and `VBA-DATETIMEPICKER` retain their required pull-request rule.
- [x] Verify that `VBA-EXCEL_UI` additionally requires its live `Repository and module checks` context.
- [x] Confirm that `VBA-PROBABILITY-DISTRIBUTIONS` retains its deliberate pre-v1.0 direct-push exception with deletion/non-fast-forward protection only.
- [x] Confirm that the template retains basic delete/force-push protection without prematurely requiring checks that do not yet exist.
- [x] Confirm that private Progress Bar and Logistic Regression rulesets are unavailable under the current GitHub plan; document rather than weaken the public baseline.
- [ ] Record the protection policy by maturity level in the post-creation checklist.

**Acceptance gate:** Live rulesets match the documented policy and every referenced status-check context exists on the default branch.

## P1.10 Pilot and certify the template

- [ ] Generate a temporary repository using the GitHub template button.
- [ ] Complete its required placeholders using the `library` profile.
- [ ] Run the quality and release gates from a clean clone.
- [ ] Import and execute the VBA example and regression harness in Excel.
- [ ] Verify issue forms, pull-request template and label synchronization.
- [ ] Apply the post-creation settings checklist and verify the resulting ruleset.
- [ ] Delete the temporary pilot only after preserving its certification evidence.
- [ ] Tag and publish `GITHUB-TEMPLATE` v0.1.0.

**P1 exit criteria:**

- the formal audit is complete;
- the template contains every required artifact;
- all generic checks are green;
- a generated pilot repository is certified;
- the live protection matrix matches the documented baseline and approved exceptions;
- the template is safe to use as the benchmark for existing repositories.

---

# P2 — Benchmark and selectively migrate the seven repositories

**Priority:** Important  
**Dependency:** P1 certification  
**Outcome:** Every repository meets the applicable baseline while retaining stronger domain-specific architecture and controls.

Early, low-risk pre-alignment is already complete for the canonical dotfiles, root governance documents, `VERSION`, MIT licence where missing, and pull-request template. This does not replace the post-certification benchmark or authorize blind structural migration.

## P2.1 Establish the migration method

- [ ] Score every repository against the certified baseline by audit axis.
- [ ] Classify each difference as:
  - `REQUIRED` — baseline defect or missing universal control;
  - `ADOPT` — beneficial improvement with low regression risk;
  - `KEEP` — valid project-specific divergence;
  - `DEFER` — beneficial but not justified in the current release;
  - `NOT APPLICABLE` — irrelevant to the repository profile.
- [ ] Produce a repository-specific change list before editing.
- [ ] Preserve existing user changes, release branches and active milestone work.
- [ ] Use one auditable commit or pull request per coherent migration unit.

**Acceptance gate:** No difference is automatically treated as a defect merely because it diverges from the template.

## P2.2 Wave A — raise the two minimal private repositories

### `VBA-PROGRESS_BAR`

- [x] Add the canonical root governance documents, `VERSION`, licence and PR template.
- [ ] Add issue forms, labels-as-code and the generic repository-quality gate after P1 certification.
- [ ] Correct the manifest/source identity mismatch: tracked source, test and demo content still belongs to Performance Manager.
- [ ] Reconstruct and independently verify the actual Progress Bar production export before any certification claim.
- [ ] Review and normalize VBA module headers and the inherited test harness.
- [ ] Document the private-repository ruleset limitation under the current GitHub plan.
- [ ] Decide whether `demo/`, `images/` and `test/` remain justified aliases.

### `EXCEL-VBA-LOGISTIC-REGRESSION`

- [x] Add the canonical root governance documents, `VERSION`, licence and PR template.
- [ ] Add issue forms, labels-as-code and the generic repository-quality gate after P1 certification.
- [ ] Restore the production source export: the production and test paths currently resolve to the same test-module blob.
- [ ] Review the analytical test harness for deterministic fixtures and tolerances.
- [ ] Document the private-repository ruleset limitation under the current GitHub plan.
- [ ] Complete repository topics and discoverability metadata before any public release.

**Wave A exit criterion:** Both repositories pass the generic baseline with documented exceptions and without changing analytical or UI behavior.

## P2.3 Wave B — align the mature UI repositories

### `VBA-EXCEL_UI`

- [x] Standardize the canonical root documents and PR template without removing UI-specific evidence requirements.
- [ ] Reconcile its checker with the canonical quality gate without weakening the public API manifest.
- [ ] Preserve the required `Repository and module checks` status context.
- [ ] Add only missing label, release or documentation controls.
- [ ] Keep SDI/window-specific documentation and regression evidence project-specific.

### `VBA-DATETIMEPICKER`

- [x] Standardize the canonical root documents and PR template without removing UI/distribution-specific requirements.
- [ ] Add a generic static quality gate alongside the traffic workflow.
- [ ] Preserve `dist/` and add-in packaging when supported by the release process.
- [ ] Reconcile issue templates and labels with the canonical taxonomy.
- [ ] Preserve the separate protection of `traffic-history`.

### `VBA-PROGRESS_BAR`

- [ ] Revisit after its Wave A baseline and compare its UI/test patterns with the two mature UI repositories.
- [ ] Adopt only proven shared UI conventions.

**Wave B exit criterion:** All UI repositories share common governance and quality controls while retaining product-specific forms, distribution models and regression cases.

## P2.4 Wave C — align the analytical and library repositories

### `VBA-PERFORMANCE_MANAGER`

- [x] Add and standardize `CONTRIBUTING.md`, `VERSION` and the remaining canonical root documents.
- [x] Standardize the PR template while preserving certification, risk and rollback evidence.
- [x] Preserve `RELEASING.md`, release provenance and measurement certification during root-document standardization.
- [ ] Reconcile label synchronization and generic static checks with the template.
- [ ] Add required status checks only after confirming the stable context name.

### `VBA-PROBABILITY-DISTRIBUTIONS`

- [x] Add and standardize the generic root release artifacts and PR template without altering numerical contracts.
- [ ] Keep accuracy, holdout and Excel execution evidence outside the generic template core.
- [ ] Reassess required specialized gates when the repository adopts PR routing; until then preserve the documented pre-v1.0 direct-push exception.
- [ ] Preserve benchmark and numerical-evidence directories.

### `KPR`

- [x] Standardize the canonical root documents and PR template without disturbing date-layer contracts.
- [ ] Reconcile its general-purpose scaffold with the final template.
- [ ] Preserve date-layer contracts, implementation plans and public/core module separation.
- [ ] Treat KPR as a validation that the template supports a growing multi-module library.
- [ ] Avoid disturbing active v0.0.2 implementation work.

### `EXCEL-VBA-LOGISTIC-REGRESSION`

- [ ] Revisit after Wave A and add analytical validation controls justified by model risk.

**Wave C exit criterion:** Generic governance is aligned, specialized scientific and release evidence remains stronger than the baseline, and all required checks are green.

## P2.5 Close the portfolio benchmark

- [ ] Re-score all seven repositories after migration.
- [ ] Document every accepted divergence from the canonical template.
- [ ] Confirm that documentation, workflows and comments reflect the new state.
- [ ] Confirm that required status contexts exactly match live workflow job names.
- [ ] Run relevant static, regression, numerical and release checks.
- [ ] Publish a portfolio conformance report.

**Deliverable:** `docs/PORTFOLIO_CONFORMANCE.md`

**P2 exit criteria:**

- every repository meets its applicable baseline;
- every exception is intentional and documented;
- no project-specific gate has been weakened;
- all repositories have appropriate metadata and protection;
- migration work is committed, pushed and traceable.

---

# P3 — Automate maintenance and prevent future drift

**Priority:** Enhancement  
**Dependency:** Stable P1 baseline and completed P2 migrations  
**Outcome:** Portfolio consistency can be monitored and updated without repeatedly performing a manual eight-repository audit.

## P3.1 Introduce a versioned template contract

- [ ] Add a machine-readable template manifest containing baseline version, required files and policy identifiers.
- [ ] Record the adopted template version in each repository.
- [ ] Distinguish copied baseline files from locally owned extensions.
- [ ] Define backward-compatible and breaking template changes.
- [ ] Publish template release notes with migration instructions.

## P3.2 Build a portfolio drift checker

- [ ] Compare canonical files semantically rather than requiring byte identity where local sections are permitted.
- [ ] Detect missing baseline controls, obsolete versions and unauthorized divergence.
- [ ] Produce repository-by-repository findings with severity and remediation guidance.
- [ ] Exclude explicitly registered project-specific differences.
- [ ] Support local dry-run and CI execution.
- [ ] Never rewrite repositories automatically without a reviewed change set.

## P3.3 Automate post-creation configuration

- [ ] Create a dry-run-first bootstrap tool for repository description, topics, merge settings, labels and rulesets.
- [ ] Require explicit repository and profile arguments.
- [ ] Validate status-check contexts before creating required-check rules.
- [ ] Protect against applying mature-release rules to experimental repositories.
- [ ] Keep credentials outside the repository and logs.
- [ ] Retain the manual post-creation checklist as the fallback and audit record.

## P3.4 Reusable workflow strategy

- [ ] Evaluate versioned reusable workflows hosted by `GITHUB-TEMPLATE`.
- [ ] Pin consumers to released tags rather than `main`.
- [ ] Keep repository-specific test jobs local when they require Excel, special fixtures or release evidence.
- [ ] Add controlled update guidance for action and workflow versions.
- [ ] Test fork and private-repository behavior before adoption.

## P3.5 Portfolio quality reporting

- [ ] Create a compact dashboard covering template version, conformance, workflow health, protection and latest release.
- [ ] Track only actionable drift; avoid popularity or traffic metrics in the quality score.
- [ ] Flag stale documentation, unresolved placeholders and disabled gates.
- [ ] Review the baseline after material changes in at least two repositories.

## P3.6 Optional maturity improvements

- [ ] Add dependency/action update policy.
- [ ] Add Markdown style and link checking where signal exceeds maintenance cost.
- [ ] Add provenance attestations or checksums for release assets where appropriate.
- [ ] Add CodeQL or security scanning only where it produces meaningful VBA/tooling coverage.
- [ ] Define archival and deprecation procedures.
- [ ] Standardize social-preview assets without forcing identical visual identity.

**P3 exit criteria:**

- template versions and repository adoption are traceable;
- drift can be detected automatically;
- new repositories can be configured reproducibly;
- reusable workflows are version-pinned and proven;
- automation preserves repository-specific exceptions and requires review before changes.

---

# 6. Recommended execution order

| Sequence | Work package | Dependency |
| ---: | --- | --- |
| 1 | Finish P1.2 root `README.md` and post-creation checklist | P1.1 audit — complete |
| 2 | P1.4 placeholder vocabulary, manifest and scanners | P1.2 documentation |
| 3 | Finish the remaining P1.5 issue forms and issue-template configuration | Frozen intake decision; label sub-gate complete |
| 4 | P1.6 generic checker, fixtures and required workflow | P1.4–P1.5 contracts |
| 5 | Finish P1.7 release gate and evidence fixtures | Generic checker |
| 6 | P1.8 reusable VBA headers, neutral modules and harness | Profile and checker contracts |
| 7 | Finish P1.9 post-creation governance policy | Stable workflow context names |
| 8 | P1.10 pilot and v0.1.0 certification | All P1 controls |
| 9 | P2.1–P2.5 selective migration and conformance report | Certified template |
| 10 | P3 automation and drift control | Stable baseline and documented exceptions |

## 7. Global definition of done

The programme is complete only when:

- `GITHUB-TEMPLATE` has a certified release and can generate a green, usable repository;
- all seven repositories have been assessed against the same evidence-backed baseline;
- every adopted improvement is implemented and verified;
- every retained difference is documented as intentional;
- required status checks refer to live, successful workflow contexts;
- GitHub settings, labels and rulesets are covered by the post-creation contract;
- generic controls and specialized project gates remain clearly separated;
- no reusable template artifact contains donor-specific names, links, badges, versions or assumptions;
- future template drift can be detected without repeating the original manual audit.

## 8. Immediate next action

Finish **P1.2** by creating the template root `README.md` and `docs/POST_CREATION_CHECKLIST.md`. Then implement **P1.4 placeholder governance** before adding the remaining issue forms and executable gates, so every later artifact uses one validated token vocabulary from its first commit. The P1.5 label sub-gate is already complete.
