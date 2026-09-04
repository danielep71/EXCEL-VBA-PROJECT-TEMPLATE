# GitHub Template and Portfolio Standardization — Code Review and Implementation Plan

**Owner:** Daniele Penza

**Repository:** `danielep71/EXCEL-VBA-PROJECT-TEMPLATE`

**Review snapshot:** `ec38113ceab37ac452707578fea0d1b1f2304e1f`

**Review date:** 4 September 2026

**Plan revision:** 18 — P1 baseline certified and released

**Plan status:** Active — P1 complete; P2 and P3 remain

**Review-snapshot weighted score:** **6.6/10**

> **Temporary execution document.** Delete this file after the global definition
> of done is satisfied. Preserve durable decisions in maintained documentation,
> release notes, the portfolio audit and the final conformance report.

## 1. Executive verdict

At the review snapshot, `EXCEL-VBA-PROJECT-TEMPLATE` had a strong governance and
static-analysis foundation but was not yet a complete or certifiable VBA project
template. The P1 baseline is now certified and published as `v1.0.0`; the
remaining plan covers material P2 hardening, portfolio adoption and P3
automation.

At that snapshot, the best implemented material—documentation, immutable workflow dependencies,
the 20-label catalogue, deterministic reports and the 21-rule checker—is already
at roughly **8.6/10**. The overall repository scores **6.6/10** because a template
must also prove that a generated repository is usable. At the reviewed snapshot:

- the static gate passes **19/19** rules with zero findings;
- that same report validates **zero VBA components** and no public API manifest;
- no neutral façade, internal module or executable regression harness exists;
- no issue forms, post-creation checklist or executable release gate exists;
- the generated-project initialization path has 108 placeholder occurrences,
  but no deterministic initializer or end-to-end template fixture;
- live `main` protection prevents deletion and non-fast-forward updates, but does
  not yet require a pull request or the now-stable `Repository integrity` check;
- there is no protected `v*` tag policy, pilot repository or published template
  release.

**Historical release verdict at the review snapshot:** **NOT READY for v1.0.0.** The then-current green static check was
valid evidence for repository-text integrity only. It is not evidence that the
template can generate, compile, test, provision and release a VBA project.

## 2. Review method and evidence

The review treated this repository as a reusable Excel/VBA project template,
not as a Python utility repository and not as a finished VBA library.

Evidence inspected:

- the complete tracked tree at the exact review snapshot;
- all root governance and installation/release documents;
- [the repository profile](../.github/repository-profile.json);
- [the repository-quality checker](../tools/check_repo.py) and its synthetic
  positive/degraded fixtures;
- [the static-check workflow](../.github/workflows/static-checks.yml);
- [the label manifest](../.github/labels.json), reconciler and workflow;
- live repository metadata, merge settings, rulesets and workflow results;
- local checker execution, Python compilation and deterministic report output;
- the frozen donor decisions in [the portfolio audit](PORTFOLIO_AUDIT.md).

Measured baseline:

| Measure | Result |
| --- | ---: |
| Tracked files | 28 |
| VBA exports | 0 |
| Repository-quality rules | 19/19 pass |
| Quality findings | 0 |
| Label core | 20 labels |
| External action references | 4, all immutable SHA pins |
| Placeholder catalogue | 45 unique tokens |
| Placeholder use | 108 occurrences across 8 files |
| Issue forms | 0 |
| Release-gate workflows | 0 |
| Published template releases | 0 |
| Live `main` rules | deletion and non-fast-forward protection only |

The latest reviewed static workflow completed successfully at the review SHA.
The score nevertheless reflects the difference between **a green source tree**
and **a proven generated VBA repository**.

## 3. Detailed scoring

| Axis | Weight | Score | Weighted points | Review judgment |
| --- | ---: | ---: | ---: | --- |
| Documentation and first-use guidance | 12% | 9.0 | 10.8 | Excellent coverage, navigation, evidence boundaries and source-first guidance; too much duplicated operational text remains. |
| Structure and profile contract | 12% | 8.0 | 9.6 | Clear six-directory model and three profiles; profile requirements do not yet prove substantive source or tests. |
| Reusable VBA assets and harness | 18% | 2.0 | 3.6 | The required design is documented, but no importable VBA source, premium header or harness exists. |
| Static checker | 16% | 8.2 | 13.1 | Deterministic, read-only, dependency-free and unusually broad; several parsers and fixtures still have material false-negative paths. |
| Workflows and supply-chain controls | 10% | 8.6 | 8.6 | Least privilege, exact-SHA actions, bounded runs, evidence upload, terminal enforcement and content-pinned authoritative workflow validation are strong. Release validation remains missing. |
| Collaboration intake | 8% | 7.0 | 5.6 | PR intake is excellent; structured issue intake and routing configuration are absent. |
| Labels | 6% | 8.8 | 5.3 | Canonical core and reconciler are strong; profile/domain overlay selection is not wired to the repository profile. |
| Release and provenance | 8% | 6.0 | 4.8 | Documentation is strong; executable version/tag/evidence enforcement and protected release tags are absent. |
| Metadata and live governance | 5% | 5.5 | 2.8 | Template flag and merge settings are correct; description, topics and branch/tag enforcement are below the intended baseline. |
| Pilot and certification evidence | 5% | 3.0 | 1.5 | Hosted text checks are green, but no generated-repository, Excel or release certification exists. |
| **Overall** | **100%** |  | **65.7/100** | **6.6/10 — strong foundation, incomplete template product.** |

### 3.1 Strongest current elements

1. The checker reports deterministic JSON and Markdown, distinguishes policy
   findings from operational failure, and documents what it cannot prove.
2. Its self-test demonstrates read-only execution, repeated-output determinism,
   one positive fixture and one deliberately degraded fixture for every named
   rule.
3. The static workflow checks the exact SHA, disables persisted credentials,
   uses read-only permissions, uploads evidence even on failure and has one
   explicit terminal verdict.
4. External actions are pinned to immutable 40-character SHAs with audited
   semantic-version comments.
5. The label reconciler is dependency-free, idempotent, least-privilege and
   verifies an exact post-run match.
6. The root documentation consistently separates source inspection, VBA
   compilation, regression execution, specialist assurance and release
   certification.
7. The directory and profile documentation preserves valid differences between
   libraries, UI components and applications.
8. The repository avoids opaque Office binaries, secrets, locks and donor VBA
   source.

## 4. P1 findings — certification blockers

P1 findings must close before the template is tagged `v1.0.0` or used as the
formal benchmark for portfolio migration.

### P1-01 — Reusable VBA starter and harness — **complete**

**Original evidence:** `src/` and `tests/` contained instructional READMEs only.
The quality report said “All 0 VBA components” and “Observed 0 public
declarations.” The component map in repository policy was empty.

**Original impact:** The repository documented a VBA architecture without
demonstrating that its exported format, headers, façade/core boundary,
assertions, error handling and cleanup contract compiled together.

**Required implementation:**

- add a neutral public façade module, internal/core module and test module;
- use valid fixed VBA identifiers rather than placeholders inside exported names;
- add premium headers covering responsibility, public surface, dependencies,
  state ownership, error policy, worksheet safety and test seam;
- implement deterministic equality, tolerance and expected-error assertions;
- report cases, assertions, failures, environment and cleanup outcome;
- provide one small, passing, importable example with no donor dependency;
- register every component and role in the repository profile.

**Acceptance gate:** Clean VBE import and compilation; one test entry point runs
without editing and reports zero failures; source and test exports have distinct
identities; the static report validates at least one public, one internal and one
test component.

**Completion evidence — 2026-09-03:** Commit
`179024b0d29ae40775662bd582bebc82e710bd68` contains fixed-identifier
`ProjectFacade`, `ProjectCore`, `ProjectTests`, and `ProjectExample` exports, a
two-declaration public API manifest, exact import/test instructions, and policy
roles for one public, one internal, one test, and one example component. All
four exports were imported without edits into a clean VBE project and the
project compiled. `ProjectTests.RunProjectTests` then passed on Microsoft Excel
16.0, Windows 64-bit, Office 64-bit, VBA7+, with four cases, six assertions,
zero failures, complete execution, and cleanup passing. The separate example
returned `ProjectRatio(12, 4) = 3`. Hosted static checks passed 19/19 at the same
source SHA; all checker degradation fixtures and all three initialized-profile
fixtures also passed.

### P1-02 — Generated profiles can pass structure checks without substantive VBA — **complete**

**Evidence:** Profile definitions require non-empty directories, while the VBA
component map and public API manifest may remain empty. An instructional README
is sufficient to make a directory non-empty.

**Impact:** After replacing textual placeholders, a generated repository can be
green while still containing no production VBA and no regression source. This
is a template-level false assurance.

**Required implementation:**

- add generated-mode invariants for a minimum common component set;
- require at least one production and one test component for every profile;
- require the public façade/internal-core relationship for the starter baseline;
- define profile-specific component expectations without forcing unused forms,
  Ribbon or workbook modules;
- exercise `library`, `ui-component` and `application` against full-tree fixtures.

**Acceptance gate:** A README-only generated profile fails for the intended
reason; all three initialized profile fixtures pass; removing any mandatory
starter component fails one named rule.

**Completion evidence — 2026-09-03:** Source commit
`376b164f31da1c06bf5aef55fb618d9a00775223` makes each profile declare a
machine-readable `vba_contract` with minimum `public`, `internal`, and `test`
roles plus the exact façade, core, and regression starter paths. The new
`generated-vba-contract` rule evaluates all profiles in template mode and only
the selected profile in generated mode; its evidence records role minima,
observed counts, and required paths. The checker self-test passes 20 rules. The
initializer's three complete fixtures pass, 12 README-only or mandatory-removal
fixtures fail only the named contract rule with profile and role diagnostics,
and three example-removal fixtures pass. Hosted run
[`33806787297`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/actions/runs/33806787297)
checked out that exact source SHA and passed every initializer, checker,
repository, summary, artifact, and terminal-enforcement step.

### P1-03 — Placeholder and initialization governance — closed

**Evidence:** The catalogue contains 45 square-bracket tokens used 108 times in
eight files, but only two tokens are classified as required. Required, optional,
profile-specific and repeatable values are not represented as distinct schema
categories. The quick start requires manual edits across documents and policy.

The current Unreleased changelog also contains template-construction history
that a generated project must not present as its own product history.

**Impact:** Initialization is error-prone and can leave inconsistent names,
repository paths, evidence claims, unused profile sections or inherited template
history even when the operator follows the prose guide.

**Required implementation:**

- adopt one unambiguous machine-safe token syntax, preferably double-brace
  tokens, and remove competing square-bracket semantics;
- classify tokens as required, optional, profile-specific or repeatable;
- keep placeholders out of VBA identifiers and executable syntax;
- add a deterministic dry-run-first initializer that validates inputs, applies
  one profile, removes optional/template-only blocks and resets the changelog;
- fail atomically on missing, unknown or unused substitutions;
- retain manual instructions as a transparent fallback.

**Acceptance gate:** Each profile can be initialized from a clean template tree
with one documented command; a second run is idempotent; no token, template
identity, construction changelog entry or irrelevant profile block remains.

**Verdict: PASS.** The repository now uses one classified 14-token double-brace
schema. `tools/initialize_repository.py` validates missing, unknown, duplicated,
category-incompatible and unused inputs before mutation; defaults to a
content-digested dry-run; applies one explicit profile with rollback; removes
optional, repeatable, non-selected-profile and template-only content; resets the
generated changelog; records initialization; and accepts an identical second
run as a no-op. Its self-test exercises positive and negative fixtures for all
three profiles and runs the repository-quality gate over each generated tree.

### P1-04 — Structured issue intake — **complete**

**Evidence:** `.github/ISSUE_TEMPLATE/` does not exist. There are no bug,
feature or documentation forms and no issue-template configuration.

**Impact:** Generated repositories lose the evidence, environment, security
routing and scope discipline already required by the PR and contribution guides.

**Required implementation:** Add concise YAML forms for bug, feature and
documentation requests plus `config.yml`; use canonical labels; route security
reports privately; do not hard-code an assignee into the reusable forms.

**Acceptance gate:** GitHub accepts every form; titles, labels, required fields,
support links and security guidance are valid; blank issues follow the documented
policy; no donor identity or unavailable URL remains.

**Completion evidence — 2026-09-04:** Three native issue forms now require
bounded source, environment, expected-result, scope and validation evidence as
appropriate. Their labels come from the canonical manifest, reusable assignees
remain empty, blank issues are disabled, and the security contact link is
rewritten by the initializer for the generated repository without leaving a
token in YAML. Final source commit
[`bbc9209778cc8ad791a6e50604b457b9c7b5eabd`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/commit/bbc9209778cc8ad791a6e50604b457b9c7b5eabd)
passed exact-SHA static run
[`33850927030`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/actions/runs/33850927030).
The live chooser rendered all three forms, their prefixes, mandatory fields and
canonical labels with no assignee; blank intake was maintainer-only and both
security entries resolved to this repository's policy. The 21-rule checker
also includes a positive fixture and a targeted degraded-form fixture.

### P1-05 — Release integrity gate — complete

**Evidence:** `VERSION`, the changelog and the release guide exist, but no release
gate validates the candidate tag, dated changelog heading, required files,
placeholder state, evidence bundle, asset manifest or source/artifact identity.
No `v*` tag ruleset exists.

**Impact:** A release can contradict its source, documentation or evidence while
still passing the generic repository gate.

**Required implementation:**

- add a dependency-free release checker and positive/negative fixtures;
- require equality among `VERSION`, the dated changelog heading and `v*` tag;
- reject `0.0.0` and reject unresolved tokens or template identity in a
  generated-project release candidate;
- validate profile-specific evidence requirements and SHA-256 manifests;
- distinguish source-only library releases from optional UI/application binaries;
- add an immutable protected-tag policy before the first release.

**Acceptance gate:** A valid synthetic release passes; mismatched version/tag,
invalid changelog date, missing evidence, unapproved binary, wrong digest and
mutable-tag cases each fail deterministically.

**Completion evidence — 2026-09-04:** `tools/check_release.py` validates
an initialized candidate, canonical SemVer and dated changelog, exact full SHA,
external evidence, profile-required checks, annotated-tag type and target, and
optional staged assets against a sorted SHA-256 manifest. The versioned
`.github/release-policy.json` distinguishes source-only libraries from optional
UI/application binary distributions. Twenty deterministic fixtures cover all
three valid profiles and 17 named failure paths, including the `0.0.0` sentinel,
missing evidence/manifest, candidate and asset binding mismatches, unapproved
binaries, bad digests, lightweight tags and moved tags. The static workflow now
requires the self-test report and outcome. Source commit
[`8eab253ba9c5557bf3304928f2a33616b5ef3785`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/commit/8eab253ba9c5557bf3304928f2a33616b5ef3785)
passed exact-SHA hosted run
[`33854256898`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/actions/runs/33854256898),
including the release fixtures and fail-closed terminal verdict. Active ruleset
[`Protect releases`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/rules/22258556)
targets `refs/tags/v*`, has no bypass actors, and prevents updates, deletions and
force pushes. The documented pre-tag and annotated-tag verification commands
bind the release path to the certified candidate; publication remains P1-08.

### P1-06 — Authoritative workflow and structured-data validation — complete

**Evidence:** `validate_yaml_subset` checks indentation parity, mapping-like lines,
quotes and flow-bracket balance. It is not a complete YAML or GitHub Actions
parser and can accept invalid plain scalars or semantically invalid workflow
structures. The structured-data degraded fixture corrupts JSON only, so the YAML
and XML paths have no direct negative fixture.

**Impact:** The repository may report a green “structured data” rule for a
workflow GitHub cannot load, causing the required status itself to disappear.

**Required implementation:** Keep the dependency-free conservative check for
local portability, and add an authoritative, version-pinned workflow validation
step with verified provenance. Add negative YAML and XML fixtures and test the
workflow/action schema paths used by the template.

**Acceptance gate:** Known invalid YAML, duplicate/invalid workflow structure,
an invalid action reference and malformed XML all fail locally or in the hosted
gate before certification; current workflows remain green.

**Implementation status — 2026-09-04:** The portable 21-rule checker now adds
direct malformed-YAML and malformed-XML branch fixtures. The hosted gate
downloads actionlint 1.7.12, verifies the upstream Linux archive against its
published SHA-256 before execution, checks the reported version, validates all
tracked workflows, and exercises a valid local action plus five targeted
negative fixtures for invalid YAML, duplicate jobs, invalid job structure,
missing local-action metadata and a missing local entry point. Validator setup,
execution, summary publication and evidence upload are all terminal outcomes.
The initial hosted run correctly failed closed on two ShellCheck findings in the
workflow summary block. Commit
[`4982647aecf4a02aca5fd4c759fcd32ab6081a01`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/commit/4982647aecf4a02aca5fd4c759fcd32ab6081a01)
corrected those findings without weakening enforcement, and exact-SHA hosted
[`Static repository checks` run 33852502452](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/actions/runs/33852502452)
passed all steps. P1-06 is complete; the broader P2 checker-hardening findings
remain independently open.

### P1-07 — Post-creation and profile provisioning — complete

**Original evidence:** Template generation does not inherit labels, rulesets,
topics, security settings or merge settings. The original tree had no
`docs/POST_CREATION_CHECKLIST.md`, and the label workflow did not consume the
selected profile/domain overlay from repository policy.

**Resolved impact:** The checklist has now been applied to and read back from a
real generated application-profile pilot. The source-level label-selection and
evidence gaps are closed under P2-11, and the live pilot proves that repository
settings are provisioned separately from generated files.

**Required implementation:**

- create the post-creation checklist with exact expected settings and evidence;
- make profile/domain label selection explicit and end-to-end;
- document the first trusted label reconciliation for a new repository;
- include repository description, topics, issue/security settings, merge policy,
  branch rules, required contexts and tag rules;
- verify settings after application rather than assuming inheritance.

**Acceptance gate — passed:** A generated pilot’s files, labels, metadata,
merge settings, branch protection and tag policy match the selected profile and
checklist.

**Certification status — 2026-09-04:**
[`EXCEL-VBA-PILOT-APPLICATION`](https://github.com/danielep71/EXCEL-VBA-PILOT-APPLICATION)
records generated mode, profile `application`, repository identity and no domain
overlays at exact commit
[`286c329f267342627823ba969bc5c129244c830b`](https://github.com/danielep71/EXCEL-VBA-PILOT-APPLICATION/commit/286c329f267342627823ba969bc5c129244c830b).
Its exact-SHA
[`Repository integrity` run 33859831731](https://github.com/danielep71/EXCEL-VBA-PILOT-APPLICATION/actions/runs/33859831731)
passed; trusted
[`Sync issue labels` run 33859831711](https://github.com/danielep71/EXCEL-VBA-PILOT-APPLICATION/actions/runs/33859831711)
resolved 20 core labels, the application profile, no domain overlays and an exact
post-run match with no changes. UI/API read-back confirms the canonical
description and topics, Issues and private vulnerability reporting enabled,
Wiki, Projects, Discussions and Sponsorships disabled, all three merge methods,
branch updates and automatic head-branch deletion enabled, and auto-merge
disabled. Active rulesets
[`Protect main`](https://github.com/danielep71/EXCEL-VBA-PILOT-APPLICATION/rules/22263639)
and
[`Protect releases`](https://github.com/danielep71/EXCEL-VBA-PILOT-APPLICATION/rules/22264488)
have empty bypass lists; the former requires pull requests and strict
`Repository integrity`, while the latter targets `refs/tags/v*` and blocks tag
updates, deletion and force-pushes. Generated-mode initializer fixtures also
prove that maintained README/CHANGELOG evolution and removal of a once-recorded
optional preview asset do not invalidate the durable initialization contract.

### P1-08 — Complete: live governance, three pilots and `v1.0.0` certification

**Original evidence:** `main` had deletion and non-fast-forward protection only.
The stable `Repository integrity` job was not required and pull requests were
not required. There was no generated pilot, Excel evidence or template release.

**Impact:** The canonical source can change without the control it instructs
generated repositories to require, and there is no end-to-end evidence that the
template works outside its own template-mode tree.

**Required implementation:**

- require pull-request routing and strict `Repository integrity` on `main`, with
  zero mandatory approvals appropriate to the single-maintainer model;
- retain no default bypass and keep deletion/non-fast-forward protection;
- generate clean pilots for all three profiles and certify the common starter;
- run the VBA compile/regression path in a stated Excel environment;
- preserve exact-SHA evidence, then tag and publish `v1.0.0`.

**Acceptance gate:** All file and live-setting controls are green at the exact
release SHA; generated pilots pass; Excel evidence names host/build/bitness and
test counts; the protected tag targets that SHA.

**Live-governance status — 2026-09-04:** Active ruleset
[`Protect main`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/rules/22183319)
targets the default branch, has no bypass actor, requires pull requests with
zero approvals, requires strict GitHub Actions context `Repository integrity`,
and retains deletion and non-fast-forward protection. API read-back reports
`current_user_can_bypass: never`. All three profile pilots are fully provisioned
and green: application at `286c329f` (static run `33859831731`, label run
`33859831711`), library at `c5cb563a` (static run `33866091921`, label run
`33866091863`) and UI-component at `5ba6afb9` (static run `33866118364`, label
run `33866118333`). Each pilot has pull-request routing, strict
`Repository integrity`, immutable `v*` protection, an empty bypass list and
exact live settings/label read-back. The full commits, collection time, operator, profile,
domains, workflow links and live-setting links are retained in
[`PILOT_CERTIFICATION.md`](PILOT_CERTIFICATION.md).

**Release-candidate correction — 2026-09-04:** The generic release gate
originally accepted only initialized generated repositories, so it could not
certify the canonical template repository required by P1-08. The release policy
now defines a source-only `template` profile requiring all-profile pilot and
live-governance evidence in addition to the common static, compile and regression
checks. Template releases deliberately retain registered placeholders,
template-only blocks and their own construction history; generated-project
releases continue to reject inherited template state. Initialization resets the
template's released `VERSION` to the `0.0.0` generated-project sentinel. The
candidate sets `VERSION` to `1.0.0` and creates the dated initial release
changelog section. Protected
[`PR #36`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/pull/36)
established exact candidate
[`1144dd69112f6c238488c7888158d58b014fdc70`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/commit/1144dd69112f6c238488c7888158d58b014fdc70).
Its independent push-triggered
[`Repository integrity` run 33879467260](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/actions/runs/33879467260)
passed all steps, including 22 deterministic release fixtures, all three
initializer profiles, all 21 repository rules and six authoritative workflow
fixtures. Excel compilation and `ProjectTests.RunProjectTests` passed on
Microsoft Excel for Microsoft 365 MSO Version 2607, Build 16.0.20228.20188,
Windows 64-bit, Office 64-bit and VBA7+, with four cases, six assertions, zero
failures, complete execution and cleanup passing. The external evidence bundle
passed pre-tag and annotated-tag-reference validation with zero findings.
Protected annotated tag `v1.0.0` peels exactly to the certified candidate, and
the source-only
[`EXCEL-VBA-PROJECT-TEMPLATE v1.0.0`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/releases/tag/v1.0.0)
release was published on 4 September 2026. Both generated source archives were
read back successfully with the expected consumable tree; repository plumbing
is intentionally excluded by `.gitattributes`.

## 5. P2 findings — material hardening and portfolio adoption

P2 work begins after P1 certification unless a finding is safely corrected while
touching the same P1 code. It must not delay a necessary P1 safety correction.

| ID | Finding | Evidence and risk | Required correction | Acceptance |
| --- | --- | --- | --- | --- |
| P2-01 | The committed-whitespace rule is ineffective in CI | `git diff --check HEAD --` compares a clean checkout with itself. Its fixture changes the working tree, so the self-test does not expose the CI blind spot. | Check the candidate commit/range or the tracked tree against an empty tree; retain a separate working-tree mode locally. | A committed trailing-whitespace fixture fails in a clean checkout. |
| P2-02 | VBA jump labels are scoped to the file, not the procedure | The structure checker collects all labels in a component before validating jumps. A jump can incorrectly resolve to a label in another procedure. | Parse procedure boundaries and resolve each `GoTo`/`Resume` target within its owning procedure. | Cross-procedure-label fixture fails; valid local handlers pass. |
| P2-03 | Conditional-compilation analysis is approximate | Nested `#If`, `#ElseIf` and parent-active state are not modelled as a full branch stack, which can miss a non-PtrSafe declaration in a reachable VBA7 branch. | Track parent activity, prior branch selection and nested alternatives; add VBA6/VBA7/Win64 fixtures. | Each reachable VBA7 declaration requires `PtrSafe`; legacy branches remain allowed when policy permits. |
| P2-04 | Public API extraction is incomplete | Only explicitly `Public` Sub/Function/Property/Enum/Type/Const declarations on one physical line are recorded. Default-public procedures, variables, events, declares and continuations can escape the manifest. | Either prohibit implicit visibility and unsupported public declarations or extend the extractor; define when a manifest becomes mandatory. | Every supported public declaration appears exactly once in the manifest and collision checks. |
| P2-05 | Local action references are not validated | Workflow references beginning with `./` bypass pin checks but are not checked for existence or action metadata. | Resolve local action paths and require tracked `action.yml` or `action.yaml`. | Missing or malformed local actions fail. |
| P2-06 | SemVer and changelog checks are intentionally shallow | Pre-release numeric identifiers with leading zero, actual calendar validity, ordering, duplicate versions and comparison links are not enforced. | Put strict release semantics in the release gate and keep the generic rule narrowly documented. | SemVer edge fixtures and invalid dates fail the release gate. |
| P2-07 | Rule-level fixtures do not cover important branches | One degraded fixture per rule proves routing, not every sub-check: YAML/XML, secret patterns, BOM/VBA endings, changelog, Office locks and several VBA paths remain unexercised. | Add a table-driven branch matrix and coverage report for checker logic. | Every blocking branch has a positive or negative fixture; coverage exclusions are explicit. |
| P2-08 | The checker is a 2,600-plus-line single module | Portability is excellent, but configuration, parsers, reporters and fixtures are tightly coupled and costly to review. | Preserve a single-file distributable if desired, but develop from small tested modules or clearly separated internal sections with generated bundling. | Runtime artifact stays dependency-light; maintainers can test parsers independently. |
| P2-09 | Documentation duplicates evolving contracts | README, contribution, installation, release, security and conduct files repeat some source/evidence rules; the combined onboarding surface exceeds 10,000 words. | Assign one authority per contract, shorten root navigation and link to specialized guidance; retain all substantive protections. | No contradictory duplicate; first-use path is short; detailed policy remains discoverable. |
| P2-10 — **complete** | Repository metadata was below template quality | The live repository now has a specific Excel/VBA template description, ten relevant topics, an uploaded social preview, intentional feature/merge settings and enabled private vulnerability reporting. | Complete. Preserve the tracked preview source and initializer lifecycle contract; re-verify live settings during release certification. | Metadata explains Excel/VBA scope and template profiles; enabled features are intentional. |
| P2-11 — **complete** | Label overlays were declarative but not operationally selected | The reconciler now resolves profile and optional domain overlays only from `.github/repository-profile.json`; unknown, duplicate and unsorted selections fail before mutation. Plan/apply summaries name both policy files and enumerate the core, selected profile, selected domain and complete resolved label sets. | Complete. Preserve the versioned-policy trust boundary and the catalogue assertions in the self-test. | Non-empty profile/domain fixtures resolve through policy, summary evidence contains every selected label, reconciliation remains idempotent and exact-match verification prevents unintended pruning. |
| P2-12 | Drift is not detected unless label files change | Label reconciliation runs on changes to its three files or manual dispatch only. Manual live-label edits can persist unnoticed. | Add a safe scheduled read/plan check or a portfolio drift control; reserve writes for trusted explicit events. | Live drift is reported with no untrusted mutation and can be reconciled deliberately. |

**P2-11 completion evidence — 2026-09-04:** Commit
[`978045f0187c6c0856d333c23fc459457534d7b5`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/commit/978045f0187c6c0856d333c23fc459457534d7b5)
added the resolved-catalogue evidence and assertions. Exact-SHA
[`Sync issue labels` run 33857409303](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/actions/runs/33857409303)
passed the non-empty profile/domain self-test, identified the versioned policy
and manifest paths, enumerated the complete 20-label template selection, made
zero changes and verified an exact post-run match. Exact-SHA
[`Static repository checks` run 33857409394](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/actions/runs/33857409394)
also passed every repository, workflow, release and initialization gate.

**P2-10 completion evidence — 2026-09-04:** PR
[#33](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/pull/33)
merged the tracked social-preview source and deterministic generated-profile
lifecycle as commit
[`fb77951119a70593782d7537b53a67899583213a`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/commit/fb77951119a70593782d7537b53a67899583213a).
Exact-SHA main
[`Repository integrity` run 33867905966](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/actions/runs/33867905966)
passed. Live read-back confirmed the specific description, ten topics, disabled
Wiki/Projects/Discussions/Sponsorships, intentional merge controls, the uploaded
1280-by-640 preview and private vulnerability reporting enabled.

### 5.1 P2 portfolio migration rules

After `v1.0.0`, assess every existing repository against the certified baseline.
Classify each difference as:

- `REQUIRED` — missing universal control or correctness defect;
- `ADOPT` — beneficial improvement with acceptable regression risk;
- `KEEP` — stronger or legitimate project-specific divergence;
- `DEFER` — useful but not justified in the current release;
- `NOT APPLICABLE` — irrelevant to the selected profile.

Migration order:

1. repair the two repositories with unresolved source/test identity defects;
2. align mature UI repositories without weakening UI, lifecycle or packaging
   evidence;
3. align analytical/library repositories without weakening numerical,
   performance, public-API or release-provenance gates;
4. publish `docs/PORTFOLIO_CONFORMANCE.md` with scores and every accepted
   exception.

No folder, workflow or document is renamed solely to imitate the template. No
domain-specific gate is replaced by a weaker generic check.

## 6. P3 findings — automation and long-term maturity

| ID | Improvement | Intended outcome |
| --- | --- | --- |
| P3-01 | Version the template contract independently from generated project versions. | Every repository records the baseline version it adopted and the applicable migration notes. |
| P3-02 | Build semantic portfolio drift detection. | Missing controls and unauthorized divergence are reported without requiring byte identity for locally owned sections. |
| P3-03 | Add a dry-run-first repository provisioner. | Description, topics, labels, merge settings and rulesets can be applied and verified reproducibly with explicit repository/profile arguments. |
| P3-04 | Publish versioned reusable workflows. | Consumers pin stable workflow releases while keeping Excel, numerical and UI jobs local where necessary. |
| P3-05 | Add controlled dependency-update policy. | Action SHAs and tool versions are updated through reviewed changes with provenance and rollback information. |
| P3-06 | Add portfolio quality reporting. | Template version, conformance, workflow health, protection and release state are visible without popularity metrics contaminating quality scores. |
| P3-07 | Add advanced release provenance where justified. | UI/application assets can carry checksums, build-environment records, signatures or attestations without imposing binary releases on libraries. |
| P3-08 | Add external-link and documentation-drift review. | Local links remain blocking; network-dependent links are checked separately with retry and allow-list policy. |
| P3-09 | Add an optional Windows/Excel reusable evidence pattern. | Eligible projects can produce exact-SHA compilation/regression evidence without pretending every repository has the same runner or trust boundary. |

## 7. Findings-driven implementation sequence

| Sequence | Work package | Findings closed | Dependency | Estimated effort |
| ---: | --- | --- | --- | ---: |
| 1 | Finalize placeholder schema and deterministic initializer — **complete** | P1-03 | None | 1.5–2.5 days |
| 2 | Add neutral VBA façade/core modules, premium headers and harness — **complete** | P1-01 | Placeholder identifier policy | 2.5–4 days |
| 3 | Enforce substantive generated-profile contracts and full-tree fixtures — **complete** | P1-02 | Starter assets | 1.5–2.5 days |
| 4 | Add issue forms and post-creation checklist; wire label profile selection — **complete** | P1-04, P1-07, P2-11 | Final token/profile schema | 1–2 days |
| 5 | Add authoritative workflow validation and close checker false negatives — **P1-06 complete; P2 hardening remains** | P1-06, P2-01–P2-07 | Stable checker/profile schema | 2–3.5 days |
| 6 | Add release gate, fixtures and protected-tag specification — **complete** | P1-05 | Static and placeholder gates | 1.5–2.5 days |
| 7 | Strengthen live `main` protection and complete metadata — **complete** | P1-08, P2-10 | Stable required context | 0.5–1 day |
| 8 | Generate profile pilots, run Excel evidence and publish `v1.0.0` — **complete** | P1-08 | All prior P1 packages | 1.5–3 days plus Excel access |
| 9 | Perform selective portfolio migration and conformance review | P2 portfolio work | Certified `v1.0.0` | Repository-specific |
| 10 | Add versioned drift/provisioning/reusable-workflow automation | P3-01–P3-09 | Stable P1/P2 contracts | 4–8 days incrementally |

The earlier **12–21 person-day** estimate described the work from the review
snapshot to certification and is now retired. Remaining effort belongs to the
P2/P3 packages and portfolio-specific adoption work below.

## 8. Delivery status

| Package | Status after review | Evidence or remaining boundary |
| --- | --- | --- |
| Portfolio audit and donor decisions | **Complete** | Frozen decision record remains authoritative |
| Root documentation | **Strong / incomplete** | Core documents and exact post-creation checklist exist; deduplication remains |
| Canonical directories | **Complete** | All profiles retain substantive façade, core, and test assets; optional UI assets remain profile-driven |
| Placeholder governance | **Complete** | Classified schema, dry-run/apply initializer, manual fallback and all-profile fixtures pass |
| Collaboration files | **Complete** | PR template, three issue forms/config, policy-bound label selection, complete resolved-catalogue evidence and all-profile pilot live read-back are certified |
| Repository-quality checker | **P1 complete / P2 hardening open** | The portable 21-rule gate includes direct YAML/XML fixtures; content-pinned authoritative workflow validation passed at the exact hosted SHA, while identified P2 false negatives remain |
| Release controls | **Complete** | Dependency-free gate, four-profile policy, evidence contract and 22 fixtures pass locally and at the exact release SHA; active immutable `v*` protection was read back with no bypass |
| Reusable VBA assets | **Complete** | Four governed exports, public API manifest, clean import/compile, 4-case/6-assertion Excel pass, and exact-SHA hosted checks evidenced |
| Live governance | **Complete** | The template and all three pilots require PRs and strict `Repository integrity` with no bypass; each also has active immutable release-tag policy and verified settings |
| Pilot and v1.0.0 | **Complete** | Application, library and UI-component pilots are green and fully provisioned; exact-candidate Excel evidence, annotated tag and source-only release are certified |
| Portfolio migration | **Pre-alignment only** | Starts formally after v1.0.0 |
| Drift automation | **Not started** | P3 dependency on stable contracts |

The earlier 19-rule checker verdict is not discarded: it remains a valid pass
for the scope it measured. This review reopens the package because template
certification requires additional end-to-end and parser guarantees that the
synthetic rule fixtures did not claim to provide.

## 9. P1 certification gate

P1 passes only when all statements below are true:

- [x] One token grammar and category schema governs every reusable placeholder.
- [x] The initializer is dry-run-first, atomic and idempotent for all profiles.
- [x] Generated changelog and documentation contain no template-construction history.
- [x] Neutral façade, core and test modules import and compile without editing.
- [x] The deterministic harness passes and reports exact environment/count evidence.
- [x] Generated mode rejects repositories with no substantive production/test VBA.
- [x] Bug, feature and documentation forms plus configuration are valid.
- [x] Profile/domain labels resolve from versioned repository policy.
- [x] YAML/workflow validation is authoritative and version-pinned. *(Exact-SHA hosted run 33852502452 passed.)*
- [x] The release gate passes positive fixtures and rejects every named negative fixture. *(Exact-SHA hosted run 33854256898 passed all 20 fixtures.)*
- [x] `main` requires PR routing and the exact live `Repository integrity` context.
- [x] `v*` tags are protected against deletion and mutation; the documented creation path verifies the certified exact SHA.
- [x] The post-creation checklist is executed and evidenced on generated pilots.
- [x] Static, compile, regression and release evidence bind to the same candidate SHA.
- [x] `v1.0.0` is published from that protected exact SHA.

## 10. P2 and P3 exit gates

P2 passes when every portfolio repository has a profile, a conformance result and
documented deviations; generic controls are aligned; stronger specialist gates
remain intact; the two source-identity defects are resolved; and all required
live contexts are green.

P3 passes when template adoption is versioned, drift is detected semantically,
repository provisioning is reproducible, reusable workflows are released and
pinned, and automation cannot rewrite a repository without a reviewed change.

## 11. Immediate next action

Begin **P2-01**: make committed trailing whitespace fail in a clean hosted
checkout by validating the candidate commit/range rather than comparing `HEAD`
with itself, and add a deterministic committed-whitespace fixture.
