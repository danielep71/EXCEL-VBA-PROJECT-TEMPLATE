<div align="center">

# ⚡ [PROJECT_NAME]

### [PROJECT_TAGLINE]

**[ONE-SENTENCE PROJECT PURPOSE]**

<br>

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-source--first-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#requirements)
[![Profile](https://img.shields.io/badge/Profile-select_one-6f42c1?style=for-the-badge)](#supported-profiles)
[![Status](https://img.shields.io/badge/Status-template-d97706?style=for-the-badge)](#project-status)
[![Version](https://img.shields.io/badge/Version-VERSION_file-0969da?style=for-the-badge)](VERSION)
[![License](https://img.shields.io/badge/License-MIT-2ea44f?style=for-the-badge)](LICENSE)

<br>

[![Static checks](https://github.com/[REPOSITORY PATH]/actions/workflows/static-checks.yml/badge.svg?branch=main)](https://github.com/[REPOSITORY PATH]/actions/workflows/static-checks.yml)
[![Release](https://img.shields.io/github/v/release/[REPOSITORY PATH]?style=flat-square&label=release&color=217346)](https://github.com/[REPOSITORY PATH]/releases)
[![Issues](https://img.shields.io/github/issues/[REPOSITORY PATH]?style=flat-square&color=d73a49)](https://github.com/[REPOSITORY PATH]/issues)
[![Last commit](https://img.shields.io/github/last-commit/[REPOSITORY PATH]?style=flat-square&color=0969da)](https://github.com/[REPOSITORY PATH]/commits/main)

<br>

**[PUBLIC API OR USER SURFACE] · [INTERNAL OR CORE ENGINE] · [TEST OR EVIDENCE SYSTEM]**

[Overview](#what-this-project-is)
&nbsp;·&nbsp;
[Quick start](#quick-start)
&nbsp;·&nbsp;
[Profiles](#supported-profiles)
&nbsp;·&nbsp;
[Architecture](#architecture)
&nbsp;·&nbsp;
[Quality](#quality-and-assurance)
&nbsp;·&nbsp;
[Installation](INSTALLATION.md)
&nbsp;·&nbsp;
[Documentation](#documentation-map)
&nbsp;·&nbsp;
[Security](SECURITY.md)

</div>

---

<!-- TEMPLATE OPTIONAL: Replace [OPTIONAL PATH] with a tracked banner image and
     remove this comment, or delete the complete block.
<p align="center">
  <img src="[OPTIONAL PATH]"
       alt="[PROJECT_NAME] — [PROJECT_TAGLINE]"
       width="100%">
</p>

---
-->

> [!IMPORTANT]
> This repository is still in template mode. Before publishing or developing a
> generated project, select one profile, replace every registered placeholder,
> remove the remaining GITHUB-TEMPLATE identity text, and run the complete
> quality sequence. A green static gate does not constitute Excel execution or
> release certification.

<a id="what-this-project-is"></a>

## ✨ What this project is

[PROJECT_NAME] is an exported-source-first Excel/VBA project designed around
reviewable source, explicit contracts, deterministic checks and evidence bound
to an exact commit.

The repository treats text source, tests, documentation, workflows and evidence
as authoritative. Office workbooks and add-ins are generated or release
artifacts unless an exact path is deliberately approved in the repository
profile.

### At a glance

| Principle | Canonical expectation |
| --- | --- |
| Source of truth | Exported VBA and versioned text, not an opaque workbook |
| Public surface | Document supported entry points separately from internal and host callbacks |
| Failure behavior | Define invalid-input, error, cleanup and state-restoration contracts |
| Verification | Separate static inspection, Excel execution and release certification |
| Evidence | Bind every claim to the exact source SHA, environment and artifact tested |
| Portability | State supported Excel hosts, Office bitness and deployment models explicitly |

## ⭐ Why use this structure

- **Reviewable by design.** Exported components, contracts and workflows produce
  meaningful diffs.
- **Profile-aware.** Libraries, UI components and applications share one core
  without pretending that their runtime evidence is identical.
- **Source-first.** Production source is separated from tests, examples,
  generated workbooks and release packages.
- **Evidence-led.** Repository checks, Excel regression and specialist assurance
  have distinct meanings.
- **Safe to extend.** Domain-specific workflows remain additive and may never be
  replaced by a weaker generic check.
- **Honest about boundaries.** Untested environments and unresolved limitations
  remain visible.

<a id="project-status"></a>

## 🚦 Project status

| Item | Current value |
| --- | --- |
| Lifecycle | Template — initialize before project use |
| Version source | [VERSION](VERSION) |
| Repository profile | Select exactly one supported profile |
| Static integrity | [Static repository checks](.github/workflows/static-checks.yml) |
| Excel/VBA execution | Project-specific evidence required |
| Release state | No generated project is certified by the template itself |

When this README belongs to a generated project, replace this table with the
project’s actual lifecycle, supported version, evidence and release status.
Never publish inherited test counts, compatibility claims or badges that were
not produced for that project.

<a id="quick-start"></a>

# ⚡ Quick start

## 1. Generate and clone the repository

Create a new repository with **Use this template**, then clone the generated
repository:

~~~bash
git clone https://github.com/[REPOSITORY PATH].git
~~~

Do not place a generated project's domain implementation in GITHUB-TEMPLATE.
Maintain template policy separately from the repositories created from it.

## 2. Select and configure one profile

Edit [.github/repository-profile.json](.github/repository-profile.json):

1. change <code>mode</code> from <code>template</code> to
   <code>generated</code>;
2. set <code>profile</code> to <code>library</code>,
   <code>ui-component</code>, or <code>application</code>;
3. replace the repository value with the generated <code>owner/name</code>;
4. declare required files and directories;
5. assign every governed VBA component a role; and
6. allow an Office binary only through an exact, justified path.

Configuration changes are policy changes and receive the same review as checker
logic.

## 3. Remove template identity

- Replace every registered square-bracket placeholder.
- Remove this initialization guidance when it is no longer useful.
- Delete template-construction records that do not belong to the project,
  especially <code>docs/PORTFOLIO_AUDIT.md</code> and
  <code>docs/IMPLEMENTATION_PLAN.md</code>.
- Retain only profile-relevant examples and optional sections.
- Replace all template badges with project evidence or remove them.

## 4. Add authoritative source and tests

- Add [VERSIONED SOURCE FILES].
- Import components in [IMPORT ORDER].
- Keep production components under [src/](src/).
- Keep regression components, independent fixtures and test documentation under
  [tests/](tests/).
- Update [.github/repository-profile.json](.github/repository-profile.json) so
  the configured component manifest matches the tracked tree.

## 5. Run the local quality sequence

~~~bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root . \
  --output test-results/static-checks.json \
  --summary test-results/static-checks.md
[STATIC CHECK COMMAND]
~~~

Then execute the applicable Excel/VBA entry point:

~~~text
[EXCEL REGRESSION ENTRY POINT]
[REQUIRED UI OR MANUAL SMOKE TESTS]
~~~

Record only evidence actually produced for the exact candidate.

## 6. Complete repository provisioning

GitHub template generation copies files, not repository settings. After
creation:

- apply the canonical labels;
- configure repository topics, description and social preview;
- enable the approved merge methods and automatic branch deletion;
- create the applicable branch and tag rulesets;
- select <code>Repository integrity</code> and any stronger project checks as
  required; and
- verify security, issue-routing and release settings.

<a id="supported-profiles"></a>

# 🧭 Supported profiles

Choose one profile. A project may add specialist controls, but it must not
silently combine profiles to avoid a requirement.

| Profile | Intended use | Typical production structure | Additional evidence |
| --- | --- | --- | --- |
| <code>library</code> | Reusable VBA functions, services or numerical components | Public façade modules plus internal/core modules | Public API, caller contract and focused regression |
| <code>ui-component</code> | UserForms, Ribbon controls, window or worksheet UI | Modules, classes, forms and host callbacks | UI state, cleanup, recovery, DPI/accessibility and lifecycle checks |
| <code>application</code> | End-to-end workbook or add-in workflows | Modules, classes and workbook/application integration | Startup, shutdown, migration, scenario and package tests |

The selected profile changes required structure and evidence. It never weakens
documentation, security, source integrity, action pinning or release provenance.

<a id="architecture"></a>

# 🏗️ Architecture

## Canonical repository map

| Path | Responsibility | Keep when |
| --- | --- | --- |
| [src/](src/) | Authoritative exported production VBA | Always |
| [tests/](tests/) | Regression modules, fixtures and evidence instructions | Always |
| [examples/](examples/) | Minimal supported examples or demonstrations | The project can maintain them |
| [assets/](assets/) | Documentation images and non-source visual assets | Referenced by tracked documentation |
| [docs/](docs/) | Architecture, API, compatibility and durable contracts | Material exceeds the root overview |
| [tools/](tools/) | Deterministic validation, evidence and packaging utilities | Tooling is versioned and documented |
| [.github/](.github/) | Workflows, contribution templates and declarative policy | Always |

See [REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md) for ownership rules,
VBA separation and legitimate profile-specific alternatives.

## VBA component separation

| Layer | Responsibility | Dependency direction |
| --- | --- | --- |
| Public façade | Stable consumer-facing procedures, functions and types | May call internal/core code |
| Internal/core | Algorithms, state management and implementation details | Must not depend on tests |
| UI and host adapters | Forms, Ribbon callbacks and workbook/application integration | Translate host events into supported operations |
| Regression tests | Contract, boundary, failure and cleanup verification | May exercise governed seams without becoming production API |

~~~mermaid
flowchart LR
    Consumer["Workbook or add-in"] --> Public["Public façade"]
    UI["UI and host adapters"] --> Public
    Public --> Core["Internal core"]
    Tests["Regression tests"] --> Public
    Tests --> Core
~~~

The diagram is a dependency guide, not a mandatory module count. Record any
justified alternative in the project architecture documentation.

## Source contract

- Exported <code>.bas</code>, <code>.cls</code> and <code>.frm</code> files are
  the reviewable source of truth.
- Production and regression components have distinct tracked identities.
- Exported component names match filenames and remain unique
  case-insensitively.
- VBA source uses Windows-1252-compatible text, CRLF line endings and
  <code>Option Explicit</code>.
- Internal standard modules declare <code>Option Private Module</code>.
- UserForm <code>.frx</code> companions are tracked only with their authoritative
  <code>.frm</code> export.
- Generated Office packages remain untracked unless the profile and release
  process explicitly govern an exact path.

# 🛡️ Engineering contracts

## Source-first

A workbook or add-in is not a substitute for exported source. Release artifacts
must be reproducible from, or cryptographically bound to, the exact versioned
source.

## Caller-owned state

Code must restore only state it successfully acquired or changed. Cleanup paths
receive the same attention as successful paths, particularly for
<code>Application</code> state, Windows resources, callbacks and modeless UI.

## Deterministic failures

Public behavior defines invalid inputs, expected error values or numbers,
partial-success rules, cleanup and diagnostics. Tests assert the contract rather
than implementation accidents.

## No self-referencing evidence

Expected results come from an independent oracle, frozen contract or explicitly
reviewed reference—not from the implementation under test.

## Environment honesty

Source inspection does not compile VBA. One Office bitness does not execute the
other conditional branch. A workbook opened manually does not prove a headless
runner. State these boundaries directly.

<a id="quality-and-assurance"></a>

# ✅ Quality and assurance

## Assurance ladder

| Layer | What it can establish | What it cannot establish |
| --- | --- | --- |
| Repository integrity | Structure, links, placeholders, identities, policy files, action pins and exported-source hygiene | VBA compilation or Excel behavior |
| VBA compilation | Importability and compile-clean source in a stated host | Functional correctness |
| Regression execution | Contract behavior for executed cases and environment | Untested cases or platforms |
| Specialist assurance | UI, numerical, performance, lifecycle or packaging claims | Broader claims than its measured scope |
| Release certification | Version, source, artifact and evidence consistency | Correctness beyond the recorded evidence |

## Canonical static gate

The [static-check workflow](.github/workflows/static-checks.yml):

- runs against the exact SHA under review;
- uses pinned external actions and read-only repository permissions;
- self-tests one passing fixture and one degraded fixture per rule;
- validates the configured repository tree;
- publishes a readable job summary;
- uploads deterministic JSON and Markdown evidence; and
- enforces every intermediate outcome through one terminal verdict.

Run it locally with the commands shown in [Quick start](#quick-start).

## Evidence record

Replace the entries below with facts from the exact candidate:

| Evidence | Result |
| --- | --- |
| Commit SHA | Full 40-character candidate SHA |
| Static checks | Command, rule count and workflow URL |
| VBA compilation | Host, version, build, bitness and outcome |
| Regression | Entry point, cases, assertions, failures and cleanup |
| Specialist checks | [EVIDENCE ACTUALLY PRODUCED, INCLUDING ENVIRONMENT AND LIMITATIONS.] |
| Release artifact | Name, SHA-256 digest and source relationship |

> [!NOTE]
> Hosted static checks validate repository evidence only. They do not execute
> Excel, prove numerical accuracy, exercise UI state, validate a workbook
> lifecycle, or certify a release package.

<a id="requirements"></a>

# 💻 Requirements

Replace this table with the environments actually supported and tested.

| Item | Requirement |
| --- | --- |
| Excel / Office host | [SUPPORTED EXCEL / OFFICE HOSTS] |
| Office bitness | [SUPPORTED OFFICE BITNESS] |
| Operating system | State every supported platform and material limitation |
| VBA references | List required references or state that none are non-standard |
| Deployment model | Embedded source, add-in, workbook, or another supported package |
| Trust settings | Document macro, trusted-location and signing requirements |

# 📦 Installation

Use [INSTALLATION.md](INSTALLATION.md) as the authoritative deployment and
removal guide. At minimum it must state:

1. the exact production source manifest;
2. the required import order;
3. workbook, add-in or application integration steps;
4. references, trust settings and host requirements;
5. verification and smoke-test procedures; and
6. rollback, removal and recovery steps.

Do not instruct users to import test modules as production dependencies.

<a id="documentation-map"></a>

# 📚 Documentation map

| Document | Purpose |
| --- | --- |
| [Installation](INSTALLATION.md) | Requirements, import, deployment, verification and removal |
| [Contributing](CONTRIBUTING.md) | Branch, source, testing, evidence and review standards |
| [Changelog](CHANGELOG.md) | Unreleased work, version history and compatibility |
| [Security](SECURITY.md) | Supported versions and private vulnerability reporting |
| [Releasing](RELEASING.md) | Versioning, certification, artifacts, provenance and recovery |
| [Repository structure](docs/REPOSITORY_STRUCTURE.md) | Canonical directory and VBA ownership |
| [Documentation index](docs/README.md) | Durable architecture, API and specialist contracts |
| [Code of Conduct](CODE_OF_CONDUCT.md) | Participation expectations |

Keep one authoritative statement for each contract and link to it. Do not copy
the same evolving rule into the README, Wiki and multiple documents.

# 🆘 Recovery and troubleshooting

Project-specific recovery guidance belongs in
[INSTALLATION.md](INSTALLATION.md). It should cover:

- incomplete imports or missing references;
- failed startup and partially initialized state;
- workbook or application state left altered after an error;
- UI controls, callbacks or shortcuts that remain installed;
- removal of generated packages and local evidence; and
- the safe route back to the last certified source and artifact.

Never advise deleting evidence, suppressing errors or bypassing a gate merely to
obtain a green result.

# 🔐 Security and trust

- Report suspected vulnerabilities privately through
  [SECURITY.md](SECURITY.md).
- Never commit credentials, private keys, confidential workbooks, customer data
  or private review material.
- Treat macros, add-ins, Ribbon XML, WinAPI declarations and generated Office
  packages as trust-boundary changes.
- Keep workflow permissions minimal and external actions pinned to audited
  immutable SHAs.
- Do not present unsigned or unverified packages as certified releases.

# ⚠️ Known limitations and boundaries

- [UNRESOLVED OR DELIBERATE BOUNDARY.]
- Add every untested Excel version, Office bitness, locale, deployment mode or
  specialist scenario.
- Distinguish unsupported behavior from behavior that is merely untested.
- Link each material defect or deferred boundary to its owning issue.

# 🧭 Versioning and releases

- [VERSION](VERSION) is the authoritative version without a leading
  <code>v</code>.
- [CHANGELOG.md](CHANGELOG.md) stages material changes under
  <code>Unreleased</code>.
- Release tags use <code>vMAJOR.MINOR.PATCH</code> and target the exact certified
  commit.
- Release notes identify evidence, known limitations and compatibility impact.
- Published artifacts record SHA-256 digests and their relationship to source.

Follow [RELEASING.md](RELEASING.md). A tag or uploaded workbook is not certified
merely because it exists.

# 🤝 Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) before changing source, tests,
documentation, workflows or release evidence.

Every pull request should:

- state one coherent outcome and its supported-surface impact;
- identify the exact source and environment tested;
- update tests and documentation with the change;
- record skipped checks and limitations honestly;
- explain compatibility, migration, risk and rollback; and
- leave the repository-quality gate green.

# 📄 License

This template uses the [MIT License](LICENSE). Confirm that MIT is appropriate
for the generated project before publication. If the license changes, update
the root license, badges, documentation, package notices and release statements
consistently.
