<div align="center">

# [PROJECT_NAME]

### [PROJECT_TAGLINE]

[![Static checks](https://github.com/[REPOSITORY PATH]/actions/workflows/static-checks.yml/badge.svg)](https://github.com/[REPOSITORY PATH]/actions/workflows/static-checks.yml)
[![VBA](https://img.shields.io/badge/VBA-source--first-217346?style=flat-square)](src/)
[![Profiles](https://img.shields.io/badge/profiles-library%20%7C%20UI%20%7C%20application-0969da?style=flat-square)](#supported-profiles)
[![License](https://img.shields.io/badge/license-MIT-6f42c1?style=flat-square)](LICENSE)

**[ONE-SENTENCE PROJECT PURPOSE]**

</div>

---

> [!IMPORTANT]
> This repository was created from GITHUB-TEMPLATE. Before publishing or
> developing the project, complete the initialization checklist below, replace
> every registered placeholder, select one repository profile, and run the
> canonical quality gate.

## Overview

[PROJECT_NAME] is an exported-source-first Excel/VBA repository. Text source,
tests, documentation, workflows and evidence are authoritative; Office files
are generated or release artifacts unless an exact path is explicitly allowed
by the selected profile.

The template provides a common repository contract without pretending that
static inspection executes Excel. Project-specific compilation, regression,
numerical, UI and packaging checks remain additive requirements.

<a id="supported-profiles"></a>

## Supported profiles

Choose exactly one profile in
[repository-profile.json](.github/repository-profile.json).

| Profile | Use when | Typical governed content |
| --- | --- | --- |
| library | The repository exposes reusable VBA functions or services. | Public façade modules, internal/core modules, API contract and regression modules |
| ui-component | The repository supplies forms, Ribbon controls or host UI behavior. | Standard modules, class modules, UserForms, callbacks and UI-state tests |
| application | The repository owns an end-to-end workbook or add-in workflow. | Modules, classes, workbook integration, application lifecycle and scenario tests |

The profile changes required structure and evidence; it does not weaken the
common documentation, security, source-integrity or workflow rules.

## Initialize a generated repository

1. Create a new repository using **Use this template**, then clone it:

   ~~~bash
   git clone https://github.com/[REPOSITORY PATH].git
   ~~~

2. Edit [.github/repository-profile.json](.github/repository-profile.json):
   - change mode from template to generated;
   - set profile to library, ui-component, or application;
   - replace the repository value with the new owner/name;
   - declare required paths, permitted Office binaries and every governed VBA
     component with its role.
3. Replace all registered square-bracket placeholders, including the project
   name, tagline, purpose, repository path, supported hosts and test commands.
   Remove remaining GITHUB-TEMPLATE identity text after these initialization
   instructions are no longer needed.
4. Remove template-construction records that do not belong to the new project,
   particularly docs/PORTFOLIO_AUDIT.md and docs/IMPLEMENTATION_PLAN.md.
5. Retain only profile-relevant sample content, then add the authoritative
   exported VBA source and independent tests.
6. Review installation, security, contribution, changelog and release
   statements so they describe facts that the project can support.
7. Run the checker self-test and the repository gate:

   ~~~bash
   python3 tools/check_repo.py --root . --self-test
   python3 tools/check_repo.py --root . \
     --output test-results/static-checks.json \
     --summary test-results/static-checks.md
   ~~~

8. Run the label synchronization workflow, configure repository metadata and
   rulesets, and select the required checks. GitHub template generation does not
   inherit those repository settings.

Do not suppress a failing rule merely to make initialization green. Correct the
repository or update the versioned profile configuration with a documented,
reviewable reason.

## Repository map

| Path | Responsibility |
| --- | --- |
| [src/](src/) | Authoritative exported production VBA |
| [tests/](tests/) | Regression modules, fixtures and test documentation |
| [examples/](examples/) | Minimal supported examples or demonstrations |
| [assets/](assets/) | Documentation images and other non-source assets |
| [docs/](docs/) | Architecture, API, compatibility and durable technical contracts |
| [tools/](tools/) | Deterministic repository, evidence and packaging utilities |
| [.github/](.github/) | Workflows, contribution templates and declarative repository policy |

See [REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md) for directory
ownership, VBA separation and legitimate profile-specific alternatives.

## Source and test contract

- Exported .bas, .cls and .frm files are the reviewable source of truth.
- Production components belong under src/; regression components belong under
  tests/.
- Public façade, internal/core, UI and test roles are declared explicitly in
  .github/repository-profile.json.
- Exported component names must match their filenames and remain unique
  case-insensitively.
- VBA source uses Windows-1252-compatible text, CRLF line endings and
  Option Explicit.
- UserForm .frx companions are tracked only when required by an authoritative
  .frm export.
- Generated workbooks and add-ins remain untracked unless the profile permits an
  exact path and the release process defines provenance.

## Quality boundaries

The [static-check workflow](.github/workflows/static-checks.yml)
checks required structure, placeholder and donor leakage, local Markdown links,
structured files, dotfile policy, line endings, forbidden artifacts, immutable
workflow actions, version/changelog consistency and exported VBA structure.

Passing that gate proves only the repository evidence available at the tested
commit. It does not prove that VBA compiles, Excel behavior is correct, both
Office bitness branches execute, numerical tolerances hold, UI state is safely
restored, or a release package matches source. Record those results separately
and bind them to the exact commit or artifact tested.

## Documentation

- [Installation](INSTALLATION.md) — supported hosts, import order, deployment
  and removal
- [Contributing](CONTRIBUTING.md) — branch, source, test and review expectations
- [Changelog](CHANGELOG.md) — unreleased and published behavior
- [Security](SECURITY.md) — supported versions and private reporting
- [Releasing](RELEASING.md) — versioning, evidence, provenance and recovery
- [Code of Conduct](CODE_OF_CONDUCT.md) — participation expectations

## License

This template uses the [MIT License](LICENSE). Confirm that the selected license
is appropriate for the generated project before publication and update all
project statements consistently if it changes.
