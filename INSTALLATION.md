<div align="center">

# 📦 Installation and Upgrade Guide

### Install, validate, upgrade, troubleshoot, and remove {{PROJECT_NAME}}

[![Deployment](https://img.shields.io/badge/Deployment-Source--first-0969da?style=flat-square)](#deployment-model)
[![Validation](https://img.shields.io/badge/Validation-Required-d97706?style=flat-square)](#validation)
[![Security](https://img.shields.io/badge/Security-Review_before_enabling-d73a49?style=flat-square)](SECURITY.md)
[![Version](https://img.shields.io/badge/Version-VERSION_file-6f42c1?style=flat-square)](VERSION)
[![License](https://img.shields.io/badge/License-MIT-217346?style=flat-square)](LICENSE)

<br>

**Back up · Import one coherent version · Compile · Validate · Preserve caller state**

</div>

---

This guide covers installation, validation, upgrade, recovery, and removal of
**{{PROJECT_NAME}}**.

> [!IMPORTANT]
> VBA source can execute with the user's Office permissions. Review the exact
> source or use a trusted tagged release, follow the organization's macro
> security policy, and never enable macros in an untrusted workbook.

---

## 🧭 Support baseline

| Item | Requirement |
|---|---|
| Host | State every supported and tested Excel/Office version |
| Office bitness | State each supported and tested bitness |
| Version identity | Root `VERSION` file and the selected tag/commit |
| Source policy | Exported repository source is authoritative |
| Licence | MIT |
| Current deployment status | Initialized {{PROFILE_NAME}} profile; document the actual package model before release |

Compatibility claims apply only to environments actually certified for the
selected release. Read [README.md](README.md), [CHANGELOG.md](CHANGELOG.md), and
the release notes before installation.

<a id="deployment-model"></a>

## 🎯 Deployment model

Choose and document the supported deployment models before publishing: embedded source, add-in, workbook, or another explicitly supported package.

Choose one supported model and keep its source identity explicit:

| Model | Use when | Trust boundary |
|---|---|---|
| Embedded source | The component must travel with a workbook or add-in | Destination project contains the reviewed source |
| Tagged source | You build or integrate the component yourself | Tag/commit and exported files define identity |
| Published binary | The project explicitly ships a workbook/add-in asset | Hash, tag binding, and package smoke evidence are required |
| Development source | Focused testing or contribution work | Not a supported release unless the project says otherwise |

Do not combine files from different tags, commits, release assets, local exports,
or copied workbooks.

---

## 📂 Production source package

| Order | Repository source | VBE component | Responsibility |
|---:|---|---|---|
| 1 | `src/core/ProjectCore.bas` | `ProjectCore` | Internal, host-independent implementation; `Option Private Module` |
| 2 | `src/modules/ProjectFacade.bas` | `ProjectFacade` | Supported public façade and stable error contract |

Optional material is not part of the normal runtime unless stated otherwise:

| Source | Purpose |
|---|---|
| `tests/modules/ProjectTests.bas` | Development-only regression harness; never part of the production package |
| `examples/modules/ProjectExample.bas` | Optional minimal consumer example; never required by production source |

> [!CAUTION]
> A `.frm` and its `.frx` companion are one logical component. Keep them in
> the same directory during import, never import the `.frx` separately, and
> never process it as text.

---

## 🚀 Fresh installation

1. Back up the destination host and user data.
2. Import every required source component in the documented dependency order.
3. Configure only the documented references, callbacks, and workbook lifecycle code.
4. Compile the complete VBA project.
5. Run the documented consumer smoke and cleanup tests.

### VBE import procedure

1. Open the destination workbook or add-in and press `Alt+F11`.
2. Select the intended project in Project Explorer.
3. Use **File → Import File…** for exported modules, classes, and forms.
4. Confirm component names match the repository source.
5. Run **Debug → Compile VBAProject**.
6. Save in a macro-capable format such as `.xlsm`, `.xlsb`, or `.xlam`
   when the project requires executable VBA.
7. Close and reopen the host before the clean-session smoke test.

Do not paste source into arbitrarily named modules when an exported component is
available. VBE attributes, component identity, form resources, and line endings
are part of a reproducible source installation.

---

<a id="validation"></a>

## ✅ Validation

A successful import is not sufficient evidence that the installation is correct.

- Run the documented startup or public-API smoke test.
- Exercise one expected-error path and verify cleanup.
- Run every applicable platform, bitness, UI, lifecycle, or numerical check.

For the neutral starter, import `ProjectTests` after the two production modules,
run **Debug → Compile VBAProject**, and execute
`ProjectTests.RunProjectTests`. A passing run ends with:

~~~text
RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS
~~~

The harness reports the Excel host, version, operating system, Office bitness,
VBA generation, named cases, and cleanup detail immediately before the verdict.
Run `ProjectExample.RunProjectExample` separately for the consumer smoke; it
prints `ProjectRatio(12, 4) = 3` without reading or changing Excel state.

### Minimum installation evidence

~~~text
Source tag or full commit SHA:
VERSION:
Files imported:
Excel version/build:
Office bitness:
Operating system:
Compile:
Consumer smoke:
Regression/certification:
Cleanup:
Skipped or unverified:
~~~

Treat a skipped, incomplete, cleanup-failed, or wrong-environment run as not
certified. Static checks and source review do not replace execution in Excel.

---

## ⬆️ Upgrade

Before upgrading:

1. read the complete version-to-version changelog;
2. back up the host and export any local modifications;
3. stop or clean up active component state;
4. identify every required production component;
5. decide whether stored configuration or generated assets are compatible.

- Replace the complete production package from one release; never mix component versions.
- Review the changelog for behavior, defaults, errors, migration, and known limitations.
- Back up the host and preserve user configuration only when the project documents it as compatible.

After replacement, compile and repeat the full installation validation. Do not
claim an upgrade is non-breaking solely because VBA signatures compile.

### Local modifications

A locally modified copy is a fork. Diff it against the old and new exported
source, reapply changes deliberately, and retest. Do not overwrite it and assume
the local behavior survived.

---

## 🧯 Troubleshooting

| Symptom | Check |
|---|---|
| Compile error or missing procedure | Confirm every required component was imported from one version and optional dependencies are present. |
| Ambiguous name | Remove duplicate/legacy modules; do not paste new source beside old components. |
| Form missing controls or corrupt UI | Re-import the `.frm` with its exact adjacent `.frx`. |
| Behavior differs by workbook | Check caller, active-object, settings namespace, references, locale, and date-system assumptions. |
| 32/64-bit failure | Confirm the tested Office bitness and conditional WinAPI declarations. |
| Excel left altered after failure | Run the documented recovery/cleanup path; do not blindly force global state. |
| Security warning | Verify source origin, signature/hash where provided, trusted location policy, and macro settings. |
| Output differs from reference | Confirm exact version, inputs, parameterization, tolerance, environment, and reference independence. |

If recovery is uncertain, save user data separately, close Excel, reopen a clean
session, and reproduce with a minimal sanitized workbook before changing code.

Report suspected vulnerabilities privately under [SECURITY.md](SECURITY.md).

---

## 🗑️ Removal

1. Run the documented project-specific cleanup or restore procedure.
2. Remove all production components and optional integrations.
3. Compile the remaining VBA project and reopen the host for a clean-session check.

Removing files does not automatically remove workbook formulas, Ribbon XML,
registry settings, trusted-location configuration, cached add-ins, shortcuts,
scheduled callbacks, or other integrations. Remove only state the component
owns and document anything intentionally retained.

---

## 🔐 Security and privacy

- Obtain source and assets from the official repository or a verified release.
- Compare the selected tag, `VERSION`, release notes, and any published hash.
- Review VBA before enabling macros.
- Do not test with client, personal, regulated, or confidential workbooks.
- Inspect example and release workbooks for links, connections, names,
  properties, hidden content, and embedded code.
- Follow organizational macro, add-in, trusted-location, and signing policy.
- Report vulnerabilities through [SECURITY.md](SECURITY.md), not publicly.

---

## 📚 Related documentation

- [README.md](README.md) — capabilities, requirements, and public API
- [CHANGELOG.md](CHANGELOG.md) — version history and compatibility
- [CONTRIBUTING.md](CONTRIBUTING.md) — source and validation standards
- [RELEASING.md](RELEASING.md) — maintainer release and provenance procedure
- [SECURITY.md](SECURITY.md) — private vulnerability reporting
- [LICENSE](LICENSE) — MIT licence terms

---

### Installation principle

> Install one identifiable source version, compile it, exercise its real host
> behavior, and keep evidence of what was—and was not—validated.
