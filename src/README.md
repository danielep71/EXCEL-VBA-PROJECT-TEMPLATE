# 🧩 Production Source

`src/` contains the authoritative production VBA source. A user must be able to reconstruct the supported workbook or add-in from this directory and the instructions in the root documentation.

## Canonical separation

Create only the subdirectories the project actually needs:

| Location | Contents | Excludes |
| --- | --- | --- |
| `src/modules/` | Public standard modules: supported procedures, worksheet functions, constants, and thin facades | Test harnesses and private implementation engines |
| `src/core/` | Internal standard modules: parsing, calculation, validation, and other implementation details | Supported public entry points |
| `src/classes/` | Production class modules, including state managers, event sinks, and UI hooks | Test doubles used only by the harness |
| `src/forms/` | Production UserForms; each `.frm` stays beside its required `.frx` | Screenshots and distributable workbooks |

A small project may keep production components directly in `src/` when further subdivision would add no clarity. If it does, document each component's role in `INSTALLATION.md`.

## Neutral starter

| Import order | Path | Component | Role |
| ---: | --- | --- | --- |
| 1 | `core/ProjectCore.bas` | `ProjectCore` | Internal implementation guarded by `Option Private Module` |
| 2 | `modules/ProjectFacade.bas` | `ProjectFacade` | Supported public façade recorded in `docs/PUBLIC_API.txt` |

The starter implements one stateless ratio operation only to prove the
façade/core, error, import, and test contracts. Replace it with real project
behavior before release, or document why the sample remains supported.

<!-- template:profile:library:start -->
## Library reference

| Import order | Path | Component | Role |
| ---: | --- | --- | --- |
| 3 | `core/TextCore.bas` | `TextCore` | Internal text and path implementation guarded by `Option Private Module` |
| 4 | `modules/TextFacade.bas` | `TextFacade` | Supported text and path façade recorded in `docs/PUBLIC_API.txt` |

The reference domain is deliberately generic: padding, whitespace
normalization, comparison, counting, splitting and joining text, and composing
and splitting paths as strings. Every Excel VBA library needs some of these,
none of them carries business meaning, and path rules genuinely differ between
Windows and macOS. Worksheet data enters only as `Range.Value` arrays; the
façade never reads ambient Excel state or the file system.
<!-- template:profile:library:end -->

<!-- template:profile:ui-component:start -->
## UI-component reference

| Import order | Path | Component | Role |
| ---: | --- | --- | --- |
| 3 | `core/ProgressCore.bas` | `ProgressCore` | Session state, snapshot/restore and the native timer, guarded by `Option Private Module` |
| 4 | `modules/ProgressFacade.bas` | `ProgressFacade` | Supported progress-session façade recorded in `docs/PUBLIC_API.txt` |

One session at a time owns these Application properties and restores each at
`ProgressEnd`; a second `ProgressBegin` is refused.

| Property | During the session | Restored to |
| --- | --- | --- |
| `StatusBar` | Progress text | The caller's text, or `False` when Excel owned the bar |
| `DisplayStatusBar` | `True` | The caller's value |
| `Cursor` (Windows) | `xlWait` | The caller's value |
| `ScreenUpdating` | `False`, unless `ProgressKeepScreenUpdating` | The caller's value |
| `EnableCancelKey` | `xlErrorHandler`: Esc raises error 18 in the caller | The caller's value |

`Calculation`, `EnableEvents`, `DisplayAlerts`, the selection and every
workbook and worksheet stay caller-owned and are never changed.
`ProgressRecover` ends a live session, or restores Excel defaults when a VBA
reset has discarded the snapshot.

Elapsed time uses the Windows performance counter, declared for VBA7 with
`PtrSafe` and for legacy VBA without it, because VBA's `Timer` wraps at
midnight; macOS uses `Timer` with the wrap corrected. The interface is status
text only: no layout or pixel geometry, so there is no DPI or scaling
assumption. Esc is the keyboard cancel. Excel exposes the status bar to
assistive technology; no further accessibility is claimed. The cursor is left
unchanged on macOS.
<!-- template:profile:ui-component:end -->

## Rules

- Preserve exported VBE component names, headers, and text encoding.
- Keep public facades thin and move reusable implementation logic into `core/`.
- Mark every production class as public-surface or internal in its header or architecture documentation.
- Keep a UserForm's `.frm` and `.frx` together and import only the `.frm` through the VBE.
- Do not place tests, examples, release binaries, generated evidence, or local workbooks here.
- Document the exact production manifest and import order in `INSTALLATION.md`.

Delete this README only if real source files and equivalent project documentation make the directory's purpose equally explicit.
