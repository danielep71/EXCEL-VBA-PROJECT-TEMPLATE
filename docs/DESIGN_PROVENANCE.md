# 🧭 Design Provenance

Some solutions in this template are independent reimplementations of designs
first written in other repositories of the same maintainer. No code is copied:
every file listed here is written in this repository's own terms and stays
under its MIT licence. This note records where each design came from, at which
revision, so a reviewer can compare the two.

| File in this repository | Design origin | Origin revision | Issue |
| --- | --- | --- | --- |
| `.github/scripts/labels-sync.mjs` (prune guard: `labelsInUse`, `guardPrune`, `executePlan`) and `.github/workflows/labels-sync.yml` (`pull-requests: read`) | `danielep71/MyTemplate`: `.github/scripts/labels-sync.mjs` and `.github/workflows/labels-sync.yml` (MPL-2.0) | `f1d0d17ecede525941cfedf07f0cf737bf1346ef` | #157 |
| `.github/workflows/labels-sync.yml` (reconcile job only on the default branch) and `tools/test_workflow_validation.py` (`label_reconcile_contract`) | `danielep71/MyTemplate`: `.github/workflows/labels-sync.yml` (MPL-2.0) | `f1d0d17ecede525941cfedf07f0cf737bf1346ef` | #158 |
| `.github/scripts/labels-sync.mjs` and `.github/scripts/labels-drift.mjs` (`inlineCode`, `cellText` in the summaries) | `danielep71/MyTemplate`: `.github/scripts/labels-sync.mjs` and `.github/scripts/labels-drift.mjs` (MPL-2.0) | `f1d0d17ecede525941cfedf07f0cf737bf1346ef` | #159 |
| `.github/scripts/labels-sync.mjs` and `.github/scripts/labels-drift.mjs` (`apiUrl` honours `GITHUB_API_URL`) | `danielep71/MyTemplate`: `.github/scripts/labels-sync.mjs` and `.github/scripts/labels-drift.mjs` (MPL-2.0) | `f1d0d17ecede525941cfedf07f0cf737bf1346ef` | #160 |

This note belongs to the template and is removed from generated projects.
