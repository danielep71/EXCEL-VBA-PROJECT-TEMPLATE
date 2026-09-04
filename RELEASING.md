# 🚀 {{PROJECT_NAME}} Release Guide

[![Release model: exact source](https://img.shields.io/badge/release-exact%20source-0969da)](#release-invariants)
[![SemVer contract](https://img.shields.io/badge/versioning-SemVer-3f4551)](docs/RELEASE_SEMANTICS.md)
[![Evidence contract](https://img.shields.io/badge/evidence-required-success)](docs/RELEASE_EVIDENCE.md)
[![Security policy](https://img.shields.io/badge/security-private-d73a49)](SECURITY.md)

This document is authoritative for the **maintainer release sequence**. Strict
version/changelog semantics are owned by
[`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md); external evidence JSON,
profile evidence and asset-manifest schemas are owned by
[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

> [!IMPORTANT]
> Generated-project releases must be fully initialized. The canonical template's
> own release is the documented exception: it preserves registered template
> tokens and uses the `template` release-evidence profile.

## 🧭 Release identity

| Property | Authority |
| --- | --- |
| Project/profile | {{PROJECT_NAME}} / {{PROFILE_NAME}} |
| Current version | [`VERSION`](VERSION) |
| User-visible history | [`CHANGELOG.md`](CHANGELOG.md) |
| SemVer/changelog policy | [`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md) |
| Installation contract | [`INSTALLATION.md`](INSTALLATION.md) |
| Evidence schema | [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) |
| Vulnerability handling | [`SECURITY.md`](SECURITY.md) |

<a id="release-invariants"></a>

## 🔒 Release invariants

A release is valid only when:

1. one exact candidate SHA is frozen and reviewable;
2. version/changelog/tag semantics pass the strict release-semantic gate;
3. repository/static checks pass on that candidate;
4. VBA compile and applicable regression/specialist checks pass on that candidate;
5. every distributed artifact is derived from and tested against that candidate;
6. external evidence and optional asset hashes bind to the candidate;
7. the annotated lower-case `v*` tag targets the certified commit; and
8. post-publication retrieval/installation checks pass.

If source changes after certification, the affected evidence is stale and must be
rerun. Never compensate by manually editing an already-tested artifact.

## 1. Freeze and identify the candidate

Start from the repository's protected release path, freeze scope, and record the
exact base/candidate revisions.

```bash
git fetch --tags --prune
git rev-parse HEAD
git status --short
git diff --stat <previous-tag>...HEAD
```

A dirty tree, unexplained generated file or unreviewed binary delta is blocking.

## 2. Synchronize version and user-visible change surfaces

Update the applicable release surfaces in one reviewable change:

- `VERSION`;
- the dated `CHANGELOG.md` release section and comparison links;
- user-facing documentation/examples affected by the release; and
- package metadata where the project actually has one.

Do not duplicate SemVer/order/link rules here. Run the authoritative semantic
contract:

```bash
python3 tools/check_release_semantics.py --root . --self-test
python3 tools/check_release_semantics.py --root .
```

Historical changelog sections and immutable evidence remain historical.

## 3. Verify documentation and installation

From a clean environment:

- follow [`INSTALLATION.md`](INSTALLATION.md);
- verify source paths, component names, prerequisites and supported upgrade path;
- verify README examples and the supported public surface;
- confirm the security and license links; and
- remove stale compatibility or evidence claims.

For a generated-project release, verify no unresolved template state remains.
For the canonical template's own release, preserve registered template state and
use the `template` evidence profile.

## 4. Run repository and release gates

At minimum:

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root .
python3 tools/check_release.py --root . --self-test
```

Run every project-specific numerical, UI, lifecycle, performance or packaging
gate as well. A stronger specialist gate is additive; the generic repository
gate never replaces it.

For changes to checker behavior, also run the checker-development and semantic
policy-coverage contracts documented in
[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md).

## 5. Certify in Excel

Use the exact candidate source in each advertised Excel environment:

1. import only candidate-controlled exports;
2. run **Debug → Compile VBAProject**;
3. execute the documented regression entry point;
4. run applicable UI/lifecycle/platform/manual checks; and
5. record environment, counts, failures, completeness and cleanup.

The neutral starter baseline is `ProjectTests.RunProjectTests`; until replaced by
the generated project's own contract it reports four cases, six assertions,
zero failures, complete execution and passing cleanup.

Source inspection is not Excel execution. If code changes, recertify.

## 6. Build and test release artifacts

Source-only libraries do not need an artificial binary asset. When the project
ships a workbook/add-in/package:

1. build from a clean location using only candidate-controlled inputs;
2. preserve required binary companions such as `.frx` files;
3. exclude development-only material unless promised;
4. reopen and smoke/regression-test the packaged artifact;
5. record filename, size and SHA-256; and
6. never edit the artifact after hashing.

The exact manifest format and profile-specific evidence requirements are defined
only in [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

## 7. Create and validate external evidence

Keep candidate-binding release evidence outside the candidate tree to avoid
self-referential commit hashes. Prepare the evidence JSON and, when applicable,
the sorted asset manifest according to
[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

The evidence must identify the exact candidate, executed checks, environment and
material limitations. A release note summarizes evidence; it does not replace it.

## 8. Review and merge the release candidate

The release review should make these facts easy to verify:

- target version and previous tag;
- candidate SHA and final diff;
- semantic/repository/release gate results;
- Excel and specialist evidence;
- artifact manifest/hashes when applicable;
- compatibility/migration/security notes; and
- remaining limitations.

If the merge changes source identity, certify the merged commit before tagging.

## 9. Create the protected annotated tag

Tag only the certified commit. Run the release-integrity checker before and after
creating the local annotated tag:

```bash
git switch main
git pull --ff-only
candidate_sha="$(git rev-parse HEAD)"
release_version="$(tr -d '\r\n' < VERSION)"
release_tag="v${release_version}"

python3 tools/check_release.py \
  --root . \
  --tag "$release_tag" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --output test-results/release-integrity.json \
  --summary test-results/release-integrity.md

git tag -a "$release_tag" -m "{{PROJECT_NAME}} ${release_version}"

python3 tools/check_release.py \
  --root . \
  --tag "$release_tag" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --require-tag-ref

git push origin "$release_tag"
```

Add `--asset-manifest ../release-assets.sha256` when the release distributes
binary assets. Do not push the tag if either check fails. Never move or recreate
a public tag to hide an error.

## 10. Publish the GitHub Release

Create the release from the protected annotated tag. Include:

- user-facing summary/highlights;
- upgrade or migration notes;
- supported platform statement;
- known limitations;
- installation link;
- artifact/hash table when applicable;
- changelog comparison link; and
- security-reporting link.

Upload the already-tested, already-hashed artifacts. Do not rebuild between
certification/tagging and publication.

## 11. Verify after publication

- [ ] Tag resolves to the certified SHA.
- [ ] `VERSION` and changelog agree with the tag.
- [ ] Published assets download and hashes match.
- [ ] Installation and documentation links work.
- [ ] Packaged artifact, when present, passes its published smoke test.
- [ ] Source archive contains the expected release tree.
- [ ] Default branch is ready for the next Unreleased cycle.

Do not announce broad availability until these checks pass.

## 🧯 Recovery

Before publication, repair the candidate and rerun every affected gate. After a
public release, never silently replace assets or move the tag: document the
problem and publish a corrected patch release. Vulnerability handling follows
[`SECURITY.md`](SECURITY.md).

## 📚 Related authorities

- [`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md) — exact version/changelog semantics
- [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) — evidence and asset-manifest schema
- [`INSTALLATION.md`](INSTALLATION.md) — clean install/upgrade validation
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — change/review workflow
- [`SECURITY.md`](SECURITY.md) — vulnerability handling
- [`docs/README.md`](docs/README.md) — complete documentation authority map

---

**Release principle:** certify one exact source revision, derive artifacts from it
once, and publish only evidence-backed output bound to that revision.
