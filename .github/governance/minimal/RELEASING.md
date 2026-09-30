# 🚀 {{PROJECT_NAME}} Release Guide

[![SemVer contract](https://img.shields.io/badge/versioning-SemVer-3f4551)](docs/RELEASE_SEMANTICS.md)
[![Governance: minimal](https://img.shields.io/badge/governance-minimal-0969da)](docs/INITIALIZATION.md#governance-tiers)
[![Security policy](https://img.shields.io/badge/security-private-d73a49)](SECURITY.md)

This document is authoritative for the **maintainer release sequence** of a
project on the **minimal governance tier**. Strict version/changelog semantics
are owned by [`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md).

The minimal tier keeps source discipline, VBA structure and public-API checks,
release semantics and a verified Excel regression run. It does not include
machine-readable release evidence, build provenance, host-evidence validation,
release certification or post-release closeout. Adopt those with the documented
[upgrade to the full tier](docs/INITIALIZATION.md#governance-tiers); no new
repository is needed.

## 🔒 Release invariants

A release is valid only when:

1. one exact candidate SHA is frozen and reviewable;
2. version/changelog/tag semantics pass the strict release-semantic gate;
3. repository, VBA public-API and procedure-scoped jump checks pass on that
   candidate;
4. VBA compile and the regression suite pass on that candidate;
5. every distributed artifact is derived from and tested against that candidate;
6. the annotated lower-case `v*` tag targets the certified commit; and
7. post-publication retrieval checks pass.

If source changes after certification, the affected checks are stale and must be
rerun. Never compensate by manually editing an already-tested artifact.

## 1. Freeze and identify the candidate

Before a functional release, verify the required branch/tag protection through
live settings or read-back. Files in the repository do not prove that
server-side protection is active.

```bash
git fetch --tags --prune
git rev-parse HEAD
git status --short
git diff --stat <previous-tag>...HEAD
```

A dirty tree, unexplained generated file or unreviewed binary delta is blocking.

## 2. Synchronize version and user-visible change surfaces

Update `VERSION`, the dated `CHANGELOG.md` release section and comparison links,
and any user-facing documentation affected by the release, in one reviewable
change. Use the calendar date on which the release section is cut/frozen for the
candidate. Then run the authoritative semantic contract:

```bash
python3 tools/check_release_semantics.py --root . --self-test
python3 tools/check_release_semantics.py --root .
```

## 3. Run the repository gates

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root .
python3 tools/check_vba_public_api.py --root .
python3 tools/check_vba_jumps.py --root .
```

The hosted *Static repository checks* workflow runs the same gates, with their
self-tests, on every pull request.

## 4. Certify in Excel

Use the exact candidate source in each advertised Excel environment:

1. import only candidate-controlled exports;
2. run **Debug → Compile VBAProject**;
3. execute the documented regression entry point; and
4. record environment, counts, failures, completeness and cleanup in the release
   pull request.

The neutral starter baseline is `ProjectTests.RunProjectTests`; until replaced by
the project's own contract it reports four cases, six assertions, zero failures,
complete execution and passing cleanup. Source inspection is not Excel
execution. If code changes, recertify.

## 5. Review and merge the release candidate

Use **Squash and merge** for release pull requests into `main`. The review should
show the target version and previous tag, the candidate SHA and final diff, the
gate results, the Excel run and the remaining limitations. Required checks must
pass, and the merge must use the reviewed head SHA.

## 6. Create the release tag

Tag only the certified commit with an annotated tag:

```bash
(
git switch main || exit $?
git pull --ff-only || exit $?
candidate_sha="$(git rev-parse HEAD)" || exit $?
release_version="$(tr -d '\r\n' < VERSION)" || exit $?
release_tag="v${release_version}"
python3 tools/check_release_semantics.py --root . || exit $?
git tag -a "$release_tag" "$candidate_sha" -m "{{PROJECT_NAME}} ${release_version}" || exit $?
test "$(git rev-parse "${release_tag}^{commit}")" = "$candidate_sha" || exit $?
git push origin "refs/tags/$release_tag:refs/tags/$release_tag" || exit $?
)
```

## 7. Publish and verify

Create the GitHub Release from the annotated tag, with a curated summary,
upgrade notes, known limitations and the changelog comparison link. Then check:

- [ ] Tag resolves to the certified SHA.
- [ ] `VERSION` and changelog agree with the tag.
- [ ] Published assets, when present, download intact.
- [ ] Installation and documentation links work.
- [ ] Default branch is ready for the next Unreleased cycle.

## 🧯 Recovery

Before publication, repair the candidate and rerun every affected gate. After a
public release, never silently replace assets or move the tag: document the
problem and publish a corrected patch release. Vulnerability handling follows
[`SECURITY.md`](SECURITY.md).

---

**Release principle:** certify one exact source revision and publish only output
derived from it.
