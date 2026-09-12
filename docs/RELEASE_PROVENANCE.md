# 🔐 Release Provenance

[![Contract: 1.2.0](https://img.shields.io/badge/contract-1.2.0-217346)](TEMPLATE_CONTRACT.md)
[![Digest: SHA-256](https://img.shields.io/badge/digest-SHA--256-1D76DB)](#-minimum-levels)
[![Signatures: policy-selected](https://img.shields.io/badge/signatures-policy--selected-6f42c1)](#-signature-policy)

This is the authority for build records and signature verification under template
contract 1.2.0. The existing [release evidence contract](RELEASE_EVIDENCE.md)
continues to own runtime checks and the SHA-256 asset manifest. The release gate
calls `tools/release_provenance.py`; there is no second publication verdict.

## 📋 Minimum Levels

| Profile | Source-only distribution | Generated binary assets |
| --- | --- | --- |
| `library` | Base release evidence; empty `assets`; no manifest or build record required | Disallowed by the base release policy |
| `template` | Base release evidence, including pilots and governance | Disallowed by the base release policy |
| `ui-component` | Base release evidence; no artificial workbook required | Base evidence, SHA-256 manifest and build record |
| `application` | Base release evidence; no artificial workbook required | Base evidence, SHA-256 manifest and build record |

Selecting SSH verification raises every distribution's minimum to a signed build
record, including source-only releases. The selection is explicit in
`.github/release-policy.json`; omission, an unsupported value, or disagreement
with `.github/release-provenance.json` fails closed. Contracts 1.0.0 and 1.1.0
retain their previous evidence rules; advanced inputs require 1.2.0. Unsupported
future contracts fail instead of inheriting today's policy.

`dist/` is the complete payload boundary. Its regular files must exactly equal
the evidence asset list and digests. Missing, additional, changed, duplicate or
symlinked payloads fail. Keep reports, provenance and signatures outside `dist/`.
For source-only releases, `dist/` must be empty or absent.

## 🧭 Trust Policy

Commit `.github/release-provenance.json` before freezing the candidate. The
default identifies `.github/workflows/static-checks.yml` in the candidate's own
repository. `@repository` and `@candidate` resolve to the committed repository
identity and exact release SHA. They are policy references, not initialization
placeholders.

Set `workflow.path` to the workflow responsible for the retained release
validation or build record. For a reusable provider, set `workflow.repository`
and `workflow.sha` to its actual repository and full immutable commit. Do not
use a branch, tag, or a workflow that did not participate. Record the actual
invocation's run ID and attempt. A local default workflow must exist in the
candidate; a remote provider's content and run must be reviewed separately.

The gate reads policy and signing keys from candidate Git objects. Editing the
working copy or an external record cannot disable required verification. The
reviewed candidate SHA and trusted verifier installation are the trust roots:
review changes to that policy as carefully as changes to the workflow itself.

## 🔏 Signature Policy

`.github/release-policy.json` is the selector for the provenance-record signature
requirement. Its top-level `provenance_signature_mode` field is mandatory and
supports exactly:

- `"none"` — detached provenance signing is disabled; a supplied provenance
  signature is rejected rather than silently ignored;
- `"ssh"` — a provenance record and detached OpenSSH signature are mandatory,
  including for source-only releases.

`.github/release-provenance.json` supplies the candidate-bound trust details and
must declare the **same** mode in its `signature` object. The duplication is
intentional: the release policy selects the required assurance level, while the
provenance trust policy supplies the verification configuration. Neither file may
silently raise or lower the other. A missing selector, unsupported selector, or
mode mismatch is a blocking release finding.

The default canonical/template baseline selects `"none"` explicitly in both
files. This is an explicit unsigned provenance policy, not permission inferred
from an absent setting. Git-tag signing is a separate control and does not satisfy
or replace this provenance-record policy.

## 🧾 Build Record

After producing and testing assets, finalize the external evidence JSON and
manifest. Hash their exact bytes with SHA-256, then create an external UTF-8
`release-provenance.json`. Replace every example value with observed facts:

```json
{
  "schema_version": 1,
  "repository": "owner/project",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "template_contract": {"version": "1.2.0", "source": "owner/template"},
  "profile": "application",
  "tag": "v1.0.0",
  "distribution": "binary",
  "digest_algorithm": "sha256",
  "evidence_sha256": "REPLACE_WITH_EXACT_EVIDENCE_DIGEST",
  "manifest_sha256": "REPLACE_WITH_EXACT_MANIFEST_DIGEST",
  "assets": [{"path": "dist/project.xlsm", "sha256": "REPLACE_WITH_ASSET_DIGEST"}],
  "workflow": {
    "repository": "owner/project",
    "path": ".github/workflows/static-checks.yml",
    "sha": "0123456789abcdef0123456789abcdef01234567",
    "run_id": 123,
    "run_attempt": 1
  },
  "environment": {
    "os": "Windows 10",
    "architecture": "x64",
    "host": "Microsoft Excel",
    "host_version": "16.0",
    "office_bitness": "64",
    "runtime": "VBA7+",
    "builder": "Controlled workstation identifier",
    "procedure": "Retained build log identifier and exact import/package procedure"
  }
}
```

The shape is exact; duplicate JSON keys fail. Copy `template_contract` from the
candidate. List assets in path order. For optional source-only records use
`assets: []`, `distribution: "source-only"`, and `manifest_sha256: null` if the
manifest is omitted. All environment fields must be nonempty; use an explicit
`not applicable: source-only` explanation for unused Office/build fields.
Keep runtime test environments in the separate base evidence even if identical.

## ✍️ Optional Signatures

This implementation supports detached OpenSSH signatures, without a signing
service dependency. It does not claim SLSA or verify vendor attestations.
Use an approved release signing key kept outside the repository. Commit its
public key in an allowed-signers file, for example `.github/release-signers`:

```text
release@example.org ssh-ed25519 REPLACE_WITH_APPROVED_PUBLIC_KEY
```

To enable signing, change the release selector before freezing the candidate:

```json
{"provenance_signature_mode": "ssh"}
```

and configure the matching provenance trust policy:

```json
{"mode": "ssh", "principal": "release@example.org", "allowed_signers": ".github/release-signers"}
```

Both changes belong in the reviewed candidate. Changing only one file is a
policy mismatch and fails before signature verification.

Sign the finalized record with the dedicated namespace:

```bash
ssh-keygen -Y sign -f /secure/path/release-key -n excel-vba-release ../release-provenance.json
```

The gate invokes `ssh-keygen -Y verify` with the committed allowed signers,
configured principal, namespace and exact record bytes. See the
[OpenSSH manual](https://man.openbsd.org/ssh-keygen.1) for key and allowed-signers
formats. An absent tool, unsupported operation, timeout, missing signature,
wrong key/namespace, changed record, or policy mismatch blocks publication.
Supplying a signature while both policies say `none` also fails: no signature is
silently left unchecked.

## ✅ Verification and Retention

1. Check out the reviewed candidate; retain the exact source SHA and both release
   and provenance trust policies.
2. Stage only the approved downloadable payloads in `dist/`. Build/test them
   using the recorded environment and keep the actual logs.
3. Finalize base evidence, manifest and build record; sign last if enabled.
4. Run the integrated gate, adding `--require-tag-ref` after the annotated tag
   exists. Omit `--provenance-signature` when signing is disabled:

```bash
python3 tools/check_release.py --root . --tag v1.0.0 \
  --candidate-sha FULL_CANDIDATE_SHA \
  --evidence ../release-evidence.json \
  --asset-manifest ../release-assets.sha256 \
  --provenance ../release-provenance.json \
  --provenance-signature ../release-provenance.json.sig
```

5. Retain the checked evidence, manifest, record, signature if enabled, build
   logs, gate output and workflow run/attempt alongside the release. Recheck
   downloaded payloads using the same candidate and those exact bytes.

Checksums prove byte identity. A verified signature authenticates the approved
signer's assertions. Neither independently proves that Excel imported the
recorded source, that an environment description is truthful, or that a named
workflow ran successfully. Review its exact source, logs and outcome separately.
The default static workflow supplies validation identity, not an automated
Excel builder. Optional Windows/Excel automation is a separate contract.

## ↩️ Rollback and Key Changes

If validation fails before publication, stop, correct the candidate or rebuild
assets, and regenerate every affected digest and signature. After publication,
never replace assets silently or move the tag. Restore a previously certified
version after verifying its retained bundle; publish corrections under a new
version with fresh evidence. Commit and review signer rotation or policy
changes before a new candidate. If a key is compromised, remove it from the
current trust policy and document which historical releases require independent
revalidation: an old candidate's embedded key is historical trust, not proof
that the key remains approved today.

## 🧪 Validation Scope

`python3 tools/test_release_provenance.py -v` exercises synthetic candidates,
payload tampering, explicit unsigned policy, missing/unsupported selectors,
both policy-mismatch directions, and real ephemeral SSH signatures. It runs with
the existing release self-test in repository CI. It does not build or execute
Office files, request signing credentials, or publish a release.
