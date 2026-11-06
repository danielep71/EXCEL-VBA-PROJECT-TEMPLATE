# 🧩 Repository-local GitHub Action Validation

[![Scope: tracked local actions](https://img.shields.io/badge/scope-tracked%20local%20actions-217346)](#local-action-contract)

`check_local_actions.py` complements the pinned authoritative workflow validator by enforcing repository-boundary and tracked-state rules that schema validation alone cannot prove.

## Authority split

- **`check_local_actions.py`** owns local `uses: ./...` containment, tracked reusable-workflow files, tracked action metadata, metadata-file uniqueness, supported local runtime shapes, and tracked entrypoints.
- **actionlint 1.7.12** remains authoritative for GitHub Actions YAML/schema semantics.
- **`check_repo.py`** continues to require immutable full-SHA pins and audited version comments for external actions.

No one gate substitutes for another; the hosted terminal verdict requires all of them.

## Local action contract

A step-level local action reference must:

1. start with `./` and resolve inside the repository without `.` or `..` traversal segments;
2. point to an existing directory;
3. contain exactly one tracked metadata file: `action.yml` or `action.yaml`;
4. provide non-empty `name`, `description`, and `runs.using` metadata;
5. use a supported baseline runtime shape:
   - `composite`, with `runs.steps`;
   - `node20` or `node24`, with a tracked in-directory `runs.main` and any declared `pre`/`post` entrypoints; or
   - `docker`, with a tracked in-directory `Dockerfile`;
6. keep all required entrypoints inside the action directory and under Git control.

The checker fails closed on missing directories, path traversal, untracked metadata, dual `action.yml`/`action.yaml`, malformed required metadata, unsupported runtimes, missing entrypoints, untracked entrypoints, or entrypoint traversal.

## Local reusable workflow contract

A job-level `uses: ./.github/workflows/reusable.yml` calls a reusable workflow. It must name an existing tracked `.yml` or `.yaml` file directly inside `.github/workflows`, without traversal segments. A workflow file used as a step action, or an action directory used as a job workflow, fails the corresponding contract.

The gate classifies block-style YAML references by their job/step context, supports indentless step sequences, and ignores reference-like text inside block scalars. The authoritative workflow validator owns YAML/schema semantics, including reusable-workflow calling rules.

## Commands

```bash
python3 tools/check_local_actions.py --root . --self-test
python3 tools/check_local_actions.py \
  --root . \
  --output test-results/local-actions.json \
  --summary test-results/local-actions.md
```

The self-test covers valid tracked composite, Node and Docker actions (including quoted references) and these rejections: a missing path, `..` traversal, an untracked action, an untracked Dockerfile, two metadata files, an empty name or description, malformed metadata, and a missing, untracked or traversing entrypoint. Its reusable-workflow cases accept a tracked job-level workflow and reject a missing, untracked, traversing or nested one, an action directory used as a job and a workflow file used as a step. An unsupported runtime, a composite action without `runs.steps`, a Docker image other than `Dockerfile` and Node `pre`/`post` entrypoints are enforced by the gate but have no fixture yet. External action references are intentionally ignored by this tool because their immutable-pin contract is enforced separately.
