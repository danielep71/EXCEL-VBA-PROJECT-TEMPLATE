# Repository-local GitHub Action validation

`check_local_actions.py` complements the pinned authoritative workflow validator by enforcing repository-boundary and tracked-state rules that schema validation alone cannot prove.

## Authority split

- **`check_local_actions.py`** owns local `uses: ./...` containment, tracked metadata, metadata-file uniqueness, supported local runtime shapes, and tracked entrypoints.
- **actionlint 1.7.12** remains authoritative for GitHub Actions YAML/schema semantics.
- **`check_repo.py`** continues to require immutable full-SHA pins and audited version comments for external actions.

No one gate substitutes for another; the hosted terminal verdict requires all of them.

## Local action contract

A local action reference must:

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

## Commands

```bash
python3 tools/check_local_actions.py --root . --self-test
python3 tools/check_local_actions.py \
  --root . \
  --output test-results/local-actions.json \
  --summary test-results/local-actions.md
```

The self-test covers valid tracked composite and Node actions plus every blocking boundary above. External action references are intentionally ignored by this tool because their immutable-pin contract is enforced separately.
