#!/usr/bin/env node

import assert from "node:assert/strict";
import { appendFile, mkdir, readFile, writeFile } from "node:fs/promises";
import process from "node:process";
import {
  planChanges,
  resolveDesired,
  resolvePolicySelection,
  validateManifest
} from "./labels-sync.mjs";

const TOOL_NAME = "Repository label drift";
const API_VERSION = "2022-11-28";
const DEFAULT_MANIFEST = ".github/labels.json";
const DEFAULT_POLICY = ".github/repository-profile.json";

function compareNames(left, right) {
  const leftKey = left.toLowerCase();
  const rightKey = right.toLowerCase();
  if (leftKey < rightKey) return -1;
  if (leftKey > rightKey) return 1;
  return left < right ? -1 : left > right ? 1 : 0;
}

function normalizedLiveLabel(label) {
  return {
    name: String(label.name ?? ""),
    color: String(label.color ?? "").toUpperCase(),
    description: label.description == null ? "" : String(label.description)
  };
}

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    policy: DEFAULT_POLICY,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    output: null,
    summary: null,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--self-test") {
      parsed.selfTest = true;
      continue;
    }
    if (["--manifest", "--policy", "--live", "--repository", "--output", "--summary"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument.slice(2).replace(/-([a-z])/g, (_, letter) => letter.toUpperCase())] = value;
      index += 1;
      continue;
    }
    throw new Error(`Unknown argument: ${argument}`);
  }
  return parsed;
}

async function loadJson(path) {
  let text;
  try {
    text = await readFile(path, "utf8");
  } catch (error) {
    throw new Error(`Cannot read ${path}: ${error.message}`);
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Invalid JSON in ${path}: ${error.message}`);
  }
}

function splitRepository(repository) {
  const parts = String(repository ?? "").split("/");
  if (parts.length !== 2 || parts.some(part => part.length === 0)) {
    throw new Error("repository must be in owner/name form");
  }
  return parts.map(encodeURIComponent);
}

function apiUrl(repository, suffix) {
  const [owner, repo] = splitRepository(repository);
  return `https://api.github.com/repos/${owner}/${repo}${suffix}`;
}

async function listLiveLabels(repository, token) {
  const labels = [];
  for (let page = 1; ; page += 1) {
    const response = await fetch(apiUrl(repository, `/labels?per_page=100&page=${page}`), {
      method: "GET",
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "X-GitHub-Api-Version": API_VERSION,
        "User-Agent": "canonical-vba-label-drift"
      }
    });
    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`GET /labels failed with HTTP ${response.status}: ${detail.slice(0, 500)}`);
    }
    const batch = await response.json();
    if (!Array.isArray(batch)) throw new Error("GitHub labels response is not an array");
    labels.push(...batch.map(normalizedLiveLabel));
    if (batch.length < 100) break;
  }
  return labels.sort((left, right) => compareNames(left.name, right.name));
}

function changeEvidence(change, observedByName) {
  const currentKey = String(change.currentName ?? change.name).toLowerCase();
  const current = observedByName.get(currentKey) ?? null;
  const desired = change.target ?? null;
  return {
    action: change.action,
    name: change.name,
    fields: change.fields ?? [],
    observed: current,
    desired
  };
}

export function buildReport({
  repository,
  manifestPath,
  policyPath,
  manifest,
  selection,
  desiredLabels,
  liveLabels
}) {
  const observed = liveLabels
    .map(normalizedLiveLabel)
    .sort((left, right) => compareNames(left.name, right.name));
  const desired = structuredClone(desiredLabels)
    .sort((left, right) => compareNames(left.name, right.name));
  const changes = planChanges(desired, observed, { prune: manifest.prune });
  const observedByName = new Map(observed.map(label => [label.name.toLowerCase(), label]));
  const evidence = changes.map(change => changeEvidence(change, observedByName));
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    evidence.filter(item => item.action === action).length
  ]));
  return {
    schema_version: 1,
    tool: TOOL_NAME,
    status: evidence.length === 0 ? "pass" : "drift",
    repository,
    manifest: manifestPath,
    policy: policyPath,
    selection: {
      profile: selection.profile,
      domains: selection.domains
    },
    prune: manifest.prune,
    counts: {
      desired: desired.length,
      observed: observed.length,
      differences: evidence.length,
      create: counts.create,
      update: counts.update,
      delete: counts.delete
    },
    desired,
    observed,
    differences: evidence
  };
}

function markdownEscape(value) {
  return String(value).replaceAll("|", "\\|").replaceAll("\n", " ");
}

function compactLabel(label) {
  if (label === null) return "—";
  return `\`${markdownEscape(label.name)}\` / \`${label.color}\` / ${markdownEscape(label.description)}`;
}

export function markdownReport(report) {
  const lines = [
    "# Repository label drift",
    "",
    `- Status: **${String(report.status).toUpperCase()}**`,
    `- Repository: \`${report.repository ?? "fixture"}\``,
    `- Manifest: \`${report.manifest}\``,
    `- Policy: \`${report.policy}\``,
    `- Profile overlay: **${report.selection.profile ?? "none"}**`,
    `- Domain overlays: **${report.selection.domains.length > 0 ? report.selection.domains.join(", ") : "none"}**`,
    `- Desired / observed: **${report.counts.desired} / ${report.counts.observed}**`,
    `- Differences: **${report.counts.differences}**`,
    `- Create / update / delete: **${report.counts.create} / ${report.counts.update} / ${report.counts.delete}**`,
    ""
  ];
  if (report.differences.length === 0) {
    lines.push("No live label drift detected.");
  } else {
    lines.push(
      "| Action | Label | Fields | Observed | Desired |",
      "| --- | --- | --- | --- | --- |"
    );
    for (const difference of report.differences) {
      lines.push(
        `| ${difference.action} | \`${markdownEscape(difference.name)}\` | `
        + `${difference.fields.length > 0 ? difference.fields.join(", ") : "—"} | `
        + `${compactLabel(difference.observed)} | ${compactLabel(difference.desired)} |`
      );
    }
  }
  return `${lines.join("\n")}\n`;
}

async function writeEvidence(path, content) {
  if (!path) return;
  const slash = path.lastIndexOf("/");
  if (slash > 0) await mkdir(path.slice(0, slash), { recursive: true });
  await writeFile(path, content, "utf8");
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) {
    await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
  }
}

async function prepareContract(manifestPath, policyPath) {
  const manifest = validateManifest(await loadJson(manifestPath));
  const selection = resolvePolicySelection(await loadJson(policyPath), manifest);
  const desiredLabels = resolveDesired(manifest, selection);
  return { manifest, selection, desiredLabels };
}

async function runSelfTest(manifestPath, policyPath) {
  const { manifest, selection, desiredLabels } = await prepareContract(manifestPath, policyPath);

  const baseline = structuredClone(desiredLabels);
  const baselineBefore = JSON.stringify(baseline);
  const cleanFirst = buildReport({
    repository: "fixture/clean",
    manifestPath,
    policyPath,
    manifest,
    selection,
    desiredLabels,
    liveLabels: baseline
  });
  const cleanSecond = buildReport({
    repository: "fixture/clean",
    manifestPath,
    policyPath,
    manifest,
    selection,
    desiredLabels,
    liveLabels: baseline
  });
  assert.equal(cleanFirst.status, "pass");
  assert.equal(cleanFirst.counts.differences, 0);
  assert.equal(JSON.stringify(cleanFirst), JSON.stringify(cleanSecond));
  assert.equal(markdownReport(cleanFirst), markdownReport(cleanSecond));
  assert.equal(JSON.stringify(baseline), baselineBefore, "clean check must not mutate live input");

  const drifted = structuredClone(desiredLabels);
  const missing = drifted.shift();
  assert.ok(missing, "fixture requires at least one desired label");
  drifted[0].color = drifted[0].color === "000000" ? "FFFFFF" : "000000";
  drifted[0].description = "simulated out-of-band change";
  drifted.push({ name: "out-of-band", color: "ABCDEF", description: "simulated extra label" });
  const driftedBefore = JSON.stringify(drifted);
  const driftFirst = buildReport({
    repository: "fixture/drift",
    manifestPath,
    policyPath,
    manifest,
    selection,
    desiredLabels,
    liveLabels: drifted
  });
  const driftSecond = buildReport({
    repository: "fixture/drift",
    manifestPath,
    policyPath,
    manifest,
    selection,
    desiredLabels,
    liveLabels: drifted
  });
  assert.equal(driftFirst.status, "drift");
  assert.deepEqual(
    driftFirst.differences.map(item => item.action),
    ["update", "create", "delete"]
  );
  assert.equal(JSON.stringify(driftFirst), JSON.stringify(driftSecond));
  assert.equal(markdownReport(driftFirst), markdownReport(driftSecond));
  assert.equal(JSON.stringify(drifted), driftedBefore, "drift check must not mutate live input");
  assert.match(markdownReport(driftFirst), /simulated out-of-band change/);
  assert.match(markdownReport(driftFirst), new RegExp(missing.name));

  process.stdout.write(
    "SELF-TEST PASS: deterministic no-drift and create/update/delete drift fixtures are read-only.\n"
  );
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest, options.policy);
    return;
  }

  const { manifest, selection, desiredLabels } = await prepareContract(
    options.manifest,
    options.policy
  );

  let liveLabels;
  if (options.live) {
    const liveDocument = await loadJson(options.live);
    liveLabels = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(liveLabels)) {
      throw new Error("live fixture must be an array or contain a labels array");
    }
  } else {
    if (!options.repository) {
      throw new Error("live check requires --repository or GITHUB_REPOSITORY");
    }
    if (!process.env.GITHUB_TOKEN) {
      throw new Error("live check requires GITHUB_TOKEN");
    }
    liveLabels = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  }

  const report = buildReport({
    repository: options.repository,
    manifestPath: options.manifest,
    policyPath: options.policy,
    manifest,
    selection,
    desiredLabels,
    liveLabels
  });
  const json = `${JSON.stringify(report, null, 2)}\n`;
  const markdown = markdownReport(report);
  await writeEvidence(options.output, json);
  await writeEvidence(options.summary, markdown);
  await publishSummary(markdown);
  process.exitCode = report.status === "pass" ? 0 : 1;
}

main().catch(error => {
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`ERROR: ${message}\n`);
  process.exitCode = 2;
});
