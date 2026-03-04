#!/usr/bin/env node
/**
 * Office.js Manifest Validator
 *
 * Pass 1 — Microsoft's official validation gateway (office-addin-manifest)
 * Pass 2 — Static analysis: resid integrity, duplicate IDs, URL reachability
 *
 * Usage:
 *   node validate-manifest.mjs <manifest.xml> [--timeout=5000] [--skip-urls] [--only-local]
 */

import { createRequire } from "module";
import { readFileSync } from "fs";
import { parseArgs } from "util";
import { XMLParser } from "fast-xml-parser";
import fetch from "node-fetch";
import https from "https";
import path from "path";

const require = createRequire(import.meta.url);
const { validateManifest } = require("./node_modules/office-addin-manifest/lib/validate.js");

// ── CLI ───────────────────────────────────────────────────────────────────────
const { values, positionals } = parseArgs({
  args: process.argv.slice(2),
  options: {
    timeout:      { type: "string",  default: "5000"  },
    "skip-urls":  { type: "boolean", default: false   },
    "only-local": { type: "boolean", default: false   },
    help:         { type: "boolean", default: false   },
  },
  allowPositionals: true,
});

if (values.help || positionals.length === 0) {
  console.log(`
Usage: node validate-manifest.mjs <manifest.xml> [options]

Options:
  --timeout <ms>   HTTP request timeout per URL (default: 5000)
  --skip-urls      Skip URL reachability checks entirely
  --only-local     Only probe localhost / 127.x.x.x URLs
  --help           Show this help
`);
  process.exit(0);
}

const MANIFEST_PATH = positionals[0];
const TIMEOUT_MS    = parseInt(values["timeout"], 10);
const SKIP_URLS     = values["skip-urls"];
const ONLY_LOCAL    = values["only-local"];

// ── Colours ───────────────────────────────────────────────────────────────────
const C = {
  reset: "\x1b[0m", bold: "\x1b[1m",
  red: "\x1b[31m",  green: "\x1b[32m",
  yellow: "\x1b[33m", cyan: "\x1b[36m",
  grey: "\x1b[90m",
};
const fmt = {
  ok:   (s) => `${C.green}✔${C.reset}  ${s}`,
  warn: (s) => `${C.yellow}⚠${C.reset}  ${s}`,
  fail: (s) => `${C.red}✖${C.reset}  ${s}`,
  info: (s) => `${C.cyan}ℹ${C.reset}  ${s}`,
  head: (s) => `\n${C.bold}${C.cyan}━━ ${s} ━━${C.reset}`,
  grey: (s) => `${C.grey}${s}${C.reset}`,
};

const summary = { errors: 0, warnings: 0 };
const addError   = (msg) => { summary.errors++;   console.log(fmt.fail(msg)); };
const addWarning = (msg) => { summary.warnings++; console.log(fmt.warn(msg)); };
const addOk      = (msg) => console.log(fmt.ok(msg));
const addInfo    = (msg) => console.log(fmt.info(msg));

// ── Load & parse XML ──────────────────────────────────────────────────────────
console.log(fmt.head("Loading manifest"));

let raw;
try {
  raw = readFileSync(MANIFEST_PATH, "utf8");
  addOk(`Read ${path.resolve(MANIFEST_PATH)} (${(raw.length / 1024).toFixed(1)} KB)`);
} catch (e) {
  console.log(fmt.fail(`Cannot read file: ${e.message}`));
  process.exit(1);
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  isArray: (_name, _jpath, _isLeaf, isAttr) => !isAttr,  // wrap elements, not attributes
  removeNSPrefix: false,
});

let doc;
try {
  doc = parser.parse(raw);
  addOk("XML is well-formed");
} catch (e) {
  addError(`XML parse error: ${e.message}`);
  process.exit(1);
}

// ── Traversal helpers ─────────────────────────────────────────────────────────
function collectAttribs(obj, attrName, result = []) {
  if (!obj || typeof obj !== "object") return result;
  if (Array.isArray(obj)) { obj.forEach(i => collectAttribs(i, attrName, result)); return result; }
  for (const [k, v] of Object.entries(obj)) {
    if (k === attrName) result.push(String(v));
    else collectAttribs(v, attrName, result);
  }
  return result;
}

function collectElements(obj, localName, result = []) {
  if (!obj || typeof obj !== "object") return result;
  if (Array.isArray(obj)) { obj.forEach(i => collectElements(i, localName, result)); return result; }
  for (const [k, v] of Object.entries(obj)) {
    const local = k.includes(":") ? k.split(":").pop() : k;
    if (local === localName) {
      Array.isArray(v) ? v.forEach(i => result.push(i)) : result.push(v);
    }
    collectElements(v, localName, result);
  }
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 1 — Microsoft official schema validation
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 1 · Microsoft schema validation"));

try {
  const validation = await validateManifest(MANIFEST_PATH, false);

  if (validation.status && validation.status !== 200) {
    addWarning(`Validation gateway returned HTTP ${validation.status} — skipping schema check`);
  } else if (validation.report) {
    const notes    = validation.report.notes    ?? [];
    const warnings = validation.report.warnings ?? [];
    const errors   = validation.report.errors   ?? [];

    notes.forEach(n    => addInfo(`[gateway] ${n?.title ?? ""}: ${n?.content ?? n}`));
    warnings.forEach(w => addWarning(`[gateway] ${w?.title ?? ""}: ${w?.content ?? w}`));
    errors.forEach(e   => addError(`[gateway] ${e?.title ?? ""}: ${e?.content ?? e}`));

    if (validation.isValid) {
      addOk("Gateway: manifest is schema-valid");
      const products = validation.report.addInDetails?.supportedProducts ?? [];
      if (products.length) {
        addInfo(`Supported products: ${products.map(p => p.name ?? p.displayName ?? JSON.stringify(p)).join(", ")}`);
      }
    } else {
      addError("Gateway: manifest failed schema validation");
    }
  } else {
    addWarning("Gateway returned no report (possibly offline) — continuing with static checks");
  }
} catch (e) {
  addWarning(`Could not reach validation gateway: ${e.message}`);
}

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2a — Resource ID integrity
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2a · Resource ID integrity"));

// Collect all declared IDs from bt:Image, bt:Url, bt:String
const declaredResids = new Map(); // id → [elementType, ...]
["Image", "Url", "String"].forEach(tag => {
  collectElements(doc, tag).forEach(el => {
    const id = el?.["@_id"];
    if (!id) return;
    if (!declaredResids.has(id)) declaredResids.set(id, []);
    declaredResids.get(id).push(tag);
  });
});

addInfo(`Declared resource IDs: ${declaredResids.size}`);

// Duplicate resource IDs
const duplicateResids = [...declaredResids.entries()].filter(([, types]) => types.length > 1);
if (duplicateResids.length === 0) {
  addOk("No duplicate resource IDs");
} else {
  duplicateResids.forEach(([id, types]) =>
    addError(`Duplicate resource ID "${id}" declared ${types.length}× (types: ${types.join(", ")})`));
}

// All resid="…" usages
const usedResids = new Set(collectAttribs(doc, "@_resid"));
addInfo(`Resid references in use: ${usedResids.size}`);

// Broken: resid that points to no declared ID
const broken = [...usedResids].filter(id => !declaredResids.has(id));
broken.length === 0
  ? addOk("All resid references resolve to a declared resource")
  : broken.forEach(id => addError(`Unresolved resid="${id}" — no matching resource declared`));

// Orphaned: declared but never referenced
const orphaned = [...declaredResids.keys()].filter(id => !usedResids.has(id));
orphaned.length === 0
  ? addOk("No orphaned (unreferenced) resource IDs")
  : orphaned.forEach(id => addWarning(`Orphaned resource ID "${id}" declared but never referenced`));

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2b — Control & group ID uniqueness
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2b · Control & group ID uniqueness"));

// All @_id values that are NOT resource IDs (those are in bt:* elements)
const allIds = collectAttribs(doc, "@_id");
const nonResIds = allIds.filter(id => !declaredResids.has(id));
const idCount = new Map();
nonResIds.forEach(id => idCount.set(id, (idCount.get(id) ?? 0) + 1));
const dupControls = [...idCount.entries()].filter(([, c]) => c > 1);

dupControls.length === 0
  ? addOk("No duplicate control/group IDs")
  : dupControls.forEach(([id, c]) => addError(`Duplicate id="${id}" appears ${c}×`));

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2c — Required top-level manifest fields
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2c · Required manifest fields"));

const REQUIRED = ["Id", "Version", "ProviderName", "DefaultLocale", "DisplayName", "Description"];
const root = doc["OfficeApp"]?.[0] ?? {};

REQUIRED.forEach(field => {
  const el  = root[field]?.[0];
  if (el === undefined || el === null) {
    addError(`Missing required element <${field}>`);
    return;
  }
  const val = typeof el === "object"
    ? (el["#text"] ?? el["@_DefaultValue"] ?? JSON.stringify(el))
    : el;
  addOk(`<${field}> = "${String(val).slice(0, 80)}"`);
});

// GUID format check for <Id>
const rawId = root["Id"]?.[0];
const idStr = typeof rawId === "object" ? (rawId["#text"] ?? "") : String(rawId ?? "");
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
if (idStr && !GUID_RE.test(idStr.trim())) {
  addError(`<Id> "${idStr}" is not a valid GUID format`);
} else if (idStr) {
  addOk(`<Id> is a valid GUID`);
}

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2d — FunctionFile / ExecuteFunction consistency
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2d · FunctionFile consistency"));

const functionFiles  = collectElements(doc, "FunctionFile");
const execActions    = collectElements(doc, "Action")
  .filter(a => a?.["@_xsi:type"] === "ExecuteFunction");
const funcNames      = new Set(
  collectElements(doc, "FunctionName").map(el =>
    typeof el === "string" ? el : el?.["#text"] ?? "")
);

if (functionFiles.length === 0 && execActions.length > 0) {
  addError(`${execActions.length} button(s) use ExecuteFunction but no <FunctionFile> is declared`);
} else if (functionFiles.length === 0) {
  addWarning("No <FunctionFile> declared — required if any button uses ExecuteFunction");
} else {
  functionFiles.forEach(ff => {
    const resid = ff?.["@_resid"];
    if (!resid) {
      addError("<FunctionFile> missing resid attribute");
    } else if (!declaredResids.has(resid)) {
      addError(`<FunctionFile resid="${resid}"> references unknown resource ID`);
    } else {
      addOk(`<FunctionFile resid="${resid}"> resolves`);
    }
  });
  addOk(`${execActions.length} ExecuteFunction action(s) covered by <FunctionFile>`);
}

if (funcNames.size > 0) {
  addInfo(`Function names declared: ${[...funcNames].join(", ")}`);
}

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2e — URL reachability
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2e · URL reachability"));

if (SKIP_URLS) {
  addInfo("URL checks skipped (--skip-urls)");
} else {
  const urlMap = new Map(); // url → resId

  collectElements(doc, "Url").forEach(el => {
    if (!el || typeof el !== "object") return;
    const url = el["@_DefaultValue"];
    if (url) urlMap.set(url, el["@_id"] ?? "bt:Url");
  });
  collectElements(doc, "Image").forEach(el => {
    if (!el || typeof el !== "object") return;
    const url = el["@_DefaultValue"];
    if (url) urlMap.set(url, el["@_id"] ?? "bt:Image");
  });

  const isLocal = (url) => /localhost|127\.\d+\.\d+\.\d+|::1/.test(url);
  const targets = [...urlMap.entries()].filter(([url]) => ONLY_LOCAL ? isLocal(url) : true);

  if (targets.length === 0) {
    addInfo("No URLs to probe");
  } else {
    addInfo(`Probing ${targets.length} URL(s) with ${TIMEOUT_MS}ms timeout…`);
    const localAgent = new https.Agent({ rejectUnauthorized: false });

    const probe = async ([url, resId]) => {
      const label = `${fmt.grey(resId)} → ${url}`;
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
      try {
        const res = await fetch(url, {
          method: "HEAD",
          signal: controller.signal,
          agent: isLocal(url) ? localAgent : undefined,
          redirect: "follow",
        });
        clearTimeout(timer);
        if (res.ok || res.status === 405) {
          addOk(`HTTP ${res.status}  ${label}`);
        } else if (res.status >= 400) {
          addError(`HTTP ${res.status}  ${label}`);
        } else {
          addWarning(`HTTP ${res.status}  ${label}`);
        }
      } catch (e) {
        clearTimeout(timer);
        const reason = e.name === "AbortError" ? `timeout after ${TIMEOUT_MS}ms` : e.message;
        isLocal(url)
          ? addWarning(`Unreachable (${reason})  ${label}  ${fmt.grey("← local dev server, may not be running")}`)
          : addError(`Unreachable (${reason})  ${label}`);
      }
    };

    // Probe in batches of 8
    for (let i = 0; i < targets.length; i += 8) {
      await Promise.all(targets.slice(i, i + 8).map(probe));
    }
  }

  // ── PASS 2f — AppDomain coverage ───────────────────────────────────────────
  console.log(fmt.head("Pass 2f · AppDomain coverage"));

  const appDomains = collectElements(doc, "AppDomain").map(el =>
    typeof el === "string" ? el.trim() : String(el?.["#text"] ?? "").trim());

  const externalOrigins = new Set(
    [...urlMap.keys()]
      .filter(u => !isLocal(u))
      .map(u => { try { return new URL(u).origin; } catch { return null; } })
      .filter(Boolean)
  );

  if (externalOrigins.size === 0) {
    addInfo("No external origins to check against AppDomains");
  } else {
    externalOrigins.forEach(origin => {
      const covered = appDomains.some(d => d === origin || d.startsWith(origin));
      covered
        ? addOk(`AppDomain covers ${origin}`)
        : addWarning(`Host "${origin}" used in resources but not in <AppDomains>`);
    });
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// Summary
// ═══════════════════════════════════════════════════════════════════════════════
console.log(`\n${C.bold}${C.cyan}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${C.reset}`);
if (summary.errors === 0 && summary.warnings === 0) {
  console.log(`${C.bold}${C.green}✔  Manifest looks perfect — 0 errors, 0 warnings${C.reset}\n`);
} else {
  console.log(`${C.bold}Errors  : ${summary.errors > 0 ? C.red : C.green}${summary.errors}${C.reset}`);
  console.log(`${C.bold}Warnings: ${summary.warnings > 0 ? C.yellow : C.green}${summary.warnings}${C.reset}`);
  if (summary.errors === 0) {
    console.log(`\n${C.bold}${C.yellow}⚠  Manifest valid with warnings${C.reset}\n`);
  } else {
    console.log(`\n${C.bold}${C.red}✖  Manifest has errors${C.reset}\n`);
  }
}

process.exit(summary.errors > 0 ? 1 : 0);
