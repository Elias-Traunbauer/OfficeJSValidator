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
    "skip-urls":     { type: "boolean", default: false   },
    "check-images":  { type: "boolean", default: false   },
    "only-local":    { type: "boolean", default: false   },
    verbose:         { type: "boolean", short: "v", default: false },
    help:            { type: "boolean", default: false   },
  },
  allowPositionals: true,
});

if (values.help || positionals.length === 0) {
  console.log(`
Usage: node validate-manifest.mjs <manifest.xml> [options]

Options:
  --timeout <ms>   HTTP request timeout per URL (default: 5000)
  --skip-urls      Skip URL reachability checks entirely
  --check-images   Also probe bt:Image URLs (skipped by default)
  --only-local     Only probe localhost / 127.x.x.x URLs
  -v, --verbose    Show informational messages
  --help           Show this help
`);
  process.exit(0);
}

const MANIFEST_PATH = positionals[0];
const TIMEOUT_MS    = parseInt(values["timeout"], 10);
const SKIP_URLS     = values["skip-urls"];
const CHECK_IMAGES  = values["check-images"];
const ONLY_LOCAL    = values["only-local"];
const VERBOSE       = values["verbose"];

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
const addOk      = (msg) => { if (VERBOSE) console.log(fmt.ok(msg)); };
const addInfo    = (msg) => { if (VERBOSE) console.log(fmt.info(msg)); };

let _stageErrors = 0, _stageWarnings = 0;
const stageStart = () => { _stageErrors = summary.errors; _stageWarnings = summary.warnings; };
const stageEnd   = (name) => {
  if (summary.errors === _stageErrors && summary.warnings === _stageWarnings) {
    console.log(fmt.ok(`${name} — passed`));
  }
};

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
stageStart();

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
stageEnd("Pass 1 · Microsoft schema validation");

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2a — Resource ID integrity
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2a · Resource ID integrity"));
stageStart();

// Collect all declared IDs from bt:Images, bt:Urls, bt:ShortStrings, bt:LongStrings
const declaredResids = new Map(); // id → [elementType, ...]

// Images and Urls are straightforward
["Image", "Url"].forEach(tag => {
  collectElements(doc, tag).forEach(el => {
    const id = el?.["@_id"];
    if (!id) return;
    if (!declaredResids.has(id)) declaredResids.set(id, []);
    declaredResids.get(id).push(tag);
  });
});

// bt:ShortStrings and bt:LongStrings both contain bt:String children — differentiate by parent
function collectStringsByParent(obj, result = [], parentLocal = null) {
  if (!obj || typeof obj !== "object") return result;
  if (Array.isArray(obj)) { obj.forEach(i => collectStringsByParent(i, result, parentLocal)); return result; }
  for (const [k, v] of Object.entries(obj)) {
    const local = k.includes(":") ? k.split(":").pop() : k;
    if (local === "ShortStrings" || local === "LongStrings") {
      // Children of this node tagged as String are short/long
      collectStringsByParent(v, result, local);
    } else if (local === "String" && parentLocal) {
      const type = parentLocal === "ShortStrings" ? "ShortString" : "LongString";
      const items = Array.isArray(v) ? v : [v];
      items.forEach(el => {
        const id = el?.["@_id"];
        if (!id) return;
        if (!declaredResids.has(id)) declaredResids.set(id, []);
        declaredResids.get(id).push(type);
      });
    } else {
      collectStringsByParent(v, result, parentLocal);
    }
  }
  return result;
}
collectStringsByParent(doc);

const strCount  = [...declaredResids.values()].filter(t => t.includes("ShortString")).length;
const lstrCount = [...declaredResids.values()].filter(t => t.includes("LongString")).length;
const urlCount  = [...declaredResids.values()].filter(t => t.includes("Url")).length;
const imgCount  = [...declaredResids.values()].filter(t => t.includes("Image")).length;
addInfo(`Declared resource IDs: ${declaredResids.size} (${strCount} ShortString, ${lstrCount} LongString, ${urlCount} Url, ${imgCount} Image)`);

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

// ── Resource type expectations ───────────────────────────────────────────────
const RESID_TYPE_EXPECTATIONS = {
  Label:          "ShortString",
  Title:          "ShortString",
  Description:    "LongString",
  FunctionFile:   "Url",
  SourceLocation: "Url",
  Icon:           "Image",
  Image:          "Image",
};

// Walk the parsed doc looking for elements with a resid attribute
function collectResidUsages(obj, parentName = null, result = []) {
  if (!obj || typeof obj !== "object") return result;
  if (Array.isArray(obj)) { obj.forEach(i => collectResidUsages(i, parentName, result)); return result; }
  for (const [k, v] of Object.entries(obj)) {
    if (k === "@_resid") {
      result.push({ parent: parentName, resid: String(v) });
    } else {
      const local = k.includes(":") ? k.split(":").pop() : k;
      collectResidUsages(v, local, result);
    }
  }
  return result;
}

const residUsages = collectResidUsages(doc);

// Build a lookup: resid → expected type (from the element that references it)
const residExpectedType = new Map();
residUsages.forEach(({ parent, resid }) => {
  const expected = RESID_TYPE_EXPECTATIONS[parent];
  if (expected) residExpectedType.set(resid, expected);
});

// Broken: resid that points to no declared ID
const broken = [...usedResids].filter(id => !declaredResids.has(id));
if (broken.length === 0) {
  addOk("All resid references resolve to a declared resource");
} else {
  broken.forEach(id => {
    const expected = residExpectedType.get(id);
    const hint = expected ? ` (expected in ${expected === "ShortString" ? "bt:ShortStrings" : expected === "LongString" ? "bt:LongStrings" : "bt:" + expected + "s"})` : "";
    addError(`Unresolved resid="${id}" — no matching resource declared${hint}`);
  });
}

// Orphaned: declared but never referenced
const orphaned = [...declaredResids.keys()].filter(id => !usedResids.has(id));
orphaned.length === 0
  ? addOk("No orphaned (unreferenced) resource IDs")
  : orphaned.forEach(id => addInfo(`Orphaned resource ID "${id}" declared but never referenced`));

// Type-mismatch: resid resolves but to the wrong resource type
let typeMismatches = 0;

residUsages.forEach(({ parent, resid }) => {
  const expected = RESID_TYPE_EXPECTATIONS[parent];
  if (!expected) return;
  const actual = declaredResids.get(resid);
  if (!actual) return; // already reported as unresolved above
  if (!actual.includes(expected)) {
    addError(`<${parent} resid="${resid}"> expects a ${expected} resource but found ${actual.join(", ")}`);
    typeMismatches++;
  }
});

if (typeMismatches === 0 && broken.length === 0) {
  addOk("All resid references point to the correct resource type (ShortString/LongString/Url/Image)");
}
stageEnd("Pass 2a · Resource ID integrity");

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2b — Control & group ID uniqueness
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2b · Control & group ID uniqueness"));
stageStart();

// All @_id values that are NOT resource IDs (those are in bt:* elements)
const allIds = collectAttribs(doc, "@_id");
const nonResIds = allIds.filter(id => !declaredResids.has(id));
const idCount = new Map();
nonResIds.forEach(id => idCount.set(id, (idCount.get(id) ?? 0) + 1));
const dupControls = [...idCount.entries()].filter(([, c]) => c > 1);

dupControls.length === 0
  ? addOk("No duplicate control/group IDs")
  : dupControls.forEach(([id, c]) => addError(`Duplicate id="${id}" appears ${c}×`));
stageEnd("Pass 2b · Control & group ID uniqueness");

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2c — Required top-level manifest fields
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2c · Required manifest fields"));
stageStart();

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
stageEnd("Pass 2c · Required manifest fields");

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2d — FunctionFile / ExecuteFunction consistency
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2d · FunctionFile consistency"));
stageStart();

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
stageEnd("Pass 2d · FunctionFile consistency");

// ═══════════════════════════════════════════════════════════════════════════════
// PASS 2e — URL reachability
// ═══════════════════════════════════════════════════════════════════════════════
console.log(fmt.head("Pass 2e · URL reachability"));
stageStart();

if (SKIP_URLS) {
  addInfo("URL checks skipped (--skip-urls)");
  stageEnd("Pass 2e · URL reachability");
} else {
  const urlMap = new Map(); // url → resId

  collectElements(doc, "Url").forEach(el => {
    if (!el || typeof el !== "object") return;
    const url = el["@_DefaultValue"];
    if (url) urlMap.set(url, el["@_id"] ?? "bt:Url");
  });
  if (CHECK_IMAGES) {
    collectElements(doc, "Image").forEach(el => {
      if (!el || typeof el !== "object") return;
      const url = el["@_DefaultValue"];
      if (url) urlMap.set(url, el["@_id"] ?? "bt:Image");
    });
  } else {
    addInfo("Skipping bt:Image URL checks (use --check-images to include)");
  }

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
          addWarning(`HTTP ${res.status}  ${label}`);
        } else {
          addWarning(`HTTP ${res.status}  ${label}`);
        }
      } catch (e) {
        clearTimeout(timer);
        const reason = e.name === "AbortError" ? `timeout after ${TIMEOUT_MS}ms` : e.message;
        const suffix = isLocal(url) ? `  ${fmt.grey("← local dev server, may not be running")}` : "";
        addWarning(`Unreachable (${reason})  ${label}${suffix}`);
      }
    };

    // Probe in batches of 8
    for (let i = 0; i < targets.length; i += 8) {
      await Promise.all(targets.slice(i, i + 8).map(probe));
    }
  }

  stageEnd("Pass 2e · URL reachability");

  // ── PASS 2f — AppDomain coverage ───────────────────────────────────────────
  console.log(fmt.head("Pass 2f · AppDomain coverage"));
  stageStart();

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
  stageEnd("Pass 2f · AppDomain coverage");
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
