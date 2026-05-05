// ── XML helpers using native DOMParser ───────────────────────────────────────

function parseXml(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const err = doc.querySelector("parsererror");
  if (err) throw new Error(err.textContent.split("\n")[0]);
  return doc;
}

function localName(el) { return el.localName || el.nodeName.split(":").pop(); }

// Collect all elements by local name anywhere in the doc
function qAll(doc, local) {
  return [...doc.getElementsByTagName("*")].filter(el => localName(el) === local);
}

// Collect all attribute values by local name from the entire doc
function collectAttr(doc, attrLocal) {
  const result = [];
  for (const el of doc.getElementsByTagName("*")) {
    for (const attr of el.attributes) {
      const aLocal = attr.localName || attr.name.split(":").pop();
      if (aLocal === attrLocal) result.push(attr.value);
    }
  }
  return result;
}

// Get attribute by local name from a single element
function attr(el, name) {
  for (const a of el.attributes) {
    const aLocal = a.localName || a.name.split(":").pop();
    if (aLocal === name) return a.value;
  }
  return null;
}

// Get text content of a direct child element by local name
function childText(el, childLocal) {
  for (const c of el.children) {
    if (localName(c) === childLocal) return c.textContent.trim() || c.getAttribute("DefaultValue") || "";
  }
  return null;
}

// ── Resource collection (ShortString vs LongString by parent) ────────────────
function collectDeclaredResids(doc) {
  const map = new Map(); // id -> [type, ...]

  // Images and Urls — children of bt:Images / bt:Urls
  for (const el of qAll(doc, "Image")) {
    const id = attr(el, "id");
    if (!id) continue;
    if (!map.has(id)) map.set(id, []);
    map.get(id).push("Image");
  }
  for (const el of qAll(doc, "Url")) {
    const id = attr(el, "id");
    if (!id) continue;
    if (!map.has(id)) map.set(id, []);
    map.get(id).push("Url");
  }

  // Strings — differentiate by parent container
  for (const el of qAll(doc, "String")) {
    const id = attr(el, "id");
    if (!id) continue;
    const parentLocal = el.parentElement ? localName(el.parentElement) : "";
    let type;
    if (parentLocal === "ShortStrings") type = "ShortString";
    else if (parentLocal === "LongStrings") type = "LongString";
    else continue; // not a resource string
    if (!map.has(id)) map.set(id, []);
    map.get(id).push(type);
  }

  return map;
}

// Collect resid usages with parent element context
function collectResidUsages(doc) {
  const result = [];
  for (const el of doc.getElementsByTagName("*")) {
    const resid = attr(el, "resid");
    if (resid) result.push({ parent: localName(el), resid });
  }
  return result;
}

// ── Constants ────────────────────────────────────────────────────────────────
const RESID_TYPE_EXPECTATIONS = {
  Label:          "ShortString",
  Title:          "ShortString",
  Description:    "LongString",
  FunctionFile:   "Url",
  SourceLocation: "Url",
  Icon:           "Image",
  Image:          "Image",
};

const REQUIRED_FIELDS = ["Id", "Version", "ProviderName", "DefaultLocale", "DisplayName", "Description"];
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function r(level, message) { return { level, message }; }

// ── Pass 2a: Resource ID integrity ───────────────────────────────────────────
function pass2a(doc) {
  const results = [];
  const declaredResids = collectDeclaredResids(doc);

  const strCount  = [...declaredResids.values()].filter(t => t.includes("ShortString")).length;
  const lstrCount = [...declaredResids.values()].filter(t => t.includes("LongString")).length;
  const urlCount  = [...declaredResids.values()].filter(t => t.includes("Url")).length;
  const imgCount  = [...declaredResids.values()].filter(t => t.includes("Image")).length;
  results.push(r("info", `Declared resource IDs: ${declaredResids.size} (${strCount} ShortString, ${lstrCount} LongString, ${urlCount} Url, ${imgCount} Image)`));

  // Duplicates
  const dups = [...declaredResids.entries()].filter(([, t]) => t.length > 1);
  if (dups.length === 0) {
    results.push(r("ok", "No duplicate resource IDs"));
  } else {
    dups.forEach(([id, types]) =>
      results.push(r("error", `Duplicate resource ID "${id}" declared ${types.length}\u00d7 (types: ${types.join(", ")})`)));
  }

  // All resid usages
  const usedResids = new Set(collectAttr(doc, "resid"));
  results.push(r("info", `Resid references in use: ${usedResids.size}`));

  // Type expectations lookup
  const residUsages = collectResidUsages(doc);
  const residExpectedType = new Map();
  residUsages.forEach(({ parent, resid }) => {
    const expected = RESID_TYPE_EXPECTATIONS[parent];
    if (expected) residExpectedType.set(resid, expected);
  });

  // Unresolved
  const broken = [...usedResids].filter(id => !declaredResids.has(id));
  if (broken.length === 0) {
    results.push(r("ok", "All resid references resolve to a declared resource"));
  } else {
    broken.forEach(id => {
      const expected = residExpectedType.get(id);
      const hint = expected
        ? ` (expected in ${expected === "ShortString" ? "bt:ShortStrings" : expected === "LongString" ? "bt:LongStrings" : "bt:" + expected + "s"})`
        : "";
      results.push(r("error", `Unresolved resid="${id}" \u2014 no matching resource declared${hint}`));
    });
  }

  // Orphaned
  const orphaned = [...declaredResids.keys()].filter(id => !usedResids.has(id));
  if (orphaned.length === 0) {
    results.push(r("ok", "No orphaned (unreferenced) resource IDs"));
  } else {
    orphaned.forEach(id => results.push(r("info", `Orphaned resource ID "${id}" declared but never referenced`)));
  }

  // Type mismatches
  let typeMismatches = 0;
  residUsages.forEach(({ parent, resid }) => {
    const expected = RESID_TYPE_EXPECTATIONS[parent];
    if (!expected) return;
    const actual = declaredResids.get(resid);
    if (!actual) return;
    if (!actual.includes(expected)) {
      results.push(r("error", `<${parent} resid="${resid}"> expects a ${expected} resource but found ${actual.join(", ")}`));
      typeMismatches++;
    }
  });
  if (typeMismatches === 0 && broken.length === 0) {
    results.push(r("ok", "All resid references point to the correct resource type"));
  }

  return { results, declaredResids };
}

// ── Pass 2b: Control & group ID uniqueness ───────────────────────────────────
function pass2b(doc, declaredResids) {
  const results = [];
  const allIds = collectAttr(doc, "id");
  const nonResIds = allIds.filter(id => !declaredResids.has(id));
  const idCount = new Map();
  nonResIds.forEach(id => idCount.set(id, (idCount.get(id) ?? 0) + 1));
  const dups = [...idCount.entries()].filter(([, c]) => c > 1);

  if (dups.length === 0) {
    results.push(r("ok", "No duplicate control/group IDs"));
  } else {
    dups.forEach(([id, c]) => results.push(r("error", `Duplicate id="${id}" appears ${c}\u00d7`)));
  }
  return results;
}

// ── Pass 2c: Required manifest fields ────────────────────────────────────────
function pass2c(doc) {
  const results = [];
  const root = doc.documentElement;

  REQUIRED_FIELDS.forEach(field => {
    const els = qAll(root, field);
    // Only match direct children or near-top-level (not deep nested like Supertip > Description)
    const el = els.find(e => {
      const pLocal = localName(e.parentElement);
      return pLocal === "OfficeApp" || pLocal === root.nodeName.split(":").pop();
    });
    if (!el) {
      results.push(r("error", `Missing required element <${field}>`));
      return;
    }
    const val = el.getAttribute("DefaultValue") || el.textContent.trim();
    results.push(r("ok", `<${field}> = "${val.slice(0, 80)}"`));
  });

  const idEl = qAll(root, "Id").find(e => {
    const pLocal = localName(e.parentElement);
    return pLocal === "OfficeApp" || pLocal === root.nodeName.split(":").pop();
  });
  const idStr = idEl ? idEl.textContent.trim() : "";
  if (idStr && !GUID_RE.test(idStr)) {
    results.push(r("error", `<Id> "${idStr}" is not a valid GUID format`));
  } else if (idStr) {
    results.push(r("ok", "<Id> is a valid GUID"));
  }
  return results;
}

// ── Pass 2d: FunctionFile consistency ────────────────────────────────────────
function pass2d(doc, declaredResids) {
  const results = [];
  const functionFiles = qAll(doc, "FunctionFile");
  const execActions = qAll(doc, "Action").filter(a => {
    const type = attr(a, "type") || a.getAttribute("xsi:type");
    return type === "ExecuteFunction";
  });
  const funcNames = new Set(qAll(doc, "FunctionName").map(el => el.textContent.trim()).filter(Boolean));

  if (functionFiles.length === 0 && execActions.length > 0) {
    results.push(r("error", `${execActions.length} button(s) use ExecuteFunction but no <FunctionFile> is declared`));
  } else if (functionFiles.length === 0) {
    results.push(r("warning", "No <FunctionFile> declared \u2014 required if any button uses ExecuteFunction"));
  } else {
    functionFiles.forEach(ff => {
      const resid = attr(ff, "resid");
      if (!resid) {
        results.push(r("error", "<FunctionFile> missing resid attribute"));
      } else if (!declaredResids.has(resid)) {
        results.push(r("error", `<FunctionFile resid="${resid}"> references unknown resource ID`));
      } else {
        results.push(r("ok", `<FunctionFile resid="${resid}"> resolves`));
      }
    });
    results.push(r("ok", `${execActions.length} ExecuteFunction action(s) covered by <FunctionFile>`));
  }

  if (funcNames.size > 0) {
    results.push(r("info", `Function names declared: ${[...funcNames].join(", ")}`));
  }
  return results;
}

// ── Pass 2f: AppDomain coverage ──────────────────────────────────────────────
function pass2f(doc) {
  const results = [];
  const urls = new Map();
  for (const el of qAll(doc, "Url")) {
    const dv = el.getAttribute("DefaultValue");
    const id = attr(el, "id");
    if (dv) urls.set(dv, id ?? "bt:Url");
  }

  const isLocal = (u) => /localhost|127\.\d+\.\d+\.\d+|::1/.test(u);
  const appDomains = qAll(doc, "AppDomain").map(el => el.textContent.trim());

  const externalOrigins = new Set(
    [...urls.keys()]
      .filter(u => !isLocal(u))
      .map(u => { try { return new URL(u).origin; } catch { return null; } })
      .filter(Boolean)
  );

  if (externalOrigins.size === 0) {
    results.push(r("info", "No external origins to check against AppDomains"));
  } else {
    externalOrigins.forEach(origin => {
      const covered = appDomains.some(d => d === origin || d.startsWith(origin));
      if (covered) {
        results.push(r("ok", `AppDomain covers ${origin}`));
      } else {
        results.push(r("warning", `Host "${origin}" used in resources but not in <AppDomains>`));
      }
    });
  }
  return results;
}

// ── Run all passes ───────────────────────────────────────────────────────────
function validate(xml) {
  let doc;
  try {
    doc = parseXml(xml);
  } catch (e) {
    return [{ name: "XML Parse Error", results: [r("error", `Malformed XML: ${e.message}`)] }];
  }

  const { results: r2a, declaredResids } = pass2a(doc);

  return [
    { name: "Resource ID integrity",        results: r2a },
    { name: "Control & group ID uniqueness", results: pass2b(doc, declaredResids) },
    { name: "Required manifest fields",     results: pass2c(doc) },
    { name: "FunctionFile consistency",     results: pass2d(doc, declaredResids) },
    { name: "AppDomain coverage",           results: pass2f(doc) },
  ];
}

// ── Icons ────────────────────────────────────────────────────────────────────
const ICONS = {
  error:   "\u2716",
  warning: "\u26a0",
  ok:      "\u2714",
  info:    "\u2139",
};

const CHEVRON = `<svg class="chevron" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg>`;

// ── Render ───────────────────────────────────────────────────────────────────
function render(passes) {
  const passesEl = document.getElementById("passes");
  const summaryEl = document.getElementById("summary");
  const resultsEl = document.getElementById("results");

  passesEl.innerHTML = "";
  summaryEl.innerHTML = "";
  resultsEl.hidden = false;

  let totalErrors = 0, totalWarnings = 0;

  passes.forEach(pass => {
    const errors   = pass.results.filter(r => r.level === "error").length;
    const warnings = pass.results.filter(r => r.level === "warning").length;
    totalErrors += errors;
    totalWarnings += warnings;

    const statusClass = errors > 0 ? "has-err" : warnings > 0 ? "has-warn" : "clean";
    const hasIssues = errors > 0 || warnings > 0;

    const details = document.createElement("details");
    details.className = "pass";
    if (hasIssues) details.open = true;

    const summary = document.createElement("summary");
    summary.innerHTML = `${CHEVRON}<span class="status-dot ${statusClass}"></span>${escHtml(pass.name)}`;
    details.appendChild(summary);

    const body = document.createElement("div");
    body.className = "pass-body";

    pass.results.forEach(item => {
      const line = document.createElement("div");
      line.className = `result-line ${item.level}`;
      line.innerHTML = `<span class="icon">${ICONS[item.level]}</span><span class="msg">${escHtml(item.message)}</span>`;
      body.appendChild(line);
    });

    details.appendChild(body);
    passesEl.appendChild(details);
  });

  const verdictClass = totalErrors > 0 ? "fail" : totalWarnings > 0 ? "warn" : "clean";
  const verdictText = totalErrors > 0
    ? "Manifest has errors"
    : totalWarnings > 0
      ? "Valid with warnings"
      : "All checks passed";

  if (totalErrors === 0) {
    summaryEl.innerHTML = `
      <div class="hero-pass">
        <svg class="hero-check" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
          <polyline points="22 4 12 14.01 9 11.01"/>
        </svg>
        <span class="hero-text">Looks good</span>
      </div>
    `;
  } else {
    summaryEl.innerHTML = `
      <span class="badge errors">${ICONS.error} ${totalErrors} error${totalErrors !== 1 ? "s" : ""}</span>
      <span class="badge warnings">${ICONS.warning} ${totalWarnings} warning${totalWarnings !== 1 ? "s" : ""}</span>
      <span class="verdict ${verdictClass}">${verdictText}</span>
    `;
  }
}

function escHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ── File handling ────────────────────────────────────────────────────────────
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const filenameEl = document.getElementById("filename");

function handleFile(file) {
  if (!file) return;
  filenameEl.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  filenameEl.hidden = false;

  const reader = new FileReader();
  reader.onload = () => {
    const passes = validate(reader.result);
    render(passes);
  };
  reader.readAsText(file);
}

dropzone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => handleFile(fileInput.files[0]));

// Prevent browser from opening dropped files
document.addEventListener("dragover", (e) => e.preventDefault());
document.addEventListener("drop", (e) => e.preventDefault());

dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  e.stopPropagation();
  dropzone.classList.remove("dragover");
  handleFile(e.dataTransfer.files[0]);
});
