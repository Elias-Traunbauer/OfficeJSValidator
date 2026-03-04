# Office.js Manifest Validator

A CLI tool for validating Office.js XML manifests. Runs Microsoft's official schema validation, then a second static analysis pass covering things the gateway doesn't catch.

## Setup

```bash
npm install office-addin-manifest fast-xml-parser node-fetch
```

## Usage

```bash
node validate-manifest.mjs <manifest.xml> [options]
```

| Option | Default | Description |
|---|---|---|
| `--timeout <ms>` | `5000` | HTTP request timeout per URL |
| `--skip-urls` | off | Skip URL reachability checks entirely |
| `--only-local` | off | Only probe `localhost` / `127.x.x.x` URLs |
| `--help` | | Print usage |

Exit code `0` = clean (or warnings only), `1` = errors found — suitable for CI.

## What it checks

**Pass 1 — Microsoft schema validation**  
POSTs the manifest to `validationgateway.omex.office.net`, the same endpoint used by the official `office-addin-manifest validate` command. If the gateway is unreachable (offline, on-prem CI), this pass degrades to a warning and the static checks still run.

**Pass 2 — Static analysis**

| Check | Errors on | Warns on |
|---|---|---|
| resid integrity | `resid="X"` with no matching declared resource | Declared resource IDs never referenced anywhere |
| Duplicate IDs | Same `id=` on multiple controls or groups | — |
| Required fields | Missing `<Id>`, `<Version>`, `<ProviderName>`, `<DefaultLocale>`, `<DisplayName>`, `<Description>` | — |
| GUID format | `<Id>` value that isn't a valid GUID | — |
| FunctionFile | `ExecuteFunction` buttons with no `<FunctionFile>` declared, or `<FunctionFile resid>` pointing to unknown resource | — |
| URL reachability | External URLs returning 4xx or unreachable | `localhost` URLs unreachable (dev server may just not be running) |
| AppDomain coverage | — | External host used in resources but missing from `<AppDomains>` |

## Notes

- Self-signed certificates on `localhost` are accepted (no `rejectUnauthorized`).
- URL probes use `HEAD` requests with a concurrency of 8.
- Orphaned resource IDs (declared but unreferenced) are informational only — they cause no runtime issues, the Office host only resolves IDs that are actually referenced.