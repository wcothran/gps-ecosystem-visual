# GPS Ecosystem Visual — developer notes

A short technical reference for the next person who touches this project.

## What this is

A statically-hosted, publicly-viewable visualization of the **GPS (Growing Pathways for Students)** ecosystem. Content is editable by any SCDE staff member via an Excel workbook on SharePoint. The visual itself is a single HTML page that fetches the workbook and renders it in the browser.

## Files

| File                      | Purpose                                                                                          |
|---------------------------|--------------------------------------------------------------------------------------------------|
| `index.html`              | The public viewer. Self-contained — no build, no bundler. Loads React/Babel/SheetJS from CDN.   |
| `data.xlsx`               | The published copy of the ecosystem content. Replaced by the publisher on each release.         |
| `publish.html`            | Drag-and-drop preview helper for publishers. Also useful for local development (see below).     |
| `GPS-Ecosystem.xlsx`      | The seeded initial workbook. After first delivery, the canonical version lives on SharePoint.   |
| `build-xlsx.ps1`          | **Historical seed only — do not run against an edited workbook.** Generated the initial `.xlsx`. The SharePoint copy has been authoritative since first publish; re-running this would overwrite editor changes. Kept for reference / disaster recovery only. |
| `merge-add-scope.ps1`     | One-time migration that read a downloaded copy of the live SharePoint workbook and added the `Scope` column without touching editor edits. Kept as an example for any future column-add operation. |
| `README-FOR-EDITORS.md`   | Editing guide for non-technical SCDE staff.                                                     |
| `README-FOR-PUBLISHERS.md`| Manual publishing workflow for the X-Team publisher.                                            |
| `README-FOR-DEVS.md`      | This file.                                                                                      |

## Architecture in one paragraph

The viewer (`index.html`) calls `fetch(DATA_URL)` on mount where `DATA_URL` is currently `'./data.xlsx'`. The returned ArrayBuffer is parsed with [SheetJS](https://sheetjs.com/) into two JSON arrays (one per sheet). The `Items` rows are grouped by `Section` and sorted by `Order`; the `Settings` rows are flattened into a `{Key: Value}` object. The whole thing is rendered as the layered framework view from the original prototype, re-skinned to SCDE design tokens.

## The migration seam

```js
async function loadData() {
  const res = await fetch(DATA_URL);
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  // …group, sort, return { itemsBySection, settings }
}
```

This is the **only** function that touches the data source. To migrate to a different backend:

- **Azure Blob:** change `DATA_URL` to the blob URL. Set up Power Automate to copy the SharePoint file there on save. No code change beyond the URL.
- **REST API / database:** replace the function body to return `{ itemsBySection, settings }` from whatever endpoint. The renderer above it doesn't know or care where the data came from.
- **Static JSON:** if `.xlsx` parsing is overkill, pre-export to JSON at publish time and `fetch().then(r => r.json())` instead.

The shape contract is the only thing that matters: an object of arrays of `{name, short, desc, future, dashboards: string[], docURL, order}` plus a flat settings map.

## Stack

- **React 18** (UMD build from unpkg)
- **Babel standalone** for inline JSX (matches the prototype's no-build runtime)
- **SheetJS 0.20.3** (`xlsx.full.min.js`) from `cdn.sheetjs.com`
- **Poppins** from Google Fonts; **Aptos** falls back to **Segoe UI** then `system-ui`

All loaded from CDN. There is no build step and no node_modules. To deploy: drop the files on any static host.

## Design tokens

Pulled from `SCDE_Application_Style_Guide.html` in `OneDrive/Delete/Claude_Cowork/`.

| Token       | Value     | Used for                                  |
|-------------|-----------|-------------------------------------------|
| slate       | `#2F3D4C` | Page title, Foundation accent             |
| navy        | `#234058` | Tools & Apps accent                       |
| teal        | `#43718B` | Data Services accent                      |
| gold        | `#F1BA55` | Audiences rail, `*` future marker         |
| slate-blue  | `#778BA5` | Governance rail                           |
| cream       | `#FFEDCC` | Dashboard chip background                 |
| alice       | `#F0F8FF` | Page background                           |
| ink         | `#0a0a0b` | Card text                                 |
| muted       | `#6b7280` | Subtitle text                             |
| border      | `#e5e7eb` | Card resting border                       |

Fonts:
- Heading: `'Poppins', 'Aptos', 'Segoe UI', system-ui, sans-serif`
- Body: `'Aptos', 'Segoe UI', system-ui, sans-serif`

Brandon Grotesque is the official SCDE display font but isn't on Google Fonts; Poppins is the closest free CDN equivalent. Aptos ships with modern Windows/Office; non-Windows visitors fall back to Segoe UI / system sans.

## Section → band mapping

The seven `Section` values in the spreadsheet map to four visual bands plus two rails:

| Section          | Rendered in          | Accent     |
|------------------|----------------------|------------|
| Audiences        | Left rail            | gold       |
| Tools & Apps     | Tools band (with Outputs)   | navy       |
| Outputs          | Tools band (with Tools & Apps) | navy   |
| Data Services    | Data Services band (with Inputs) | teal |
| Inputs           | Data Services band (with Data Services) | teal |
| Foundation       | Foundation band (4-col grid) | slate    |
| Governance       | Right rail           | slate-blue |

This mapping is hard-coded in `index.html` and `publish.html`. Changing it requires a code edit.

## Local development

Browsers block `file://`-served pages from fetching adjacent files via CORS. Two options for local dev:

**Option A (recommended): use `publish.html`.**
Open `publish.html` in any browser directly from disk, then drag `data.xlsx` onto it. No server needed; SheetJS reads the file via the File API, not fetch.

**Option B: run a tiny static server.**
```powershell
# Built-in PowerShell static server (Windows)
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add('http://localhost:8000/')
$listener.Start()
# …(full snippet in a gist if needed)
```
Or install `npx serve`, `python -m http.server`, or VS Code's Live Server extension if any of those are available.

Once served over `http://`, open `http://localhost:8000/index.html` and the fetch will work.

## Regenerating the seed workbook

If `ecosystem-data.jsx` ever changes upstream and you want to re-seed the workbook:

```powershell
powershell -ExecutionPolicy Bypass -File .\build-xlsx.ps1
```

This writes a fresh `GPS-Ecosystem.xlsx` next to the script. It requires Excel installed locally (uses Excel COM automation; no extra modules needed).

The script's seed data is currently inlined — keeping it in sync with `ecosystem-data.jsx` after future edits is a manual step. The expectation is that after first delivery, the SharePoint workbook becomes the canonical source and `build-xlsx.ps1` is rarely re-run.

## Deployment to GitHub Pages

1. Push these files to a GitHub repo (e.g. `scde/gps-ecosystem-visual`).
2. Settings → Pages → enable Pages from the `main` branch root.
3. The public URL is `<org>.github.io/gps-ecosystem-visual/`.

That's it. There's no build pipeline.

## Things to be aware of

- **CDN dependencies are unpinned to a major version range in our usage.** They're pinned to specific versions in the script tags (React 18.3.1, SheetJS 0.20.3, Babel 7.29.0), but if any of those CDNs go down, the viewer breaks. Mitigation: self-host the JS files in the same repo if availability matters.
- **Babel in-browser compilation has a perceptible startup cost** (~100-300ms). If this becomes annoying, pre-compile the JSX once and replace `<script type="text/babel">` with a regular `<script>`.
- **`fetch('./data.xlsx')` is cached aggressively by browsers and GitHub Pages CDN.** A hard refresh (Ctrl+Shift+R) usually clears it; if not, append a `?v=<timestamp>` query string. Long-term, the right fix is HTTP cache headers, but GitHub Pages doesn't let you customize those.

## What is NOT being built

These have been explicitly deferred until concrete need:

- In-page editor UI (Excel is the editor)
- Approval / review workflow (SharePoint permissions handle who can edit; version history handles mistakes)
- Adding/removing sections from the sheet (sections are visual-design decisions, not data)
- Authentication or access control on the viewer (it's intentionally public)
- A backend API or database (the migration seam means we can add one later)
