# GPS Ecosystem Visual

An interactive, single-screen diagram of the **GPS (Growing Pathways for Students)** ecosystem at the South Carolina Department of Education. Communicates what GPS is — a coordinated portfolio of foundation, data services, tools, and the audiences they serve — for SCDE leadership and the public.

**Live site:** https://wcothran.github.io/gps-ecosystem-visual/

## What this is

A static HTML page that reads its content from an Excel workbook. Anyone at SCDE can update the visual by editing the workbook in SharePoint — no code changes required. A small "publisher" workflow drops the edited file into this repo, and GitHub Pages serves it to the public.

```
SharePoint (private)            GitHub Pages (public)
─────────────────────           ─────────────────────
GPS-Ecosystem.xlsx   ──────▶    index.html ← viewer
 • edited by SCDE staff         data.xlsx ← swapped here
 • SharePoint version history   publish.html ← publisher helper
```

## Three roles

| Role | Who | What they do |
|------|-----|--------------|
| **Editor** | Anyone with edit access to the SharePoint workbook | Edits cells / adds rows in Excel. SharePoint auto-saves and versions. |
| **Publisher** | One X-Team member, on a routine cadence | Downloads the latest `.xlsx` from SharePoint, opens `publish.html` to preview, drops `data.xlsx` into this repo. ~30 seconds. |
| **Viewer** | Anyone on the internet | Opens the live URL. No sign-in. |

## Files

| File | Purpose |
|------|---------|
| `index.html` | The public viewer. Loads `data.xlsx` and renders the visual. |
| `data.xlsx` | The published copy of the workbook. **This is the file the live site reads.** |
| `publish.html` | Local-only preview / export helper for the publisher. Not used at runtime. |
| `GPS-Ecosystem.xlsx` | A reference copy of the source-of-truth workbook (the canonical copy lives in SharePoint). |
| `README-FOR-EDITORS.md` | Plain-English guide for SCDE staff editing the SharePoint workbook. |
| `README-FOR-PUBLISHERS.md` | Step-by-step for the publisher when content changes. |
| `README-FOR-DEVS.md` | Technical notes for anyone changing the viewer code. |
| `build-xlsx.ps1` | One-time script that built the initial workbook from the legacy `ecosystem-data.jsx`. |
| `serve.ps1` | Local development server (`file://` won't fetch `.xlsx`). |

## Local development

```powershell
# Serve the folder on http://localhost:8000
./serve.ps1
```

See [`README-FOR-DEVS.md`](README-FOR-DEVS.md) for architecture details.

## Design system

The visual is styled to the SCDE Application Style Guide — Slate (`#2F3D4C`) primary, Gold (`#F1BA55`) accent, Teal (`#43718B`) and Navy (`#234058`) for sections, Poppins for headings, Aptos for body. All tokens are CSS custom properties in `index.html`.

## Future migration

The viewer's data load is a single async function — swap the `DATA_URL` constant and `loadData()` function and the rest of the code is unchanged. This means moving to Azure Blob, a SharePoint REST API, or a proper backend database with auth is a one-function change, not a rebuild.

## Contact

For content questions, contact the GPS X-Team. For viewer / code questions, open an issue.
