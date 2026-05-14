# Publisher guide — pushing edits to the public visual

You are the bridge between the **private SharePoint workbook** (where SCDE staff edit the content) and the **public visual on GitHub Pages** (what everyone sees).

Total time per publish: ~30 seconds. No software to install. No command line.

## The setup, one-time

The public visual lives in a GitHub repository — e.g. `github.com/scde/gps-ecosystem-visual`. The repository contains:

- `index.html` — the viewer page itself
- `data.xlsx` — the public copy of the ecosystem content
- `publish.html` — a preview tool (this is what you're reading about)

GitHub Pages serves these files at a public URL like `scde.github.io/gps-ecosystem-visual/`.

You will need:
- **Edit access** to the SharePoint workbook (`GPS-Ecosystem.xlsx`)
- **Write access** to the GitHub repository (a GitHub account on the SCDE org with at least Write permission on the repo)

Ask the X-Team lead to set this up if you don't have it yet.

## Publishing a change

When an editor tells you the SharePoint workbook is ready:

1. **Download the latest workbook from SharePoint.**
   - Open the SharePoint document library.
   - Click `GPS-Ecosystem.xlsx`.
   - **File → Download a copy** (or the three-dot menu → Download).
   - You'll get a file named `GPS-Ecosystem.xlsx` (or similar) in your Downloads folder.

2. *(Optional but recommended for big changes)* **Preview it.**
   - Open `publish.html` from the GitHub repo (or just open a local copy in your browser).
   - Drag the file you just downloaded onto the page.
   - You'll see exactly what the public viewer will show. Click around. Confirm it looks right.
   - Click **Download as data.xlsx** — this saves the same file but with the exact filename the viewer expects.

3. **Replace `data.xlsx` in the GitHub repo.**
   - Open the GitHub repo in your browser.
   - Click the `data.xlsx` file in the file list.
   - Click the **pencil icon** (Edit) or the three-dot menu → **Upload a new version** (varies by GitHub UI).
   - Drop the downloaded file into the upload box.
   - Add a short commit message: e.g. "Update content - 2026-05-13" or "Add new Data Service: Foo".
   - Click **Commit changes**.

4. **Wait ~1 minute** for GitHub Pages to redeploy.
5. **Reload the public URL.** Confirm the change is live.

That's it.

## Rollback if a bad publish goes out

Two layers of rollback:

**Layer 1 — GitHub.** In the GitHub repo:
1. Click `data.xlsx` → **History**.
2. Find the previous good commit, click it.
3. Click **View raw** to download that version.
4. Re-upload it through the normal publish flow.

**Layer 2 — SharePoint.** If you want to roll back the *source of truth* too:
1. In SharePoint, right-click `GPS-Ecosystem.xlsx` → **Version history**.
2. Pick an earlier version → **Restore**.
3. Then do a normal publish from that restored version.

## Migration to automated publishing (when the team is ready)

The manual step (#3 above) is the one we'd eventually replace with automation:

1. Create an **Azure Blob Storage** container in the SCDE tenant with **public read** access.
2. Set up a **Power Automate flow** with these steps:
   - Trigger: `When a file is modified` on the SharePoint library, filtered to `GPS-Ecosystem.xlsx`.
   - Action: `Create blob (V2)` — copy the file into the public blob container as `data.xlsx`, overwriting.
3. Change one line in `index.html`:
   ```js
   const DATA_URL = 'https://YOUR-STORAGE-ACCOUNT.blob.core.windows.net/gps/data.xlsx';
   ```
4. Commit and publish that change once. From then on, every save to SharePoint automatically updates the public visual within a minute.

When that's live, this guide collapses to a one-liner: **edit SharePoint, save — that's the publish.** No more downloading or uploading.

## Things to watch out for

- **Don't edit `data.xlsx` directly in GitHub.** It's a generated file. Edits there will be overwritten the next time someone publishes. Always edit the SharePoint workbook.
- **Browser cache.** If you publish and the public URL still shows old content, try a hard refresh (Ctrl+Shift+R / Cmd+Shift+R). GitHub Pages can serve from CDN cache for a minute or two.
- **Filename matters.** The file in the repo must be named `data.xlsx` (lowercase, no spaces). The viewer looks for that exact name. The `publish.html` preview tool's download button names it correctly automatically.
