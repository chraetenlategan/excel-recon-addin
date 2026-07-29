# Reconcile — Excel add-in

A Microsoft Excel **task-pane add-in** that reconciles a cashbook against a bank
statement and/or general ledger by **amount + date**, right inside the workbook.
It reuses the matching engine from the
[static-reconciliation](https://github.com/chraetenlategan/static-reconciliation)
web app, unchanged — the only new code is the Office.js glue that reads your
worksheets and writes a **Recon Results** sheet.

Everything runs locally in Excel. **No data leaves your machine** — the hosting
(GitHub Pages) only serves the add-in's code, never your numbers.

---

## Using it

1. Put each dataset on its own worksheet — e.g. tabs named `Cashbook`,
   `Bank Statement`, and (optionally) `Ledger`. Each just needs a header row
   with **Date**, **Description**, and **Amount** columns (or **Debit** +
   **Credit** instead of Amount). Column order doesn't matter — they're
   auto-detected.
2. **Home ▸ Reconcile** to open the pane.
3. Pick which sheet is the cashbook, and at least one of statement / ledger.
4. Click **Reconcile**. Results land on a new `Recon Results` sheet, colour-coded:
   - 🟩 **Matched** — amount, date and description all line up.
   - 🟨 **Check description** — amount + date match but the description doesn't.
   - 🟥 **Not found** — no matching amount + date on the other side.
5. The pane also shows the **detected columns** so you can sanity-check the
   mapping. If it guessed wrong, rename your headers (Date / Description /
   Amount, or Debit + Credit) and reconcile again.

---

## Installing (sideload)

The add-in is hosted on GitHub Pages; you install a small **manifest** file once
per machine. After that, any update pushed to this repo is picked up
automatically — nobody re-installs anything.

### Windows desktop Excel — shared-folder catalog (best for a small team)

1. Download **[`manifest.xml`](https://chraetenlategan.github.io/excel-recon-addin/manifest.xml)**
   and drop it in a folder everyone can reach — a network share
   (`\\server\addins\`) for a team, or any local folder to try it yourself.
2. In Excel: **File ▸ Options ▸ Trust Center ▸ Trust Center Settings… ▸
   Trusted Add-in Catalogs**.
3. Paste the folder path in **Catalog Url**, click **Add catalog**, tick
   **Show in Menu**, **OK**, and **restart Excel**.
4. **Home ▸ Add-ins ▸ More Add-ins ▸ Shared Folder** (tab) ▸ select
   **Reconcile** ▸ **Add**.

### Excel on the web / Mac — upload the manifest

1. Download `manifest.xml` (link above).
2. **Home ▸ Add-ins ▸ More Add-ins ▸ Upload My Add-in** ▸ choose the file.

> The same `manifest.xml` works on all platforms. Only the install step differs.

### Later: central deployment (no per-machine setup)

When you're ready to roll it out properly, a Microsoft 365 admin can push it to
everyone via **Microsoft 365 admin center ▸ Settings ▸ Integrated apps ▸
Upload custom apps**, using this same manifest. No code changes needed.

---

## Repo layout

| File | Role |
|---|---|
| `manifest.xml` | What you sideload. Points Excel at the hosted `taskpane.html`. |
| `taskpane.html` / `.css` | The pane UI (loads Office.js from Microsoft's CDN). |
| `taskpane.js` | **The only Excel-aware code** — reads sheets, calls the engine, writes results. |
| `engine.js` | Reconciliation engine, ported verbatim from the web app. |
| `utils.js` | Value/date parsing helpers the engine needs. |
| `assets/icon-*.png` | Ribbon icons. Regenerate with `python make_icons.py`. |
| `index.html` | Plain landing page for the GitHub Pages root. |

## Developing

There's no build step. Edit the files and push — GitHub Pages redeploys, and
Excel picks up the new `taskpane.*` on next load (the manifest URL is unchanged).

To debug against a local server instead of GitHub Pages, run any static server
in this folder over **HTTPS** (Office requires https) and point the manifest's
URLs at it — e.g. `npx office-addin-dev-certs install` then serve on
`https://localhost:3000`.
