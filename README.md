# Reconcile — Excel add-in

A Microsoft Excel **task-pane add-in** that reconciles worksheets against each
other, right inside the workbook. It reuses the matching engines from the
[static-reconciliation](https://github.com/chraetenlategan/static-reconciliation)
web app, unchanged — the only new code is the Office.js glue that reads your
worksheets and writes the result sheets.

It does one job: **cashbook vs bank statement and/or general ledger**, matched
by **amount + date**.

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
4. Click **Reconcile**. Results land on new `Recon - …` sheets. The outcome is
   carried by **colour on the amount cell**, not by status text:
   - 🟩 amount, date and description all line up.
   - 🟨 amount + date match but the description doesn't.
   - 🟥 no matching amount + date on the other side.

   The only text added to a copied sheet is the row number the match was found
   on, under a short `BS` / `GL` / `CB` column.
5. The pane also shows the **detected columns** so you can sanity-check the
   mapping. If it guessed wrong, rename your headers (Date / Description /
   Amount, or Debit + Credit) and reconcile again.

### In-cell formulas

You can also check a single row without running a full reconcile, using three
custom functions:

```
=RECON.COMPARETOBS(amount, date, [description], [sheetName])
=RECON.COMPARETOCB(amount, date, [description], [sheetName])
=RECON.COMPARETOGL(amount, date, [description], [sheetName])
```

Each returns a single mark against the target sheet — `✓` matched, `⚠` matched
but the description differs (or no amount), `✗` not found — so a filled-down
column reads at a glance and takes ordinary conditional formatting. The target
sheet is auto-detected by name (bank/statement, cashbook, ledger); pass an
explicit `sheetName` to override. If you've already clicked **Load** in the
pane, the formulas reuse that side's exact column mapping.

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
| `taskpane.js` | **The Excel-aware pane code** — reads sheets, calls the engine, writes results. Owns `writeSpecs`, the Office.js sheet writer. |
| `functions.js` / `functions.json` | The `=RECON.COMPARETO…` custom functions + their metadata. |
| `engine.js` | The reconciliation engine, ported verbatim from the web app. |
| `comparison.js` / `sheets.js` | Build the comparison / unmatched output sheet specs. |
| `utils.js` | Value/date parsing helpers the engine needs. |
| `assets/icon-*.png` | Ribbon icons. Regenerate with `python make_icons.py`. |
| `index.html` / `guide.html` | Landing page for the GitHub Pages root, and the how-to for coworkers. |

### Output sheets

A sheet spec is `{ name, aoa, colWidths, bandRows[], titleRows[], paintRects[],
autofilter }`. `paintRects` is the outcome colouring: `sheets.js` collects the
cells to fill and `_painter()` merges each column's consecutive same-colour
cells into one rectangle, which `paintCells` then applies a colour at a time
through a single `getRanges` call per 50 blocks — so colouring a long sheet is
a handful of Office.js operations rather than one per cell.

Office reports a date cell as a serial number; `_display_date` (`utils.js`)
converts anything above the 1900 epoch back to a date, so the engine sees the
same values the user does.

`writeSpecs` deletes every existing `Recon - ` sheet before writing, so each run
is clean; those sheets are also skipped when listing input sheets.

## Developing

There's no build step. Edit the files and push — GitHub Pages redeploys, and
Excel picks up the new `taskpane.*` on next load (the manifest URL is unchanged).

To debug against a local server instead of GitHub Pages, run any static server
in this folder over **HTTPS** (Office requires https) and point the manifest's
URLs at it — e.g. `npx office-addin-dev-certs install` then serve on
`https://localhost:3000`.
