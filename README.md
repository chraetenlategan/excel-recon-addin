# Reconcile — Excel add-in

A Microsoft Excel **task-pane add-in** that reconciles worksheets against each
other, right inside the workbook. It reuses the matching engines from the
[static-reconciliation](https://github.com/chraetenlategan/static-reconciliation)
web app, unchanged — the only new code is the Office.js glue that reads your
worksheets and writes the result sheets.

Two flows share the pane, switched at the top:

| Mode | What it does |
|---|---|
| **Classic** | Cashbook vs bank statement and/or general ledger, matched by **amount + date**. Knows what a cashbook is. |
| **Modular** | *Any* sheets, *any* columns. You draw the comparison — "this column equals that column" — and it reconciles on that composite key. Knows nothing about accounting. |

Everything runs locally in Excel. **No data leaves your machine** — the hosting
(GitHub Pages) only serves the add-in's code, never your numbers.

---

## Using it — Classic

1. Put each dataset on its own worksheet — e.g. tabs named `Cashbook`,
   `Bank Statement`, and (optionally) `Ledger`. Each just needs a header row
   with **Date**, **Description**, and **Amount** columns (or **Debit** +
   **Credit** instead of Amount). Column order doesn't matter — they're
   auto-detected.
2. **Home ▸ Reconcile** to open the pane.
3. Pick which sheet is the cashbook, and at least one of statement / ledger.
4. Click **Reconcile**. Results land on new `Recon - …` sheets, colour-coded:
   - 🟩 **Matched** — amount, date and description all line up.
   - 🟨 **Check description** — amount + date match but the description doesn't.
   - 🟥 **Not found** — no matching amount + date on the other side.
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

Each returns that row's status against the target sheet — e.g.
`=RECON.COMPARETOBS(D2, B2, C2)` in a cashbook row reports *Matched to Bank
Statement*, *Check description (Bank Statement)*, or *Not found on Bank
Statement*. The target sheet is auto-detected by name (bank/statement,
cashbook, ledger); pass an explicit `sheetName` to override. If you've already
run **Load & detect** in the pane, the formulas reuse that side's exact column
mapping.

---

## Using it — Modular Recon

The format-free flow. The classic engine needs a Date, a Description and an
Amount; this one asks you to describe the comparison instead, so it will
reconcile a stock count against a delivery note as happily as a cashbook against
a bank feed.

1. Switch the pane to **Modular**.
2. **Sheets** — **Load worksheets** pulls in every tab of the workbook. For each,
   click the row holding its **column names** (guessed on load), and untick the
   sheets that aren't part of the job.
3. **Model** — every included sheet is an entity box, every column an attribute.
   Drag from a column on one box to a column on another: that line *is* the
   instruction "compare these two". Several lines between the same two sheets
   build a composite key — a row matches only when **all** of them agree. Click a
   line to change how it's compared (text / number / date / digits only, plus
   options like ignoring sign or punctuation) or to remove it. **Suggest** links
   columns that share a name; the pickers under the canvas add the same links
   without dragging, which is easier in a narrow pane.
4. **Results** — **Reconcile** gives a block per pair of sheets: pairs matched,
   and every row that didn't. Click a row number to open that sheet in the Data
   view, where the compared values are painted by outcome; a row number there
   selects the real row in Excel. **Write sheets** puts the lot into the workbook
   as `Modular - …` tabs with the same colouring, plus a `Modular - Summary`.

Statuses are deliberately four, not two:

| Status | Means |
|---|---|
| **Matched** | Exactly one row on the other side carries the same linked values. |
| **Matched (repeated value)** | It matched, but that key occurs more than once — the counts are right, which row paired with which is arbitrary. |
| **Not found** | Nothing on the other side carries those values. |
| **No value** | A linked column was blank or unreadable, so the row was never compared. Two blanks never match. |

Values are compared **exactly, once normalised** — dates to a canonical
`YYYYMMDD`, numbers to a chosen number of decimals, text lower-cased. Nothing is
matched approximately, which is what lets 30 rows against 29 mean precisely one
missing row rather than a similarity score.

> A date column's day/month order is settled once for the whole column, so a
> sheet that mixes `07/17/2024` with `15/07/2024` will read some rows wrong.
> Format the column as real dates in Excel and reload the worksheets.

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
| `taskpane.js` | **The Excel-aware pane code** — reads sheets, calls the engine, writes results. Owns `writeSpecs`, the Office.js sheet writer both flows use. |
| `functions.js` / `functions.json` | The `=RECON.COMPARETO…` custom functions + their metadata. |
| `engine.js` | Classic reconciliation engine, ported verbatim from the web app. |
| `comparison.js` / `sheets.js` | Build the comparison / unmatched output sheet specs. |
| `flex-*.js` / `flex.css` | **Modular Recon**, ported from the web app's `js/flex-*.js`: `-model` (state), `-engine` (matching), `-setup` (step 1 + wiring), `-erd` (the canvas), `-results` (output + sheet specs). |
| `utils.js` | Value/date parsing helpers both engines need. |
| `assets/icon-*.png` | Ribbon icons. Regenerate with `python make_icons.py`. |
| `index.html` / `guide.html` | Landing page for the GitHub Pages root, and the how-to for coworkers. |

### How the two flows coexist

`body.flex-mode` decides which of the two `<div>`s in `<main>` is on screen, and
that is the entire integration: neither flow touches the other's state. The
classic flow owns worksheets prefixed `Recon - `, Modular owns `Modular - `, and
`writeSpecs(ctx, specs, prefix)` clears only its own prefix — so running one
never eats the other's output. Both are skipped when listing input sheets.

A sheet spec is `{ name, aoa, colWidths, bandRows[], titleRows[], rowFills{},
paintRects[], autofilter }`. `rowFills` is the classic three-colour row shading;
`paintRects` is Modular's arbitrary-colour rectangles, pre-merged into maximal
blocks by `flexFillRects` so a painted sheet costs a handful of Office.js range
operations rather than one per cell.

### Porting notes

The `flex-*.js` files are a close port of the web app's, so fixes travel between
the two. What genuinely differs:

- **Ingest.** The web app reads uploaded workbooks through SheetJS; here every
  sheet is a worksheet of the open workbook, read via Office.js in
  `flexReadWorkbook()`. Office reports a date cell as a serial number, so cells
  whose *number format* says "date" are converted back to `Date` objects
  (`isExcelDateFormat` / `excelSerialToDate` in `utils.js`) — otherwise a date
  column would be typed as numbers.
- **Output.** `XLSX.writeFile` is replaced by sheet specs handed to `writeSpecs`.
- **Layout.** A pane is narrow: the comparisons rail sits under the canvas rather
  than beside it, the two result tables stack, the per-sheet Data tabs became a
  picker, and the rail carries a two-dropdown form that adds the same link a drag
  would.

## Developing

There's no build step. Edit the files and push — GitHub Pages redeploys, and
Excel picks up the new `taskpane.*` on next load (the manifest URL is unchanged).

To debug against a local server instead of GitHub Pages, run any static server
in this folder over **HTTPS** (Office requires https) and point the manifest's
URLs at it — e.g. `npx office-addin-dev-certs install` then serve on
`https://localhost:3000`.
