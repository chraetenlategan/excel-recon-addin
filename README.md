# Compare — Excel add-in

A Microsoft Excel **task-pane add-in** that compares **any column on any sheet
against any column on any other sheet**, right inside the workbook, and colours
the result on your own cells.

No fixed layout, no headers to get right, no cashbook/bank/ledger roles — pick
two columns and it matches them.

Everything runs locally in Excel. **No data leaves your machine** — the hosting
(GitHub Pages) only serves the add-in's code, never your numbers.

---

## Using it

1. **Home ▸ Reconcile** to open the pane.
2. **Side A**: pick the sheet and the column (e.g. `Sheet1`, column `E`).
   **Side B**: pick the other sheet and column (e.g. `Sheet3`, column `F`).
3. Optionally **limit the rows** on either side. The box takes
   `B12:B25` (column *and* rows — the column overrides the picker),
   `12:25` (rows only, on the picked column), or `B` / `B:B` (whole column).
   Leave it blank to use the whole used column. The line under each side always
   shows the exact range that will be read.
4. Options:
   - **Ignore + / −** — `-100` matches `100` (on by default).
   - **Ignore case & spacing** — for text values.
5. Click **Compare**. Every non-blank cell in both ranges is filled:
   - 🟩 found on the other side,
   - 🟥 not found.

   Blank cells are left alone. Values are matched **one-for-one**: three `100`s
   on A against two on B leave the third one red.

Numbers compare as numbers (rounded to cents, `(50)`, `50-`, `R1 234,56` and
`50 DR` all understood); anything else compares as text.

**Nothing else in your workbook is touched.** The add-in only ever sets a cell's
**fill colour** — no values are written, nothing is merged or unmerged, and
fonts, borders, number formats and column widths are left exactly as they were.
**Clear colours** removes the fill from the last two ranges it painted, and
nothing else.

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
explicit `sheetName` to override. They work independently of the compare pane.

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
