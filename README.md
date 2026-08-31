# Compare — Excel add-in

A Microsoft Excel **task-pane add-in** that compares **any column on any sheet
against any column on any other sheet**, right inside the workbook, and colours
the result on your own cells — and, on its **PDF finder** tab, matches those
same cells against a PDF.

No fixed layout, no headers to get right, no cashbook/bank/ledger roles — pick
two columns and it matches them.

Everything runs locally in Excel. **No data leaves your machine** — the hosting
(GitHub Pages) only serves the add-in's code, never your numbers.

---

## Using it

1. **Home ▸ Reconcile** to open the pane.
2. **Fastest way — select the cells.** Highlight the cells you want on the
   sheet, click **Use selected cells** under Side A, then highlight the second
   lot of cells (any sheet) and click **Use selected cells** under Side B. Each
   button fills in that side's sheet, column(s) and row range from the
   selection — ctrl-click several blocks and they all come across, and a whole
   selected column is trimmed to the used rows. Everything it fills in can
   still be edited by hand afterwards. Or set it up manually:
3. **Side A**: pick the sheet and the column (e.g. `Sheet1`, column `E`).
   **Side B**: pick the other sheet and column (e.g. `Sheet3`, column `F`).
   Either side can take **more than one column** — ctrl-click them. The columns
   of a side are pooled, so a value on A counts as found if it turns up in
   *any* of side B's columns (e.g. `E` on one sheet against `F` **or** `H` on
   the other).
4. Optionally **limit the rows** on either side. The box takes
   `B12:B25` (column *and* rows — the column overrides the picker),
   `12:25` (rows only, applied to every picked column), or `B` / `B:B` (whole
   column). Separate several pieces with commas — `F12:F25, H12:H25`.
   Leave it blank to use the whole used column. The line under each side always
   shows the exact range that will be read.
5. Options:
   - **Ignore + / −** — `-100` matches `100` (on by default).
   - **Ignore case & spacing** — for text values.
6. Click **Compare**. Every non-blank cell in both ranges is filled:
   - 🟩 found on the other side,
   - 🟥 not found.

   Blank cells are left alone. Values are matched **one-for-one** across the
   whole of each side: three `100`s on A against two on B leave the third one
   red, whichever of B's columns they sit in.

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

## PDF finder

Reconcile a column against a **PDF** without reading the statement line by line.

1. Highlight the cells you want to find, on any sheet.
2. **PDF finder ▸ Use selected cells**, then **Open PDF finder**.
3. The finder opens in a window of its own beside Excel. Pick a PDF (or drop one
   on it) and every one of those values is outlined where it is printed.
4. Tick an occurrence off and **its cell in Excel is filled**, straight away.

Leave the pane open while you work — it is the pane, not the finder window, that
colours the cells. **Clear ticks** takes the colour off exactly the cells it put
it on, and nothing else.

| Action | Result |
| --- | --- |
| Click a row, or `↑` `↓` | Selects it; its free occurrences turn navy and the page scrolls to one |
| Double-click a row (or `Enter`) | Ticks off **one** occurrence — and fills that cell in Excel |
| Double-click it again (or `→`) | Steps that row on to the next free occurrence |
| Click an outline on the page | Ticks it off against the first row waiting for it |
| Double-click a highlight on the page | Selects its cell in Excel |
| Right-click either side (or `Delete`) | Releases that row's tick, and its fill |
| **Use Excel selection** | Re-reads whatever is selected in Excel now |
| The coloured dot | Picks the tick colour — the same colour Excel fills with |
| **Exact** | Whole values only: `ABC Trading` stops matching `ABC Trading CC` |
| **Follow** | Selecting a row moves the Excel cursor to its cell |

Each cell claims **one** printed occurrence, so a value that appears three times
in the column ticks off three separate printings. Column `#` reads `2/4` — this
cell holds the second of four printed occurrences — or a bare `4` while it is
untouched. A greyed italic value with a red corner is not in the PDF at all; a
small circle means it is printed, but every printing is already spoken for. The
counter at the top reads *cells ticked / cells picked* — that number is the
reconciliation.

Amounts are compared as numbers, so the sheet's `1234.56` finds the statement's
`1 234,56`, `R1,234.56`, `(1 234,56)` and `1234.56-`, and a credit matches its
opposite debit. Matches never run across a token boundary, so searching `234.56`
never lights up part of `1 234.56`. Scanned PDFs with no text layer are read with
OCR automatically — locally, offline, from the copy of Tesseract in `vendor/`.

**Nothing leaves your machine.** The PDF is opened in the window itself; pdf.js
and Tesseract are vendored rather than fetched, so no page, amount or client
name is ever sent anywhere. The only thing the add-in writes to your workbook is
a cell's fill colour.

### How it talks to Excel

The finder is far too big for a task pane, so it opens as an **Office dialog** —
a window of its own that gets exactly one wire back to the pane. `pdffinder/wire.js`
splits each message into chunks small enough for that wire and puts them back
together on the other side; `pdffinder/bridge.js` is the finder's end and
`pdffinder-pane.js` is Excel's. The finder sends the *whole* set of ticked cells
on every change rather than each tick, so a dropped message can never leave the
sheet disagreeing with the column.

This needs the **DialogApi 1.2** requirement set — Excel 2019 or Microsoft 365.
The pane says so plainly on older builds; the other two tabs still work there.

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
| `pdffinder.html` / `.css` | The PDF finder window — opened as an Office dialog, not a pane. |
| `pdffinder/finder.js` | The finder itself: the page, the marks, and which printing belongs to which cell. |
| `pdffinder/pdfdoc.js` | The only file that touches pdf.js or Tesseract. Word boxes in page points; marks positioned in percentages. |
| `pdffinder/match.js` | Value normalisation and word-sequence matching. Pure — testable with node. |
| `pdffinder/claim.js` | Which printed occurrence belongs to which cell. Pure — testable with node. |
| `pdffinder/wire.js` / `bridge.js` | The chunked message channel between the finder and the pane. |
| `pdffinder-pane.js` | Excel's end of it: reads the selection, colours the cells. |
| `vendor/` | pdf.js and Tesseract, vendored so the finder works offline. |
| `dev.py` | Static server for looking at `pdffinder.html` in a browser while working on it. |

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

`python dev.py` serves the folder at `http://localhost:5174/` with the MIME types
pdf.js and Tesseract need, which is enough to work on `pdffinder.html` in an
ordinary browser: it renders with no Excel behind it, the column simply stays
empty. `bridge.js` exports `receive()` so a column can be pushed in from the
console. Office itself still requires HTTPS — that is what `serve.py` is for.

To debug against a local server instead of GitHub Pages, run any static server
in this folder over **HTTPS** (Office requires https) and point the manifest's
URLs at it — e.g. `npx office-addin-dev-certs install` then serve on
`https://localhost:3000`.
