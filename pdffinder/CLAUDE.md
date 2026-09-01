# PDF finder

The add-in's third tab: reconcile cells of the open workbook against a **PDF**,
in a window of its own beside Excel.

This file covers the whole feature, both ends of it — the window's code lives in
`pdffinder/`, the pane's beside it at the repository root.

## How it works, from the user's side

1. The user opens the add-in and picks the **PDF finder** tab.
2. That tab holds the constraints, and nothing else: a **sheet**, and optionally
   the **columns** (`C`, `A,B`, `C:E`) and **rows** (`12:250`) to narrow it to.
   Blank means every used column, every used row. These are settled *before* the
   window opens — they are the user's whole say over what is searched for.
3. **Open PDF finder** reads those cells off the sheet and opens the finder
   window. A mistyped column is said in the pane; the window only opens once the
   scope is good.
4. The user picks (or drops) a PDF. Every value in scope is outlined where it is
   printed. Scanned pages are OCR'd **automatically** — nobody has to know what
   OCR is or press anything to get it.
5. Clicking an outline ticks it off, and **its cell in Excel is filled** in the
   marker colour, straight away. Double-clicking any value on the page — in
   scope or not — puts the Excel cursor on it.

## Rules this feature is built on

**No spreadsheet in the finder window.** The cells are in Excel, one window
across, where the auditor can already see them. The window shows the statement,
a count, and the few switches that change how matching works. Never add a column
of cells, a list of rows, a formula bar or anything else that duplicates Excel
inside it — that was the old shape and it was removed on purpose.

**The pane stays small.** It is a task pane a few centimetres wide, sharing
space with two other tools. It holds the scope, one button and one line of
status. Instructions, marker swatches and tick-clearing belong in the finder
window, which has room for them.

**Only fill colour is ever written to a cell.** Same promise as the compare
tabs. `pfPainted` in `pdffinder-pane.js` is the exact list of cells the pane
filled, so releasing a tick takes the colour off that cell and nothing else.
No value, format or formula is touched.

**One cell claims one printed occurrence.** Three 50s in scope tick off three
separate printed 50s; a fourth printing stays open. That is `claim.js`, and it
is what makes the count at the top a reconciliation rather than a search hit
count.

**The finder sends the whole set of ticks, never a delta.** A dropped message on
the Office dialog channel can then never leave the sheet disagreeing with the
page; `pfPaint` diffs against what it last painted.

**Nothing leaves the machine.** pdf.js and Tesseract are vendored under
`vendor/`. No page, amount or file name is uploaded anywhere.

## The two ends, and the wire between them

| File | Role |
| --- | --- |
| `taskpane.html` (`#tab-pdf`) / `taskpane.css` | The three constraint fields and the button. |
| `pdffinder-pane.js` | Excel's end: reads the scope, colours cells, moves the cursor, hunts values down in the workbook. |
| `pdffinder.html` / `pdffinder.css` | The finder window — an Office dialog, not a pane. |
| `pdffinder/finder.js` | The window's brain: state, marks, ticking, and every event. |
| `pdffinder/pdfdoc.js` | pdf.js loading, page rendering, word boxes, OCR. |
| `pdffinder/match.js` | Value normalisation, word-sequence matching, `valueAt` for a click. |
| `pdffinder/claim.js` | Which printed occurrence belongs to which cell. |
| `pdffinder/store.js` | Ticks and marker colour in `localStorage`, per PDF + scope. |
| `pdffinder/bridge.js` | The dialog's end of the Office channel. |
| `pdffinder/wire.js` | The chunking codec both ends speak — loaded by both. |

Office gives a dialog one string-sized wire in each direction, so every message
is chunked by `wire.js` and reassembled on the other side. Messages:

| From | Message | Meaning |
| --- | --- | --- |
| finder | `ready` | The window is up; send it the colour and the cells. |
| pane | `rows` | `{sheet, scope, cells:[{ref,v}]}` — the scope, in reading order. |
| pane | `colour` | The marker colour the pane last knew about. |
| finder | `pull` | Re-read the scope (the **Re-read cells** button). |
| finder | `colour` | The user changed the marker; remember it and repaint. |
| finder | `ticks` | `{sheet, hex, refs, all, hit}` — every ticked cell, every time. |
| finder | `goto` | Put the Excel cursor on one cell of the scope. |
| finder | `find` | Look a printed value up anywhere in the workbook. |
| pane | `found` | What that lookup landed on, for the line over the page. |

## Where it is served from

`PF_BASE` in `pdffinder-pane.js` is the finder's home and the one line to
change. Office refuses a plain `http://` dialog (localhost excepted) and needs
any other origin listed in the manifest's `<AppDomains>`, with both ends naming
each other in `targetOrigin` — the pane puts its own origin on the query string
and `bridge.js` replies to that. If the configured home will not open, the pane
falls back to the copy beside it and says so.
