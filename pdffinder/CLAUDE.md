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
6. And the other way about: **selecting a cell in Excel outlines that value on
   the page** and scrolls to it. The two windows follow each other's cursor.

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

**Following the cursor never ticks anything.** A cell selected in Excel is
outlined on the page and scrolled to, and that is all. Looking at a value and
reconciling it are different acts, and only the second may colour a cell. The
same rule the other way: `goto` moves the Excel cursor and paints nothing.

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
| `pdffinder/debug.js` | The black box recorder both ends keep — loaded by both, first. No UI: read it from the console with `PFDebug.report()`. |

Office gives a dialog one string-sized wire in each direction, so every message
is chunked by `wire.js` and reassembled on the other side. Both inbound handlers
are handed an **event object**, never the string that was sent — the chunk is on
`arg.message`, and feeding the reader anything else drops every message in
silence, greeting included. Messages:

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
| pane | `look` | `{sheet, ref, v}` — the cell the Excel cursor is on; show it on the page. |
| either | `ping` / `pong` | The bridge test — see below. Answered by `bridge.js` itself. |
| pane | `report` | Send me your log. |
| finder | `log` | `{text}` — this window's log, for the pane's report. |

## Where it is served from

Beside the pane, on the add-in's own origin — `PF_BASE` in `pdffinder-pane.js`
is blank and that is what blank means. The finder ships with the add-in, so it
is always the same version as the pane and always on the HTTPS the pane already
loads over.

`PF_BASE` is the one line to change to serve it from elsewhere. Office refuses a
plain `http://` dialog (localhost excepted) and needs any other origin listed in
the manifest's `<AppDomains>`, with both ends naming each other in
`targetOrigin` — the pane puts its own origin on the query string and
`bridge.js` replies to that. Such a host must also carry the whole finder
(`pdffinder.html`, `pdffinder.css`, `pdffinder/`, `vendor/`, `assets/`) at this
same version: a server holding some other app answers with its own page or a
404, and the window opens on nothing. If the configured home will not open at
all, the pane falls back to the copy beside it and says so.

## When the bridge goes quiet

The two ends are two windows, and when the wire between them stops there is
nothing on either screen to say **which** direction stopped. A pane that never
answers and a window that never asks look identical from the page: a
double-click sits on *Looking for 5 400,00 in Excel…* for ever.

So each end keeps its own log, from its first line of script, in `debug.js`.
Neither end shows it on screen — there is no diagnostics drawer and no bridge
test any more — so read it from that window's console:
`PFDebug.report()` in the pane, and the same in the finder window.

Tag names are a dotted path so a log can be read down its left column:
`dialog.*` opening the window, `out.*` / `in.*` chunks each way, `wire.*` the
codec (`wire.notString` is the classic — a handler's event object fed to the
reader instead of the string on `.message`), `excel.*` what the pane did to the
workbook, `window.error` anything that threw where nobody was catching.

Every file of the bridge registers its build (`PFDebug.file`), and the report
prints them on its second line. The window also puts its build in `ready`, and
the pane says so in plain words when the two disagree.

GitHub Pages serves these files with `Cache-Control: max-age=600`, so for ten
minutes after a push the WebView can keep handing out the previous copy — long
enough to test a fix that is already deployed and conclude it did not work. To
force the issue, close Excel and empty
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, or simply wait the ten minutes.
Either way, **read the `builds:` line before reading anything else**: a report
with no `builds:` line at all is an old build by definition. A WebView holding one file back while the others
move on produces symptoms that make no sense — most of all a codec that drops
messages without complaint, because the copy doing the dropping is the copy
without the complaints in it. **Check that line first.**

Two rules this section is built on, both learned the hard way:

- **Logging must never be able to break the thing it logs.** `wire.js` shipped
  once calling a `say` that was never defined; the throw landed on the line
  after a successful decode, so every message arrived and none was delivered,
  and the log that would have said so was the code that was broken. `say` and
  `codes` there swallow their own faults for that reason.
- **A handler's fault is not a dropped message.** `wire.js` catches what
  `onMessage` throws and logs it as `wire.handlerThrew`, so the two can never
  again be mistaken for one another.

Recording is always on. It is a few hundred short strings in memory, and it
means a failure five minutes old is still there when somebody thinks to look.

## Following the Excel cursor

Office gives an add-in **no double-click event on a cell** — `onSelectionChanged`
is the whole of what it offers — so the gesture that finds a value on the page is
*selecting* the cell, which is the one an auditor makes anyway while reading down
a column. It is a checkbox in the pane (*Show the cell I select on the PDF*),
remembered between sessions.

Three things make it behave rather than thrash:

- **The echo is ignored.** The pane moves the Excel cursor itself, for `goto` and
  `find`. Those moves come back as selection changes like any other, and
  following them would have the two windows chasing each other, so a move the
  pane made is ignored for half a second (`pfOwnMove`).
- **Arrowing down a column is debounced**, and the same cell selected twice
  running is not sent twice.
- **Only the first cell of a selection** is sent. A dragged block is one gesture
  and its corner is what the user pointed at.

On the window's side, a cell **in scope** is taken in hand exactly as clicking its
mark would — so the arrow keys step it on to its next printing — and it lands on
its own tick if it has one, otherwise on the first printing still going spare. A
cell **outside the scope** (another sheet, another column) is outlined without
being adopted: it has no row to belong to, and inventing one would put a value in
the count that nobody asked to reconcile.
