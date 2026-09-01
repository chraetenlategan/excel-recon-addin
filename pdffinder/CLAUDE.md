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
| `pdffinder/debug.js` | The black box recorder both ends keep — loaded by both, first. |
| `pdffinder/debugpanel.js` | The finder window's **Debug** drawer. |

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

So each end keeps its own log, from its first line of script, and each can print
that log without the other's help — the one channel that would carry a joint
report is the thing under suspicion.

**Where to find it.** In the pane, under **Bridge diagnostics** on the PDF
finder tab. In the window, the **Debug** button (or Ctrl+Alt+D). Both hold a
textarea that can simply be selected and copied where the clipboard API is
locked down inside a dialog.

**How to read it.** Press **Test bridge** in the pane. The same `ping` goes down
all three forms of `messageChild` Office offers, each labelled; `bridge.js`
answers every ping it receives three times over, once down each form of
`messageParent`, labelled the same way. Four seconds later:

| What the two logs show | What is broken |
| --- | --- |
| Pane logs `test.ping`, window's log has no `in.ping` | **pane → window**. Look at `dialog.opened` (cross-origin?) and `in.handler` in the window. |
| Window logs `in.ping`, pane logs no `test.pong` | **window → pane**. This is the failure that leaves a double-click hanging. |
| Some labels come back, others do not | The channel works in one form only — make `pfPost` / `rawPost` prefer that one. |
| Both, in a few ms | The bridge is fine; the fault is in what the message asked for — read the `excel.*` lines. |

Tag names are a dotted path so a log can be read down its left column:
`dialog.*` opening the window, `out.*` / `in.*` chunks each way, `wire.*` the
codec (`wire.notString` is the classic — a handler's event object fed to the
reader instead of the string on `.message`), `excel.*` what the pane did to the
workbook, `test.*` the bridge test, `window.error` anything that threw where
nobody was catching.

**Fetch window log** (pane) and **Send to Excel** (window) put both logs in one
report — but only where the return leg works at all, which is exactly the case
where it is least needed. When it does not, copy each end separately.

Recording is always on. It is a few hundred short strings in memory, and it
means a failure five minutes old is still there when somebody thinks to look.
