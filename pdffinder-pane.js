"use strict";

/**
 * pdffinder-pane.js — the workbook's end of the PDF finder.
 *
 * The finder itself is a page (`pdffinder.html`) too big to live in a task
 * pane, so it opens as an Office dialog: a window of its own, beside Excel,
 * talking back down the one wire Office gives it. This file is the other end.
 * It reads the selected cells, sends them over, and colours the sheet as the
 * auditor ticks them off on the page.
 *
 * The only property it ever sets on a cell is its fill colour, the same promise
 * the compare tabs make. `pfPainted` is the exact list of cells it filled, so
 * "Clear ticks" takes off those and nothing else.
 *
 * Loaded after taskpane.js, and it borrows that file's selection parsing
 * (`parseSelection`, `colLetter`, `supports`) rather than repeating it.
 */

// What was sent to the finder, in the order it was sent.
const pfState = { sheet: "", cells: [] };
// What is currently coloured on the sheet, so it can be taken off again.
let pfPainted = { sheet: "", refs: [] };
let pfHex = "#FFE94D";
let pfDialog = null;

const PF_TICK_KEY = "vdm-pdffinder:tick";
const PF_URL = new URL("pdffinder.html", window.location.href).href;

function initPdfFinder() {
  try {
    const saved = localStorage.getItem(PF_TICK_KEY);
    if (saved) pfHex = JSON.parse(saved);
  } catch { /* first run */ }
  pfSwatch();

  $("pf-use-selection").onclick = () => pfReadSelection(true);
  $("pf-open").onclick = pfOpen;
  $("pf-clear").onclick = pfClear;
}

const pfSwatch = () => { $("pf-dot").style.background = pfHex; };

function pfSetStatus(msg, isError) {
  const box = $("pf-status");
  box.textContent = msg || "";
  box.classList.toggle("err", !!isError);
}

/* ---------- reading the selection ---------- */

/**
 * The cells the user has highlighted, in reading order, blanks dropped.
 *
 * Values come off the sheet as `text` — what the auditor sees, which is what is
 * printed on the statement — except for plain numbers, where the underlying
 * value is used instead: a column displayed to the rand still has to find its
 * cents on the page.
 */
async function pfReadSelection(announce) {
  const btn = $("pf-use-selection");
  btn.disabled = true;
  try {
    let address = "";
    await Excel.run(async (ctx) => {
      const sel = supports("1.9") ? ctx.workbook.getSelectedRanges() : ctx.workbook.getSelectedRange();
      sel.load("address");
      await ctx.sync();
      address = sel.address;
    });

    const { sheet, areas } = parseSelection(address);
    if (!sheet || !areas.length) throw new Error("select some cells in Excel first.");

    const cells = [];
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const used = ws.getUsedRangeOrNullObject(true);
      used.load(["rowIndex", "rowCount"]);
      await ctx.sync();

      // Clicking a column header selects a million rows; there is no sense
      // reading past the data.
      const lastUsed = used.isNullObject ? 0 : used.rowIndex + used.rowCount;
      const blocks = [];
      for (const a of areas) {
        const first = Math.max(a.firstRow === null ? 1 : a.firstRow, 1);
        const last = Math.min(a.lastRow === null ? lastUsed : a.lastRow, lastUsed);
        if (last < first) continue;
        const letter = colLetter(a.column);
        const range = ws.getRange(`${letter}${first}:${letter}${last}`);
        range.load(["values", "text", "numberFormat"]);
        blocks.push({ letter, first, range });
      }
      if (!blocks.length) throw new Error("that selection holds no data.");
      await ctx.sync();

      for (const b of blocks) {
        const vals = b.range.values, texts = b.range.text, fmts = b.range.numberFormat;
        for (let i = 0; i < vals.length; i++) {
          const v = pfCellValue(vals[i][0], texts[i][0], fmts[i][0]);
          if (v !== "") cells.push({ ref: b.letter + (b.first + i), v });
        }
      }
    });

    if (!cells.length) throw new Error("every cell in that selection is blank.");
    pfState.sheet = sheet;
    pfState.cells = cells;
    $("pf-range").textContent = `${sheet} · ${cells.length} cell${cells.length === 1 ? "" : "s"}`;
    pfSend({ t: "rows", sheet, cells });
    if (announce) pfSetStatus(`${cells.length} cells ready.`);
    return true;
  } catch (e) {
    pfSetStatus("Could not read the selection: " + e.message, true);
    return false;
  } finally {
    btn.disabled = false;
  }
}

// Dates arrive as a serial number, so anything on a date format is taken as it
// is displayed. A blank cell is "" either way.
function pfCellValue(value, text, format) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number" && /[ymd]/i.test(String(format || ""))) return String(text || "").trim();
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return String(text || value).trim();
  return String(text || value).trim();
}

/* ---------- the dialog ---------- */

function pfOpen() {
  if (!supportsDialog()) {
    pfSetStatus("This version of Excel cannot open the finder window. Excel 2019 or Microsoft 365 is needed.", true);
    return;
  }
  if (pfDialog) { pfSetStatus("The finder is already open."); return; }

  $("pf-open").disabled = true;
  Office.context.ui.displayDialogAsync(PF_URL, { height: 88, width: 88, displayInIframe: false }, (res) => {
    $("pf-open").disabled = false;
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      pfSetStatus("Could not open the finder: " + res.error.message, true);
      return;
    }
    pfDialog = res.value;
    pfDialog.addEventHandler(Office.EventType.DialogMessageReceived, pfRead);
    pfDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      // 12006 is the user closing the window; anything else is a failure worth saying.
      pfDialog = null;
      pfSetStatus(arg.error === 12006 ? "Finder closed." : "Finder closed (" + arg.error + ").", arg.error !== 12006);
    });
    pfSetStatus("Finder open — pick a PDF in that window.");
  });
}

function supportsDialog() {
  try { return Office.context.requirements.isSetSupported("DialogApi", "1.2"); } catch { return false; }
}

function pfSend(msg) {
  if (!pfDialog) return;
  for (const chunk of PFWire.encode(msg)) {
    try { pfDialog.messageChild(chunk, { targetOrigin: window.location.origin }); }
    catch { try { pfDialog.messageChild(chunk); } catch { /* window gone */ } }
  }
}

const pfRead = PFWire.reader(pfHandle);

function pfHandle(msg) {
  if (msg.t === "ready") {
    pfSend({ t: "colour", hex: pfHex });
    if (pfState.cells.length) pfSend({ t: "rows", sheet: pfState.sheet, cells: pfState.cells });
    else pfReadSelection(false);
    return;
  }
  if (msg.t === "pull") { pfReadSelection(false); return; }
  if (msg.t === "colour") {
    pfHex = msg.hex;
    pfSwatch();
    try { localStorage.setItem(PF_TICK_KEY, JSON.stringify(pfHex)); } catch { /* quota */ }
    return;
  }
  if (msg.t === "goto") { pfGoto(msg.sheet, msg.ref); return; }
  if (msg.t === "ticks") { pfPaint(msg); return; }
}

/* ---------- colouring the sheet ---------- */

// Excel takes at most a handful of areas comfortably in one multi-range call.
const PF_AREAS_PER_CALL = 50;

/**
 * The finder sends the whole set of ticked cells rather than each change, so
 * this is a straight diff against what is already coloured: newly ticked cells
 * get the marker colour, released ones have their fill taken off, and a colour
 * change repaints every one of them.
 */
let pfBusy = false;
let pfQueued = null;

async function pfPaint(msg) {
  if (pfBusy) { pfQueued = msg; return; }
  pfBusy = true;
  try {
    const sheet = msg.sheet || pfState.sheet;
    const want = new Set(msg.refs || []);
    const hex = /^#[0-9a-f]{6}$/i.test(msg.hex || "") ? msg.hex : pfHex;
    const repaint = hex.toLowerCase() !== pfHex.toLowerCase() || pfPainted.sheet !== sheet;
    pfHex = hex;
    pfSwatch();

    const had = pfPainted.sheet === sheet ? new Set(pfPainted.refs) : new Set();
    const fill = [...want].filter((r) => repaint || !had.has(r));
    const strip = pfPainted.sheet === sheet
      ? pfPainted.refs.filter((r) => !want.has(r))
      : pfPainted.refs.slice();
    const stripSheet = pfPainted.sheet || sheet;

    if (fill.length || strip.length) {
      await Excel.run(async (ctx) => {
        if (strip.length) pfApply(ctx.workbook.worksheets.getItem(stripSheet), strip, null);
        if (fill.length) pfApply(ctx.workbook.worksheets.getItem(sheet), fill, hex);
        await ctx.sync();
      });
    }

    pfPainted = { sheet, refs: [...want] };
    pfSetStatus(`${msg.hit || 0} of ${msg.all || 0} ticked off.`);
  } catch (e) {
    pfSetStatus("Could not colour the sheet: " + e.message, true);
  } finally {
    pfBusy = false;
    const next = pfQueued;
    pfQueued = null;
    if (next) pfPaint(next);
  }
}

/** Fill (or clear, when `hex` is null) a list of single-cell refs. */
function pfApply(ws, refs, hex) {
  const act = (range) => { if (hex === null) range.format.fill.clear(); else range.format.fill.color = hex; };
  if (!supports("1.9")) {
    for (const r of refs) act(ws.getRange(r));
    return;
  }
  for (let k = 0; k < refs.length; k += PF_AREAS_PER_CALL) {
    act(ws.getRanges(refs.slice(k, k + PF_AREAS_PER_CALL).join(",")));
  }
}

/** Put the Excel cursor on one cell, switching sheets if it is on another. */
async function pfGoto(sheet, ref) {
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet || pfState.sheet);
      ws.activate();
      ws.getRange(ref).select();
      await ctx.sync();
    });
  } catch { /* the sheet was renamed or the cell is out of range */ }
}

/** Take the finder's colour off every cell it put it on. */
async function pfClear() {
  if (!pfPainted.refs.length) { pfSetStatus("Nothing to clear."); return; }
  $("pf-clear").disabled = true;
  try {
    await Excel.run(async (ctx) => {
      pfApply(ctx.workbook.worksheets.getItem(pfPainted.sheet), pfPainted.refs, null);
      await ctx.sync();
    });
    pfPainted = { sheet: pfPainted.sheet, refs: [] };
    pfSend({ t: "cleared" });
    pfSetStatus("Ticks cleared.");
  } catch (e) {
    pfSetStatus("Could not clear: " + e.message, true);
  } finally {
    $("pf-clear").disabled = false;
  }
}
