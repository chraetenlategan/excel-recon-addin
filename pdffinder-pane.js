"use strict";

/**
 * pdffinder-pane.js — the workbook's end of the PDF finder.
 *
 * The finder itself is a page (`pdffinder.html`) too big to live in a task
 * pane, so it opens as an Office dialog: a window of its own, beside Excel,
 * showing the statement and nothing else. This file is the other end.
 *
 * The pane is deliberately small. It holds one thing the finder cannot know —
 * **which cells to look for** — as a sheet and, optionally, the columns and
 * rows to narrow it to. That scope is read when the window opens, sent over,
 * and the pane then does one job for the rest of the session: colouring the
 * cells the auditor ticks off on the page.
 *
 * The only property it ever sets on a cell is its fill colour, the same promise
 * the compare tabs make. `pfPainted` is the exact list of cells it filled, so
 * releasing a tick takes the colour off that cell and nothing else.
 *
 * Loaded after taskpane.js, and it borrows that file's helpers (`colLetter`,
 * `colIndex`, `supports`) rather than repeating them.
 */

// What was sent to the finder, in the order it was sent.
const pfState = { sheet: "", scope: "", cells: [] };
// What is currently coloured on the sheet, so it can be taken off again.
let pfPainted = { sheet: "", refs: [] };
let pfHex = "#FFE94D";
let pfDialog = null;

const PF_TICK_KEY = "vdm-pdffinder:tick";

// A whole column of a large book is more than anyone reconciles against one
// statement, and every value has to be searched for on every page.
const PF_MAX_CELLS = 10000;

/**
 * Where the finder window is served from.
 *
 * Blank means **beside the pane** — the copy of `pdffinder.html` sitting next to
 * this file, on the add-in's own origin. That is the right answer nearly always:
 * the finder ships with the add-in, so wherever the pane came from the finder is
 * already there, at the same version, over the same HTTPS the pane is trusted on.
 *
 * Set it to another origin only to serve the finder from elsewhere, and mind the
 * two conditions Office puts on a dialog before doing so:
 *
 *  - **It must be HTTPS.** Office refuses a plain `http://` dialog outright
 *    (localhost is the only exemption). A LAN box needs a certificate the
 *    workstations trust — an internal CA cert, the same thing `serve.py` was
 *    written for — before it can host this.
 *  - **Cross-origin needs declaring.** A finder on a different origin from the
 *    pane can still message it, but only if that origin is listed in the
 *    manifest's `<AppDomains>` and both ends name each other in `targetOrigin`.
 *    The pane's origin rides in on the query string for `bridge.js` to answer to;
 *    the `<AppDomains>` entry has to be added by hand.
 *
 * That host must also carry the whole finder — `pdffinder.html`, `pdffinder.css`,
 * `pdffinder/`, `vendor/` and `assets/` — at this same version. A server holding
 * some other app answers with its own page or a 404, and the window opens on
 * nothing. If the configured home cannot be opened at all, the pane falls back to
 * the copy beside it and says so.
 */
const PF_BASE = "";

// Blank base means beside the pane, which is also where the fallback lives.
const pfUrl = (base) => {
  const u = new URL("pdffinder.html", base || window.location.href);
  u.searchParams.set("parent", window.location.origin);
  return u.href;
};

const PF_HOME = pfUrl(PF_BASE);
const PF_FALLBACK = pfUrl(window.location.href);
// The origin of whichever of the two actually opened — messageChild must name it.
let pfOrigin = new URL(PF_HOME).origin;

function initPdfFinder() {
  try {
    const saved = localStorage.getItem(PF_TICK_KEY);
    if (saved) pfHex = JSON.parse(saved);
  } catch { /* first run */ }

  $("pf-open").onclick = pfOpen;
}

/** Fill the sheet picker, keeping whatever was chosen if it is still there. */
function pfSetSheets(names) {
  const sel = $("pf-sheet");
  const keep = sel.value;
  sel.innerHTML = "";
  for (const n of names) sel.appendChild(new Option(n, n));
  sel.value = names.includes(keep) ? keep : (names[0] || "");
}

function pfSetStatus(msg, isError) {
  const box = $("pf-status");
  box.textContent = msg || "";
  box.classList.toggle("err", !!isError);
}

/* ---------- the scope ---------- */

/** "C", "A,B", "C:E" -> column indexes. Blank means every used column. */
function pfColumns(raw) {
  const text = String(raw || "").trim();
  if (!text) return null;
  const out = new Set();
  for (const token of text.split(",").map((t) => t.trim()).filter(Boolean)) {
    const m = token.match(/^([A-Za-z]{1,3})(?:\s*[:-]\s*([A-Za-z]{1,3}))?$/);
    if (!m) throw new Error(`“${token}” is not a column like C or C:E.`);
    const a = colIndex(m[1].toUpperCase());
    const b = colIndex((m[2] || m[1]).toUpperCase());
    for (let i = Math.min(a, b); i <= Math.max(a, b); i++) out.add(i);
  }
  return [...out].sort((x, y) => x - y);
}

/** "12:250" -> those rows. Blank means every used row. */
function pfRows(raw) {
  const text = String(raw || "").trim();
  if (!text) return null;
  const m = text.match(/^(\d+)\s*[:-]\s*(\d+)$/);
  if (!m) throw new Error(`“${text}” is not a row range like 12:250.`);
  const first = Math.min(+m[1], +m[2]), last = Math.max(+m[1], +m[2]);
  if (first < 1) throw new Error("rows start at 1.");
  return { first, last };
}

/** How the scope reads in the finder's header: "Bank · A, B · 340 cells". */
function pfLabel(sheet, cols, rows, n) {
  const parts = [sheet];
  parts.push(cols ? cols.map(colLetter).join(", ") : "every column");
  if (rows) parts.push("rows " + rows.first + "–" + rows.last);
  parts.push(n + (n === 1 ? " cell" : " cells"));
  return parts.join(" · ");
}

/**
 * The cells to look for, read straight off the sheet — no selection needed, and
 * none disturbed. The constraints in the pane are the whole of the user's say
 * over what is searched for.
 *
 * Values come off the sheet as `text` — what the auditor sees, which is what is
 * printed on the statement — except for plain numbers, where the underlying
 * value is used instead: a column displayed to the rand still has to find its
 * cents on the page.
 */
async function pfReadScope(announce) {
  try {
    const sheet = $("pf-sheet").value;
    if (!sheet) throw new Error("pick a sheet.");
    const cols = pfColumns($("pf-cols").value);
    const rows = pfRows($("pf-rows").value);

    const cells = [];
    let capped = false;
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const used = ws.getUsedRangeOrNullObject(true);
      used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
      await ctx.sync();
      if (used.isNullObject) throw new Error(`“${sheet}” is empty.`);

      // Never read past the data: a column named on its own still stops at the
      // last used row.
      const usedFirst = used.rowIndex + 1, usedLast = used.rowIndex + used.rowCount;
      const first = Math.max(rows ? rows.first : usedFirst, 1);
      const last = Math.min(rows ? rows.last : usedLast, usedLast);
      if (last < first) throw new Error("those rows hold no data.");

      const usedCols = [];
      for (let i = 0; i < used.columnCount; i++) usedCols.push(used.columnIndex + i);
      const want = (cols || usedCols).filter((c) => usedCols.includes(c));
      if (!want.length) throw new Error("those columns hold no data.");

      const blocks = [];
      for (const c of want) {
        const letter = colLetter(c);
        const range = ws.getRange(`${letter}${first}:${letter}${last}`);
        range.load(["values", "text", "numberFormat"]);
        blocks.push({ letter, range });
      }
      await ctx.sync();

      for (const b of blocks) {
        const vals = b.range.values, texts = b.range.text, fmts = b.range.numberFormat;
        for (let i = 0; i < vals.length; i++) {
          const v = pfCellValue(vals[i][0], texts[i][0], fmts[i][0]);
          if (v === "") continue;
          if (cells.length >= PF_MAX_CELLS) { capped = true; break; }
          cells.push({ ref: b.letter + (first + i), v });
        }
        if (capped) break;
      }
    });

    if (!cells.length) throw new Error("every cell in that scope is blank.");
    // in reading order, so the page and the sheet are walked the same way
    cells.sort((a, b) => (parseInt(a.ref.replace(/\D/g, ""), 10) - parseInt(b.ref.replace(/\D/g, ""), 10)));

    pfState.sheet = sheet;
    pfState.scope = pfLabel(sheet, cols, rows, cells.length);
    pfState.cells = cells;
    pfSend({ t: "rows", sheet, scope: pfState.scope, cells });
    if (announce) {
      pfSetStatus(capped
        ? `First ${cells.length} cells sent — narrow the columns or rows for the rest.`
        : `${pfState.scope}.`);
    }
    return true;
  } catch (e) {
    pfSetStatus("Could not read those cells: " + e.message, true);
    return false;
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

async function pfOpen() {
  if (!supportsDialog()) {
    pfSetStatus("This version of Excel cannot open the finder window. Excel 2019 or Microsoft 365 is needed.", true);
    return;
  }
  if (pfDialog) { pfSetStatus("The finder is already open."); return; }

  $("pf-open").disabled = true;
  // The scope is settled before the window opens, so a mistyped column is said
  // here rather than by an empty page over there.
  if (!(await pfReadScope(true))) { $("pf-open").disabled = false; return; }
  pfTry(PF_HOME, () => pfTry(PF_FALLBACK, null));
}

/** Open one candidate URL, handing off to `next` (if any) when it will not open. */
function pfTry(url, next) {
  Office.context.ui.displayDialogAsync(url, { height: 88, width: 88, displayInIframe: false }, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      if (next) { next(); return; }
      $("pf-open").disabled = false;
      pfSetStatus("Could not open the finder: " + res.error.message, true);
      return;
    }
    $("pf-open").disabled = false;
    pfOrigin = new URL(url).origin;
    if (url === PF_FALLBACK && PF_HOME !== PF_FALLBACK) {
      pfSetStatus(PF_BASE + " could not be opened — using the copy beside the pane.");
    }
    pfDialog = res.value;
    pfDialog.addEventHandler(Office.EventType.DialogMessageReceived, pfRead);
    pfDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      // 12006 is the user closing the window; anything else is a failure worth saying.
      pfDialog = null;
      pfSetStatus(arg.error === 12006 ? "Finder closed." : "Finder closed (" + arg.error + ").", arg.error !== 12006);
    });
  });
}

function supportsDialog() {
  try { return Office.context.requirements.isSetSupported("DialogApi", "1.2"); } catch { return false; }
}

function pfSend(msg) {
  if (!pfDialog) return;
  for (const chunk of PFWire.encode(msg)) {
    try { pfDialog.messageChild(chunk, { targetOrigin: pfOrigin }); }
    catch { try { pfDialog.messageChild(chunk); } catch { /* window gone */ } }
  }
}

const pfRead = PFWire.reader(pfHandle);

function pfHandle(msg) {
  if (msg.t === "ready") {
    pfSend({ t: "colour", hex: pfHex });
    if (pfState.cells.length) pfSend({ t: "rows", sheet: pfState.sheet, scope: pfState.scope, cells: pfState.cells });
    else pfReadScope(false);
    return;
  }
  if (msg.t === "pull") { pfReadScope(false); return; }
  if (msg.t === "colour") {
    pfHex = msg.hex;
    try { localStorage.setItem(PF_TICK_KEY, JSON.stringify(pfHex)); } catch { /* quota */ }
    return;
  }
  if (msg.t === "goto") { pfGoto(msg.sheet, msg.ref); return; }
  if (msg.t === "find") { pfFind(msg); return; }
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

/* ---------- hunting a printed value down on the sheet ---------- */

// A value double-clicked on the page need not be one of the cells in scope, so
// this looks through the whole workbook. Enough matches to point at, not enough
// to take all afternoon selecting.
const PF_FIND_MAX = 200;

/** The pane's own copy of the finder's number parser — 1 234,56 / (1.234,56) / 12.34- */
function pfToNumber(raw) {
  if (raw === null || raw === undefined) return null;
  let s = String(raw).trim().replace(/[R$€£¥]/gi, "").replace(/\s| |'/g, "");
  if (!s) return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  if (s.startsWith("-")) { neg = !neg; s = s.slice(1); }
  if (s.endsWith("-")) { neg = !neg; s = s.slice(0, -1); }
  if (s.endsWith("%")) return null;
  if (!/^[\d.,]+$/.test(s) || !/\d/.test(s)) return null;
  const last = Math.max(s.lastIndexOf("."), s.lastIndexOf(","));
  let intPart = s, decPart = "";
  if (last > -1 && /^\d{1,2}$/.test(s.slice(last + 1))) { intPart = s.slice(0, last); decPart = s.slice(last + 1); }
  intPart = intPart.replace(/[.,]/g, "");
  if (!/^\d*$/.test(intPart) || !/^\d*$/.test(decPart)) return null;
  if (intPart === "" && decPart === "") return null;
  const n = Number((intPart || "0") + (decPart ? "." + decPart : ""));
  if (!isFinite(n)) return null;
  return neg ? -n : n;
}

const pfNorm = (raw) => String(raw === null || raw === undefined ? "" : raw).toLowerCase().replace(/[^a-z0-9]/g, "");

/** Every cell of one used range that says the same thing as the printed value. */
function pfScan(used, want, wantText) {
  const refs = [];
  const vals = used.values, texts = used.text;
  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      const raw = vals[r][c];
      if (raw === null || raw === undefined || raw === "") continue;
      const shown = String(texts[r][c] === undefined ? raw : texts[r][c]);
      let ok;
      if (want !== null) {
        // an amount matches on its value, however either side prints it
        const n = typeof raw === "number" ? raw : pfToNumber(shown);
        ok = n !== null && Math.abs(n - want) < 0.005;
      } else {
        ok = wantText.length > 0 && pfNorm(shown) === wantText;
      }
      if (ok) {
        refs.push(colLetter(used.columnIndex + c) + (used.rowIndex + r + 1));
        if (refs.length >= PF_FIND_MAX) return refs;
      }
    }
  }
  return refs;
}

/**
 * Find a value printed on the PDF somewhere in the workbook and put the Excel
 * selection on it — the scope's own sheet first, then the rest, so a figure on
 * the statement that falls outside the chosen columns can still be pointed at.
 *
 * Only the selection moves. Nothing is coloured, so nothing has to be undone.
 */
async function pfFind(msg) {
  const v = String((msg && msg.v) || "").trim();
  if (!v) return;
  const want = pfToNumber(v);
  const wantText = pfNorm(v);
  if (want === null && !wantText) { pfSend({ t: "found", msg: "Nothing to look for." }); return; }

  try {
    const first = msg.sheet || pfState.sheet;
    let hit = null;
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();

      const names = sheets.items.map((w) => w.name);
      names.sort((a, b) => (b === first ? 1 : 0) - (a === first ? 1 : 0));

      for (const name of names) {
        const ws = sheets.getItem(name);
        const used = ws.getUsedRangeOrNullObject(true);
        used.load(["rowIndex", "columnIndex", "values", "text"]);
        await ctx.sync();
        if (used.isNullObject) continue;

        const refs = pfScan(used, want, wantText);
        if (!refs.length) continue;

        ws.activate();
        pfPick(ws, refs);
        await ctx.sync();
        hit = { sheet: name, refs };
        break;
      }
    });

    if (!hit) {
      pfSetStatus(v + " is not on any sheet.");
      pfSend({ t: "found", msg: v + " — not found in this workbook." });
      return;
    }
    const many = hit.refs.length > 1 ? " (" + hit.refs.length + " cells)" : "";
    pfSetStatus("Found " + v + " at " + hit.sheet + "!" + hit.refs[0] + many + ".");
    pfSend({ t: "found", msg: v + "  →  " + hit.sheet + "!" + hit.refs[0] + many });
  } catch (e) {
    pfSetStatus("Could not look for " + v + ": " + e.message, true);
    pfSend({ t: "found", msg: "Could not look for " + v + "." });
  }
}

/** Put the selection on a list of single-cell refs — all of them where Excel can. */
function pfPick(ws, refs) {
  if (supports("1.9") && refs.length > 1) {
    ws.getRanges(refs.slice(0, PF_AREAS_PER_CALL).join(",")).select();
    return;
  }
  ws.getRange(refs[0]).select();
}
