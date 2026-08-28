"use strict";

/**
 * taskpane.js — generic column comparison.
 *
 * Pick any sheet and any column(s) on each side, optionally limit the rows, and
 * every value on side A is matched against the values on side B. Matches go
 * green, non-matches go red — on the user's own cells.
 *
 * Either/or: a side can carry several columns. They are pooled, so a value in
 * one column on A counts as found if it turns up in *any* of side B's columns
 * (and the same the other way round).
 *
 * Nothing else about the workbook is touched: no values are written, no cells
 * are merged or unmerged, no fonts, borders, number formats or column widths
 * are changed. The only property this file ever sets is a cell's fill colour
 * (and "Clear colours" only ever clears that same fill).
 */

const $ = (id) => document.getElementById(id);

const GREEN = "C6EFCE";
const RED = "FFC7CE";

// Per side: what the user picked plus the used-range geometry of that sheet.
// `columns` is a list of 0-based column indexes — several means either/or.
const state = {
  a: { sheet: "", columns: [], limit: "", used: null },
  b: { sheet: "", columns: [], limit: "", used: null },
};

// The ranges the last Compare painted, so "Clear colours" can undo exactly them.
let lastPainted = [];

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in only runs in Excel.", true);
    return;
  }
  for (const side of ["a", "b"]) {
    el(side, "sheet").onchange = () => { state[side].sheet = el(side, "sheet").value; loadColumns(side); };
    el(side, "column").onchange = () => { state[side].columns = pickedColumns(side); updatePreview(side); };
    el(side, "limit").oninput = () => { state[side].limit = el(side, "limit").value; updatePreview(side); };
    el(side, "use-selection").onclick = () => useSelection(side);
  }
  $("refresh-sheets").onclick = loadSheetList;
  $("compare").onclick = compare;
  $("clear-colours").onclick = clearColours;
  initColourRecon();
  loadSheetList();
});

const el = (side, key) => document.querySelector(`#side-${side} [data-k="${key}"]`);
const pickedColumns = (side) =>
  [...el(side, "column").selectedOptions].map((o) => parseInt(o.value, 10)).sort((x, y) => x - y);

function setStatus(msg, isError) {
  const box = $("status");
  box.textContent = msg || "";
  box.classList.toggle("err", !!isError);
}

/* ---------- sheet + column pickers ---------- */

async function loadSheetList() {
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      const names = sheets.items.map((s) => s.name);
      for (const side of ["a", "b"]) {
        const sel = el(side, "sheet");
        const keep = state[side].sheet;
        sel.innerHTML = "";
        for (const n of names) sel.appendChild(new Option(n, n));
        // Default the two sides to different sheets where the book has two.
        const fallback = side === "a" ? names[0] : (names[1] || names[0]);
        sel.value = names.includes(keep) ? keep : (fallback || "");
        state[side].sheet = sel.value;
      }
      crSetSheets(names);
    });
    for (const side of ["a", "b"]) await loadColumns(side);
    setStatus("");
  } catch (e) {
    setStatus("Could not read worksheets: " + e.message, true);
  }
}

// Fill the column picker from the sheet's used range, labelling each column
// with whatever sits in its first used row ("E — Amount").
async function loadColumns(side) {
  const s = state[side];
  const sel = el(side, "column");
  if (!s.sheet) { sel.innerHTML = ""; s.used = null; updatePreview(side); return; }
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(s.sheet);
      const used = ws.getUsedRangeOrNullObject(true);
      used.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values"]);
      await ctx.sync();
      if (used.isNullObject) { s.used = null; sel.innerHTML = ""; return; }

      s.used = {
        row: used.rowIndex, col: used.columnIndex,
        rows: used.rowCount, cols: used.columnCount,
      };
      const header = used.values[0] || [];
      const last = s.used.col + s.used.cols - 1;
      const keep = s.columns.filter((c) => c >= s.used.col && c <= last);
      sel.innerHTML = "";
      for (let i = 0; i < s.used.cols; i++) {
        const index = s.used.col + i;
        const label = _text(header[i]);
        sel.appendChild(new Option(colLetter(index) + (label ? " — " + label : ""), String(index)));
      }
      s.columns = keep.length ? keep : [s.used.col];
      for (const opt of sel.options) opt.selected = s.columns.includes(parseInt(opt.value, 10));
    });
  } catch (e) {
    setStatus("Could not read “" + s.sheet + "”: " + e.message, true);
  }
  updatePreview(side);
}

/* ---------- "Use selected cells" ---------- */

/**
 * Fill one side in from whatever the user has highlighted in Excel: highlight
 * the cells for side A, press its button, highlight the cells for side B,
 * press its button, then Compare.
 *
 * The selection sets the sheet, the column picker and the limit box, so the
 * result is exactly as if it had been typed in by hand — and can still be
 * edited afterwards. A multi-column or multi-area selection becomes an
 * either/or pool, the same as ctrl-clicking several columns.
 */
async function useSelection(side) {
  const btn = el(side, "use-selection");
  btn.disabled = true;
  try {
    let address = "";
    await Excel.run(async (ctx) => {
      // getSelectedRanges() carries every area of a ctrl-click selection;
      // older hosts only offer the single-area getSelectedRange().
      const sel = supports("1.9")
        ? ctx.workbook.getSelectedRanges()
        : ctx.workbook.getSelectedRange();
      sel.load("address");
      await ctx.sync();
      address = sel.address;
    });

    const { sheet, areas } = parseSelection(address);
    if (!areas.length) throw new Error("select some cells in Excel first.");

    const sel = el(side, "sheet");
    if (sheet && ![...sel.options].some((o) => o.value === sheet)) await loadSheetList();
    if (sheet) { sel.value = sheet; state[side].sheet = sheet; }
    await loadColumns(side);

    const columns = [...new Set(areas.map((a) => a.column))].sort((x, y) => x - y);
    state[side].columns = columns;
    for (const opt of el(side, "column").options) opt.selected = columns.includes(parseInt(opt.value, 10));

    const limit = areas.map((a) => selectionToken(a, state[side].used)).join(", ");
    el(side, "limit").value = limit;
    state[side].limit = limit;
    updatePreview(side);
    setStatus(`Side ${side.toUpperCase()} set to ${state[side].sheet}!${limit}`);
  } catch (e) {
    setStatus("Could not use the selection: " + e.message, true);
  } finally {
    btn.disabled = false;
  }
}

/**
 * Split a selection address — "Sheet1!B12:B25", or several areas as
 * "'My Book'!B2:D9,'My Book'!F2:F9" — into a sheet name plus one
 * { column, firstRow, lastRow } per column covered. Rows are null where the
 * whole column is selected.
 */
function parseSelection(address) {
  let sheet = "";
  const areas = [];
  for (const piece of splitAreas(String(address || ""))) {
    let ref = piece;
    const bang = ref.lastIndexOf("!");
    if (bang >= 0) {
      let name = ref.slice(0, bang);
      ref = ref.slice(bang + 1);
      if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
      if (!sheet) sheet = name;
    }
    areas.push(...refToAreas(ref.replace(/\$/g, "").trim().toUpperCase()));
  }
  return { sheet, areas };
}

// Areas are comma-separated, but a quoted sheet name may itself hold a comma
// ('Jan, Feb'!A1:A9), so only commas outside the quotes split.
function splitAreas(address) {
  const out = [];
  let cur = "", quoted = false;
  for (const ch of address) {
    if (ch === "'") quoted = !quoted;
    if (ch === "," && !quoted) { out.push(cur); cur = ""; continue; }
    cur += ch;
  }
  out.push(cur);
  return out.map((t) => t.trim()).filter(Boolean);
}

// One A1 area -> one entry per column it covers, so a block like "B2:D9"
// becomes B, C and D over rows 2-9.
function refToAreas(ref) {
  const [start, end = start] = ref.split(":");
  const s = start.match(/^([A-Z]{1,3})?(\d+)?$/);
  const e = end.match(/^([A-Z]{1,3})?(\d+)?$/);
  if (!s || !e) return [];
  const [, c1, r1] = s;
  const [, c2, r2] = e;
  if (!c1 || !c2) return [];                       // whole rows carry no column
  const rows = r1 && r2
    ? { firstRow: Math.min(+r1, +r2), lastRow: Math.max(+r1, +r2) }
    : { firstRow: null, lastRow: null };
  const out = [];
  for (let c = Math.min(colIndex(c1), colIndex(c2)); c <= Math.max(colIndex(c1), colIndex(c2)); c++) {
    out.push({ column: c, ...rows });
  }
  return out;
}

/**
 * Render one selected area as a limit-box token, trimmed to the sheet's used
 * range — clicking a column header selects a million rows, and there's no
 * sense reading past the data.
 */
function selectionToken(area, used) {
  const letter = colLetter(area.column);
  let { firstRow, lastRow } = area;
  if (firstRow === null) return letter;
  if (used) {
    firstRow = Math.max(firstRow, used.row + 1);
    lastRow = Math.min(lastRow, used.row + used.rows);
    if (lastRow < firstRow) return letter;         // selection misses the data
  }
  return `${letter}${firstRow}:${letter}${lastRow}`;
}

/* ---------- the row limiter ---------- */

/**
 * Turn a side into { sheet, parts: [{ column, firstRow, lastRow }] } — one part
 * per column being read, all pooled together when matching.
 *
 * The limit box is a comma-separated list of range pieces, each of which may
 * name a column, rows, or both:
 *   ""                   whole used rows of every picked column
 *   "12:25"              rows 12-25 of every picked column
 *   "B12:B25"            that one range (the column overrides the picker)
 *   "F12:F25, H12:H25"   either/or across two ranges
 *   "F, H"               those two whole columns
 */
function resolveSide(side) {
  const s = state[side];
  const where = `Side ${side.toUpperCase()}`;
  if (!s.sheet) throw new Error("Pick a sheet on both sides.");
  if (!s.used) throw new Error(`“${s.sheet}” looks empty.`);

  const picked = s.columns.length ? s.columns : [s.used.col];
  const usedFirst = s.used.row + 1;
  const usedLast = s.used.row + s.used.rows;

  const raw = (s.limit || "").trim().toUpperCase().replace(/\$/g, "");
  if (!raw) {
    return { sheet: s.sheet, parts: picked.map((column) => ({ column, firstRow: usedFirst, lastRow: usedLast })) };
  }

  const parts = [];
  for (const token of raw.split(",").map((t) => t.trim()).filter(Boolean)) {
    const m = token.match(/^([A-Z]{1,3})?(\d+)?\s*:\s*([A-Z]{1,3})?(\d+)?$/) || token.match(/^([A-Z]{1,3})()()()$/);
    if (!m) throw new Error(`${where}: “${token}” is not a range like B12:B25.`);
    const [, c1, r1, , r2] = m;
    let firstRow = usedFirst, lastRow = usedLast;
    if (r1 && r2) { firstRow = Math.min(+r1, +r2); lastRow = Math.max(+r1, +r2); }
    else if (r1 || r2) throw new Error(`${where}: give both ends, e.g. B12:B25.`);
    if (lastRow < firstRow) throw new Error(`${where}: “${token}” has no rows.`);
    // A piece that names a column is that column; one that only gives rows
    // applies those rows to every column picked above.
    for (const column of (c1 ? [colIndex(c1)] : picked)) parts.push({ column, firstRow, lastRow });
  }
  if (!parts.length) throw new Error(`${where}: nothing to read.`);
  return { sheet: s.sheet, parts };
}

// "B" -> 1, "AA" -> 26. The inverse of colLetter().
function colIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

const addressOf = (p) => `${colLetter(p.column)}${p.firstRow}:${colLetter(p.column)}${p.lastRow}`;
const cellCount = (side) => side.parts.reduce((n, p) => n + (p.lastRow - p.firstRow + 1), 0);

function updatePreview(side) {
  const box = el(side, "preview");
  try {
    const r = resolveSide(side);
    box.textContent = `${r.sheet}!${r.parts.map(addressOf).join(", ")}  ·  ${cellCount(r)} cells`;
    box.classList.remove("err");
  } catch (e) {
    box.textContent = state[side].sheet ? e.message : "";
    box.classList.add("err");
  }
}

/* ---------- matching ---------- */

/**
 * Every matching rule the pane offers, read off the checkboxes once per run and
 * handed to keyOf()/matchValues(). Both the column compare and the colour recon
 * use the same set, so a rule you tick applies to whichever you press.
 */
function readOpts() {
  const num = (id) => {
    const v = parseFloat($(id).value);
    return isFinite(v) && v > 0 ? v : 0;
  };
  return {
    ignoreSign: $("ignore-sign").checked,
    ignoreCase: $("ignore-case").checked,
    roundWhole: $("round-whole").checked,
    ignorePunct: $("ignore-punct").checked,
    ignoreZeros: $("ignore-zeros").checked,
    tolerance: num("tolerance"),
    firstChars: Math.round(num("first-chars")),
  };
}

/**
 * The comparison key for one cell. Numbers compare as numbers (in cents, so
 * 100 and 100.00 are one value); anything else compares as text. Blank cells
 * return null and are skipped entirely — never coloured, never matched.
 *
 * The rules bend the key before it is compared: dropping the sign, the cents,
 * the case, the punctuation, the leading zeros, or everything past the first
 * few characters. A tolerance is *not* a key rule — near numbers can't share a
 * key — and is handled by matchValues() instead.
 */
function keyOf(value, opts) {
  const n = _amount(value);
  if (n !== null) {
    let v = opts.ignoreSign ? Math.abs(n) : n;
    if (opts.roundWhole) v = Math.round(v);
    return "n:" + Math.round(v * 100);
  }
  let t = _text(value);
  if (!t) return null;
  if (opts.ignoreCase) t = t.toLowerCase().replace(/\s+/g, " ");
  if (opts.ignorePunct) t = t.replace(/[^\p{L}\p{N}]+/gu, "");
  // "INV-00123" -> "INV-123": zeros that open a run of digits, not zeros
  // sitting inside a number ("1004" is left alone).
  if (opts.ignoreZeros) t = t.replace(/(^|[^\p{N}])0+(?=\p{N})/gu, "$1");
  t = t.trim();
  if (!t) return null;
  if (opts.firstChars) t = t.slice(0, opts.firstChars);
  return "t:" + t;
}

const _cents = (key) => (key !== null && key.startsWith("n:") ? parseInt(key.slice(2), 10) : null);

/**
 * Pair the two sides off one-for-one: a value on A claims one equal value from
 * B's pool, so three 100s on A against two on B leave the third 100 red. Where
 * a side spans several columns (or several colours) the pool spans them too —
 * that's the either/or.
 *
 * Returns a boolean per cell on each side (null = blank, leave alone) plus
 * pairA: for each A cell, the index of the B cell it claimed, or -1.
 */
function matchValues(valuesA, valuesB, opts) {
  const keysA = valuesA.map((v) => keyOf(v, opts));
  const keysB = valuesB.map((v) => keyOf(v, opts));
  const tol = Math.round((opts.tolerance || 0) * 100);

  const hitA = keysA.map((k) => (k === null ? null : false));
  const hitB = keysB.map((k) => (k === null ? null : false));
  const pairA = keysA.map(() => -1);
  const claim = (i, j) => { hitA[i] = true; hitB[j] = true; pairA[i] = j; };

  // Exact keys go through a pool, one queue per key. With a tolerance set the
  // numbers are held back for the sweep below and only text pools here.
  const pool = new Map();
  keysB.forEach((k, j) => {
    if (k === null) return;
    if (tol > 0 && _cents(k) !== null) return;
    if (!pool.has(k)) pool.set(k, []);
    pool.get(k).push(j);
  });
  keysA.forEach((k, i) => {
    if (k === null) return;
    if (tol > 0 && _cents(k) !== null) return;
    const queue = pool.get(k);
    if (queue && queue.length) claim(i, queue.shift());
  });

  // Tolerance: an amount matches anything within ± of it. Both sides are walked
  // in ascending order and each A takes the smallest B still in reach, which
  // pairs off as many rows as can be paired.
  if (tol > 0) {
    const numbered = (keys) => keys
      .map((k, i) => ({ i, cents: _cents(k) }))
      .filter((e) => e.cents !== null)
      .sort((x, y) => x.cents - y.cents);
    const numA = numbered(keysA);
    const numB = numbered(keysB);
    let j = 0;
    for (const a of numA) {
      while (j < numB.length && numB[j].cents < a.cents - tol) j++;
      if (j < numB.length && numB[j].cents <= a.cents + tol) claim(a.i, numB[j++].i);
    }
  }

  return { hitA, hitB, pairA };
}

/* ---------- compare + paint ---------- */

async function compare() {
  let sa, sb;
  try {
    sa = resolveSide("a");
    sb = resolveSide("b");
  } catch (e) {
    setStatus(e.message, true);
    return;
  }
  const opts = readOpts();

  $("compare").disabled = true;
  setStatus("Reading…");
  try {
    await Excel.run(async (ctx) => {
      const wsA = ctx.workbook.worksheets.getItem(sa.sheet);
      const wsB = ctx.workbook.worksheets.getItem(sb.sheet);
      const rangesA = sa.parts.map((p) => wsA.getRange(addressOf(p)));
      const rangesB = sb.parts.map((p) => wsB.getRange(addressOf(p)));
      for (const r of [...rangesA, ...rangesB]) r.load("values");
      await ctx.sync();

      // Every column of a side flattens into one pool, in part order; the hits
      // come back in that same order and are handed back out to the parts.
      const flatten = (ranges) => [].concat(...ranges.map((r) => r.values.map((row) => row[0])));
      const { hitA, hitB } = matchValues(flatten(rangesA), flatten(rangesB), opts);

      setStatus("Colouring…");
      paintSide(wsA, sa, hitA);
      paintSide(wsB, sb, hitB);
      await ctx.sync();
    });

    lastPainted = [sa, sb];
    setStatus("Done.");
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    $("compare").disabled = false;
  }
}

// Excel takes at most a handful of areas comfortably in one multi-range call.
const AREAS_PER_CALL = 50;

/**
 * Fill the cells of one side. Consecutive rows of the same colour are merged
 * into runs, and the runs of a colour go in as one multi-area call where the
 * host supports it, so a thousand-row column is a few operations.
 */
function paintSide(ws, side, hits) {
  const runs = { [GREEN]: [], [RED]: [] };
  let offset = 0;
  for (const p of side.parts) {
    const n = p.lastRow - p.firstRow + 1;
    const mine = hits.slice(offset, offset + n);
    offset += n;
    const letter = colLetter(p.column);
    let i = 0;
    while (i < mine.length) {
      const hit = mine[i];
      if (hit === null) { i++; continue; }
      let j = i;
      while (j + 1 < mine.length && mine[j + 1] === hit) j++;
      runs[hit ? GREEN : RED].push(`${letter}${p.firstRow + i}:${letter}${p.firstRow + j}`);
      i = j + 1;
    }
  }

  const multi = supports("1.9");
  for (const colour of [GREEN, RED]) {
    const addrs = runs[colour];
    if (!addrs.length) continue;
    if (!multi) {
      for (const a of addrs) ws.getRange(a).format.fill.color = "#" + colour;
      continue;
    }
    for (let k = 0; k < addrs.length; k += AREAS_PER_CALL) {
      ws.getRanges(addrs.slice(k, k + AREAS_PER_CALL).join(",")).format.fill.color = "#" + colour;
    }
  }
}

function supports(version) {
  try { return Office.context.requirements.isSetSupported("ExcelApi", version); } catch { return false; }
}

// Undo: clear the fill on exactly the ranges the last Compare painted (or, if
// there hasn't been one this session, on the ranges currently picked).
async function clearColours() {
  let sides = lastPainted;
  if (!sides.length) {
    try { sides = [resolveSide("a"), resolveSide("b")]; }
    catch (e) { setStatus(e.message, true); return; }
  }
  $("clear-colours").disabled = true;
  setStatus("Clearing…");
  try {
    await Excel.run(async (ctx) => {
      for (const side of sides) {
        const ws = ctx.workbook.worksheets.getItem(side.sheet);
        for (const p of side.parts) ws.getRange(addressOf(p)).format.fill.clear();
      }
      await ctx.sync();
    });
    setStatus("Colours cleared.");
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    $("clear-colours").disabled = false;
  }
}
