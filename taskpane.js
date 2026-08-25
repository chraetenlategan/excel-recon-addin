"use strict";

/**
 * taskpane.js — generic two-column comparison.
 *
 * Pick any two sheets, any column on each, optionally limit the rows, and every
 * value in column A is matched against the values in column B. Matches go
 * green, non-matches go red — on the user's own cells.
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
const state = {
  a: { sheet: "", column: null, limit: "", used: null },
  b: { sheet: "", column: null, limit: "", used: null },
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
    el(side, "column").onchange = () => { state[side].column = intOrNull(el(side, "column").value); updatePreview(side); };
    el(side, "limit").oninput = () => { state[side].limit = el(side, "limit").value; updatePreview(side); };
  }
  $("refresh-sheets").onclick = loadSheetList;
  $("compare").onclick = compare;
  $("clear-colours").onclick = clearColours;
  loadSheetList();
});

const el = (side, key) => document.querySelector(`#side-${side} [data-k="${key}"]`);
const intOrNull = (v) => (v === "" || v === null || v === undefined ? null : parseInt(v, 10));

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
      const keep = s.column;
      sel.innerHTML = "";
      for (let i = 0; i < s.used.cols; i++) {
        const index = s.used.col + i;
        const label = _text(header[i]);
        sel.appendChild(new Option(colLetter(index) + (label ? " — " + label : ""), String(index)));
      }
      s.column = (keep !== null && keep >= s.used.col && keep < s.used.col + s.used.cols) ? keep : s.used.col;
      sel.value = String(s.column);
    });
  } catch (e) {
    setStatus("Could not read “" + s.sheet + "”: " + e.message, true);
  }
  updatePreview(side);
}

/* ---------- the row limiter ---------- */

/**
 * Turn the limit box into { column, firstRow, lastRow } (all 1-based rows,
 * 0-based column), falling back to the picked column and the sheet's used rows.
 * Accepts "B12:B25", "12:25", "B:B", "B" or "" (blank = whole used column).
 */
function resolveRange(side) {
  const s = state[side];
  if (!s.sheet) throw new Error("Pick a sheet on both sides.");
  if (!s.used) throw new Error(`“${s.sheet}” looks empty.`);

  let column = s.column === null ? s.used.col : s.column;
  let firstRow = s.used.row + 1;
  let lastRow = s.used.row + s.used.rows;

  const raw = (s.limit || "").trim().toUpperCase().replace(/\$/g, "");
  if (raw) {
    const m = raw.match(/^([A-Z]{1,3})?(\d+)?\s*:\s*([A-Z]{1,3})?(\d+)?$/) || raw.match(/^([A-Z]{1,3})()()()$/);
    if (!m) throw new Error(`Side ${side.toUpperCase()}: “${s.limit}” is not a range like B12:B25.`);
    const [, c1, r1, , r2] = m;
    if (c1) column = colIndex(c1);
    if (r1 && r2) { firstRow = Math.min(+r1, +r2); lastRow = Math.max(+r1, +r2); }
    else if (r1 || r2) throw new Error(`Side ${side.toUpperCase()}: give both ends, e.g. B12:B25.`);
  }
  if (lastRow < firstRow) throw new Error(`Side ${side.toUpperCase()}: that range has no rows.`);
  return { sheet: s.sheet, column, firstRow, lastRow };
}

// "B" -> 1, "AA" -> 26. The inverse of colLetter().
function colIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

const addressOf = (r) => `${colLetter(r.column)}${r.firstRow}:${colLetter(r.column)}${r.lastRow}`;

function updatePreview(side) {
  const box = el(side, "preview");
  try {
    const r = resolveRange(side);
    box.textContent = `${r.sheet}!${addressOf(r)}  ·  ${r.lastRow - r.firstRow + 1} rows`;
    box.classList.remove("err");
  } catch (e) {
    box.textContent = state[side].sheet ? e.message : "";
    box.classList.add("err");
  }
}

/* ---------- matching ---------- */

/**
 * The comparison key for one cell. Numbers compare as numbers (rounded to
 * cents, optionally unsigned); anything else compares as text. Blank cells
 * return null and are skipped entirely — never coloured, never matched.
 */
function keyOf(value, opts) {
  const n = _amount(value);
  if (n !== null) {
    const v = opts.ignoreSign ? Math.abs(n) : n;
    return "n:" + Math.round(v * 100);
  }
  let t = _text(value);
  if (!t) return null;
  if (opts.ignoreCase) t = t.toLowerCase().replace(/\s+/g, " ");
  return "t:" + t;
}

/**
 * Pair the two columns off one-for-one: a value on A claims one equal value on
 * B, so three 100s on A against two on B leave the third 100 red. Returns a
 * boolean per cell on each side (null = blank, leave alone).
 */
function matchColumns(valuesA, valuesB, opts) {
  const keysA = valuesA.map((v) => keyOf(v, opts));
  const keysB = valuesB.map((v) => keyOf(v, opts));

  const pool = new Map();
  keysB.forEach((k, i) => {
    if (k === null) return;
    if (!pool.has(k)) pool.set(k, []);
    pool.get(k).push(i);
  });

  const hitB = keysB.map((k) => (k === null ? null : false));
  const hitA = keysA.map((k) => {
    if (k === null) return null;
    const queue = pool.get(k);
    if (!queue || !queue.length) return false;
    hitB[queue.shift()] = true;
    return true;
  });
  return { hitA, hitB };
}

/* ---------- compare + paint ---------- */

async function compare() {
  let ra, rb;
  try {
    ra = resolveRange("a");
    rb = resolveRange("b");
  } catch (e) {
    setStatus(e.message, true);
    return;
  }
  const opts = { ignoreSign: $("ignore-sign").checked, ignoreCase: $("ignore-case").checked };

  $("compare").disabled = true;
  setStatus("Reading…");
  try {
    await Excel.run(async (ctx) => {
      const wsA = ctx.workbook.worksheets.getItem(ra.sheet);
      const wsB = ctx.workbook.worksheets.getItem(rb.sheet);
      const rngA = wsA.getRange(addressOf(ra));
      const rngB = wsB.getRange(addressOf(rb));
      rngA.load("values");
      rngB.load("values");
      await ctx.sync();

      const valuesA = rngA.values.map((row) => row[0]);
      const valuesB = rngB.values.map((row) => row[0]);
      const { hitA, hitB } = matchColumns(valuesA, valuesB, opts);

      setStatus("Colouring…");
      paintColumn(wsA, ra, hitA);
      paintColumn(wsB, rb, hitB);
      await ctx.sync();
    });

    lastPainted = [ra, rb];
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
 * Fill the cells of one column. Consecutive rows of the same colour are merged
 * into runs, and the runs of a colour go in as one multi-area call where the
 * host supports it, so a thousand-row column is a few operations.
 */
function paintColumn(ws, range, hits) {
  const runs = { [GREEN]: [], [RED]: [] };
  let i = 0;
  while (i < hits.length) {
    const hit = hits[i];
    if (hit === null) { i++; continue; }
    let j = i;
    while (j + 1 < hits.length && hits[j + 1] === hit) j++;
    const letter = colLetter(range.column);
    runs[hit ? GREEN : RED].push(`${letter}${range.firstRow + i}:${letter}${range.firstRow + j}`);
    i = j + 1;
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
  let ranges = lastPainted;
  if (!ranges.length) {
    try { ranges = [resolveRange("a"), resolveRange("b")]; }
    catch (e) { setStatus(e.message, true); return; }
  }
  $("clear-colours").disabled = true;
  setStatus("Clearing…");
  try {
    await Excel.run(async (ctx) => {
      for (const r of ranges) {
        ctx.workbook.worksheets.getItem(r.sheet).getRange(addressOf(r)).format.fill.clear();
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
