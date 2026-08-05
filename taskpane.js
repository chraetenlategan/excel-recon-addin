"use strict";

/**
 * taskpane.js — the only Excel-aware code. Three jobs:
 *   1. list worksheets + let the user map columns (role pickers, header row,
 *      amount layout, month filter) per side,
 *   2. read the chosen sheets and run the ported engine,
 *   3. write the full result-sheet set (comparison, reverse-direction per-file
 *      sheets, unmatched correlations) back into the workbook.
 *
 * engine.js / comparison.js / sheets.js are untouched web-app logic; everything
 * Office.js lives here.
 */

const $ = (id) => document.getElementById(id);
const SIDES = ["cashbook", "statement", "ledger"];
const LABELS = { cashbook: "Cashbook", statement: "Bank statement", ledger: "General ledger" };
const ROLE_LABEL = { date: "Date", description: "Description", amount: "Amount", debit: "Debit (out)", credit: "Credit (in)" };

// Per side after "Load": { sheetName, rows, columns, mapping }.
const loaded = { cashbook: null, statement: null, ledger: null };

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in only runs in Excel.", true);
    return;
  }
  $("refresh-sheets").onclick = loadSheetList;
  $("load-detect").onclick = loadAndDetect;
  $("reconcile").onclick = reconcile;
  $("clear-colours").onclick = clearColours;
  loadSheetList();
});

function setStatus(msg, isError) {
  const el = $("status");
  el.textContent = msg || "";
  el.classList.toggle("err", !!isError);
}

/* ---------- 1a. list worksheets ---------- */

async function loadSheetList() {
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      const names = sheets.items.map((s) => s.name).filter((n) => !n.startsWith(RESULT_PREFIX));
      const guess = (needles) => names.find((n) => needles.some((k) => n.toLowerCase().includes(k))) || "";
      fillSelect("sel-cashbook", names, false, guess(["cashbook", "cash book", "cash"]));
      fillSelect("sel-statement", names, true, guess(["bank", "statement"]));
      fillSelect("sel-ledger", names, true, guess(["ledger", "gl", "general"]));
    });
    setStatus("");
  } catch (e) {
    setStatus("Could not read worksheets: " + e.message, true);
  }
}

function fillSelect(id, names, allowNone, selected) {
  const sel = $(id);
  sel.innerHTML = "";
  if (allowNone) sel.appendChild(new Option("— none —", ""));
  for (const n of names) sel.appendChild(new Option(n, n));
  if (selected) sel.value = selected;
}

/* ---------- read a sheet's used range as rows ---------- */

// Returns the used range's values plus where it sits on the sheet, since
// colour-only mode writes back to those very cells.
async function readSheet(ctx, name) {
  const ws = ctx.workbook.worksheets.getItem(name);
  const used = ws.getUsedRangeOrNullObject(true);
  used.load(["values", "rowIndex", "columnIndex"]);
  await ctx.sync();
  if (used.isNullObject) return { rows: [], origin: { row: 0, col: 0 } };
  return {
    rows: used.values.map((row) => row.map((v) => (v === null ? "" : v))),
    origin: { row: used.rowIndex, col: used.columnIndex },
  };
}

/* ---------- 1b. load + auto-detect, then render mapping ---------- */

async function loadAndDetect() {
  const cbName = $("sel-cashbook").value;
  const stName = $("sel-statement").value;
  const glName = $("sel-ledger").value;
  if (!cbName || !(stName || glName)) {
    setStatus("Pick a cashbook and one sheet to match it against.", true);
    return;
  }
  setStatus("Reading…");
  try {
    await Excel.run(async (ctx) => {
      const chosen = { cashbook: cbName, statement: stName, ledger: glName };
      for (const side of SIDES) {
        const name = chosen[side];
        loaded[side] = null;
        if (!name) continue;
        const { rows, origin } = await readSheet(ctx, name);
        if (!rows.length) throw new Error(`"${name}" looks empty.`);
        const analysis = await runLocalAnalyze(rows, [], name);
        loaded[side] = {
          sheetName: name,
          rows,
          origin,
          columns: analysis.columns,
          mapping: { ...analysis.mapping, headerRow: analysis.headerRow, monthFilter: [] },
        };
      }
    });
    SIDES.forEach(renderMapping);
    $("mapping-wrap").classList.remove("hidden");
    setStatus("Check the columns, then Reconcile.");
  } catch (e) {
    setStatus("Error: " + e.message, true);
  }
}

// Unique "YYYYMM" months present in the mapped date column, sorted.
function detectMonths(L) {
  const m = L.mapping;
  if (m.date === null || m.date === undefined) return [];
  const start = m.headerRow || 0;
  const set = new Set();
  for (let i = start; i < L.rows.length; i++) {
    for (const k of monthKeys(L.rows[i][m.date])) set.add(k);
  }
  return [...set].sort();
}

function monthLabel(key) {
  const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const mm = +key.slice(4, 6);
  return mm >= 1 && mm <= 12 ? `${names[mm - 1]} ${key.slice(0, 4)}` : key;
}

function colOptions(columns, selected) {
  const none = `<option value=""${selected === null || selected === undefined ? " selected" : ""}>— none —</option>`;
  const opts = columns.map((c) => {
    const label = `${c.letter}${c.header ? " — " + c.header : ""}`;
    return `<option value="${c.index}"${c.index === selected ? " selected" : ""}>${escapeHtml(label)}</option>`;
  });
  return none + opts.join("");
}

function renderMapping(side) {
  const card = $("map-" + side);
  const L = loaded[side];
  if (!L) { card.classList.add("hidden"); card.innerHTML = ""; return; }
  card.classList.remove("hidden");
  const m = L.mapping;
  const dc = m.mode === "debit_credit";
  const monthable = side === "cashbook" || side === "ledger";

  let html = `<h3>${LABELS[side]} · ${escapeHtml(L.sheetName)}</h3>`;

  html += `<div class="field grid2">
      <label style="margin:0"><span>Header row</span>
        <input type="number" min="0" data-k="headerRow" value="${m.headerRow || 0}"></label>
      <label style="margin:0"><span>Amount</span>
        <select data-k="mode">
          <option value="single"${dc ? "" : " selected"}>Single</option>
          <option value="debit_credit"${dc ? " selected" : ""}>Debit + Credit</option>
        </select></label>
    </div>`;

  html += `<div class="field"><span>Date</span><select data-role="date">${colOptions(L.columns, m.date)}</select></div>`;
  html += `<div class="field"><span>Description</span><select data-role="description">${colOptions(L.columns, m.description)}</select></div>`;

  if (dc) {
    html += `<div class="field grid2">
        <label style="margin:0"><span>Debit (out)</span><select data-role="debit">${colOptions(L.columns, m.debit)}</select></label>
        <label style="margin:0"><span>Credit (in)</span><select data-role="credit">${colOptions(L.columns, m.credit)}</select></label>
      </div>`;
  } else {
    html += `<div class="field"><span>Amount</span><select data-role="amount">${colOptions(L.columns, m.amount)}</select></div>`;
  }

  if (monthable) {
    const months = detectMonths(L);
    const active = new Set(_monthFilterKeys(m.monthFilter));
    const chips = months.length
      ? months.map((k) => `<label><input type="checkbox" data-month="${k}"${active.has(k) ? " checked" : ""}> ${monthLabel(k)}</label>`).join("")
      : `<span class="none">no dates</span>`;
    html += `<div class="field"><span>Months</span><div class="months">${chips}</div></div>`;
  }

  card.innerHTML = html;
  wireMapping(side);
}

function wireMapping(side) {
  const card = $("map-" + side);
  const m = loaded[side].mapping;
  const val = (el) => (el.value === "" ? null : parseInt(el.value, 10));

  card.querySelectorAll("select[data-role]").forEach((sel) => {
    sel.onchange = () => { m[sel.dataset.role] = val(sel); };
  });
  const modeSel = card.querySelector('[data-k="mode"]');
  if (modeSel) modeSel.onchange = () => {
    m.mode = modeSel.value;
    // Switching layout swaps which role pickers are relevant; clear the others.
    if (m.mode === "debit_credit") { m.amount = null; }
    else { m.debit = null; m.credit = null; }
    renderMapping(side);
  };
  const hdr = card.querySelector('[data-k="headerRow"]');
  if (hdr) hdr.onchange = () => {
    m.headerRow = Math.max(0, parseInt(hdr.value, 10) || 0);
    renderMapping(side); // month list depends on header row
  };
  card.querySelectorAll("input[data-month]").forEach((cb) => {
    cb.onchange = () => {
      const set = new Set(_monthFilterKeys(m.monthFilter));
      if (cb.checked) set.add(cb.dataset.month); else set.delete(cb.dataset.month);
      m.monthFilter = [...set];
    };
  });
}

/* ---------- 2 + 3. reconcile and write result sheets ---------- */

function validSide(m) {
  const hasAmount = (m.mode === "debit_credit") ? (m.debit !== null || m.credit !== null) : (m.amount !== null);
  return m.date !== null && m.date !== undefined && hasAmount;
}

async function reconcile() {
  if (!loaded.cashbook) { setStatus("Load a cashbook first.", true); return; }
  for (const side of SIDES) {
    if (loaded[side] && !validSide(loaded[side].mapping)) {
      setStatus(`${LABELS[side]}: needs a Date and an Amount column.`, true);
      return;
    }
  }
  $("reconcile").disabled = true;
  setStatus("Matching…");
  try {
    const cb = loaded.cashbook, st = loaded.statement, gl = loaded.ledger;
    const result = runLocalReconcile(
      cb.rows, st ? st.rows : null, cb.mapping, st ? st.mapping : null,
      gl ? gl.rows : null, gl ? gl.mapping : null
    );

    if ($("colour-only").checked) {
      setStatus("Colouring…");
      let painted = 0;
      await Excel.run(async (ctx) => { painted = await paintSourceSheets(ctx, result); });
      showSummary(result);
      setStatus(`Done — ${painted} amount${painted === 1 ? "" : "s"} coloured on your own sheets.`);
      return;
    }

    const specs = buildResultSheets(result);
    setStatus("Writing…");
    await Excel.run(async (ctx) => { await writeSpecs(ctx, specs); });
    showSummary(result);
    setStatus(`Done — see the “${RESULT_PREFIX}…” sheets.`);
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    $("reconcile").disabled = false;
  }
}

/* ---------- Office.js sheet writer ---------- */

const NAVY = "1F3864";
const GRID = "D9D9D9";

function supports(v) {
  try { return Office.context.requirements.isSetSupported("ExcelApi", v); } catch { return false; }
}

/**
 * Write a set of sheet specs into the workbook, replacing whatever the previous
 * run left behind (every sheet whose name starts with `prefix`).
 */
async function writeSpecs(ctx, specs, prefix = RESULT_PREFIX) {
  // Clear any previous result sheets so each run is clean.
  const all = ctx.workbook.worksheets;
  all.load("items/name");
  await ctx.sync();
  for (const ws of all.items) {
    if (ws.name.startsWith(prefix)) ws.delete();
  }
  await ctx.sync();

  for (const spec of specs) writeOne(ctx, spec);
  await ctx.sync();

  // Land the user on the sheet that reads as the answer.
  const landing = specs.find((s) => s.name === prefix + "Comparison") || specs[0];
  if (landing) {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(landing.name);
    ws.load("name");
    await ctx.sync();
    if (!ws.isNullObject) ws.activate();
  }
}

function writeOne(ctx, spec) {
  const ws = ctx.workbook.worksheets.add(spec.name);

  // Rectangular width: widest of data, declared column widths, and any band row.
  let width = spec.colWidths.length;
  for (const r of spec.aoa) width = Math.max(width, r.length);
  const nRows = spec.aoa.length;
  if (!nRows || !width) return;

  const values = spec.aoa.map((r) => {
    const out = new Array(width);
    for (let c = 0; c < width; c++) {
      const v = r[c];
      out[c] = (v === undefined || v === null) ? "" : v;
    }
    return out;
  });

  const body = ws.getRangeByIndexes(0, 0, nRows, width);
  body.values = values;

  // Light gridlines across the whole block.
  for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"]) {
    const b = body.format.borders.getItem(edge);
    b.style = "Continuous"; b.color = "#" + GRID; b.weight = "Thin";
  }

  // The outcome colouring: green/amber/red cells, already merged into maximal
  // blocks by sheets.js (_painter).
  paintCells(ws, spec.paintRects || [], nRows, width);

  // Navy header band(s).
  for (const r of spec.bandRows || []) {
    const rng = ws.getRangeByIndexes(r, 0, 1, width);
    rng.format.fill.color = "#" + NAVY;
    rng.format.font.color = "white";
    rng.format.font.bold = true;
    rng.format.horizontalAlignment = "Center";
  }

  // Bold-navy section title cells (column 0).
  for (const r of spec.titleRows || []) {
    const rng = ws.getRangeByIndexes(r, 0, 1, 1);
    rng.format.font.bold = true;
    rng.format.font.color = "#" + NAVY;
  }

  // Column widths (chars → points ≈ ×7).
  spec.colWidths.forEach((w, c) => {
    if (c < width) ws.getRangeByIndexes(0, c, 1, 1).getEntireColumn().format.columnWidth = w * 7;
  });

  // Freeze under the last band row, and add an AutoFilter, where supported.
  const bands = spec.bandRows || [];
  if (bands.length && supports("1.7")) {
    try { ws.freezePanes.freezeRows(Math.max(...bands) + 1); } catch (e) { /* older Excel */ }
  }
  if (spec.autofilter && supports("1.9")) {
    const a = spec.autofilter;
    try {
      ws.autoFilter.apply(ws.getRangeByIndexes(a.headerRow, 0, a.lastRow - a.headerRow + 1, a.width));
    } catch (e) { /* older Excel */ }
  }
}

/**
 * Fill a sheet's coloured cells. Rectangles that share a colour are set through
 * one RangeAreas ("A2:A9,C4") call where Excel supports it, so a long sheet is
 * a few operations rather than one per block.
 */
const AREAS_PER_CALL = 50;

function paintCells(ws, rects, nRows, width) {
  const groups = new Map();
  for (const rect of rects) {
    if (rect.r0 >= nRows || rect.c0 >= width) continue;
    const rows = Math.min(rect.rows, nRows - rect.r0);
    const cols = Math.min(rect.cols, width - rect.c0);
    const key = rect.fill + "|" + (rect.font || "");
    if (!groups.has(key)) groups.set(key, { fill: rect.fill, font: rect.font, boxes: [] });
    groups.get(key).boxes.push({ r0: rect.r0, c0: rect.c0, rows, cols });
  }

  const areas = supports("1.9");
  for (const g of groups.values()) {
    if (!areas) {
      for (const b of g.boxes) {
        const rng = ws.getRangeByIndexes(b.r0, b.c0, b.rows, b.cols);
        rng.format.fill.color = "#" + g.fill;
        if (g.font) rng.format.font.color = "#" + g.font;
      }
      continue;
    }
    const addrs = g.boxes.map((b) => {
      const from = colLetter(b.c0) + (b.r0 + 1);
      const to = colLetter(b.c0 + b.cols - 1) + (b.r0 + b.rows);
      return from === to ? from : from + ":" + to;
    });
    for (let i = 0; i < addrs.length; i += AREAS_PER_CALL) {
      const rngs = ws.getRanges(addrs.slice(i, i + AREAS_PER_CALL).join(","));
      rngs.format.fill.color = "#" + g.fill;
      if (g.font) rngs.format.font.color = "#" + g.font;
    }
  }
}

/* ---------- colour-only mode ---------- */

// Excel's grid, used as the clamp when painting straight onto a source sheet.
const SHEET_ROWS = 1048576;
const SHEET_COLS = 16384;

// A cashbook row counts as found only if every side it was compared against
// claimed it; "Check description" still means the amount itself was found.
function foundOnAllSides(result, r) {
  const ok = (s) => s === "Matched" || s === "Check description";
  if (result.hasStatement && !ok(r.status)) return false;
  if (result.hasLedger && !ok(r.ledgerStatus)) return false;
  return true;
}

// Per side: 1-based row number within that sheet's used range -> found?
function colourOnlyPlan(result) {
  const plan = {};
  const withAmount = (rows) => rows.filter((r) => r.amount !== null);
  plan.cashbook = new Map(withAmount(result.rows).map((r) => [r.row, foundOnAllSides(result, r)]));
  if (result.hasStatement) plan.statement = new Map(withAmount(result.statement).map((s) => [s.row, s.matched]));
  if (result.hasLedger) plan.ledger = new Map(withAmount(result.ledger).map((s) => [s.row, s.matched]));
  return plan;
}

/**
 * Colour-only mode: no result sheets at all — just fill the amount cell of every
 * row on the user's own sheets, green where the amount was found on the other
 * side and red where it wasn't. Fills only; fonts and values stay untouched.
 */
async function paintSourceSheets(ctx, result) {
  const plan = colourOnlyPlan(result);
  const sources = result.sources || {};
  let painted = 0;

  for (const side of SIDES) {
    const found = plan[side];
    const src = sources[side];
    const L = loaded[side];
    if (!found || !src || !L) continue;
    const cols = src.amountCols || [];
    if (!cols.length) continue;

    const paint = _painter();
    for (const [rowNum, ok] of found) {
      const raw = src.rows[rowNum - 1] || [];
      // In debit/credit layout only the side that carries a figure is coloured.
      const filled = cols.filter((c) => String(raw[c] ?? "").trim() !== "");
      for (const c of (filled.length ? filled : cols)) {
        paint.set(L.origin.row + rowNum - 1, L.origin.col + c, ok ? "green" : "red");
      }
      painted++;
    }

    const ws = ctx.workbook.worksheets.getItem(L.sheetName);
    const rects = paint.rects().map((r) => ({ ...r, font: null }));
    paintCells(ws, rects, SHEET_ROWS, SHEET_COLS);
  }
  await ctx.sync();
  return painted;
}

// Undo colour-only mode: strip the fill from the amount column(s) of every
// loaded sheet, over the data rows only.
async function clearColours() {
  const sides = SIDES.filter((s) => loaded[s]);
  if (!sides.length) { setStatus("Load your sheets first.", true); return; }
  $("clear-colours").disabled = true;
  setStatus("Clearing…");
  try {
    await Excel.run(async (ctx) => {
      for (const side of sides) {
        const L = loaded[side];
        const start = L.mapping.headerRow || 0;
        const n = L.rows.length - start;
        const cols = amountCols(L.mapping);
        if (n <= 0 || !cols.length) continue;
        const ws = ctx.workbook.worksheets.getItem(L.sheetName);
        for (const c of cols) {
          ws.getRangeByIndexes(L.origin.row + start, L.origin.col + c, n, 1).format.fill.clear();
        }
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

/* ---------- summary ---------- */

// One line per side: green / amber / red counts in the same colours the cells
// were written in, plus how many of that side's own rows went unclaimed.
function showSummary(result) {
  const box = $("summary");
  const lines = [];

  const tally = (who, counts, leftover) => lines.push(
    `<div class="tally"><span class="who">${who}</span>` +
    `<span class="n green">${counts["Matched"]}</span>` +
    `<span class="n amber">${counts["Check description"]}</span>` +
    `<span class="n red">${counts["Not found"] + counts["No amount"]}</span>` +
    `<span class="n plain">${leftover}</span></div>`
  );

  if (result.hasStatement) tally("BS", result.summary, result.unmatchedStatementCount);
  if (result.hasLedger) tally("GL", result.ledgerSummary, result.unmatchedLedgerCount);

  box.innerHTML = lines.join("");
  box.classList.remove("hidden");
}
