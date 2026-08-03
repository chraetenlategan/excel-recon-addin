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

// Per side after "Load & detect": { sheetName, rows, columns, mapping }.
const loaded = { cashbook: null, statement: null, ledger: null };

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in only runs in Excel.", true);
    return;
  }
  $("refresh-sheets").onclick = loadSheetList;
  $("load-detect").onclick = loadAndDetect;
  $("reconcile").onclick = reconcile;
  initFlex();          // Modular Recon wires itself up (flex-setup.js)
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
      const names = sheets.items.map((s) => s.name)
        .filter((n) => !n.startsWith(RESULT_PREFIX) && !n.startsWith(FLEX_PREFIX));
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

async function readSheet(ctx, name) {
  const ws = ctx.workbook.worksheets.getItem(name);
  const used = ws.getUsedRangeOrNullObject(true);
  used.load(["values"]);
  await ctx.sync();
  if (used.isNullObject) return [];
  return used.values.map((row) => row.map((v) => (v === null ? "" : v)));
}

/* ---------- 1b. load + auto-detect, then render mapping ---------- */

async function loadAndDetect() {
  const cbName = $("sel-cashbook").value;
  const stName = $("sel-statement").value;
  const glName = $("sel-ledger").value;
  if (!cbName || !(stName || glName)) {
    setStatus("Pick a cashbook plus at least one of bank statement / ledger.", true);
    return;
  }
  setStatus("Reading sheets & detecting columns…");
  try {
    await Excel.run(async (ctx) => {
      const chosen = { cashbook: cbName, statement: stName, ledger: glName };
      for (const side of SIDES) {
        const name = chosen[side];
        loaded[side] = null;
        if (!name) continue;
        const rows = await readSheet(ctx, name);
        if (!rows.length) throw new Error(`"${name}" looks empty.`);
        const analysis = await runLocalAnalyze(rows, [], name);
        loaded[side] = {
          sheetName: name,
          rows,
          columns: analysis.columns,
          mapping: { ...analysis.mapping, headerRow: analysis.headerRow, monthFilter: [] },
        };
      }
    });
    SIDES.forEach(renderMapping);
    $("mapping-wrap").classList.remove("hidden");
    setStatus("Check the detected columns, then Reconcile.");
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

  let html = `<h3>${LABELS[side]} — sheet “${escapeHtml(L.sheetName)}”</h3>`;

  html += `<div class="field grid2">
      <label style="margin:0"><span>Header row (0 = none)</span>
        <input type="number" min="0" data-k="headerRow" value="${m.headerRow || 0}"></label>
      <label style="margin:0"><span>Amount layout</span>
        <select data-k="mode">
          <option value="single"${dc ? "" : " selected"}>Single amount</option>
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
      : `<span class="none">no dates detected</span>`;
    html += `<div class="field"><span>Month filter (none = all months)</span><div class="months">${chips}</div></div>`;
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
      setStatus(`${LABELS[side]}: pick at least a Date column and an Amount (or Debit/Credit).`, true);
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
    const specs = buildResultSheets(result);
    setStatus(`Writing ${specs.length} result sheets…`);
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

const FILL = {
  green: { fill: "C6EFCE", font: "006100" },
  amber: { fill: "FFEB9C", font: "9C6500" },
  red:   { fill: "FFC7CE", font: "9C0006" },
};
const NAVY = "1F3864";
const GRID = "D9D9D9";

function supports(v) {
  try { return Office.context.requirements.isSetSupported("ExcelApi", v); } catch { return false; }
}

/**
 * Write a set of sheet specs into the workbook, replacing whatever the previous
 * run of the same flow left behind. `prefix` says which sheets that is — the
 * classic flow owns "Recon - …" and Modular Recon owns "Modular - …", so
 * running one never eats the other's output.
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

  // Land the user on the sheet that reads as the answer: the side-by-side
  // Comparison for the classic flow, otherwise whatever was written first.
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

  // Coloured rows (group consecutive same-colour rows into one range).
  const fills = spec.rowFills || {};
  let run = null;
  const flush = () => {
    if (!run) return;
    const rng = ws.getRangeByIndexes(run.start, 0, run.count, width);
    rng.format.fill.color = "#" + FILL[run.color].fill;
    rng.format.font.color = "#" + FILL[run.color].font;
    run = null;
  };
  for (let i = 0; i < nRows; i++) {
    const color = fills[i] || null;
    if (run && run.color === color && color) { run.count++; continue; }
    flush();
    if (color) run = { start: i, count: 1, color };
  }
  flush();

  // Free-form coloured rectangles (Modular Recon paints individual compared
  // cells, in colours the user picked, so it can't use the three-colour runs
  // above). Each rect is already a maximal block — see flexFillRects.
  for (const rect of spec.paintRects || []) {
    if (rect.r0 >= nRows || rect.c0 >= width) continue;
    const rows = Math.min(rect.rows, nRows - rect.r0);
    const cols = Math.min(rect.cols, width - rect.c0);
    const rng = ws.getRangeByIndexes(rect.r0, rect.c0, rows, cols);
    rng.format.fill.color = "#" + rect.fill;
    if (rect.font) rng.format.font.color = "#" + rect.font;
  }

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

/* ---------- summary chips ---------- */

function showSummary(result) {
  const s = result.summary;
  const box = $("summary");
  const chips = [];
  if (result.hasStatement) {
    chips.push(
      ["matched", "Matched (BS)", s["Matched"]],
      ["check", "Check desc (BS)", s["Check description"]],
      ["missing", "Not found (BS)", s["Not found"]],
      ["plain", "Unmatched BS rows", result.unmatchedStatementCount],
    );
  }
  if (result.hasLedger) {
    const g = result.ledgerSummary;
    chips.push(
      ["matched", "Matched (GL)", g["Matched"] + g["Check description"]],
      ["missing", "Not found (GL)", g["Not found"]],
      ["plain", "Unmatched GL rows", result.unmatchedLedgerCount],
    );
  }
  const noAmount = (result.hasStatement ? s : result.ledgerSummary)["No amount"];
  if (noAmount) chips.push(["noamount", "No amount", noAmount]);

  box.innerHTML = chips
    .filter(([, , n]) => n !== undefined && n !== null)
    .map(([cls, label, n]) => `<span class="chip ${cls}">${label}: ${n}</span>`)
    .join("");
  box.classList.remove("hidden");
}
