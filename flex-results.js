"use strict";

/**
 * flex-results.js — step 3 of Modular Recon: run it, show it, colour it, write it.
 *
 * Three views come out of one run:
 *  - Results     one block per relationship: the counts, then every row on each
 *                side with its status and the row it paired with.
 *  - Data        the sheet as loaded, with the matched values painted green and
 *                the unmatched ones red (the colours are the user's to choose,
 *                and can paint just the compared cells or the whole row).
 *  - Write       the same thing as new "Modular - …" worksheets, so the
 *                colouring survives outside the pane.
 *
 * "Repeated value" (amber) is deliberately its own status rather than being
 * folded into matched: when a value occurs several times on a side the counts
 * are still right, but which row paired with which is arbitrary, and an auditor
 * needs to be told that rather than reassured.
 *
 * Where the web app downloads a SheetJS workbook, this builds sheet specs and
 * hands them to taskpane.js's Office.js writer — the same path the classic flow
 * uses for its "Recon - …" sheets.
 */

const FLEX_ROW_CAP = 300;    // rows rendered per side before the list is capped
const FLEX_DATA_CAP = 1000;  // rows rendered in the Data view

/* ---------- running ---------- */

async function flexRun() {
  const rels = flexRelationships();
  if (!rels.length) {
    setStatus("Draw at least one comparison between two sheets first.", true);
    return;
  }
  setStatus("Comparing the columns you linked…");
  await new Promise((resolve) => setTimeout(resolve, 0));   // let the status paint
  try {
    flexState.result = flexRunRecon();
    flexRenderSheetPicker();
    renderFlexResults();
    flexUpdateButtons();
    flexSwitchTab("flex-results-panel");
    setStatus("Done — review the counts, then write the sheets if you want them in the workbook.");
  } catch (e) {
    setStatus("Error: " + e.message, true);
  }
}

/* ---------- colours ---------- */

function flexHexToRgb(hex) {
  const clean = String(hex).replace("#", "");
  const full = clean.length === 3 ? clean.split("").map((c) => c + c).join("") : clean;
  return [0, 2, 4].map((i) => parseInt(full.slice(i, i + 2), 16) || 0);
}

function flexTint(hex, alpha) {
  const [r, g, b] = flexHexToRgb(hex);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

// A pale version of the colour for Excel cell fills, so the text stays readable.
function flexPaleHex(hex, keep = 0.22) {
  const [r, g, b] = flexHexToRgb(hex);
  const mix = (c) => Math.round(c * keep + 255 * (1 - keep)).toString(16).padStart(2, "0");
  return (mix(r) + mix(g) + mix(b)).toUpperCase();
}

function flexDarkHex(hex) {
  return flexHexToRgb(hex).map((c) => Math.round(c * 0.7).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function flexApplyColors() {
  const root = document.documentElement.style;
  for (const [status, hex] of Object.entries(flexState.colors)) {
    root.setProperty(`--flex-${status}`, hex);
    root.setProperty(`--flex-${status}-bg`, flexTint(hex, 0.16));
    root.setProperty(`--flex-${status}-line`, flexTint(hex, 0.45));
  }
}

/* ---------- results view ---------- */

function renderFlexResults() {
  const body = $("flex-results-body");
  if (!body) return;
  flexApplyColors();
  const result = flexState.result;

  if (!result) {
    body.innerHTML = `<p class="flex-rail-empty">Nothing has been reconciled yet — draw your comparisons on the Model tab, then press Reconcile.</p>`;
    return;
  }
  body.replaceChildren(...result.relationships.map(flexRelBlock));
}

function flexRelBlock(rel) {
  const block = document.createElement("div");
  block.className = "flex-result-block";

  const fields = rel.fields.map((f) =>
    `${escapeHtml(flexColumnName(rel.left, f.leftCol))} = ${escapeHtml(flexColumnName(rel.right, f.rightCol))}`).join(" · ");

  const s = rel.summary;
  block.innerHTML = `
    <div class="flex-result-head">
      <h3>${escapeHtml(rel.leftLabel)} <span class="flex-eq">⇄</span> ${escapeHtml(rel.rightLabel)}</h3>
      <p class="flex-result-fields">${fields}${rel.fields.length > 1 ? " <em>(all must agree)</em>" : ""}</p>
    </div>
    <div class="flex-chips">
      <span class="chip flex-chip matched">${s.pairs} pair${s.pairs === 1 ? "" : "s"} matched</span>
      ${s.left.ambiguous || s.right.ambiguous ? `<span class="chip flex-chip ambiguous">${s.left.ambiguous + s.right.ambiguous} on a repeated value</span>` : ""}
      <span class="chip flex-chip unmatched">${s.left.unmatched} not found from ${escapeHtml(rel.leftLabel)}</span>
      <span class="chip flex-chip unmatched">${s.right.unmatched} not found from ${escapeHtml(rel.rightLabel)}</span>
      ${s.left.novalue || s.right.novalue ? `<span class="chip flex-chip novalue">${s.left.novalue + s.right.novalue} with no value to compare</span>` : ""}
    </div>`;

  const controls = document.createElement("div");
  controls.className = "flex-result-controls";
  controls.innerHTML = `
    <label>Show
      <select class="flex-filter">
        <option value="all">every row</option>
        <option value="unmatched" selected>only what didn't match</option>
        <option value="matched">only what matched</option>
      </select>
    </label>`;
  block.appendChild(controls);

  const tables = document.createElement("div");
  tables.className = "flex-result-tables";
  block.appendChild(tables);

  const draw = (filter) => {
    tables.replaceChildren(
      flexSideTable(rel, "left", filter),
      flexSideTable(rel, "right", filter));
  };
  controls.querySelector(".flex-filter").addEventListener("change", (e) => draw(e.target.value));
  draw("unmatched");

  if (rel.duplicates.length) {
    const dup = document.createElement("details");
    dup.className = "flex-dupes";
    dup.innerHTML = `<summary>${rel.duplicates.length} value${rel.duplicates.length === 1 ? "" : "s"} appear more than once — pairing inside these groups is arbitrary</summary>`;
    const table = document.createElement("table");
    table.className = "flex-table";
    table.innerHTML = `<thead><tr><th>Value</th><th>${escapeHtml(rel.leftLabel)}</th><th>${escapeHtml(rel.rightLabel)}</th><th>Difference</th></tr></thead>`;
    const tbody = document.createElement("tbody");
    for (const d of rel.duplicates.slice(0, 100)) {
      const tr = document.createElement("tr");
      const diff = d.left - d.right;
      tr.innerHTML = `<td>${escapeHtml(d.label)}</td><td>${d.left}</td><td>${d.right}</td>
        <td class="${diff ? "flex-cell-unmatched" : ""}">${diff === 0 ? "balanced" : `${Math.abs(diff)} extra on ${diff > 0 ? escapeHtml(rel.leftLabel) : escapeHtml(rel.rightLabel)}`}</td>`;
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    dup.appendChild(table);
    block.appendChild(dup);
  }

  return block;
}

function flexSideTable(rel, side, filter) {
  const sheetId = rel[side];
  const rows = rel[side === "left" ? "leftRows" : "rightRows"];
  const cols = rel.fields.map((f) => side === "left" ? f.leftCol : f.rightCol);
  const otherId = side === "left" ? rel.right : rel.left;

  const wanted = rows.filter((e) => {
    if (filter === "unmatched") return e.status === "unmatched" || e.status === "novalue";
    if (filter === "matched") return e.status === "matched" || e.status === "ambiguous";
    return true;
  });

  const wrap = document.createElement("div");
  wrap.className = "flex-side";
  wrap.innerHTML = `<h4>${escapeHtml(flexSheet(sheetId).label)} <span class="flex-side-count">${wanted.length} of ${rows.length}</span></h4>`;

  const scroll = document.createElement("div");
  scroll.className = "flex-side-scroll";
  const table = document.createElement("table");
  table.className = "flex-table";
  table.innerHTML = `<thead><tr>
      <th class="rownum">#</th>
      ${cols.map((c) => `<th>${escapeHtml(flexColumnName(sheetId, c))}</th>`).join("")}
      <th>Status</th>
      <th>Row on ${escapeHtml(flexSheet(otherId).label)}</th>
    </tr></thead>`;

  const tbody = document.createElement("tbody");
  for (const e of wanted.slice(0, FLEX_ROW_CAP)) {
    const tr = document.createElement("tr");
    tr.className = `flex-row-${e.status}`;
    const jump = `<td class="rownum flex-jump" data-sheet="${sheetId}" data-row="${e.row}" title="Open this sheet at row ${e.row}">${e.row}</td>`;
    const partner = e.partner
      ? `<td class="flex-jump" data-sheet="${otherId}" data-row="${e.partner}" title="Open ${escapeHtml(flexSheet(otherId).label)} at row ${e.partner}">${e.partner}</td>`
      : `<td></td>`;
    tr.innerHTML = `${jump}${e.values.map((v) => `<td>${escapeHtml(v)}</td>`).join("")}
      <td class="flex-status flex-cell-${e.status}">${FLEX_STATUS_LABEL[e.status]}</td>${partner}`;
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  scroll.appendChild(table);

  if (wanted.length > FLEX_ROW_CAP) {
    const note = document.createElement("p");
    note.className = "hint";
    note.textContent = `Showing the first ${FLEX_ROW_CAP} — write the sheets for the full list.`;
    scroll.appendChild(note);
  }
  if (!wanted.length) {
    const note = document.createElement("p");
    note.className = "hint";
    note.textContent = "Nothing to show here.";
    scroll.appendChild(note);
  }

  wrap.appendChild(scroll);
  return wrap;
}

/* ---------- the Data view ---------- */

// The web app gives each sheet its own tab; a pane has no room for a strip of
// them, so the same choice is a picker.
function flexRenderSheetPicker() {
  const picker = $("flex-data-sheet");
  if (!picker) return;
  const sheets = flexIncludedSheets();
  const current = flexState.activeSheetId;
  picker.innerHTML = sheets.map((s) => `<option value="${s.id}">${escapeHtml(s.label)}</option>`).join("");
  if (sheets.length) {
    flexState.activeSheetId = sheets.some((s) => s.id === current) ? current : sheets[0].id;
    picker.value = flexState.activeSheetId;
  }
}

function flexShowSheet(sheetId, scrollToRow = null) {
  flexState.activeSheetId = sheetId;
  const picker = $("flex-data-sheet");
  if (picker) picker.value = sheetId;
  renderFlexData(scrollToRow);
  flexSwitchTab("flex-data-panel");
}

function renderFlexData(scrollToRow = null) {
  const sheet = flexSheet(flexState.activeSheetId);
  const table = $("flex-data-grid");
  const title = $("flex-data-title");
  const hint = $("flex-data-hint");
  if (!sheet || !table) return;
  flexApplyColors();

  const store = (flexState.result && flexState.result.bySheet[sheet.id]) || { rows: new Map(), cols: new Set() };
  const painted = flexState.highlightMode === "row" ? null : store.cols;

  title.textContent = sheet.label;
  const counts = { matched: 0, ambiguous: 0, unmatched: 0, novalue: 0 };
  for (const entry of store.rows.values()) if (entry.status) counts[entry.status]++;
  hint.textContent = `${counts.matched} matched · ${counts.ambiguous} on a repeated value · ${counts.unmatched} not found`
    + (counts.novalue ? ` · ${counts.novalue} with no value to compare` : "")
    + ` — ${flexState.highlightMode === "row" ? "whole rows" : "compared values"} are coloured.`;

  const width = sheet.columns.length;
  table.querySelector("thead").innerHTML = `<tr><th class="rownum">#</th>${sheet.columns.map((c) =>
    `<th${painted && painted.has(c.index) ? ' class="flex-col-used"' : ""}>${escapeHtml(c.name)}</th>`).join("")}</tr>`;

  const tbody = table.querySelector("tbody");
  tbody.replaceChildren();
  const limit = Math.min(sheet.rows.length, sheet.headerRow + FLEX_DATA_CAP);
  for (let r = sheet.headerRow; r < limit; r++) {
    const cells = sheet.rows[r] || [];
    const rowNumber = r + 1;
    const entry = store.rows.get(rowNumber);
    const tr = document.createElement("tr");
    tr.dataset.row = rowNumber;
    if (entry && entry.status && flexState.highlightMode === "row") tr.className = `flex-row-${entry.status}`;
    const status = entry && entry.status;

    const num = document.createElement("td");
    num.className = "rownum flex-jump";
    num.textContent = rowNumber;
    num.title = `Select row ${rowNumber} on “${sheet.name}”`;
    num.dataset.sheet = sheet.id;
    num.dataset.row = rowNumber;
    tr.appendChild(num);

    for (let c = 0; c < width; c++) {
      const td = document.createElement("td");
      td.textContent = _display_date(cells[c] ?? "");
      if (status && painted && painted.has(c)) {
        td.className = `flex-cell-${status}`;
        td.title = flexRowTitle(entry);
      }
      tr.appendChild(td);
    }
    if (status && flexState.highlightMode === "row") tr.title = flexRowTitle(entry);
    tbody.appendChild(tr);
  }

  if (sheet.rows.length > limit) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td class="rownum"></td><td colspan="${width}" class="hint">Showing the first ${FLEX_DATA_CAP} rows — write the sheets for all of them.</td>`;
    tbody.appendChild(tr);
  }

  if (scrollToRow) {
    const target = tbody.querySelector(`tr[data-row="${scrollToRow}"]`);
    if (target) {
      tbody.querySelectorAll("tr.hl").forEach((el) => el.classList.remove("hl"));
      target.classList.add("hl");
      const wrap = $("flex-data-panel").querySelector(".grid-wrap");
      wrap.scrollTop = target.offsetTop - wrap.clientHeight / 2;
    }
  }
}

// What a coloured cell says when you hover it: one line per relationship.
function flexRowTitle(entry) {
  const lines = [];
  for (const [relId, info] of Object.entries(entry.byRel)) {
    const other = flexSheet(info.partnerSheet);
    const where = info.partner ? ` (row ${info.partner})` : "";
    lines.push(`${FLEX_STATUS_LABEL[info.status]} vs ${other ? other.label : relId}${where}`);
  }
  return lines.join("\n");
}

// The sheets are the workbook's own, so a row number can point at the real
// thing rather than only at the copy in the pane.
async function flexSelectInExcel(sheetId, row) {
  const sheet = flexSheet(sheetId);
  if (!sheet) return;
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet.name);
      ws.activate();
      ws.getRangeByIndexes(row - 1, 0, 1, Math.max(1, sheet.columns.length)).select();
      await ctx.sync();
    });
  } catch (e) {
    setStatus(`Could not select row ${row} on “${sheet.name}”: ${e.message}`, true);
  }
}

/* ---------- writing the result sheets ---------- */

// One rectangle per run of same-status cells. A per-cell fill would be one
// Office.js range operation per cell; grouping identical adjacent rows and then
// identical adjacent columns turns a painted sheet into a handful of ranges.
function flexFillRects(nRows, nCols, statusAt) {
  const grid = [];
  for (let r = 0; r < nRows; r++) {
    const row = new Array(nCols);
    for (let c = 0; c < nCols; c++) row[c] = statusAt(r, c) || null;
    grid.push(row);
  }
  const sameRow = (a, b) => a.every((v, i) => v === b[i]);

  const rects = [];
  let r = 0;
  while (r < nRows) {
    let end = r + 1;
    while (end < nRows && sameRow(grid[r], grid[end])) end++;
    let c = 0;
    while (c < nCols) {
      const status = grid[r][c];
      if (!status) { c++; continue; }
      let cEnd = c + 1;
      while (cEnd < nCols && grid[r][cEnd] === status) cEnd++;
      rects.push({ r0: r, c0: c, rows: end - r, cols: cEnd - c, status });
      c = cEnd;
    }
    r = end;
  }
  return rects;
}

// "novalue" has no colour of its own — nothing was compared, so nothing is
// claimed about the row.
function flexStatusStyle(status) {
  const hex = flexState.colors[status];
  if (!hex) return null;
  return { fill: flexPaleHex(hex), font: flexDarkHex(hex) };
}

// Excel: 31 characters, none of  [ ] : * ? / \  — and no duplicates.
function flexSheetName(used, label) {
  let name = String(label).replace(/[\[\]:*?/\\]/g, " ").slice(0, 31).trim() || "Sheet";
  let n = 2;
  while (used.has(name.toLowerCase())) name = `${name.slice(0, 28)} ${n++}`;
  used.add(name.toLowerCase());
  return name;
}

// One output sheet per included sheet: the rows as loaded, plus a status and a
// partner-row column for every relationship the sheet takes part in.
function buildFlexSheetSpecs(result) {
  const specs = [];
  const usedNames = new Set();

  for (const sheet of flexIncludedSheets()) {
    const store = result.bySheet[sheet.id];
    if (!store) continue;
    const rels = result.relationships.filter((r) => r.left === sheet.id || r.right === sheet.id);
    const width = sheet.columns.length;

    const header = [...sheet.columns.map((c) => c.name)];
    for (const rel of rels) {
      const other = flexSheet(rel.left === sheet.id ? rel.right : rel.left);
      header.push(`Status vs ${other.label}`, `Row on ${other.label}`);
    }

    const aoa = [header];
    const rowStatus = [];   // status per data row, indexed the same way as aoa - 1
    for (let r = sheet.headerRow; r < sheet.rows.length; r++) {
      const cells = sheet.rows[r] || [];
      if (!cells.some((v) => _text(v) !== "")) continue;
      const entry = store.rows.get(r + 1);
      const line = [];
      for (let c = 0; c < width; c++) {
        const v = cells[c];
        line.push(v instanceof Date ? _display_date(v) : (typeof v === "number" ? v : _text(v)));
      }
      for (const rel of rels) {
        const info = entry && entry.byRel[rel.id];
        line.push(info ? FLEX_STATUS_LABEL[info.status] : "");
        line.push(info && info.partner ? info.partner : "");
      }
      aoa.push(line);
      rowStatus.push(entry ? entry.status : null);
    }

    // Which columns get painted: the compared ones plus the appended status
    // columns, or the whole row when the user asked for that.
    const totalWidth = header.length;
    const wholeRow = flexState.highlightMode === "row";
    const rects = flexFillRects(rowStatus.length, totalWidth, (r, c) =>
      (wholeRow || store.cols.has(c) || c >= width) ? rowStatus[r] : null);

    const paintRects = [];
    for (const rect of rects) {
      const style = flexStatusStyle(rect.status);
      if (!style) continue;
      paintRects.push({ r0: rect.r0 + 1, c0: rect.c0, rows: rect.rows, cols: rect.cols, ...style });
    }

    specs.push({
      name: flexSheetName(usedNames, FLEX_PREFIX + sheet.name),
      aoa,
      colWidths: Array.from({ length: width }, () => 16).concat(
        rels.flatMap(() => [26, 12])),
      bandRows: [0],
      titleRows: [],
      rowFills: {},
      paintRects,
      autofilter: { headerRow: 0, width: totalWidth, lastRow: aoa.length - 1 },
    });
  }

  // A summary sheet so the output explains itself.
  const summary = [["Modular reconciliation"], ["Run", new Date().toLocaleString()], []];
  const titleRows = [0];
  for (const rel of result.relationships) {
    titleRows.push(summary.length);
    summary.push([`${rel.leftLabel} vs ${rel.rightLabel}`]);
    for (const f of rel.fields) {
      summary.push(["", `${flexColumnName(rel.left, f.leftCol)} = ${flexColumnName(rel.right, f.rightCol)}`, f.link.mode]);
    }
    summary.push(["", "Pairs matched", rel.summary.pairs]);
    summary.push(["", `Not found on ${rel.rightLabel}`, rel.summary.left.unmatched]);
    summary.push(["", `Not found on ${rel.leftLabel}`, rel.summary.right.unmatched]);
    if (rel.summary.left.ambiguous || rel.summary.right.ambiguous) {
      summary.push(["", "Rows on a repeated value", rel.summary.left.ambiguous + rel.summary.right.ambiguous]);
    }
    if (rel.summary.left.novalue || rel.summary.right.novalue) {
      summary.push(["", "Rows with no value to compare", rel.summary.left.novalue + rel.summary.right.novalue]);
    }
    summary.push([]);
  }
  specs.push({
    name: flexSheetName(usedNames, FLEX_PREFIX + "Summary"), aoa: summary, colWidths: [30, 40, 14],
    bandRows: [], titleRows, rowFills: {}, paintRects: [], autofilter: null,
  });

  return specs;
}

async function flexWriteSheets() {
  if (!flexState.result) return;
  $("flex-write").disabled = true;
  try {
    const specs = buildFlexSheetSpecs(flexState.result);
    setStatus(`Writing ${specs.length} sheet${specs.length === 1 ? "" : "s"}…`);
    await Excel.run(async (ctx) => { await writeSpecs(ctx, specs, FLEX_PREFIX); });
    setStatus(`Done — see the “${FLEX_PREFIX}…” sheets.`);
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    $("flex-write").disabled = false;
  }
}

/* ---------- wiring ---------- */

function initFlexResults() {
  flexApplyColors();

  $("flex-run").addEventListener("click", flexRun);
  $("flex-write").addEventListener("click", flexWriteSheets);

  $("flex-data-sheet").addEventListener("change", (e) => flexShowSheet(e.target.value));
  $("flex-data-open").addEventListener("click", () => {
    if (!flexState.activeSheetId) return;
    renderFlexData();
    flexSwitchTab("flex-data-panel");
  });

  // Colour picker popover.
  const colorsBtn = $("flex-colors-btn");
  const colorsMenu = $("flex-colors-menu");
  colorsBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    colorsMenu.classList.toggle("hidden");
  });
  document.addEventListener("click", (e) => {
    if (!colorsMenu.contains(e.target) && e.target !== colorsBtn) colorsMenu.classList.add("hidden");
  });
  colorsMenu.querySelectorAll('input[type="color"]').forEach((input) => {
    input.value = flexState.colors[input.dataset.status];
    input.addEventListener("input", () => {
      flexState.colors[input.dataset.status] = input.value;
      flexApplyColors();
    });
  });
  colorsMenu.querySelectorAll('input[name="flex-highlight"]').forEach((radio) => {
    radio.checked = radio.value === flexState.highlightMode;
    radio.addEventListener("change", () => {
      if (!radio.checked) return;
      flexState.highlightMode = radio.value;
      if (flexState.activeSheetId) renderFlexData();
    });
  });

  // Row numbers in the results open that sheet in the pane; row numbers in the
  // Data view select the real row in Excel.
  $("flex-results-body").addEventListener("click", (e) => {
    const cell = e.target.closest(".flex-jump");
    if (!cell) return;
    flexShowSheet(cell.dataset.sheet, parseInt(cell.dataset.row, 10));
  });
  $("flex-data-grid").addEventListener("click", (e) => {
    const cell = e.target.closest(".flex-jump");
    if (!cell) return;
    flexSelectInExcel(cell.dataset.sheet, parseInt(cell.dataset.row, 10));
  });
}
