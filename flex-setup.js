"use strict";

/**
 * flex-setup.js — entering Modular Recon and step 1 of it: the sheets.
 *
 * Step 1 asks the only question the tool genuinely cannot answer on its own —
 * which row holds each sheet's column names — and lets the user drop the sheets
 * that aren't part of the job. Everything else (which columns exist, what they
 * hold, how they should be compared) follows from that one answer, which is why
 * it is a wizard step of its own rather than a setting buried in a menu.
 *
 * A header row is guessed on load, so the common case is "look, then Continue".
 *
 * Where the web app uploads workbooks, here every candidate sheet is already in
 * front of the user: the worksheets of the open workbook, read through
 * Office.js. That reading is the only Excel-aware code in the flex-* files.
 *
 * This file also owns initFlex(): all Modular Recon wiring lives in one place,
 * called once from taskpane.js.
 */

const FLEX_PREVIEW_ROWS = 10;
const FLEX_PREVIEW_COLS = 10;

/* ---------- mode switching ---------- */

// Modular Recon lives in its own <section>. `body.flex-mode` decides which
// section is on screen (see taskpane.css), so the classic flow keeps its own
// DOM and state untouched behind it.
function flexModeActive() {
  return document.body.classList.contains("flex-mode");
}

function flexEnterMode() {
  document.body.classList.add("flex-mode");
  document.querySelectorAll("#mode-switch button").forEach((b) => b.classList.toggle("active", b.dataset.mode === "flex"));
  setStatus("");
  renderFlexSetup();
  // Coming back to a canvas that was already open: its wires are measured from
  // the DOM, which had no sizes while the section was hidden, so redraw now.
  if (!$("flex-model-panel").classList.contains("hidden")) renderFlexErd();
  flexUpdateButtons();
}

function flexExitMode() {
  document.body.classList.remove("flex-mode");
  document.querySelectorAll("#mode-switch button").forEach((b) => b.classList.toggle("active", b.dataset.mode === "classic"));
  setStatus("");
}

// The flex panels share one tab strip; the Data panel is reached from the
// sheet picker rather than a tab of its own.
function flexSwitchTab(panelId) {
  document.querySelectorAll("#flex .tab-panel").forEach((p) => p.classList.toggle("hidden", p.id !== panelId));
  document.querySelectorAll("#flex-tabs button").forEach((t) => t.classList.toggle("active", t.dataset.panel === panelId));
}

// Toolbar state: Run needs at least one drawn comparison, Write needs a result.
function flexUpdateButtons() {
  const runBtn = $("flex-run");
  const writeBtn = $("flex-write");
  const toModel = $("flex-to-model");
  if (runBtn) runBtn.disabled = flexRelationships().length === 0;
  if (writeBtn) writeBtn.disabled = !flexState.result;
  if (toModel) toModel.disabled = flexIncludedSheets().length < 2;
  $("flex-tab-results").classList.toggle("hidden", !flexState.result);
  $("flex-data-nav").classList.toggle("hidden", !flexState.result);
}

/* ---------- reading the workbook ---------- */

// Every worksheet's used range, with date-formatted cells turned back into real
// Dates. Office.js reports a date cell as a serial number, which every parser
// here would otherwise read as an ordinary number — the cell's number format is
// the only thing that says it is a date, so it is what decides.
async function flexReadWorkbook() {
  const out = [];
  await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    const wanted = sheets.items
      .map((s) => s.name)
      .filter((n) => !n.startsWith(RESULT_PREFIX) && !n.startsWith(FLEX_PREFIX));

    const ranges = wanted.map((name) => {
      const used = ctx.workbook.worksheets.getItem(name).getUsedRangeOrNullObject(true);
      used.load(["values", "numberFormat"]);
      return { name, used };
    });
    await ctx.sync();

    for (const { name, used } of ranges) {
      if (used.isNullObject) { out.push({ name, rows: [] }); continue; }
      const values = used.values, formats = used.numberFormat || [];
      out.push({
        name,
        rows: values.map((row, r) => row.map((v, c) => {
          if (v === null || v === "") return null;
          if (typeof v === "number" && isExcelDateFormat((formats[r] || [])[c])) return excelSerialToDate(v);
          return v;
        })),
      });
    }
  });
  return out;
}

/**
 * Load (or reload) the workbook's sheets.
 *
 * A reload keeps everything the user has told us — header row, whether the
 * sheet is included, where its box sits, and the links drawn to it — and just
 * refreshes the values underneath. Links pointing past a sheet's new column
 * list are dropped, the same rule as changing a header row.
 */
async function flexLoadWorkbook() {
  setStatus("Reading the worksheets…");
  try {
    const read = await flexReadWorkbook();
    const seen = new Set();

    for (const { name, rows } of read) {
      seen.add(name);
      const existing = flexState.sheets.find((s) => s.name === name);
      if (existing) {
        existing.rows = flexTrim(rows);
        flexRebuildColumns(existing);
        flexDropStaleLinks(existing);
      } else {
        flexAddSheet(name, rows);
      }
    }
    // A worksheet that has since been deleted or renamed takes its links with it.
    for (const sheet of [...flexState.sheets]) {
      if (!seen.has(sheet.name)) flexRemoveSheet(sheet.id);
    }

    flexState.result = null;
    flexAutoLayout();
    renderFlexSetup();
    flexRenderSheetPicker();
    flexUpdateButtons();
    const n = flexIncludedSheets().length;
    setStatus(n < 2
      ? "Include at least two sheets to draw a comparison."
      : `${n} sheets ready — check each header row, then Continue.`);
  } catch (e) {
    setStatus("Could not read the worksheets: " + e.message, true);
  }
}

// Links can outlive the column they point at when a sheet gets narrower.
function flexDropStaleLinks(sheet) {
  const width = sheet.columns.length;
  flexState.links = flexState.links.filter((l) =>
    !((l.from.sheet === sheet.id && l.from.col >= width) || (l.to.sheet === sheet.id && l.to.col >= width)));
}

/* ---------- step 1: the sheet cards ---------- */

function renderFlexSetup() {
  const body = $("flex-setup-body");
  if (!body) return;

  if (!flexState.sheets.length) {
    body.innerHTML = `<p class="flex-empty">Nothing loaded yet. <b>Load worksheets</b> pulls in every sheet of this
      workbook — pick the ones you want and say where each one's column names are.</p>`;
    return;
  }
  body.replaceChildren(...flexState.sheets.map(flexSheetCard));
}

function flexSheetCard(sheet) {
  const card = document.createElement("div");
  card.className = `flex-sheet-card${sheet.skip ? " skipped" : ""}`;
  card.dataset.sheet = sheet.id;

  const rowCount = Math.max(0, sheet.rows.length - sheet.headerRow);
  const colCount = sheet.columns.length;

  const head = document.createElement("div");
  head.className = "flex-sheet-head";
  head.innerHTML = `
    <label class="flex-include" title="Uncheck to leave this sheet out of the reconciliation">
      <input type="checkbox" ${sheet.skip ? "" : "checked"}> Include
    </label>
    <div class="flex-sheet-title">
      ${escapeHtml(sheet.name)}
      <span class="flex-sheet-sub">${rowCount} row${rowCount === 1 ? "" : "s"} · ${colCount} column${colCount === 1 ? "" : "s"}</span>
    </div>`;
  card.appendChild(head);

  head.querySelector("input").addEventListener("change", (e) => {
    sheet.skip = !e.target.checked;
    flexState.result = null;
    renderFlexSetup();
    flexUpdateButtons();
  });

  card.appendChild(flexHeaderPicker(sheet));

  const foot = document.createElement("div");
  foot.className = "flex-sheet-foot";
  const names = sheet.columns.slice(0, 10).map((c) =>
    `<span class="flex-col-chip" title="${escapeHtml(c.type)}${c.samples.length ? " · e.g. " + escapeHtml(c.samples.join(", ")) : ""}">${escapeHtml(c.name)}<i>${c.type[0]}</i></span>`);
  if (sheet.columns.length > 10) names.push(`<span class="flex-col-chip more">+${sheet.columns.length - 10} more</span>`);
  foot.innerHTML = `
    <div class="flex-foot-line">
      <span class="flex-foot-label">${sheet.headerRow ? `Column names from row ${sheet.headerRow}` : "No header row — columns are A, B, C…"}</span>
      <button type="button" class="flex-link-btn flex-no-header">${sheet.headerRow ? "No header row" : "Undo"}</button>
    </div>
    <div class="flex-col-chips">${names.join("")}</div>`;
  card.appendChild(foot);

  foot.querySelector(".flex-no-header").addEventListener("click", () => {
    flexSetHeaderRow(sheet, sheet.headerRow ? 0 : flexGuessHeaderRow(sheet.rows));
  });

  return card;
}

// The preview: the first rows of the sheet with a clickable row-number gutter.
// Clicking a row says "these are my column names" — the whole of step 1.
function flexHeaderPicker(sheet) {
  const wrap = document.createElement("div");
  wrap.className = "flex-preview-wrap";

  const shown = Math.min(sheet.rows.length, Math.max(FLEX_PREVIEW_ROWS, sheet.headerRow + 2));
  const width = Math.min(sheet.rows.reduce((max, r) => Math.max(max, r.length), 0), FLEX_PREVIEW_COLS);

  const table = document.createElement("table");
  table.className = "flex-preview";
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr><th class="rownum"></th>${
    Array.from({ length: width }, (_, i) => `<th>${colLetter(i)}</th>`).join("")}</tr>`;
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (let r = 0; r < shown; r++) {
    const tr = document.createElement("tr");
    tr.dataset.row = r + 1;
    if (r + 1 === sheet.headerRow) tr.className = "is-header";
    else if (sheet.headerRow && r + 1 < sheet.headerRow) tr.className = "above-header";
    const cells = sheet.rows[r] || [];
    tr.innerHTML = `<td class="rownum">${r + 1}</td>${
      Array.from({ length: width }, (_, c) => `<td>${escapeHtml(_display_date(cells[c] ?? ""))}</td>`).join("")}`;
    tr.title = "Click to use this row as the column names";
    tr.addEventListener("click", () => flexSetHeaderRow(sheet, r + 1));
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  wrap.appendChild(table);
  return wrap;
}

// Changing the header row changes what the columns are, so any link pointing
// past the new column list has to go — silently keeping a link to a column that
// no longer exists would produce a recon nobody drew.
function flexSetHeaderRow(sheet, headerRow) {
  sheet.headerRow = headerRow;
  flexRebuildColumns(sheet);
  flexDropStaleLinks(sheet);
  flexState.result = null;
  renderFlexSetup();
  flexUpdateButtons();
}

/* ---------- wiring (called once from taskpane.js) ---------- */

function initFlex() {
  document.querySelectorAll("#mode-switch button").forEach((btn) => {
    btn.addEventListener("click", () => (btn.dataset.mode === "flex" ? flexEnterMode() : flexExitMode()));
  });

  $("flex-load").addEventListener("click", flexLoadWorkbook);
  $("flex-to-model").addEventListener("click", () => {
    flexSwitchTab("flex-model-panel");
    renderFlexErd();
  });

  // Re-draw the canvas whenever its tab is opened: entity boxes are measured
  // from the DOM, which only has real sizes once the panel is visible.
  document.querySelectorAll("#flex-tabs button").forEach((tab) => {
    tab.addEventListener("click", () => {
      flexSwitchTab(tab.dataset.panel);
      if (tab.dataset.panel === "flex-model-panel") renderFlexErd();
    });
  });

  initFlexErd();
  initFlexResults();
  renderFlexSetup();
}
