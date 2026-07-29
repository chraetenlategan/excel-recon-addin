"use strict";

/**
 * taskpane.js — the Office.js glue.
 *
 * This is the ONLY file that knows about Excel. It:
 *   1. lists the workbook's worksheets into three dropdowns,
 *   2. reads the chosen sheets' used ranges as rows (arrays of arrays),
 *   3. hands those rows to the ported engine (runLocalAnalyze / runLocalReconcile),
 *   4. writes the result rows back into a "Recon Results" worksheet.
 *
 * The engine itself (engine.js / utils.js) is untouched from the web app — it
 * neither knows nor cares that the rows came from a live sheet instead of a
 * dropped .xlsx.
 */

const $ = (id) => document.getElementById(id);
const RESULT_SHEET = "Recon Results";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in only runs in Excel.", true);
    return;
  }
  $("refresh-sheets").onclick = loadSheetList;
  $("reconcile").onclick = reconcile;
  ["sel-cashbook", "sel-statement", "sel-ledger"].forEach((id) =>
    $(id).addEventListener("change", updateEnabled));
  loadSheetList();
});

function setStatus(msg, isError) {
  const el = $("status");
  el.textContent = msg || "";
  el.classList.toggle("err", !!isError);
}

// A cashbook sheet plus at least one of statement/ledger is required.
function updateEnabled() {
  const cb = $("sel-cashbook").value;
  const other = $("sel-statement").value || $("sel-ledger").value;
  $("reconcile").disabled = !(cb && other);
}

/* ---------- 1. list worksheets ---------- */

async function loadSheetList() {
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      const names = sheets.items.map((s) => s.name).filter((n) => n !== RESULT_SHEET);

      const guess = (needles) =>
        names.find((n) => needles.some((k) => n.toLowerCase().includes(k))) || "";

      fillSelect("sel-cashbook", names, false, guess(["cashbook", "cash book", "cash"]));
      fillSelect("sel-statement", names, true, guess(["bank", "statement"]));
      fillSelect("sel-ledger", names, true, guess(["ledger", "gl", "general"]));
    });
    updateEnabled();
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

/* ---------- 2. read a sheet's used range as rows ---------- */

async function readSheet(ctx, name) {
  const ws = ctx.workbook.worksheets.getItem(name);
  const used = ws.getUsedRangeOrNullObject(true);
  used.load(["values", "rowCount"]);
  await ctx.sync();
  if (used.isNullObject) return [];
  // Excel gives dates as serial numbers; the engine's _display_date handles
  // both those and text dates, so we feed .values through unchanged.
  return used.values.map((row) => row.map((v) => (v === null ? "" : v)));
}

/* ---------- 3 + 4. reconcile and write back ---------- */

async function reconcile() {
  const cbName = $("sel-cashbook").value;
  const stName = $("sel-statement").value;
  const glName = $("sel-ledger").value;
  $("reconcile").disabled = true;
  setStatus("Reading sheets…");

  try {
    await Excel.run(async (ctx) => {
      const cbRows = await readSheet(ctx, cbName);
      const stRows = stName ? await readSheet(ctx, stName) : null;
      const glRows = glName ? await readSheet(ctx, glName) : null;

      if (!cbRows.length) throw new Error(`"${cbName}" looks empty.`);

      setStatus("Detecting columns…");
      const cbMap = await mapFor(cbRows);
      const stMap = stRows ? await mapFor(stRows) : null;
      const glMap = glRows ? await mapFor(glRows) : null;

      setStatus("Matching…");
      const result = runLocalReconcile(cbRows, stRows, cbMap, stMap, glRows, glMap);

      setStatus("Writing results…");
      await writeResults(ctx, result);
      await ctx.sync();

      showSummary(result);
      showMapping({ Cashbook: cbMap, "Bank statement": stMap, "General ledger": glMap });
      setStatus(`Done. See the "${RESULT_SHEET}" sheet.`);
    });
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    updateEnabled();
  }
}

// runLocalAnalyze returns the mapping without headerRow; _extract needs it on
// the mapping, so we merge it in (exactly what the web app's grid.js does).
async function mapFor(rows) {
  const analysis = await runLocalAnalyze(rows, [], "");
  return { ...analysis.mapping, headerRow: analysis.headerRow };
}

const STATUS_FILL = {
  "Matched": "#DFF6DD",
  "Check description": "#FFF4CE",
  "Not found": "#FDE7E9",
  "No amount": "#F3F2F1",
};

async function writeResults(ctx, result) {
  // Recreate the results sheet from scratch each run.
  const existing = ctx.workbook.worksheets.getItemOrNullObject(RESULT_SHEET);
  existing.load("name");
  await ctx.sync();
  if (!existing.isNullObject) existing.delete();
  const ws = ctx.workbook.worksheets.add(RESULT_SHEET);

  const withLedger = result.hasLedger;
  const header = ["CB row", "Date", "Description", "Amount"];
  if (result.hasStatement) header.push("Bank statement");
  if (withLedger) header.push("General ledger");

  const body = result.rows.map((r) => {
    const row = [r.row, r.date, r.description, r.amount === null ? "" : r.amount];
    if (result.hasStatement) row.push(r.statusLabel || "");
    if (withLedger) row.push(r.ledgerStatusLabel || "");
    return row;
  });

  const all = [header, ...body];
  const cols = header.length;
  const range = ws.getRangeByIndexes(0, 0, all.length, cols);
  range.values = all;

  // Header styling.
  const headRange = ws.getRangeByIndexes(0, 0, 1, cols);
  headRange.format.font.bold = true;
  headRange.format.fill.color = "#217346";
  headRange.format.font.color = "white";

  // Colour each status cell by outcome.
  const stCol = result.hasStatement ? 4 : -1;
  const glCol = withLedger ? (result.hasStatement ? 5 : 4) : -1;
  result.rows.forEach((r, i) => {
    if (stCol >= 0 && STATUS_FILL[r.status]) {
      ws.getRangeByIndexes(i + 1, stCol, 1, 1).format.fill.color = STATUS_FILL[r.status];
    }
    if (glCol >= 0 && STATUS_FILL[r.ledgerStatus]) {
      ws.getRangeByIndexes(i + 1, glCol, 1, 1).format.fill.color = STATUS_FILL[r.ledgerStatus];
    }
  });

  ws.getUsedRange().format.autofitColumns();
  ws.activate();
}

/* ---------- pane readouts ---------- */

function showSummary(result) {
  const s = result.summary;
  const box = $("summary");
  const chips = [
    ["matched", "Matched", s["Matched"]],
    ["check", "Check description", s["Check description"]],
    ["missing", "Not found", s["Not found"]],
    ["noamount", "No amount", s["No amount"]],
  ];
  box.innerHTML = chips
    .filter(([, , n]) => n > 0)
    .map(([cls, label, n]) => `<span class="chip ${cls}">${label}: ${n}</span>`)
    .join("");
  box.classList.remove("hidden");
}

function showMapping(maps) {
  const roleName = { date: "Date", description: "Description", amount: "Amount", debit: "Debit", credit: "Credit" };
  const rows = [];
  for (const [side, m] of Object.entries(maps)) {
    if (!m) continue;
    const parts = [];
    for (const role of ["date", "description", "amount", "debit", "credit"]) {
      if (m[role] !== null && m[role] !== undefined) {
        parts.push(`${roleName[role]} = ${colLetter(m[role])}`);
      }
    }
    const hdr = m.headerRow ? ` (header row ${m.headerRow})` : "";
    rows.push(`<tr><td>${side}</td><td>${parts.join(", ")}${hdr}</td></tr>`);
  }
  const box = $("mapping");
  box.innerHTML = `<h3>Detected columns</h3><table>${rows.join("")}</table>` +
    `<p>Wrong? Rename headers (Date / Description / Amount, or Debit + Credit) and reconcile again.</p>`;
  box.classList.remove("hidden");
}
