"use strict";

/**
 * sheets.js — turn a reconcile result into "sheet specs": plain descriptions of
 * each output worksheet (values + which rows to colour). Ported from the web
 * app's export.js, but where export.js built SheetJS worksheets, this returns
 * data that taskpane.js writes into the live workbook via Office.js.
 *
 * A spec: { name, aoa, colWidths, bandRows[], titleRows[], rowFills{idx:color},
 *           autofilter:{headerRow,width,lastRow}|null }
 * color is "green" | "amber" | "red".
 *
 * Output sheets are prefixed so they never collide with the user's input sheets.
 */

const RESULT_PREFIX = "Recon: ";

/* ---------- status text + colour helpers (from export.js) ---------- */

function _plainStatus(status, target) {
  if (status === "Matched") return `Matched on ${target}`;
  if (status === "Check description") return `Matched on ${target} - check description`;
  if (status === "Not found") return `Not matched on ${target}`;
  if (status === "No amount") return "No amount";
  return "";
}

function _statusColor(status) {
  if (status === "Matched") return "green";
  if (status === "Check description") return "amber";
  if (status === "Not found" || status === "No amount") return "red";
  return null;
}

function _classColor(cls) {
  if (cls === "status-matched") return "green";
  if (cls === "status-check") return "amber";
  if (cls === "status-missing" || cls === "status-noamount") return "red";
  return null;
}

function _combineColors(colors) {
  const c = colors.filter(Boolean);
  if (!c.length) return null;
  if (c.every(x => x === "green")) return "green";
  if (c.every(x => x === "red")) return "red";
  return "amber";
}

// Drop columns that hold nothing but auto-added "=compareTo..." helper formulas.
function _withoutFormulaColumns(rows) {
  let width = 0;
  for (const r of rows) width = Math.max(width, r.length);
  const keep = [];
  for (let c = 0; c < width; c++) {
    let hasContent = false, onlyFormulas = true;
    for (const r of rows) {
      const v = String(r[c] ?? "").trim();
      if (!v) continue;
      hasContent = true;
      if (!v.toLowerCase().startsWith("=compareto")) { onlyFormulas = false; break; }
    }
    if (!(hasContent && onlyFormulas)) keep.push(c);
  }
  return rows.map(r => keep.map(c => (r[c] !== undefined && r[c] !== null) ? r[c] : ""));
}

/* ---------- per-file sheet (uploaded rows + appended statuses) ---------- */

function _appendStatusSpec(name, source, statusHeaders, statusForRow, colorForRow) {
  const data = _withoutFormulaColumns(source.rows);
  let width = 0;
  for (const r of data) width = Math.max(width, r.length);

  const rowFills = {};
  const aoa = data.map((r, i) => {
    const row = [];
    for (let c = 0; c < width; c++) row.push(r[c] !== undefined ? r[c] : "");
    if (source.headerRow && i + 1 === source.headerRow) {
      row.push(...statusHeaders);
    } else {
      row.push(...statusForRow(i + 1));
      const color = colorForRow ? colorForRow(i + 1) : null;
      if (color) rowFills[i] = color;
    }
    return row;
  });

  const totalWidth = width + statusHeaders.length;
  const colWidths = Array.from({ length: width }, () => 16).concat(statusHeaders.map(() => 30));
  const bandRows = source.headerRow ? [source.headerRow - 1] : [];
  const autofilter = source.headerRow
    ? { headerRow: source.headerRow - 1, width: totalWidth, lastRow: aoa.length - 1 }
    : null;

  return { name: RESULT_PREFIX + name, aoa, colWidths, bandRows, titleRows: [], rowFills, autofilter };
}

/* ---------- top-level: build every output sheet ---------- */

function buildResultSheets(result) {
  const specs = [];
  const hasLedger = result.hasLedger;
  const hasStatement = result.hasStatement !== false;
  const sources = result.sources || {};

  // Reverse direction: which cashbook row(s) consumed each statement/ledger row.
  const reverseSpec = (name, source, sideRows, matchedRowsOf) => {
    const matchedBy = new Map(sideRows.map(s => [s.row, s.matched]));
    const cbRefs = new Map();
    for (const r of result.rows) {
      for (const n of matchedRowsOf(r)) {
        if (!cbRefs.has(n)) cbRefs.set(n, []);
        cbRefs.get(n).push(r.row);
      }
    }
    return _appendStatusSpec(name, source, ["Cashbook Status", "Cashbook Rows"], (rowNum) => {
      if (!matchedBy.has(rowNum)) return ["", ""];
      return [matchedBy.get(rowNum) ? "Matched on cashbook" : "Not matched on cashbook",
              (cbRefs.get(rowNum) || []).join(", ")];
    }, (rowNum) => {
      if (!matchedBy.has(rowNum)) return null;
      return matchedBy.get(rowNum) ? "green" : "red";
    });
  };

  if (hasStatement && sources.statement) {
    specs.push(reverseSpec("Bank Statement", sources.statement, result.statement, r => r.matchedStatementRows));
  }

  if (sources.cashbook) {
    const byRow = new Map(result.rows.map(r => [r.row, r]));
    const cbHeaders = [];
    if (hasStatement) cbHeaders.push("Bank Statement Status", "Statement Rows");
    if (hasLedger) cbHeaders.push("General Ledger Status", "Ledger Rows");
    specs.push(_appendStatusSpec("Cashbook", sources.cashbook, cbHeaders, (rowNum) => {
      const r = byRow.get(rowNum);
      const out = [];
      if (hasStatement) out.push(r ? _plainStatus(r.status, "bank statement") : "",
                                 r ? r.matchedStatementRows.join(", ") : "");
      if (hasLedger) out.push(r ? _plainStatus(r.ledgerStatus, "general ledger") : "",
                              r ? r.matchedLedgerRows.join(", ") : "");
      return out;
    }, (rowNum) => {
      const r = byRow.get(rowNum);
      if (!r) return null;
      const cs = [];
      if (hasStatement) cs.push(_statusColor(r.status));
      if (hasLedger) cs.push(_statusColor(r.ledgerStatus));
      return _combineColors(cs);
    }));
  }

  if (hasLedger && sources.ledger) {
    specs.push(reverseSpec("General Ledger", sources.ledger, result.ledger, r => r.matchedLedgerRows));
  }

  // --- Comparison: cashbook / statement (/ ledger) rows side by side ---
  const compHeaders = ["CB Row", "CB Date", "CB Description", "CB Amount"];
  if (hasStatement) compHeaders.push("BS Row", "BS Date", "BS Description", "BS Amount", "Status");
  if (hasLedger) compHeaders.push("GL Row", "GL Date", "GL Description", "GL Amount", "GL Status");
  const compAoa = [compHeaders];
  const compFills = {};
  let compIdx = 1;
  for (const c of buildComparisonRows(result)) {
    const sideVals = (s) => s ? [s.row, s.date, s.description, s.amount] : ["", "", "", ""];
    const row = [...sideVals(c.cb)];
    const cs = [];
    if (hasStatement) { row.push(...sideVals(c.st), c.stStatusLabel); cs.push(_classColor(c.stStatusClass)); }
    if (hasLedger) { row.push(...sideVals(c.ld), c.ldStatusLabel); cs.push(_classColor(c.ldStatusClass)); }
    compAoa.push(row);
    const color = _combineColors(cs);
    if (color) compFills[compIdx] = color;
    compIdx++;
  }
  const sideCols = () => [8, 12, 40, 12];
  const compCols = sideCols();
  if (hasStatement) compCols.push(...sideCols(), 28);
  if (hasLedger) compCols.push(...sideCols(), 28);
  specs.push({
    name: RESULT_PREFIX + "Comparison", aoa: compAoa, colWidths: compCols,
    bandRows: [0], titleRows: [], rowFills: compFills,
    autofilter: { headerRow: 0, width: compHeaders.length, lastRow: compAoa.length - 1 }
  });

  // --- Unmatched Comparison: listing + correlations + leftovers ---
  const unmatchedAoa = unmatchedComparisonSheetRows(result);
  const titleRows = [0];
  unmatchedAoa.forEach((row, i) => {
    if (i > 1 && row.length === 1 && String(row[0] ?? "").trim()) titleRows.push(i);
  });
  specs.push({
    name: RESULT_PREFIX + "Unmatched", aoa: unmatchedAoa,
    colWidths: [...sideCols(), ...sideCols(), 60, 12],
    bandRows: [1], titleRows, rowFills: {}, autofilter: null
  });

  return specs;
}
