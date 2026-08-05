"use strict";

/**
 * sheets.js — turn a reconcile result into "sheet specs": plain descriptions of
 * each output worksheet (values + which cells to colour). Ported from the web
 * app's export.js, but where export.js built SheetJS worksheets, this returns
 * data that taskpane.js writes into the live workbook via Office.js.
 *
 * The outcome is carried by colour, not words: the amount cell of every row is
 * filled green (matched), amber (matched, description differs) or red (not
 * found / no amount). The only text added to a source sheet is the row number
 * the match was found on.
 *
 * A spec: { name, aoa, colWidths, bandRows[], titleRows[], paintRects[],
 *           autofilter:{headerRow,width,lastRow}|null }
 *
 * Output sheets are prefixed so they never collide with the user's input sheets.
 */

// NB: Excel worksheet names forbid the characters  : \ / ? * [ ]  — so this
// prefix must not contain a colon (an earlier "Recon: " made every add() throw
// "The argument is invalid, missing or has an incorrect format.").
const RESULT_PREFIX = "Recon - ";

// The three outcome colours, as Excel's own "good / neutral / bad" pair of
// fill + font so they read the same as the built-in cell styles.
// Orange and blue are only used by colour-only mode, where all three sheets are
// present and a cashbook row can be found on one side but not the other.
const STATUS_FILL = {
  green:  { fill: "C6EFCE", font: "006100" },
  amber:  { fill: "FFEB9C", font: "9C6500" },
  red:    { fill: "FFC7CE", font: "9C0006" },
  orange: { fill: "F8CBAD", font: "833C0C" },
  blue:   { fill: "BDD7EE", font: "1F4E79" },
};

function _statusColor(status) {
  if (status === "Matched") return "green";
  if (status === "Check description") return "amber";
  if (status === "Not found" || status === "No amount") return "red";
  return null;
}

function _combineColors(colors) {
  const c = colors.filter(Boolean);
  if (!c.length) return null;
  if (c.every(x => x === "green")) return "green";
  if (c.every(x => x === "red")) return "red";
  return "amber";
}

/**
 * Collects coloured cells and hands back rectangles. Consecutive cells of one
 * colour in the same column merge into a single rectangle, so colouring a long
 * sheet costs a handful of Office.js range operations rather than thousands.
 */
function _painter() {
  const byCol = new Map();  // column index -> Map(row index -> colour)
  return {
    set(r, c, color) {
      if (!color || c === undefined || c === null || c < 0) return;
      if (!byCol.has(c)) byCol.set(c, new Map());
      byCol.get(c).set(r, color);
    },
    rects() {
      const out = [];
      for (const [c, rows] of byCol) {
        let run = null;
        const flush = () => {
          if (run) out.push({
            r0: run.r0, c0: c, rows: run.n, cols: 1,
            fill: STATUS_FILL[run.color].fill, font: STATUS_FILL[run.color].font,
          });
          run = null;
        };
        for (const r of [...rows.keys()].sort((a, b) => a - b)) {
          const color = rows.get(r);
          if (run && run.color === color && run.r0 + run.n === r) { run.n++; continue; }
          flush();
          run = { r0: r, n: 1, color };
        }
        flush();
      }
      return out;
    },
  };
}

// Drop columns that hold nothing but auto-added "=compareTo..." helper formulas.
// Returns the trimmed rows plus old-column-index -> new-column-index, since the
// amount columns to colour are recorded against the untrimmed sheet.
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
  const index = new Map(keep.map((c, i) => [c, i]));
  return {
    rows: rows.map(r => keep.map(c => (r[c] !== undefined && r[c] !== null) ? r[c] : "")),
    index,
  };
}

/* ---------- per-file sheet (source rows, amounts coloured) ---------- */

/**
 * A copy of one source sheet with the matched row number(s) appended per side,
 * its amount cell(s) coloured by outcome and each reference cell coloured by
 * that side's own outcome.
 *
 *  refHeaders   short column titles, one per side ("BS", "GL", "CB")
 *  refsForRow   1-based sheet row -> array of cell values, one per side
 *  colorsForRow 1-based sheet row -> array of colours, one per side
 */
function _sideSpec(name, source, refHeaders, refsForRow, colorsForRow) {
  const { rows: data, index } = _withoutFormulaColumns(source.rows);
  let width = 0;
  for (const r of data) width = Math.max(width, r.length);

  const amountCols = (source.amountCols || [])
    .map(c => index.get(c))
    .filter(c => c !== undefined);

  const paint = _painter();
  const aoa = data.map((r, i) => {
    const row = [];
    for (let c = 0; c < width; c++) row.push(r[c] !== undefined ? r[c] : "");
    if (source.headerRow && i + 1 === source.headerRow) {
      row.push(...refHeaders);
      return row;
    }
    row.push(...refsForRow(i + 1));
    const colors = colorsForRow(i + 1);
    for (const c of amountCols) paint.set(i, c, _combineColors(colors));
    colors.forEach((color, k) => paint.set(i, width + k, color));
    return row;
  });

  const totalWidth = width + refHeaders.length;
  const colWidths = Array.from({ length: width }, () => 16).concat(refHeaders.map(() => 10));
  const bandRows = source.headerRow ? [source.headerRow - 1] : [];
  const autofilter = source.headerRow
    ? { headerRow: source.headerRow - 1, width: totalWidth, lastRow: aoa.length - 1 }
    : null;

  return {
    name: RESULT_PREFIX + name, aoa, colWidths, bandRows,
    titleRows: [], paintRects: paint.rects(), autofilter,
  };
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
    return _sideSpec(name, source, ["CB"],
      (rowNum) => [matchedBy.has(rowNum) ? (cbRefs.get(rowNum) || []).join(", ") : ""],
      (rowNum) => [matchedBy.has(rowNum) ? (matchedBy.get(rowNum) ? "green" : "red") : null]);
  };

  if (hasStatement && sources.statement) {
    specs.push(reverseSpec("Bank Statement", sources.statement, result.statement, r => r.matchedStatementRows));
  }

  if (sources.cashbook) {
    const byRow = new Map(result.rows.map(r => [r.row, r]));
    const cbHeaders = [];
    if (hasStatement) cbHeaders.push("BS");
    if (hasLedger) cbHeaders.push("GL");
    specs.push(_sideSpec("Cashbook", sources.cashbook, cbHeaders, (rowNum) => {
      const r = byRow.get(rowNum);
      const out = [];
      if (hasStatement) out.push(r ? r.matchedStatementRows.join(", ") : "");
      if (hasLedger) out.push(r ? r.matchedLedgerRows.join(", ") : "");
      return out;
    }, (rowNum) => {
      const r = byRow.get(rowNum);
      const out = [];
      if (hasStatement) out.push(r ? _statusColor(r.status) : null);
      if (hasLedger) out.push(r ? _statusColor(r.ledgerStatus) : null);
      return out;
    }));
  }

  if (hasLedger && sources.ledger) {
    specs.push(reverseSpec("General Ledger", sources.ledger, result.ledger, r => r.matchedLedgerRows));
  }

  // --- Comparison: cashbook / statement (/ ledger) rows side by side ---
  const compHeaders = ["CB Row", "Date", "Description", "Amount"];
  if (hasStatement) compHeaders.push("BS Row", "Date", "Description", "Amount");
  if (hasLedger) compHeaders.push("GL Row", "Date", "Description", "Amount");
  const stAmountCol = 7;                            // 4 cashbook columns + 3
  const ldAmountCol = hasStatement ? 11 : 7;
  const compAoa = [compHeaders];
  const compPaint = _painter();
  let compIdx = 1;
  for (const c of buildComparisonRows(result)) {
    const sideVals = (s) => s ? [s.row, s.date, s.description, s.amount] : ["", "", "", ""];
    const row = [...sideVals(c.cb)];
    const colors = [];
    if (hasStatement) {
      row.push(...sideVals(c.st));
      const color = _statusColor(c.stStatus);
      colors.push(color);
      if (c.st) compPaint.set(compIdx, stAmountCol, color);
    }
    if (hasLedger) {
      row.push(...sideVals(c.ld));
      const color = _statusColor(c.ldStatus);
      colors.push(color);
      if (c.ld) compPaint.set(compIdx, ldAmountCol, color);
    }
    compAoa.push(row);
    if (c.cb) compPaint.set(compIdx, 3, _combineColors(colors));
    compIdx++;
  }
  const sideCols = () => [8, 12, 40, 12];
  const compCols = sideCols();
  if (hasStatement) compCols.push(...sideCols());
  if (hasLedger) compCols.push(...sideCols());
  specs.push({
    name: RESULT_PREFIX + "Comparison", aoa: compAoa, colWidths: compCols,
    bandRows: [0], titleRows: [], paintRects: compPaint.rects(),
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
    bandRows: [1], titleRows, paintRects: [], autofilter: null
  });

  return specs;
}
