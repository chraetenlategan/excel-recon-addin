"use strict";

/**
 * flex-model.js — the data model behind Modular Recon.
 *
 * Modular Recon is the format-free sibling of the classic three-sheet flow.
 * Instead of a fixed Date/Description/Amount layout it takes *any* worksheet,
 * asks only where each sheet's column names live, and then lets the user draw
 * the comparison themselves on an ERD-style canvas: every included sheet is an
 * entity, every column an attribute, and every line the user draws between two
 * attributes is one field of a composite key.
 *
 * This file owns the state and the pure model operations (ingest a worksheet,
 * guess the header row, describe columns, create/remove links, group links
 * into relationships). Matching lives in flex-engine.js, the wizard in
 * flex-setup.js, the canvas in flex-erd.js, and the output in flex-results.js.
 *
 * Ported from the static-reconciliation web app. The one real difference: the
 * web app ingests uploaded workbooks through SheetJS, while here every sheet
 * comes from the open workbook via Office.js (flex-setup.js does the reading).
 */

// Modular Recon writes its output to its own set of sheets, kept apart from the
// classic flow's "Recon - " ones so a run of either never eats the other's work.
// Excel worksheet names forbid  : \ / ? * [ ]  — hence the dash, not a colon.
const FLEX_PREFIX = "Modular - ";

const flexState = {
  sheets: [],            // FlexSheet[] — every worksheet pulled in from the workbook
  links: [],             // FlexLink[] — the attribute-to-attribute comparisons
  view: { x: 20, y: 20, scale: 0.85 },   // ERD canvas pan/zoom
  colors: { matched: "#12805c", unmatched: "#c02626", ambiguous: "#b57d10" },
  highlightMode: "cell", // "cell" = paint the compared values, "row" = the whole row
  result: null,          // last flexRunRecon() output
  selectedLink: null,    // link whose editor popover is open (drawn highlighted)
  activeSheetId: null,   // which sheet the Data tab is showing
  seq: 1,                // id counter for sheets and links
};

function flexId(prefix) {
  return `${prefix}${flexState.seq++}`;
}

/* ---------- sheet ingest ---------- */

/**
 * FlexSheet:
 *   id         stable id used by links and the DOM
 *   name       worksheet name
 *   label      display name (same as name — worksheet names are unique)
 *   rows       raw cell values, array of arrays (date cells already Date objects)
 *   headerRow  1-based row holding the column names; 0 = no header row
 *   skip       true when the user excluded this sheet from the recon
 *   columns    FlexColumn[] rebuilt whenever headerRow changes
 *   pos        {x,y} position of the entity box on the ERD canvas
 */
function flexAddSheet(name, rawRows) {
  const rows = flexTrim(rawRows || []);
  const sheet = {
    id: flexId("s"),
    name,
    label: name,
    rows,
    headerRow: flexGuessHeaderRow(rows),
    // An empty sheet has nothing to reconcile — start it excluded.
    skip: rows.length === 0,
    columns: [],
    pos: null,
  };
  flexRebuildColumns(sheet);
  flexState.sheets.push(sheet);
  return sheet;
}

// Drop trailing blank rows/columns but keep the original values (numbers and
// Dates all stay as-is so the parsers in utils.js can read them).
function flexTrim(rows) {
  let lastRow = -1, lastCol = -1;
  rows.forEach((row, r) => (row || []).forEach((value, c) => {
    if (String(value ?? "").trim() !== "") {
      if (r > lastRow) lastRow = r;
      if (c > lastCol) lastCol = c;
    }
  }));
  if (lastRow < 0) return [];
  return rows.slice(0, lastRow + 1).map((row) => {
    const out = [];
    for (let c = 0; c <= lastCol; c++) out.push((row || [])[c] ?? null);
    return out;
  });
}

function flexRemoveSheet(sheetId) {
  flexState.links = flexState.links.filter((l) => l.from.sheet !== sheetId && l.to.sheet !== sheetId);
  flexState.sheets = flexState.sheets.filter((s) => s.id !== sheetId);
  if (flexState.activeSheetId === sheetId) flexState.activeSheetId = null;
}

function flexSheet(sheetId) {
  return flexState.sheets.find((s) => s.id === sheetId) || null;
}

function flexIncludedSheets() {
  return flexState.sheets.filter((s) => !s.skip);
}

/* ---------- header row detection ---------- */

// The column-name row is the one that reads like labels: several short, non
// numeric, distinct cells, with real data underneath it. Scored over the first
// 15 rows; 1 (the first row) is the fallback, which is right most of the time.
function flexGuessHeaderRow(rows) {
  if (!rows.length) return 0;
  let best = 0, bestScore = 0;
  const limit = Math.min(15, rows.length - 1);
  for (let r = 0; r < limit; r++) {
    const cells = (rows[r] || []).map((v) => _text(v)).filter(Boolean);
    if (cells.length < 2) continue;
    const labels = cells.filter((t) => t.length <= 40 && _amount(t) === null && !_is_date_like(t));
    const distinct = new Set(labels.map((t) => t.toLowerCase())).size;
    // Reward label-ish rows, and rows whose width matches the data below them.
    const below = (rows[r + 1] || []).filter((v) => _text(v) !== "").length;
    const score = distinct * 2 + (below >= cells.length - 1 ? 2 : 0) + (labels.length === cells.length ? 1 : 0);
    if (score > bestScore) { bestScore = score; best = r + 1; }
  }
  return bestScore > 0 ? best : 1;
}

/* ---------- column description ---------- */

// FlexColumn: { index, letter, name, type, dateOrder, samples }
// `type` is what the column mostly holds — it drives the default compare mode
// and the little tag shown on the ERD attribute.
function flexRebuildColumns(sheet) {
  const width = sheet.rows.reduce((max, r) => Math.max(max, r.length), 0);
  const headers = sheet.headerRow ? (sheet.rows[sheet.headerRow - 1] || []) : [];
  const start = sheet.headerRow;
  sheet.columns = [];
  for (let i = 0; i < width; i++) {
    const values = [];
    for (let r = start; r < sheet.rows.length && values.length < 200; r++) {
      const v = sheet.rows[r] ? sheet.rows[r][i] : null;
      if (_text(v) !== "") values.push(v);
    }
    const name = _text(headers[i]) || colLetter(i);
    sheet.columns.push({
      index: i,
      letter: colLetter(i),
      name,
      type: flexColumnType(values),
      dateOrder: flexDateOrder(values),
      samples: values.slice(0, 3).map((v) => _display_date(v)),
      filled: values.length,
    });
  }
}

// Majority vote over the column's populated cells.
function flexColumnType(values) {
  if (!values.length) return "text";
  let dates = 0, numbers = 0;
  for (const v of values) {
    if (v instanceof Date || _is_date_like(v)) dates++;
    else if (_amount(v) !== null) numbers++;
  }
  if (dates >= values.length / 2) return "date";
  if (numbers >= values.length / 2) return "number";
  return "text";
}

// Which way round a numeric date column reads. Matching needs ONE canonical key
// per value, so we settle the D/M vs M/D question once for the whole column
// (a single "13/07/2024" anywhere in it proves the day comes first) instead of
// carrying every plausible reading the way the classic engine does.
function flexDateOrder(values) {
  let dmy = 0, mdy = 0, ymd = 0;
  for (const v of values) {
    if (v instanceof Date || (typeof v === "number" && v > 10000)) { ymd++; continue; }
    const parts = String(_display_date(v)).trim().split(/[-/.\s]+/).filter(Boolean);
    if (parts.length < 3 || !parts.slice(0, 3).every((p) => /^\d+$/.test(p))) continue;
    const [a, b] = parts.map(Number);
    if (parts[0].length === 4) ymd++;
    else if (a > 12) dmy++;
    else if (b > 12) mdy++;
  }
  if (ymd > dmy && ymd > mdy) return "ymd";
  if (mdy > dmy) return "mdy";
  return "dmy"; // day-first is the local convention and the safer default
}

function flexColumn(sheetId, index) {
  const sheet = flexSheet(sheetId);
  return sheet ? sheet.columns[index] || null : null;
}

function flexColumnName(sheetId, index) {
  const col = flexColumn(sheetId, index);
  return col ? col.name : colLetter(index);
}

// Rows that actually take part: everything below the header row that isn't blank.
// Returns [{ row (1-based sheet row), cells }].
function flexDataRows(sheet) {
  const out = [];
  for (let r = sheet.headerRow; r < sheet.rows.length; r++) {
    const cells = sheet.rows[r] || [];
    if (!cells.some((v) => _text(v) !== "")) continue;
    out.push({ row: r + 1, cells });
  }
  return out;
}

/* ---------- links ---------- */

/**
 * FlexLink:
 *   id     stable id
 *   from   { sheet, col }   — an attribute on one entity
 *   to     { sheet, col }   — the attribute it is compared against
 *   mode   "auto" | "text" | "number" | "date" | "digits"
 *   opts   { caseInsensitive, loose, decimals, absolute }
 *
 * Every link between the same pair of sheets belongs to the same relationship
 * and becomes one more field of that relationship's composite key: two rows
 * match only when *all* the drawn fields agree.
 */
function flexDefaultMode(fromSheet, fromCol, toSheet, toCol) {
  const a = flexColumn(fromSheet, fromCol);
  const b = flexColumn(toSheet, toCol);
  if (!a || !b) return "auto";
  if (a.type === b.type) return a.type === "date" ? "date" : a.type === "number" ? "number" : "text";
  // Mixed types: compare as the stricter of the two rather than as raw text.
  if (a.type === "number" || b.type === "number") return "number";
  if (a.type === "date" || b.type === "date") return "date";
  return "text";
}

function flexAddLink(from, to) {
  if (from.sheet === to.sheet) return null;             // a sheet cannot match itself
  if (flexFindLink(from, to)) return null;              // already drawn
  const link = {
    id: flexId("l"),
    from: { ...from },
    to: { ...to },
    mode: flexDefaultMode(from.sheet, from.col, to.sheet, to.col),
    opts: { caseInsensitive: true, loose: false, decimals: 2, absolute: false },
  };
  flexState.links.push(link);
  flexState.result = null;
  return link;
}

function flexFindLink(from, to) {
  return flexState.links.find((l) =>
    (l.from.sheet === from.sheet && l.from.col === from.col && l.to.sheet === to.sheet && l.to.col === to.col) ||
    (l.from.sheet === to.sheet && l.from.col === to.col && l.to.sheet === from.sheet && l.to.col === from.col));
}

function flexLink(linkId) {
  return flexState.links.find((l) => l.id === linkId) || null;
}

function flexRemoveLink(linkId) {
  flexState.links = flexState.links.filter((l) => l.id !== linkId);
  flexState.result = null;
}

/* ---------- relationships ---------- */

// Group the links by the pair of sheets they join. Each group is one recon:
// left/right are fixed by the order the sheets were loaded, so the output
// reads the same way every run.
function flexRelationships() {
  const order = new Map(flexState.sheets.map((s, i) => [s.id, i]));
  const groups = new Map();
  for (const link of flexState.links) {
    const a = link.from, b = link.to;
    if (!flexSheet(a.sheet) || !flexSheet(b.sheet)) continue;
    if (flexSheet(a.sheet).skip || flexSheet(b.sheet).skip) continue;
    const flip = (order.get(a.sheet) ?? 0) > (order.get(b.sheet) ?? 0);
    const left = flip ? b : a, right = flip ? a : b;
    const key = `${left.sheet}|${right.sheet}`;
    if (!groups.has(key)) {
      groups.set(key, { id: key, left: left.sheet, right: right.sheet, fields: [] });
    }
    groups.get(key).fields.push({ link, leftCol: left.col, rightCol: right.col });
  }
  return [...groups.values()];
}

// Every relationship a sheet takes part in.
function flexRelationshipsFor(sheetId) {
  return flexRelationships().filter((rel) => rel.left === sheetId || rel.right === sheetId);
}

function flexRelLabel(rel) {
  const left = flexSheet(rel.left), right = flexSheet(rel.right);
  return `${left ? left.label : "?"} ⇄ ${right ? right.label : "?"}`;
}
