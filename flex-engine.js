"use strict";

/**
 * flex-engine.js — the matching engine for Modular Recon.
 *
 * The classic engine (engine.js) knows what a cashbook is. This one knows
 * nothing: it is handed a relationship — two sheets plus the list of attribute
 * pairs the user drew between them — and answers one question per row, "does
 * this row have a partner on the other sheet?".
 *
 * How it matches
 *   Each drawn pair contributes one part of a composite key. A row's key is the
 *   parts joined together, so two rows match only when *every* drawn field
 *   agrees. Each part is normalised first (dates to YYYYMMDD, numbers to a fixed
 *   number of decimals, text lower-cased), which is what lets "15/07/2024" meet
 *   "2024-07-15" and "ACME  Ltd" meet "acme ltd".
 *
 * Why keys and not fuzzy scoring
 *   Matching is duplicate-aware and one-to-one: rows are bucketed by key and a
 *   bucket with 30 rows on the left and 29 on the right yields 29 pairs and one
 *   unmatched row on the left — which is the whole point of the exercise. That
 *   counting only works if a value maps to exactly one key, so every compare
 *   mode here is a key transform. Nothing is matched "approximately".
 *
 * Row statuses: "matched" | "ambiguous" | "unmatched" | "novalue".
 * "ambiguous" is a real match whose key occurs more than once on a side — the
 * count is right but which row paired with which is arbitrary, so it is
 * reported separately instead of being passed off as certain.
 *
 * Ported verbatim from the static-reconciliation web app.
 */

const FLEX_MODES = [
  ["auto", "Auto"],
  ["text", "Text"],
  ["number", "Number"],
  ["date", "Date"],
  ["digits", "Digits only"],
];

const FLEX_STATUS_LABEL = {
  matched: "Matched",
  ambiguous: "Matched (repeated value)",
  unmatched: "Not found",
  novalue: "No value",
};

/* ---------- value normalisation ---------- */

// One canonical YYYYMMDD key for a date cell. `order` is the column's settled
// reading (see flexDateOrder) so a numeric date is never counted twice.
function flexCanonicalDate(value, order) {
  const text = String(_display_date(value)).trim();
  if (!text) return null;
  const pad = (n) => String(n).padStart(2, "0");
  const fullYear = (y) => (y >= 100 ? y : (y >= 50 ? 1900 + y : 2000 + y));
  const make = (y, m, d) => (m >= 1 && m <= 12 && d >= 1 && d <= 31) ? `${fullYear(y)}${pad(m)}${pad(d)}` : null;

  let m = text.match(/^(\d{1,2})[\s.\-]+([A-Za-z]{3,9})\.?[\s.\-,]+(\d{2,4})$/);   // "15 Jul 2024"
  if (m) {
    const month = _MONTH_NUM[m[2].slice(0, 3).toLowerCase()];
    return month ? make(+m[3], month, +m[1]) : null;
  }
  m = text.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{2,4})$/);                // "Jul 15, 2024"
  if (m) {
    const month = _MONTH_NUM[m[1].slice(0, 3).toLowerCase()];
    return month ? make(+m[3], month, +m[2]) : null;
  }

  const parts = text.split(/[-/.\s]+/).filter(Boolean);
  if (parts.length >= 3 && parts.slice(0, 3).every((p) => /^\d+$/.test(p))) {
    const [a, b, c] = parts.map(Number);
    if (parts[0].length === 4) return make(a, b, c);        // Y-M-D
    if (order === "mdy") return make(c, a, b);
    return make(c, b, a);                                   // D/M/Y (the default)
  }

  const digits = text.replace(/\D/g, "");
  if (digits.length === 8) {
    if (order === "ymd") return make(+digits.slice(0, 4), +digits.slice(4, 6), +digits.slice(6));
    if (order === "mdy") return make(+digits.slice(4), +digits.slice(0, 2), +digits.slice(2, 4));
    return make(+digits.slice(4), +digits.slice(2, 4), +digits.slice(0, 2));
  }
  return null;
}

// Normalise one cell into a key part, or null when the cell can't take part
// (blank, or not readable as the mode expects).
function flexKeyPart(value, link, column) {
  const mode = link.mode === "auto" ? (column ? column.type : "text") : link.mode;
  const opts = link.opts || {};

  if (mode === "number") {
    let n = _amount(value);
    if (n === null) return null;
    if (opts.absolute) n = Math.abs(n);
    const dec = Number.isFinite(opts.decimals) ? opts.decimals : 2;
    const factor = Math.pow(10, dec);
    let rounded = Math.round(n * factor) / factor;
    if (rounded === 0) rounded = 0;           // fold -0 into 0
    return rounded.toFixed(dec);
  }

  if (mode === "date") {
    return flexCanonicalDate(value, column ? column.dateOrder : "dmy");
  }

  if (mode === "digits") {
    const digits = String(_display_date(value)).replace(/\D/g, "");
    return digits || null;
  }

  // Text (and the auto fallback).
  let text = String(_display_date(value)).trim();
  if (!text) return null;
  if (opts.caseInsensitive !== false) text = text.toLowerCase();
  text = opts.loose
    ? text.replace(/[^a-z0-9]/gi, "")        // ignore spacing and punctuation entirely
    : text.replace(/\s+/g, " ");
  return text || null;
}

// A separator that cannot appear in a normalised part, so ["ab","c"] can
// never collide with ["a","bc"].
const FLEX_KEY_SEP = String.fromCharCode(1);

// Build the composite key for one row. Any part that can't be read makes the
// whole row unusable for this relationship ("No value") — matching a row on
// half its key would invent matches that aren't there.
function flexRowKey(cells, fields, side) {
  const parts = [];
  const shown = [];
  for (const field of fields) {
    const col = side === "left" ? field.leftCol : field.rightCol;
    const sheetId = side === "left" ? field.leftSheet : field.rightSheet;
    const value = cells[col] ?? null;
    const part = flexKeyPart(value, field.link, flexColumn(sheetId, col));
    shown.push(_display_date(value));
    if (part === null) return { key: null, shown };
    parts.push(part);
  }
  return { key: parts.join(FLEX_KEY_SEP), shown };
}

/* ---------- relationship matching ---------- */

function flexEntries(sheet, fields, side) {
  return flexDataRows(sheet).map(({ row, cells }) => {
    const { key, shown } = flexRowKey(cells, fields, side);
    return { row, key, values: shown, status: null, partner: null };
  });
}

function flexMatchRelationship(rel) {
  const leftSheet = flexSheet(rel.left);
  const rightSheet = flexSheet(rel.right);
  // Carry the sheet ids on each field so flexRowKey can look up column metadata
  // (a column's date order and detected type live on the sheet, not the link).
  const fields = rel.fields.map((f) => ({ ...f, leftSheet: rel.left, rightSheet: rel.right }));

  const left = flexEntries(leftSheet, fields, "left");
  const right = flexEntries(rightSheet, fields, "right");

  const leftCounts = new Map();
  for (const e of left) if (e.key !== null) leftCounts.set(e.key, (leftCounts.get(e.key) || 0) + 1);

  const buckets = new Map();
  for (const e of right) {
    if (e.key === null) continue;
    if (!buckets.has(e.key)) buckets.set(e.key, []);
    buckets.get(e.key).push(e);
  }

  // Walk the left rows in sheet order, taking the next unused row from the
  // matching bucket. Whatever is left over on either side is genuinely
  // unmatched — the count difference the user is looking for.
  const cursor = new Map();
  for (const e of left) {
    if (e.key === null) { e.status = "novalue"; continue; }
    const bucket = buckets.get(e.key);
    const i = cursor.get(e.key) || 0;
    if (!bucket || i >= bucket.length) { e.status = "unmatched"; continue; }
    const partner = bucket[i];
    cursor.set(e.key, i + 1);
    const repeated = bucket.length > 1 || (leftCounts.get(e.key) || 0) > 1;
    e.status = partner.status = repeated ? "ambiguous" : "matched";
    e.partner = partner.row;
    partner.partner = e.row;
  }
  for (const e of right) if (!e.status) e.status = e.key === null ? "novalue" : "unmatched";

  // Keys present more than once on either side: the pairing inside these groups
  // is arbitrary, so they are surfaced rather than buried.
  const duplicates = [];
  const seen = new Set();
  for (const e of [...left, ...right]) {
    if (e.key === null || seen.has(e.key)) continue;
    seen.add(e.key);
    const l = leftCounts.get(e.key) || 0;
    const r = (buckets.get(e.key) || []).length;
    if (l > 1 || r > 1) duplicates.push({ key: e.key, label: e.values.join(" · "), left: l, right: r });
  }
  duplicates.sort((a, b) => Math.abs(b.left - b.right) - Math.abs(a.left - a.right));

  const tally = (rows) => {
    const out = { total: rows.length, matched: 0, ambiguous: 0, unmatched: 0, novalue: 0 };
    for (const e of rows) out[e.status]++;
    return out;
  };

  return {
    id: rel.id,
    left: rel.left,
    right: rel.right,
    leftLabel: leftSheet.label,
    rightLabel: rightSheet.label,
    fields: rel.fields,
    leftRows: left,
    rightRows: right,
    duplicates,
    summary: {
      pairs: left.filter((e) => e.partner !== null).length,
      left: tally(left),
      right: tally(right),
    },
  };
}

/* ---------- top level ---------- */

// A row can sit in more than one relationship (a cashbook matched against both
// a statement and a ledger, say). The sheet-level status is the honest
// combination: any miss makes the row unmatched, any caveat makes it ambiguous.
function flexCombineStatus(statuses) {
  if (!statuses.length) return null;
  if (statuses.includes("unmatched")) return "unmatched";
  if (statuses.every((s) => s === "novalue")) return "novalue";
  if (statuses.includes("ambiguous") || statuses.includes("novalue")) return "ambiguous";
  return "matched";
}

function flexRunRecon() {
  const relationships = flexRelationships().map(flexMatchRelationship);

  // Per sheet: the status of each row, plus which columns took part (the Data
  // view paints those cells when highlighting is set to "cell").
  const bySheet = {};
  const ensure = (sheetId) => (bySheet[sheetId] || (bySheet[sheetId] = { rows: new Map(), cols: new Set() }));

  for (const rel of relationships) {
    for (const side of ["left", "right"]) {
      const sheetId = rel[side];
      const store = ensure(sheetId);
      for (const field of rel.fields) store.cols.add(side === "left" ? field.leftCol : field.rightCol);
      for (const e of rel[side === "left" ? "leftRows" : "rightRows"]) {
        if (!store.rows.has(e.row)) store.rows.set(e.row, { status: null, byRel: {} });
        store.rows.get(e.row).byRel[rel.id] = {
          status: e.status,
          partner: e.partner,
          partnerSheet: side === "left" ? rel.right : rel.left,
          values: e.values,
        };
      }
    }
  }
  for (const store of Object.values(bySheet)) {
    for (const entry of store.rows.values()) {
      entry.status = flexCombineStatus(Object.values(entry.byRel).map((r) => r.status));
    }
  }

  return { relationships, bySheet, generatedAt: new Date() };
}
