"use strict";

/**
 * utils.js — value parsing/formatting helpers, ported verbatim from the
 * static-reconciliation web app so the reconcile engine behaves identically.
 * Only the pure helpers the engine needs are kept here; grid/DOM helpers from
 * the original file are dropped (the add-in reads rows straight from Excel).
 */

/* ---------- value parsing & formatting ---------- */

// Parse a cell into a signed number. Handles "1,234.56", "(1,234.56)",
// trailing-minus "1,234.56-", DR/CR markers ("1,234.56 DR" = money out) and a
// leading currency code/symbol ("R1 234.56"). Requires the whole cell to be
// numeric, so a date like "15/07/2024" is never misread as the amount 15.
function _amount(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return isFinite(value) ? value : null;
  let text = String(value).trim();
  if (!text) return null;
  let negative = false;
  const marker = text.match(/(dr|cr)\.?$/i);
  if (marker) {
    if (marker[1].toLowerCase() === "dr") negative = true;
    text = text.slice(0, marker.index).trim();
  }
  if (text.startsWith("(") && text.endsWith(")")) { negative = true; text = text.slice(1, -1).trim(); }
  text = text.replace(/^(zar|usd|gbp|eur|[R$£€])\s*/i, "").replace(/[, ]/g, "");
  if (text.endsWith("-")) { negative = true; text = text.slice(0, -1); }
  if (text.startsWith("-")) { negative = true; text = text.slice(1); }
  else if (text.startsWith("+")) text = text.slice(1);
  if (!/^(\d+(\.\d*)?|\.\d+)$/.test(text)) return null;
  return negative ? -parseFloat(text) : parseFloat(text);
}

function _text(value) {
  return value == null ? "" : String(value).trim();
}

function _is_date_like(value) {
  if (value instanceof Date) return true;
  if (!value) return false;
  const text = String(value).trim();
  if (/^\d{1,4}[-/.\s]\d{1,2}[-/.\s]\d{1,4}$/.test(text)) return true;
  if (/^\d{6,8}$/.test(text)) return true;
  if (/^\d{1,2}\s+[A-Za-z]{3,9}(\s+\d{2,4})?$/.test(text)) return true;
  return false;
}

// Normalise a date value to "YYYY-MM-DD" for display. Handles JS Date and
// Excel serial numbers; otherwise returns the trimmed text.
function _display_date(value) {
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  // Excel numeric dates (often > 20000 for 1900 dates)
  if (typeof value === 'number' && value > 10000) {
     const dt = new Date((value - (25569)) * 86400 * 1000);
     const utc = new Date(dt.getTime() + dt.getTimezoneOffset() * 60000);
     return `${utc.getFullYear()}-${String(utc.getMonth()+1).padStart(2,'0')}-${String(utc.getDate()).padStart(2,'0')}`;
  }
  return _text(value);
}

const _MONTH_NUM = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

// Key(s) for matching two dates regardless of formatting. Always includes the
// raw digit string plus a canonical "YYYYMMDD" key for every plausible reading
// — D/M/Y, M/D/Y, Y-M-D, named months.
function _date_keys(value) {
  const keys = new Set();
  if (value == null) return keys;
  const text = String(_display_date(value)).trim();
  if (!text) return keys;
  const digits = text.replace(/\D/g, "");
  if (digits) keys.add(digits);

  const pad = (n) => String(n).padStart(2, "0");
  const fullYear = (y) => (y >= 100 ? y : (y >= 50 ? 1900 + y : 2000 + y));
  const add = (y, m, d) => {
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) keys.add(`${fullYear(y)}${pad(m)}${pad(d)}`);
  };

  let m = text.match(/^(\d{1,2})[\s.\-]+([A-Za-z]{3,9})\.?[\s.\-,]+(\d{2,4})$/);   // "15 Jul 2024"
  if (m) {
    const month = _MONTH_NUM[m[2].slice(0, 3).toLowerCase()];
    if (month) add(+m[3], month, +m[1]);
  } else if ((m = text.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{2,4})$/))) { // "Jul 15, 2024"
    const month = _MONTH_NUM[m[1].slice(0, 3).toLowerCase()];
    if (month) add(+m[3], month, +m[2]);
  } else {
    const parts = text.split(/[-/.\s]+/).filter(Boolean);
    if (parts.length === 3 && parts.every((p) => /^\d+$/.test(p))) {
      const [a, b, c] = parts.map(Number);
      if (parts[0].length === 4) add(a, b, c);          // Y-M-D
      else { add(c, b, a); add(c, a, b); }              // D/M/Y and M/D/Y
    } else if (/^\d{8}$/.test(text)) {
      add(+text.slice(0, 4), +text.slice(4, 6), +text.slice(6));  // YYYYMMDD
      add(+text.slice(4), +text.slice(2, 4), +text.slice(0, 2));  // DDMMYYYY
      add(+text.slice(4), +text.slice(0, 2), +text.slice(2, 4));  // MMDDYYYY
    }
  }
  return keys;
}

function _description_tail(description) {
  const parts = description.trim().split(/\s+/);
  if (parts.length <= 1) return "";
  return parts.slice(1).join(" ").trim();
}

function _is_balance_row(description) {
  return /^(opening\s+|closing\s+)?balance(\s+(b\/?f(wd)?|c\/?f(wd)?|brought\s+forward|carried\s+forward))?$/i.test(description.trim());
}

// Year-month key(s) ("YYYYMM") a date could belong to.
function monthKeys(value) {
  const text = String(_display_date(value)).trim();
  if (!text) return [];
  const pad = (n) => String(n).padStart(2, "0");
  const named = text.match(/([A-Za-z]{3,9})\.?\s+(\d{4})/);
  if (named) {
    const month = _MONTH_NUM[named[1].slice(0, 3).toLowerCase()];
    return month ? [`${named[2]}${pad(month)}`] : [];
  }
  const keys = new Set();
  const fullYear = (y) => String(+y >= 100 ? +y : (+y >= 50 ? 1900 + +y : 2000 + +y));
  const parts = text.split(/[-/.\s]+/);
  if (parts.length >= 3 && parts.slice(0, 3).every((p) => /^\d+$/.test(p))) {
    const [first, middle, last] = parts;
    if (first.length === 4) {
      if (+middle >= 1 && +middle <= 12) keys.add(first + pad(+middle));
    } else {
      for (const candidate of new Set([middle, first])) {
        if (+candidate >= 1 && +candidate <= 12) keys.add(fullYear(last) + pad(+candidate));
      }
    }
    return [...keys];
  }
  const digits = text.replace(/\D/g, "");
  if (digits.length === 8) {
    if (+digits.slice(4, 6) >= 1 && +digits.slice(4, 6) <= 12) keys.add(digits.slice(0, 6));
    if (+digits.slice(2, 4) >= 1 && +digits.slice(2, 4) <= 12) keys.add(digits.slice(4) + digits.slice(2, 4));
  } else if (digits.length === 6) {
    if (+digits.slice(2, 4) >= 1 && +digits.slice(2, 4) <= 12) keys.add(fullYear(digits.slice(4)) + digits.slice(2, 4));
    if (+digits.slice(0, 2) >= 1 && +digits.slice(0, 2) <= 12) keys.add(fullYear(digits.slice(4)) + digits.slice(0, 2));
  }
  return [...keys];
}

// Normalise a stored monthFilter into a list of canonical 6-digit "YYYYMM" keys.
function _monthFilterKeys(monthFilter) {
  if (!monthFilter) return [];
  const list = Array.isArray(monthFilter) ? monthFilter : [monthFilter];
  return list.map((m) => String(m).replace(/\D/g, "")).filter((d) => d.length === 6);
}

function formatAmount(value) {
  if (value === null || value === undefined || value === "") return "";
  return Number(value).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// 0-based column index -> spreadsheet letter (0 -> "A", 26 -> "AA").
function colLetter(index) {
  let n = index + 1, letters = "";
  while (n > 0) {
    n -= 1;
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26);
  }
  return letters;
}
