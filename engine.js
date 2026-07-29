"use strict";

/**
 * engine.js — the reconciliation engine, ported verbatim from the
 * static-reconciliation web app. Same matching rule, same one-to-one
 * consumption. The only environment dependency is that _desc_similar reads two
 * checkboxes by id ("global-desc-ignore" / "global-desc-strict") — the task
 * pane provides them, so this file is otherwise untouched.
 *
 * Matching rule: same absolute amount (to the cent) AND a shared date key.
 * Among amount+date candidates, description similarity decides
 * "Matched" vs "Check description". No candidates => "Not found".
 */

// Header-keyword hints used to auto-detect which column plays which role.
// (Originally lived in config.js.)
const KEYWORDS = {
  date: ["date", "datum", "value date", "posting"],
  description: ["description", "details", "narrative", "particulars", "reference", "memo", "transaction", "payee", "beskrywing"],
  amount: ["amount", "value", "bedrag"],
  debit: ["debit", "withdrawal", "money out", "payments", "paid out", "dr"],
  credit: ["credit", "deposit", "money in", "receipts", "paid in", "cr"],
};

/* ---------- header & column detection ---------- */

// For one row, find the first column index whose text matches each role's keywords.
function _header_scores(row) {
  const hits = {};
  for (let index = 0; index < row.length; index++) {
    const text = _text(row[index]).toLowerCase();
    if (!text) continue;
    for (const [role, words] of Object.entries(KEYWORDS)) {
      if (role in hits) continue;
      for (const word of words) {
        if ((word.length <= 2 && word === text) || (word.length > 2 && text.includes(word))) {
          hits[role] = index;
          break;
        }
      }
    }
  }
  return hits;
}

// Fallback when no header row is recognised: classify columns by content.
function _detect_by_content(rows, width) {
  const dateCounts = Array(width).fill(0);
  const numberCounts = Array(width).fill(0);
  const textLengths = Array(width).fill(0);

  const sample = rows.slice(0, 200);
  for (const row of sample) {
    for (let i = 0; i < width; i++) {
      const val = row[i] !== undefined ? row[i] : null;
      if (val == null) continue;
      if (_is_date_like(val)) dateCounts[i]++;
      else if (_amount(val) !== null) numberCounts[i]++;
      else textLengths[i] += _text(val).length;
    }
  }

  function best(counts, exclude) {
    let maxCount = -1;
    let bestIdx = null;
    for (let i = 0; i < width; i++) {
      if (!exclude.has(i) && counts[i] > maxCount && counts[i] > 0) {
        maxCount = counts[i];
        bestIdx = i;
      }
    }
    return bestIdx;
  }

  const dateCol = best(dateCounts, new Set());
  const descCol = best(textLengths, new Set([dateCol]));
  const amountCol = best(numberCounts, new Set([dateCol, descCol]));

  return { date: dateCol, description: descCol, amount: amountCol, debit: null, credit: null };
}

async function runLocalAnalyze(rows, sheets, sheetName) {
  const width = rows.reduce((max, r) => Math.max(max, r.length), 0);
  let headerRow = null;
  let mapping = {};
  let bestScore = 0;

  for (let i = 0; i < Math.min(20, rows.length); i++) {
    const row = rows[i];
    const hits = _header_scores(row);
    const score = Object.keys(hits).length;
    if (hits.date !== undefined && score > bestScore) {
      bestScore = score;
      headerRow = i + 1; // 1-based
      mapping = hits;
    }
  }

  if (headerRow === null) {
    mapping = _detect_by_content(rows, width);
  } else {
    const dataRows = rows.slice(headerRow);
    ["date", "description", "amount", "debit", "credit"].forEach(role => {
      if (mapping[role] === undefined) mapping[role] = null;
    });

    if (mapping.description === null || (mapping.amount === null && mapping.debit === null)) {
      const content = _detect_by_content(dataRows, width);
      const used = new Set(Object.values(mapping).filter(v => v !== null));
      for (const role of ["description", "amount"]) {
        if (mapping[role] === null && content[role] !== null && !used.has(content[role])) {
          mapping[role] = content[role];
        }
      }
    }
  }

  const mode = (mapping.debit !== null && mapping.credit !== null) ? "debit_credit" : "single";
  const headers = headerRow ? rows[headerRow - 1] : [];
  const start = headerRow ? headerRow : 0;

  const columns = [];
  for (let i = 0; i < width; i++) {
    columns.push({
      index: i,
      letter: colLetter(i),
      header: _text(headers[i] !== undefined ? headers[i] : ""),
      samples: rows.slice(start, start + 4).map(r => _display_date(r[i] !== undefined ? r[i] : null))
    });
  }

  return {
    sheets,
    sheet: sheetName,
    headerRow: headerRow || 0,
    columns,
    mapping: { mode, ...mapping },
    rowCount: Math.max(0, rows.length - start)
  };
}

/* ---------- extraction ---------- */

// Turn raw rows + mapping into normalised entries { row, date_raw, description, amount }.
function _extract(rows, mapping) {
  const start = mapping.headerRow ? mapping.headerRow : 0;
  const dateCol = mapping.date;
  const descCol = mapping.description;
  const mode = mapping.mode || "single";

  const entries = [];
  for (let i = start; i < rows.length; i++) {
    const row = rows[i];
    const cell = (col) => (col !== null && col !== undefined && col < row.length) ? row[col] : null;

    let amount = null;
    if (mode === "debit_credit") {
      const debit = _amount(cell(mapping.debit));
      const credit = _amount(cell(mapping.credit));
      if (debit !== null || credit !== null) {
        amount = (credit || 0.0) - Math.abs(debit || 0.0);
      }
    } else {
      amount = _amount(cell(mapping.amount));
    }

    entries.push({
      row: i + 1, // 1-based index to match python
      date_raw: cell(dateCol),
      description: _text(cell(descCol)),
      amount: amount
    });
  }
  return entries;
}

// Extract + enrich with date/month keys, ready to be matched against.
function _prepare_side(rows, mapping) {
  const entries = _extract(rows, mapping);
  return entries.map(e => ({
    row: e.row,
    date_keys: _date_keys(e.date_raw),
    month_keys: monthKeys(e.date_raw),
    date: _display_date(e.date_raw),
    description: e.description,
    amount: e.amount
  }));
}

/* ---------- matching ---------- */

// Rows on `side` with the same absolute amount (to the cent) AND a shared date key.
function _find_candidates(amount, date_keys, side) {
  const cents = Math.round(Math.abs(amount) * 100);
  return side.filter(s =>
    s.amount !== null &&
    Math.round(Math.abs(s.amount) * 100) === cents &&
    [...date_keys].some(dk => s.date_keys.has(dk))
  );
}

// Description similarity, honouring the global "Ignore descriptions" and
// "Strict Descriptions" toggles.
function _desc_similar(sourceDescRaw, targetDescRaw, tail) {
  const ignoreEl = document.getElementById("global-desc-ignore");
  if (ignoreEl && ignoreEl.checked) return true;

  const targetDesc = targetDescRaw.trim().toLowerCase();
  const sourceDesc = sourceDescRaw.trim().toLowerCase();
  if (targetDesc === sourceDesc) return true;

  const strictEl = document.getElementById("global-desc-strict");
  if (strictEl && strictEl.checked) return false;

  const digitWords = sourceDesc.split(/\s+/).filter(w => /\d/.test(w) && !/^\d{1,2}$/.test(w));
  if (digitWords.length > 0) {
    return digitWords.some(w => targetDesc.includes(w));
  }

  if (!tail) return false;
  if (targetDesc.includes(tail)) return true;
  return tail.split(/\s+/).filter(w => w.length > 2).some(w => targetDesc.includes(w));
}

// Two-pass one-to-one matcher used by reconcile.
function _assign_matches(items, side, apply, monthFilter = null) {
  const used = new Set();
  const pending = [];

  for (const it of items) {
    if (monthFilter && monthFilter.length &&
        !monthKeys(it.e.date_raw).some((k) => monthFilter.includes(k))) continue;
    const candidates = _find_candidates(it.e.amount, it.keys, side);
    if (candidates.length === 0) { apply(it.entry, "Not found", []); continue; }
    const pick = candidates.find(s => !used.has(s.row) && _desc_similar(it.e.description, s.description, it.tail));
    if (pick) { used.add(pick.row); apply(it.entry, "Matched", [pick.row]); }
    else pending.push({ it, candidates });
  }

  for (const { it, candidates } of pending) {
    const pick = candidates.find(s => !used.has(s.row));
    if (pick) { used.add(pick.row); apply(it.entry, "Check description", [pick.row]); }
    else apply(it.entry, "Not found", []);
  }
  return used;
}

// Build the per-side (statement/ledger) row list flagged with matched state.
function _side_rows(side, matched_rows) {
  return side
    .filter(s => s.amount !== null || s.description)
    .map(s => ({
      row: s.row,
      date: s.date,
      description: s.description,
      amount: s.amount,
      matched: matched_rows.has(s.row),
      statusLabel: matched_rows.has(s.row) ? "✓ Matched on cashbook" : "✗ Not on cashbook"
    }));
}

/* ---------- top-level reconcile ---------- */

function runLocalReconcile(cashbookRows, statementRows, cashbookMap, statementMap, ledgerRows = null, ledgerMap = null) {
  const cashbook = _extract(cashbookRows, cashbookMap);
  const hasStatement = statementRows !== null && statementMap !== null;
  const statement = hasStatement ? _prepare_side(statementRows, statementMap) : [];
  const hasLedger = ledgerRows !== null && ledgerMap !== null;
  let ledger = hasLedger ? _prepare_side(ledgerRows, ledgerMap) : [];

  let monthFilter = null;
  if (hasLedger) {
    monthFilter = _monthFilterKeys(ledgerMap.monthFilter);

    ledger = ledger.filter(s => s.month_keys.length > 0 && !_is_balance_row(s.description));
    if (monthFilter.length) {
      ledger = ledger.filter(s => s.month_keys.some((k) => monthFilter.includes(k)));
    }
  }

  const cbMonthFilter = _monthFilterKeys(cashbookMap.monthFilter);

  const results = [];
  const matchable = [];

  for (const e of cashbook) {
    if (e.date_raw == null && e.amount == null && !e.description) continue;
    if (cbMonthFilter.length && !monthKeys(e.date_raw).some((k) => cbMonthFilter.includes(k))) continue;

    const entry = {
      row: e.row,
      date: _display_date(e.date_raw),
      description: e.description,
      amount: e.amount,
      status: "",
      matchedStatementRows: [],
      ledgerStatus: "",
      matchedLedgerRows: []
    };
    results.push(entry);

    if (e.date_raw == null) continue;
    if (e.amount === null) {
      if (hasStatement) entry.status = "No amount";
      if (hasLedger) entry.ledgerStatus = "No amount";
      continue;
    }
    matchable.push({
      e, entry,
      keys: _date_keys(e.date_raw),
      tail: _description_tail(e.description).toLowerCase()
    });
  }

  const matchedStatement = hasStatement
    ? _assign_matches(matchable, statement,
        (entry, status, rows) => { entry.status = status; entry.matchedStatementRows = rows; })
    : new Set();
  const matchedLedger = hasLedger
    ? _assign_matches(matchable, ledger,
        (entry, status, rows) => { entry.ledgerStatus = status; entry.matchedLedgerRows = rows; },
        monthFilter)
    : new Set();

  function count(statuses) {
    const counts = { "Matched": 0, "Check description": 0, "Not found": 0, "No amount": 0 };
    for (const st of statuses) if (counts[st] !== undefined) counts[st]++;
    return counts;
  }

  function label(status, target) {
    if (status === "Matched") return `✓ Matched on ${target}`;
    if (status === "Check description") return `⚠ Matched on ${target} - check description`;
    if (status === "Not found") return `✗ Not on ${target}`;
    if (status === "No amount") return `⚠ No amount`;
    return "";
  }

  for (const r of results) {
    r.statusLabel = label(r.status, "bank statement");
    r.ledgerStatusLabel = label(r.ledgerStatus, "general ledger");
  }

  const stRowsOut = _side_rows(statement, matchedStatement);
  const ldRowsOut = _side_rows(ledger, matchedLedger);

  return {
    rows: results,
    statement: stRowsOut,
    ledger: ldRowsOut,
    hasStatement,
    hasLedger,
    summary: count(results.map(r => r.status)),
    ledgerSummary: hasLedger ? count(results.map(r => r.ledgerStatus)) : null,
    unmatchedStatementCount: stRowsOut.filter(s => !s.matched).length,
    unmatchedLedgerCount: ldRowsOut.filter(s => !s.matched).length,
    sources: {
      cashbook: { rows: cashbookRows, headerRow: cashbookMap.headerRow || 0 },
      statement: hasStatement ? { rows: statementRows, headerRow: statementMap.headerRow || 0 } : null,
      ledger: hasLedger ? { rows: ledgerRows, headerRow: ledgerMap.headerRow || 0 } : null
    }
  };
}
