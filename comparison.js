"use strict";

/**
 * comparison.js — the pure "what belongs together" builders, ported verbatim
 * from the web app's results.js + unmatched.js with every DOM/rendering
 * function removed. Produces plain data; sheets.js turns it into worksheets.
 *
 *  - buildComparisonRows(result)         : cashbook rows aligned with their
 *                                          matched statement/ledger row.
 *  - buildUnmatchedComparison(result)    : unmatched listings + correlation groups.
 *  - uncorrelatedBySheet(result, data)   : leftovers, split per sheet.
 *  - unmatchedComparisonSheetRows(result): the whole Unmatched sheet as an aoa.
 */

/* ---------- comparison (results.js) ---------- */

// Align each cashbook row with its matched statement/ledger row, then append
// the side rows nothing matched.
function buildComparisonRows(result) {
  const stByRow = new Map(result.statement.map(s => [s.row, s]));
  const ldByRow = new Map(result.ledger.map(l => [l.row, l]));
  const shownSt = new Set();
  const shownLd = new Set();
  const out = [];
  for (const r of result.rows) {
    const st = r.matchedStatementRows.map(n => stByRow.get(n)).find(Boolean) || null;
    const ld = r.matchedLedgerRows.map(n => ldByRow.get(n)).find(Boolean) || null;
    if (st) shownSt.add(st.row);
    if (ld) shownLd.add(ld.row);
    out.push({ cb: r, st, ld, stStatus: r.status, ldStatus: r.ledgerStatus });
  }
  // Side rows nothing on the cashbook claimed: shown on their own, unmatched.
  for (const s of result.statement) {
    if (!s.matched && !shownSt.has(s.row)) {
      out.push({ cb: null, st: s, ld: null, stStatus: "Not found", ldStatus: "" });
    }
  }
  for (const l of result.ledger) {
    if (!l.matched && !shownLd.has(l.row)) {
      out.push({ cb: null, st: null, ld: l, stStatus: "", ldStatus: "Not found" });
    }
  }
  return out;
}

/* ---------- correlation tuning (unmatched.js) ---------- */

const SUM_MAX_MEMBERS = 4;
const SUM_POOL_LIMIT = 14;
const SUM_STEP_LIMIT = 20000;

const _cents = (amount) => Math.round(Math.abs(amount) * 100);

function _dayNumbers(value) {
  const out = [];
  for (const key of _date_keys(value)) {
    if (!/^\d{8}$/.test(key)) continue;
    const year = +key.slice(0, 4), month = +key.slice(4, 6), day = +key.slice(6);
    if (year < 1900 || year > 2200 || month < 1 || month > 12 || day < 1 || day > 31) continue;
    out.push(Math.round(Date.UTC(year, month - 1, day) / 86400000));
  }
  return out;
}

function _dayGap(a, b) {
  const left = _dayNumbers(a), right = _dayNumbers(b);
  if (!left.length || !right.length) return null;
  let best = Infinity;
  for (const x of left) for (const y of right) best = Math.min(best, Math.abs(x - y));
  return best;
}

function _dateScore(gap) {
  if (gap === null) return 0.25;
  return Math.max(0, 1 - gap / 31);
}

function _descTokens(text) {
  return [...new Set(
    String(text || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").split(" ")
      .filter((word) => word && (word.length > 2 || /\d/.test(word)))
  )];
}

function _descScore(a, b) {
  const left = _descTokens(a), right = _descTokens(b);
  if (!left.length || !right.length) return 0;
  const set = new Set(right);
  const shared = left.filter((word) => set.has(word));
  if (!shared.length) return 0;
  const base = shared.length / Math.min(left.length, right.length);
  return shared.some((word) => /\d/.test(word)) ? Math.max(base, 0.9) : base;
}

function _gapText(gap) {
  if (gap === null) return "date unreadable";
  if (gap === 0) return "same date";
  return `${gap} day${gap === 1 ? "" : "s"} apart`;
}

function _descText(score) {
  if (score >= 0.9) return "description matches";
  if (score >= 0.4) return "description partly matches";
  return "description differs";
}

function _unmatchedSides(result) {
  const useStatement = result.hasStatement !== false;
  const sideRows = useStatement ? result.statement : result.ledger;
  const statusOf = (r) => (useStatement ? r.status : r.ledgerStatus);
  const pick = (r) => ({ row: r.row, date: r.date, description: r.description, amount: r.amount });

  return {
    leftLabel: useStatement ? "Bank Statement" : "General Ledger",
    leftPrefix: useStatement ? "BS" : "GL",
    rightLabel: "Cashbook",
    rightPrefix: "CB",
    left: (sideRows || []).filter((s) => !s.matched).map(pick),
    right: result.rows.filter((r) => statusOf(r) === "Not found").map(pick)
  };
}

function _countByAmount(rows) {
  const counts = new Map();
  for (const r of rows) {
    if (r.amount === null) continue;
    const key = _cents(r.amount);
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  return counts;
}

function _pairOnAmount(left, right, usedLeft, usedRight) {
  const leftCounts = _countByAmount(left);
  const rightCounts = _countByAmount(right);

  const pairs = [];
  for (const l of left) {
    if (l.amount === null) continue;
    for (const r of right) {
      if (r.amount === null || _cents(l.amount) !== _cents(r.amount)) continue;
      const gap = _dayGap(l.date, r.date);
      const desc = _descScore(l.description, r.description);
      pairs.push({ l, r, gap, desc, score: _dateScore(gap) * 10 + desc });
    }
  }
  pairs.sort((a, b) => b.score - a.score || a.l.row - b.l.row || a.r.row - b.r.row);

  const groups = [];
  for (const p of pairs) {
    if (usedLeft.has(p.l.row) || usedRight.has(p.r.row)) continue;
    usedLeft.add(p.l.row);
    usedRight.add(p.r.row);

    const key = _cents(p.l.amount);
    const duplicated = (leftCounts.get(key) || 0) > 1 || (rightCounts.get(key) || 0) > 1;
    const reasons = [`same amount, ${_gapText(p.gap)}, ${_descText(p.desc)}`];
    if (duplicated) reasons.push("⚠ amount appears more than once — check which row is which");

    groups.push({ kind: "one-to-one", left: [p.l], right: [p.r], reason: reasons.join("; "), score: p.score });
  }
  return groups;
}

function _findSumGroup(target, pool) {
  const goal = _cents(target.amount);
  if (!goal) return null;

  const candidates = pool
    .filter((p) => p.amount !== null && _cents(p.amount) > 0 && _cents(p.amount) < goal)
    .map((p) => ({ p, gap: _dayGap(target.date, p.date) }))
    .sort((a, b) => (a.gap === null ? 99999 : a.gap) - (b.gap === null ? 99999 : b.gap))
    .slice(0, SUM_POOL_LIMIT);
  if (candidates.length < 2) return null;

  let best = null;
  let steps = 0;
  const chosen = [];

  const search = (start, remaining) => {
    if (steps++ > SUM_STEP_LIMIT) return;
    if (remaining === 0 && chosen.length >= 2) {
      const spread = Math.max(...chosen.map((c) => (c.gap === null ? 999 : c.gap)));
      if (!best || chosen.length < best.members.length ||
          (chosen.length === best.members.length && spread < best.spread)) {
        best = { members: chosen.map((c) => c.p), spread };
      }
      return;
    }
    if (chosen.length >= SUM_MAX_MEMBERS) return;
    for (let i = start; i < candidates.length; i++) {
      const c = candidates[i];
      const cents = _cents(c.p.amount);
      if (cents > remaining) continue;
      if (chosen.length && Math.sign(chosen[0].p.amount) !== Math.sign(c.p.amount)) continue;
      chosen.push(c);
      search(i + 1, remaining - cents);
      chosen.pop();
    }
  };
  search(0, goal);
  return best;
}

function _collectSumGroups(oneRows, manyRows, usedOne, usedMany, oneIsLeft, manyLabel) {
  const groups = [];
  const targets = oneRows
    .filter((r) => !usedOne.has(r.row) && r.amount !== null)
    .sort((a, b) => _cents(b.amount) - _cents(a.amount));

  for (const target of targets) {
    const pool = manyRows.filter((r) => !usedMany.has(r.row));
    if (pool.length < 2) break;
    const found = _findSumGroup(target, pool);
    if (!found) continue;

    usedOne.add(target.row);
    for (const member of found.members) usedMany.add(member.row);
    found.members.sort((a, b) => a.row - b.row);

    const desc = Math.max(...found.members.map((m) => _descScore(target.description, m.description)));
    const gap = found.spread >= 999 ? null : found.spread;
    groups.push({
      kind: "one-to-many",
      left: oneIsLeft ? [target] : found.members,
      right: oneIsLeft ? found.members : [target],
      reason: `${found.members.length} ${manyLabel} rows add up to ${formatAmount(Math.abs(target.amount))}` +
              ` (${gap === null ? "dates unreadable" : `within ${gap} day${gap === 1 ? "" : "s"}`})`,
      score: _dateScore(gap) * 10 + desc - 5
    });
  }
  return groups;
}

function buildUnmatchedComparison(result) {
  const sides = _unmatchedSides(result);
  const usedLeft = new Set();
  const usedRight = new Set();

  const groups = _pairOnAmount(sides.left, sides.right, usedLeft, usedRight);
  groups.push(..._collectSumGroups(sides.right, sides.left, usedRight, usedLeft, false, sides.leftLabel.toLowerCase()));
  groups.push(..._collectSumGroups(sides.left, sides.right, usedLeft, usedRight, true, sides.rightLabel.toLowerCase()));

  groups.sort((a, b) => b.score - a.score || a.left[0].row - b.left[0].row);

  const leftGroupNo = new Map();
  const rightGroupNo = new Map();
  groups.forEach((g, i) => {
    g.number = i + 1;
    for (const r of g.left) leftGroupNo.set(r.row, g.number);
    for (const r of g.right) rightGroupNo.set(r.row, g.number);
  });

  return { ...sides, groups, leftGroupNo, rightGroupNo };
}

function uncorrelatedBySheet(result, data) {
  const pick = (r) => ({ row: r.row, date: r.date, description: r.description, amount: r.amount });
  const useStatement = result.hasStatement !== false;
  const sheets = [];

  if (useStatement) {
    sheets.push({
      label: "Bank Statement",
      rows: (result.statement || []).filter((s) => !s.matched && !data.leftGroupNo.has(s.row)).map(pick)
    });
  }
  sheets.push({
    label: "Cashbook",
    rows: data.right.filter((r) => !data.rightGroupNo.has(r.row))
  });
  if (result.hasLedger) {
    const inGroup = useStatement ? () => false : (row) => data.leftGroupNo.has(row);
    sheets.push({
      label: "General Ledger",
      rows: (result.ledger || []).filter((l) => !l.matched && !inGroup(l.row)).map(pick)
    });
  }
  return sheets;
}

const _groupLines = (g) => Math.max(g.left.length, g.right.length);

// The whole Unmatched sheet as an array-of-arrays (same layout as the web app).
function unmatchedComparisonSheetRows(result) {
  const data = buildUnmatchedComparison(result);
  const blank = ["", "", "", ""];
  const side = (entry) => (entry ? [entry.row, entry.date, entry.description, entry.amount] : blank);

  const aoa = [];
  aoa.push([`Unmatched — ${data.left.length} ${data.leftPrefix}, ${data.right.length} ${data.rightPrefix}`]);
  aoa.push([`${data.leftPrefix} Row`, "Date", "Description", "Amount",
            `${data.rightPrefix} Row`, "Date", "Description", "Amount",
            "Correlation"]);

  const listLength = Math.max(data.left.length, data.right.length);
  for (let i = 0; i < listLength; i++) {
    const l = data.left[i] || null;
    const r = data.right[i] || null;
    const refs = [];
    if (l && data.leftGroupNo.has(l.row)) refs.push(`${data.leftPrefix} -> group ${data.leftGroupNo.get(l.row)}`);
    if (r && data.rightGroupNo.has(r.row)) refs.push(`${data.rightPrefix} -> group ${data.rightGroupNo.get(r.row)}`);
    aoa.push([...side(l), ...side(r), refs.join(" | ")]);
  }

  aoa.push([]);
  aoa.push([`Possible correlations — ${data.groups.length} group(s), paired on amount, date, description`]);
  for (const g of data.groups) {
    const lines = _groupLines(g);
    for (let i = 0; i < lines; i++) {
      aoa.push([
        ...side(g.left[i] || null),
        ...side(g.right[i] || null),
        i === 0 ? `Group ${g.number}: ${g.reason}` : ""
      ]);
    }
  }

  aoa.push([]);
  const sheets = uncorrelatedBySheet(result, data);
  const totalUncorrelated = sheets.reduce((n, s) => n + s.rows.length, 0);
  aoa.push([`No correlation — ${totalUncorrelated} row(s)`]);
  for (const sheet of sheets) {
    aoa.push([`${sheet.label} — ${sheet.rows.length} row(s)`]);
    for (const entry of sheet.rows) aoa.push(side(entry));
  }
  return aoa;
}
