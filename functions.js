"use strict";

/**
 * functions.js — Excel custom functions:
 *   =RECON.COMPARETOBS(amount, date, [description], [sheetName])
 *   =RECON.COMPARETOCB(amount, date, [description], [sheetName])
 *   =RECON.COMPARETOGL(amount, date, [description], [sheetName])
 *
 * Each takes the current row's amount + date (and, optionally, description) and
 * returns that row's match status against the chosen target sheet, using the
 * same non-consuming rule as the web app's in-cell "=compareTo…" formulas
 * (engine.js `_match_entry`).
 *
 * These run in the add-in's SHARED runtime (see manifest.xml), so:
 *   - if you've clicked "Load & detect" in the pane, the function reuses that
 *     side's exact column mapping and rows;
 *   - otherwise it reads the target sheet live and auto-detects columns.
 *
 * Only this file + engine.js/utils.js are needed at call time; nothing here
 * touches the task-pane DOM.
 */

const CF_TARGET = {
  bs: { key: "statement", label: "Bank Statement", needles: ["bank", "statement"] },
  cb: { key: "cashbook",  label: "Cashbook",       needles: ["cashbook", "cash book", "cash"] },
  gl: { key: "ledger",    label: "General Ledger", needles: ["ledger", "gl", "general"] },
};

// Short-lived cache of prepared target sides keyed by sheet name, so filling a
// whole column of formulas doesn't re-read the sheet once per cell.
const _cfCache = new Map();
const _CF_TTL_MS = 4000;

// Resolve the target side (array of prepared entries) + a display label.
async function _cfPreparedSide(which, sheetNameArg) {
  const t = CF_TARGET[which];
  const explicit = sheetNameArg != null && String(sheetNameArg).trim() !== "";

  // 1) Prefer whatever the pane already loaded + mapped (honours manual remap
  //    and month filters), unless an explicit sheet name was passed.
  const L = (typeof loaded !== "undefined") ? loaded[t.key] : null;
  if (L && !explicit) {
    return { side: _prepare_side(L.rows, L.mapping), label: t.label };
  }

  // 2) Otherwise read the sheet live and auto-detect its columns.
  return await Excel.run(async (ctx) => {
    const wss = ctx.workbook.worksheets;
    wss.load("items/name");
    await ctx.sync();
    const names = wss.items.map((s) => s.name);

    let name = explicit
      ? names.find((n) => n.toLowerCase() === String(sheetNameArg).trim().toLowerCase())
      : null;
    if (!name) {
      const pool = names.filter((n) => !n.startsWith(RESULT_PREFIX));
      name = pool.find((n) => t.needles.some((k) => n.toLowerCase().includes(k)));
    }
    if (!name) {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        `No ${t.label} sheet found — pass the sheet name as the last argument.`
      );
    }

    const cached = _cfCache.get(name);
    if (cached && Date.now() - cached.at < _CF_TTL_MS) return { side: cached.side, label: t.label };

    const used = ctx.workbook.worksheets.getItem(name).getUsedRangeOrNullObject(true);
    used.load(["values"]);
    await ctx.sync();
    if (used.isNullObject) return { side: [], label: t.label };

    const rows = used.values.map((r) => r.map((v) => (v === null ? "" : v)));
    const analysis = await runLocalAnalyze(rows, [], name);
    const mapping = { ...analysis.mapping, headerRow: analysis.headerRow };
    const side = _prepare_side(rows, mapping);
    _cfCache.set(name, { side, at: Date.now() });
    return { side, label: t.label };
  });
}

async function _compareTo(which, amount, date, description, sheetName) {
  const amt = _amount(amount);
  const hasDate = date !== null && date !== undefined && String(date).trim() !== "";
  if (amt === null && !hasDate) return "";
  if (amt === null) return "No amount";

  const { side, label } = await _cfPreparedSide(which, sheetName);
  const desc = _text(description);
  const keys = _date_keys(date);
  const tail = _description_tail(desc).toLowerCase();
  const [status] = _match_entry(amt, keys, tail, side, desc);

  if (status === "Matched") return `Matched to ${label}`;
  if (status === "Check description") return `Check description (${label})`;
  return `Not found on ${label}`;
}

// Thin per-target wrappers (Excel associates by id, below).
function comparetobs(amount, date, description, sheetName) { return _compareTo("bs", amount, date, description, sheetName); }
function comparetocb(amount, date, description, sheetName) { return _compareTo("cb", amount, date, description, sheetName); }
function comparetogl(amount, date, description, sheetName) { return _compareTo("gl", amount, date, description, sheetName); }

if (typeof CustomFunctions !== "undefined") {
  CustomFunctions.associate("COMPARETOBS", comparetobs);
  CustomFunctions.associate("COMPARETOCB", comparetocb);
  CustomFunctions.associate("COMPARETOGL", comparetogl);
}
