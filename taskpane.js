"use strict";

/**
 * taskpane.js — the shared machinery behind the pane's tools.
 *
 * The matching rules, the key builder, the one-for-one pairing and the painter
 * live here; ColourCode (colourrecon.js) and the PDF finder (pdffinder-pane.js)
 * sit on top of them.
 *
 * Nothing else about the workbook is touched: no values are written, no cells
 * are merged or unmerged, no fonts, borders, number formats or column widths
 * are changed. The only property this file ever sets is a cell's fill colour
 * (and "Clear colours" only ever clears that same fill).
 */

const $ = (id) => document.getElementById(id);

const GREEN = "C6EFCE";
const RED = "FFC7CE";
// A third verdict, used by ColourCode: a cell that was only ever an
// alternative the match didn't need. It is not a match and not a miss, so it
// gets no colour at all — the fill simply comes off.
const CLEAR = "clear";

// The ranges the last run painted, so "Clear colours" can undo exactly them.
let lastPainted = [];

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in only runs in Excel.", true);
    return;
  }
  for (const tab of document.querySelectorAll(".tab")) tab.onclick = () => showTab(tab.dataset.tab);
  $("refresh-sheets").onclick = loadSheetList;
  $("clear-colours").onclick = clearColours;
  initColourRecon();
  initPdfFinder();
  loadSheetList();
});

// Two tools, one pane: ColourCode and the PDF finder each get a tab. The
// matching rules underneath belong to ColourCode — the finder does its own
// matching, against a page — so they come off with it.
function showTab(name) {
  for (const tab of document.querySelectorAll(".tab")) tab.classList.toggle("active", tab.dataset.tab === name);
  for (const panel of document.querySelectorAll(".panel")) panel.classList.toggle("hidden", panel.id !== "tab-" + name);
  $("opts").classList.toggle("hidden", name === "pdf");
  setStatus("");
}

function setStatus(msg, isError) {
  const box = $("status");
  box.textContent = msg || "";
  box.classList.toggle("err", !!isError);
}

/* ---------- sheet list ---------- */

async function loadSheetList() {
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      const names = sheets.items.map((s) => s.name);
      crSetSheets(names);
      pfSetSheets(names);
    });
    setStatus("");
  } catch (e) {
    setStatus("Could not read worksheets: " + e.message, true);
  }
}

// "B" -> 1, "AA" -> 26. The inverse of colLetter().
function colIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

const addressOf = (p) => `${colLetter(p.column)}${p.firstRow}:${colLetter(p.column)}${p.lastRow}`;

/* ---------- matching ---------- */

/**
 * Every matching rule the pane offers, read off the checkboxes once per run and
 * handed to keyOf()/matchValues().
 */
function readOpts() {
  const num = (id) => {
    const v = parseFloat($(id).value);
    return isFinite(v) && v > 0 ? v : 0;
  };
  return {
    ignoreSign: $("ignore-sign").checked,
    ignoreCase: $("ignore-case").checked,
    roundWhole: $("round-whole").checked,
    ignorePunct: $("ignore-punct").checked,
    ignoreZeros: $("ignore-zeros").checked,
    tolerance: num("tolerance"),
    firstChars: Math.round(num("first-chars")),
  };
}

/**
 * The comparison key for one cell. Numbers compare as numbers (in cents, so
 * 100 and 100.00 are one value); anything else compares as text. Blank cells
 * return null and are skipped entirely — never coloured, never matched.
 *
 * The rules bend the key before it is compared: dropping the sign, the cents,
 * the case, the punctuation, the leading zeros, or everything past the first
 * few characters. A tolerance is *not* a key rule — near numbers can't share a
 * key — and is handled by matchValues() instead.
 */
function keyOf(value, opts) {
  const n = _amount(value);
  if (n !== null) {
    let v = opts.ignoreSign ? Math.abs(n) : n;
    if (opts.roundWhole) v = Math.round(v);
    return "n:" + Math.round(v * 100);
  }
  let t = _text(value);
  if (!t) return null;
  if (opts.ignoreCase) t = t.toLowerCase().replace(/\s+/g, " ");
  if (opts.ignorePunct) t = t.replace(/[^\p{L}\p{N}]+/gu, "");
  // "INV-00123" -> "INV-123": zeros that open a run of digits, not zeros
  // sitting inside a number ("1004" is left alone).
  if (opts.ignoreZeros) t = t.replace(/(^|[^\p{N}])0+(?=\p{N})/gu, "$1");
  t = t.trim();
  if (!t) return null;
  if (opts.firstChars) t = t.slice(0, opts.firstChars);
  return "t:" + t;
}

const _keyCents = (key) => (key !== null && key.startsWith("n:") ? parseInt(key.slice(2), 10) : null);

/**
 * Pair the two sides off one-for-one: a value on A claims one equal value from
 * B's pool, so three 100s on A against two on B leave the third 100 red. Where
 * a side spans several columns (or several colours) the pool spans them too —
 * that's the either/or.
 *
 * Returns a boolean per cell on each side (null = blank, leave alone) plus
 * pairA: for each A cell, the index of the B cell it claimed, or -1.
 */
function matchValues(valuesA, valuesB, opts) {
  const keysA = valuesA.map((v) => keyOf(v, opts));
  const keysB = valuesB.map((v) => keyOf(v, opts));
  const tol = Math.round((opts.tolerance || 0) * 100);

  const hitA = keysA.map((k) => (k === null ? null : false));
  const hitB = keysB.map((k) => (k === null ? null : false));
  const pairA = keysA.map(() => -1);
  const claim = (i, j) => { hitA[i] = true; hitB[j] = true; pairA[i] = j; };

  // Exact keys go through a pool, one queue per key. With a tolerance set the
  // numbers are held back for the sweep below and only text pools here.
  const pool = new Map();
  keysB.forEach((k, j) => {
    if (k === null) return;
    if (tol > 0 && _keyCents(k) !== null) return;
    if (!pool.has(k)) pool.set(k, []);
    pool.get(k).push(j);
  });
  keysA.forEach((k, i) => {
    if (k === null) return;
    if (tol > 0 && _keyCents(k) !== null) return;
    const queue = pool.get(k);
    if (queue && queue.length) claim(i, queue.shift());
  });

  // Tolerance: an amount matches anything within ± of it. Both sides are walked
  // in ascending order and each A takes the smallest B still in reach, which
  // pairs off as many rows as can be paired.
  if (tol > 0) {
    const numbered = (keys) => keys
      .map((k, i) => ({ i, cents: _keyCents(k) }))
      .filter((e) => e.cents !== null)
      .sort((x, y) => x.cents - y.cents);
    const numA = numbered(keysA);
    const numB = numbered(keysB);
    let j = 0;
    for (const a of numA) {
      while (j < numB.length && numB[j].cents < a.cents - tol) j++;
      if (j < numB.length && numB[j].cents <= a.cents + tol) claim(a.i, numB[j++].i);
    }
  }

  return { hitA, hitB, pairA };
}

/**
 * Row-level matching, for an AND of several colour rules: a row is a list of
 * values — one per rule — and two rows only match when *every* one of those
 * values matches. That is what makes "blue AND purple" mean "on the same row":
 * the blue cell and the purple cell of a row on side A must both find their
 * partner on one single row of side B.
 *
 * A row with a blank in any of its rules can never match — there is nothing to
 * compare on that rule — so it comes back red.
 *
 * Rows are paired one-for-one, like matchValues(). Exact keys go in a bucket
 * map; where a tolerance is set the numbers can't be hashed, so they are
 * checked inside the bucket their exact parts land in.
 */
function matchRows(rowsA, rowsB, opts) {
  const tol = Math.round((opts.tolerance || 0) * 100);
  const prep = (row) => {
    const exact = [], near = [];
    for (const value of row) {
      const key = keyOf(value, opts);
      if (key === null) return null;
      const cents = _keyCents(key);
      if (tol > 0 && cents !== null) near.push(cents);
      else exact.push(key);
    }
    return { exact: exact.join(""), near };
  };

  const prepA = rowsA.map(prep);
  const prepB = rowsB.map(prep);
  const buckets = new Map();
  prepB.forEach((p, j) => {
    if (!p) return;
    if (!buckets.has(p.exact)) buckets.set(p.exact, []);
    buckets.get(p.exact).push(j);
  });

  const hitA = rowsA.map(() => false);
  const hitB = rowsB.map(() => false);
  prepA.forEach((p, i) => {
    if (!p) return;
    const bucket = buckets.get(p.exact);
    if (!bucket) return;
    for (let n = 0; n < bucket.length; n++) {
      const j = bucket[n];
      if (!p.near.every((c, k) => Math.abs(c - prepB[j].near[k]) <= tol)) continue;
      bucket.splice(n, 1);
      hitA[i] = true;
      hitB[j] = true;
      return;
    }
  });
  return { hitA, hitB };
}

/* ---------- painting ---------- */

// Excel takes at most a handful of areas comfortably in one multi-range call.
const AREAS_PER_CALL = 50;

/**
 * Fill the cells of one side. Consecutive rows with the same verdict are merged
 * into runs, and the runs of a verdict go in as one multi-area call where the
 * host supports it, so a thousand-row column is a few operations.
 */
function paintSide(ws, side, hits) {
  // A verdict per cell -> what happens to its fill. true is a match, false a
  // miss, CLEAR takes the fill off, and null (a blank cell) leaves it alone.
  const paintOf = (hit) =>
    (hit === null || hit === undefined) ? null : hit === CLEAR ? CLEAR : hit ? GREEN : RED;

  const runs = { [GREEN]: [], [RED]: [], [CLEAR]: [] };
  let offset = 0;
  for (const p of side.parts) {
    const n = p.lastRow - p.firstRow + 1;
    const mine = hits.slice(offset, offset + n).map(paintOf);
    offset += n;
    const letter = colLetter(p.column);
    let i = 0;
    while (i < mine.length) {
      const act = mine[i];
      if (act === null) { i++; continue; }
      let j = i;
      while (j + 1 < mine.length && mine[j + 1] === act) j++;
      runs[act].push(`${letter}${p.firstRow + i}:${letter}${p.firstRow + j}`);
      i = j + 1;
    }
  }

  const multi = supports("1.9");
  const apply = (range, act) => {
    if (act === CLEAR) range.format.fill.clear();
    else range.format.fill.color = "#" + act;
  };
  for (const act of [GREEN, RED, CLEAR]) {
    const addrs = runs[act];
    if (!addrs.length) continue;
    if (!multi) {
      for (const a of addrs) apply(ws.getRange(a), act);
      continue;
    }
    for (let k = 0; k < addrs.length; k += AREAS_PER_CALL) {
      apply(ws.getRanges(addrs.slice(k, k + AREAS_PER_CALL).join(",")), act);
    }
  }
}

function supports(version) {
  try { return Office.context.requirements.isSetSupported("ExcelApi", version); } catch { return false; }
}

// Undo: clear the fill on exactly the ranges the last run painted.
async function clearColours() {
  const sides = lastPainted;
  if (!sides.length) {
    setStatus("Nothing has been coloured from here yet.", true);
    return;
  }
  $("clear-colours").disabled = true;
  setStatus("Clearing…");
  try {
    await Excel.run(async (ctx) => {
      for (const side of sides) {
        const ws = ctx.workbook.worksheets.getItem(side.sheet);
        for (const p of side.parts) ws.getRange(addressOf(p)).format.fill.clear();
      }
      await ctx.sync();
    });
    setStatus("Colours cleared.");
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    $("clear-colours").disabled = false;
  }
}
