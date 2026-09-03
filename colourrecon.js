"use strict";

/**
 * colourrecon.js — "ColourCode": the user says which cells to reconcile by
 * colouring them in Excel, rather than by naming ranges.
 *
 * The pane shows one line per colour, reading the way it is said out loud:
 *
 *     Sheet1  [purple]  compare with  Sheet3  [purple]
 *                        AND
 *     Sheet1  [blue]    compare with  Sheet3  [blue]
 *
 * and the word between the rules is what they mean together:
 *
 *   OR   each rule is reconciled on its own. A purple cell only cares about the
 *        purple cells on the other sheet, and blue about blue.
 *   AND  the rules must come true together *on the same row*: row 12's purple
 *        and row 12's blue must both find their partner on one single row of
 *        the other sheet. Either both cells of a row go green, or neither does.
 *
 * One colour can say the same thing on its own. Cells of a colour that sit side
 * by side in a row belong together — colouring A23, B23 and C23 purple means
 * "these three, on one row of the other sheet, in any order" — which is the AND
 * again, without needing three colours to say it. Where one side has fewer
 * cells of a colour on its row than the other, it reads as either/or: the cell
 * that pairs goes green and the spare is left blank, because it is neither a
 * match nor a miss. The "cells side by side belong together" tick turns this
 * off, going back to pooling every cell of a colour.
 *
 * Matching itself uses the pane's own matchValues()/matchRows(), with whatever
 * rules are ticked below (ignore sign, ignore cents, tolerance, and the rest).
 *
 * As with the rest of the add-in the only thing ever written is a cell's fill
 * colour. Note that the green/red repaint replaces the colour you marked with,
 * so mark again before running a second pass.
 */

// Reading a cell's fill one cell at a time is far too slow, so this leans on
// getCellProperties (ExcelApi 1.9) to read a whole block of fills in one call.
const CR_CELLS_PER_CALL = 2000;

const CR_DEFAULTS = ["#7030A0", "#00B0F0", "#FFC000", "#92D050"];

// One line of the pane: this colour on this sheet against that colour on that
// sheet. `join` is shared by all of them — the word shown between the lines.
let crRules = [];
let crJoin = "or";

// Sheet names as of the last refresh, so a rule added later still gets a full
// dropdown without another trip to Excel.
let crSheetNames = [];

function initColourRecon() {
  crRules = [newRule(0)];
  $("cr-add").onclick = () => { crRules.push(newRule(crRules.length)); crDrawRules(); };
  $("cr-run").onclick = reconColours;
  crDrawRules();
}

// A new rule lands on the same two sheets as the one above it — a second
// colour is nearly always the same recon seen a second way — with the next
// unused colour on both sides.
function newRule(i) {
  const colour = CR_DEFAULTS[i % CR_DEFAULTS.length];
  const above = crRules[i - 1];
  return {
    a: { sheet: above ? above.a.sheet : (crSheetNames[0] || ""), colour },
    b: { sheet: above ? above.b.sheet : (crSheetNames[1] || crSheetNames[0] || ""), colour },
  };
}

const crHex = (c) => String(c || "").replace("#", "").toUpperCase();

/* ---------- the rule list ---------- */

// Remember the workbook's sheets and give every rule a sheet to point at. Side
// A defaults to the first sheet and side B to the second, which is the usual
// shape of a recon; anything the user has already chosen is kept.
function crSetSheets(names) {
  crSheetNames = names.slice();
  for (const rule of crRules) {
    if (!names.includes(rule.a.sheet)) rule.a.sheet = names[0] || "";
    if (!names.includes(rule.b.sheet)) rule.b.sheet = names[1] || names[0] || "";
  }
  crDrawRules();
}

function crDrawRules() {
  const box = $("cr-rules");
  box.innerHTML = "";
  crRules.forEach((rule, i) => {
    if (i > 0) box.appendChild(crJoinChip());
    box.appendChild(crRuleCard(rule, i));
  });
}

/**
 * The AND / OR between two rules, as a chip the user can press. It also spells
 * out what it means, because the two read very differently.
 */
function crJoinChip() {
  const wrap = document.createElement("div");
  wrap.className = "join " + crJoin;

  const chip = document.createElement("button");
  chip.type = "button";
  chip.className = "join-chip";
  chip.textContent = crJoin.toUpperCase();
  chip.title = "Press to switch between AND and OR";
  chip.onclick = () => { crJoin = crJoin === "and" ? "or" : "and"; crDrawRules(); };

  const note = document.createElement("span");
  note.className = "join-note";
  note.textContent = crJoin === "and"
    ? "both must match, on the same row"
    : "each colour is checked on its own";

  wrap.append(chip, note);
  return wrap;
}

function crRuleCard(rule, i) {
  const card = document.createElement("div");
  card.className = "rule";

  const head = document.createElement("div");
  head.className = "rule-head";
  const title = document.createElement("span");
  title.textContent = "Rule " + (i + 1);
  head.appendChild(title);
  if (crRules.length > 1) {
    const drop = document.createElement("button");
    drop.type = "button";
    drop.className = "drop";
    drop.textContent = "×";
    drop.title = "Remove this rule";
    drop.onclick = () => { crRules.splice(i, 1); crDrawRules(); };
    head.appendChild(drop);
  }

  const middle = document.createElement("p");
  middle.className = "rule-mid";
  middle.textContent = "compare with";

  card.append(head, crRuleSide(rule, "a"), middle, crRuleSide(rule, "b"));
  return card;
}

// "Sheet1 [swatch] ⊙" — the sheet, the colour, and a button that reads both off
// whatever is selected in Excel.
function crRuleSide(rule, side) {
  const row = document.createElement("div");
  row.className = "rule-side";

  const sheet = document.createElement("select");
  for (const n of crSheetNames) sheet.appendChild(new Option(n, n));
  sheet.value = rule[side].sheet;
  sheet.onchange = () => { rule[side].sheet = sheet.value; };

  const swatch = document.createElement("input");
  swatch.type = "color";
  swatch.value = rule[side].colour;
  swatch.title = "The colour to look for";
  swatch.oninput = () => { rule[side].colour = swatch.value; };

  const pick = document.createElement("button");
  pick.type = "button";
  pick.className = "icon pick";
  pick.textContent = "⊙";
  pick.title = "Use the selected cell's sheet and colour";
  pick.onclick = () => crPickColour(rule, side, sheet, swatch, pick);

  row.append(sheet, swatch, pick);
  return row;
}

/* ---------- "use the selected cell" ---------- */

// Read the sheet and fill of whatever is selected in Excel, so the user never
// has to find the sheet in the list or match the shade by eye.
async function crPickColour(rule, side, sheetSel, swatch, btn) {
  btn.disabled = true;
  try {
    let colour = "", sheet = "";
    await Excel.run(async (ctx) => {
      const cell = ctx.workbook.getSelectedRange().getCell(0, 0);
      cell.load(["format/fill/color", "worksheet/name"]);
      await ctx.sync();
      colour = cell.format.fill.color;
      sheet = cell.worksheet.name;
    });
    if (!colour || crHex(colour) === "FFFFFF") {
      throw new Error("that cell has no fill — colour it first, then press this.");
    }
    if (sheet && !crSheetNames.includes(sheet)) await loadSheetList();

    rule[side].colour = "#" + crHex(colour);
    if (sheet) rule[side].sheet = sheet;
    // Redraw rather than poke the two controls: the sheet list may have grown.
    crDrawRules();
    setStatus(`${sheet} · #${crHex(colour)}`);
  } catch (e) {
    setStatus("Could not read the selection: " + e.message, true);
  } finally {
    btn.disabled = false;
  }
}

/* ---------- finding the coloured cells ---------- */

/**
 * Walk `sheet`'s used range once and hand back one group per colour asked for:
 * { hex, sheet, parts, values, rows, cellGroups }. `parts` is the same
 * { column, firstRow, lastRow } shape paintSide() reads — runs of
 * consecutive coloured cells in one column — so paintSide() and "Clear colours"
 * work on it unchanged, while `values` and `rows` line up cell by cell in that
 * same order. `cellGroups` cuts the same cells the other way: side-by-side runs
 * within one row (see crSideBySide).
 */
async function crFindCells(ctx, sheet, hexes) {
  const groups = hexes.map((hex) => ({ hex, sheet, parts: [], values: [], rows: [], cellGroups: [] }));
  const ws = ctx.workbook.worksheets.getItem(sheet);
  const used = ws.getUsedRangeOrNullObject(true);
  used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await ctx.sync();
  if (used.isNullObject) return groups;

  const top = used.rowIndex, left = used.columnIndex;
  const rows = used.rowCount, cols = used.columnCount;
  const block = Math.max(1, Math.floor(CR_CELLS_PER_CALL / cols));

  // per colour: column index -> 1-based rows, plus the value of each hit cell
  const found = hexes.map(() => new Map());
  const valueAt = hexes.map(() => new Map());

  for (let start = 0; start < rows; start += block) {
    const height = Math.min(block, rows - start);
    const range = ws.getRangeByIndexes(top + start, left, height, cols);
    range.load("values");
    const props = range.getCellProperties({ format: { fill: { color: true } } });
    await ctx.sync();

    for (let r = 0; r < height; r++) {
      for (let c = 0; c < cols; c++) {
        const g = hexes.indexOf(crHex(props.value[r][c].format.fill.color));
        if (g < 0) continue;
        const value = range.values[r][c];
        if (_text(value) === "" && _amount(value) === null) continue;   // blanks are not data
        const column = left + c;
        const row = top + start + r + 1;
        if (!found[g].has(column)) found[g].set(column, []);
        found[g].get(column).push(row);
        valueAt[g].set(column + ":" + row, value);
      }
    }
  }

  hexes.forEach((hex, g) => {
    const at = new Map();                          // "column:row" -> index into values
    for (const column of [...found[g].keys()].sort((x, y) => x - y)) {
      const list = found[g].get(column);
      let i = 0;
      while (i < list.length) {
        let j = i;
        while (j + 1 < list.length && list[j + 1] === list[j] + 1) j++;
        groups[g].parts.push({ column, firstRow: list[i], lastRow: list[j] });
        for (let row = list[i]; row <= list[j]; row++) {
          at.set(column + ":" + row, groups[g].values.length);
          groups[g].values.push(valueAt[g].get(column + ":" + row));
          groups[g].rows.push(row);
        }
        i = j + 1;
      }
    }
    groups[g].cellGroups = crSideBySide(found[g], at);
  });
  return groups;
}

/**
 * The cells of one colour cut by row instead of by column: each row's cells are
 * split into runs of neighbouring columns, and each run is one entry
 * { row, cells: [index into the group's values] }.
 *
 * A colour used down a single column gives runs of one cell, which is the plain
 * "this value against those values" recon. A colour spread across A, B and C
 * gives a run of three, and those three are then matched as one thing.
 */
function crSideBySide(byColumn, at) {
  const byRow = new Map();
  for (const [column, rows] of byColumn) {
    for (const row of rows) {
      if (!byRow.has(row)) byRow.set(row, []);
      byRow.get(row).push(column);
    }
  }
  const out = [];
  for (const row of [...byRow.keys()].sort((x, y) => x - y)) {
    const columns = byRow.get(row).sort((x, y) => x - y);
    let i = 0;
    while (i < columns.length) {
      let j = i;
      while (j + 1 < columns.length && columns[j + 1] === columns[j] + 1) j++;
      out.push({ row, cells: columns.slice(i, j + 1).map((c) => at.get(c + ":" + row)) });
      i = j + 1;
    }
  }
  return out;
}

// Every colour a side needs on one sheet, read in a single pass, handed back as
// "which group is rule i's colour" for that sheet.
async function crReadSide(ctx, wants) {
  const bySheet = new Map();
  for (const w of wants) {
    if (!bySheet.has(w.sheet)) bySheet.set(w.sheet, []);
    if (!bySheet.get(w.sheet).includes(w.hex)) bySheet.get(w.sheet).push(w.hex);
  }
  const groups = new Map();                        // "sheet|hex" -> group
  for (const [sheet, hexes] of bySheet) {
    const found = await crFindCells(ctx, sheet, hexes);
    found.forEach((g, i) => groups.set(sheet + "|" + hexes[i], g));
  }
  return wants.map((w) => groups.get(w.sheet + "|" + w.hex));
}

/* ---------- OR: every rule on its own ---------- */

function crMatchOr(groupsA, groupsB, opts) {
  const hitsA = [], hitsB = [];
  groupsA.forEach((ga, i) => {
    const { hitA, hitB } = matchValues(ga.values, groupsB[i].values, opts);
    hitsA.push(hitA.map((h) => h === true));
    hitsB.push(hitB.map((h) => h === true));
  });
  return { hitsA, hitsB };
}

/* ---------- AND: the rules must come true on one row ---------- */

/**
 * Gather one side's groups into rows: row 12 carries rule 1's purple value and
 * rule 2's blue value. Where a colour turns up twice on a row the leftmost cell
 * is the one compared, and the rest of that row's cells simply follow its
 * verdict.
 *
 * Returns the rows (each a value per rule, in rule order) plus, for every rule,
 * which of its cells belong to which row — that is what the green/red is
 * painted from.
 */
function crRowsOf(groups) {
  const order = [];                                // row numbers, first seen first
  const seen = new Map();                          // row number -> index in order
  const values = [];                               // per row: value per rule
  const cells = groups.map(() => []);              // per rule: row index per cell

  groups.forEach((g, r) => {
    g.rows.forEach((row, i) => {
      if (!seen.has(row)) {
        seen.set(row, order.length);
        order.push(row);
        values.push(groups.map(() => undefined));
      }
      const at = seen.get(row);
      if (values[at][r] === undefined) values[at][r] = g.values[i];
      cells[r][i] = at;
    });
  });
  return { order, values, cells };
}

function crMatchAnd(groupsA, groupsB, opts) {
  const a = crRowsOf(groupsA);
  const b = crRowsOf(groupsB);
  const { hitA, hitB } = matchRows(a.values, b.values, opts);
  const spread = (side, hits) => side.cells.map((rowOf) => rowOf.map((at) => hits[at] === true));
  return { hitsA: spread(a, hitA), hitsB: spread(b, hitB) };
}

/* ---------- side by side: several cells of one colour on one row ---------- */

/**
 * A "unit" is one thing to be matched as a whole: a handful of cells, each
 * tagged with the rule (the colour) it came from. Two units match when, colour
 * by colour, the shorter list pairs off one-for-one against the longer one, in
 * any order:
 *
 *   3 cells against 3   every one must find a partner — all green, or all red.
 *   1 cell against 2    either will do — the one that pairs goes green and the
 *                       spare is left blank: no match, but no miss either.
 *
 * A colour that only one of the two units carries can never pair, so that unit
 * misses outright. That is what keeps "purple AND blue" honest: a row with no
 * blue on it has nothing to offer the blue half of the rule.
 */
function crCell(rule, index, value, opts) {
  const key = keyOf(value, opts);
  return { rule, index, key, cents: _keyCents(key) };
}

function crUnit(cells) {
  const bySlot = new Map();                        // rule index -> that rule's cells
  for (const c of cells) {
    if (!bySlot.has(c.rule)) bySlot.set(c.rule, []);
    bySlot.get(c.rule).push(c);
  }
  return { cells, bySlot };
}

// OR: one unit per side-by-side run, each about a single colour.
function crUnitsOr(group, rule, opts) {
  return group.cellGroups.map((cg) =>
    crUnit(cg.cells.map((i) => crCell(rule, i, group.values[i], opts))));
}

// AND: one unit per row, carrying every rule's cells on that row — which is
// what makes the rules come true together on one row.
function crUnitsAnd(groups, opts) {
  const byRow = new Map();
  groups.forEach((g, r) => {
    for (const cg of g.cellGroups) {
      if (!byRow.has(cg.row)) byRow.set(cg.row, []);
      for (const i of cg.cells) byRow.get(cg.row).push(crCell(r, i, g.values[i], opts));
    }
  });
  return [...byRow.keys()].sort((x, y) => x - y).map((row) => crUnit(byRow.get(row)));
}

// Two cells are the same value under the rules that are ticked. A tolerance is
// the one rule a key can't carry, so numbers are compared in cents instead.
const crSame = (x, y, tol) =>
  (tol > 0 && x.cents !== null && y.cents !== null)
    ? Math.abs(x.cents - y.cents) <= tol
    : (x.key !== null && x.key === y.key);

/**
 * Pair the shorter list off against the longer one, in any order, covering the
 * shorter one completely. Returns the pairs as [cell on A, cell on B], or null
 * when it can't be done. The lists are a few cells long, so a plain backtrack
 * is both exact and quick.
 */
function crPairLists(listA, listB, tol) {
  const flip = listA.length > listB.length;
  const small = flip ? listB : listA;
  const large = flip ? listA : listB;
  const used = new Array(large.length).fill(false);
  const chosen = new Array(small.length).fill(-1);

  const walk = (i) => {
    if (i === small.length) return true;
    for (let j = 0; j < large.length; j++) {
      if (used[j] || !crSame(small[i], large[j], tol)) continue;
      used[j] = true; chosen[i] = j;
      if (walk(i + 1)) return true;
      used[j] = false; chosen[i] = -1;
    }
    return false;
  };
  if (!walk(0)) return null;
  return chosen.map((j, i) => (flip ? [large[j], small[i]] : [small[i], large[j]]));
}

// Two units match when every colour in them pairs off.
function crPairUnits(ua, ub, tol) {
  if (ua.bySlot.size !== ub.bySlot.size) return null;
  const pairs = [];
  for (const [rule, listA] of ua.bySlot) {
    const listB = ub.bySlot.get(rule);
    if (!listB) return null;
    const found = crPairLists(listA, listB, tol);
    if (!found) return null;
    pairs.push(...found);
  }
  return pairs;
}

/**
 * Pair the units off one-for-one across the two sides, the way matchValues()
 * pairs single cells: a unit on A claims one unit on B and no other A can have
 * it. Units are shortlisted through an index of their cells' keys so this isn't
 * everything against everything; with an amount tolerance set, numbers can't be
 * hashed and a unit made only of numbers falls back to a scan.
 *
 * Any valid pairing covers the shorter side's cells completely, so at least one
 * of A's keys is always a key of B's unit too — shortlisting on A's keys never
 * loses a match.
 */
function crMatchUnits(unitsA, unitsB, opts) {
  const tol = Math.round((opts.tolerance || 0) * 100);
  const hashable = (c) => c.key !== null && !(tol > 0 && c.cents !== null);

  const index = new Map();
  unitsB.forEach((u, j) => {
    for (const c of u.cells) {
      if (!hashable(c)) continue;
      if (!index.has(c.key)) index.set(c.key, []);
      const list = index.get(c.key);
      if (list[list.length - 1] !== j) list.push(j);
    }
  });
  const everything = unitsB.map((_, j) => j);

  const usedB = new Set();
  const okA = unitsA.map(() => false);
  const okB = unitsB.map(() => false);
  const paired = new Set();                        // "side:rule:index" of cells that found a partner

  unitsA.forEach((ua, i) => {
    const keys = ua.cells.filter(hashable).map((c) => c.key);
    let candidates = everything;
    if (keys.length) {
      const set = new Set();
      for (const key of keys) for (const j of index.get(key) || []) set.add(j);
      candidates = [...set].sort((x, y) => x - y);
    }
    for (const j of candidates) {
      if (usedB.has(j)) continue;
      const pairs = crPairUnits(ua, unitsB[j], tol);
      if (!pairs) continue;
      usedB.add(j);
      okA[i] = true;
      okB[j] = true;
      for (const [ca, cb] of pairs) {
        paired.add("a:" + ca.rule + ":" + ca.index);
        paired.add("b:" + cb.rule + ":" + cb.index);
      }
      break;
    }
  });
  return { okA, okB, paired };
}

/**
 * The verdict per cell: green where the cell found a partner, red where its
 * unit found nothing at all, and blank for a cell sitting in a unit that
 * matched without needing it.
 */
function crMatchGrouped(groupsA, groupsB, opts, join) {
  const hitsA = groupsA.map((g) => g.values.map(() => null));
  const hitsB = groupsB.map((g) => g.values.map(() => null));

  const apply = (units, ok, paired, side, hits) => {
    units.forEach((u, i) => {
      for (const c of u.cells) {
        hits[c.rule][c.index] = !ok[i] ? false
          : paired.has(side + ":" + c.rule + ":" + c.index) ? true : CLEAR;
      }
    });
  };
  const run = (unitsA, unitsB) => {
    const { okA, okB, paired } = crMatchUnits(unitsA, unitsB, opts);
    apply(unitsA, okA, paired, "a", hitsA);
    apply(unitsB, okB, paired, "b", hitsB);
  };

  if (join === "and") run(crUnitsAnd(groupsA, opts), crUnitsAnd(groupsB, opts));
  else groupsA.forEach((ga, r) => run(crUnitsOr(ga, r, opts), crUnitsOr(groupsB[r], r, opts)));

  return { hitsA, hitsB };
}

/* ---------- run ---------- */

async function reconColours() {
  const wantsA = crRules.map((r) => ({ sheet: r.a.sheet, hex: crHex(r.a.colour) }));
  const wantsB = crRules.map((r) => ({ sheet: r.b.sheet, hex: crHex(r.b.colour) }));

  if (wantsA.concat(wantsB).some((w) => !w.sheet)) {
    setStatus("Pick a sheet on both sides of every rule.", true);
    return;
  }
  const clash = crRules.findIndex((r, i) =>
    wantsA[i].sheet === wantsB[i].sheet && wantsA[i].hex === wantsB[i].hex);
  if (clash >= 0) {
    setStatus(`Rule ${clash + 1} is the same colour on the same sheet on both sides — use two sheets, or two colours.`, true);
    return;
  }
  if (!supports("1.9")) {
    setStatus("ColourCode needs a newer Excel (ExcelApi 1.9).", true);
    return;
  }
  const opts = readOpts();

  $("cr-run").disabled = true;
  setStatus("Looking for the coloured cells...");
  try {
    let groupsA, groupsB;
    await Excel.run(async (ctx) => {
      groupsA = await crReadSide(ctx, wantsA);
      groupsB = await crReadSide(ctx, wantsB);
      const missing = [...groupsA, ...groupsB].find((g) => !g.values.length);
      if (missing) throw new Error(`nothing on “${missing.sheet}” is filled with #${missing.hex}.`);

      // Where a colour is used across neighbouring columns the cells of a row
      // are matched as one thing; where every colour sits in a single column
      // there is nothing to group, and the plainer matchers do the same job
      // with less work.
      const join = crRules.length > 1 && crJoin === "and" ? "and" : "or";
      const grouped = $("cr-side-by-side").checked &&
        [...groupsA, ...groupsB].some((g) => g.cellGroups.some((cg) => cg.cells.length > 1));
      const { hitsA, hitsB } = grouped
        ? crMatchGrouped(groupsA, groupsB, opts, join)
        : join === "and"
          ? crMatchAnd(groupsA, groupsB, opts)
          : crMatchOr(groupsA, groupsB, opts);

      setStatus("Colouring...");
      const paint = (groups, hits) => groups.forEach((g, i) => {
        paintSide(ctx.workbook.worksheets.getItem(g.sheet), g, hits[i]);
      });
      paint(groupsA, hitsA);
      paint(groupsB, hitsB);
      await ctx.sync();
    });

    lastPainted = [...groupsA, ...groupsB];
    const count = (groups) => groups.reduce((n, g) => n + g.values.length, 0);
    setStatus(`Done — ${count(groupsA)} cells on the left, ${count(groupsB)} on the right.`);
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    $("cr-run").disabled = false;
  }
}
