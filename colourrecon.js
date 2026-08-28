"use strict";

/**
 * colourrecon.js — "Colour recon": the same comparison as the Columns tab, but
 * the user says which cells to read by colouring them in Excel instead of
 * naming ranges.
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
 * { hex, sheet, parts, values, rows }. `parts` is the same { column, firstRow,
 * lastRow } shape the Columns tab uses — runs of consecutive coloured cells in
 * one column — so paintSide() and "Clear colours" work on it unchanged, while
 * `values` and `rows` line up cell by cell in that same order.
 */
async function crFindCells(ctx, sheet, hexes) {
  const groups = hexes.map((hex) => ({ hex, sheet, parts: [], values: [], rows: [] }));
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
    for (const column of [...found[g].keys()].sort((x, y) => x - y)) {
      const list = found[g].get(column);
      let i = 0;
      while (i < list.length) {
        let j = i;
        while (j + 1 < list.length && list[j + 1] === list[j] + 1) j++;
        groups[g].parts.push({ column, firstRow: list[i], lastRow: list[j] });
        for (let row = list[i]; row <= list[j]; row++) {
          groups[g].values.push(valueAt[g].get(column + ":" + row));
          groups[g].rows.push(row);
        }
        i = j + 1;
      }
    }
  });
  return groups;
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
    setStatus("Colour recon needs a newer Excel (ExcelApi 1.9). Use the Columns tab instead.", true);
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

      const { hitsA, hitsB } = crRules.length > 1 && crJoin === "and"
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
