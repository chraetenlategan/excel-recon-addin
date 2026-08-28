"use strict";

/**
 * colourrecon.js — "Colour recon": the same comparison as the main pane, but
 * the user says which cells to read by colouring them in Excel instead of
 * naming ranges.
 *
 * The flow is: fill the cells you care about on one sheet with a colour (say
 * blue), fill the ones they should match on the other sheet with a colour too,
 * tell the pane "recon blue against blue", and every filled cell is repainted
 * green (matched) or red (not found on the other side).
 *
 * A side can carry several colours, and the mode says what that means:
 *
 *   any   (OR)  the colours pool together — a value is found if it turns up in
 *               any coloured cell on the other side. Blue *or* yellow.
 *   all   (AND) the colours are separate demands — a value is only green if it
 *               is found in every one of the other side's colours. Blue *and*
 *               yellow.
 *   pair        first colour against first colour, second against second, each
 *               pair reconciled on its own. Blue vs blue, yellow vs yellow.
 *
 * Matching itself is the pane's own matchValues(), with whatever rules are
 * ticked above (ignore sign, ignore cents, tolerance, and the rest), so a value
 * appearing twice on one side needs to appear twice on the other.
 *
 * As with the rest of the add-in the only thing ever written is a cell's fill
 * colour. Note that the green/red repaint replaces the colour you marked with,
 * so mark again before running a second pass.
 */

// Reading a cell's fill one cell at a time is far too slow, so this leans on
// getCellProperties (ExcelApi 1.9) to read a whole block of fills in one call.
const CR_CELLS_PER_CALL = 2000;

const CR_DEFAULTS = ["#00B0F0", "#FFFF00", "#92D050", "#FFC000"];

const crState = {
  a: { sheet: "", colours: ["#00B0F0"] },
  b: { sheet: "", colours: ["#00B0F0"] },
};

const crEl = (side, key) => document.querySelector(`#cr-${side} [data-k="${key}"]`);

function initColourRecon() {
  for (const side of ["a", "b"]) {
    crEl(side, "sheet").onchange = () => { crState[side].sheet = crEl(side, "sheet").value; };
    crEl(side, "add").onclick = () => crAddColour(side);
    crEl(side, "pick").onclick = () => crPickColour(side);
    crDrawColours(side);
  }
  $("cr-run").onclick = reconColours;
}

// Keep the colour-recon sheet pickers in step with the main ones.
function crSetSheets(names) {
  for (const side of ["a", "b"]) {
    const sel = crEl(side, "sheet");
    const keep = crState[side].sheet;
    sel.innerHTML = "";
    for (const n of names) sel.appendChild(new Option(n, n));
    const fallback = side === "a" ? names[0] : (names[1] || names[0]);
    sel.value = names.includes(keep) ? keep : (fallback || "");
    crState[side].sheet = sel.value;
  }
}

const crHex = (c) => String(c || "").replace("#", "").toUpperCase();

/* ---------- the colour list ---------- */

// One swatch per colour, each with an x that drops it. The last one keeps its x
// hidden: a side always carries at least one colour.
function crDrawColours(side) {
  const box = crEl(side, "colours");
  const list = crState[side].colours;
  box.innerHTML = "";
  list.forEach((colour, i) => {
    const chip = document.createElement("span");
    chip.className = "chip" + (list.length === 1 ? " only" : "");

    const swatch = document.createElement("input");
    swatch.type = "color";
    swatch.value = colour;
    swatch.title = `Colour ${i + 1}`;
    swatch.oninput = () => { list[i] = swatch.value; };

    const drop = document.createElement("button");
    drop.type = "button";
    drop.textContent = "×";
    drop.title = "Remove this colour";
    drop.onclick = () => { list.splice(i, 1); crDrawColours(side); };

    chip.append(swatch, drop);
    box.appendChild(chip);
  });
}

function crAddColour(side) {
  const list = crState[side].colours;
  const next = CR_DEFAULTS.find((c) => !list.includes(c)) || "#FF0000";
  list.push(next);
  crDrawColours(side);
  setStatus("Set the new swatch, or highlight a cell and use the colour picker button.");
}

/* ---------- "Use selected cell's colour" ---------- */

// Read the fill of whatever is selected, so the user never has to match the
// shade by eye. It fills the last swatch on that side (add one first to keep
// the colour already there) and sets the side's sheet.
async function crPickColour(side) {
  const btn = crEl(side, "pick");
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
    const list = crState[side].colours;
    list[list.length - 1] = "#" + crHex(colour);
    crDrawColours(side);

    const sel = crEl(side, "sheet");
    if (sheet && ![...sel.options].some((o) => o.value === sheet)) await loadSheetList();
    if (sheet) { sel.value = sheet; crState[side].sheet = sheet; }
    setStatus(`Side ${side.toUpperCase()}: #${crHex(colour)} on ${crState[side].sheet}.`);
  } catch (e) {
    setStatus("Could not read the colour: " + e.message, true);
  } finally {
    btn.disabled = false;
  }
}

/* ---------- finding the coloured cells ---------- */

/**
 * Walk `sheet`'s used range once and hand back one group per colour asked for:
 * { hex, sheet, parts, values }, in the same { sheet, parts } shape the main
 * pane uses, so paintSide() and "Clear colours" work on it unchanged. Runs of
 * consecutive coloured cells in a column become one part, and `values` lines up
 * with those parts in order.
 */
async function crFindCells(ctx, sheet, hexes) {
  const groups = hexes.map((hex) => ({ hex, sheet, parts: [], values: [] }));
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
        }
        i = j + 1;
      }
    }
  });
  return groups;
}

/* ---------- the three modes ---------- */

// Flatten a side's groups into one pool, remembering which group each value
// came from so the hits can be handed back out.
function crPool(groups) {
  const values = [];
  const owner = [];
  groups.forEach((g, gi) => g.values.forEach((v, i) => { values.push(v); owner.push([gi, i]); }));
  return { values, owner };
}

const crBlankHits = (groups) => groups.map((g) => g.values.map(() => false));

function crScatter(hits, owner, into) {
  hits.forEach((hit, k) => {
    if (!hit) return;
    const [gi, i] = owner[k];
    into[gi][i] = true;
  });
}

/**
 * Work out which cells are green, per mode. Returns a hit array per group on
 * each side, lined up with that group's values.
 */
function crMatchGroups(mode, groupsA, groupsB, opts) {
  const hitsA = crBlankHits(groupsA);
  const hitsB = crBlankHits(groupsB);

  if (mode === "pair") {
    for (let i = 0; i < groupsA.length; i++) {
      const { hitA, hitB } = matchValues(groupsA[i].values, groupsB[i].values, opts);
      hitA.forEach((h, k) => { if (h) hitsA[i][k] = true; });
      hitB.forEach((h, k) => { if (h) hitsB[i][k] = true; });
    }
    return { hitsA, hitsB };
  }

  if (mode === "any") {
    const a = crPool(groupsA), b = crPool(groupsB);
    const { hitA, hitB } = matchValues(a.values, b.values, opts);
    crScatter(hitA, a.owner, hitsA);
    crScatter(hitB, b.owner, hitsB);
    return { hitsA, hitsB };
  }

  // "all": every colour on the far side is a separate demand, so a value is
  // green only where it was found in each of them. The far side's own cells go
  // green where the value that claimed them came out green overall — a partial
  // hit colours nothing. Both directions are worked out the same way: A against
  // B's colours, then B against A's.
  const demand = (mine, theirs) => {
    const pool = crPool(mine);
    const runs = [];
    let keep = pool.values.map(() => true);
    for (const g of theirs) {
      const { hitA, pairA } = matchValues(pool.values, g.values, opts);
      runs.push(pairA);
      keep = keep.map((k, i) => k && hitA[i] === true);
    }
    return { pool, keep, runs };
  };

  const paintFar = (from, farGroups, farHits) => {
    from.runs.forEach((pairA, g) => {
      pairA.forEach((j, i) => {
        if (j >= 0 && from.keep[i]) farHits[g][j] = true;
      });
    });
  };

  const fromA = demand(groupsA, groupsB);
  const fromB = demand(groupsB, groupsA);
  crScatter(fromA.keep, fromA.pool.owner, hitsA);
  crScatter(fromB.keep, fromB.pool.owner, hitsB);
  paintFar(fromA, groupsB, hitsB);
  paintFar(fromB, groupsA, hitsA);
  return { hitsA, hitsB };
}

/* ---------- run ---------- */

async function reconColours() {
  const a = crState.a, b = crState.b;
  const mode = $("cr-mode").value;
  const hexA = [...new Set(a.colours.map(crHex))];
  const hexB = [...new Set(b.colours.map(crHex))];

  if (!a.sheet || !b.sheet) { setStatus("Pick a sheet on both sides.", true); return; }
  if (mode === "pair" && hexA.length !== hexB.length) {
    setStatus("Colour by colour needs the same number of colours on both sides.", true);
    return;
  }
  if (a.sheet === b.sheet && hexA.some((h) => hexB.includes(h))) {
    setStatus("The same colour is on both sides of one sheet — use two sheets, or two colours.", true);
    return;
  }
  if (!supports("1.9")) {
    setStatus("Colour recon needs a newer Excel (ExcelApi 1.9). Use the column compare above instead.", true);
    return;
  }
  const opts = readOpts();

  $("cr-run").disabled = true;
  setStatus("Looking for the coloured cells...");
  try {
    let groupsA, groupsB;
    await Excel.run(async (ctx) => {
      groupsA = await crFindCells(ctx, a.sheet, hexA);
      groupsB = await crFindCells(ctx, b.sheet, hexB);
      const missing = [...groupsA, ...groupsB].find((g) => !g.values.length);
      if (missing) throw new Error(`nothing on “${missing.sheet}” is filled with #${missing.hex}.`);

      const { hitsA, hitsB } = crMatchGroups(mode, groupsA, groupsB, opts);
      setStatus("Colouring...");
      const wsA = ctx.workbook.worksheets.getItem(a.sheet);
      const wsB = ctx.workbook.worksheets.getItem(b.sheet);
      groupsA.forEach((g, i) => paintSide(wsA, g, hitsA[i]));
      groupsB.forEach((g, i) => paintSide(wsB, g, hitsB[i]));
      await ctx.sync();
    });

    lastPainted = [...groupsA, ...groupsB];
    const count = (groups) => groups.reduce((n, g) => n + g.values.length, 0);
    setStatus(`Done — ${count(groupsA)} cells on A, ${count(groupsB)} on B.`);
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    $("cr-run").disabled = false;
  }
}
