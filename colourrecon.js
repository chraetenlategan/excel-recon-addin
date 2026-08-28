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
 * Matching itself is the pane's own matchValues() — one-for-one pairing, with
 * the same ignore-sign / ignore-case options — so a value that appears twice on
 * one side needs to appear twice on the other.
 *
 * As with the rest of the add-in the only thing ever written is a cell's fill
 * colour. Note that the green/red repaint replaces the colour you marked with,
 * so mark again before running a second pass.
 */

// Reading a cell's fill one cell at a time is far too slow, so this leans on
// getCellProperties (ExcelApi 1.9) to read a whole block of fills in one call.
const CR_CELLS_PER_CALL = 2000;

const crState = { a: { sheet: "", colour: "#00B0F0" }, b: { sheet: "", colour: "#00B0F0" } };

const crEl = (side, key) => document.querySelector(`#cr-${side} [data-k="${key}"]`);

function initColourRecon() {
  for (const side of ["a", "b"]) {
    crEl(side, "sheet").onchange = () => { crState[side].sheet = crEl(side, "sheet").value; };
    crEl(side, "colour").oninput = () => { crState[side].colour = crEl(side, "colour").value; };
    crEl(side, "pick").onclick = () => crPickColour(side);
    crState[side].colour = crEl(side, "colour").value;
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

/* ---------- "Use selected cell's colour" ---------- */

// Read the fill of whatever is selected, so the user never has to match the
// shade by eye. The selection also sets that side's sheet.
async function crPickColour(side) {
  const btn = crEl(side, "pick");
  btn.disabled = true;
  try {
    let colour = "", sheet = "";
    await Excel.run(async (ctx) => {
      const sel = ctx.workbook.getSelectedRange();
      const cell = sel.getCell(0, 0);
      cell.load(["format/fill/color", "worksheet/name"]);
      await ctx.sync();
      colour = cell.format.fill.color;
      sheet = cell.worksheet.name;
    });
    if (!colour || crHex(colour) === "FFFFFF") {
      throw new Error("that cell has no fill — colour it first, then press this.");
    }
    crEl(side, "colour").value = "#" + crHex(colour);
    crState[side].colour = crEl(side, "colour").value;
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
 * Every non-blank cell on `sheet` whose fill is `hex`, walked across the used
 * range in blocks. Returned as the same { sheet, parts, values } shape the main
 * pane uses, so paintSide() and "Clear colours" work on it unchanged: runs of
 * consecutive coloured cells in a column become one part, and `values` lines up
 * with those parts in order.
 */
async function crFindCells(ctx, sheet, hex) {
  const ws = ctx.workbook.worksheets.getItem(sheet);
  const used = ws.getUsedRangeOrNullObject(true);
  used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await ctx.sync();
  if (used.isNullObject) return { sheet, parts: [], values: [] };

  const top = used.rowIndex, left = used.columnIndex;
  const rows = used.rowCount, cols = used.columnCount;
  const block = Math.max(1, Math.floor(CR_CELLS_PER_CALL / cols));

  // column index -> sorted list of 1-based row numbers that carry the colour
  const found = new Map();
  const valueAt = new Map();                      // "col:row" -> cell value

  for (let start = 0; start < rows; start += block) {
    const height = Math.min(block, rows - start);
    const range = ws.getRangeByIndexes(top + start, left, height, cols);
    range.load("values");
    const props = range.getCellProperties({ format: { fill: { color: true } } });
    await ctx.sync();

    for (let r = 0; r < height; r++) {
      for (let c = 0; c < cols; c++) {
        if (crHex(props.value[r][c].format.fill.color) !== hex) continue;
        const value = range.values[r][c];
        if (_text(value) === "" && _amount(value) === null) continue;   // blanks are not data
        const column = left + c;
        const row = top + start + r + 1;
        if (!found.has(column)) found.set(column, []);
        found.get(column).push(row);
        valueAt.set(column + ":" + row, value);
      }
    }
  }

  const parts = [];
  const values = [];
  for (const column of [...found.keys()].sort((x, y) => x - y)) {
    const list = found.get(column);
    let i = 0;
    while (i < list.length) {
      let j = i;
      while (j + 1 < list.length && list[j + 1] === list[j] + 1) j++;
      parts.push({ column, firstRow: list[i], lastRow: list[j] });
      for (let row = list[i]; row <= list[j]; row++) values.push(valueAt.get(column + ":" + row));
      i = j + 1;
    }
  }
  return { sheet, parts, values };
}

/* ---------- run ---------- */

async function reconColours() {
  const a = crState.a, b = crState.b;
  if (!a.sheet || !b.sheet) { setStatus("Pick a sheet on both sides.", true); return; }
  if (a.sheet === b.sheet && crHex(a.colour) === crHex(b.colour)) {
    setStatus("Both sides are the same colour on the same sheet — use two sheets, or two colours.", true);
    return;
  }
  if (!supports("1.9")) {
    setStatus("Colour recon needs a newer Excel (ExcelApi 1.9). Use the column compare above instead.", true);
    return;
  }
  const opts = { ignoreSign: $("ignore-sign").checked, ignoreCase: $("ignore-case").checked };

  $("cr-run").disabled = true;
  setStatus("Looking for the coloured cells…");
  try {
    let sa, sb;
    await Excel.run(async (ctx) => {
      sa = await crFindCells(ctx, a.sheet, crHex(a.colour));
      sb = await crFindCells(ctx, b.sheet, crHex(b.colour));
      if (!sa.values.length) throw new Error(`nothing on “${a.sheet}” is filled with that colour.`);
      if (!sb.values.length) throw new Error(`nothing on “${b.sheet}” is filled with that colour.`);

      const { hitA, hitB } = matchValues(sa.values, sb.values, opts);
      setStatus("Colouring…");
      paintSide(ctx.workbook.worksheets.getItem(sa.sheet), sa, hitA);
      paintSide(ctx.workbook.worksheets.getItem(sb.sheet), sb, hitB);
      await ctx.sync();
    });

    lastPainted = [sa, sb];
    setStatus(`Done — ${sa.values.length} cells on A, ${sb.values.length} on B.`);
  } catch (e) {
    setStatus("Error: " + e.message, true);
  } finally {
    $("cr-run").disabled = false;
  }
}
