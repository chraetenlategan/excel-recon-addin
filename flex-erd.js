"use strict";

/**
 * flex-erd.js — step 2 of Modular Recon: the ERD canvas.
 *
 * Every included sheet is drawn as an entity box and every column as an
 * attribute inside it, exactly like a database diagram. The user drags from an
 * attribute on one table to an attribute on another, and that line *is* the
 * instruction: "compare these two columns". Several lines between the same two
 * tables build a composite key — all of them have to agree for two rows to
 * match — which is how the classic date+description+amount recon is expressed
 * here without the app knowing what a date or a description is.
 *
 * Drawing rather than form-filling is the point: the user can see at a glance
 * what will be compared with what, which is the part that used to be locked
 * into the fixed three-sheet layout.
 *
 * Layout maths lives in world coordinates: entity positions are plain numbers
 * on `sheet.pos`, and the whole world is pan/zoomed with one CSS transform, so
 * nothing has to be recalculated when the user moves around.
 *
 * Ported from the web app. A task pane is narrow enough that dragging a wire
 * can be fiddly, so the rail also carries a two-dropdown form that adds exactly
 * the same link — the canvas stays the way you read the model, not the only way
 * to build it.
 */

const FLEX_ENTITY_WIDTH = 190;
const FLEX_COLLAPSE_AFTER = 10;   // columns before an entity offers "linked only"

/* ---------- layout ---------- */

// Place any entity that hasn't got a position yet: a simple grid, wide enough
// apart that the wires between them have room to breathe.
function flexAutoLayout(force = false) {
  const sheets = flexIncludedSheets();
  sheets.forEach((sheet, i) => {
    if (sheet.pos && !force) return;
    sheet.pos = { x: 30 + (i % 2) * (FLEX_ENTITY_WIDTH + 120), y: 30 + Math.floor(i / 2) * 360 };
  });
}

/* ---------- rendering ---------- */

function renderFlexErd() {
  const world = $("flex-world");
  const svg = $("flex-wires");
  if (!world || !svg) return;

  flexAutoLayout();
  const sheets = flexIncludedSheets();

  $("flex-canvas-empty").classList.toggle("hidden", sheets.length >= 2);

  // Rebuild the entity boxes (the <svg> of wires stays as world's first child).
  world.querySelectorAll(".flex-entity").forEach((el) => el.remove());
  for (const sheet of sheets) world.appendChild(flexEntityBox(sheet));

  flexApplyView();
  flexDrawWires();
  renderFlexRail();
  flexUpdateButtons();
}

function flexEntityBox(sheet) {
  const box = document.createElement("div");
  box.className = "flex-entity";
  box.dataset.sheet = sheet.id;
  box.style.left = `${sheet.pos.x}px`;
  box.style.top = `${sheet.pos.y}px`;
  box.style.width = `${FLEX_ENTITY_WIDTH}px`;

  const linked = flexLinkedCols(sheet.id);
  const collapsible = sheet.columns.length > FLEX_COLLAPSE_AFTER;
  const collapsed = collapsible && sheet.collapsed;
  const rowCount = Math.max(0, sheet.rows.length - sheet.headerRow);

  const head = document.createElement("div");
  head.className = "flex-entity-head";
  head.innerHTML = `
    <span class="flex-entity-name">${escapeHtml(sheet.label)}</span>
    <span class="flex-entity-count">${rowCount}</span>`;
  if (collapsible) {
    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "flex-entity-toggle";
    toggle.textContent = collapsed ? "all" : "linked";
    toggle.title = collapsed ? "Show every column" : "Show only the columns you have linked";
    toggle.addEventListener("pointerdown", (e) => e.stopPropagation());
    toggle.addEventListener("click", (e) => {
      e.stopPropagation();
      sheet.collapsed = !sheet.collapsed;
      renderFlexErd();
    });
    head.appendChild(toggle);
  }
  box.appendChild(head);

  const list = document.createElement("div");
  list.className = "flex-attrs";
  const visible = collapsed
    ? sheet.columns.filter((c) => linked.has(c.index))
    : sheet.columns;

  if (!visible.length) {
    list.innerHTML = `<div class="flex-attr-empty">No columns linked yet</div>`;
  }
  for (const col of visible) {
    const attr = document.createElement("div");
    attr.className = `flex-attr${linked.has(col.index) ? " linked" : ""}`;
    attr.dataset.col = col.index;
    attr.dataset.sheet = sheet.id;
    attr.title = col.samples.length ? `e.g. ${col.samples.join(", ")}` : "";
    attr.innerHTML = `
      <span class="flex-port flex-port-l" data-side="left"></span>
      <span class="flex-attr-name">${escapeHtml(col.name)}</span>
      <span class="flex-attr-type type-${col.type}">${col.type}</span>
      <span class="flex-port flex-port-r" data-side="right"></span>`;
    list.appendChild(attr);
  }
  if (collapsed) {
    const more = document.createElement("div");
    more.className = "flex-attr-more";
    more.textContent = `${sheet.columns.length - visible.length} more column${sheet.columns.length - visible.length === 1 ? "" : "s"} hidden`;
    list.appendChild(more);
  }
  box.appendChild(list);
  return box;
}

// Column indices on a sheet that already take part in a link.
function flexLinkedCols(sheetId) {
  const out = new Set();
  for (const link of flexState.links) {
    if (link.from.sheet === sheetId) out.add(link.from.col);
    if (link.to.sheet === sheetId) out.add(link.to.col);
  }
  return out;
}

/* ---------- wires ---------- */

// The anchor point of one end of a wire, in world coordinates. Falls back to
// the entity header when the attribute is hidden by a collapsed box.
function flexAnchor(sheetId, col, side) {
  const sheet = flexSheet(sheetId);
  const box = document.querySelector(`.flex-entity[data-sheet="${sheetId}"]`);
  if (!sheet || !sheet.pos || !box) return null;
  const attr = box.querySelector(`.flex-attr[data-col="${col}"]`);
  const target = attr || box.querySelector(".flex-entity-head");
  const y = sheet.pos.y + target.offsetTop + target.offsetHeight / 2;
  const x = side === "right" ? sheet.pos.x + box.offsetWidth : sheet.pos.x;
  return { x, y };
}

// Which sides the wire should leave from: whichever pair faces the other box.
function flexSidesFor(fromSheetId, toSheetId) {
  const a = flexSheet(fromSheetId), b = flexSheet(toSheetId);
  if (!a || !b || !a.pos || !b.pos) return ["right", "left"];
  return a.pos.x + FLEX_ENTITY_WIDTH / 2 <= b.pos.x + FLEX_ENTITY_WIDTH / 2
    ? ["right", "left"] : ["left", "right"];
}

function flexCurve(p1, side1, p2, side2) {
  const dx = Math.max(50, Math.abs(p2.x - p1.x) * 0.45);
  const c1 = { x: p1.x + (side1 === "right" ? dx : -dx), y: p1.y };
  const c2 = { x: p2.x + (side2 === "right" ? dx : -dx), y: p2.y };
  return {
    d: `M ${p1.x} ${p1.y} C ${c1.x} ${c1.y}, ${c2.x} ${c2.y}, ${p2.x} ${p2.y}`,
    mid: { x: (p1.x + 3 * c1.x + 3 * c2.x + p2.x) / 8, y: (p1.y + 3 * c1.y + 3 * c2.y + p2.y) / 8 },
  };
}

function flexDrawWires(temp = null) {
  const svg = $("flex-wires");
  if (!svg) return;
  const NS = "http://www.w3.org/2000/svg";
  svg.replaceChildren();

  for (const link of flexState.links) {
    const [sideA, sideB] = flexSidesFor(link.from.sheet, link.to.sheet);
    const p1 = flexAnchor(link.from.sheet, link.from.col, sideA);
    const p2 = flexAnchor(link.to.sheet, link.to.col, sideB);
    if (!p1 || !p2) continue;
    const { d, mid } = flexCurve(p1, sideA, p2, sideB);

    // A fat invisible path under the visible one, so the line is easy to click.
    const hit = document.createElementNS(NS, "path");
    hit.setAttribute("d", d);
    hit.setAttribute("class", "flex-wire-hit");
    hit.dataset.link = link.id;
    svg.appendChild(hit);

    const path = document.createElementNS(NS, "path");
    path.setAttribute("d", d);
    path.setAttribute("class", `flex-wire${flexState.selectedLink === link.id ? " selected" : ""}`);
    path.dataset.link = link.id;
    svg.appendChild(path);

    for (const p of [p1, p2]) {
      const dot = document.createElementNS(NS, "circle");
      dot.setAttribute("cx", p.x);
      dot.setAttribute("cy", p.y);
      dot.setAttribute("r", 4);
      dot.setAttribute("class", "flex-wire-dot");
      svg.appendChild(dot);
    }

    const text = document.createElementNS(NS, "text");
    text.setAttribute("x", mid.x);
    text.setAttribute("y", mid.y - 6);
    text.setAttribute("class", "flex-wire-label");
    text.setAttribute("text-anchor", "middle");
    text.dataset.link = link.id;
    text.textContent = flexLinkLabel(link);
    svg.appendChild(text);
  }

  if (temp) {
    const path = document.createElementNS(NS, "path");
    path.setAttribute("d", temp);
    path.setAttribute("class", "flex-wire temp");
    svg.appendChild(path);
  }
}

function flexLinkLabel(link) {
  const col = flexColumn(link.from.sheet, link.from.col);
  const mode = link.mode === "auto" ? (col ? col.type : "text") : link.mode;
  const bits = [mode];
  if (mode === "number" && link.opts.absolute) bits.push("±");
  if (mode === "text" && link.opts.loose) bits.push("loose");
  return bits.join(" ");
}

/* ---------- pan / zoom ---------- */

function flexApplyView() {
  const world = $("flex-world");
  const { x, y, scale } = flexState.view;
  world.style.transform = `translate(${x}px, ${y}px) scale(${scale})`;
}

function flexZoom(factor, centre) {
  const canvas = $("flex-canvas");
  const rect = canvas.getBoundingClientRect();
  const cx = centre ? centre.x - rect.left : rect.width / 2;
  const cy = centre ? centre.y - rect.top : rect.height / 2;
  const view = flexState.view;
  const next = Math.min(2, Math.max(0.25, view.scale * factor));
  // Keep the point under the cursor fixed while the scale changes.
  view.x = cx - (cx - view.x) * (next / view.scale);
  view.y = cy - (cy - view.y) * (next / view.scale);
  view.scale = next;
  flexApplyView();
}

function flexFit() {
  const sheets = flexIncludedSheets();
  if (!sheets.length) return;
  const rect = $("flex-canvas").getBoundingClientRect();
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const sheet of sheets) {
    const box = document.querySelector(`.flex-entity[data-sheet="${sheet.id}"]`);
    const h = box ? box.offsetHeight : 200;
    minX = Math.min(minX, sheet.pos.x);
    minY = Math.min(minY, sheet.pos.y);
    maxX = Math.max(maxX, sheet.pos.x + FLEX_ENTITY_WIDTH);
    maxY = Math.max(maxY, sheet.pos.y + h);
  }
  const pad = 16;
  const scale = Math.min(1.2, Math.max(0.2,
    Math.min((rect.width - pad * 2) / (maxX - minX), (rect.height - pad * 2) / (maxY - minY))));
  flexState.view = {
    scale,
    x: pad + (rect.width - pad * 2 - (maxX - minX) * scale) / 2 - minX * scale,
    y: pad + (rect.height - pad * 2 - (maxY - minY) * scale) / 2 - minY * scale,
  };
  flexApplyView();
}

// Screen point -> world point.
function flexToWorld(clientX, clientY) {
  const rect = $("flex-canvas").getBoundingClientRect();
  const { x, y, scale } = flexState.view;
  return { x: (clientX - rect.left - x) / scale, y: (clientY - rect.top - y) / scale };
}

/* ---------- the comparisons rail ---------- */

// The same information as the wires, in words. A diagram is quick to read but
// slow to audit; this side list says exactly what will be compared, and is the
// reliable place to remove a link you can't easily click.
function renderFlexRail() {
  const rail = $("flex-rail");
  if (!rail) return;
  rail.replaceChildren(flexAddLinkForm());

  const rels = flexRelationships();
  if (!rels.length) {
    const empty = document.createElement("p");
    empty.className = "flex-rail-empty";
    empty.textContent = "Nothing drawn yet. Drag from a column on one table to a column on another — or use the two pickers above.";
    rail.appendChild(empty);
    return;
  }

  const heading = document.createElement("h3");
  heading.innerHTML = `Comparisons <span class="flex-rail-count">${rels.length}</span>`;
  rail.appendChild(heading);

  for (const rel of rels) {
    const block = document.createElement("div");
    block.className = "flex-rel";
    block.innerHTML = `<div class="flex-rel-head">${escapeHtml(flexRelLabel(rel))}</div>`;
    const list = document.createElement("div");
    list.className = "flex-rel-fields";
    for (const field of rel.fields) {
      const row = document.createElement("div");
      row.className = "flex-rel-field";
      row.dataset.link = field.link.id;
      row.innerHTML = `
        <span class="flex-rel-col">${escapeHtml(flexColumnName(rel.left, field.leftCol))}</span>
        <span class="flex-rel-arrow">=</span>
        <span class="flex-rel-col">${escapeHtml(flexColumnName(rel.right, field.rightCol))}</span>
        <span class="flex-rel-mode">${escapeHtml(flexLinkLabel(field.link))}</span>
        <button type="button" class="flex-rel-del" title="Remove this comparison">✕</button>`;
      row.addEventListener("click", (e) => {
        if (e.target.classList.contains("flex-rel-del")) {
          if (flexState.selectedLink === field.link.id) flexCloseLinkMenu();
          flexRemoveLink(field.link.id);
          renderFlexErd();
          return;
        }
        flexOpenLinkMenu(field.link.id, row);
      });
      list.appendChild(row);
    }
    block.appendChild(list);
    if (rel.fields.length > 1) {
      const note = document.createElement("p");
      note.className = "flex-rel-note";
      note.textContent = `Rows match only when all ${rel.fields.length} agree.`;
      block.appendChild(note);
    }
    rail.appendChild(block);
  }
}

// Two pickers and a button: the same flexAddLink() the drag gesture calls.
// A pane is narrow, and a precise drag across a zoomed canvas is a poor way to
// have to work — this is the keyboard-and-mouse route to the same model.
function flexAddLinkForm() {
  const form = document.createElement("div");
  form.className = "flex-add-link";
  const options = () => flexIncludedSheets().map((sheet) =>
    `<optgroup label="${escapeHtml(sheet.label)}">${sheet.columns.map((c) =>
      `<option value="${sheet.id}:${c.index}">${escapeHtml(c.name)}</option>`).join("")}</optgroup>`).join("");

  form.innerHTML = `
    <h3>Add a comparison</h3>
    <div class="flex-add-row">
      <select id="flex-add-a">${options()}</select>
      <span class="flex-rel-arrow">=</span>
      <select id="flex-add-b">${options()}</select>
      <button type="button" id="flex-add-go" class="flex-add-go">Add</button>
    </div>`;

  form.querySelector("#flex-add-go").addEventListener("click", () => {
    const parse = (id) => {
      const value = form.querySelector(id).value;
      if (!value) return null;
      const [sheet, col] = value.split(":");
      return { sheet, col: parseInt(col, 10) };
    };
    const a = parse("#flex-add-a"), b = parse("#flex-add-b");
    if (!a || !b) return;
    if (a.sheet === b.sheet) { setStatus("Pick columns on two different sheets.", true); return; }
    const link = flexAddLink(a, b);
    if (!link) { setStatus("Those two columns are already compared.", true); return; }
    setStatus("");
    renderFlexErd();
  });
  return form;
}

/* ---------- link editor ---------- */

const flexLinkMenu = document.createElement("div");
flexLinkMenu.id = "flex-link-menu";
flexLinkMenu.className = "flex-link-menu hidden";
document.body.appendChild(flexLinkMenu);

// Where the menu was opened from, so re-rendering it after a mode change (which
// swaps which options are relevant) keeps it in the same place.
let flexMenuAnchor = null;

function flexCloseLinkMenu() {
  flexLinkMenu.classList.add("hidden");
  if (!flexState.selectedLink) return;
  flexState.selectedLink = null;
  flexDrawWires();
}

// `anchor` is either an element to hang the menu under, or a {x, y} client point
// (a wire has no box worth measuring — the whole SVG is thousands of pixels wide).
function flexOpenLinkMenu(linkId, anchor) {
  const link = flexLink(linkId);
  if (!link) return;
  flexState.selectedLink = linkId;
  flexMenuAnchor = anchor || flexMenuAnchor;

  const fromCol = flexColumn(link.from.sheet, link.from.col);
  const mode = link.mode === "auto" ? (fromCol ? fromCol.type : "text") : link.mode;

  flexLinkMenu.innerHTML = `
    <div class="flex-menu-title">
      ${escapeHtml(flexSheet(link.from.sheet).label)}.<b>${escapeHtml(flexColumnName(link.from.sheet, link.from.col))}</b>
      <span class="flex-menu-eq">=</span>
      ${escapeHtml(flexSheet(link.to.sheet).label)}.<b>${escapeHtml(flexColumnName(link.to.sheet, link.to.col))}</b>
    </div>
    <label class="flex-menu-row">Compare as
      <select data-field="mode">${FLEX_MODES.map(([value, label]) =>
        `<option value="${value}"${link.mode === value ? " selected" : ""}>${label}</option>`).join("")}</select>
    </label>
    <label class="flex-menu-row ${mode === "text" ? "" : "hidden"}">
      <input type="checkbox" data-field="caseInsensitive" ${link.opts.caseInsensitive !== false ? "checked" : ""}> Ignore capitals
    </label>
    <label class="flex-menu-row ${mode === "text" ? "" : "hidden"}">
      <input type="checkbox" data-field="loose" ${link.opts.loose ? "checked" : ""}> Ignore spaces &amp; punctuation
    </label>
    <label class="flex-menu-row ${mode === "number" ? "" : "hidden"}">Decimals
      <input type="number" data-field="decimals" min="0" max="6" value="${Number.isFinite(link.opts.decimals) ? link.opts.decimals : 2}">
    </label>
    <label class="flex-menu-row ${mode === "number" ? "" : "hidden"}">
      <input type="checkbox" data-field="absolute" ${link.opts.absolute ? "checked" : ""}> Ignore + / − sign
    </label>
    <p class="flex-menu-note">Values are compared exactly once normalised — never approximately.</p>
    <button type="button" class="flex-menu-delete">Remove this comparison</button>`;

  flexLinkMenu.querySelector('[data-field="mode"]').addEventListener("change", (e) => {
    link.mode = e.target.value;
    flexState.result = null;
    flexOpenLinkMenu(linkId);   // re-open so the right options show
    renderFlexRail();
    flexDrawWires();
  });
  flexLinkMenu.querySelectorAll("input[data-field]").forEach((input) => {
    input.addEventListener("change", () => {
      const field = input.dataset.field;
      link.opts[field] = input.type === "checkbox" ? input.checked : Number(input.value);
      flexState.result = null;
      renderFlexRail();
      flexDrawWires();
      flexUpdateButtons();
    });
  });
  flexLinkMenu.querySelector(".flex-menu-delete").addEventListener("click", () => {
    flexRemoveLink(linkId);
    flexCloseLinkMenu();
    renderFlexErd();
  });

  flexLinkMenu.classList.remove("hidden");
  const from = flexMenuAnchor && flexMenuAnchor.getBoundingClientRect
    ? (() => { const r = flexMenuAnchor.getBoundingClientRect(); return { x: r.left, y: r.bottom }; })()
    : (flexMenuAnchor || { x: 20, y: 80 });
  const menuRect = flexLinkMenu.getBoundingClientRect();
  const left = Math.min(from.x + window.scrollX, window.innerWidth - menuRect.width - 8);
  const top = Math.min(from.y + window.scrollY + 6, window.innerHeight - menuRect.height - 8);
  flexLinkMenu.style.left = `${Math.max(4, left)}px`;
  flexLinkMenu.style.top = `${Math.max(4, top)}px`;
  flexDrawWires();
}

/* ---------- suggestions ---------- */

// Offer the obvious links: columns that carry the same name on two sheets.
// Deliberately conservative — a wrong suggestion the user doesn't notice is
// worse than no suggestion at all.
function flexSuggestLinks() {
  const norm = (s) => s.toLowerCase().replace(/[^a-z0-9]/g, "");
  const sheets = flexIncludedSheets();
  let added = 0;
  for (let i = 0; i < sheets.length; i++) {
    for (let j = i + 1; j < sheets.length; j++) {
      for (const a of sheets[i].columns) {
        if (!norm(a.name) || a.name === a.letter) continue;
        for (const b of sheets[j].columns) {
          if (norm(a.name) !== norm(b.name)) continue;
          if (flexAddLink({ sheet: sheets[i].id, col: a.index }, { sheet: sheets[j].id, col: b.index })) added++;
        }
      }
    }
  }
  setStatus(added
    ? `Added ${added} comparison${added === 1 ? "" : "s"} from matching column names — check them before running.`
    : "No columns share a name across your sheets — draw the comparisons by hand.", !added);
  renderFlexErd();
}

/* ---------- interaction wiring ---------- */

function initFlexErd() {
  const canvas = $("flex-canvas");

  $("flex-suggest").addEventListener("click", flexSuggestLinks);
  $("flex-arrange").addEventListener("click", () => {
    flexAutoLayout(true);
    renderFlexErd();
    flexFit();
  });
  $("flex-zoom-in").addEventListener("click", () => flexZoom(1.2));
  $("flex-zoom-out").addEventListener("click", () => flexZoom(1 / 1.2));
  $("flex-fit").addEventListener("click", flexFit);
  $("flex-clear-links").addEventListener("click", () => {
    if (!flexState.links.length) return;
    flexState.links = [];
    flexState.result = null;
    renderFlexErd();
  });

  canvas.addEventListener("wheel", (e) => {
    e.preventDefault();
    flexZoom(e.deltaY < 0 ? 1.1 : 1 / 1.1, { x: e.clientX, y: e.clientY });
  }, { passive: false });

  // One pointerdown handler decides what the gesture is: open a wire's settings,
  // draw a new wire (started on a port), move an entity (started on its header),
  // or pan the canvas.
  //
  // Wires are opened on pointerdown rather than click on purpose: opening the
  // menu redraws the SVG, which would detach the very path a later click event
  // needed to land on.
  canvas.addEventListener("pointerdown", (e) => {
    if (e.button !== 0) return;
    const wireId = e.target.dataset && e.target.dataset.link;
    if (wireId) return flexOpenLinkMenu(wireId, { x: e.clientX, y: e.clientY });
    const port = e.target.closest(".flex-port");
    const head = e.target.closest(".flex-entity-head");
    if (port) return flexStartWire(e, port);
    if (head) return flexStartMove(e, head.closest(".flex-entity"));
    if (!e.target.closest(".flex-entity")) flexStartPan(e);
  });

  document.addEventListener("pointerdown", (e) => {
    if (!flexLinkMenu.contains(e.target) && !e.target.closest(".flex-wire-hit, .flex-wire, .flex-wire-label, .flex-rel-field")) {
      flexCloseLinkMenu();
    }
  });

  // Keep the canvas usable when the pane is resized.
  window.addEventListener("resize", () => { if (flexModeActive()) flexDrawWires(); });
}

// Drag from a port to another table's attribute.
function flexStartWire(e, port) {
  e.preventDefault();
  const attr = port.closest(".flex-attr");
  const from = { sheet: attr.dataset.sheet, col: parseInt(attr.dataset.col, 10) };
  const side = port.dataset.side;
  const start = flexAnchor(from.sheet, from.col, side);
  if (!start) return;
  document.body.classList.add("flex-wiring");

  const onMove = (ev) => {
    const p = flexToWorld(ev.clientX, ev.clientY);
    const endSide = p.x < start.x ? "right" : "left";
    flexDrawWires(flexCurve(start, side, p, endSide).d);
    const over = document.elementFromPoint(ev.clientX, ev.clientY);
    const hover = over && over.closest(".flex-attr");
    document.querySelectorAll(".flex-attr.drop-target").forEach((el) => el.classList.remove("drop-target"));
    if (hover && hover.dataset.sheet !== from.sheet) hover.classList.add("drop-target");
  };

  const onUp = (ev) => {
    document.removeEventListener("pointermove", onMove);
    document.removeEventListener("pointerup", onUp);
    document.body.classList.remove("flex-wiring");
    document.querySelectorAll(".flex-attr.drop-target").forEach((el) => el.classList.remove("drop-target"));

    const over = document.elementFromPoint(ev.clientX, ev.clientY);
    const target = over && over.closest(".flex-attr");
    if (target && target.dataset.sheet !== from.sheet) {
      const to = { sheet: target.dataset.sheet, col: parseInt(target.dataset.col, 10) };
      const link = flexAddLink(from, to);
      renderFlexErd();
      if (link) flexOpenLinkMenu(link.id, target);
      return;
    }
    flexDrawWires();
  };

  document.addEventListener("pointermove", onMove);
  document.addEventListener("pointerup", onUp);
}

// Drag an entity box by its header.
function flexStartMove(e, box) {
  e.preventDefault();
  const sheet = flexSheet(box.dataset.sheet);
  const startX = e.clientX, startY = e.clientY;
  const origin = { ...sheet.pos };
  box.classList.add("moving");

  const onMove = (ev) => {
    sheet.pos.x = Math.round(origin.x + (ev.clientX - startX) / flexState.view.scale);
    sheet.pos.y = Math.round(origin.y + (ev.clientY - startY) / flexState.view.scale);
    box.style.left = `${sheet.pos.x}px`;
    box.style.top = `${sheet.pos.y}px`;
    flexDrawWires();
  };
  const onUp = () => {
    document.removeEventListener("pointermove", onMove);
    document.removeEventListener("pointerup", onUp);
    box.classList.remove("moving");
  };
  document.addEventListener("pointermove", onMove);
  document.addEventListener("pointerup", onUp);
}

// Drag the empty canvas to pan.
function flexStartPan(e) {
  const canvas = $("flex-canvas");
  const startX = e.clientX, startY = e.clientY;
  const origin = { x: flexState.view.x, y: flexState.view.y };
  canvas.classList.add("panning");

  const onMove = (ev) => {
    flexState.view.x = origin.x + (ev.clientX - startX);
    flexState.view.y = origin.y + (ev.clientY - startY);
    flexApplyView();
  };
  const onUp = () => {
    document.removeEventListener("pointermove", onMove);
    document.removeEventListener("pointerup", onUp);
    canvas.classList.remove("panning");
  };
  document.addEventListener("pointermove", onMove);
  document.addEventListener("pointerup", onUp);
}
