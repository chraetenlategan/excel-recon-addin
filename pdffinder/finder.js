// PDF finder — the Excel-side port of the PDF/Excel reconciliation app.
//
// The half of the original that read and rewrote an .xlsx is gone: the workbook
// is open in Excel next door, so the column arrives over the bridge as live
// cells and a tick is a fill colour Excel puts on straight away. What is left
// is the part that matters — the page, the marks, and which printed occurrence
// belongs to which cell.
//
// One row is one cell of the Excel selection and claims at most one printed
// occurrence, so three 50s tick off three separate printed 50s.

import { openPdf, renderPage, ocrPage, needsOcr } from './pdfdoc.js';
import { findAll, toNumber } from './match.js';
import { load, save, drop } from './store.js';
import { sig, claims, ordinal, free, claimNext, claimHit } from './claim.js';
import * as bridge from './bridge.js';

const $ = id => document.getElementById(id);
const el = {
  open: $('btnOpen'), pull: $('btnPull'), clear: $('btnClear'), file: $('filePdf'),
  tHit: $('tHit'), tAll: $('tAll'), exact: $('optExact'), follow: $('optFollow'),
  cells: $('cells'), nameBox: $('nameBox'), fxVal: $('fxVal'), srcName: $('srcName'),
  empty: $('empty'), split: $('split'),
  tick: $('btnTick'), pick: $('tickPick'), any: $('tickAny'),
  viewer: $('viewer'), pages: $('pages'), drop: $('drop'),
  bar: $('bar'), barFill: $('barFill')
};

const S = {
  doc: null,     // {name,size,pages} of the open PDF, or null
  key: null,     // localStorage key for this PDF against this selection
  pages: [],     // {page,width,height,words,ocr,node,marks}
  rows: [],      // {v, ref, mark:null|{p,x,y,w,h}, hand}
  active: -1,
  sheet: '',     // the worksheet the cells came from
  peek: null     // a value outlined on the page without being ticked: {v, ref}
};

/** Ticked off — against the page, or by hand when there is no page to match. */
const done = r => !!(r && (r.mark || r.hand));

/* ---------- occurrences ---------- */
// Every hit of a value, in reading order. Cached; only a rescan changes it and
// that clears the cache.
let hitCache = new Map();
const forget = () => { hitCache = new Map(); };

function hitsFor(v){
  let h = hitCache.get(v);
  if(!h){
    h = findAll(S.pages, v, el.exact.checked).map(x => ({ p: x.page, x: x.x, y: x.y, w: x.w, h: x.h }));
    hitCache.set(v, h);
  }
  return h;
}
const scanned = v => hitCache.has(v);

/** Step a row on to the next occurrence nobody else holds. */
const step = (ri, dir) => S.rows[ri] && claimNext(S.rows, ri, hitsFor(S.rows[ri].v), dir);
/** Give one occurrence to whichever row wants it. @returns the row index, or -1 */
const give = (v, hit) => claimHit(S.rows, v, hit, S.active);

/* ---------- the marker ---------- */
// One colour for every tick, the user's own — the same colour Excel fills the
// cell with, so the two sides read as one document.
const TICK_KEY = 'vdm-pdffinder:tick';
const DEFAULT_TICK = '#FFE94D';

/** A darker cast of the same colour, for the hairline around a tick. */
function edgeOf(hex){
  const n = parseInt(hex.slice(1), 16);
  const dim = v => Math.round(v * 0.62).toString(16).padStart(2, '0');
  return '#' + dim((n >> 16) & 255) + dim((n >> 8) & 255) + dim(n & 255);
}

let tickHex = DEFAULT_TICK;

function setTick(hex, keep){
  const c = /^#[0-9a-f]{6}$/i.test(hex || '') ? hex : DEFAULT_TICK;
  tickHex = c;
  document.documentElement.style.setProperty('--tick', c);
  document.documentElement.style.setProperty('--tick-edge', edgeOf(c));
  el.any.value = c;
  el.pick.querySelectorAll('.sw[data-c]').forEach(b => {
    b.classList.toggle('on', b.dataset.c.toLowerCase() === c.toLowerCase());
  });
  if(keep === false) return;
  save(TICK_KEY, c);
  // Excel repaints in the new colour: the ticks already on the sheet are the
  // same ticks, and an auditor changing pen mid-document means all of them.
  bridge.send({ t: 'colour', hex: c });
  pushTicks(true);
}

/* ---------- progress ---------- */
function progress(p){
  if(p === null){ el.bar.hidden = true; return; }
  el.bar.hidden = false;
  el.barFill.style.width = Math.round(p * 100) + '%';
}

/* ---------- the bridge ---------- */

// Excel is told the whole set of ticked cells rather than each change, so a
// dropped message can never leave the sheet disagreeing with the column.
let pushSoon = 0;
function pushTicks(now){
  clearTimeout(pushSoon);
  const fire = () => bridge.send({
    t: 'ticks',
    sheet: S.sheet,
    hex: tickHex,
    refs: S.rows.filter(done).map(r => r.ref).filter(Boolean),
    all: S.rows.length,
    hit: S.rows.filter(done).length
  });
  if(now) fire(); else pushSoon = setTimeout(fire, 140);
}

/** Put the Excel cursor on a row's own cell. */
function follow(ri){
  const r = S.rows[ri];
  if(r && r.ref && el.follow.checked) bridge.send({ t: 'goto', sheet: S.sheet, ref: r.ref });
}

// The column, straight off the sheet. A tick survives where the same value
// still sits on the same cell, which is what makes re-sending a selection cheap.
bridge.on('rows', msg => {
  const was = new Map(S.rows.map(r => [r.ref, r]));
  S.sheet = msg.sheet || '';
  S.rows = (msg.cells || []).map(c => {
    const v = String(c.v);
    const old = was.get(c.ref);
    const same = old && old.v === v;
    return { v, ref: c.ref, mark: same ? old.mark : null, hand: !!(same && old.hand) };
  });
  S.active = -1;
  S.peek = null;
  el.srcName.textContent = S.sheet || 'Cell';
  el.srcName.title = S.sheet ? 'From ' + S.sheet : 'Where these cells came from';
  el.empty.hidden = S.rows.length > 0;
  el.clear.disabled = !S.rows.length;
  rekey();
  if(!applySaved()) rescan();
  pushTicks(true);
});

bridge.on('colour', msg => setTick(msg.hex, false));

// The pane's own "Clear ticks": the sheet is bare, so the column must be too.
bridge.on('cleared', () => {
  S.rows.forEach(r => { r.mark = null; r.hand = false; });
  redraw();
  if(S.key) drop(S.key);
});

/* ---------- document ---------- */
async function loadFile(file){
  if(!file) return;
  if(file.type ? file.type !== 'application/pdf' : !/\.pdf$/i.test(file.name)) return;
  progress(0.05);
  let pages;
  try{
    ({ pages } = await openPdf(await file.arrayBuffer()));
  }catch(err){
    progress(null);
    alert('That PDF could not be read.\n\n' + err.message);
    return;
  }
  S.pages = pages;
  S.doc = { name: file.name, size: file.size, pages: pages.length };
  S.active = -1;
  S.peek = null;
  forget();
  rekey();

  el.pages.innerHTML = '';
  el.drop.classList.add('gone');
  for(let i = 0; i < pages.length; i++){
    const wrap = document.createElement('div');
    wrap.className = 'page';
    const canvas = document.createElement('canvas');
    const marks = document.createElement('div');
    marks.className = 'marks';
    wrap.append(canvas, marks);
    el.pages.appendChild(wrap);
    pages[i].node = wrap;
    pages[i].marks = marks;
    await renderPage(pages[i], canvas);
    progress(0.05 + 0.9 * (i + 1) / pages.length);
  }
  progress(null);

  el.clear.disabled = !S.rows.length;
  if(!applySaved()) rescan();
  autoOcr();
}

/** This PDF, against this selection — ticks come back when either is reopened. */
function rekey(){
  S.key = S.doc && S.rows.length
    ? 'vdm-pdffinder:' + [S.doc.name, S.doc.size, S.doc.pages, S.sheet, S.rows.length, S.rows[0].ref].join('|')
    : null;
}

/** Scanned pages carry no text layer — read them without being asked. */
async function autoOcr(){
  const token = S.doc && S.doc.name;
  const todo = S.pages.map((p, i) => i).filter(i => needsOcr(S.pages[i]) && !S.pages[i].ocr);
  if(!todo.length) return;
  for(let n = 0; n < todo.length; n++){
    if(!S.doc || S.doc.name !== token) return;     // another document was opened
    try{ await ocrPage(S.pages[todo[n]], p => progress((n + p) / todo.length)); }
    catch{ /* OCR unavailable — the embedded text layer is all there is */ }
    progress((n + 1) / todo.length);
  }
  progress(null);
  if(S.doc && S.doc.name === token){ forget(); rescan(); }
}

/* ---------- the column ---------- */
function paintCells(){
  el.cells.innerHTML = '';
  S.rows.forEach((r, i) => {
    const known = scanned(r.v);
    const hits = known ? hitsFor(r.v) : [];
    const open = known ? free(S.rows, hits) : 0;

    const li = document.createElement('li');
    li.dataset.i = i;
    li.className = [
      i === S.active ? 'active' : '',
      done(r) ? 'done' : '',
      known && !hits.length ? 'miss' : '',
      known && hits.length && !done(r) && !open ? 'short' : ''
    ].filter(Boolean).join(' ');

    li.innerHTML = '<span class="rn"></span><span class="v"></span>' +
      '<span class="n"></span><span class="k"></span>';
    li.querySelector('.rn').textContent = r.ref || (i + 1);

    const v = li.querySelector('.v');
    v.textContent = r.v;
    if(toNumber(r.v) !== null) v.classList.add('num');

    // "2/4" — this row holds the second of four printings; a bare "4" while untouched
    const n = li.querySelector('.n');
    if(!known) n.textContent = '';
    else if(r.mark) n.textContent = ordinal(r, hits) + '/' + hits.length;
    else if(hits.length) n.textContent = String(hits.length);
    if(known && hits.length) n.title = hits.length + ' on the PDF, ' + open + ' still free';

    li.querySelector('.k').title = done(r) ? 'Release this tick' : 'Tick off one occurrence';
    el.cells.appendChild(li);
  });

  const a = S.rows[S.active];
  el.nameBox.textContent = a ? (a.ref || 'A' + (S.active + 1)) : ' ';
  el.fxVal.textContent = a ? a.v : '';
}

/** The reconciliation itself: how much of the column has been ticked off. */
function tally(){
  el.tAll.textContent = S.rows.length;
  el.tHit.textContent = S.rows.filter(done).length;
}

/* ---------- scanning ---------- */
let scanToken = 0;
function rescan(){
  const token = ++scanToken;
  if(!S.pages.length || !S.rows.length){ redraw(); return; }
  const values = [...new Set(S.rows.map(r => r.v))].filter(v => !scanned(v));
  if(!values.length){ redraw(); return; }
  let i = 0;

  const chunk = () => {
    if(token !== scanToken) return;
    const until = Date.now() + 40;
    while(i < values.length && Date.now() < until) hitsFor(values[i++]);
    progress(i / values.length);
    if(i < values.length) setTimeout(chunk, 0);
    else { progress(null); redraw(); }
  };
  chunk();
}

/* ---------- marks ---------- */
function place(node, m, page){
  node.style.left = (m.x / page.width * 100) + '%';
  node.style.top = (m.y / page.height * 100) + '%';
  node.style.width = (m.w / page.width * 100) + '%';
  node.style.height = (m.h / page.height * 100) + '%';
}

/**
 * One box per printed occurrence of every value in the column. The class is the
 * only thing that colours it:
 *   open       — printed here, no row has claimed it (hairline)
 *   candidate  — an open occurrence of the row in hand (navy outline)
 *   claimed    — ticked off by some row (marker fill)
 *   held       — ticked off by the row in hand (marker fill, navy ring)
 */
function drawMarks(){
  S.pages.forEach(p => { if(p.marks) p.marks.innerHTML = ''; });
  if(!S.pages.length) return;
  const taken = claims(S.rows);
  const here = S.rows[S.active];
  const values = new Set(S.rows.map(r => r.v));
  if(S.peek) values.add(S.peek.v);

  for(const v of values){
    if(!scanned(v)) continue;
    const mine = (!!here && here.v === v) || (!!S.peek && S.peek.v === v);
    hitsFor(v).forEach((h, n) => {
      const page = S.pages[h.p];
      if(!page || !page.marks) return;
      const by = taken.get(sig(h));
      const held = by !== undefined;
      const node = document.createElement('div');
      node.className = 'mk ' + (held ? (by === S.active ? 'claimed held' : 'claimed') : (mine ? 'candidate' : 'open'));
      node.dataset.v = v;
      node.dataset.hit = n;
      if(held) node.dataset.ri = by;
      node.title = held ? v + '  —  ' + (S.rows[by].ref || 'row ' + (by + 1)) : v;
      place(node, h, page);
      page.marks.appendChild(node);
    });
  }
}

/** The single path that repaints anything. */
function redraw(){
  drawMarks();
  paintCells();
  tally();
}

/** Bring a box into view; `near` leaves it be when it is comfortably on screen. */
function scrollTo(t, near){
  if(!t) return;
  const page = S.pages[t.p];
  if(!page || !page.node) return;
  const y = page.node.offsetTop + (t.y / page.height) * page.node.clientHeight;
  const h = (t.h / page.height) * page.node.clientHeight;
  const top = el.viewer.scrollTop, bottom = top + el.viewer.clientHeight;
  if(near && y > top + 40 && y + h < bottom - 40) return;
  el.viewer.scrollTo({ top: y - el.viewer.clientHeight * 0.38, behavior: 'smooth' });
}

const reveal = (ri, block) => el.cells.children[ri]?.scrollIntoView({ block: block || 'nearest' });

/** Briefly ring the box a row has just taken. */
function flash(ri){
  const r = S.rows[ri];
  if(!r || !r.mark) return;
  const page = S.pages[r.mark.p];
  const node = page && page.marks && page.marks.querySelector('.mk.held');
  if(node){ node.classList.add('flash'); setTimeout(() => node.classList.remove('flash'), 560); }
}

/* ---------- selection ---------- */
function select(ri, opts){
  const o = opts || {};
  S.peek = null;
  S.active = ri;
  redraw();
  reveal(ri, o.block);
  const r = S.rows[ri];
  if(!r) return;
  follow(ri);
  if(o.silent) return;
  const taken = claims(S.rows);
  const target = r.mark || hitsFor(r.v).find(h => !taken.has(sig(h))) || hitsFor(r.v)[0];
  scrollTo(target, o.near !== false);
}

/** Tick off one occurrence for this row, or step it on to the next one. */
function advance(ri){
  const r = S.rows[ri];
  if(!r) return;
  S.peek = null;
  if(!S.pages.length){                 // no document open — tick the cell by hand
    r.hand = !r.hand;
    S.active = ri;
    redraw();
    settle();
    return;
  }
  if(!step(ri, 1)){ select(ri); return; }
  S.active = ri;
  redraw();
  reveal(ri);
  scrollTo(S.rows[ri].mark, true);
  flash(ri);
  settle();
}

function release(ri){
  if(!done(S.rows[ri])) return;
  S.rows[ri].mark = null;
  S.rows[ri].hand = false;
  S.active = ri;
  redraw();
  reveal(ri);
  settle();
}

/** Everything a change of ticks entails: remember it, and colour the sheet. */
function settle(){
  persist();
  pushTicks();
}

function persist(){
  if(!S.key) return;
  save(S.key, { ticks: S.rows.filter(done).map(r => ({ ref: r.ref, v: r.v, mark: r.mark, hand: !!r.hand })) });
}

/** Put a document's saved ticks back on the cells they were made on. */
function applySaved(){
  if(!S.key) return false;
  const saved = load(S.key);
  if(!saved || !Array.isArray(saved.ticks)) return false;
  const at = new Map(S.rows.map((r, i) => [r.ref, i]));
  for(const t of saved.ticks){
    const ri = at.get(t.ref);
    if(ri === undefined || S.rows[ri].v !== t.v) continue;
    S.rows[ri].mark = t.mark || null;
    S.rows[ri].hand = !!t.hand;
  }
  rescan();
  return true;
}

/* ---------- events ---------- */
el.open.addEventListener('click', () => el.file.click());
el.file.addEventListener('change', e => { loadFile(e.target.files[0]); e.target.value = ''; });
el.pull.addEventListener('click', () => bridge.send({ t: 'pull' }));

el.viewer.addEventListener('dragover', e => { e.preventDefault(); el.drop.classList.remove('gone'); el.drop.classList.add('over'); });
el.viewer.addEventListener('dragleave', () => { el.drop.classList.remove('over'); if(S.pages.length) el.drop.classList.add('gone'); });
el.viewer.addEventListener('drop', e => {
  e.preventDefault();
  el.drop.classList.remove('over');
  if(S.pages.length) el.drop.classList.add('gone');
  loadFile(e.dataTransfer.files[0]);
});

el.exact.addEventListener('change', () => { forget(); rescan(); });

el.tick.addEventListener('click', e => { e.stopPropagation(); el.pick.hidden = !el.pick.hidden; });
el.pick.addEventListener('click', e => {
  const b = e.target.closest('.sw[data-c]');
  if(!b) return;
  setTick(b.dataset.c);
  el.pick.hidden = true;
});
el.any.addEventListener('input', () => setTick(el.any.value));
document.addEventListener('click', e => {
  if(!el.pick.hidden && !e.target.closest('.picker') && e.target !== el.tick) el.pick.hidden = true;
});

el.clear.addEventListener('click', () => {
  S.rows.forEach(r => { r.mark = null; r.hand = false; });
  redraw();
  if(S.key) drop(S.key);
  pushTicks(true);
});

/* the handle between the two windows */
el.split.addEventListener('pointerdown', e => {
  e.preventDefault();
  el.split.setPointerCapture(e.pointerId);
  el.split.classList.add('dragging');
  const move = ev => {
    const w = Math.max(240, Math.min(window.innerWidth - 300, ev.clientX));
    document.documentElement.style.setProperty('--side', w + 'px');
  };
  const up = () => {
    el.split.classList.remove('dragging');
    el.split.removeEventListener('pointermove', move);
    el.split.removeEventListener('pointerup', up);
  };
  el.split.addEventListener('pointermove', move);
  el.split.addEventListener('pointerup', up);
});

/* the column */
el.cells.addEventListener('click', e => {
  const li = e.target.closest('li');
  if(!li) return;
  const ri = +li.dataset.i;
  if(e.target.closest('.k')){
    if(done(S.rows[ri])) release(ri); else advance(ri);
    return;
  }
  select(ri);
});
el.cells.addEventListener('dblclick', e => {
  const li = e.target.closest('li');
  if(!li || e.target.closest('.k')) return;
  e.preventDefault();
  advance(+li.dataset.i);
});
el.cells.addEventListener('contextmenu', e => {
  const li = e.target.closest('li');
  if(!li) return;
  e.preventDefault();
  release(+li.dataset.i);
});

/* the page */
el.pages.addEventListener('click', e => {
  const mk = e.target.closest('.mk');
  if(!mk) return;
  if(mk.dataset.ri !== undefined){ select(+mk.dataset.ri, { near: true }); return; }
  const v = mk.dataset.v;
  const ri = give(v, hitsFor(v)[+mk.dataset.hit]);
  if(ri < 0) return;
  S.active = ri;
  redraw();
  reveal(ri);
  flash(ri);
  follow(ri);
  settle();
});
el.pages.addEventListener('dblclick', e => {
  const mk = e.target.closest('.mk');
  if(!mk) return;
  e.preventDefault();
  // a double-click on the page finds the value on the sheet
  const ri = mk.dataset.ri !== undefined
    ? +mk.dataset.ri
    : give(mk.dataset.v, hitsFor(mk.dataset.v)[+mk.dataset.hit]);
  if(ri < 0) return;
  S.active = ri;
  redraw();
  reveal(ri, 'center');
  follow(ri);
  settle();
});
el.pages.addEventListener('contextmenu', e => {
  const mk = e.target.closest('.mk');
  if(!mk || mk.dataset.ri === undefined) return;
  e.preventDefault();
  release(+mk.dataset.ri);
});

document.addEventListener('keydown', e => {
  if(e.target.tagName === 'TEXTAREA' || e.target.tagName === 'INPUT') return;
  const keys = ['ArrowDown', 'ArrowUp', 'ArrowRight', 'ArrowLeft', 'Enter', ' ', 'Escape', 'Delete', 'Backspace'];
  if(!keys.includes(e.key)) return;
  if(e.key === 'Escape'){ el.pick.hidden = true; S.active = -1; S.peek = null; redraw(); return; }
  if(!S.rows.length) return;
  e.preventDefault();

  if(e.key === 'ArrowDown' || e.key === 'ArrowUp'){
    select(Math.max(0, Math.min(S.rows.length - 1, S.active + (e.key === 'ArrowDown' ? 1 : -1))));
    return;
  }
  if(S.active < 0) return;
  if(e.key === 'Enter' || e.key === ' '){ advance(S.active); return; }
  if(e.key === 'Delete' || e.key === 'Backspace'){ release(S.active); return; }
  // step this row across the other printings of the same value
  if(step(S.active, e.key === 'ArrowRight' ? 1 : -1)){
    redraw();
    scrollTo(S.rows[S.active].mark, true);
    flash(S.active);
    settle();
  }
});

window.addEventListener('beforeunload', persist);

setTick(load(TICK_KEY), false);
redraw();
bridge.start();
