// PDF finder — the page half of the add-in's PDF tab.
//
// The workbook is open in Excel next door, so this window shows one thing: the
// statement. There is no column of cells on it and no list to keep in step —
// the cells live in the sheet, where the auditor can already see them. What
// arrives over the bridge is the scope the pane was told to send (a sheet, and
// the columns the user narrowed it to); every value in it is outlined where it
// is printed, and ticking one off fills its cell in Excel straight away.
//
// One cell claims at most one printed occurrence, so three 50s in the scope
// tick off three separate printed 50s and a fourth printing stays open.

import { openPdf, renderPage, ocrPage, needsOcr } from './pdfdoc.js';
import { findAll, toNumber, toText, valueAt } from './match.js';
import { load, save, drop } from './store.js';
import { sig, claims, claimNext, claimHit } from './claim.js';
import * as bridge from './bridge.js';

const $ = id => document.getElementById(id);
const el = {
  open: $('btnOpen'), pull: $('btnPull'), clear: $('btnClear'), file: $('filePdf'),
  tHit: $('tHit'), tAll: $('tAll'), exact: $('optExact'), follow: $('optFollow'),
  scope: $('scope'), tick: $('btnTick'), pick: $('tickPick'), any: $('tickAny'),
  viewer: $('viewer'), pages: $('pages'), drop: $('drop'), toast: $('toast'),
  bar: $('bar'), barFill: $('barFill')
};

const S = {
  doc: null,     // {name,size,pages} of the open PDF, or null
  key: null,     // localStorage key for this PDF against this scope
  pages: [],     // {page,width,height,words,ocr,node,marks}
  rows: [],      // one per cell in scope: {v, ref, mark:null|{p,x,y,w,h}}
  active: -1,    // the cell whose value is outlined in navy
  sheet: '',     // the worksheet the cells came from
  scope: '',     // how the pane described that scope, for the header
  peek: null     // a value outlined without being ticked: {v}
};

const done = r => !!(r && r.mark);

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

/** Step a cell on to the next occurrence nobody else holds. */
const step = (ri, dir) => S.rows[ri] && claimNext(S.rows, ri, hitsFor(S.rows[ri].v), dir);
/** Give one occurrence to whichever cell wants it. @returns the row index, or -1 */
const give = (v, hit) => claimHit(S.rows, v, hit, S.active);

/* ---------- the marker ---------- */
// One colour for every tick, the user's own — the same colour Excel fills the
// cell with, so the two windows read as one document.
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
// dropped message can never leave the sheet disagreeing with the page.
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

/** Put the Excel cursor on a cell. */
function follow(ri){
  const r = S.rows[ri];
  if(r && r.ref && el.follow.checked) bridge.send({ t: 'goto', sheet: S.sheet, ref: r.ref });
}

// The cells in scope, straight off the sheet. A tick survives where the same
// value still sits on the same cell, which is what makes re-reading cheap.
bridge.on('rows', msg => {
  const was = new Map(S.rows.map(r => [r.ref, r]));
  S.sheet = msg.sheet || '';
  S.scope = msg.scope || '';
  S.rows = (msg.cells || []).map(c => {
    const v = String(c.v);
    const old = was.get(c.ref);
    return { v, ref: c.ref, mark: old && old.v === v ? old.mark : null };
  });
  S.active = -1;
  S.peek = null;
  el.scope.textContent = S.scope || 'the cells from Excel, found on the page';
  el.clear.disabled = !S.rows.length;
  rekey();
  if(!applySaved()) rescan();
  pushTicks(true);
});

bridge.on('colour', msg => setTick(msg.hex, false));

/* ---------- a cell picked in Excel, found on the page ---------- */

/** The mark node for one occurrence, so it can be rung without redrawing. */
function markNode(v, n){
  return [...el.pages.querySelectorAll('.mk')].find(m => m.dataset.v === v && +m.dataset.hit === n);
}

/** Ring one occurrence briefly, whether or not any cell has claimed it. */
function flashHit(v, n){
  const node = markNode(v, n);
  if(node){ node.classList.add('flash'); setTimeout(() => node.classList.remove('flash'), 560); }
}

/** The first occurrence nobody has ticked off, or the first one at all. */
function freeHit(hits){
  const taken = claims(S.rows);
  const n = hits.findIndex(h => taken.get(sig(h)) === undefined);
  return n < 0 ? 0 : n;
}

/**
 * The other direction: Excel says which cell the cursor is on, and the page
 * shows where that value is printed.
 *
 * A cell of the scope is taken in hand, exactly as clicking its mark would —
 * so the arrow keys step it on to its next printing, and its own tick shows
 * navy. A cell outside the scope is outlined without being adopted: it has no
 * row to belong to, and inventing one would put a value in the count that the
 * user never asked to reconcile.
 *
 * Nothing is ever ticked here. Excel moved a cursor; it did not reconcile.
 */
bridge.on('look', msg => {
  const ref = msg.ref || '';
  const v = String(msg.v === undefined || msg.v === null ? '' : msg.v);
  const where = (msg.sheet ? msg.sheet + '!' : '') + ref;

  if(msg.blank || !v.trim()){ toast(where + ' is blank.'); return; }
  if(!S.pages.length){ toast('Open a PDF and ' + where + ' can be found on it.'); return; }

  const mine = (!msg.sheet || msg.sheet === S.sheet) ? S.rows.findIndex(r => r.ref === ref) : -1;
  const hits = hitsFor(v);

  if(mine > -1){ S.active = mine; S.peek = null; }
  else { S.active = -1; S.peek = { v }; }
  redraw();

  if(!hits.length){
    toast(where + '  —  ' + v + ' is not printed on this document.');
    return;
  }

  // its own tick if it has one, otherwise the first printing still going spare
  const n = (mine > -1 && S.rows[mine].mark)
    ? Math.max(0, hits.findIndex(h => sig(h) === sig(S.rows[mine].mark)))
    : freeHit(hits);
  scrollTo(hits[n], false);
  flashHit(v, n);
  toast(where + '  →  ' + v + (hits.length > 1 ? '  (' + hits.length + ' on the page)' : ''));
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

/** This PDF, against this scope — ticks come back when either is reopened. */
function rekey(){
  S.key = S.doc && S.rows.length
    ? 'vdm-pdffinder:' + [S.doc.name, S.doc.size, S.doc.pages, S.sheet, S.scope, S.rows.length, S.rows[0].ref].join('|')
    : null;
}

/**
 * Scanned pages carry no text layer, so they are read for the user without
 * being asked: nobody has to know what OCR is, or press anything to get it.
 */
async function autoOcr(){
  const token = S.doc && S.doc.name;
  const todo = S.pages.map((p, i) => i).filter(i => needsOcr(S.pages[i]) && !S.pages[i].ocr);
  if(!todo.length) return;
  toast(todo.length === S.pages.length ? 'Reading the scan…' : 'Reading ' + todo.length + ' scanned pages…');
  for(let n = 0; n < todo.length; n++){
    if(!S.doc || S.doc.name !== token) return;     // another document was opened
    try{ await ocrPage(S.pages[todo[n]], p => progress((n + p) / todo.length)); }
    catch{ /* OCR unavailable — the embedded text layer is all there is */ }
    progress((n + 1) / todo.length);
  }
  progress(null);
  if(S.doc && S.doc.name === token){ forget(); rescan(); }
}

/* ---------- the count ---------- */
/** The reconciliation itself: how much of the scope has been ticked off. */
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
 * One box per printed occurrence of every value in scope. The class is the only
 * thing that colours it:
 *   open       — printed here, no cell has claimed it (hairline)
 *   candidate  — an open occurrence of the value in hand (navy outline)
 *   claimed    — ticked off by some cell (marker fill)
 *   held       — ticked off by the cell in hand (marker fill, navy ring)
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
      node.title = held ? v + '  —  ' + (S.rows[by].ref || '') : v;
      place(node, h, page);
      page.marks.appendChild(node);
    });
  }
}

/** The single path that repaints anything. */
function redraw(){
  drawMarks();
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

/** Briefly ring the box a cell has just taken. */
function flash(){
  const node = el.pages.querySelector('.mk.held');
  if(node){ node.classList.add('flash'); setTimeout(() => node.classList.remove('flash'), 560); }
}

/* ---------- ticking ---------- */

/** Tick one printed occurrence off against the cell it belongs to. */
function take(v, hit){
  const ri = give(v, hit);
  if(ri < 0) return;                       // no cell in scope carries this value
  S.active = ri;
  S.peek = null;
  redraw();
  flash();
  follow(ri);
  settle();
  toast(v + '  →  ' + where(ri));
}

function release(ri){
  if(!done(S.rows[ri])) return;
  S.rows[ri].mark = null;
  S.active = ri;
  redraw();
  settle();
}

/** How a cell reads in a message: Sheet1!B12. */
const where = ri => (S.sheet ? S.sheet + '!' : '') + (S.rows[ri] ? S.rows[ri].ref || '' : '');

/** Everything a change of ticks entails: remember it, and colour the sheet. */
function settle(){
  persist();
  pushTicks();
}

function persist(){
  if(!S.key) return;
  save(S.key, { ticks: S.rows.filter(done).map(r => ({ ref: r.ref, v: r.v, mark: r.mark })) });
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
  }
  rescan();
  return true;
}

/* ---------- looking a printed value up in Excel ---------- */

// A line of feedback over the page: where a value landed on the sheet, or that
// the workbook does not carry it. It says itself and goes away.
let toastSoon = 0;
function toast(text){
  if(!el.toast) return;
  el.toast.textContent = text;
  el.toast.hidden = false;
  clearTimeout(toastSoon);
  toastSoon = setTimeout(() => { el.toast.hidden = true; }, 3200);
}

// Excel is asked things and answers them; when it stops answering, the page
// should say so rather than leave "Looking for …" standing for ever. Nothing
// but a silent bridge produces that, so the message points straight at it.
let waiting = 0;
function expectAnswer(what){
  clearTimeout(waiting);
  waiting = setTimeout(() => {
    if(window.PFDebug) window.PFDebug.log('bridge.silent', what + ' — Excel did not answer in 6s');
    toast('Excel has not answered — close the finder and open it again.');
  }, 6000);
}

bridge.on('found', msg => { clearTimeout(waiting); toast(msg.msg || ''); });

/** The printed value under a click, whatever the page prints there. */
function valueUnder(e){
  const node = e.target.closest('.page');
  if(!node) return null;
  const p = S.pages.findIndex(pg => pg.node === node);
  const pg = S.pages[p];
  if(!pg || !pg.words.length) return null;
  const box = node.getBoundingClientRect();
  const found = valueAt(pg.words,
    (e.clientX - box.left) / box.width * pg.width,
    (e.clientY - box.top) / box.height * pg.height);
  return found && found.t.trim() ? { ...found, p } : null;
}

/** Two printed strings, or a printed string and a cell, saying the same thing. */
function sameValue(a, b){
  const na = toNumber(a), nb = toNumber(b);
  if(na !== null || nb !== null) return na !== null && nb !== null && Math.abs(na - nb) < 0.005;
  const ta = toText(a);
  return ta.length > 0 && ta === toText(b);
}

/**
 * Take a value off the page and put the Excel cursor on it.
 *
 * A value the scope already carries is one of these cells, so its own cell is
 * the answer. Anything else was never in scope at all — the pane is asked to
 * hunt it down in the workbook, which is how a figure printed on the statement
 * but outside the chosen columns can still be found on the sheet.
 */
function lookUp(v){
  const ri = S.rows.findIndex(r => sameValue(r.v, v));
  if(ri > -1 && S.rows[ri].ref){
    S.active = ri;
    S.peek = null;
    redraw();
    bridge.send({ t: 'goto', sheet: S.sheet, ref: S.rows[ri].ref });
    toast(v + '  →  ' + where(ri));
    return;
  }
  // outline every printing of it while Excel is looked through
  S.peek = { v };
  hitsFor(v);
  redraw();
  bridge.send({ t: 'find', sheet: S.sheet, v });
  expectAnswer('find ' + v);
  toast('Looking for ' + v + ' in Excel…');
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
  S.rows.forEach(r => { r.mark = null; });
  S.active = -1;
  redraw();
  if(S.key) drop(S.key);
  pushTicks(true);
});

/* the page — the whole of the interface */
el.pages.addEventListener('click', e => {
  const mk = e.target.closest('.mk');
  if(!mk) return;
  if(mk.dataset.ri !== undefined){          // already ticked: take that cell in hand
    S.active = +mk.dataset.ri;
    S.peek = null;
    redraw();
    follow(S.active);
    return;
  }
  take(mk.dataset.v, hitsFor(mk.dataset.v)[+mk.dataset.hit]);
});

el.pages.addEventListener('dblclick', e => {
  const mk = e.target.closest('.mk');
  // whatever is printed under the pointer, found on the sheet — a marked value
  // by the cell that holds it, anything else by asking the pane to look
  const v = mk ? mk.dataset.v : (valueUnder(e) || {}).t;
  if(!v) return;
  e.preventDefault();
  lookUp(v);
});

el.pages.addEventListener('contextmenu', e => {
  const mk = e.target.closest('.mk');
  if(!mk || mk.dataset.ri === undefined) return;
  e.preventDefault();
  release(+mk.dataset.ri);
});

document.addEventListener('keydown', e => {
  if(e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
  if(e.key === 'Escape'){ el.pick.hidden = true; S.active = -1; S.peek = null; redraw(); return; }
  if(S.active < 0) return;
  if(e.key === 'Delete' || e.key === 'Backspace'){ e.preventDefault(); release(S.active); return; }
  // step the cell in hand across the other printings of its value
  if(e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
  e.preventDefault();
  if(step(S.active, e.key === 'ArrowRight' ? 1 : -1)){
    redraw();
    scrollTo(S.rows[S.active].mark, true);
    flash();
    settle();
  }
});

window.addEventListener('beforeunload', persist);

setTick(load(TICK_KEY), false);
redraw();
bridge.start();
