// The finder window's Debug drawer.
//
// The window cannot rely on the pane to report for it: the one channel that
// would carry the report is the thing most likely to be broken. So this end
// prints its own log, on its own, out of a textarea the user can select and
// copy even where the clipboard API is locked down inside a dialog.
//
// It is its own file so `finder.js` stays about the statement and the ticks.

import * as bridge from './bridge.js';

const D = window.PFDebug;
const $ = id => document.getElementById(id);

if (D) {
  D.side = 'finder window';

  const el = {
    btn: $('btnDebug'), box: $('dbg'), log: $('dbgLog'),
    test: $('dbgTest'), copy: $('dbgCopy'), send: $('dbgSend'), clear: $('dbgClear'),
    note: $('dbgNote')
  };

  const paint = () => { if (el.log && el.box && !el.box.hidden) { el.log.value = D.report(); el.log.scrollTop = el.log.scrollHeight; } };
  let soon = 0;
  D.onLine(() => { clearTimeout(soon); soon = setTimeout(paint, 200); });

  const note = (text) => { if (el.note) el.note.textContent = text; };

  if (el.btn) el.btn.addEventListener('click', () => {
    el.box.hidden = !el.box.hidden;
    el.btn.classList.toggle('on', !el.box.hidden);
    paint();
  });

  // The same test the pane runs, from this side: `ping` down the one wire home,
  // answered by the pane with `pong`. No answer means the window→pane leg is
  // dead — which is exactly what a double-click that never gets past "Looking
  // for … in Excel" looks like.
  let n = 0, pending = null;
  if (el.test) el.test.addEventListener('click', () => {
    const id = ++n, at = Date.now();
    clearTimeout(pending);
    D.log('test.ping', '#' + id);
    bridge.send({ t: 'ping', n: id, at, via: 'finder' });
    note('Pinging Excel…');
    pending = setTimeout(() => {
      D.log('test.silence', 'no pong in 5s — this window cannot reach the task pane');
      note('No answer in 5 seconds. This window cannot reach Excel — copy this report.');
    }, 5000);
  });
  bridge.on('pong', msg => {
    clearTimeout(pending);
    const ms = Date.now() - (msg.sentAt || Date.now());
    D.log('test.pong', '#' + msg.n + ' in ' + ms + ' ms');
    note('Excel answered in ' + ms + ' ms — the bridge is working both ways.');
  });

  // The pane asking for this window's log, so one report can hold both ends.
  bridge.on('report', () => bridge.send({ t: 'log', text: D.report() }));

  if (el.copy) el.copy.addEventListener('click', () => {
    const text = D.report();
    if (el.log) { el.log.value = text; el.log.focus(); el.log.select(); }
    const ok = (good) => note(good
      ? 'Copied — paste it into a message.'
      : 'Could not copy automatically; the report is selected above, press Ctrl+C.');
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => ok(true), () => ok(fallback()));
      return;
    }
    ok(fallback());
  });

  if (el.send) el.send.addEventListener('click', () => {
    bridge.send({ t: 'log', text: D.report() });
    note('Sent to the Excel pane — if it arrives, it is in the pane’s report too.');
  });

  if (el.clear) el.clear.addEventListener('click', () => { D.clear(); paint(); });

  function fallback() {
    try { return document.execCommand('copy'); } catch { return false; }
  }

  // Ctrl+Alt+D, for when the button is not where someone is looking.
  document.addEventListener('keydown', e => {
    if (e.ctrlKey && e.altKey && (e.key === 'd' || e.key === 'D') && el.btn) { e.preventDefault(); el.btn.click(); }
  });
}
