// The dialog's end of the channel back to the task pane.
//
// Office gives a dialog exactly one wire home — `messageParent` — and one the
// other way, `DialogParentMessageReceived`. Everything the finder knows about
// the workbook arrives here, and every tick it makes leaves here. Nothing else
// in the dialog touches Office.js.
//
// The finder runs perfectly well with no pane on the other end (opened straight
// in a browser, say): `ready` simply never resolves anything and the column
// stays empty until one is sent.

const handlers = new Map();
let post = null;                       // set once Office is up
const outbox = [];                     // messages made before that happened

/** Register a handler for one message type from the pane. */
export function on(type, fn) {
  handlers.set(type, fn);
}

/** Send one message to the pane. Safe to call before Office is ready. */
export function send(msg) {
  if (!post) { outbox.push(msg); return; }
  for (const chunk of window.PFWire.encode(msg)) post(chunk);
}

/**
 * Hand one decoded message to its handler. Office's own inbound path ends here,
 * and it is exported so the finder can be driven — and looked at — in an
 * ordinary browser tab with no Excel behind it.
 */
export function receive(msg) {
  const fn = handlers.get(msg && msg.t);
  if (fn) fn(msg);
}

export function start() {
  if (typeof Office === "undefined" || !Office.onReady) return;   // opened outside Office
  Office.onReady(() => {
    const ui = Office.context && Office.context.ui;
    if (!ui || !ui.messageParent) return;

    const read = window.PFWire.reader(receive);
    if (ui.addHandlerAsync) {
      ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => read(arg.message));
    }

    post = (chunk) => {
      // targetOrigin is ignored on desktop and required on the web; the pane
      // and the dialog are served from the same origin either way.
      try { ui.messageParent(chunk, { targetOrigin: window.location.origin }); }
      catch { ui.messageParent(chunk); }
    };

    send({ t: "ready" });
    while (outbox.length) send(outbox.shift());
  });
}
