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
//
// Every step of the handshake is written to `PFDebug`, because when this file
// goes quiet there is nothing on either screen to say so. `ping` is answered
// here rather than in the finder, and answered three different ways, so a
// bridge that is half working says exactly which half.

const handlers = new Map();
let post = null;                       // set once Office is up
const outbox = [];                     // messages made before that happened
let parentOrigin = "";
let ui = null;
let inbound = 0, outbound = 0;

const BUILD = "2026-09-01c";
const say = (tag, d) => { if (window.PFDebug) window.PFDebug.log(tag, d); };
const codes = (s, n) => (window.PFDebug ? window.PFDebug.codes(s, n) : String(s).slice(0, n));
if (window.PFDebug) window.PFDebug.file("bridge", BUILD);

/**
 * A string sent home with no codec around it at all — no chunk header, no JSON,
 * nothing that could be blamed on this end's own encoding. What comes back from
 * a round trip of one of these is what the Office channel really does to a
 * string, and that is the one fact the whole bridge rests on.
 */
const RAW_PROBE = "PFRAW|";
const RAW_BACK = "PFRAWBACK|";
// tab, then a space, a non-breaking space, an em dash and an accent — the
// characters most likely to be eaten, and the tab is the one that matters.
const RAW_SAMPLE = "a\tb c\u00a0d—é|end";

/** Register a handler for one message type from the pane. */
export function on(type, fn) {
  handlers.set(type, fn);
}

/** Send one message to the pane. Safe to call before Office is ready. */
export function send(msg) {
  if (!post) {
    outbox.push(msg);
    say("send.queued", (msg && msg.t) + " — Office is not up yet, " + outbox.length + " waiting");
    return;
  }
  for (const chunk of window.PFWire.encode(msg)) post(chunk);
}

/**
 * Hand one decoded message to its handler. Office's own inbound path ends here,
 * and it is exported so the finder can be driven — and looked at — in an
 * ordinary browser tab with no Excel behind it.
 */
/** Answer a raw probe with what actually arrived, uncoded, so nothing can hide it. */
function rawEcho(text) {
  const body = text.slice(RAW_PROBE.length);
  const intact = body === RAW_SAMPLE;
  say("probe.in", (intact ? "INTACT" : "MANGLED") + " " + body.length + "ch: " + codes(body, 60));
  const reply = RAW_BACK + (intact ? "intact" : "mangled") + "|" + codes(body, 60) + "|" + body;
  for (const via of ["targetOrigin", "bare", "star"]) rawPost(reply, via);
}

export function receive(msg) {
  inbound++;
  const t = msg && msg.t;
  if (t === "ping") { pong(msg); return; }
  const fn = handlers.get(t);
  say("in." + t, fn ? "handled" : "NO HANDLER for this type");
  if (fn) {
    try { fn(msg); }
    catch (e) { say("in.threw", t + " — " + (e && e.stack || e)); }
  }
}

/**
 * Answer the pane's bridge test — down all three forms of `messageParent` at
 * once, each labelled. Whichever labels come out the other end are the forms
 * this host actually delivers; if none do, the return leg is dead and the
 * finder's own log will still show the ping arriving.
 */
function pong(msg) {
  say("in.ping", "#" + (msg && msg.n));
  const body = { t: "pong", n: msg && msg.n, sentAt: msg && msg.at, at: Date.now() };
  for (const via of ["targetOrigin", "bare", "star"]) {
    for (const chunk of window.PFWire.encode({ ...body, via })) rawPost(chunk, via);
  }
}

/** One chunk out, by one named form of messageParent. @returns whether it threw */
function rawPost(chunk, via) {
  if (!ui || !ui.messageParent) { say("out.noUi", via); return false; }
  try {
    if (via === "bare") ui.messageParent(chunk);
    else if (via === "star") ui.messageParent(chunk, { targetOrigin: "*" });
    else ui.messageParent(chunk, { targetOrigin: parentOrigin });
    outbound++;
    say("out.chunk", via + " " + chunk.length + "ch");
    return true;
  } catch (e) {
    say("out.threw", via + " — " + (e && e.message));
    return false;
  }
}

export function probe() {
  const req = (name, v) => {
    try { return Office.context.requirements.isSetSupported(name, v); }
    catch (e) { return "?" + e.message; }
  };
  const has = typeof Office !== "undefined" && Office.context;
  return {
    "window.url": window.location.href,
    "window.origin": window.location.origin,
    "parent origin": parentOrigin || "(not set)",
    "Office loaded": typeof Office !== "undefined",
    "Office.context": !!has,
    "Office host": has ? String(Office.context.host) : "-",
    "Office platform": has ? String(Office.context.platform) : "-",
    "DialogApi 1.1": has ? req("DialogApi", "1.1") : "-",
    "DialogApi 1.2": has ? req("DialogApi", "1.2") : "-",
    "messageParent": !!(ui && ui.messageParent),
    "addHandlerAsync": !!(ui && ui.addHandlerAsync),
    "channel open": !!post,
    "queued unsent": outbox.length,
    "chunks in / out": inbound + " msgs in, " + outbound + " chunks out"
  };
}

export function start() {
  say("bridge.start");
  if (window.PFDebug) window.PFDebug.env(probe);

  if (typeof Office === "undefined" || !Office.onReady) {
    say("bridge.noOffice", "office.js did not load, or this window is not a dialog — running standalone");
    return;
  }

  // Office.onReady that never resolves is indistinguishable from one that
  // resolves late, unless somebody is counting.
  let ready = false;
  setTimeout(() => { if (!ready) say("office.stillWaiting", "Office.onReady has not fired after 5s"); }, 5000);

  Office.onReady((info) => {
    ready = true;
    say("office.ready", info ? (info.host + " / " + info.platform) : "no info");
    ui = Office.context && Office.context.ui;
    if (!ui || !ui.messageParent) {
      say("office.noUi", "Office.context.ui.messageParent is missing — nothing can be sent home");
      return;
    }

    const read = window.PFWire.reader(receive);
    if (ui.addHandlerAsync) {
      ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        (arg) => {
          const raw = arg && arg.message;
          if (typeof raw !== "string") {
            say("in.raw", "no .message — keys [" + (arg ? Object.keys(arg).join(",") : "null") + "]");
            return;
          }
          say("in.raw", raw.length + "ch — head: " + codes(raw, 44));
          if (raw.slice(0, RAW_PROBE.length) === RAW_PROBE) { rawEcho(raw); return; }
          read(raw);
        },
        // Without this callback a refused registration is silent, and a finder
        // that hears nothing looks exactly like a pane that says nothing.
        (res) => say("in.handler", res && res.status === "failed"
          ? "REGISTRATION FAILED — " + (res.error && res.error.message)
          : "registered")
      );
    } else {
      say("in.noHandler", "ui.addHandlerAsync is missing — this dialog cannot be sent anything");
    }

    // The finder may be served from somewhere other than the pane, and a
    // cross-origin reply has to name who it is for. The pane puts its own
    // origin on the query string; failing that the two share one.
    parentOrigin = window.location.origin;
    try { parentOrigin = new URLSearchParams(window.location.search).get("parent") || parentOrigin; }
    catch { say("bridge.noQuery"); }
    say("bridge.parent", parentOrigin + (parentOrigin === window.location.origin ? " (same origin)" : " (CROSS origin)"));

    post = (chunk) => {
      // targetOrigin is ignored on desktop and required on the web. A host that
      // dislikes the options object throws; one that dislikes the origin inside
      // it often does not, so the bare form is tried on a throw only.
      if (!rawPost(chunk, "targetOrigin")) rawPost(chunk, "bare");
    };

    send({ t: "ready", build: BUILD, wire: window.PFWire && window.PFWire.BUILD });
    say("bridge.open", "flushing " + outbox.length + " queued");
    while (outbox.length) send(outbox.shift());
  });
}
