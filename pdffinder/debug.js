"use strict";

/**
 * debug.js — the black box recorder for the PDF-finder bridge.
 *
 * The bridge has two ends in two different windows, and when it goes quiet
 * there is nothing on screen to say which end went quiet. Neither end can be
 * trusted to report on the other — the very channel that would carry the report
 * is the thing under suspicion — so each end keeps its own log, from its own
 * first line of script, and each can print it out on its own.
 *
 * Everything is recorded always. The cost is a few hundred short strings held
 * in memory; the benefit is that a failure that happened five minutes ago is
 * still in the buffer when someone thinks to look.
 *
 * Loaded as a classic script by both the task pane and the dialog, before
 * anything else of the finder's, so nothing can fail before the recorder is
 * listening.
 */
(function (global) {
  const MAX = 600;                 // lines kept; the oldest fall off the end
  const builds = {};               // which build of each file is actually running
  const t0 = Date.now();
  const lines = [];
  const envs = [];                 // providers of "what this end looks like now"
  let listener = null;

  const clip = (s, n) => {
    const text = typeof s === "string" ? s : JSON.stringify(s);
    if (text === undefined) return "undefined";
    return text.length > n ? text.slice(0, n) + "…(" + text.length + ")" : text;
  };

  /** Anything at all, rendered short enough to read in a list. */
  function detail(d) {
    if (d === undefined || d === null) return "";
    if (typeof d === "string") return clip(d, 220);
    if (d instanceof Error) return d.name + ": " + d.message;
    try { return clip(d, 220); } catch { return String(d); }
  }

  const stamp = (ms) => {
    const s = ms / 1000;
    return (s < 10 ? "  " : s < 100 ? " " : "") + s.toFixed(3) + "s";
  };

  /**
   * Record one thing that happened. `tag` is a short dotted path — `send.chunk`,
   * `office.ready`, `excel.fail` — so a report can be read by scanning the left
   * column.
   */
  function log(tag, d) {
    const line = { at: Date.now() - t0, tag: String(tag), d: detail(d) };
    lines.push(line);
    if (lines.length > MAX) lines.shift();
    if (listener) { try { listener(line); } catch { /* a broken panel must not break the log */ } }
    return line;
  }

  /** Register a function returning `{label: value}` about this end, for the report header. */
  function env(fn) { envs.push(fn); }

  /**
   * Every file of the bridge says which build of itself is running. A browser
   * or a WebView holding one file back while the others move on is invisible
   * from the outside and produces symptoms that make no sense — a codec that
   * drops messages with no complaint, because the copy doing the dropping is
   * the copy without the complaints in it.
   */
  function file(name, build) {
    builds[name] = build;
    log("file." + name, build);
  }

  /** The exact characters of a string, for when a string is not what it seems. */
  function codes(s, n) {
    const text = String(s === undefined || s === null ? "" : s).slice(0, n || 40);
    const out = [];
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      out.push(c < 32 || c > 126 ? "<" + c + ">" : text[i]);
    }
    return out.join("");
  }

  function snapshot() {
    const out = {};
    for (const fn of envs) {
      try { Object.assign(out, fn() || {}); }
      catch (e) { out["(env failed)"] = e.message; }
    }
    return out;
  }

  /** The log as one plain-text block, ready to be pasted into a message. */
  function report() {
    const head = ["=== PDF finder — " + (API.side || "?") + " — " + new Date().toISOString() + " ==="];
    head.push("builds: " + (Object.keys(builds).map((k) => k + " " + builds[k]).join(", ") || "(none registered)"));
    const snap = snapshot();
    for (const k of Object.keys(snap)) head.push(k + ": " + detail(snap[k]));
    head.push("--- " + lines.length + " events ---");
    for (const l of lines) head.push(stamp(l.at) + "  " + l.tag + (l.d ? "  " + l.d : ""));
    return head.join("\n");
  }

  function clear() { lines.length = 0; log("log.cleared"); }

  /** Whatever the panels want to be told when a line lands. One at a time. */
  function onLine(fn) { listener = fn; }

  // Anything that throws where nobody catches it is the most useful line in the
  // log, and the one nobody remembers to write.
  if (global.addEventListener) {
    global.addEventListener("error", (e) =>
      log("window.error", (e.message || "") + " @ " + (e.filename || "") + ":" + (e.lineno || "")));
    global.addEventListener("unhandledrejection", (e) =>
      log("window.reject", (e.reason && (e.reason.stack || e.reason.message)) || String(e.reason)));
  }

  const API = { side: "?", log, env, file, codes, report, clear, onLine, lines, detail, since: () => Date.now() - t0 };
  global.PFDebug = API;
  log("log.started");
})(typeof window !== "undefined" ? window : globalThis);
