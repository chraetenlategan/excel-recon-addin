"use strict";

/**
 * wire.js — the codec both ends of the PDF-finder bridge speak.
 *
 * Office's dialog channel carries one string at a time and caps how long that
 * string may be, so every message goes out as `id \t i \t n \t chunk` and is
 * put back together on the other side. Plain tabs rather than a JSON envelope:
 * the payload is already JSON and does not want escaping twice.
 *
 * Loaded as a classic script by both the task pane and the dialog, so the two
 * ends can never drift apart.
 */
(function (global) {
  const BUILD = "2026-09-01c";

  // The recorder is loaded before this file, but never assume it: a codec that
  // throws while logging is worse than a codec that does not log, and that is
  // exactly the way this file was broken once.
  const say = (tag, d) => { try { if (global.PFDebug) global.PFDebug.log(tag, d); } catch { /* never break the wire */ } };
  const codes = (s, n) => { try { return global.PFDebug ? global.PFDebug.codes(s, n) : String(s).slice(0, n); } catch { return "?"; } };
  if (global.PFDebug) global.PFDebug.file("wire", BUILD);

  // Office documents no hard limit for messageParent, but hosts have been
  // unreliable well below 32 KB. 16 000 characters is comfortably under every
  // one of them and costs a handful of extra round trips on a long column.
  const LIMIT = 16000;

  let seq = 0;

  /** @returns {string[]} the chunks to send, in order */
  function encode(obj) {
    const body = JSON.stringify(obj);
    const id = (++seq) + "." + Math.random().toString(36).slice(2, 8);
    const n = Math.max(1, Math.ceil(body.length / LIMIT));
    const out = [];
    for (let i = 0; i < n; i++) {
      out.push(id + "\t" + i + "\t" + n + "\t" + body.substr(i * LIMIT, LIMIT));
    }
    say("wire.encode", (obj && obj.t) + " id=" + id + " " + body.length + "ch in " + n);
    return out;
  }

  /**
   * A reassembler. Feed it every raw string that arrives; it calls `onMessage`
   * once per complete message.
   */
  function reader(onMessage) {
    const pending = new Map();
    return function (raw) {
      if (typeof raw !== "string") {
        // The commonest way to break this bridge: handing the reader the event
        // object Office passes a handler instead of the string on `.message`.
        say("wire.notString", raw && typeof raw === "object"
          ? "got an object with keys [" + Object.keys(raw).join(",") + "]"
          : typeof raw);
        return;
      }
      const a = raw.indexOf("\t"), b = raw.indexOf("\t", a + 1), c = raw.indexOf("\t", b + 1);
      if (a < 0 || b < 0 || c < 0) {
        // The separators did not survive the trip. What arrived instead is the
        // whole answer, so it is printed character by character.
        say("wire.badFrame", "tabs at " + a + "/" + b + "/" + c + " — head: " + codes(raw, 60));
        return;
      }
      const id = raw.slice(0, a);
      const i = parseInt(raw.slice(a + 1, b), 10);
      const n = parseInt(raw.slice(b + 1, c), 10);
      const part = raw.slice(c + 1);
      if (!isFinite(i) || !isFinite(n) || n < 1) { say("wire.badCount", raw.slice(0, 60)); return; }

      let bin = pending.get(id);
      if (!bin) { bin = { n, got: 0, parts: new Array(n) }; pending.set(id, bin); }
      if (bin.parts[i] === undefined) { bin.parts[i] = part; bin.got++; }
      if (bin.got < bin.n) { say("wire.part", id + " " + bin.got + "/" + bin.n); return; }

      pending.delete(id);
      let msg;
      try { msg = JSON.parse(bin.parts.join("")); }
      catch (e) {
        const body = bin.parts.join("");
        say("wire.badJson", id + " — " + e.message +
          " — " + body.length + "ch, head: " + codes(body, 50) + " tail: " + codes(body.slice(-25), 25));
        return;
      }
      say("wire.decode", (msg && msg.t) + " id=" + id + " of " + n);
      // The handler's own faults are its own; they must never be mistaken for
      // the codec dropping a message.
      try { onMessage(msg); }
      catch (e) { say("wire.handlerThrew", (msg && msg.t) + " — " + (e && (e.stack || e.message))); }
    };
  }

  global.PFWire = { encode, reader, LIMIT, BUILD };
})(typeof window !== "undefined" ? window : globalThis);
