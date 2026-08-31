// Which printed occurrence belongs to which line of the pasted column.
// Pure functions, no DOM: a row is {v, mark}, an occurrence is {p,x,y,w,h}.
//
// The rule is one occurrence per row. Three 50s in the column tick off three
// separate printed 50s, and the fourth printed 50 stays open.

/** An occurrence's identity — page and position, rounded to survive a rescan. */
export const sig = m => [m.p, Math.round(m.x), Math.round(m.y)].join(',');

/** sig -> index of the row holding it. */
export function claims(rows){
  const m = new Map();
  rows.forEach((r, i) => { if(r.mark) m.set(sig(r.mark), i); });
  return m;
}

/** Which occurrence (1-based) of its value a row holds, or 0. */
export function ordinal(row, hits){
  if(!row || !row.mark) return 0;
  return hits.findIndex(h => sig(h) === sig(row.mark)) + 1;
}

/** How many occurrences of this value nobody holds. */
export function free(rows, hits){
  const taken = claims(rows);
  return hits.filter(h => !taken.has(sig(h))).length;
}

/**
 * Move a row onto the next occurrence nobody else holds, wrapping around.
 * Called again on the same row it steps to the one after — that is how one
 * duplicate walks through the four printed 50s.
 */
export function claimNext(rows, ri, hits, dir = 1){
  const r = rows[ri];
  if(!r || !hits.length) return false;
  const taken = claims(rows);
  const from = r.mark ? hits.findIndex(h => sig(h) === sig(r.mark)) : (dir > 0 ? -1 : 0);
  for(let k = 1; k <= hits.length; k++){
    const at = ((from + dir * k) % hits.length + hits.length) % hits.length;
    const by = taken.get(sig(hits[at]));
    if(by === undefined || by === ri){ r.mark = hits[at]; return true; }
  }
  return false;   // every printing of this value is already spoken for
}

/**
 * Hand one particular occurrence to the row that wants it: the row in hand when
 * it is still empty, else the first empty row carrying that value.
 * @returns the row index, or -1 when the column does not carry the value
 */
export function claimHit(rows, v, hit, prefer){
  if(!hit) return -1;
  const held = claims(rows).get(sig(hit));
  if(held !== undefined) return held;
  const p = rows[prefer];
  if(p && p.v === v && !p.mark){ p.mark = hit; return prefer; }
  let ri = rows.findIndex(r => r.v === v && !r.mark);
  if(ri < 0) ri = rows.findIndex(r => r.v === v);
  if(ri < 0) return -1;
  rows[ri].mark = hit;
  return ri;
}
