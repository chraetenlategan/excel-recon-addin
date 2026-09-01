// Value normalisation + word-sequence matching.
// Words carry their own boxes, so matches never straddle a token boundary
// (searching 234.56 can never hit inside 1234.56).

const CURRENCY = /[R$€£¥]/gi;

/** Parse a printed amount into a number, or null. Handles 1 234,56 / 1,234.56 / (1.234,56) / 1234.56- */
export function toNumber(raw){
  if(raw == null) return null;
  let s = String(raw).trim().replace(CURRENCY, '').replace(/\s| |'/g, '');
  if(!s) return null;

  let neg = false;
  if(/^\(.*\)$/.test(s)){ neg = true; s = s.slice(1, -1); }
  if(s.startsWith('-')){ neg = !neg; s = s.slice(1); }
  if(s.endsWith('-')){ neg = !neg; s = s.slice(0, -1); }
  if(s.endsWith('%')) return null;
  if(!/^[\d.,]+$/.test(s) || !/\d/.test(s)) return null;

  const last = Math.max(s.lastIndexOf('.'), s.lastIndexOf(','));
  let intPart = s, decPart = '';
  if(last > -1){
    const tail = s.slice(last + 1);
    // a final separator with 1-2 trailing digits is a decimal point; anything else groups thousands
    if(/^\d{1,2}$/.test(tail)){ intPart = s.slice(0, last); decPart = tail; }
  }
  intPart = intPart.replace(/[.,]/g, '');
  if(!/^\d*$/.test(intPart) || !/^\d*$/.test(decPart)) return null;
  if(intPart === '' && decPart === '') return null;

  const n = Number((intPart || '0') + (decPart ? '.' + decPart : ''));
  if(!isFinite(n)) return null;
  return neg ? -n : n;
}

/** Letters/digits only, lowercased — for text comparison. */
export function toText(raw){
  return String(raw == null ? '' : raw).toLowerCase().replace(/[^a-z0-9]/g, '');
}

const near = (a, b) => Math.abs(a - b) < 0.005;

function sameLine(a, b){
  const overlap = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return overlap > Math.min(a.h, b.h) * 0.5;
}

function adjacent(a, b){
  return sameLine(a, b) && b.x >= a.x - 1 && b.x - (a.x + a.w) < Math.max(a.h, b.h) * 1.5;
}

const FRAGMENT = /^[\d.,' ]+$|^%$/;

/** Reject a numeric hit that is really part of a longer printed figure. */
function bleeds(words, i, j){
  const before = words[i - 1], after = words[j + 1];
  if(before && adjacent(before, words[i]) && FRAGMENT.test(before.t)) return true;
  if(after && adjacent(words[j], after) && FRAGMENT.test(after.t)) return true;
  return false;
}

function union(words, i, j){
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for(let k = i; k <= j; k++){
    const w = words[k];
    x0 = Math.min(x0, w.x); y0 = Math.min(y0, w.y);
    x1 = Math.max(x1, w.x + w.w); y1 = Math.max(y1, w.y + w.h);
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

/**
 * Find every occurrence of `needle` in one page's words.
 * @returns [{i,j,x,y,w,h}] boxes in page-point coordinates
 */
export function findInPage(words, needle, exact){
  const hits = [];
  const num = toNumber(needle);
  const txt = toText(needle);
  if(num === null && txt.length < 2) return hits;

  for(let i = 0; i < words.length; i++){
    let joined = '';
    let end = -1;
    const span = num !== null ? 3 : 8;

    for(let j = i; j < words.length && j < i + span; j++){
      if(j > i && !adjacent(words[j - 1], words[j])) break;
      joined += words[j].t;

      if(num !== null){
        const v = toNumber(joined);
        if(v !== null && (near(v, num) || near(Math.abs(v), Math.abs(num)))) end = j;
      }else{
        const c = toText(joined);
        if(!c.length) continue;
        if(exact){
          if(c === txt){ end = j; break; }
          if(!txt.startsWith(c)) break;          // this run can never reach the needle
        }else{
          if(c.includes(txt)){ end = j; break; }
          if(c.length > txt.length * 2 + 12) break;
        }
      }
    }

    if(end > -1 && num !== null && bleeds(words, i, end)) end = -1;

    // a "contains" hit can start later than the run did — keep the tightest span
    while(end > i && num === null && !exact){
      const shorter = words.slice(i + 1, end + 1).map(w => toText(w.t)).join('');
      if(!shorter.includes(txt)) break;
      i++;
    }

    if(end > -1){
      hits.push({ i, j: end, ...union(words, i, end) });
      i = end;
    }
  }
  return hits;
}

/**
 * The printed value under a point on the page, in page-point coordinates.
 *
 * A word on its own, except for figures: an amount printed as `1 234,56` is
 * three words on the page and one value on the statement, so the numeric
 * fragments printed alongside the one under the pointer come with it.
 *
 * @returns {null|{t,i,j,x,y,w,h}} the value and its box, or null off any word
 */
export function valueAt(words, x, y){
  let at = -1;
  for(let k = 0; k < words.length; k++){
    const w = words[k];
    if(x >= w.x - 1 && x <= w.x + w.w + 1 && y >= w.y - 1 && y <= w.y + w.h + 1){ at = k; break; }
  }
  if(at < 0) return null;

  let i = at, j = at;
  if(toNumber(words[at].t) !== null){
    while(i > 0 && adjacent(words[i - 1], words[i]) && FRAGMENT.test(words[i - 1].t)) i--;
    while(j < words.length - 1 && adjacent(words[j], words[j + 1]) && FRAGMENT.test(words[j + 1].t)) j++;
  }
  return { t: words.slice(i, j + 1).map(w => w.t).join(''), i, j, ...union(words, i, j) };
}

/** Search every page. @returns [{page,x,y,w,h}] in reading order */
export function findAll(pages, needle, exact){
  const out = [];
  pages.forEach((p, n) => {
    for(const h of findInPage(p.words, needle, exact)) out.push({ page: n, x: h.x, y: h.y, w: h.w, h: h.h });
  });
  return out;
}

/** Split a pasted Excel column into trimmed, non-empty cell values. */
export function parseColumn(text){
  return String(text)
    .split(/\r\n|\r|\n/)
    .map(l => l.split('\t')[0].trim())
    .filter(l => l.length > 0);
}
