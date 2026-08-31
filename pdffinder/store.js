// Per-document persistence: which printed occurrence each cell claimed, and the
// marker colour, survive closing the finder window.

export function load(key){
  try{
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  }catch{ return null; }
}

export function save(key, state){
  try{ localStorage.setItem(key, JSON.stringify(state)); }catch{ /* quota — marks stay in memory */ }
}

export function drop(key){
  try{ localStorage.removeItem(key); }catch{ /* ignore */ }
}
