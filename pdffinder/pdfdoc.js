// PDF loading, page rendering and word extraction (embedded text layer, or OCR).
import * as pdfjs from '../vendor/pdfjs/pdf.min.mjs';

pdfjs.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.mjs';

const RENDER_SCALE = 1.5;   // on-screen crispness
const OCR_SCALE = 2.2;      // resolution handed to Tesseract

export async function openPdf(data){
  const doc = await pdfjs.getDocument({
    data,
    cMapUrl: 'vendor/pdfjs/cmaps/',
    cMapPacked: true,
    standardFontDataUrl: 'vendor/pdfjs/standard_fonts/'
  }).promise;

  const pages = [];
  for(let n = 1; n <= doc.numPages; n++){
    const page = await doc.getPage(n);
    const vp = page.getViewport({ scale: 1 });
    pages.push({ page, width: vp.width, height: vp.height, words: await textWords(page, vp), ocr: false });
  }
  return { doc, pages };
}

/** Words from the embedded text layer, split on spaces with proportional boxes. */
async function textWords(page, vp){
  const content = await page.getTextContent();
  const words = [];

  for(const item of content.items){
    if(!item.str || !item.str.trim()) continue;
    const tx = pdfjs.Util.transform(vp.transform, item.transform);
    const h = Math.hypot(tx[2], tx[3]) || item.height || 10;
    const x = tx[4];
    const y = tx[5] - h;
    const w = item.width || 0;
    if(w <= 0) continue;

    const per = w / item.str.length;      // monospace approximation, good enough for a box
    let at = 0;
    for(const part of item.str.split(/(\s+)/)){
      if(part.trim()) words.push({ t: part, x: x + at * per, y, w: part.length * per, h });
      at += part.length;
    }
  }
  return words;
}

/** True when a page has too little embedded text to reconcile against. */
export function needsOcr(pg){
  return pg.words.length < 8;
}

export async function renderPage(pg, canvas){
  const vp = pg.page.getViewport({ scale: RENDER_SCALE * (window.devicePixelRatio || 1) });
  canvas.width = vp.width;
  canvas.height = vp.height;
  canvas.style.width = (pg.width * RENDER_SCALE) + 'px';
  canvas.style.height = (pg.height * RENDER_SCALE) + 'px';
  await pg.page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
}

/* ---------- OCR ---------- */

let worker = null;

function loadTesseract(){
  if(window.Tesseract) return Promise.resolve();
  return new Promise((ok, fail) => {
    const s = document.createElement('script');
    s.src = 'vendor/tesseract/tesseract.min.js';
    s.onload = ok;
    s.onerror = () => fail(new Error('tesseract'));
    document.head.appendChild(s);
  });
}

async function getWorker(onProgress){
  if(worker) return worker;
  await loadTesseract();
  worker = await window.Tesseract.createWorker('eng', 1, {
    workerPath: 'vendor/tesseract/worker.min.js',
    corePath: 'vendor/tesseract/core',
    langPath: 'vendor/tesseract/lang',
    gzip: true,
    logger: m => { if(m.status === 'recognizing text' && onProgress) onProgress(m.progress); }
  });
  return worker;
}

/** Replace a page's words with OCR results. */
export async function ocrPage(pg, onProgress){
  const w = await getWorker(onProgress);
  const vp = pg.page.getViewport({ scale: OCR_SCALE });
  const c = document.createElement('canvas');
  c.width = vp.width; c.height = vp.height;
  await pg.page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;

  const { data } = await w.recognize(c, {}, { blocks: true });

  const raw = data.words && data.words.length
    ? data.words
    : (data.blocks || []).flatMap(b => (b.paragraphs || []).flatMap(p => (p.lines || []).flatMap(l => l.words || [])));

  const words = raw
    .filter(o => o && o.bbox && o.text && o.text.trim())
    .map(o => ({
      t: o.text.trim(),
      x: o.bbox.x0 / OCR_SCALE, y: o.bbox.y0 / OCR_SCALE,
      w: (o.bbox.x1 - o.bbox.x0) / OCR_SCALE, h: (o.bbox.y1 - o.bbox.y0) / OCR_SCALE
    }));

  pg.words = words;
  pg.ocr = true;
  return words.length;
}
