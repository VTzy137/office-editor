import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist';
import type { TextItem } from 'pdfjs-dist/types/src/display/api';
import { start, markDirty, h, button } from './common';
import type { PdfEdit } from '../src/formats/pdf';

const base = document.body.dataset.pdfjs!;
pdfjsLib.GlobalWorkerOptions.workerSrc = `${base}/pdf.worker.min.mjs`;

/**
 * VS Code serves webview resources from another origin, so pdf.js would wrap the worker in a blob that
 * `import()`s it — but requests from inside a worker bypass the webview's resource loader, and the failed
 * import never fires the Worker's error event, leaving getDocument pending forever. Fetch the (self-contained)
 * worker on the main thread instead and start it from a same-origin blob URL.
 */
async function setupWorker() {
  try {
    const res = await fetch(pdfjsLib.GlobalWorkerOptions.workerSrc);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const url = URL.createObjectURL(new Blob([await res.text()], { type: 'text/javascript' }));
    pdfjsLib.GlobalWorkerOptions.workerPort = new Worker(url, { type: 'module' });
  } catch (err) {
    console.warn('pdf.js worker blob failed, using default loading', err);
  }
}
const workerReady = setupWorker();

const app = document.getElementById('app')!;
let pdf: PDFDocumentProxy;
let pagesEl: HTMLDivElement;
let zoomLabel: HTMLSpanElement;
let scale = 1;
let renderToken = 0;
/** Edited text items keyed by "page:item". */
const edits = new Map<string, PdfEdit & { original: string }>();

function fontInfo(page: PDFPageProxy, item: TextItem, styles: Record<string, { fontFamily: string }>) {
  let name = '';
  try {
    if (page.commonObjs.has(item.fontName)) name = (page.commonObjs.get(item.fontName) as any)?.name ?? '';
  } catch {
    /* font not loaded */
  }
  const family = styles[item.fontName]?.fontFamily ?? '';
  return {
    bold: /bold|black|heavy|semibold|demi/i.test(name),
    italic: /italic|oblique/i.test(name),
    family: (/courier|mono|consol/i.test(name) || family === 'monospace' ? 'mono' : /times|serif|georgia|garamond|roman|cambria/i.test(name) || family === 'serif' ? 'serif' : 'sans') as PdfEdit['family'],
  };
}

async function renderPage(pageNo: number, container: HTMLElement, token: number) {
  const page = await pdf.getPage(pageNo);
  const viewport = page.getViewport({ scale });
  const ratio = window.devicePixelRatio || 1;
  const canvas = h('canvas', { class: 'pdf-canvas' });
  canvas.width = Math.floor(viewport.width * ratio);
  canvas.height = Math.floor(viewport.height * ratio);
  canvas.style.width = `${viewport.width}px`;
  canvas.style.height = `${viewport.height}px`;
  const textLayer = h('div', { class: 'pdf-text-layer' });
  container.replaceChildren(canvas, textLayer);

  await page.render({ canvasContext: canvas.getContext('2d')!, viewport, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined }).promise;
  if (token !== renderToken) return;

  const content = await page.getTextContent();
  const styles = content.styles as Record<string, { fontFamily: string; ascent?: number }>;
  content.items.forEach((raw, idx) => {
    const item = raw as TextItem;
    if (!('str' in item) || !item.str.trim()) return;
    const [a, b, c, d, e, f] = item.transform;
    if (Math.abs(b) > 1e-3 || Math.abs(c) > 1e-3 || a <= 0) return; // only horizontal text is editable
    const tx = pdfjsLib.Util.transform(viewport.transform, item.transform);
    const fontHeight = Math.hypot(tx[2], tx[3]);
    const ascent = styles[item.fontName]?.ascent || 0.8;
    const key = `${pageNo - 1}:${idx}`;
    const edit = edits.get(key);

    const span = h('span', { class: 'pdf-text', contenteditable: 'plaintext-only', spellcheck: 'false' }, edit?.text ?? item.str);
    span.style.left = `${tx[4]}px`;
    span.style.top = `${tx[5] - fontHeight * ascent}px`;
    span.style.fontSize = `${fontHeight}px`;
    span.style.fontFamily = styles[item.fontName]?.fontFamily ?? 'sans-serif';
    span.style.lineHeight = '1';
    textLayer.appendChild(span);

    // Stretch the transparent text to match the rendered glyphs so selection lines up.
    const targetWidth = item.width * scale;
    // When edited, the white background must hide the original glyphs underneath (as the saved file does).
    span.style.setProperty('--orig-width', `${targetWidth}px`);
    const natural = span.getBoundingClientRect().width;
    if (natural > 0 && !edit) span.style.transform = `scaleX(${targetWidth / natural})`;
    if (edit) span.classList.add('edited');

    span.addEventListener('focus', () => {
      span.style.transform = '';
      span.classList.add('editing');
    });
    span.addEventListener('blur', () => span.classList.remove('editing'));
    span.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' || ev.key === 'Escape') {
        ev.preventDefault();
        span.blur();
      }
    });
    span.addEventListener('input', () => {
      const text = (span.textContent ?? '').replace(/ /g, ' ');
      if (text === item.str) {
        edits.delete(key);
        span.classList.remove('edited');
      } else {
        const fi = fontInfo(page, item, styles);
        edits.set(key, { page: pageNo - 1, x: e, y: f, width: item.width, size: Math.abs(d) || Math.hypot(c, d), text, original: item.str, ...fi });
        span.classList.add('edited');
      }
      markDirty();
    });
  });
}

function renderAll() {
  const token = ++renderToken;
  zoomLabel.textContent = `${Math.round(scale * 100)}%`;
  pagesEl.replaceChildren();
  const observer = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const el = entry.target as HTMLElement;
        observer.unobserve(el);
        renderPage(Number(el.dataset.page), el, token).catch((err) => console.error(err));
      }
    },
    { root: pagesEl, rootMargin: '600px' },
  );
  // Size placeholders from page 1 so scrolling works before pages render.
  pdf.getPage(1).then((first) => {
    const vp = first.getViewport({ scale });
    for (let i = 1; i <= pdf.numPages; i++) {
      const el = h('div', { class: 'pdf-page', 'data-page': i });
      el.style.width = `${vp.width}px`;
      el.style.height = `${vp.height}px`;
      pagesEl.appendChild(el);
      observer.observe(el);
      // Pages can differ in size: fix each once known.
      pdf.getPage(i).then((p) => {
        const v = p.getViewport({ scale });
        el.style.width = `${v.width}px`;
        el.style.height = `${v.height}px`;
      });
    }
  });
}

function zoom(next: number) {
  scale = Math.max(0.25, Math.min(4, next));
  renderAll();
}

start<{ data: Uint8Array }>({
  async init(model, fileName) {
    const raw: any = model.data;
    const data =
      raw instanceof Uint8Array ? raw.slice()
      : raw instanceof ArrayBuffer ? new Uint8Array(raw)
      : raw?.type === 'Buffer' && Array.isArray(raw.data) ? new Uint8Array(raw.data)
      : new Uint8Array(Object.values(raw));
    await workerReady;
    pdf = await pdfjsLib.getDocument({
      data,
      cMapUrl: `${base}/cmaps/`,
      cMapPacked: true,
      standardFontDataUrl: `${base}/standard_fonts/`,
      useWorkerFetch: false,
      isEvalSupported: false,
    }).promise;

    zoomLabel = h('span', { class: 'tb-zoom' });
    pagesEl = h('div', { class: 'pdf-pages' });
    app.replaceChildren(
      h(
        'div',
        { class: 'pdf-editor' },
        h(
          'div',
          { class: 'toolbar' },
          h('span', { class: 'tb-title' }, fileName),
          h('span', { class: 'tb-sep' }),
          button('−', 'Zoom out', () => zoom(scale / 1.2)),
          zoomLabel,
          button('+', 'Zoom in', () => zoom(scale * 1.2)),
          button('Fit', 'Fit width', () => fitWidth()),
          h('span', { class: 'tb-hint' }, `${pdf.numPages} pages · click text to edit`),
        ),
        pagesEl,
      ),
    );
    await fitWidth();
  },
  getContent() {
    return { edits: [...edits.values()].map(({ original, ...e }) => e) };
  },
});

async function fitWidth() {
  const first = await pdf.getPage(1);
  const width = first.getViewport({ scale: 1 }).width;
  zoom(Math.min(1.5, (pagesEl.clientWidth - 48) / width));
}
