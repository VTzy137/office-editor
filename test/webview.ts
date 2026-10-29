// Drives the real webview bundles in Chromium with a stubbed VS Code API, then saves their payloads
// through the format handlers. Run: npm run build && npm run test:webview
import { createServer } from 'node:http';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { extname, join } from 'node:path';
import { chromium, Page } from 'playwright-core';
import { xlsxHandler } from '../src/formats/xlsx';
import { docxHandler } from '../src/formats/docx';
import { pptxHandler } from '../src/formats/pptx';
import { pdfHandler } from '../src/formats/pdf';

const root = process.cwd();
const outDir = join(root, 'test/out/webview');
mkdirSync(outDir, { recursive: true });
const uri = (p: string) => ({ path: p }) as any;
const sample = (f: string) => new Uint8Array(readFileSync(join(root, 'samples', f)));

const TYPES: Record<string, string> = { '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.bcmap': 'application/octet-stream', '.pfb': 'application/octet-stream' };
const server = createServer((req, res) => {
  const url = new URL(req.url!, 'http://x');
  if (url.pathname === '/harness.html') {
    res.setHeader('content-type', 'text/html');
    res.end(`<!DOCTYPE html><html><head><meta charset="utf-8"><link rel="stylesheet" href="/media/style.css">
<script>window.__msgs=[];window.acquireVsCodeApi=()=>({postMessage:(m)=>window.__msgs.push(m)});</script></head>
<body data-pdfjs="/media/dist/pdfjs"><div id="app"></div><script type="module" src="/media/dist/${url.searchParams.get('s')}.js"></script></body></html>`);
    return;
  }
  const file = join(root, decodeURIComponent(url.pathname));
  if (!existsSync(file)) return void res.writeHead(404).end();
  res.setHeader('content-type', TYPES[extname(file)] ?? 'application/octet-stream');
  res.end(readFileSync(file));
});

async function open(page: Page, script: string, model: unknown) {
  page.on('console', (m) => m.type() === 'error' && console.log('  [console]', m.text()));
  page.on('pageerror', (e) => console.log('  [pageerror]', e.message));
  await page.goto(`http://localhost:${(server.address() as any).port}/harness.html?s=${script}`);
  await page.waitForFunction(() => (window as any).__msgs.some((m: any) => m.type === 'ready'));
  await page.evaluate((m: any) => {
    if (m && m.__bytes) m = { data: new Uint8Array(m.__bytes) };
    window.postMessage({ type: 'init', model: m, fileName: 'sample' }, '*');
  }, model as any);
}

async function content(page: Page): Promise<any> {
  await page.evaluate(() => window.postMessage({ type: 'getContent', requestId: 1 }, '*'));
  await page.waitForFunction(() => (window as any).__msgs.some((m: any) => m.type === 'content'));
  return page.evaluate(() => {
    const w = window as any;
    const i = w.__msgs.findIndex((m: any) => m.type === 'content');
    return w.__msgs.splice(i, 1)[0].payload;
  });
}

const check = (label: string, ok: boolean, detail?: unknown) => {
  console.log(`${ok ? 'PASS' : 'FAIL'} ${label}${ok ? '' : ' ' + JSON.stringify(detail)}`);
  if (!ok) process.exitCode = 1;
};

async function testTable(page: Page) {
  const model = await xlsxHandler.load(sample('sample.xlsx'), uri('a.xlsx'));
  await open(page, 'table', model);
  await page.waitForSelector('.grid td');
  await page.click('.grid tbody tr:nth-child(2) td:nth-child(3)'); // B2
  await page.keyboard.type('77');
  await page.keyboard.press('Enter');
  await page.click('button[title="Insert row above"]');
  await page.keyboard.type('Hello');
  await page.keyboard.press('Tab');
  await page.screenshot({ path: join(outDir, 'table.png') });
  let payload = await content(page);
  check('table ops recorded', payload.ops.length === 3, payload.ops);
  // VS Code-driven undo of the typed "Hello"
  await page.evaluate(() => window.postMessage({ type: 'undo' }, '*'));
  payload = await content(page);
  const saved = await xlsxHandler.save(sample('sample.xlsx'), payload, uri('a.xlsx'));
  const reloaded: any = await xlsxHandler.load(saved, uri('a.xlsx'));
  check('xlsx B2 = 77', reloaded.sheets[0].rows[1][1] === '77', reloaded.sheets[0].rows[1]);
  check('xlsx inserted row empty after undo', reloaded.sheets[0].rows[2].every((v: string) => v === ''), reloaded.sheets[0].rows[2]);
  check('xlsx Pear row shifted down', reloaded.sheets[0].rows[3][0] === 'Pear', reloaded.sheets[0].rows[3]);
}

async function testDocx(page: Page) {
  const model = await docxHandler.load(sample('sample.docx'), uri('a.docx'));
  await open(page, 'docx', model);
  await page.waitForSelector('.doc-page p');
  // Append text to the mixed-format paragraph and bold one word.
  const para = page.locator('p[data-pid="2"]');
  await para.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Appended');
  await page.keyboard.down('Shift');
  for (let i = 0; i < 8; i++) await page.keyboard.press('ArrowLeft');
  await page.keyboard.up('Shift');
  await page.keyboard.press('Control+b');
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Brand new paragraph');
  await page.selectOption('select[title="Paragraph style"]', 'Heading1');
  // Font size on a selection in the final paragraph
  const last = page.locator('p[data-pid="11"]');
  await last.dblclick();
  await page.selectOption('select[title="Font size"]', '24');
  await page.screenshot({ path: join(outDir, 'docx.png'), fullPage: true });
  const payload = await content(page);
  const saved = await docxHandler.save(sample('sample.docx'), payload, uri('a.docx'));
  writeFileSync(join(outDir, 'edited.docx'), saved);
  const reloaded: any = await docxHandler.load(saved, uri('a.docx'));
  const paras = reloaded.blocks.filter((b: any) => b.t === 'p');
  const p2 = paras[2];
  check('docx appended text', p2.runs.map((r: any) => r.text).join('') === 'This is bold and italic and big text. Appended', p2.runs);
  check('docx bold applied', p2.runs.some((r: any) => r.text.includes('Appended') && r.b), p2.runs);
  check('docx original formatting kept', p2.runs.some((r: any) => r.text === 'italic' && r.i) && p2.runs.some((r: any) => r.sz === 20), p2.runs);
  check('docx new paragraph w/ style', paras[3].style === 'Heading1' && paras[3].runs.map((r: any) => r.text).join('') === 'Brand new paragraph', paras[3]);
  const final = paras.find((p: any) => p.runs.some((r: any) => r.text.includes('Final') || r.text.includes('paragraph.')));
  check('docx font size applied', final?.runs.some((r: any) => r.sz === 24), final?.runs);
  check('docx untouched paragraphs unchanged', payload.paras.filter((p: any) => p.same).length >= 8, payload.paras.length);
}

async function testPptx(page: Page) {
  const model = await pptxHandler.load(sample('sample.pptx'), uri('a.pptx'));
  await open(page, 'pptx', model);
  await page.waitForSelector('.slide-para[data-id="0:0"]');
  await page.click('.slide-para[data-id="0:0"]');
  await page.keyboard.press('Control+a');
  await page.keyboard.type('Tiêu đề mới');
  await page.click('.slide-para[data-id="1:1"]');
  await page.keyboard.press('End');
  await page.keyboard.type(' (edited)');
  await page.screenshot({ path: join(outDir, 'pptx.png'), fullPage: true });
  const payload = await content(page);
  const saved = await pptxHandler.save(sample('sample.pptx'), payload, uri('a.pptx'));
  writeFileSync(join(outDir, 'edited.pptx'), saved);
  const reloaded: any = await pptxHandler.load(saved, uri('a.pptx'));
  const text = (s: number) => reloaded.slides[s].elements.flatMap((e: any) => e.paras ?? []).filter((p: any) => p.id).map((p: any) => p.items.map((i: any) => i.text ?? '').join(''));
  check('pptx title edited', text(0)[0] === 'Tiêu đề mới', text(0));
  check('pptx bullet edited', text(1)[1] === 'Point one (edited)', text(1));
  check('pptx only changed paragraphs sent', payload.paras.length === 2, payload.paras);
}

async function testPdf(page: Page) {
  const bytes = sample('sample.pdf');
  await open(page, 'pdf', { __bytes: Array.from(bytes) });
  const span = page.locator('.pdf-text', { hasText: 'This line of text can be edited.' });
  await span.waitFor({ timeout: 20000 });
  await span.click();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('Đã sửa dòng này');
  await page.keyboard.press('Enter');
  await page.screenshot({ path: join(outDir, 'pdf.png') });
  const payload = await content(page);
  check('pdf edit captured', payload.edits.length === 1 && payload.edits[0].text === 'Đã sửa dòng này' && Math.round(payload.edits[0].y) === 720, payload.edits);
  const saved = await pdfHandler.save(bytes, payload, uri('a.pdf'));
  writeFileSync(join(outDir, 'edited.pdf'), saved);
  check('pdf saved', saved.length > bytes.length);
}

server.listen(0, async () => {
  const browser = await chromium.launch({ channel: process.env.BROWSER_CHANNEL ?? 'msedge' });
  try {
    for (const t of [testTable, testDocx, testPptx, testPdf]) {
      const page = await browser.newPage({ viewport: { width: 1200, height: 900 } });
      console.log(`-- ${t.name}`);
      await t(page).catch((e) => check(t.name, false, String(e)));
      await page.close();
    }
  } finally {
    await browser.close();
    server.close();
  }
});
