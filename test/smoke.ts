import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { csvHandler } from '../src/formats/csv';
import { xlsxHandler } from '../src/formats/xlsx';
import { docxHandler } from '../src/formats/docx';
import { pptxHandler } from '../src/formats/pptx';
import { pdfHandler } from '../src/formats/pdf';
const dir = process.argv[2], out = process.argv[3];
mkdirSync(out, { recursive: true });
const uri = (p: string) => ({ path: p }) as any;
const rd = (f: string) => new Uint8Array(readFileSync(`${dir}/${f}`));
(async () => {
  // CSV
  const csv: any = await csvHandler.load(rd('sample.csv'), uri('sample.csv'));
  console.log('CSV', JSON.stringify(csv.sheets[0].rows));
  const g = csv.sheets[0].rows.map((r: string[]) => [...r]); g[1][0] = 'Ân, "quoted"'; g.push(['New', '1', 'x']);
  const csvOut = await csvHandler.save(rd('sample.csv'), { ops: [], grids: [g] }, uri('sample.csv'));
  writeFileSync(`${out}/out.csv`, csvOut); console.log(JSON.stringify(new TextDecoder().decode(csvOut)));

  // XLSX
  const x: any = await xlsxHandler.load(rd('sample.xlsx'), uri('sample.xlsx'));
  console.log('XLSX', x.sheets.map((s: any) => s.name), JSON.stringify(x.sheets[0].rows), JSON.stringify(x.sheets[0].edit));
  const xOut = await xlsxHandler.save(rd('sample.xlsx'), { grids: [], ops: [
    { k: 'set', s: 0, r: 1, c: 1, v: '10' }, { k: 'insRow', s: 0, r: 1 }, { k: 'set', s: 0, r: 1, c: 0, v: 'Inserted' },
    { k: 'set', s: 0, r: 4, c: 3, v: '=SUM(D3:D4)' }, { k: 'insCol', s: 0, c: 0 }, { k: 'delCol', s: 0, c: 0 }, { k: 'set', s: 1, r: 0, c: 0, v: '2026-10-29' } ] }, uri('x'));
  writeFileSync(`${out}/out.xlsx`, xOut);
  const x2: any = await xlsxHandler.load(xOut, uri('x'));
  console.log('XLSX after', JSON.stringify(x2.sheets[0].rows), JSON.stringify(x2.sheets[0].edit), JSON.stringify(x2.sheets[1].rows));

  // DOCX
  const d: any = await docxHandler.load(rd('sample.docx'), uri('d'));
  const flat = (bs: any[]): any[] => bs.flatMap((b) => (b.t === 'p' ? [b] : b.rows.flatMap((r: any) => r.flatMap((c: any) => flat(c.blocks)))));
  for (const p of flat(d.blocks)) console.log('P', p.id, p.style, p.label ?? '', p.locked ? 'LOCKED' : '', JSON.stringify(p.runs.map((r: any) => [r.text, r.b, r.i, r.sz])));
  console.log('styles', d.styles.length, 'css', d.css.split('\n').length, 'default', d.defaultStyle);
  const all = flat(d.blocks);
  const paras = all.map((p: any) => ({ id: p.id, same: true, style: p.style, runs: [] }));
  // edit paragraph 2 (bold mix), split it, delete bullet 4, restyle 1, edit a table cell
  paras[2] = { id: 2, style: 'Normal', runs: [{ text: 'Edited ', ref: '2:0', b: null, i: null, u: false, sz: null }, { text: 'BOLD', ref: '2:1', b: true, i: null, u: true, sz: 16 }] };
  paras.splice(3, 0, { id: 2, style: 'Normal', runs: [{ text: 'Split-off paragraph\twith tab', ref: '2:4', b: null, i: true, u: false, sz: null }] });
  const del = paras.findIndex((p: any) => p.id === 4); paras.splice(del, 1);
  paras[1] = { id: 1, style: 'Heading2', runs: [{ text: '1. Giới thiệu (sửa)', ref: '1:0', b: null, i: null, u: false, sz: null }] };
  const cell = paras.find((p: any) => p.id === all.find((q: any) => q.runs[0]?.text === '42').id);
  Object.assign(cell, { same: false, runs: [{ text: '99', ref: null, b: null, i: null, u: false, sz: null }] });
  const dOut = await docxHandler.save(rd('sample.docx'), { paras }, uri('d'));
  writeFileSync(`${out}/out.docx`, dOut);
  const d2: any = await docxHandler.load(dOut, uri('d'));
  for (const p of flat(d2.blocks)) console.log('P2', p.id, p.style, JSON.stringify(p.runs.map((r: any) => [r.text, r.b, r.i, r.u, r.sz])));

  // PPTX
  const pp: any = await pptxHandler.load(rd('sample.pptx'), uri('p'));
  console.log('PPTX', pp.width, pp.height, pp.slides.length);
  for (const s of pp.slides) for (const e of s.elements) console.log(' el', e.k, Math.round(e.x), Math.round(e.y), Math.round(e.w), Math.round(e.h), JSON.stringify((e.paras ?? e.rows?.flatMap((r: any) => r.cells.flatMap((c: any) => c.paras)) ?? []).map((p: any) => [p.id, p.align, p.bullet, p.items.map((i: any) => [i.text, Math.round(i.sz), i.color, i.font])])));
  const pOut = await pptxHandler.save(rd('sample.pptx'), { paras: [
    { id: '0:0', runs: { 0: 'Xin chào Slides' }, extra: null },
    { id: '1:2', runs: {}, extra: 'Replaced point two' },
    { id: '1:4', runs: { 0: 'cell edited' }, extra: null },
  ] }, uri('p'));
  writeFileSync(`${out}/out.pptx`, pOut);
  const pp2: any = await pptxHandler.load(pOut, uri('p'));
  console.log('PPTX after', JSON.stringify(pp2.slides.map((s: any) => s.elements.flatMap((e: any) => (e.paras ?? e.rows?.flatMap((r: any) => r.cells.flatMap((c: any) => c.paras)) ?? []).filter((p: any) => p.id).map((p: any) => p.items.map((i: any) => i.text).join(''))))));

  // PDF
  const pdfOut = await pdfHandler.save(rd('sample.pdf'), { edits: [
    { page: 0, x: 72, y: 720, width: 170, size: 12, text: 'Dòng này đã được sửa (Vietnamese ✓)', bold: false, italic: false, family: 'sans' },
    { page: 1, x: 72, y: 750, width: 90, size: 14, text: 'Edited page two', bold: true, italic: false, family: 'serif' },
  ] }, uri('pdf'));
  writeFileSync(`${out}/out.pdf`, pdfOut);
  console.log('PDF bytes', rd('sample.pdf').length, '->', pdfOut.length);
})().catch((e) => { console.error(e); process.exit(1); });
