import type { FormatHandler } from '../provider';
import type { TableModel, TablePayload } from './table';

interface CsvInfo {
  delimiter: string;
  eol: string;
  bom: boolean;
  trailingEol: boolean;
  rows: string[][];
}

function detectDelimiter(text: string, isTsv: boolean): string {
  if (isTsv) return '\t';
  const sample = text.slice(0, 10000).split(/\r?\n/).slice(0, 20);
  let best = ',';
  let bestScore = -1;
  for (const d of [',', ';', '\t', '|']) {
    const counts = sample.map((l) => l.split(d).length - 1).filter((n) => n > 0);
    if (!counts.length) continue;
    // Prefer delimiters that appear consistently on every line.
    const consistent = counts.filter((n) => n === counts[0]).length;
    const score = consistent * 1000 + counts[0];
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

export function parseCsv(bytes: Uint8Array, isTsv: boolean): CsvInfo {
  let text = new TextDecoder('utf-8').decode(bytes);
  const bom = text.charCodeAt(0) === 0xfeff;
  if (bom) text = text.slice(1);
  const delimiter = detectDelimiter(text, isTsv);
  const eol = text.includes('\r\n') ? '\r\n' : '\n';
  const trailingEol = /\r?\n$/.test(text);

  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else inQuotes = false;
      } else field += ch;
    } else if (ch === '"' && field === '') {
      inQuotes = true;
    } else if (ch === delimiter) {
      row.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  if (!rows.length) rows.push(['']);

  // Pad to a rectangle so the grid is regular.
  const width = Math.max(...rows.map((r) => r.length));
  for (const r of rows) while (r.length < width) r.push('');
  return { delimiter, eol, bom, trailingEol, rows };
}

function serialize(rows: string[][], info: CsvInfo): Uint8Array {
  const d = info.delimiter;
  const quote = (v: string) => (v.includes(d) || v.includes('"') || v.includes('\n') || v.includes('\r') ? `"${v.replace(/"/g, '""')}"` : v);
  let text = rows.map((r) => r.map(quote).join(d)).join(info.eol);
  if (info.trailingEol) text += info.eol;
  if (info.bom) text = '﻿' + text;
  return new TextEncoder().encode(text);
}

const isTsv = (path: string) => path.toLowerCase().endsWith('.tsv');

export const csvHandler: FormatHandler = {
  script: 'table',
  customUndo: true,
  async load(bytes, uri): Promise<TableModel> {
    const { rows } = parseCsv(bytes, isTsv(uri.path));
    return { kind: 'csv', sheets: [{ name: 'Sheet1', rows }] };
  },
  async save(base, payload: TablePayload, uri) {
    return serialize(payload.grids[0], parseCsv(base, isTsv(uri.path)));
  },
};
