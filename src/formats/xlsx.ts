import ExcelJS from 'exceljs';
import type { FormatHandler } from '../provider';
import type { SheetModel, TableModel, TablePayload } from './table';

const pad = (n: number) => String(n).padStart(2, '0');

function formatDate(d: Date): string {
  const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const h = d.getUTCHours(), m = d.getUTCMinutes(), s = d.getUTCSeconds();
  if (!h && !m && !s) return date;
  return `${date} ${pad(h)}:${pad(m)}${s ? ':' + pad(s) : ''}`;
}

function formatNumber(n: number, numFmt?: string): string {
  if (numFmt && numFmt.includes('%')) return `${Number((n * 100).toPrecision(12))}%`;
  return String(Number(n.toPrecision(15)));
}

function displayOf(value: ExcelJS.CellValue, numFmt?: string): string {
  if (value == null) return '';
  if (typeof value === 'number') return formatNumber(value, numFmt);
  if (typeof value === 'string') return value;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value instanceof Date) return formatDate(value);
  const v = value as any;
  if (v.richText) return v.richText.map((t: any) => t.text).join('');
  if (v.error) return String(v.error);
  if ('formula' in v || 'sharedFormula' in v) return v.result === undefined ? '' : displayOf(v.result, numFmt);
  if ('hyperlink' in v) return displayOf(v.text, numFmt);
  return String(value);
}

function editOf(cell: ExcelJS.Cell): string {
  const v = cell.value as any;
  if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
    const f = cell.formula;
    if (f) return '=' + f;
  }
  return displayOf(cell.value, cell.numFmt);
}

function parseInput(text: string, old: ExcelJS.CellValue): ExcelJS.CellValue {
  const t = text.trim();
  if (text === '') return null;
  if (t.startsWith('=') && t.length > 1) return { formula: t.slice(1) } as ExcelJS.CellFormulaValue;
  let value: ExcelJS.CellValue = text;
  if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) value = Number(t);
  else if (/^[-+]?(\d+\.?\d*|\.\d+)%$/.test(t)) value = Number(t.slice(0, -1)) / 100;
  else if (/^(true|false)$/i.test(t)) value = t.toLowerCase() === 'true';
  else {
    const m = t.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) value = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0)));
  }
  // Keep hyperlinks attached when only the label changes.
  const o = old as any;
  if (typeof value === 'string' && o && typeof o === 'object' && o.hyperlink) return { text: value, hyperlink: o.hyperlink };
  return value;
}

async function loadWorkbook(bytes: Uint8Array) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(bytes) as any);
  return wb;
}

export const xlsxHandler: FormatHandler = {
  script: 'table',
  customUndo: true,
  async load(bytes): Promise<TableModel> {
    const wb = await loadWorkbook(bytes);
    const sheets: SheetModel[] = wb.worksheets.map((ws) => {
      const rowCount = Math.max(ws.rowCount, 1);
      const colCount = Math.max(ws.columnCount, 1);
      const rows: string[][] = [];
      const edit: Record<string, string> = {};
      for (let r = 1; r <= rowCount; r++) {
        const row = ws.getRow(r);
        const out: string[] = [];
        for (let c = 1; c <= colCount; c++) {
          const cell = row.getCell(c);
          // Only the master cell of a merged range carries the value.
          if (cell.isMerged && cell.master !== cell) {
            out.push('');
            continue;
          }
          const display = displayOf(cell.value, cell.numFmt);
          const e = editOf(cell);
          out.push(display);
          if (e !== display) edit[`${r - 1},${c - 1}`] = e;
        }
        rows.push(out);
      }
      return { name: ws.name, rows, edit };
    });
    return { kind: 'xlsx', sheets };
  },
  async save(base, payload: TablePayload) {
    const wb = await loadWorkbook(base);
    let hasFormula = false;
    for (const op of payload.ops) {
      const ws = wb.worksheets[op.s];
      if (!ws) continue;
      switch (op.k) {
        case 'set': {
          const cell = ws.getCell(op.r + 1, op.c + 1);
          const value = parseInput(op.v, cell.value);
          if (value && typeof value === 'object' && 'formula' in value) hasFormula = true;
          cell.value = value;
          break;
        }
        case 'insRow':
          ws.spliceRows(op.r + 1, 0, []);
          break;
        case 'delRow':
          ws.spliceRows(op.r + 1, 1);
          break;
        case 'insCol':
          ws.spliceColumns(op.c + 1, 0, []);
          break;
        case 'delCol':
          ws.spliceColumns(op.c + 1, 1);
          break;
      }
    }
    // Formulas we write have no cached result, so ask Excel to recalculate on open.
    if (hasFormula) (wb as any).calcProperties = { ...((wb as any).calcProperties ?? {}), fullCalcOnLoad: true };
    const out = await wb.xlsx.writeBuffer();
    return new Uint8Array(out as ArrayBuffer);
  },
};
