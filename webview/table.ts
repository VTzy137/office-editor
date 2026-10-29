import { start, markDirty, h, button } from './common';
import type { TableModel, TableOp, TablePayload } from '../src/formats/table';

type LocalOp = TableOp & { d?: string };

interface Sheet {
  name: string;
  edit: string[][];
  disp: string[][];
}

const CHUNK = 300;
const colName = (c: number) => {
  let s = '';
  for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
};
const isNumeric = (v: string) => /^[-+]?[\d,]*\.?\d+(e[-+]?\d+)?%?$/i.test(v.trim());

let sheets: Sheet[] = [];
let active = 0;
let sel = { r: 0, c: 0 };
let rendered = CHUNK;
let editing: { input: HTMLInputElement; r: number; c: number } | null = null;

// Undo history (driven by VS Code's undo/redo) and the net op log sent on save.
const history: { redo: LocalOp[]; undo: LocalOp[] }[] = [];
let historyPos = 0;
const log: TableOp[] = [];

const app = document.getElementById('app')!;
let grid: HTMLTableElement;
let scroller: HTMLDivElement;
let addressBox: HTMLSpanElement;
let formulaInput: HTMLInputElement;
let tabs: HTMLDivElement;

const sheet = () => sheets[active];
const width = (s: Sheet) => s.edit[0]?.length ?? 0;

/** Apply an op to the local model and return the ops that revert it. */
function apply(op: LocalOp): LocalOp[] {
  const s = sheets[op.s];
  switch (op.k) {
    case 'set': {
      const inverse: LocalOp = { k: 'set', s: op.s, r: op.r, c: op.c, v: s.edit[op.r][op.c], d: s.disp[op.r][op.c] };
      s.edit[op.r][op.c] = op.v;
      s.disp[op.r][op.c] = op.d ?? op.v;
      return [inverse];
    }
    case 'insRow':
      s.edit.splice(op.r, 0, new Array(width(s)).fill(''));
      s.disp.splice(op.r, 0, new Array(width(s)).fill(''));
      return [{ k: 'delRow', s: op.s, r: op.r }];
    case 'delRow': {
      const edit = s.edit.splice(op.r, 1)[0];
      const disp = s.disp.splice(op.r, 1)[0];
      const restore: LocalOp[] = edit.map((v, c) => ({ k: 'set', s: op.s, r: op.r, c, v, d: disp[c] }) as LocalOp).filter((o) => (o as any).v !== '');
      return [{ k: 'insRow', s: op.s, r: op.r }, ...restore];
    }
    case 'insCol':
      for (const row of s.edit) row.splice(op.c, 0, '');
      for (const row of s.disp) row.splice(op.c, 0, '');
      return [{ k: 'delCol', s: op.s, c: op.c }];
    case 'delCol': {
      const restore: LocalOp[] = [];
      s.edit.forEach((row, r) => {
        const v = row.splice(op.c, 1)[0];
        const d = s.disp[r].splice(op.c, 1)[0];
        if (v !== '') restore.push({ k: 'set', s: op.s, r, c: op.c, v, d });
      });
      return [{ k: 'insCol', s: op.s, c: op.c }, ...restore];
    }
  }
}

const strip = (op: LocalOp): TableOp => {
  const { d, ...rest } = op;
  return rest as TableOp;
};

function run(ops: LocalOp[]): LocalOp[] {
  const inverses: LocalOp[] = [];
  for (const op of ops) {
    inverses.unshift(...apply(op));
    log.push(strip(op));
  }
  return inverses;
}

function perform(ops: LocalOp[], label: string, structural = false) {
  if (!ops.length) return;
  const undo = run(ops);
  history.length = historyPos;
  history.push({ redo: ops, undo });
  historyPos++;
  markDirty(label);
  structural ? renderGrid() : ops.forEach((o) => o.k === 'set' && refreshCell(o.r, o.c));
  updateFormulaBar();
}

function setCell(r: number, c: number, v: string) {
  if (sheet().edit[r][c] === v) return;
  perform([{ k: 'set', s: active, r, c, v }], 'Edit cell');
}

// ---------- Rendering ----------

function cellEl(r: number, c: number): HTMLTableCellElement | undefined {
  return grid.tBodies[0]?.rows[r]?.cells[c + 1];
}

function fillCell(td: HTMLTableCellElement, r: number, c: number) {
  const s = sheet();
  const d = s.disp[r][c];
  td.textContent = d;
  td.title = d.length > 30 ? d : '';
  td.classList.toggle('num', isNumeric(d));
  td.classList.toggle('formula', s.edit[r][c].startsWith('=') && s.edit[r][c] === d);
}

function refreshCell(r: number, c: number) {
  const td = cellEl(r, c);
  if (td) fillCell(td, r, c);
}

function renderRows(tbody: HTMLTableSectionElement, from: number, to: number) {
  const s = sheet();
  const frag = document.createDocumentFragment();
  for (let r = from; r < Math.min(to, s.edit.length); r++) {
    const tr = document.createElement('tr');
    tr.appendChild(h('th', { class: 'row-head' }, String(r + 1)));
    for (let c = 0; c < width(s); c++) {
      const td = document.createElement('td');
      fillCell(td, r, c);
      tr.appendChild(td);
    }
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

function renderGrid() {
  const s = sheet();
  if (!s.edit.length) {
    s.edit.push(['']);
    s.disp.push(['']);
  }
  sel.r = Math.min(sel.r, s.edit.length - 1);
  sel.c = Math.min(sel.c, width(s) - 1);
  rendered = Math.max(rendered, sel.r + 50);
  const head = h('tr', {}, h('th', { class: 'corner' }));
  for (let c = 0; c < width(s); c++) head.appendChild(h('th', { class: 'col-head' }, colName(c)));
  const tbody = document.createElement('tbody');
  grid.replaceChildren(h('thead', {}, head), tbody);
  renderRows(tbody, 0, rendered);
  select(sel.r, sel.c, false);
  renderTabs();
}

function renderTabs() {
  tabs.replaceChildren(
    ...sheets.map((s, i) => {
      const t = h('button', { class: 'sheet-tab' + (i === active ? ' active' : ''), type: 'button' }, s.name);
      t.addEventListener('click', () => {
        commitEdit();
        active = i;
        sel = { r: 0, c: 0 };
        rendered = CHUNK;
        renderGrid();
        scroller.scrollTo(0, 0);
      });
      return t;
    }),
  );
  tabs.hidden = sheets.length < 2;
}

function ensureRendered(r: number) {
  if (r < rendered) return;
  const tbody = grid.tBodies[0];
  renderRows(tbody, rendered, r + CHUNK);
  rendered = r + CHUNK;
}

function select(r: number, c: number, scroll = true) {
  const s = sheet();
  r = Math.max(0, Math.min(r, s.edit.length - 1));
  c = Math.max(0, Math.min(c, width(s) - 1));
  ensureRendered(r);
  grid.querySelectorAll('.selected').forEach((e) => e.classList.remove('selected'));
  grid.querySelectorAll('.head-active').forEach((e) => e.classList.remove('head-active'));
  sel = { r, c };
  const td = cellEl(r, c);
  td?.classList.add('selected');
  grid.tHead?.rows[0].cells[c + 1]?.classList.add('head-active');
  grid.tBodies[0].rows[r]?.cells[0]?.classList.add('head-active');
  if (scroll && td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  updateFormulaBar();
}

function updateFormulaBar() {
  if (!addressBox) return;
  addressBox.textContent = `${colName(sel.c)}${sel.r + 1}`;
  if (document.activeElement !== formulaInput) formulaInput.value = sheet().edit[sel.r]?.[sel.c] ?? '';
}

// ---------- Editing ----------

function beginEdit(initial?: string) {
  commitEdit();
  const { r, c } = sel;
  const td = cellEl(r, c);
  if (!td) return;
  const input = h('input', { class: 'cell-input', spellcheck: 'false' });
  input.value = initial ?? sheet().edit[r][c];
  td.textContent = '';
  td.appendChild(input);
  editing = { input, r, c };
  input.focus();
  input.setSelectionRange(input.value.length, input.value.length);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitEdit();
      select(sel.r + (e.shiftKey ? -1 : 1), sel.c);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commitEdit();
      select(sel.r, sel.c + (e.shiftKey ? -1 : 1));
    } else if (e.key === 'Escape') {
      cancelEdit();
    }
    e.stopPropagation();
  });
  input.addEventListener('input', () => (formulaInput.value = input.value));
  input.addEventListener('blur', () => commitEdit());
}

function commitEdit() {
  if (!editing) return;
  const { input, r, c } = editing;
  editing = null;
  input.remove();
  refreshCell(r, c);
  setCell(r, c, input.value);
  scroller.focus({ preventScroll: true });
}

function cancelEdit() {
  if (!editing) return;
  const { input, r, c } = editing;
  editing = null;
  input.remove();
  refreshCell(r, c);
  scroller.focus({ preventScroll: true });
}

function onKey(e: KeyboardEvent) {
  if (editing || e.target === formulaInput) return;
  const mod = e.ctrlKey || e.metaKey;
  const moves: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
  if (moves[e.key]) {
    e.preventDefault();
    const [dr, dc] = moves[e.key];
    const s = sheet();
    if (mod) select(dr > 0 ? s.edit.length - 1 : dr < 0 ? 0 : sel.r, dc > 0 ? width(s) - 1 : dc < 0 ? 0 : sel.c);
    else select(sel.r + dr, sel.c + dc);
  } else if (e.key === 'Tab') {
    e.preventDefault();
    select(sel.r, sel.c + (e.shiftKey ? -1 : 1));
  } else if (e.key === 'Enter' || e.key === 'F2') {
    e.preventDefault();
    beginEdit();
  } else if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    setCell(sel.r, sel.c, '');
  } else if (e.key === 'PageDown' || e.key === 'PageUp') {
    e.preventDefault();
    select(sel.r + (e.key === 'PageDown' ? 25 : -25), sel.c);
  } else if (!mod && !e.altKey && e.key.length === 1) {
    e.preventDefault();
    beginEdit(e.key);
  }
}

async function copy(e: ClipboardEvent) {
  if (editing || e.target === formulaInput) return;
  e.preventDefault();
  e.clipboardData?.setData('text/plain', sheet().edit[sel.r][sel.c]);
}

function paste(e: ClipboardEvent) {
  if (editing || e.target === formulaInput) return;
  e.preventDefault();
  const text = e.clipboardData?.getData('text/plain') ?? '';
  const rows = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
  const s = sheet();
  const ops: LocalOp[] = [];
  let structural = false;
  // Grow the sheet when the pasted block doesn't fit.
  const needCols = sel.c + Math.max(...rows.map((r) => r.length)) - width(s);
  for (let i = 0; i < needCols; i++) ops.push({ k: 'insCol', s: active, c: width(s) + i });
  const needRows = sel.r + rows.length - s.edit.length;
  for (let i = 0; i < needRows; i++) ops.push({ k: 'insRow', s: active, r: s.edit.length + i });
  if (ops.length) structural = true;
  rows.forEach((row, dr) => row.forEach((v, dc) => ops.push({ k: 'set', s: active, r: sel.r + dr, c: sel.c + dc, v })));
  perform(ops, 'Paste', structural);
}

// ---------- Structure ----------

const structure = {
  insRow: (below: boolean) => {
    const r = sel.r + (below ? 1 : 0);
    perform([{ k: 'insRow', s: active, r }], 'Insert row', true);
    select(r, sel.c);
  },
  insCol: (right: boolean) => {
    const c = sel.c + (right ? 1 : 0);
    perform([{ k: 'insCol', s: active, c }], 'Insert column', true);
    select(sel.r, c);
  },
  delRow: () => {
    if (sheet().edit.length <= 1) return;
    perform([{ k: 'delRow', s: active, r: sel.r }], 'Delete row', true);
  },
  delCol: () => {
    if (width(sheet()) <= 1) return;
    perform([{ k: 'delCol', s: active, c: sel.c }], 'Delete column', true);
  },
};

function buildUi(fileName: string) {
  addressBox = h('span', { class: 'address' });
  formulaInput = h('input', { class: 'formula-input', spellcheck: 'false', placeholder: 'Cell value or =FORMULA' });
  formulaInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      setCell(sel.r, sel.c, formulaInput.value);
      scroller.focus();
      select(sel.r + 1, sel.c);
    } else if (e.key === 'Escape') {
      updateFormulaBar();
      scroller.focus();
    }
  });
  formulaInput.addEventListener('blur', () => {
    if (formulaInput.value !== sheet().edit[sel.r][sel.c]) setCell(sel.r, sel.c, formulaInput.value);
  });

  const toolbar = h(
    'div',
    { class: 'toolbar' },
    h('span', { class: 'tb-title' }, fileName),
    h('span', { class: 'tb-sep' }),
    button('+ Row ↑', 'Insert row above', () => structure.insRow(false)),
    button('+ Row ↓', 'Insert row below', () => structure.insRow(true)),
    button('+ Col ←', 'Insert column left', () => structure.insCol(false)),
    button('+ Col →', 'Insert column right', () => structure.insCol(true)),
    button('− Row', 'Delete row', structure.delRow, 'danger'),
    button('− Col', 'Delete column', structure.delCol, 'danger'),
  );
  const formulaBar = h('div', { class: 'formula-bar' }, addressBox, h('span', { class: 'fx' }, 'fx'), formulaInput);
  grid = h('table', { class: 'grid' });
  scroller = h('div', { class: 'grid-scroll', tabindex: '0' }, grid);
  tabs = h('div', { class: 'sheet-tabs' });
  app.replaceChildren(h('div', { class: 'table-editor' }, toolbar, formulaBar, scroller, tabs));

  scroller.addEventListener('keydown', onKey);
  scroller.addEventListener('copy', copy);
  scroller.addEventListener('paste', paste);
  scroller.addEventListener('scroll', () => {
    if (scroller.scrollTop + scroller.clientHeight > scroller.scrollHeight - 400 && rendered < sheet().edit.length) {
      renderRows(grid.tBodies[0], rendered, rendered + CHUNK);
      rendered += CHUNK;
    }
  });
  grid.addEventListener('mousedown', (e) => {
    const td = (e.target as HTMLElement).closest('td');
    if (!td || td.querySelector('input')) return;
    const tr = td.parentElement as HTMLTableRowElement;
    commitEdit();
    select(tr.sectionRowIndex, td.cellIndex - 1, false);
    e.preventDefault();
    scroller.focus({ preventScroll: true });
  });
  grid.addEventListener('dblclick', (e) => {
    if ((e.target as HTMLElement).closest('td')) beginEdit();
  });
}

start<TableModel>({
  init(model, fileName) {
    sheets = model.sheets.map((s) => ({
      name: s.name,
      disp: s.rows.map((r) => [...r]),
      edit: s.rows.map((r, ri) => r.map((v, ci) => s.edit?.[`${ri},${ci}`] ?? v)),
    }));
    history.length = 0;
    historyPos = 0;
    log.length = 0;
    active = 0;
    sel = { r: 0, c: 0 };
    rendered = CHUNK;
    buildUi(fileName);
    renderGrid();
    scroller.focus();
  },
  getContent(): TablePayload {
    commitEdit();
    return { ops: log, grids: sheets.map((s) => s.edit) };
  },
  undo() {
    if (historyPos === 0) return;
    const entry = history[--historyPos];
    run(entry.undo);
    renderGrid();
  },
  redo() {
    if (historyPos >= history.length) return;
    const entry = history[historyPos++];
    run(entry.redo);
    renderGrid();
  },
});
