// Shared types between the table webview and the CSV/XLSX handlers.

export interface SheetModel {
  name: string;
  /** Display text, row-major. */
  rows: string[][];
  /** Text shown in the cell editor when it differs from the display text ("r,c" -> text), e.g. formulas. */
  edit?: Record<string, string>;
}

export interface TableModel {
  kind: 'csv' | 'xlsx';
  sheets: SheetModel[];
}

export type TableOp =
  | { k: 'set'; s: number; r: number; c: number; v: string }
  | { k: 'insRow'; s: number; r: number }
  | { k: 'delRow'; s: number; r: number }
  | { k: 'insCol'; s: number; c: number }
  | { k: 'delCol'; s: number; c: number };

export interface TablePayload {
  /** Net operations since the document was opened, in order. */
  ops: TableOp[];
  /** Current edit-text grid per sheet. */
  grids: string[][][];
}
