import type { Cell, CellValue, SheetModel, WorkbookModel } from '../../src/core/types';

type GridEntry = string | number | null | { f: string; v?: CellValue };

export function makeWorkbook(
  kind: 'xlsx' | 'csv',
  sheets: { name: string; grid: GridEntry[][] }[],
): WorkbookModel {
  return {
    kind,
    dateStyleIds: new Set<number>(),
    date1904: false,
    sheets: sheets.map(({ name, grid }): SheetModel => {
      const cells = new Map<number, Map<number, Cell>>();
      let rowCount = 0;
      let colCount = 0;
      grid.forEach((rowEntries, r) => {
        rowEntries.forEach((entry, c) => {
          if (entry === null) return;
          let cell: Cell;
          if (typeof entry === 'number') cell = { value: entry };
          else if (typeof entry === 'string') cell = { value: entry, text: true };
          else cell = { formula: entry.f, value: entry.v ?? null };
          let row = cells.get(r);
          if (!row) cells.set(r, (row = new Map()));
          row.set(c, cell);
          rowCount = Math.max(rowCount, r + 1);
          colCount = Math.max(colCount, c + 1);
        });
      });
      return { name, cells, rowCount, colCount, blockers: [], tables: [] };
    }),
  };
}
