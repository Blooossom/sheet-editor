import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import DataEditor, {
  CompactSelection,
  GridCellKind,
  type CellClickedEventArgs,
  type EditListItem,
  type GridCell,
  type GridColumn,
  type GridSelection,
  type HeaderClickedEventArgs,
  type Item,
} from '@glideapps/glide-data-grid';
import '@glideapps/glide-data-grid/dist/index.css';
import { message } from '@tauri-apps/plugin-dialog';
import { displayText, editText } from '../core/display';
import { colToLetters } from '../core/refs';
import type { DocumentStore } from '../document/store';

export interface Selection {
  sheet: number;
  row: number; // 원래 행 번호
  col: number;
  rows: number[]; // 머리글로 선택된 원래 행 번호
  cols: number[];
}

export interface ContextMenuRequest {
  x: number;
  y: number;
  target: 'row' | 'col' | 'cell';
  sel: Selection;
}

interface Props {
  store: DocumentStore;
  sheet: number;
  rowView: number[] | null;
  onSelect(sel: Selection): void;
  onContextMenu(e: ContextMenuRequest): void;
}

export function showError(e: unknown): void {
  void message(e instanceof Error ? e.message : String(e), { kind: 'error' });
}

const HEADER_WIDTH = 56;
const COLUMN_WIDTH = 120;
const EDITED_THEME = { bgCell: '#fff6d6' };
const HEADER_THEME = { bgCell: '#f4f4f5', textDark: '#8a8a8e' };
const INITIAL_SELECTION: GridSelection = {
  current: { cell: [1, 0], range: { x: 1, y: 0, width: 1, height: 1 }, rangeStack: [] },
  columns: CompactSelection.empty(),
  rows: CompactSelection.empty(),
};

function toSelection(gs: GridSelection, sheet: number, rowView: number[] | null): Selection {
  const orig = (r: number) => (rowView ? rowView[r] : r);
  const rows = gs.rows.toArray().map(orig).filter((r): r is number => r !== undefined);
  const cols = gs.columns.toArray().filter((c) => c >= 1).map((c) => c - 1);
  let row = orig(0) ?? 0;
  let col = 0;
  if (gs.current) {
    row = orig(gs.current.cell[1]) ?? row;
    col = Math.max(gs.current.cell[0] - 1, 0);
  } else if (rows.length > 0) {
    row = rows[0];
  } else if (cols.length > 0) {
    col = cols[0];
  }
  return { sheet, row, col, rows, cols };
}

// 0번 열(행 머리글)을 누르면 행 선택으로 바꾸고, 0번 열 머리글은 선택하지 않는다
function normalize(gs: GridSelection): GridSelection {
  const range = gs.current?.range;
  if (range && range.x === 0 && range.width === 1) {
    return {
      current: undefined,
      columns: CompactSelection.empty(),
      rows: CompactSelection.fromSingleSelection([range.y, range.y + range.height]),
    };
  }
  return gs.columns.hasIndex(0) ? { ...gs, columns: gs.columns.remove(0) } : gs;
}

export function Grid({ store, sheet, rowView, onSelect, onContextMenu }: Props) {
  const version = store.version;
  const saving = store.saving;
  const model = store.wb.sheets[sheet];
  const sheetCols = Math.max(model.colCount + 5, 26);
  const rowCount = rowView ? rowView.length : Math.max(model.rowCount + 50, 100);

  const [widths, setWidths] = useState<Map<number, number>>(() => new Map());
  const [selection, setSelection] = useState<GridSelection>(INITIAL_SELECTION);
  const selectionRef = useRef(selection);
  selectionRef.current = selection;

  const columns = useMemo<GridColumn[]>(() => {
    const out: GridColumn[] = [{ id: 'row', title: '', width: widths.get(0) ?? HEADER_WIDTH }];
    for (let c = 0; c < sheetCols; c++) {
      out.push({ id: String(c), title: colToLetters(c), width: widths.get(c + 1) ?? COLUMN_WIDTH });
    }
    return out;
  }, [sheetCols, widths]);

  const getCellContent = useCallback(
    ([col, row]: Item): GridCell => {
      const wb = store.wb;
      const orig = rowView ? rowView[row] : row;
      if (col === 0) {
        const label = String(orig + 1);
        return {
          kind: GridCellKind.Text,
          data: label,
          displayData: label,
          allowOverlay: false,
          readonly: true,
          contentAlign: 'center',
          themeOverride: HEADER_THEME,
        };
      }
      const cell = wb.sheets[sheet].cells.get(orig)?.get(col - 1);
      const shown = displayText(cell, wb);
      return {
        kind: GridCellKind.Text,
        data: editText(cell, wb),
        displayData: shown,
        copyData: shown,
        allowOverlay: !saving,
        readonly: saving,
        themeOverride: cell?.edited ? EDITED_THEME : undefined,
      };
    },
    // version이 바뀌면 함수가 바뀌어 표가 다시 그려진다
    [store, sheet, rowView, version, saving],
  );

  const select = useCallback(
    (gs: GridSelection) => {
      setSelection(gs);
      onSelect(toSelection(gs, sheet, rowView));
    },
    [onSelect, sheet, rowView],
  );

  useEffect(() => {
    onSelect(toSelection(selectionRef.current, sheet, rowView));
  }, [onSelect, sheet, rowView]);

  const onCellsEdited = useCallback(
    (items: readonly EditListItem[]) => {
      if (store.saving) return true;
      const edits = [];
      for (const { location, value } of items) {
        const [col, row] = location;
        const orig = rowView ? rowView[row] : row;
        if (col === 0 || orig === undefined || value.kind !== GridCellKind.Text) continue;
        edits.push({ sheet, row: orig, col: col - 1, input: value.data });
      }
      try {
        store.setCells(edits);
      } catch (e) {
        showError(e);
      }
      return true;
    },
    [store, sheet, rowView],
  );

  const onCellContextMenu = useCallback(
    ([col, row]: Item, e: CellClickedEventArgs) => {
      e.preventDefault();
      let gs = selectionRef.current;
      if (col === 0) {
        if (!gs.rows.hasIndex(row)) {
          gs = { current: undefined, columns: CompactSelection.empty(), rows: CompactSelection.fromSingleSelection(row) };
          select(gs);
        }
      } else {
        const r = gs.current?.range;
        const inside = r && col >= r.x && col < r.x + r.width && row >= r.y && row < r.y + r.height;
        if (!inside) {
          gs = {
            current: { cell: [col, row], range: { x: col, y: row, width: 1, height: 1 }, rangeStack: [] },
            columns: CompactSelection.empty(),
            rows: CompactSelection.empty(),
          };
          select(gs);
        }
      }
      onContextMenu({
        x: e.bounds.x + e.localEventX,
        y: e.bounds.y + e.localEventY,
        target: col === 0 ? 'row' : 'cell',
        sel: toSelection(gs, sheet, rowView),
      });
    },
    [select, onContextMenu, sheet, rowView],
  );

  const onHeaderContextMenu = useCallback(
    (col: number, e: HeaderClickedEventArgs) => {
      e.preventDefault();
      if (col === 0) return;
      let gs = selectionRef.current;
      if (!gs.columns.hasIndex(col)) {
        gs = { current: undefined, columns: CompactSelection.fromSingleSelection(col), rows: CompactSelection.empty() };
        select(gs);
      }
      onContextMenu({
        x: e.bounds.x + e.localEventX,
        y: e.bounds.y + e.localEventY,
        target: 'col',
        sel: toSelection(gs, sheet, rowView),
      });
    },
    [select, onContextMenu, sheet, rowView],
  );

  const onColumnResize = useCallback((_column: GridColumn, size: number, index: number) => {
    setWidths((prev) => new Map(prev).set(index, size));
  }, []);

  return (
    <div className="grid">
      <DataEditor
        width="100%"
        height="100%"
        columns={columns}
        rows={rowCount}
        getCellContent={getCellContent}
        getCellsForSelection={true}
        onCellsEdited={onCellsEdited}
        gridSelection={selection}
        onGridSelectionChange={(gs) => select(normalize(gs))}
        onCellContextMenu={onCellContextMenu}
        onHeaderContextMenu={onHeaderContextMenu}
        onColumnResize={onColumnResize}
        freezeColumns={1}
        rowMarkers="none"
        smoothScrollX
        smoothScrollY
      />
    </div>
  );
}
