import { useCallback, useEffect, useRef, useState } from 'react';
import { invoke } from '@tauri-apps/api/core';
import { getCurrentWebview } from '@tauri-apps/api/webview';
import { getCurrentWebviewWindow } from '@tauri-apps/api/webviewWindow';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { message, open, save } from '@tauri-apps/plugin-dialog';
import type { Op, StylePatch } from '../core/types';
import { buildRowView, distinctValues, keepRows, viewNeedsUpdate } from '../core/view';
import { DocumentStore } from '../document/store';
import { tauriIo } from '../document/io';
import { CanvasGrid, type CanvasGridHandle, type ResizeEvent } from './canvas/CanvasGrid';
import { ContextMenu, type MenuItem } from './ContextMenu';
import { showError } from './errors';
import { FormulaBar, type FormulaBarHandle } from './FormulaBar';
import type { CellEdit } from './paste';
import type { ContextMenuRequest, Selection } from './selectionTypes';
import { SheetTabs } from './SheetTabs';
import { SidePanel } from './SidePanel';
import { mergeAction, Toolbar, toolbarState } from './Toolbar';
import { argsToOps, shiftFilterCol, structuralArgs, type StructuralAction } from './structure';
import { TopBar } from './TopBar';
import { useStore } from './useStore';
import './app.css';

type Unlisten = () => void;

type SessionSizes = { cols: Map<number, number>; rows: Map<number, number> };
const NO_SIZES: SessionSizes = { cols: new Map(), rows: new Map() };

function fileName(path: string | null): string {
  return path ? (path.split('/').pop() ?? path) : '제목 없음';
}

function isTextInput(el: Element | null): boolean {
  // 캔버스 표의 숨은 입력창은 편집 중이 아닐 때 표 자체로 본다(되돌리기·행 삭제 단축키가 표에 적용되도록)
  if (el instanceof HTMLTextAreaElement) return !('gridIdle' in el.dataset);
  return el instanceof HTMLInputElement && ['text', 'search', 'number', 'email', 'tel', 'url', 'password'].includes(el.type);
}

function openPaths(paths: string[]): void {
  for (const path of paths) {
    invoke('open_document', { path }).catch(showError);
  }
}

/** 저장에 성공하면 true. 새 문서는 위치를 묻고, 취소하면 false. flushEdits는 입력줄·표 편집기에 남은 입력을 확정한다. */
async function saveDocument(store: DocumentStore, flushEdits: () => void): Promise<boolean> {
  if (store.saving) return false;
  try {
    flushEdits();
    if (store.source.format === 'new') {
      const path = await save({
        filters: [
          { name: 'Excel', extensions: ['xlsx'] },
          { name: 'CSV', extensions: ['csv'] },
        ],
      });
      if (!path) return false;
      if (/\.csv$/i.test(path) && store.hasFormatting) {
        const answer = await message('csv에는 서식이 저장되지 않습니다. 계속할까요?', {
          kind: 'warning',
          buttons: { ok: '계속', cancel: '취소' },
        });
        if (answer !== '계속' && answer !== 'Ok') return false;
      }
      await store.save(path);
      await invoke('set_document_path', { path });
    } else {
      await store.save();
    }
    return true;
  } catch (e) {
    showError(e);
    return false;
  }
}

function runEdit(store: DocumentStore, kind: 'undo' | 'redo', after: (ops: Op[] | null, undo: boolean) => void): void {
  if (isTextInput(document.activeElement)) {
    document.execCommand(kind);
    return;
  }
  try {
    after(kind === 'undo' ? store.undo() : store.redo(), kind === 'undo');
  } catch (e) {
    showError(e);
  }
}

interface ViewState {
  query: string;
  filterCol: number;
  off: ReadonlySet<string>; // 필터에서 꺼진 값
  rows: number[] | null; // 검색·필터 결과(원래 행 번호). 변경 시에만 다시 계산한다
}

const INITIAL_VIEW: ViewState = { query: '', filterCol: 0, off: new Set(), rows: null };

function computeRows(store: DocumentStore, sheet: number, v: Omit<ViewState, 'rows'>): number[] | null {
  const model = store.wb.sheets[sheet];
  let filter = null;
  if (v.off.size > 0) {
    // 첫 행은 머리글로 보고 값 목록에서 뺀다
    const all = distinctValues(model, store.wb, v.filterCol, Infinity, 1);
    filter = { col: v.filterCol, values: new Set(all.map((x) => x.value).filter((x) => !v.off.has(x))) };
  }
  const rows = buildRowView(model, store.wb, v.query, filter);
  // 값 필터가 걸려도 머리글 행은 항상 보인다
  if (filter && rows && rows[0] !== 0 && model.rowCount > 0) rows.unshift(0);
  return rows;
}

type FontToggle = 'bold' | 'italic' | 'underline';

const ROW_ACTIONS: [string, StructuralAction][] = [
  ['위에 행 삽입', 'insertAbove'],
  ['아래에 행 삽입', 'insertBelow'],
  ['행 삭제', 'deleteRows'],
];
const COL_ACTIONS: [string, StructuralAction][] = [
  ['왼쪽에 열 삽입', 'insertLeft'],
  ['오른쪽에 열 삽입', 'insertRight'],
  ['열 삭제', 'deleteCols'],
];

async function handleMenu(
  store: DocumentStore,
  flushEdits: () => void,
  id: string,
  structural: (action: StructuralAction) => void,
  after: (ops: Op[] | null, undo: boolean) => void,
  toggleFont: (kind: FontToggle) => void,
): Promise<void> {
  switch (id) {
    case 'new':
      await invoke('open_document', { path: null }).catch(showError);
      break;
    case 'open': {
      const picked = await open({ multiple: true, filters: [{ name: '스프레드시트', extensions: ['xlsx', 'csv'] }] });
      if (picked) openPaths(picked);
      break;
    }
    case 'save':
      await saveDocument(store, flushEdits);
      break;
    case 'close':
      await getCurrentWindow().close();
      break;
    case 'undo':
    case 'redo':
      runEdit(store, id, after);
      break;
    case 'insert':
    case 'delete':
      if (!isTextInput(document.activeElement)) structural(id === 'insert' ? 'insertShortcut' : 'deleteShortcut');
      break;
    case 'bold':
    case 'italic':
    case 'underline':
      if (!isTextInput(document.activeElement)) toggleFont(id);
      break;
  }
}

/** 비동기로 등록되는 리스너를 effect 정리 시점과 관계없이 해제한다. */
function useTauriListener(register: () => Promise<Unlisten>): void {
  const registerRef = useRef(register);
  registerRef.current = register;
  useEffect(() => {
    let disposed = false;
    let unlisten: Unlisten | undefined;
    registerRef.current().then(
      (fn) => {
        if (disposed) fn();
        else unlisten = fn;
      },
      (e) => console.error(e),
    );
    return () => {
      disposed = true;
      unlisten?.();
    };
  }, []);
}

export function App() {
  const [store] = useState(() => new DocumentStore(tauriIo));
  useStore(store);
  const [sheet, setSheet] = useState(0);
  const [selection, setSelection] = useState<Selection | null>(null);
  const [ready, setReady] = useState(false);
  const started = useRef(false);
  const closing = useRef(false);
  const selectionRef = useRef<Selection | null>(null);
  const [view, setView] = useState<ViewState>(INITIAL_VIEW);
  const viewRef = useRef(view);
  viewRef.current = view;
  const [menu, setMenu] = useState<{ x: number; y: number; items: MenuItem[] } | null>(null);
  const [sessionSizes, setSessionSizes] = useState<SessionSizes>(NO_SIZES); // csv 크기 조절(저장하지 않음)
  const gridRef = useRef<CanvasGridHandle>(null);
  const formulaRef = useRef<FormulaBarHandle>(null);
  const flushEdits = useCallback(() => {
    formulaRef.current?.commit();
    gridRef.current?.commitEditor();
  }, []);

  useEffect(() => {
    if (started.current) return; // StrictMode에서 두 번 실행되지 않도록
    started.current = true;
    const params = new URLSearchParams(window.location.search);
    const path = params.get('path');
    if (path) {
      store
        .load(path)
        .catch((e: unknown) => {
          showError(new Error(`파일을 열 수 없습니다: ${e instanceof Error ? e.message : String(e)}`));
          // 빈 문서로 남는 창이므로 Rust 쪽 경로 등록도 지운다
          invoke('clear_document_path').catch(console.error);
          getCurrentWindow().setTitle(fileName(null)).catch(console.error);
        })
        .finally(() => setReady(true));
    } else {
      store.newDocument();
      setReady(true);
    }
  }, [store]);

  const dirty = store.dirty;
  useEffect(() => {
    invoke('set_dirty', { dirty }).catch(console.error);
  }, [dirty]);

  const name = fileName(store.path);
  useEffect(() => {
    // 새 문서 제목은 Rust가 정한다. 경로가 생기거나 바뀌면(새 문서 저장) 맞춘다.
    if (store.path) getCurrentWindow().setTitle(name).catch(console.error);
  }, [store.path, name]);

  useTauriListener(() => getCurrentWebviewWindow().listen<string>('menu', (e) => {
      handleMenu(
        store,
        flushEdits,
        e.payload,
        (a) => runStructural.current(a),
        (ops, undo) => afterHistory.current(ops, undo),
        (kind) => toggleFont.current(kind),
      ).catch(showError);
    }));

  useTauriListener(() =>
    getCurrentWebview().onDragDropEvent((e) => {
      if (e.payload.type !== 'drop') return;
      const paths = e.payload.paths.filter((p) => /\.(xlsx|csv)$/i.test(p));
      if (paths.length === 0) void message('xlsx 또는 csv 파일만 열 수 있습니다.', { kind: 'warning' });
      else openPaths(paths);
    }),
  );

  useTauriListener(() =>
    getCurrentWindow().onCloseRequested(async (event) => {
      if (closing.current) {
        event.preventDefault(); // 이미 묻는 중이다
        return;
      }
      if (store.saving) {
        // 저장이 끝나면 다시 닫기를 요청한다
        event.preventDefault();
        const unsubscribe = store.subscribe(() => {
          if (store.saving) return;
          unsubscribe();
          void getCurrentWindow().close();
        });
        return;
      }
      event.preventDefault();
      closing.current = true;
      try {
        flushEdits(); // 편집 중인 셀도 수정 사항으로 친다
        if (!store.dirty) {
          await getCurrentWindow().destroy();
          return;
        }
        const answer = await message('저장하지 않은 변경 사항이 있습니다. 저장할까요?', {
          buttons: { yes: '저장', no: '저장 안 함', cancel: '취소' },
        });
        // 사용자 지정 버튼이면 누른 버튼의 문구가 돌아온다
        const close = answer === '저장' ? await saveDocument(store, flushEdits) : answer === '저장 안 함';
        if (close) await getCurrentWindow().destroy();
      } finally {
        closing.current = false;
      }
    }),
  );

  const wb = store.wb;
  const current = Math.min(sheet, wb.sheets.length - 1);
  const currentRef = useRef(current);
  currentRef.current = current;

  const changeSheet = useCallback((index: number) => {
    setSheet(index);
    setSelection(null);
    selectionRef.current = null;
    setView(INITIAL_VIEW);
  }, []);

  const updateView = useCallback(
    (patch: Partial<Omit<ViewState, 'rows'>>) => {
      const next = { ...viewRef.current, ...patch };
      setView({ ...next, rows: keepRows(viewRef.current.rows, computeRows(store, currentRef.current, next)) });
    },
    [store],
  );

  const onSelect = useCallback((sel: Selection) => {
    selectionRef.current = sel;
    setSelection(sel);
  }, []);

  // 표의 편집 확정·붙여넣기·잘라내기·지우기(저장 중이면 무시, 한 번이 되돌리기 한 단계)
  const onCommit = useCallback(
    (edits: CellEdit[]) => {
      if (store.saving) return;
      try {
        store.setCells(edits);
      } catch (e) {
        showError(e);
      }
    },
    [store],
  );
  const onWarning = useCallback((text: string) => void message(text, { kind: 'warning' }), []);

  const isCsv = wb.kind === 'csv';
  const onResize = useCallback(
    (e: ResizeEvent) => {
      if (isCsv) {
        // csv는 크기를 저장하지 않으므로 이 창에서만 적용한다
        setSessionSizes((prev) =>
          e.kind === 'col'
            ? { ...prev, cols: new Map([...prev.cols, ...e.cols.map((c): [number, number] => [c, e.px])]) }
            : { ...prev, rows: new Map([...prev.rows, ...e.rows.map((r): [number, number] => [r, e.px])]) },
        );
        return;
      }
      if (store.saving) return;
      try {
        const sheetIndex = currentRef.current;
        const reason = e.kind === 'col' ? store.setColWidth(sheetIndex, e.cols, e.px) : store.setRowHeight(sheetIndex, e.rows, e.px);
        if (reason !== null) void message(reason, { kind: 'warning' });
      } catch (err) {
        showError(err);
      }
    },
    [isCsv, store],
  );

  // ── 서식 ────────────────────────────────────────────────
  const formatDisabled = isCsv || store.saving;

  const applyFormat = useCallback(
    (patch: StylePatch) => {
      const grid = gridRef.current;
      if (!grid || store.saving || store.wb.kind === 'csv') return;
      const area = grid.formatArea();
      if (!area) return;
      try {
        const reason = store.formatCells(currentRef.current, { rows: area.rows, c1: area.c1, c2: area.c2 }, patch);
        if (reason !== null) void message(reason, { kind: 'warning' });
      } catch (e) {
        showError(e);
      }
      grid.focus();
    },
    [store],
  );

  const toggleFont = useRef<(kind: FontToggle) => void>(() => {});
  toggleFont.current = (kind) => {
    const sel = selectionRef.current;
    if (!sel || sel.stale || sel.sheet !== currentRef.current || store.saving || store.wb.kind === 'csv') return;
    const state = toolbarState(store.wb, currentRef.current, sel);
    if (kind === 'bold') applyFormat({ font: { bold: !state.bold } });
    else if (kind === 'italic') applyFormat({ font: { italic: !state.italic } });
    else applyFormat({ font: { underline: state.underline ? undefined : 'single' } });
  };

  const onUnmerge = useCallback(() => {
    const grid = gridRef.current;
    if (!grid || store.saving) return;
    const area = grid.formatArea();
    if (!area) return;
    try {
      const reason = store.unmerge(currentRef.current, area.range);
      if (reason !== null) void message(reason, { kind: 'warning' });
    } catch (e) {
      showError(e);
    }
    grid.focus();
  }, [store]);

  const onMerge = useCallback(async () => {
    const grid = gridRef.current;
    if (!grid || store.saving) return;
    const sheetIndex = currentRef.current;
    const area = grid.formatArea();
    if (!area) return;
    // 선택이 기존 병합과 정확히 같으면 Excel처럼 병합을 푼다
    if (mergeAction(store.wb.sheets[sheetIndex].layout.merges, area.range) === 'unmerge') {
      onUnmerge();
      return;
    }
    try {
      const check = store.mergeCheck(sheetIndex, area.range, area.mergeRows);
      if (!check.ok) {
        void message(check.reason, { kind: 'warning' });
        return;
      }
      if (check.lossy) {
        const answer = await message('병합하면 왼쪽 위 칸의 값만 남습니다. 계속할까요?', {
          kind: 'warning',
          buttons: { ok: '병합', cancel: '취소' },
        });
        if (answer !== '병합' && answer !== 'Ok') return;
      }
      const reason = store.merge(sheetIndex, area.range, area.mergeRows, { align: { horizontal: 'center' } });
      if (reason !== null) void message(reason, { kind: 'warning' });
    } catch (e) {
      showError(e);
    } finally {
      gridRef.current?.focus();
    }
  }, [store, onUnmerge]);

  const focusGrid = useCallback(() => gridRef.current?.focus(), []);

  const onQuery = useCallback(
    (query: string) => {
      if (query !== viewRef.current.query) updateView({ query });
    },
    [updateView],
  );

  const onFilterCol = useCallback((filterCol: number) => updateView({ filterCol, off: new Set() }), [updateView]);

  const onToggle = useCallback(
    (value: string) => {
      const off = new Set(viewRef.current.off);
      if (!off.delete(value)) off.add(value);
      updateView({ off });
    },
    [updateView],
  );

  const onClear = useCallback(() => updateView({ off: new Set() }), [updateView]);

  const runStructural = useRef<(action: StructuralAction) => void>(() => {});
  runStructural.current = (action) => {
    const sel = selectionRef.current;
    if (!sel) return;
    const args = structuralArgs(action, sel);
    if (!args) return;
    const { kind, rows } = args;
    try {
      const reason = store.structural(sel.sheet, kind, rows);
      if (reason !== null) {
        void message(reason, { kind: 'warning' });
        return;
      }
    } catch (e) {
      showError(e);
      return;
    }
    // 원래 행 번호가 바뀌었으므로 다시 계산한다
    updateView(filterPatch(argsToOps(sel.sheet, kind, rows), false));
  };

  // 열이 바뀌면 필터 열을 따라 옮기고, 필터 열이 사라지면 필터를 푼다
  const filterPatch = (ops: Op[], undo: boolean): Partial<Omit<ViewState, 'rows'>> => {
    const filterCol = shiftFilterCol(viewRef.current.filterCol, currentRef.current, ops, undo);
    return filterCol === null ? { filterCol: 0, off: new Set() } : { filterCol };
  };

  const afterHistory = useRef<(ops: Op[] | null, undo: boolean) => void>(() => {});
  afterHistory.current = (ops, undo) => {
    if (!ops || !viewNeedsUpdate(ops)) return; // 편집·서식만 바뀐 경우 보이는 행은 그대로 둔다
    updateView(filterPatch(ops, undo));
  };

  const onContextMenu = useCallback((e: ContextMenuRequest) => {
    selectionRef.current = e.sel;
    const actions = e.target === 'row' ? ROW_ACTIONS : e.target === 'col' ? COL_ACTIONS : [...ROW_ACTIONS, ...COL_ACTIONS];
    setMenu({
      x: e.x,
      y: e.y,
      items: actions.map(([label, action]) => ({ label, onClick: () => runStructural.current(action) })),
    });
  }, []);
  const closeMenu = useCallback(() => setMenu(null), []);

  return (
    <div className="app">
      <TopBar
        name={name}
        rows={wb.sheets[current].rowCount}
        dirty={dirty}
        saving={store.saving}
        onSave={() => void saveDocument(store, flushEdits)}
      />
      <Toolbar
        wb={wb}
        sheet={current}
        anchor={selection && selection.sheet === current && !selection.stale ? selection : null}
        disabled={formatDisabled}
        onFormat={applyFormat}
        onMerge={() => void onMerge()}
        onUnmerge={onUnmerge}
        onDone={focusGrid}
      />
      <FormulaBar ref={formulaRef} store={store} selection={selection} />
      <div className="main">
        {ready && (
          <CanvasGrid
            ref={gridRef}
            key={current}
            store={store}
            sheet={current}
            rowView={view.rows}
            readonly={store.saving}
            sessionSizes={isCsv ? sessionSizes : undefined}
            onSelect={onSelect}
            onContextMenu={onContextMenu}
            onResize={onResize}
            onCommit={onCommit}
            onWarning={onWarning}
          />
        )}
        {ready && (
          <SidePanel
            key={current}
            store={store}
            sheet={current}
            filterCol={view.filterCol}
            off={view.off}
            onQuery={onQuery}
            onFilterCol={onFilterCol}
            onToggle={onToggle}
            onClear={onClear}
          />
        )}
      </div>
      {menu && <ContextMenu x={menu.x} y={menu.y} items={menu.items} onClose={closeMenu} />}
      {wb.kind === 'xlsx' && wb.sheets.length >= 2 && (
        <SheetTabs names={wb.sheets.map((s) => s.name)} active={current} onChange={changeSheet} />
      )}
    </div>
  );
}
