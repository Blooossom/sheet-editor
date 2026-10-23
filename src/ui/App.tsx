import { useCallback, useEffect, useRef, useState } from 'react';
import { invoke } from '@tauri-apps/api/core';
import { getCurrentWebview } from '@tauri-apps/api/webview';
import { getCurrentWebviewWindow } from '@tauri-apps/api/webviewWindow';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { message, open, save } from '@tauri-apps/plugin-dialog';
import type { Op } from '../core/types';
import { buildRowView, distinctValues } from '../core/view';
import { DocumentStore } from '../document/store';
import { tauriIo } from '../document/io';
import { ContextMenu, type MenuItem } from './ContextMenu';
import { FormulaBar } from './FormulaBar';
import { Grid, showError, type ContextMenuRequest, type Selection } from './Grid';
import { SheetTabs } from './SheetTabs';
import { SidePanel } from './SidePanel';
import { argsToOps, shiftFilterCol, structuralArgs, type StructuralAction } from './structure';
import { TopBar } from './TopBar';
import { useStore } from './useStore';
import './app.css';

type Unlisten = () => void;

function fileName(path: string | null): string {
  return path ? (path.split('/').pop() ?? path) : '제목 없음';
}

function isTextInput(el: Element | null): boolean {
  if (el instanceof HTMLTextAreaElement) return true;
  return el instanceof HTMLInputElement && ['text', 'search', 'number', 'email', 'tel', 'url', 'password'].includes(el.type);
}

function openPaths(paths: string[]): void {
  for (const path of paths) {
    invoke('open_document', { path }).catch(showError);
  }
}

/** 입력줄과 표의 편집기에 남은 입력을 확정한다. 표 편집기는 바깥 클릭으로만 확정되므로 클릭을 흉내 낸다. */
async function flushEdits(): Promise<void> {
  const active = document.activeElement;
  if (active instanceof HTMLElement) active.blur();
  document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

/** 저장에 성공하면 true. 새 문서는 위치를 묻고, 취소하면 false. */
async function saveDocument(store: DocumentStore): Promise<boolean> {
  if (store.saving) return false;
  try {
    await flushEdits();
    if (store.source.format === 'new') {
      const path = await save({
        filters: [
          { name: 'Excel', extensions: ['xlsx'] },
          { name: 'CSV', extensions: ['csv'] },
        ],
      });
      if (!path) return false;
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
  id: string,
  structural: (action: StructuralAction) => void,
  after: (ops: Op[] | null, undo: boolean) => void,
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
      await saveDocument(store);
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
      handleMenu(store, e.payload, (a) => runStructural.current(a), (ops, undo) => afterHistory.current(ops, undo)).catch(showError);
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
        await flushEdits(); // 편집 중인 셀도 수정 사항으로 친다
        if (!store.dirty) {
          await getCurrentWindow().destroy();
          return;
        }
        const answer = await message('저장하지 않은 변경 사항이 있습니다. 저장할까요?', {
          buttons: { yes: '저장', no: '저장 안 함', cancel: '취소' },
        });
        // 사용자 지정 버튼이면 누른 버튼의 문구가 돌아온다
        const close = answer === '저장' ? await saveDocument(store) : answer === '저장 안 함';
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
      setView({ ...next, rows: computeRows(store, currentRef.current, next) });
    },
    [store],
  );

  const onSelect = useCallback((sel: Selection) => {
    selectionRef.current = sel;
    setSelection(sel);
  }, []);

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
    if (!ops?.some((o) => o.kind !== 'set')) return; // 편집만 바뀐 경우 보이는 행은 그대로 둔다
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
        onSave={() => void saveDocument(store)}
      />
      <FormulaBar store={store} selection={selection} />
      <div className="main">
        {ready && (
          <Grid
            key={current}
            store={store}
            sheet={current}
            rowView={view.rows}
            onSelect={onSelect}
            onContextMenu={onContextMenu}
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
