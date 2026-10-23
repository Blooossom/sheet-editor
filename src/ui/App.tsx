import { useCallback, useEffect, useRef, useState } from 'react';
import { invoke } from '@tauri-apps/api/core';
import { getCurrentWebview } from '@tauri-apps/api/webview';
import { getCurrentWebviewWindow } from '@tauri-apps/api/webviewWindow';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { message, open, save } from '@tauri-apps/plugin-dialog';
import { DocumentStore } from '../document/store';
import { tauriIo } from '../document/io';
import { FormulaBar } from './FormulaBar';
import { Grid, showError, type ContextMenuRequest, type Selection } from './Grid';
import { SheetTabs } from './SheetTabs';
import { TopBar } from './TopBar';
import { useStore } from './useStore';
import './app.css';

type Unlisten = () => void;

function fileName(path: string | null): string {
  return path ? (path.split('/').pop() ?? path) : '제목 없음';
}

function isTextInput(el: Element | null): boolean {
  return el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement;
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

function runEdit(store: DocumentStore, kind: 'undo' | 'redo'): void {
  if (isTextInput(document.activeElement)) {
    document.execCommand(kind);
    return;
  }
  try {
    if (kind === 'undo') store.undo();
    else store.redo();
  } catch (e) {
    showError(e);
  }
}

async function handleMenu(store: DocumentStore, id: string): Promise<void> {
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
      runEdit(store, id);
      break;
    // 'insert' / 'delete'는 행·열 메뉴(Task 15)에서 처리한다
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
      handleMenu(store, e.payload).catch(showError);
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

  const changeSheet = useCallback((index: number) => {
    setSheet(index);
    setSelection(null);
  }, []);

  const onContextMenu = useCallback((_e: ContextMenuRequest) => {
    // 행·열 메뉴는 Task 15에서 붙인다
  }, []);

  const wb = store.wb;
  const current = Math.min(sheet, wb.sheets.length - 1);

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
            rowView={null}
            onSelect={setSelection}
            onContextMenu={onContextMenu}
          />
        )}
        {/* Task 15: SidePanel */}
      </div>
      {wb.kind === 'xlsx' && wb.sheets.length >= 2 && (
        <SheetTabs names={wb.sheets.map((s) => s.name)} active={current} onChange={changeSheet} />
      )}
    </div>
  );
}
