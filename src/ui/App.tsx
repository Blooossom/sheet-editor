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

/** 저장에 성공하면 true. 새 문서는 위치를 묻고, 취소하면 false. */
async function saveDocument(store: DocumentStore): Promise<boolean> {
  if (store.saving) return false;
  try {
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
  const started = useRef(false);

  useEffect(() => {
    if (started.current) return; // StrictMode에서 두 번 실행되지 않도록
    started.current = true;
    const params = new URLSearchParams(window.location.search);
    const path = params.get('path');
    if (path) {
      store.load(path).catch((e: unknown) => {
        showError(new Error(`파일을 열 수 없습니다: ${e instanceof Error ? e.message : String(e)}`));
      });
    } else {
      store.newDocument();
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

  useTauriListener(() => getCurrentWebviewWindow().listen<string>('menu', (e) => void handleMenu(store, e.payload)));

  useTauriListener(() =>
    getCurrentWebview().onDragDropEvent((e) => {
      if (e.payload.type !== 'drop') return;
      openPaths(e.payload.paths.filter((p) => /\.(xlsx|csv)$/i.test(p)));
    }),
  );

  useTauriListener(() =>
    getCurrentWindow().onCloseRequested(async (event) => {
      if (!store.dirty && !store.saving) return;
      event.preventDefault();
      if (store.saving) return;
      const answer = await message('저장하지 않은 변경 사항이 있습니다. 저장할까요?', {
        buttons: { yes: '저장', no: '저장 안 함', cancel: '취소' },
      });
      // 사용자 지정 버튼이면 누른 버튼의 문구가 돌아온다
      const close = answer === '저장' ? await saveDocument(store) : answer === '저장 안 함';
      if (close) await getCurrentWindow().destroy();
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
        <Grid
          key={current}
          store={store}
          sheet={current}
          rowView={null}
          onSelect={setSelection}
          onContextMenu={onContextMenu}
        />
        {/* Task 15: SidePanel */}
      </div>
      {wb.kind === 'xlsx' && wb.sheets.length >= 2 && (
        <SheetTabs names={wb.sheets.map((s) => s.name)} active={current} onChange={changeSheet} />
      )}
    </div>
  );
}
