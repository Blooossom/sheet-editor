import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';
import { editText } from '../core/display';
import { getCell } from '../core/ops';
import { toA1 } from '../core/refs';
import type { DocumentStore } from '../document/store';
import { showError } from './errors';
import type { Selection } from './selectionTypes';

interface Props {
  store: DocumentStore;
  selection: Selection | null;
}

export interface FormulaBarHandle {
  /** 입력 중인 값을 확정한다(저장·닫기 전). 고치는 중이 아니면 아무것도 하지 않는다. */
  commit(): void;
}

interface Draft {
  sheet: number;
  row: number;
  col: number;
  text: string;
}

const MAX_ROWS = 3;

/** 입력줄 키가 할 일. 조합 중(isComposing·keyCode 229)에는 'none' — 조합을 확정하는 Enter가 값 확정이 되면 안 된다. */
export function formulaKeyAction(e: { key: string; altKey: boolean; isComposing: boolean; keyCode: number }): 'commit' | 'newline' | 'cancel' | 'none' {
  if (e.isComposing || e.keyCode === 229) return 'none';
  if (e.key === 'Enter') return e.altKey ? 'newline' : 'commit';
  if (e.key === 'Escape') return 'cancel';
  return 'none';
}

/** 입력줄 높이(줄 수). 접혀 있으면 한 줄, 펼치면 내용 줄 수만큼 3줄까지. */
export function formulaRows(text: string, expanded: boolean): number {
  return expanded ? Math.min(MAX_ROWS, text.split('\n').length) : 1;
}

export const FormulaBar = forwardRef<FormulaBarHandle, Props>(function FormulaBar({ store, selection }, ref) {
  const sheet = selection ? store.wb.sheets[selection.sheet] : undefined;
  const original = selection && sheet ? editText(getCell(sheet, selection.row, selection.col), store.wb) : '';
  const [text, setText] = useState(original);
  const [expanded, setExpanded] = useState(false);
  // 고치는 중인 값과 그 셀. 포커스를 잃을 때 선택이 이미 바뀌었어도 원래 셀에 넣기 위해 따로 둔다.
  const draft = useRef<Draft | null>(null);
  const caret = useRef<number | null>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);

  function commit() {
    const d = draft.current;
    draft.current = null;
    if (!d) return;
    try {
      store.setCells([{ sheet: d.sheet, row: d.row, col: d.col, input: d.text }]);
    } catch (e) {
      showError(e);
    }
  }

  const commitRef = useRef(commit);
  commitRef.current = commit;
  useImperativeHandle(ref, () => ({ commit: () => commitRef.current() }), []);

  function change(value: string) {
    setText(value);
    if (selection) draft.current = { sheet: selection.sheet, row: selection.row, col: selection.col, text: value };
  }

  // 선택이 바뀌면 다른 셀의 입력 중 값을 먼저 넣고, 현재 값을 보여 준다
  useEffect(() => {
    const d = draft.current;
    if (d && (!selection || d.sheet !== selection.sheet || d.row !== selection.row || d.col !== selection.col)) {
      commit();
    } else {
      draft.current = null;
    }
    setText(original);
  }, [original, selection?.sheet, selection?.row, selection?.col, store.version]);

  // execCommand가 안 될 때 직접 넣은 줄 바꿈 뒤로 캐럿을 옮긴다(제어 값이 다시 쓰이면 캐럿이 끝으로 간다)
  useEffect(() => {
    const el = inputRef.current;
    if (caret.current === null || !el) return;
    el.setSelectionRange(caret.current, caret.current);
    caret.current = null;
  }, [text]);

  /** ⌥Enter: 캐럿 자리에 줄 바꿈(입력줄의 되돌리기 기록에 남도록 insertText를 먼저 쓴다) */
  function insertNewline(el: HTMLTextAreaElement) {
    if (document.execCommand('insertText', false, '\n')) return;
    const { selectionStart: a, selectionEnd: b, value } = el;
    caret.current = a + 1;
    change(value.slice(0, a) + '\n' + value.slice(b));
  }

  const multiline = text.includes('\n');

  return (
    <div className="formula-bar">
      <span className="cell-address">{selection ? toA1(selection.row, selection.col) : ''}</span>
      <textarea
        ref={inputRef}
        className={`formula-input${expanded && multiline ? ' expanded' : ''}`}
        value={text}
        rows={formulaRows(text, expanded)}
        disabled={!selection || store.saving}
        spellCheck={false}
        autoComplete="off"
        aria-label="입력줄"
        onChange={(e) => change(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          const act = formulaKeyAction({ key: e.key, altKey: e.altKey, isComposing: e.nativeEvent.isComposing, keyCode: e.keyCode });
          if (act === 'none') return;
          e.preventDefault();
          if (act === 'commit') commit();
          else if (act === 'newline') insertNewline(e.currentTarget);
          else {
            draft.current = null;
            setText(original);
          }
        }}
      />
      {multiline && (
        <button
          type="button"
          className="formula-expand"
          aria-label={expanded ? '입력줄 접기' : '입력줄 펼치기'}
          title={expanded ? '접기' : '펼치기'}
          onMouseDown={(e) => e.preventDefault() /* 입력줄 포커스를 유지한다 */}
          onClick={() => setExpanded((v) => !v)}
        >
          {expanded ? '▴' : '▾'}
        </button>
      )}
    </div>
  );
});
