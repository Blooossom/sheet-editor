import { useEffect, useRef, useState } from 'react';
import { editText } from '../core/display';
import { getCell } from '../core/ops';
import { toA1 } from '../core/refs';
import type { DocumentStore } from '../document/store';
import { showError, type Selection } from './Grid';

interface Props {
  store: DocumentStore;
  selection: Selection | null;
}

interface Draft {
  sheet: number;
  row: number;
  col: number;
  text: string;
}

export function FormulaBar({ store, selection }: Props) {
  const sheet = selection ? store.wb.sheets[selection.sheet] : undefined;
  const original = selection && sheet ? editText(getCell(sheet, selection.row, selection.col), store.wb) : '';
  const [text, setText] = useState(original);
  // 고치는 중인 값과 그 셀. 포커스를 잃을 때 선택이 이미 바뀌었어도 원래 셀에 넣기 위해 따로 둔다.
  const draft = useRef<Draft | null>(null);

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

  return (
    <div className="formula-bar">
      <span className="cell-address">{selection ? toA1(selection.row, selection.col) : ''}</span>
      <input
        className="formula-input"
        value={text}
        disabled={!selection || store.saving}
        spellCheck={false}
        autoComplete="off"
        onChange={(e) => {
          setText(e.target.value);
          if (selection) draft.current = { sheet: selection.sheet, row: selection.row, col: selection.col, text: e.target.value };
        }}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.nativeEvent.isComposing) return; // 한글 조합 중 Enter는 조합 확정용
          if (e.key === 'Enter') {
            e.preventDefault();
            commit();
          } else if (e.key === 'Escape') {
            e.preventDefault();
            draft.current = null;
            setText(original);
          }
        }}
      />
    </div>
  );
}
