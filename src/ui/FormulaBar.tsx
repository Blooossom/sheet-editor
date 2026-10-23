import { useEffect, useState } from 'react';
import { editText } from '../core/display';
import { getCell } from '../core/ops';
import { toA1 } from '../core/refs';
import type { DocumentStore } from '../document/store';
import { showError, type Selection } from './Grid';

interface Props {
  store: DocumentStore;
  selection: Selection | null;
}

export function FormulaBar({ store, selection }: Props) {
  const sheet = selection ? store.wb.sheets[selection.sheet] : undefined;
  const original = selection && sheet ? editText(getCell(sheet, selection.row, selection.col), store.wb) : '';
  const [text, setText] = useState(original);

  // 선택이 바뀌거나 문서가 바뀌면 입력 중이던 내용을 버리고 현재 값을 보여 준다
  useEffect(() => {
    setText(original);
  }, [original, selection?.sheet, selection?.row, selection?.col, store.version]);

  function commit() {
    if (!selection) return;
    try {
      store.setCells([{ sheet: selection.sheet, row: selection.row, col: selection.col, input: text }]);
    } catch (e) {
      setText(original);
      showError(e);
    }
  }

  return (
    <div className="formula-bar">
      <span className="cell-address">{selection ? toA1(selection.row, selection.col) : ''}</span>
      <input
        className="formula-input"
        value={text}
        disabled={!selection || store.saving}
        spellCheck={false}
        autoComplete="off"
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.nativeEvent.isComposing) return; // 한글 조합 중 Enter는 조합 확정용
          if (e.key === 'Enter') {
            e.preventDefault();
            commit();
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setText(original);
          }
        }}
      />
    </div>
  );
}
