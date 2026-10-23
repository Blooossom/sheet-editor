import { useEffect, useMemo, useState } from 'react';
import { distinctValues } from '../core/view';
import { colToLetters } from '../core/refs';
import type { DocumentStore } from '../document/store';

const LIMIT = 500;

interface Props {
  store: DocumentStore;
  sheet: number;
  filterCol: number;
  off: ReadonlySet<string>; // 꺼진(숨길) 값
  onQuery(query: string): void;
  onFilterCol(col: number): void;
  onToggle(value: string): void;
  onClear(): void;
}

export function SidePanel({ store, sheet, filterCol, off, onQuery, onFilterCol, onToggle, onClear }: Props) {
  const [text, setText] = useState('');
  useEffect(() => {
    const timer = setTimeout(() => onQuery(text), 200);
    return () => clearTimeout(timer);
  }, [text, onQuery]);

  const model = store.wb.sheets[sheet];
  const version = store.version;
  const values = useMemo(
    () => distinctValues(model, store.wb, filterCol, LIMIT + 1, 1),
    // 편집·구조 변경 뒤에도 값 목록은 최신으로 맞춘다
    [model, store, filterCol, version],
  );
  const shown = values.slice(0, LIMIT);
  const colCount = Math.max(model.colCount, filterCol + 1, 1);

  return (
    <aside className="side-panel">
      <label className="panel-label">검색</label>
      <input className="panel-input" value={text} onChange={(e) => setText(e.target.value)} placeholder="검색" />
      <label className="panel-label">필터</label>
      <select className="panel-input" value={filterCol} onChange={(e) => onFilterCol(Number(e.target.value))}>
        {Array.from({ length: colCount }, (_, c) => (
          <option key={c} value={c}>
            {colToLetters(c)}
          </option>
        ))}
      </select>
      <button className="panel-button" onClick={onClear} disabled={off.size === 0}>
        필터 해제
      </button>
      <div className="filter-list">
        {shown.map(({ value, count }) => (
          <label key={value} className="filter-item">
            <input type="checkbox" checked={!off.has(value)} onChange={() => onToggle(value)} />
            <span className="filter-value">{value === '' ? '(빈 칸)' : value}</span>
            <span className="filter-count">{count}</span>
          </label>
        ))}
      </div>
      <div className="panel-note">첫 행(머리글)은 항상 표시됩니다</div>
      {values.length > LIMIT && <div className="panel-note">상위 500개만 표시</div>}
    </aside>
  );
}
