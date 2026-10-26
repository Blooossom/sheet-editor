import { parseA1 } from '../core/refs';
import type { CellRange, ColInfo, RowInfo, SheetLayout } from '../core/types';

export function emptyLayout(): SheetLayout {
  return {
    colRuns: [],
    rows: new Map(),
    baseColWidth: 8,
    defaultRowHeight: 15,
    merges: [],
    freeze: { rows: 0, cols: 0 },
    showGridLines: true,
  };
}

const ATTR = /([\w:]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

function attrsOf(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  ATTR.lastIndex = 0;
  for (let m = ATTR.exec(tag); m; m = ATTR.exec(tag)) out[m[1]] = m[2] ?? m[3];
  return out;
}

const isTrue = (v: string | undefined) => v === '1' || v === 'true';

function num(v: string | undefined): number | undefined {
  if (v === undefined || v === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

// 5만 행 시트에서도 빠르도록 DOM 대신 태그 단위 정규식으로 읽는다.
export function readLayout(sheetXml: string): SheetLayout {
  const L = emptyLayout();

  const view = /<sheetView\b[^>]*>/.exec(sheetXml);
  if (view) {
    const grid = attrsOf(view[0]).showGridLines;
    if (grid !== undefined) L.showGridLines = isTrue(grid);
  }
  const pane = /<pane\b[^>]*>/.exec(sheetXml);
  if (pane) {
    const a = attrsOf(pane[0]);
    if (a.state === 'frozen' || a.state === 'frozenSplit') {
      L.freeze = { rows: Math.max(0, num(a.ySplit) ?? 0), cols: Math.max(0, num(a.xSplit) ?? 0) };
    }
  }

  const fmt = /<sheetFormatPr\b[^>]*>/.exec(sheetXml);
  if (fmt) {
    const a = attrsOf(fmt[0]);
    L.baseColWidth = num(a.baseColWidth) ?? 8;
    L.defaultColWidth = num(a.defaultColWidth);
    L.defaultRowHeight = num(a.defaultRowHeight) ?? 15;
  }

  const cols = /<cols\b[^>]*>([\s\S]*?)<\/cols>/.exec(sheetXml);
  if (cols) {
    const runs: SheetLayout['colRuns'] = [];
    for (const m of cols[1].matchAll(/<col\b[^>]*>/g)) {
      const a = attrsOf(m[0]);
      const min = num(a.min);
      const max = num(a.max);
      if (min === undefined || max === undefined || min < 1 || max < min) continue;
      const info: ColInfo = {};
      const width = num(a.width);
      if (width !== undefined) info.width = width;
      if (isTrue(a.hidden)) info.hidden = true;
      const style = num(a.style);
      if (style !== undefined) info.style = style;
      runs.push({ c1: min - 1, c2: max - 1, info });
    }
    runs.sort((x, y) => x.c1 - y.c1);
    let end = -1;
    for (const run of runs) {
      if (run.c1 <= end) run.c1 = end + 1; // 겹치면 앞 구간이 우선
      if (run.c1 > run.c2) continue;
      L.colRuns.push(run);
      end = run.c2;
    }
  }

  let prev = -1;
  for (const m of sheetXml.matchAll(/<row\b[^>]*>/g)) {
    const a = attrsOf(m[0]);
    const r = num(a.r);
    const row = r !== undefined ? r - 1 : prev + 1;
    prev = row;
    const info: RowInfo = {};
    const ht = num(a.ht);
    if (ht !== undefined) info.height = ht;
    if (isTrue(a.hidden)) info.hidden = true;
    const s = num(a.s);
    if (s !== undefined && isTrue(a.customFormat)) info.style = s;
    if (Object.keys(info).length > 0) L.rows.set(row, info);
  }

  for (const m of sheetXml.matchAll(/<mergeCell\b[^>]*>/g)) {
    const ref = attrsOf(m[0]).ref;
    if (!ref) continue;
    const [a, b = a] = ref.split(':');
    const p = parseA1(a);
    const q = parseA1(b);
    const range: CellRange = {
      r1: Math.min(p.row, q.row),
      c1: Math.min(p.col, q.col),
      r2: Math.max(p.row, q.row),
      c2: Math.max(p.col, q.col),
    };
    L.merges.push(range);
  }
  return L;
}

export function colInfo(layout: SheetLayout, col: number): ColInfo {
  const runs = layout.colRuns;
  let lo = 0;
  let hi = runs.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (runs[mid].c2 < col) lo = mid + 1;
    else if (runs[mid].c1 > col) hi = mid - 1;
    else return runs[mid].info;
  }
  return {};
}
