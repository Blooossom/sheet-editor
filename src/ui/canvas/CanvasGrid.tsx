import { forwardRef, useCallback, useEffect, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { editText } from '../../core/display';
import type { DocumentStore } from '../../document/store';
import { SheetGeometry } from '../../grid/geometry';
import { colEdgeX, hitTest, pointToCell, rowEdgeY } from '../../grid/hit';
import { contextTarget, gridKeyAction, resizeTargets } from '../../grid/interaction';
import { colWidthPx, defaultColWidthPx, maxDigitWidth, rowHeightPx } from '../../grid/metrics';
import { effectiveXf, StyleResolver } from '../../grid/resolve';
import { expandForMerges, formatArea, move, toSelection, type GridSel, type MoveKey } from '../../grid/selection';
import { makeMeasure } from '../../grid/textLayout';
import { colInfo } from '../../xlsx/layout';
import { pasteEdits, type CellEdit } from '../paste';
import type { ContextMenuRequest, FormatArea, Selection } from '../selectionTypes';
import { useStore } from '../useStore';
import { CellEditor, editorBox, editorKeyAction, idleInputAction, type EditorKeyAction } from './CellEditor';
import { clearEdits, copyText, parseTsv } from './clipboard';
import { drawCells, horizontalOf, isPlainSheet, lineHeightOf, PAD_X, PAD_Y, PLAIN_FONT, type DrawView } from './draw';
import './canvas.css';

export type ResizeEvent =
  | { kind: 'col'; cols: number[]; px: number }
  | { kind: 'row'; rows: number[] /* 원래 행 */; px: number };

export interface CanvasGridProps {
  store: DocumentStore;
  sheet: number;
  rowView: number[] | null;
  readonly: boolean;
  sessionSizes?: { cols: Map<number, number>; rows: Map<number, number> }; // csv용 px 덮어쓰기
  onSelect(sel: Selection): void;
  onContextMenu(e: ContextMenuRequest): void;
  onResize?(e: ResizeEvent): void;
  /** 편집기 확정·붙여넣기·잘라내기로 바뀐 칸(원래 행). 한 번 호출이 되돌리기 한 단계다. */
  onCommit(edits: CellEdit[]): void;
  /** 붙여넣기를 막았거나 일부를 버렸을 때 사용자에게 보여 줄 문구 */
  onWarning(message: string): void;
}

export interface CanvasGridHandle {
  /** 편집 중인 칸을 확정한다(저장·닫기 전). 편집 중이 아니면 아무것도 하지 않는다. */
  commitEditor(): void;
  /** 지금 선택의 서식·병합 대상 */
  formatArea(): FormatArea;
  /** 표에 키보드 포커스를 돌려준다(도구 막대 사용 뒤) */
  focus(): void;
}

/** 편집 중인 칸. row는 화면 행, orig는 원래 행(병합이면 앵커 칸) */
interface Editing {
  row: number;
  col: number;
  orig: number;
  original: string; // 편집을 시작할 때의 편집 문자열(바뀌지 않았으면 확정하지 않는다)
}

const COMMIT_MOVE: Partial<Record<EditorKeyAction, MoveKey>> = {
  commitDown: 'Enter', commitUp: 'ShiftEnter', commitRight: 'Tab', commitLeft: 'ShiftTab',
};
const CARET_SLACK = 2; // 글자 끝 캐럿 자리

const PLAIN_COL = 120;
const PLAIN_ROW = 24;
const HEADER_H = 20;
const EXTRA_ROWS = 50;
const EXTRA_COLS = 10;
const MIN_ROWS = 100;
const MIN_COLS = 26;
const MIN_SIZE = 4; // 크기 조절 최소 px
const AUTOSCROLL_MAX = 40; // 프레임당 최대 px

const INITIAL_SEL: GridSel = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 }, mode: 'cells' };

type Drag =
  | { kind: 'select'; mode: GridSel['mode'] }
  | { kind: 'resize'; axis: 'col' | 'row'; index: number; start: number; size0: number; edge0: number; px: number };

const sameSel = (a: GridSel, b: GridSel) =>
  a.mode === b.mode && a.anchor.row === b.anchor.row && a.anchor.col === b.anchor.col && a.focus.row === b.focus.row && a.focus.col === b.focus.col;

export const CanvasGrid = forwardRef<CanvasGridHandle, CanvasGridProps>(function CanvasGrid(
  { store, sheet, rowView, readonly, sessionSizes, onSelect, onContextMenu, onResize, onCommit, onWarning },
  ref,
) {
  const version = useStore(store);
  const wb = store.wb;
  const model = wb.sheets[sheet];
  const scrollerRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });
  const [sel, setSel] = useState<GridSel>(INITIAL_SEL);
  const [editing, setEditing] = useState<Editing | null>(null);
  const editingRef = useRef<Editing | null>(null);
  const composingRef = useRef(false); // 입력창에서 한글 등 조합이 진행 중
  const blurQuiet = useRef(false);

  const plain = isPlainSheet(wb);
  const mdw = useMemo(() => (plain ? 7 : maxDigitWidth(wb.styles.fonts[0])), [plain, wb.styles]);
  const resolver = useMemo(() => new StyleResolver(wb), [wb]);

  const rowCount = rowView ? rowView.length : Math.max(model.rowCount + EXTRA_ROWS, MIN_ROWS);
  const colCount = Math.max(model.colCount + EXTRA_COLS, MIN_COLS);
  const lastRow = Math.max(0, rowCount - 1);
  const lastCol = colCount - 1;

  const geometry = useMemo(() => {
    const layout = model.layout;
    const defaultCol = defaultColWidthPx(layout, mdw);
    return new SheetGeometry({
      rowCount,
      colCount,
      rowView,
      rowHeight: (r) => {
        const info = layout.rows.get(r);
        if (info?.hidden) return 0;
        const session = sessionSizes?.rows.get(r);
        if (session !== undefined) return session;
        if (plain) return PLAIN_ROW;
        return rowHeightPx(info?.height ?? layout.defaultRowHeight);
      },
      colWidth: (c) => {
        const info = colInfo(layout, c);
        if (info.hidden) return 0;
        const session = sessionSizes?.cols.get(c);
        if (session !== undefined) return session;
        if (plain) return PLAIN_COL;
        return info.width !== undefined ? colWidthPx(info.width, mdw) : defaultCol;
      },
      merges: layout.merges,
      freeze: { rows: rowView ? 0 : Math.min(layout.freeze.rows, rowCount), cols: Math.min(layout.freeze.cols, colCount) },
    });
    // version: 같은 시트 객체 안의 값·크기가 바뀌어도 다시 계산한다
  }, [model, rowView, rowCount, colCount, plain, mdw, sessionSizes, version]);

  const headerW = useMemo(() => {
    const maxLabel = rowView ? rowView.reduce((m, r) => Math.max(m, r + 1), 1) : Math.max(model.rowCount + EXTRA_ROWS, MIN_ROWS);
    return Math.max(32, String(maxLabel).length * 7 + 14);
  }, [rowView, model.rowCount]);

  // 이벤트 처리기가 항상 최신 값을 보도록 둔다
  const live = useRef({ geometry, sel, size, headerW, lastRow, lastCol, rowView });
  live.current = { geometry, sel, size, headerW, lastRow, lastCol, rowView };

  // 글꼴 문자열마다 측정기 하나(프레임 사이에 재사용)
  const measures = useRef(new Map<string, (s: string) => number>());
  const guide = useRef<DrawView['guide']>(null);

  const drawRef = useRef<() => void>(() => {});
  drawRef.current = () => {
    const canvas = canvasRef.current;
    const scroller = scrollerRef.current;
    if (!canvas || !scroller || size.w === 0 || size.h === 0) return;
    const dpr = window.devicePixelRatio || 1;
    const pw = Math.round(size.w * dpr);
    const ph = Math.round(size.h * dpr);
    if (canvas.width !== pw || canvas.height !== ph) {
      canvas.width = pw;
      canvas.height = ph;
    }
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const scroll = { x: scroller.scrollLeft, y: scroller.scrollTop };
    const fw = geometry.colX(geometry.freeze.cols);
    const fh = geometry.rowY(geometry.freeze.rows);
    const rect = expandForMerges(sel, geometry);
    const view: DrawView = {
      geometry,
      resolver,
      wb,
      sheet,
      visible: geometry.visible(fw + scroll.x, fh + scroll.y, Math.max(0, size.w - headerW - fw), Math.max(0, size.h - HEADER_H - fh)),
      scroll,
      dpr,
      selection: { ...rect, row: sel.anchor.row, col: sel.anchor.col },
      edited: (r, c) => !!model.cells.get(r)?.get(c)?.edited,
      showGridLines: model.layout.showGridLines,
      rowView,
      size,
      header: { w: headerW, h: HEADER_H },
      mdw,
      plain,
      measure: (font) => {
        let m = measures.current.get(font);
        if (!m) measures.current.set(font, (m = makeMeasure(ctx, font)));
        return m;
      },
      guide: guide.current,
    };
    drawCells(ctx, view);
    placeRef.current();
  };

  /**
   * 입력창을 칸 위에 놓는다. 편집 중이 아니면 선택 칸 위에 보이지 않게(한글 조합 후보 창이 칸 옆에 뜨도록),
   * 편집 중이면 칸(병합 전체)의 글꼴·정렬·색으로 보이게 두고 내용에 맞춰 늘린다.
   */
  const placeRef = useRef<() => void>(() => {});
  placeRef.current = () => {
    const el = inputRef.current;
    const scroller = scrollerRef.current;
    if (!el || !scroller) return;
    const { geometry: g, sel: s, size: sz, headerW: hw } = live.current;
    const ed = editingRef.current;
    const at = ed ?? anchorOf(s.anchor.row, s.anchor.col);
    const r = g.cellRect(at.row, at.col);
    const x = hw + r.x - (at.col >= g.freeze.cols ? scroller.scrollLeft : 0);
    const y = HEADER_H + r.y - (at.row >= g.freeze.rows ? scroller.scrollTop : 0);
    const st = el.style;
    if (!ed) {
      // 화면 밖에 두면 입력 때 브라우저가 보이게 하려고 스크롤할 수 있으므로 화면 안으로 당긴다
      st.left = `${Math.min(Math.max(x, hw), Math.max(hw, sz.w - 1))}px`;
      st.top = `${Math.min(Math.max(y, HEADER_H), Math.max(HEADER_H, sz.h - 1))}px`;
      st.width = `${Math.max(1, r.w)}px`;
      st.height = `${Math.max(1, r.h)}px`;
      return;
    }
    const style = resolver.get(effectiveXf(model, ed.orig, ed.col));
    const cell = model.cells.get(ed.orig)?.get(ed.col);
    const font = plain ? PLAIN_FONT : style.font;
    const lineH = lineHeightOf(font);
    const vAlign = plain ? 'center' : style.vAlign;
    const wrap = !plain && style.wrap;
    const free = r.h - el.value.split('\n').length * lineH - 2 * PAD_Y;
    const padTop = PAD_Y + (vAlign === 'top' || free <= 0 ? 0 : vAlign === 'bottom' ? free : free / 2);
    st.font = font;
    st.lineHeight = `${lineH}px`;
    st.color = style.color;
    st.background = style.fill ?? '#ffffff';
    st.textAlign = horizontalOf(style, cell ?? { value: null });
    st.whiteSpace = wrap ? 'pre-wrap' : 'pre';
    st.padding = `${padTop}px ${PAD_X}px ${PAD_Y}px`;
    // 칸 크기로 줄여 놓고 내용 크기를 잰 뒤 늘린다
    st.left = `${x}px`;
    st.top = `${y}px`;
    st.width = `${r.w}px`;
    st.height = `${r.h}px`;
    const box = editorBox({ x, y, w: r.w, h: r.h }, { w: el.scrollWidth + CARET_SLACK, h: el.scrollHeight }, sz, wrap);
    st.width = `${box.w}px`;
    st.height = `${box.h}px`;
  };

  const frame = useRef(0);
  const schedule = useCallback(() => {
    if (frame.current) return;
    frame.current = requestAnimationFrame(() => {
      frame.current = 0;
      drawRef.current();
    });
  }, []);

  useLayoutEffect(schedule);
  useEffect(() => () => cancelAnimationFrame(frame.current), []);

  // 창이 다른 배율의 화면으로 옮겨지면 다시 그린다(조건은 현재 배율마다 새로 건다)
  useEffect(() => {
    let mq: MediaQueryList | null = null;
    const onChange = () => {
      listen();
      schedule();
    };
    const listen = () => {
      mq?.removeEventListener('change', onChange);
      mq = window.matchMedia(`(resolution: ${window.devicePixelRatio}dppx)`);
      mq.addEventListener('change', onChange);
    };
    listen();
    return () => mq?.removeEventListener('change', onChange);
  }, [schedule]);

  useEffect(() => {
    const el = scrollerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // 처음부터 방향키가 먹도록 표의 입력창에 포커스를 둔다. 입력줄·검색창에서 입력 중이면 빼앗지 않는다.
  useEffect(() => {
    const active = document.activeElement;
    if (active instanceof HTMLTextAreaElement || active instanceof HTMLInputElement) return;
    inputRef.current?.focus({ preventScroll: true });
  }, []);

  // ── 선택 ────────────────────────────────────────────────

  /** 행·열 모드는 반대 축을 처음부터 끝까지 채운다 */
  const fill = useCallback((s: GridSel): GridSel => {
    const { lastRow: lr, lastCol: lc } = live.current;
    if (s.mode === 'rows') return { ...s, anchor: { ...s.anchor, col: 0 }, focus: { ...s.focus, col: lc } };
    if (s.mode === 'cols') return { ...s, anchor: { ...s.anchor, row: 0 }, focus: { ...s.focus, row: lr } };
    return s;
  }, []);

  const select = useCallback((next: GridSel) => {
    const s = fill(next);
    if (sameSel(s, live.current.sel)) return;
    live.current.sel = s;
    setSel(s);
  }, [fill]);

  const selectionOf = useCallback(
    (s: GridSel) => toSelection(s, sheet, rowView, live.current.geometry),
    [sheet, rowView],
  );

  useEffect(() => {
    onSelect(selectionOf(sel));
  }, [sel, onSelect, selectionOf]);

  // 보이는 행이 바뀌면 화면 번호가 가리키는 행도 달라지므로 처음 칸으로 돌린다
  const prevView = useRef(rowView);
  useEffect(() => {
    if (prevView.current === rowView) return;
    prevView.current = rowView;
    select(INITIAL_SEL);
  }, [rowView, select]);

  /** 병합 안의 칸은 앵커 칸으로 */
  const anchorOf = (row: number, col: number) => {
    const m = live.current.geometry.mergeBounds(row, col);
    return m ? { row: m.r1, col: m.c1 } : { row, col };
  };

  const scroll = () => {
    const el = scrollerRef.current;
    return { x: el?.scrollLeft ?? 0, y: el?.scrollTop ?? 0 };
  };

  /** 칸이 고정 영역 밖에서 가려져 있으면 보이도록 스크롤한다 */
  const reveal = (row: number, col: number) => {
    const el = scrollerRef.current;
    if (!el) return;
    const { geometry: g, size: sz, headerW: hw } = live.current;
    const fit = (cur: number, a: number, b: number, frozen: number, view: number) => {
      if (a - frozen < cur) return a - frozen;
      if (b - view > cur) return Math.min(a - frozen, b - view);
      return cur;
    };
    if (col >= g.freeze.cols) el.scrollLeft = fit(el.scrollLeft, g.colX(col), g.colX(col + 1), g.colX(g.freeze.cols), sz.w - hw);
    if (row >= g.freeze.rows) el.scrollTop = fit(el.scrollTop, g.rowY(row), g.rowY(row + 1), g.rowY(g.freeze.rows), sz.h - HEADER_H);
  };

  const origRow = (d: number) => live.current.geometry.origRowOf(d);

  // ── 편집기 ──────────────────────────────────────────────

  /**
   * 선택 앵커 칸(병합이면 앵커)에서 편집을 시작한다. keep은 기존 값을 넣고(더블클릭·F2),
   * replace는 입력창에 이미 들어온(또는 곧 들어올) 글자로 기존 값을 대체한다(타이핑·조합 시작).
   */
  const openEditor = (mode: 'keep' | 'replace'): boolean => {
    const el = inputRef.current;
    if (readonly || !el || editingRef.current) return false;
    const a = anchorOf(live.current.sel.anchor.row, live.current.sel.anchor.col);
    const orig = origRow(a.row);
    if (orig === undefined) return false;
    const original = editText(model.cells.get(orig)?.get(a.col), wb);
    select({ anchor: a, focus: { ...a }, mode: 'cells' });
    reveal(a.row, a.col);
    const ed: Editing = { row: a.row, col: a.col, orig, original };
    editingRef.current = ed;
    setEditing(ed);
    if (mode === 'keep') {
      el.value = original;
      el.setSelectionRange(original.length, original.length);
    }
    placeRef.current();
    return true;
  };

  /** 편집을 끝낸다. commit이고 값이 바뀌었으면 onCommit으로 보낸다. 편집 중이었으면 true. */
  const closeEditor = (commit: boolean): boolean => {
    const ed = editingRef.current;
    const el = inputRef.current;
    if (!ed || !el) return false;
    if (composingRef.current && document.activeElement === el) {
      // 조합 중에 닫으면(다른 칸 클릭 등) 먼저 조합을 확정한다. 그대로 비우면 IME가 남은 글자를 빈 입력창에 다시 넣어
      // 다음 칸에 한 번 더 입력된다. 포커스를 잠깐 뺐다 돌려주면 WebKit이 조합을 끝낸다.
      blurQuiet.current = true;
      el.blur();
      el.focus({ preventScroll: true });
      blurQuiet.current = false;
    }
    const value = el.value;
    editingRef.current = null;
    setEditing(null);
    el.value = '';
    el.scrollTop = 0;
    el.scrollLeft = 0;
    if (commit && value !== ed.original) onCommit([{ sheet, row: ed.orig, col: ed.col, input: value }]);
    placeRef.current();
    return true;
  };

  const closeRef = useRef(closeEditor);
  closeRef.current = closeEditor;
  useImperativeHandle(
    ref,
    () => ({
      commitEditor: () => void closeRef.current(true),
      formatArea: () => {
        const { sel: s, rowView: rv, geometry: g, lastRow: lr, lastCol: lc } = live.current;
        return formatArea(s, rv, g, lr, lc);
      },
      focus: () => inputRef.current?.focus({ preventScroll: true }),
    }),
    [],
  );
  // 시트를 바꾸는 등으로 표가 사라질 때 편집 중이던 값을 잃지 않는다(레이아웃 정리는 입력창 ref가 떨어지기 전에 돈다)
  useLayoutEffect(() => () => void closeRef.current(true), []);

  // ── 마우스 ──────────────────────────────────────────────

  const drag = useRef<Drag | null>(null);
  const pointer = useRef({ x: 0, y: 0 }); // 캔버스 기준 마지막 포인터 위치
  const autoFrame = useRef(0);

  const local = (e: { clientX: number; clientY: number }) => {
    const r = canvasRef.current!.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };

  const dragSelect = () => {
    const d = drag.current;
    if (d?.kind !== 'select') return;
    const { geometry: g, headerW: hw, sel: s } = live.current;
    const p = pointToCell(pointer.current.x, pointer.current.y, g, hw, HEADER_H, scroll());
    if (d.mode === 'cells') select({ ...s, focus: p });
    else if (d.mode === 'rows') select({ ...s, focus: { ...s.focus, row: p.row } });
    else select({ ...s, focus: { ...s.focus, col: p.col } });
  };

  const stopAutoScroll = () => {
    cancelAnimationFrame(autoFrame.current);
    autoFrame.current = 0;
  };
  useEffect(() => stopAutoScroll, []);

  /** 드래그 중 포인터가 뷰포트 밖이면 그쪽으로 계속 스크롤한다 */
  const autoScroll = () => {
    const d = drag.current;
    const el = scrollerRef.current;
    if (d?.kind !== 'select' || !el) return stopAutoScroll();
    const { size: sz, headerW: hw } = live.current;
    const speed = (out: number) => Math.sign(out) * Math.min(AUTOSCROLL_MAX, 2 + Math.abs(out) / 2);
    const { x, y } = pointer.current;
    const outX = d.mode === 'rows' ? 0 : x < hw ? x - hw : x > sz.w ? x - sz.w : 0;
    const outY = d.mode === 'cols' ? 0 : y < HEADER_H ? y - HEADER_H : y > sz.h ? y - sz.h : 0;
    if (outX === 0 && outY === 0) return stopAutoScroll();
    if (outX) el.scrollLeft += speed(outX);
    if (outY) el.scrollTop += speed(outY);
    dragSelect();
    autoFrame.current = requestAnimationFrame(autoScroll);
  };

  const onPointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    closeEditor(true); // 다른 곳을 누르면 편집을 확정한다
    inputRef.current?.focus({ preventScroll: true });
    if (e.button !== 0 || e.ctrlKey) return; // 오른쪽 버튼·⌃클릭은 contextmenu에서 처리한다
    const { geometry: g, headerW: hw, sel: s, lastRow: lr, lastCol: lc } = live.current;
    const p = local(e);
    const sc = scroll();
    const hit = hitTest(p.x, p.y, g, hw, HEADER_H, sc);
    pointer.current = p;
    const capture = () => e.currentTarget.setPointerCapture(e.pointerId);

    if ((hit.kind === 'colEdge' || hit.kind === 'rowEdge') && !readonly && onResize) {
      const axis = hit.kind === 'colEdge' ? 'col' : 'row';
      const index = axis === 'col' ? hit.col! : hit.row!;
      const size0 = axis === 'col' ? g.colX(index + 1) - g.colX(index) : g.rowY(index + 1) - g.rowY(index);
      const edge0 = axis === 'col' ? colEdgeX(index, g, hw, sc.x) : rowEdgeY(index, g, HEADER_H, sc.y);
      drag.current = { kind: 'resize', axis, index, start: axis === 'col' ? p.x : p.y, size0, edge0, px: size0 };
      guide.current = { kind: axis, pos: edge0 };
      schedule();
      capture();
      return;
    }

    const { row, col } = hit;
    if (hit.kind === 'cell' && row !== undefined && col !== undefined) {
      if (e.shiftKey) select({ anchor: s.anchor, focus: { row, col }, mode: 'cells' });
      else {
        const a = anchorOf(row, col);
        select({ anchor: a, focus: { ...a }, mode: 'cells' });
      }
      drag.current = { kind: 'select', mode: 'cells' };
    } else if ((hit.kind === 'rowHeader' || hit.kind === 'rowEdge') && row !== undefined) {
      const anchorRow = e.shiftKey ? s.anchor.row : row;
      select({ anchor: { row: anchorRow, col: 0 }, focus: { row, col: lc }, mode: 'rows' });
      drag.current = { kind: 'select', mode: 'rows' };
    } else if ((hit.kind === 'colHeader' || hit.kind === 'colEdge') && col !== undefined) {
      const anchorCol = e.shiftKey ? s.anchor.col : col;
      select({ anchor: { row: 0, col: anchorCol }, focus: { row: lr, col }, mode: 'cols' });
      drag.current = { kind: 'select', mode: 'cols' };
    } else return; // 마지막 행·열 너머 빈 곳
    capture();
  };

  const onPointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const p = local(e);
    pointer.current = p;
    const d = drag.current;
    const canvas = e.currentTarget;
    if (!d) {
      const { geometry: g, headerW: hw } = live.current;
      const hit = hitTest(p.x, p.y, g, hw, HEADER_H, scroll());
      const resizable = !readonly && !!onResize;
      const cursor = resizable && hit.kind === 'colEdge' ? 'col-resize' : resizable && hit.kind === 'rowEdge' ? 'row-resize' : '';
      if (canvas.style.cursor !== cursor) canvas.style.cursor = cursor;
      return;
    }
    if (d.kind === 'resize') {
      d.px = Math.max(MIN_SIZE, d.size0 + (d.axis === 'col' ? p.x : p.y) - d.start);
      guide.current = { kind: d.axis, pos: d.edge0 + d.px - d.size0 };
      schedule();
      return;
    }
    dragSelect();
    if (!autoFrame.current) autoScroll();
  };

  const finishResize = (d: Extract<Drag, { kind: 'resize' }>) => {
    const { geometry: g, sel: s } = live.current;
    const px = Math.round(d.px);
    const targets = resizeTargets(s, d.axis, d.index, g);
    onResize?.(d.axis === 'col' ? { kind: 'col', cols: targets, px } : { kind: 'row', rows: targets, px });
  };

  const onPointerUp = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const d = drag.current;
    if (!d) return;
    drag.current = null;
    stopAutoScroll();
    if (e.currentTarget.hasPointerCapture(e.pointerId)) e.currentTarget.releasePointerCapture(e.pointerId);
    if (d.kind === 'resize') {
      guide.current = null;
      schedule();
      if (e.type === 'pointerup' && d.px !== d.size0) finishResize(d);
    }
  };

  const onDoubleClick = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const p = local(e);
    const { geometry: g, headerW: hw } = live.current;
    const hit = hitTest(p.x, p.y, g, hw, HEADER_H, scroll());
    if (hit.kind !== 'cell' || hit.row === undefined || hit.col === undefined) return;
    const a = anchorOf(hit.row, hit.col);
    select({ anchor: a, focus: { ...a }, mode: 'cells' });
    openEditor('keep');
  };

  const onCanvasContextMenu = (e: React.MouseEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    const p = local(e);
    const { geometry: g, headerW: hw, sel: s, lastRow: lr, lastCol: lc } = live.current;
    const hit = hitTest(p.x, p.y, g, hw, HEADER_H, scroll());
    const ctx = contextTarget(s, hit, g, lr, lc);
    if (!ctx) return;
    select(ctx.sel);
    onContextMenu({ x: e.clientX, y: e.clientY, target: ctx.target, sel: selectionOf(fill(ctx.sel)) });
  };

  // ── 키보드 ──────────────────────────────────────────────

  const moveSel = (key: MoveKey, extend: boolean, jump: boolean) => {
    const { geometry: g, sel: s, size: sz, headerW: hw, lastRow: lr, lastCol: lc } = live.current;
    const sc = scroll();
    const fh = g.rowY(g.freeze.rows);
    const v = g.visible(0, fh + sc.y, Math.max(0, sz.w - hw), Math.max(0, sz.h - HEADER_H - fh));
    const next = move(s, key, {
      extend,
      jump,
      pageRows: Math.max(1, v.r2 - v.r1),
      geo: g,
      lastRow: lr,
      lastCol: lc,
      hasValue: (d, c) => {
        const orig = g.origRowOf(d);
        const cell = orig === undefined ? undefined : model.cells.get(orig)?.get(c);
        return !!cell && cell.value !== null && cell.value !== '';
      },
    });
    select(next);
    const f = live.current.sel.focus;
    reveal(next.mode === 'cols' ? -1 : f.row, next.mode === 'rows' ? -1 : f.col);
  };

  /** ⌥Enter: 캐럿 자리에 줄바꿈을 넣는다(편집기의 되돌리기 기록에 남도록 insertText를 먼저 쓴다) */
  const insertNewline = (el: HTMLTextAreaElement) => {
    if (!document.execCommand('insertText', false, '\n')) {
      el.setRangeText('\n', el.selectionStart, el.selectionEnd, 'end');
      placeRef.current();
    }
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const { key, altKey, shiftKey, metaKey, ctrlKey, keyCode } = e;
    const isComposing = e.nativeEvent.isComposing;
    if (editingRef.current) {
      const act = editorKeyAction({ key, altKey, shiftKey, metaKey, isComposing, keyCode });
      if (act === 'none') return; // 글자 입력·캐럿 이동·조합은 textarea가 처리한다
      e.preventDefault();
      if (act === 'cancel') closeEditor(false);
      else if (act === 'newline') insertNewline(e.currentTarget);
      else {
        closeEditor(true);
        moveSel(COMMIT_MOVE[act]!, false, false);
      }
      return;
    }
    // 조합 키(229)는 null — compositionstart가 편집을 연다
    const action = gridKeyAction({ key, shiftKey, metaKey, ctrlKey, keyCode, isComposing });
    if (!action) return;
    if (action.kind === 'selectAll') {
      e.preventDefault();
      const { lastRow: lr, lastCol: lc } = live.current;
      select({ anchor: { row: 0, col: 0 }, focus: { row: lr, col: lc }, mode: 'cells' });
      return;
    }
    if (action.kind === 'clear') {
      e.preventDefault();
      clearSelection();
      return;
    }
    if (action.kind === 'edit') {
      if (action.initial === undefined) {
        e.preventDefault();
        openEditor('keep');
      } else if (!openEditor('replace')) e.preventDefault();
      // 타이핑은 기본 동작으로 둔다: 빈 입력창에 글자가 그대로 들어가 기존 값을 대체한다
      return;
    }
    e.preventDefault();
    moveSel(action.key, shiftKey, metaKey);
  };

  /** 편집 중이 아닐 때 들어온 글자(받침 글쇠·받아쓰기 등 keydown 없이 들어온 입력)도 편집으로 이어 간다 */
  const onInput = (e: React.FormEvent<HTMLTextAreaElement>) => {
    const el = e.currentTarget;
    if (editingRef.current) {
      placeRef.current();
      return;
    }
    const composing = composingRef.current;
    const act = idleInputAction({ inputType: (e.nativeEvent as InputEvent).inputType ?? '', composing, value: el.value });
    if (act === 'discard') el.value = '';
    else if (act === 'open' && !openEditor('replace') && !composing) el.value = '';
  };

  const onCompositionStart = () => {
    composingRef.current = true;
    if (!editingRef.current) openEditor('replace'); // 같은 입력창에서 조합이 이어지므로 첫 글자를 잃지 않는다
  };

  const onCompositionEnd = (e: React.CompositionEvent<HTMLTextAreaElement>) => {
    composingRef.current = false;
    if (!editingRef.current) e.currentTarget.value = ''; // 읽기 전용이라 편집을 못 연 조합은 버린다
  };

  const onBlur = () => {
    // 조합 확정용으로 잠깐 뺀 포커스, 창 전환(앱 바꾸기)으로 잃은 포커스는 편집을 유지한다
    if (!blurQuiet.current && document.hasFocus()) closeEditor(true);
  };

  // ── 복사·붙여넣기 ─────────────────────────────────────

  /** 복사·잘라내기 범위(화면 행). 행·열 선택은 데이터가 있는 곳까지로 줄인다. */
  const clipRange = () => {
    const { geometry: g, sel: s } = live.current;
    const r = expandForMerges(s, g);
    const lastDataRow = (rowView ? rowView.length : model.rowCount) - 1;
    return {
      r1: r.r1,
      c1: r.c1,
      r2: Math.max(r.r1, Math.min(r.r2, lastDataRow)),
      c2: Math.max(r.c1, Math.min(r.c2, model.colCount - 1)),
    };
  };

  const onCopy = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    if (editingRef.current) return; // 편집 중에는 글자 복사
    e.preventDefault();
    e.clipboardData.setData('text/plain', copyText(wb, sheet, clipRange(), rowView));
  };

  /** 선택 칸의 값을 지운다(잘라내기·Delete). 한 번이 되돌리기 한 단계, 필터 중에는 보이는 행만. */
  const clearSelection = () => {
    if (readonly) return;
    const edits = clearEdits(wb, sheet, clipRange(), rowView);
    if (edits.length > 0) onCommit(edits);
  };

  const onCut = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    if (editingRef.current) return;
    onCopy(e);
    clearSelection();
  };

  const onPaste = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    if (editingRef.current) return; // 편집 중에는 글자 붙여넣기
    e.preventDefault();
    if (readonly) return;
    const text = e.clipboardData.getData('text/plain');
    if (text === '') return;
    const r = expandForMerges(live.current.sel, live.current.geometry);
    const { edits, dropped, blocked } = pasteEdits({ row: r.r1, col: r.c1 }, parseTsv(text), sheet, rowView, model.layout.merges);
    if (blocked) {
      onWarning(blocked);
      return;
    }
    if (edits.length > 0) onCommit(edits);
    if (dropped > 0) onWarning(`보이는 행이 모자라 ${dropped}줄은 붙여 넣지 않았습니다.`);
  };

  // WebKit은 선택된 글자가 없으면 복사·잘라내기 메뉴를 끈다. before* 이벤트를 막으면 켜진다(편집 중이 아닐 때만).
  useEffect(() => {
    const el = inputRef.current;
    if (!el) return;
    const enable = (e: Event) => {
      if (!editingRef.current) e.preventDefault();
    };
    el.addEventListener('beforecopy', enable);
    el.addEventListener('beforecut', enable);
    return () => {
      el.removeEventListener('beforecopy', enable);
      el.removeEventListener('beforecut', enable);
    };
  }, []);

  return (
    <div className="canvas-grid" ref={scrollerRef} onScroll={schedule}>
      <div
        className="canvas-grid-spacer"
        style={{ width: Math.max(headerW + geometry.totalWidth, size.w), height: Math.max(HEADER_H + geometry.totalHeight, size.h) }}
      >
        <div className="canvas-grid-view" style={{ width: size.w, height: size.h }}>
          <canvas
            ref={canvasRef}
            style={{ width: size.w, height: size.h }}
            onMouseDown={(e) => e.preventDefault() /* 포커스를 입력창에 남긴다 */}
            onPointerDown={onPointerDown}
            onPointerMove={onPointerMove}
            onPointerUp={onPointerUp}
            onPointerCancel={onPointerUp}
            onDoubleClick={onDoubleClick}
            onContextMenu={onCanvasContextMenu}
          />
          <CellEditor
            ref={inputRef}
            editing={editing !== null}
            onKeyDown={onKeyDown}
            onInput={onInput}
            onCompositionStart={onCompositionStart}
            onCompositionEnd={onCompositionEnd}
            onBlur={onBlur}
            onCopy={onCopy}
            onCut={onCut}
            onPaste={onPaste}
          />
        </div>
      </div>
    </div>
  );
});
