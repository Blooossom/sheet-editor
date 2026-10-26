import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { DocumentStore } from '../../document/store';
import { SheetGeometry } from '../../grid/geometry';
import { colEdgeX, hitTest, pointToCell, rowEdgeY } from '../../grid/hit';
import { colWidthPx, defaultColWidthPx, maxDigitWidth, rowHeightPx } from '../../grid/metrics';
import { StyleResolver } from '../../grid/resolve';
import { expandForMerges, move, toSelection, type GridSel, type MoveKey } from '../../grid/selection';
import { makeMeasure } from '../../grid/textLayout';
import { colInfo } from '../../xlsx/layout';
import type { ContextMenuRequest, Selection } from '../selectionTypes';
import { useStore } from '../useStore';
import { drawCells, isPlainSheet, type DrawView } from './draw';
import './canvas.css';

export type ResizeEvent =
  | { kind: 'col'; cols: number[]; px: number }
  | { kind: 'row'; rows: number[] /* 원래 행 */; px: number };

export interface StartEditEvent {
  row: number; // 원래 행(병합이면 앵커 칸)
  col: number;
  initial?: string; // 타이핑으로 시작한 글자
}

export interface CanvasGridProps {
  store: DocumentStore;
  sheet: number;
  rowView: number[] | null;
  readonly: boolean;
  sessionSizes?: { cols: Map<number, number>; rows: Map<number, number> }; // csv용 px 덮어쓰기
  onSelect(sel: Selection): void;
  onContextMenu(e: ContextMenuRequest): void;
  onResize?(e: ResizeEvent): void;
  onStartEdit(e: StartEditEvent): void;
}

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

const MOVE_KEYS: Record<string, MoveKey> = {
  ArrowUp: 'ArrowUp', ArrowDown: 'ArrowDown', ArrowLeft: 'ArrowLeft', ArrowRight: 'ArrowRight',
  Home: 'Home', End: 'End', PageUp: 'PageUp', PageDown: 'PageDown',
};

type Drag =
  | { kind: 'select'; mode: GridSel['mode'] }
  | { kind: 'resize'; axis: 'col' | 'row'; index: number; start: number; size0: number; edge0: number; px: number };

const sameSel = (a: GridSel, b: GridSel) =>
  a.mode === b.mode && a.anchor.row === b.anchor.row && a.anchor.col === b.anchor.col && a.focus.row === b.focus.row && a.focus.col === b.focus.col;

const between = (v: number, a: number, b: number) => v >= Math.min(a, b) && v <= Math.max(a, b);

export function CanvasGrid({ store, sheet, rowView, readonly, sessionSizes, onSelect, onContextMenu, onResize, onStartEdit }: CanvasGridProps) {
  const version = useStore(store);
  const wb = store.wb;
  const model = wb.sheets[sheet];
  const scrollerRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });
  const [sel, setSel] = useState<GridSel>(INITIAL_SEL);

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
  const live = useRef({ geometry, sel, size, headerW, lastRow, lastCol });
  live.current = { geometry, sel, size, headerW, lastRow, lastCol };

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

  const startEdit = (initial?: string) => {
    if (readonly) return;
    const a = anchorOf(live.current.sel.anchor.row, live.current.sel.anchor.col);
    const row = origRow(a.row);
    if (row === undefined) return;
    onStartEdit(initial === undefined ? { row, col: a.col } : { row, col: a.col, initial });
  };

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
    if (e.button !== 0 || e.ctrlKey) return; // 오른쪽 버튼·⌃클릭은 contextmenu에서 처리한다
    scrollerRef.current?.focus({ preventScroll: true });
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

    if (hit.kind === 'cell' && hit.row! >= 0 && hit.col! >= 0) {
      if (e.shiftKey) select({ anchor: s.anchor, focus: { row: hit.row!, col: hit.col! }, mode: 'cells' });
      else {
        const a = anchorOf(hit.row!, hit.col!);
        select({ anchor: a, focus: { ...a }, mode: 'cells' });
      }
      drag.current = { kind: 'select', mode: 'cells' };
    } else if ((hit.kind === 'rowHeader' || hit.kind === 'rowEdge') && hit.row! >= 0) {
      const anchorRow = e.shiftKey ? s.anchor.row : hit.row!;
      select({ anchor: { row: anchorRow, col: 0 }, focus: { row: hit.row!, col: lc }, mode: 'rows' });
      drag.current = { kind: 'select', mode: 'rows' };
    } else if ((hit.kind === 'colHeader' || hit.kind === 'colEdge') && hit.col! >= 0) {
      const anchorCol = e.shiftKey ? s.anchor.col : hit.col!;
      select({ anchor: { row: 0, col: anchorCol }, focus: { row: lr, col: hit.col! }, mode: 'cols' });
      drag.current = { kind: 'select', mode: 'cols' };
    } else return;
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
    if (d.axis === 'col') {
      let cols = [d.index];
      if (s.mode === 'cols' && between(d.index, s.anchor.col, s.focus.col)) {
        cols = [];
        const [c1, c2] = [Math.min(s.anchor.col, s.focus.col), Math.max(s.anchor.col, s.focus.col)];
        for (let c = c1; c <= c2; c++) if (c === d.index || !g.isColHidden(c)) cols.push(c);
      }
      onResize?.({ kind: 'col', cols, px });
    } else {
      let display = [d.index];
      if (s.mode === 'rows' && between(d.index, s.anchor.row, s.focus.row)) {
        display = [];
        const [r1, r2] = [Math.min(s.anchor.row, s.focus.row), Math.max(s.anchor.row, s.focus.row)];
        for (let r = r1; r <= r2; r++) if (r === d.index || !g.isRowHidden(r)) display.push(r);
      }
      const rows = display.map((r) => g.origRowOf(r)).filter((r): r is number => r !== undefined);
      onResize?.({ kind: 'row', rows, px });
    }
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
    if (hit.kind !== 'cell' || hit.row! < 0 || hit.col! < 0) return;
    const a = anchorOf(hit.row!, hit.col!);
    select({ anchor: a, focus: { ...a }, mode: 'cells' });
    startEdit();
  };

  const onCanvasContextMenu = (e: React.MouseEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    const p = local(e);
    const { geometry: g, headerW: hw, sel: s, lastRow: lr, lastCol: lc } = live.current;
    const hit = hitTest(p.x, p.y, g, hw, HEADER_H, scroll());
    let next = s;
    let target: ContextMenuRequest['target'];
    if ((hit.kind === 'rowHeader' || hit.kind === 'rowEdge') && hit.row! >= 0) {
      target = 'row';
      if (!(s.mode === 'rows' && between(hit.row!, s.anchor.row, s.focus.row))) {
        next = { anchor: { row: hit.row!, col: 0 }, focus: { row: hit.row!, col: lc }, mode: 'rows' };
      }
    } else if ((hit.kind === 'colHeader' || hit.kind === 'colEdge') && hit.col! >= 0) {
      target = 'col';
      if (!(s.mode === 'cols' && between(hit.col!, s.anchor.col, s.focus.col))) {
        next = { anchor: { row: 0, col: hit.col! }, focus: { row: lr, col: hit.col! }, mode: 'cols' };
      }
    } else if (hit.kind === 'cell' && hit.row! >= 0 && hit.col! >= 0) {
      target = 'cell';
      const r = expandForMerges(s, g);
      const inside = hit.row! >= r.r1 && hit.row! <= r.r2 && hit.col! >= r.c1 && hit.col! <= r.c2;
      if (!inside) {
        const a = anchorOf(hit.row!, hit.col!);
        next = { anchor: a, focus: { ...a }, mode: 'cells' };
      }
    } else return;
    select(next);
    onContextMenu({ x: e.clientX, y: e.clientY, target, sel: selectionOf(fill(next)) });
  };

  // ── 키보드 ──────────────────────────────────────────────

  const onKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (e.nativeEvent.isComposing || e.keyCode === 229) return; // 조합 중 키는 편집 시작(Task 11)에서 다룬다
    const { geometry: g, sel: s, size: sz, headerW: hw, lastRow: lr, lastCol: lc } = live.current;
    let key: MoveKey | undefined = MOVE_KEYS[e.key];
    if (e.key === 'Tab') key = e.shiftKey ? 'ShiftTab' : 'Tab';
    else if (e.key === 'Enter') key = e.shiftKey ? 'ShiftEnter' : 'Enter';
    if (key) {
      e.preventDefault();
      const sc = scroll();
      const fh = g.rowY(g.freeze.rows);
      const v = g.visible(0, fh + sc.y, Math.max(0, sz.w - hw), Math.max(0, sz.h - HEADER_H - fh));
      const next = move(s, key, {
        extend: e.shiftKey,
        jump: e.metaKey,
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
      return;
    }
    if (e.key === 'F2') {
      e.preventDefault();
      startEdit();
      return;
    }
    if (e.key.length === 1 && !e.metaKey && !e.ctrlKey) {
      e.preventDefault();
      startEdit(e.key);
    }
  };

  return (
    <div className="canvas-grid" ref={scrollerRef} tabIndex={0} onScroll={schedule} onKeyDown={onKeyDown}>
      <div
        className="canvas-grid-spacer"
        style={{ width: Math.max(headerW + geometry.totalWidth, size.w), height: Math.max(HEADER_H + geometry.totalHeight, size.h) }}
      >
        <canvas
          ref={canvasRef}
          style={{ width: size.w, height: size.h }}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={onPointerUp}
          onPointerCancel={onPointerUp}
          onDoubleClick={onDoubleClick}
          onContextMenu={onCanvasContextMenu}
        />
      </div>
    </div>
  );
}
