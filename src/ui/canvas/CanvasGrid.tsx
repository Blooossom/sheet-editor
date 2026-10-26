import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { DocumentStore } from '../../document/store';
import { SheetGeometry } from '../../grid/geometry';
import { colWidthPx, defaultColWidthPx, maxDigitWidth, rowHeightPx } from '../../grid/metrics';
import { StyleResolver } from '../../grid/resolve';
import { makeMeasure } from '../../grid/textLayout';
import { colInfo } from '../../xlsx/layout';
import { useStore } from '../useStore';
import { drawCells, isPlainSheet, type DrawView } from './draw';
import './canvas.css';

export interface CanvasGridProps {
  store: DocumentStore;
  sheet: number;
  rowView: number[] | null;
  readonly: boolean;
  sessionSizes?: { cols: Map<number, number>; rows: Map<number, number> }; // csv용 px 덮어쓰기
}

const PLAIN_COL = 120;
const PLAIN_ROW = 24;
const HEADER_H = 20;
const EXTRA_ROWS = 50;
const EXTRA_COLS = 10;
const MIN_ROWS = 100;
const MIN_COLS = 26;

export function CanvasGrid({ store, sheet, rowView, sessionSizes }: CanvasGridProps) {
  const version = useStore(store);
  const wb = store.wb;
  const model = wb.sheets[sheet];
  const scrollerRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });

  const plain = isPlainSheet(wb, model);
  const mdw = useMemo(() => (plain ? 7 : maxDigitWidth(wb.styles.fonts[0])), [plain, wb.styles]);
  const resolver = useMemo(() => new StyleResolver(wb), [wb]);

  const geometry = useMemo(() => {
    const layout = model.layout;
    const defaultCol = defaultColWidthPx(layout, mdw);
    const rowCount = rowView ? rowView.length : Math.max(model.rowCount + EXTRA_ROWS, MIN_ROWS);
    const colCount = Math.max(model.colCount + EXTRA_COLS, MIN_COLS);
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
  }, [model, rowView, plain, mdw, sessionSizes, version]);

  const headerW = useMemo(() => {
    const maxLabel = rowView ? rowView.reduce((m, r) => Math.max(m, r + 1), 1) : Math.max(model.rowCount + EXTRA_ROWS, MIN_ROWS);
    return Math.max(32, String(maxLabel).length * 7 + 14);
  }, [rowView, model.rowCount]);

  // 글꼴 문자열마다 측정기 하나(프레임 사이에 재사용)
  const measures = useRef(new Map<string, (s: string) => number>());

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
    const view: DrawView = {
      geometry,
      resolver,
      wb,
      sheet,
      visible: geometry.visible(fw + scroll.x, fh + scroll.y, Math.max(0, size.w - headerW - fw), Math.max(0, size.h - HEADER_H - fh)),
      scroll,
      dpr,
      selection: null,
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
    };
    drawCells(ctx, view);
  };

  const frame = useRef(0);
  const schedule = () => {
    if (frame.current) return;
    frame.current = requestAnimationFrame(() => {
      frame.current = 0;
      drawRef.current();
    });
  };

  useLayoutEffect(schedule);
  useEffect(() => () => cancelAnimationFrame(frame.current), []);

  useEffect(() => {
    const el = scrollerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  return (
    <div className="canvas-grid" ref={scrollerRef} onScroll={schedule}>
      <div
        className="canvas-grid-spacer"
        style={{ width: Math.max(headerW + geometry.totalWidth, size.w), height: Math.max(HEADER_H + geometry.totalHeight, size.h) }}
      >
        <canvas ref={canvasRef} style={{ width: size.w, height: size.h }} />
      </div>
    </div>
  );
}
