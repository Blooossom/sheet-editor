import { useEffect, useRef, useState, type ReactNode } from 'react';
import type { CellRange, ColorRef, StylePatch, WorkbookModel } from '../core/types';
import { formatCode, isDateCode } from '../format/numfmt';
import { effectiveXf, resolveColor, StyleResolver } from '../grid/resolve';
import './toolbar.css';

export type NumFmtKind = 'general' | 'number' | 'currency' | 'percent' | 'date' | 'custom';

export interface ToolbarState {
  font: string;
  size: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  color: string;
  fill: string | null;
  hAlign: string;
  vAlign: string;
  wrap: boolean;
  numFmt: NumFmtKind;
  merged: boolean;
}

export function numFmtKind(code: string): NumFmtKind {
  if (code.toLowerCase() === 'general') return 'general';
  if (isDateCode(code)) return 'date';
  if (/[₩$€£¥]|\[\$[^\]-]/.test(code)) return 'currency';
  const bare = code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '');
  if (bare.includes('%')) return 'percent';
  // 글자(따옴표)가 섞인 형식은 사용자 지정으로 본다
  if (/^[#,0.]+$/.test(code.split(';')[0])) return 'number';
  return 'custom';
}

/** 선택이 기존 병합과 정확히 같으면 Excel처럼 병합 풀기 */
export function mergeAction(merges: CellRange[], range: CellRange): 'merge' | 'unmerge' {
  const same = merges.some((m) => m.r1 === range.r1 && m.c1 === range.c1 && m.r2 === range.r2 && m.c2 === range.c2);
  return same ? 'unmerge' : 'merge';
}

/** 글꼴 크기 입력값. Excel 허용 범위 1–409, 0.5 단위. 아니면 null */
function parseSize(draft: string): number | null {
  const n = Number(draft.trim());
  return draft.trim() !== '' && Number.isFinite(n) && n >= 1 && n <= 409 && Number.isInteger(n * 2) ? n : null;
}

/**
 * 크기 입력창의 Enter·Escape·blur 처리. Enter·Escape는 표로 포커스를 돌리고, 그때 생기는 blur 한 번은 건너뛴다
 * (그러지 않으면 Enter는 두 번 적용되고 Escape는 입력한 크기가 적용된다).
 * apply: 적용할 크기, draft: 입력창에 다시 둘 값(그대로면 undefined)
 */
export function sizeInputStep(
  event: 'enter' | 'escape' | 'blur',
  draft: string,
  current: number,
  skipBlur: boolean,
): { apply: number | null; draft?: string; skipBlur: boolean; focusGrid: boolean } {
  if (event === 'blur' && skipBlur) return { apply: null, skipBlur: false, focusGrid: false };
  const leaving = event !== 'blur';
  if (event === 'escape') return { apply: null, draft: String(current), skipBlur: true, focusGrid: true };
  const n = parseSize(draft);
  if (n === null) return { apply: null, draft: String(current), skipBlur: leaving, focusGrid: leaving };
  return { apply: n !== current ? n : null, skipBlur: leaving, focusGrid: leaving };
}

/** 선택 앵커 칸(병합이면 앵커)의 서식 */
export function toolbarState(wb: WorkbookModel, sheet: number, anchor: { row: number; col: number }): ToolbarState {
  const sh = wb.sheets[sheet];
  const t = wb.styles;
  const id = effectiveXf(sh, anchor.row, anchor.col);
  const xf = t.xfs[id] ?? t.xfs[0];
  const font = t.fonts[xf.fontId] ?? t.fonts[0];
  const style = new StyleResolver(wb).get(id);
  const { row, col } = anchor;
  return {
    font: font.name,
    size: font.size,
    bold: !!font.bold,
    italic: !!font.italic,
    underline: !!font.underline,
    color: style.color,
    fill: style.fill || null,
    hAlign: style.hAlign,
    vAlign: style.vAlign,
    wrap: style.wrap,
    numFmt: numFmtKind(formatCode(t, id)),
    merged: sh.layout.merges.some((m) => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2),
  };
}

const DEFAULT_FONTS = ['맑은 고딕', 'Calibri', 'Arial', '굴림', '돋움', '바탕'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
// 통화 기호는 따옴표로 감싼 글자로 쓴다(Excel이 기호를 형식 문자로 오해하지 않게)
export const CURRENCY_FMT = '"₩"#,##0';
const NUM_FORMATS: [string, string][] = [
  ['General', '일반'],
  ['#,##0', '숫자 (#,##0)'],
  [CURRENCY_FMT, '통화 (₩#,##0)'],
  ['0%', '백분율 (0%)'],
  ['yyyy-mm-dd', '날짜 (yyyy-mm-dd)'],
  ['m"월" d"일"', '날짜 (m월 d일)'],
];
const KIND_LABEL: Record<NumFmtKind, string> = {
  general: '일반', number: '숫자', currency: '통화', percent: '백분율', date: '날짜', custom: '사용자 지정',
};

// Excel 팔레트: 열은 테마 색(lt1, dk1, lt2, dk2, accent1–6), 행은 기본색과 tint 5단계
const ACCENT_TINTS = [0, 0.8, 0.6, 0.4, -0.25, -0.5];
const PALETTE_TINTS: number[][] = [
  [0, -0.05, -0.15, -0.25, -0.35, -0.5],
  [0, 0.5, 0.35, 0.25, 0.15, 0.05],
  [0, -0.1, -0.25, -0.5, -0.75, -0.9],
  ACCENT_TINTS, ACCENT_TINTS, ACCENT_TINTS, ACCENT_TINTS, ACCENT_TINTS, ACCENT_TINTS, ACCENT_TINTS,
];
const MAX_RECENT = 5;

const sameColor = (a: ColorRef, b: ColorRef) => JSON.stringify(a) === JSON.stringify(b);

type Popover = 'size' | 'color' | 'fill' | 'border' | null;
type BorderWeight = 'thin' | 'medium' | 'thick';

export interface ToolbarProps {
  wb: WorkbookModel;
  sheet: number;
  anchor: { row: number; col: number } | null;
  disabled: boolean;
  onFormat(patch: StylePatch): void;
  onMerge(): void;
  onUnmerge(): void;
  /** 도구 막대 사용이 끝나 표로 포커스를 돌릴 때 */
  onDone(): void;
}

// 단추를 눌러도 표의 포커스를 빼앗지 않는다
const keepFocus = (e: React.MouseEvent) => e.preventDefault();

function Btn(props: { title: string; active?: boolean; disabled: boolean; onClick(): void; children: ReactNode; wide?: boolean }) {
  return (
    <button
      type="button"
      className={`tb-btn${props.active ? ' active' : ''}${props.wide ? ' wide' : ''}`}
      title={props.title}
      aria-label={props.title}
      aria-pressed={props.active}
      disabled={props.disabled}
      onMouseDown={keepFocus}
      onClick={props.onClick}
    >
      {props.children}
    </button>
  );
}

const Svg = ({ children }: { children: ReactNode }) => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round">
    {children}
  </svg>
);
const Lines = ({ xs }: { xs: [number, number][] }) => (
  <Svg>
    {xs.map(([a, b], i) => (
      <line key={i} x1={a} y1={3 + i * 3.3} x2={b} y2={3 + i * 3.3} />
    ))}
  </Svg>
);
const ICONS = {
  left: <Lines xs={[[2, 14], [2, 10], [2, 14], [2, 10]]} />,
  center: <Lines xs={[[2, 14], [4, 12], [2, 14], [4, 12]]} />,
  right: <Lines xs={[[2, 14], [6, 14], [2, 14], [6, 14]]} />,
  top: (
    <Svg>
      <line x1="2" y1="2.5" x2="14" y2="2.5" />
      <path d="M8 13V5M5 8l3-3 3 3" />
    </Svg>
  ),
  middle: (
    <Svg>
      <line x1="2" y1="8" x2="14" y2="8" />
      <path d="M8 1v4M6 3.5 8 5l2-1.5M8 15v-4M6 12.5 8 11l2 1.5" />
    </Svg>
  ),
  bottom: (
    <Svg>
      <line x1="2" y1="13.5" x2="14" y2="13.5" />
      <path d="M8 3v8M5 8l3 3 3-3" />
    </Svg>
  ),
  wrap: (
    <Svg>
      <path d="M2 3.5h12M2 8h9.5a2.25 2.25 0 0 1 0 4.5H8M9.5 11 8 12.5 9.5 14M2 12.5h3" />
    </Svg>
  ),
  merge: (
    <Svg>
      <rect x="1.5" y="3" width="13" height="10" rx="1" />
      <path d="M4 8h8M5.5 6.5 4 8l1.5 1.5M10.5 6.5 12 8l-1.5 1.5" />
    </Svg>
  ),
  unmerge: (
    <Svg>
      <rect x="1.5" y="3" width="13" height="10" rx="1" />
      <path d="M8 3v10M2.5 8h3M4 6.5 5.5 8 4 9.5M13.5 8h-3M12 6.5 10.5 8l1.5 1.5" />
    </Svg>
  ),
  border: (
    <Svg>
      <rect x="2" y="2" width="12" height="12" strokeDasharray="1.5 1.5" />
      <path d="M2 14h12" strokeWidth="2" />
    </Svg>
  ),
  clear: (
    <Svg>
      <path d="M9.5 2.5 13.5 6.5 7 13H3.5L2 11.5Z M6 6l4 4" />
      <line x1="8" y1="13" x2="14" y2="13" />
    </Svg>
  ),
};

export function Toolbar({ wb, sheet, anchor, disabled, onFormat, onMerge, onUnmerge, onDone }: ToolbarProps) {
  const state = anchor ? toolbarState(wb, sheet, anchor) : null;
  const off = disabled || !state;
  const [open, setOpen] = useState<Popover>(null);
  const [recent, setRecent] = useState<ColorRef[]>([]);
  const [weight, setWeight] = useState<BorderWeight>('thin');
  const [sizeDraft, setSizeDraft] = useState('');
  const rootRef = useRef<HTMLDivElement>(null);

  const size = state ? String(state.size) : '';
  useEffect(() => setSizeDraft(size), [size]);
  useEffect(() => {
    if (off) setOpen(null);
  }, [off]);

  // 팝오버 바깥을 누르거나 Escape면 닫는다
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (!(e.target instanceof Node) || !rootRef.current?.querySelector('.tb-popover')?.parentElement?.contains(e.target)) setOpen(null);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(null);
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [open]);

  const apply = (patch: StylePatch) => {
    setOpen(null);
    onFormat(patch);
  };
  const toggle = (p: Popover) => setOpen((cur) => (cur === p ? null : p));

  const skipBlur = useRef(false);
  const sizeEvent = (event: 'enter' | 'escape' | 'blur') => {
    if (!state) return;
    const out = sizeInputStep(event, sizeDraft, state.size, skipBlur.current);
    // 적용·포커스 이동 중에 blur가 동기로 일어나므로 먼저 표시해 둔다
    skipBlur.current = out.skipBlur;
    if (out.draft !== undefined) setSizeDraft(out.draft);
    if (out.apply !== null) onFormat({ font: { size: out.apply } });
    if (out.focusGrid) onDone();
  };

  const fonts = [...new Set([...wb.styles.fonts.map((f) => f.name).filter(Boolean), ...DEFAULT_FONTS])];
  const code = anchor ? formatCode(wb.styles, effectiveXf(wb.sheets[sheet], anchor.row, anchor.col)) : 'General';
  const known = NUM_FORMATS.some(([c]) => c === code);

  const swatch = (c: ColorRef) => resolveColor(c, wb.styles, '#000000');
  // 팔레트는 이 파일의 테마 색으로 계산하되 rgb로 보낸다. 새 문서·저장한 새 xlsx에는 테마 파트가 없어
  // Excel이 자기 기본 테마로 다른 색을 보일 수 있기 때문이다.
  const paletteRef = (theme: number, tint: number): ColorRef => ({ rgb: swatch(tint === 0 ? { theme } : { theme, tint }).slice(1) });
  const pickColor = (target: 'color' | 'fill', c: ColorRef | null) => {
    if (c) setRecent((list) => [c, ...list.filter((x) => !sameColor(x, c))].slice(0, MAX_RECENT));
    apply(target === 'color' ? { font: { color: c } } : { fill: c ? { color: c } : null });
  };

  const colorPopover = (target: 'color' | 'fill') => (
    <div className="tb-popover tb-palette" role="dialog" aria-label={target === 'color' ? '글자색' : '채우기 색'}>
      <div className="tb-grid">
        {ACCENT_TINTS.map((_, row) =>
          PALETTE_TINTS.map((tints, theme) => {
            const ref = paletteRef(theme, tints[row]);
            return (
              <button
                key={`${row}-${theme}`}
                type="button"
                className="tb-swatch"
                style={{ background: swatch(ref) }}
                title={swatch(ref)}
                onMouseDown={keepFocus}
                onClick={() => pickColor(target, ref)}
              />
            );
          }),
        )}
      </div>
      {recent.length > 0 && (
        <>
          <div className="tb-caption">최근 색</div>
          <div className="tb-grid">
            {recent.map((c, i) => (
              <button
                key={i}
                type="button"
                className="tb-swatch"
                style={{ background: swatch(c) }}
                title={swatch(c)}
                onMouseDown={keepFocus}
                onClick={() => pickColor(target, c)}
              />
            ))}
          </div>
        </>
      )}
      <button type="button" className="tb-menu-item" onMouseDown={keepFocus} onClick={() => pickColor(target, null)}>
        없음
      </button>
    </div>
  );

  const hAlignBtn = (value: 'left' | 'center' | 'right', title: string) => (
    <Btn title={title} disabled={off} active={state?.hAlign === value} onClick={() => apply({ align: { horizontal: state?.hAlign === value ? null : value } })}>
      {ICONS[value]}
    </Btn>
  );
  const vAlignBtn = (value: 'top' | 'center' | 'bottom', title: string, icon: ReactNode) => (
    // 아래쪽이 기본값이므로 속성을 지운다
    <Btn title={title} disabled={off} active={state?.vAlign === value} onClick={() => apply({ align: { vertical: value === 'bottom' ? null : value } })}>
      {icon}
    </Btn>
  );

  return (
    <div className="toolbar" ref={rootRef} role="toolbar" aria-label="서식">
      <select
        className="tb-font"
        title="글꼴"
        aria-label="글꼴"
        disabled={off}
        value={state?.font ?? ''}
        onChange={(e) => {
          onFormat({ font: { name: e.target.value } });
          onDone();
        }}
      >
        {fonts.map((f) => (
          <option key={f} value={f}>
            {f}
          </option>
        ))}
      </select>

      <div className="tb-anchor">
        <div className="tb-combo">
          <input
            className="tb-size"
            type="text"
            inputMode="decimal"
            title="글꼴 크기"
            aria-label="글꼴 크기"
            disabled={off}
            value={sizeDraft}
            onChange={(e) => setSizeDraft(e.target.value)}
            onFocus={() => (skipBlur.current = false)}
            onBlur={() => sizeEvent('blur')}
            onKeyDown={(e) => {
              if (e.nativeEvent.isComposing) return;
              if (e.key === 'Enter' || e.key === 'Escape') {
                e.preventDefault();
                sizeEvent(e.key === 'Enter' ? 'enter' : 'escape');
              }
            }}
          />
          <button type="button" className="tb-drop" title="글꼴 크기 목록" aria-label="글꼴 크기 목록" disabled={off} onMouseDown={keepFocus} onClick={() => toggle('size')}>
            ▾
          </button>
        </div>
        {open === 'size' && (
          <div className="tb-popover tb-list" role="listbox" aria-label="글꼴 크기">
            {SIZES.map((s) => (
              <button
                key={s}
                type="button"
                className={`tb-menu-item${state?.size === s ? ' active' : ''}`}
                onMouseDown={keepFocus}
                onClick={() => apply({ font: { size: s } })}
              >
                {s}
              </button>
            ))}
          </div>
        )}
      </div>

      <span className="tb-sep" />
      <Btn title="굵게 (⌘B)" disabled={off} active={state?.bold} onClick={() => apply({ font: { bold: !state?.bold } })}>
        <b>가</b>
      </Btn>
      <Btn title="기울임꼴 (⌘I)" disabled={off} active={state?.italic} onClick={() => apply({ font: { italic: !state?.italic } })}>
        <i className="tb-serif">가</i>
      </Btn>
      <Btn
        title="밑줄 (⌘U)"
        disabled={off}
        active={state?.underline}
        onClick={() => apply({ font: { underline: state?.underline ? undefined : 'single' } })}
      >
        <u>가</u>
      </Btn>

      <span className="tb-sep" />
      <div className="tb-anchor">
        <Btn title="글자색" disabled={off} active={open === 'color'} onClick={() => toggle('color')}>
          <span className="tb-color-icon">
            가<span className="tb-bar" style={{ background: state?.color ?? '#000000' }} />
          </span>
        </Btn>
        {open === 'color' && colorPopover('color')}
      </div>
      <div className="tb-anchor">
        <Btn title="채우기 색" disabled={off} active={open === 'fill'} onClick={() => toggle('fill')}>
          <span className="tb-color-icon">
            <svg width="14" height="12" viewBox="0 0 14 12" fill="none" stroke="currentColor" strokeWidth="1.2">
              <path d="M2 6 6.5 1.5 11 6 6.5 10.5Z" />
              <path d="M12.5 8.5c0 .8-.5 1.3-1 1.3s-1-.5-1-1.3.5-1.3 1-2c.5.7 1 1.2 1 2Z" fill="currentColor" stroke="none" />
            </svg>
            <span className="tb-bar" style={{ background: state?.fill ?? 'transparent', borderColor: state?.fill ? 'transparent' : '#c7c7cc' }} />
          </span>
        </Btn>
        {open === 'fill' && colorPopover('fill')}
      </div>

      <div className="tb-anchor">
        <Btn title="테두리" disabled={off} active={open === 'border'} onClick={() => toggle('border')}>
          {ICONS.border}
        </Btn>
        {open === 'border' && (
          <div className="tb-popover tb-list" role="menu" aria-label="테두리">
            <button type="button" className="tb-menu-item" onMouseDown={keepFocus} onClick={() => apply({ border: { preset: 'all', style: weight } })}>
              모든 테두리
            </button>
            <button type="button" className="tb-menu-item" onMouseDown={keepFocus} onClick={() => apply({ border: { preset: 'outside', style: weight } })}>
              바깥 테두리
            </button>
            <button type="button" className="tb-menu-item" onMouseDown={keepFocus} onClick={() => apply({ border: { preset: 'none' } })}>
              테두리 없음
            </button>
            <div className="tb-caption">선 굵기</div>
            {(
              [
                ['thin', '가는 선'],
                ['medium', '중간 선'],
                ['thick', '굵은 선'],
              ] as [BorderWeight, string][]
            ).map(([w, label]) => (
              <button
                key={w}
                type="button"
                className={`tb-menu-item tb-weight${weight === w ? ' active' : ''}`}
                onMouseDown={keepFocus}
                onClick={() => setWeight(w)}
              >
                <span className={`tb-line ${w}`} />
                {label}
              </button>
            ))}
          </div>
        )}
      </div>

      <span className="tb-sep" />
      {hAlignBtn('left', '왼쪽 맞춤')}
      {hAlignBtn('center', '가운데 맞춤')}
      {hAlignBtn('right', '오른쪽 맞춤')}
      <span className="tb-gap" />
      {vAlignBtn('top', '위쪽 맞춤', ICONS.top)}
      {vAlignBtn('center', '세로 가운데 맞춤', ICONS.middle)}
      {vAlignBtn('bottom', '아래쪽 맞춤', ICONS.bottom)}

      <span className="tb-sep" />
      <Btn title="줄 바꿈" disabled={off} active={state?.wrap} onClick={() => apply({ align: { wrapText: !state?.wrap } })}>
        {ICONS.wrap}
      </Btn>
      <Btn title="병합하고 가운데 맞춤" disabled={off} active={state?.merged} onClick={onMerge}>
        {ICONS.merge}
      </Btn>
      <Btn title="병합 풀기" disabled={off} onClick={onUnmerge}>
        {ICONS.unmerge}
      </Btn>

      <span className="tb-sep" />
      <select
        className="tb-numfmt"
        title="숫자 형식"
        aria-label="숫자 형식"
        disabled={off}
        value={known ? code : ''}
        onChange={(e) => {
          onFormat({ numFmt: e.target.value });
          onDone();
        }}
      >
        {!known && (
          <option value="" disabled>
            {state ? KIND_LABEL[state.numFmt] : ''}
          </option>
        )}
        {NUM_FORMATS.map(([c, label]) => (
          <option key={c} value={c}>
            {label}
          </option>
        ))}
      </select>

      <span className="tb-sep" />
      <Btn title="서식 지우기" disabled={off} onClick={() => apply({ clear: true })}>
        {ICONS.clear}
      </Btn>
    </div>
  );
}
