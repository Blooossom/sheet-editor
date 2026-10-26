# Sheet Editor 2차 구현 계획: 서식 표시와 서식 편집

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** xlsx를 Excel에서 보던 모습 그대로(병합·크기·색·테두리·글꼴·정렬·줄 바꿈·숫자 형식) 그리고, 그 위에서 값과 서식을 고쳐 원본에 안전하게 저장한다.

**Architecture:**
- 서식은 읽기 전용 `StyleTable`(styles.xml + theme)과 시트별 `SheetLayout`(열 너비·행 높이·병합·틀 고정)으로 모델에 들어간다.
- 화면은 Glide Data Grid를 버리고, 순수 계산 모듈(`src/grid/*`: 좌표·히트 테스트·선택 이동·줄 바꿈·스타일 해석)과 그것을 그리는 캔버스 컴포넌트(`src/ui/canvas/*`)로 바꾼다.
- 서식 편집은 1차와 같이 Op로 기록한다. 화면 모델과 저장기가 **같은 함수**(`internXf`, `applyStylePatch`)로 새 서식 번호를 만들어 번호가 항상 일치한다.

**Tech Stack:** 1차와 같음(Tauri 2, React 18, TS, vitest, fflate, @xmldom/xmldom, fast-formula-parser) + `ssf`(숫자 형식, Apache-2.0). Glide Data Grid는 Task 12에서 제거한다.

**Spec:** `docs/superpowers/specs/2026-10-26-sheet-editor-formatting-design.md` (1차: `docs/superpowers/specs/2026-10-23-sheet-editor-design.md`)

## Global Constraints

- 1차 Global Constraints(1차 계획 문서 참조)는 모두 유지된다. 특히 다음은 그대로다.
  - 손대지 않은 zip 파트는 바이트 그대로 둔다.
  - 고친 파트 안의 손대지 않은 부분은 의미가 같게 유지한다.
  - 숫자처럼 보이는 문자열을 보존한다.
  - csv 방식을 보존한다.
  - 저장 중 편집 잠금, 되돌리기는 replay 방식이다.
- styles.xml은 각 목록 **끝에만 추가**하고 `count`를 갱신한다. 기존 항목은 바꾸거나 지우지 않는다. 같은 정의가 이미 있으면 재사용한다.
- 새로 넣는 XML 요소는 스키마 순서를 지킨다.
  - worksheet: `sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData, …, mergeCells, …`
  - styleSheet: `numFmts, fonts, fills, borders, cellStyleXfs, cellXfs, cellStyles, dxfs, tableStyles, colors, extLst`
- 고친 칸 표시는 **왼쪽 위 주황 삼각형**(`#f59e0b`, 한 변 6px)으로 한다. 노란 배경은 쓰지 않는다. 서식만 바꾼 칸도 표시한다.
- 서식 변경도 dirty이며 되돌리기 대상이다.
- 한 번의 서식 변경 대상이 200,000칸을 넘으면 막고 다음 문구를 보여 준다: `"한 번에 서식을 바꿀 수 있는 칸은 200,000개까지입니다."`
- 필터가 걸려 있으면 서식 변경은 **보이는 행에만** 적용한다.
- csv는 서식 도구 막대가 꺼지고, 열 너비·행 높이 조절은 저장하지 않는다.
- 숫자 형식 로캘은 한국어로 고정한다(`aaa`→`월`, `aaaa`→`월요일`, `AM/PM`→`오전/오후`, `[$-412]`/`[$-ko-KR]` 무시).
- 화면 문구는 한국어다. 커밋마다 `change/YYYY-MM-DD.md`(작업한 날짜)에 한 줄씩 기록하고, 커밋 끝에 `Co-Authored-By` 줄을 붙인다(구현 모델 이름).

## Review Focus

1. **한글 IME 조합 중 Enter·방향키·Tab**: 조합 중에는 확정·이동으로 처리하면 안 된다. 마지막 글자가 사라지거나 두 번 입력되면 안 된다. → Task 11 테스트 + 수동 확인
2. **병합 칸 편집·붙여넣기·행 삭제**: 앵커 칸에 들어가야 한다. 일부만 덮으면 거부한다. 삭제 후에는 화면 모델과 저장 파일의 병합 범위가 같아야 한다. → Task 3, 11, 14 테스트
3. **같은 서식을 반복 적용하거나 되돌린 뒤 다시 적용**: styles.xml 항목이 늘지 않아야 한다. 저장 결과를 Excel이 복구 없이 열어야 한다(요소 순서·count). → Task 13, 14 테스트
4. **필터 상태의 서식·병합**: 원래 행 번호의 보이는 행에만 적용한다. 병합은 필터 상태에서 막는다(아래 Task 15 규칙). → Task 15 테스트
5. **5만 행 xlsx/csv 스크롤**: 화면 그리기가 보이는 칸 수에만 비례해야 한다. 스타일 해석과 숫자 형식은 캐시해야 한다. → Task 6, 8 성능 테스트

---

## 파일 구조 (새로 생기거나 크게 바뀌는 것)

```
src/core/types.ts            # StyleTable, Xf, FontDef, FillDef, BorderDef, SheetLayout, 새 Op 종류 추가
src/xlsx/styles.ts           # readStyles(parts) → StyleTable, readTheme
src/xlsx/layout.ts           # readLayout(sheetXml) → SheetLayout
src/xlsx/stylesWrite.ts      # 저장: styles.xml에 새 항목 추가, 시트에 s/cols/row ht/mergeCells 반영
src/core/styleOps.ts         # internXf, applyStylePatch, StylePatch (모델·저장기 공용)
src/core/layoutOps.ts        # 구조 연산 시 SheetLayout 이동(병합·행·열 크기)
src/format/numfmt.ts         # formatNumber(value, code, date1904) → {text, color?} (ssf + 한국어, 캐시)
src/grid/metrics.ts          # colWidthPx, rowHeightPx, maxDigitWidth, fontCss, 글꼴 대체표
src/grid/resolve.ts          # resolveStyle(wb, sheet, row, col) → ResolvedStyle (상속·캐시), resolveColor
src/grid/geometry.ts         # SheetGeometry: 좌표·보이는 범위·병합 색인·히트 테스트
src/grid/textLayout.ts       # wrapLines, measure 캐시
src/grid/selection.ts        # GridSel, 키 이동, 병합 확장
src/ui/canvas/CanvasGrid.tsx # 캔버스 표(그리기·마우스·스크롤·크기 조절)
src/ui/canvas/draw.ts        # 그리기 함수(칸·테두리·머리글·선택)
src/ui/canvas/CellEditor.tsx # 표 위 textarea 편집기(IME·⌥Enter)
src/ui/Toolbar.tsx           # 서식 도구 막대
```

`src/ui/Grid.tsx`(Glide)는 Task 12에서 지운다. `Selection` 타입은 `src/ui/selectionTypes.ts`로 옮긴다.

---

## 1단계: 서식 표시

### Task 1: 서식 모델 타입과 styles.xml·theme 읽기

**Files:**
- Modify: `src/core/types.ts`, `src/xlsx/read.ts`, `src/core/ops.ts`(`cloneWorkbook`), `src/document/store.ts`·`src/csv/csv.ts`(빈 모델 생성부), `tests/helpers/model.ts`
- Create: `src/xlsx/styles.ts`, `tests/xlsx/styles.test.ts`

**Interfaces:**
- Produces (`types.ts`):
  ```ts
  export type ColorRef = { rgb: string } | { theme: number; tint?: number } | { indexed: number } | { auto: true };
  export interface FontDef { name: string; size: number; bold?: boolean; italic?: boolean; underline?: 'single' | 'double'; strike?: boolean; color?: ColorRef; family?: number; charset?: number; scheme?: string }
  export interface FillDef { pattern: string /* 'none' | 'solid' | 'gray125' | … */; fg?: ColorRef; bg?: ColorRef }
  export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';
  export interface BorderSide { style: BorderStyle; color?: ColorRef }
  export interface BorderDef { left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide }
  export interface Alignment { horizontal?: string; vertical?: string; wrapText?: boolean; indent?: number; textRotation?: number; shrinkToFit?: boolean }
  export interface Xf { numFmtId: number; fontId: number; fillId: number; borderId: number; xfId?: number; alignment?: Alignment; protection?: string /* 원문 보존용 */; applyFlags?: Record<string, string> /* applyFont 등 원문 속성 */ }
  export interface StyleTable {
    numFmts: Map<number, string>;   // 사용자 정의 형식(id ≥ 164)
    fonts: FontDef[]; fills: FillDef[]; borders: BorderDef[]; xfs: Xf[]; // cellXfs
    themeColors: string[];          // 'RRGGBB' 12개 (lt1/dk1 순서 보정 포함: 0=lt1? → 아래 규칙)
    indexedColors?: string[];       // <colors><indexedColors> 사용자 정의가 있을 때
    baseCounts: { numFmts: number; fonts: number; fills: number; borders: number; xfs: number }; // 읽었을 때의 개수(새 항목 판별)
  }
  // WorkbookModel에 추가: styles: StyleTable  (csv·새 문서는 defaultStyleTable())
  export function defaultStyleTable(): StyleTable  // src/xlsx/styles.ts에서 export (create.ts의 STYLES와 같은 내용: Calibri 11, fills none/gray125)
  ```
- Produces (`styles.ts`): `export function readStyles(parts: Parts): StyleTable`, `export function defaultStyleTable(): StyleTable`
- `dateStyleIds`는 그대로 두되 `readStyles` 결과에서 계산한다(기존 규칙 그대로 이동).

테마 색 순서 규칙: `<a:clrScheme>`의 자식 순서는 `dk1, lt1, dk2, lt2, accent1…6, hlink, folHlink`이다. Excel의 theme 색 번호는 `0=lt1, 1=dk1, 2=lt2, 3=dk2, 4…=accent1…`이므로 0↔1, 2↔3을 바꿔 저장한다. `sysClr`는 `lastClr`를 쓴다.

- [ ] **Step 1: 실패하는 테스트 작성** (`tests/xlsx/styles.test.ts`, buildXlsx의 `styles`/`extraParts`로 theme 포함)

```ts
it('글꼴·채우기·테두리·xf·사용자 형식·테마 색을 읽는다', () => {
  const t = readStyles(unzipParts(buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData/>' }], styles: RICH_STYLES, extraParts: { 'xl/theme/theme1.xml': THEME } })));
  expect(t.fonts[1]).toMatchObject({ name: '맑은 고딕', size: 18, bold: true });
  expect(t.fills[2]).toEqual({ pattern: 'solid', fg: { theme: 9, tint: 0.5999938962981048 }, bg: { indexed: 64 } });
  expect(t.borders[1].bottom).toEqual({ style: 'dotted', color: { indexed: 64 } });
  expect(t.xfs[3]).toMatchObject({ numFmtId: 176, fontId: 1, fillId: 2, borderId: 1, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } });
  expect(t.numFmts.get(176)).toBe('m"월" d"일"');
  expect(t.themeColors.slice(0, 4)).toEqual(['FFFFFF', '000000', 'E7E6E6', '44546A']); // lt1, dk1, lt2, dk2
  expect(t.baseCounts).toEqual({ numFmts: 1, fonts: 2, fills: 3, borders: 2, xfs: 4 });
});
it('styles.xml이 없으면 기본 표', () => expect(readStyles(new Map()).fonts[0]).toMatchObject({ name: 'Calibri', size: 11 }));
```
`RICH_STYLES`와 `THEME`은 테스트 파일 안에 직접 작성한다(Office 기본 테마의 clrScheme만 있으면 된다).

- [ ] **Step 2: 실패 확인** — `npx vitest run tests/xlsx/styles.test.ts` → FAIL
- [ ] **Step 3: 구현**: `readXlsx`가 `styles: readStyles(parts)`를 채우고, `cloneWorkbook`이 StyleTable을 깊은 복사하게 한다(배열·Map 복사, 객체는 얕은 복사 후 고정 취급). 빈 모델들은 `defaultStyleTable()`을 쓴다.
- [ ] **Step 4: 전체 통과 확인** — `npm test && npx tsc --noEmit`
- [ ] **Step 5: 커밋** — `feat(xlsx): 서식 표와 테마 색 읽기`

---

### Task 2: 시트 레이아웃 읽기(열 너비·행 높이·병합·숨김·틀 고정)

**Files:**
- Modify: `src/core/types.ts`, `src/xlsx/read.ts`, `src/core/ops.ts`(clone), 빈 모델 생성부, `tests/helpers/model.ts`
- Create: `src/xlsx/layout.ts`, `tests/xlsx/layout.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface CellRange { r1: number; c1: number; r2: number; c2: number } // 0부터, 포함
  export interface ColInfo { width?: number /* 파일의 width(문자 단위, 패딩 포함) */; hidden?: boolean; style?: number }
  export interface RowInfo { height?: number /* pt */; hidden?: boolean; style?: number /* customFormat="1"일 때만 */ }
  export interface SheetLayout {
    colRuns: { c1: number; c2: number; info: ColInfo }[]; // 0부터, c1 오름차순, 겹치지 않음
    rows: Map<number, RowInfo>;
    defaultColWidth?: number; baseColWidth: number /* 기본 8 */; defaultRowHeight: number /* pt, 기본 15 */;
    merges: CellRange[];
    freeze: { rows: number; cols: number };   // pane state="frozen"/"frozenSplit"의 ySplit/xSplit
    showGridLines: boolean;                   // sheetView showGridLines (기본 true)
  }
  // SheetModel에 추가: layout: SheetLayout
  export function readLayout(sheetXml: string): SheetLayout     // src/xlsx/layout.ts
  export function emptyLayout(): SheetLayout
  export function colInfo(layout: SheetLayout, col: number): ColInfo   // colRuns 이분 탐색, 없으면 {}
  ```
  `<col min max>` 구간은 펼치지 않고 그대로 `colRuns`에 둔다(`max=16384` 같은 큰 구간 때문).

- [ ] **Step 1: 실패하는 테스트**

```ts
it('열 구간·행 높이·병합·틀 고정·격자선', () => {
  const L = readLayout(`<worksheet xmlns="${NS_MAIN}"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="11" xSplit="1" topLeftCell="B12" state="frozen"/></sheetView></sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="16.5"/><cols><col min="1" max="1" width="11.5" customWidth="1"/><col min="2" max="3" width="4.5" customWidth="1" style="3"/><col min="7" max="16384" width="9" hidden="1"/></cols>
    <sheetData><row r="1" ht="33" customHeight="1"/><row r="4" hidden="1"/><row r="5" s="2" customFormat="1"/></sheetData>
    <mergeCells count="2"><mergeCell ref="A1:F1"/><mergeCell ref="A12:A15"/></mergeCells></worksheet>`);
  expect(colInfo(L, 0)).toEqual({ width: 11.5 });
  expect(colInfo(L, 2)).toEqual({ width: 4.5, style: 3 });
  expect(colInfo(L, 9000)).toEqual({ width: 9, hidden: true });
  expect(colInfo(L, 4)).toEqual({});
  expect(L.rows.get(0)).toEqual({ height: 33 });
  expect(L.rows.get(3)).toEqual({ hidden: true });
  expect(L.rows.get(4)).toEqual({ style: 2 });
  expect(L.defaultRowHeight).toBe(16.5);
  expect(L.merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 5 }, { r1: 11, c1: 0, r2: 14, c2: 0 }]);
  expect(L.freeze).toEqual({ rows: 11, cols: 1 });
  expect(L.showGridLines).toBe(false);
});
```
- [ ] **Step 2–4**: 실패 확인 → 구현(정규식/xmldom 중 택1, 5만 행 시트에서 `readLayout`이 300ms 이내 — 성능 테스트 추가) → `npm test && npx tsc --noEmit`
- [ ] **Step 5: 커밋** — `feat(xlsx): 열 너비·행 높이·병합·틀 고정 읽기`

---

### Task 3: 구조 연산 시 레이아웃 이동(화면 모델)

**Files:**
- Create: `src/core/layoutOps.ts`, `tests/core/layoutOps.test.ts`
- Modify: `src/core/ops.ts`(structural applyOp에서 호출)

**Interfaces:**
- Produces: `export function shiftLayout(layout: SheetLayout, axis: 'row' | 'col', at: number, count: number /* 삽입 +, 삭제 - */): void`

규칙은 1차 저장기(`src/xlsx/structure.ts`)와 **같아야 한다**.
- 병합은 `shiftRange` 규칙을 따른다. 결과가 한 칸이 되거나 사라지면 제거한다.
- 행 정보는 함께 이동하고, 삭제된 행은 제거한다.
- 열 구간의 처리:
  - 삽입 위치를 포함하는 구간은 늘린다.
  - 구간의 시작 위치에 삽입하면 구간 전체를 민다.
  - 삭제는 구간을 줄이고, 빈 구간은 지운다.
- 틀 고정 행·열 수는 고정 영역 안에서 삽입·삭제되면 함께 늘리거나 줄인다.

- [ ] **Step 1: 실패하는 테스트** — 병합 `A12:A15`(r 11–14)에서 행 13 삭제 → `{r1:11,r2:13}`; `A1:B1`에서 열 B 삭제 → 제거; 열 구간 `{c1:1,c2:2}`에 at=1 삽입 → `{c1:2,c2:3}`, at=2 삽입 → `{c1:1,c2:3}`; 행 높이 이동. 그리고 **동일성 테스트**: `buildXlsx` 픽스처에 무작위 구조 연산 50개를 적용해 `readXlsx(saveXlsx(src, ops)).sheets[0].layout.merges`가 `replay(readXlsx(src), ops)`의 merges와 같다.
- [ ] **Step 2–4**: 실패 확인 → 구현 → `npm test`
- [ ] **Step 5: 커밋** — `feat(core): 행·열 구조 연산 시 병합·크기 이동`

---

### Task 4: 숫자 형식(`ssf` + 한국어)

**Files:**
- Create: `src/format/numfmt.ts`, `tests/format/numfmt.test.ts`
- Modify: `src/core/display.ts`(styles가 있으면 형식 코드로 표시), `package.json`(`ssf` 추가)

**Interfaces:**
- Produces:
  ```ts
  export function formatCode(styles: StyleTable, xfId: number | undefined): string  // 내장 id(0–49) 표 + numFmts
  export function formatNumber(value: number, code: string, date1904: boolean): { text: string; color?: string }  // 캐시: Map<code, Map<value, result>> 크기 제한(LRU 10,000)
  export function isDateCode(code: string): boolean
  ```
- `displayText(cell, wb)`: 숫자이고 `wb.styles`가 있으면 `formatNumber(v, formatCode(wb.styles, cell.styleId), wb.date1904).text`를 쓴다. 문자열 칸에 `@` 이외 형식이 있어도 문자열은 그대로 둔다. `editText`는 1차 규칙(날짜는 ISO)을 유지한다.
- 내장 형식 표: 0 General, 1 `0`, 2 `0.00`, 3 `#,##0`, 4 `#,##0.00`, 9 `0%`, 10 `0.00%`, 11 `0.00E+00`, 14 `yyyy-mm-dd`(한국어 Excel 기준), 15 `d-mmm-yy`, 20 `h:mm`, 21 `h:mm:ss`, 22 `yyyy-mm-dd h:mm`, 27–36·50–58은 한국어 Excel 기본(27 `yyyy"년" m"월"`, 28 `m"월" d"일"`, 29 `m"월" d"일"`, 30 `mm-dd-yy`, 31 `yyyy"년" m"월" d"일"`, 32 `h"시" mm"분"`, 33 `h"시" mm"분" ss"초"`, 34 `yyyy-mm-dd`, 35 `yyyy-mm-dd`, 36 `yyyy"년" m"월"`, 50–58은 27–36과 같은 대응), 37–40 회계 정수/소수, 45 `mm:ss`, 46 `[h]:mm:ss`, 47 `mmss.0`, 48 `##0.0E+0`, 49 `@`.

- [ ] **Step 1: 실패하는 테스트**

```ts
it.each([
  [46293, 'm"월" d"일"', '9월 28일'],
  [46293, 'yyyy"년" m"월" d"일" aaa', '2026년 9월 28일 월'],
  [46293.75, 'yyyy-mm-dd AM/PM h:mm', '2026-09-28 오후 6:00'],
  [1234567.891, '#,##0', '1,234,568'],
  [0.256, '0.0%', '25.6%'],
  [-5, '#,##0;[Red]-#,##0', '-5'],
  [1234, '₩#,##0', '₩1,234'],
  [42, '@', '42'],
  [0.5, 'General', '0.5'],
])('%s | %s → %s', (v, code, text) => expect(formatNumber(v as number, code as string, false).text).toBe(text));
it('색 구간', () => expect(formatNumber(-5, '#,##0;[Red]-#,##0', false).color).toBe('#FF0000'));
it('[$-412] 로캘 태그 무시', () => expect(formatNumber(46293, '[$-412]yyyy"년" m"월"', false).text).toBe('2026년 9월'));
it('1904 날짜 체계', () => expect(formatNumber(0, 'yyyy-mm-dd', true).text).toBe('1904-01-01'));
```
- [ ] **Step 2–4**: 실패 확인 → 구현(ssf 결과를 한국어 후처리, 지원 안 되는 코드는 General로 대체하고 던지지 않음) → `npm test`. 1차 `tests/core/display.test.ts`가 그대로 통과해야 한다(필요하면 styles 없는 경우의 기존 경로 유지).
- [ ] **Step 5: 커밋** — `feat: Excel 숫자 형식 표시(한국어)`

---

### Task 5: 크기 환산·글꼴·색·스타일 해석(순수 함수)

**Files:**
- Create: `src/grid/metrics.ts`, `src/grid/resolve.ts`, `tests/grid/metrics.test.ts`, `tests/grid/resolve.test.ts`

**Interfaces:**
- Produces:
  ```ts
  // metrics.ts
  export function maxDigitWidth(font: FontDef): number      // 표: Calibri 11→7, 맑은 고딕/Malgun Gothic 11→7, 맑은 고딕 10→7, Arial 10→7, 돋움/Dotum 9→6, 굴림/Gulim 9→6 — 표에 없으면 measureDigitWidth(대체 글꼴) 반올림
  export function colWidthPx(width: number, mdw: number): number      // trunc(((256*width + trunc(128/mdw)) / 256) * mdw)
  export function defaultColWidthPx(layout: SheetLayout, mdw: number): number // defaultColWidth 있으면 그것, 없으면 baseColWidth+5px 규칙: trunc((baseColWidth*mdw + 5)/mdw*256)/256 → colWidthPx
  export function rowHeightPx(pt: number): number           // Math.round(pt * 4 / 3)
  export function fontCss(font: FontDef, zoom?: number): string  // "bold italic 14.67px 'Apple SD Gothic Neo', …"
  export const FONT_FALLBACKS: Record<string, string[]>     // '맑은 고딕'|'Malgun Gothic' → ['Apple SD Gothic Neo','AppleGothic','sans-serif'], 'Calibri'/'Arial' → ['Helvetica Neue','Arial','sans-serif'], '굴림'/'돋움' → ['AppleGothic','Apple SD Gothic Neo','sans-serif'], '바탕'/'궁서' → ['AppleMyungjo','serif']
  // resolve.ts
  export interface ResolvedStyle { font: string /* css */; color: string; fill: string | null; borders: { left?: ResolvedBorder; right?: ResolvedBorder; top?: ResolvedBorder; bottom?: ResolvedBorder }; hAlign: 'left'|'center'|'right'|'fill'|'justify'|'centerContinuous'|'general'; vAlign: 'top'|'center'|'bottom'|'justify'; wrap: boolean; indent: number; underline?: 'single'|'double'; strike: boolean; numFmt: string }
  export interface ResolvedBorder { width: 1 | 2 | 3; dash: number[]; double: boolean; color: string }
  export function resolveColor(c: ColorRef | undefined, styles: StyleTable, fallback: string): string // theme+tint(HSL 휘도 규칙), indexed(기본 64색 표), auto→fallback
  export function effectiveXf(sheet: SheetModel, row: number, col: number): number   // 칸 s → 행 style → 열 style → 0
  export class StyleResolver { constructor(wb: WorkbookModel); get(xfId: number): ResolvedStyle /* xfId별 캐시 */; invalidate(): void }
  ```
- tint 규칙(Excel): RGB→HSL에서 L을 0–1로 두고, tint<0이면 `L*(1+tint)`, 아니면 `L*(1-tint) + tint`로 바꾼 뒤 RGB로 돌린다.
- 테두리 굵기: thin·hair·dotted·dashed·dashDot·dashDotDot = 1, medium·mediumDashed·mediumDashDot·mediumDashDotDot·slantDashDot = 2, thick = 3, double = 3(이중선). 대시 패턴: dotted [1,1], dashed [3,1], hair [1,1], mediumDashed [6,2] 등.

- [ ] **Step 1: 실패하는 테스트**: `colWidthPx(9.140625, 7) === 64`, `colWidthPx(13.7109375, 7) === 96`, `rowHeightPx(15) === 20`, `rowHeightPx(16.5) === 22`. `resolveColor({theme:9, tint:0.5999938962981048})`가 Office 기본 테마에서 `'#C6E0B4'`(±1 허용)로 나와야 한다. `{indexed: 64}`은 fallback(system foreground)이다. `effectiveXf`는 칸이 없을 때 행 style, 그다음 열 style 순서를 따른다. 그리고 `StyleResolver.get`은 같은 id에서 같은 객체를 돌려준다(캐시).
- [ ] **Step 2–4**: 실패 확인 → 구현 → `npm test`
- [ ] **Step 5: 커밋** — `feat(grid): 크기 환산·글꼴 대체·색·스타일 해석`

---

### Task 6: 좌표 계산과 병합 색인(`SheetGeometry`)

**Files:**
- Create: `src/grid/geometry.ts`, `tests/grid/geometry.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface GeometryInput {
    rowCount: number; colCount: number;                 // 화면에 둘 행·열 수(여분 포함)
    rowView: number[] | null;                           // 화면 행 → 원래 행
    rowHeight(origRow: number): number;                 // px, 숨김이면 0
    colWidth(col: number): number;                      // px, 숨김이면 0
    merges: CellRange[];                                // 원래 좌표
    freeze: { rows: number; cols: number };             // 화면 행·열 기준(필터 중이면 rows=0으로 둔다)
  }
  export class SheetGeometry {
    constructor(input: GeometryInput);
    readonly totalWidth: number; readonly totalHeight: number;
    colX(col: number): number; rowY(displayRow: number): number;              // 누적 오프셋(O(1))
    colAt(x: number): number; rowAt(y: number): number;                       // 이분 탐색
    visible(scrollX: number, scrollY: number, w: number, h: number): { r1: number; r2: number; c1: number; c2: number };
    mergeAt(origRow: number, col: number): CellRange | null;                  // O(1) 평균(행별 색인)
    cellRect(displayRow: number, col: number): { x: number; y: number; w: number; h: number }; // 병합이면 병합 전체(화면에서 연속일 때)
    displayRowOf(origRow: number): number | null;
  }
  ```
- 필터(rowView) 상태에서는 병합을 **화면에서 연속인 경우에만** 합쳐 그리고, 그 외에는 각 칸에 앵커 값만 보여 준다(결정).
- 성능: 5만 행 × 30열 입력에서 생성 50ms 이내, `visible`/`rowAt` 1µs 수준. 성능 테스트를 포함한다.

- [ ] **Step 1: 실패하는 테스트**: 숨김 행 높이 0이 누적에 반영된다. `rowAt` 경계값을 확인한다. `mergeAt`이 병합 안 모든 칸에서 같은 범위를 돌려준다. `cellRect`가 병합 크기를 돌려준다. rowView=[0,2,3]에서 병합(1–2)은 연속이 아니므로 합치지 않는다. 5만 행 생성은 50ms 미만이어야 한다.
- [ ] **Step 2–4**: 실패 확인 → 구현 → `npm test`
- [ ] **Step 5: 커밋** — `feat(grid): 좌표·히트 테스트·병합 색인`

---

### Task 7: 칸 안 줄 바꿈 계산

**Files:**
- Create: `src/grid/textLayout.ts`, `tests/grid/textLayout.test.ts`

**Interfaces:**
- Produces: `export function wrapLines(text: string, maxWidth: number, measure: (s: string) => number): string[]` — 명시적 `\n`·`\r\n`으로 먼저 나누고, 각 줄은 공백 단위로 감싸되 한 단어가 넘치면 글자 단위로 나눈다. 한글(가–힣)과 CJK는 글자 단위로 끊어도 된다. `export function makeMeasure(ctx: CanvasRenderingContext2D, font: string): (s: string) => number` (글꼴별 결과 캐시)
- [ ] **Step 1: 실패하는 테스트**: 고정폭 가짜 measure(글자당 10px)로 확인한다. `'가나다라마바'`를 폭 30으로 나누면 `['가나다','라마바']`가 된다. `'hello world foo'`를 폭 60으로 나누면 `['hello','world','foo']`가 된다. `'a\nb'`는 `['a','b']`가 되고, 빈 문자열은 `['']`가 된다.
- [ ] **Step 2–5**: 실패 확인 → 구현 → 통과 → 커밋 `feat(grid): 칸 안 줄 바꿈`

---

### Task 8: 캔버스 표 — 그리기(읽기 전용)

**Files:**
- Create: `src/ui/canvas/CanvasGrid.tsx`, `src/ui/canvas/draw.ts`, `src/ui/canvas/canvas.css`, `tests/ui/draw.test.ts`

**Interfaces:**
- Consumes: Task 4–7 전부, `store.wb`
- Produces (Props, 2단계에서 확장):
  ```ts
  interface CanvasGridProps { store: DocumentStore; sheet: number; rowView: number[] | null; readonly: boolean; sessionSizes?: { cols: Map<number, number>; rows: Map<number, number> } /* csv용 px 덮어쓰기 */ }
  export function drawCells(ctx: CanvasRenderingContext2D, view: DrawView): void   // draw.ts, DrawView = { geometry, resolver, wb, sheet, visible, scroll, dpr, selection, edited(origRow,col):boolean, showGridLines }
  ```

그리기 순서:
1. 배경과 격자선. 격자선은 채우기가 있는 칸 위에 그리지 않는다.
2. 채우기
3. 글자. 줄 바꿈, 정렬, 들여쓰기(1단계 = mdw×3 px)를 적용한다. 줄 바꿈이 없는 왼쪽 정렬 글자가 칸을 넘고 오른쪽 이웃이 비어 있으면 Excel처럼 넘쳐서 그린다. 오른쪽 정렬과 가운데 정렬도 같은 규칙을 따른다. 칸을 넘는 숫자는 `#`으로 채운다.
4. 테두리
5. 병합 칸. 병합 안의 격자선은 지우고, 앵커 칸의 값과 서식으로 병합 전체를 그린다.
6. 고친 칸 표시(주황 삼각형)
7. 틀 고정 영역과 구분선
8. 머리글. 열 머리글은 A, B…, 행 머리글은 원래 행 번호+1이다.
9. 선택 표시

- 스크롤: 바깥 `div`(overflow:auto) 안에 전체 크기 spacer를 두고, 캔버스는 sticky로 화면 크기만 그린다. 스크롤 이벤트마다 `requestAnimationFrame`으로 한 번만 그린다. `devicePixelRatio`를 반영한다.
- csv 또는 styles 없는 시트: 기본 글꼴 `-apple-system 13px`, 열 120px, 행 24px(1차와 비슷한 모습)로 그린다.
- 이 Task에서는 App의 기본 표를 바꾸지 않는다(교체는 Task 12). 개발 확인용으로, 창 URL에 `canvas=1`이 있을 때만 App이 CanvasGrid를 읽기 전용으로 쓰는 분기를 넣는다(Task 12에서 제거).

- [ ] **Step 1: 실패하는 테스트**: `draw.ts`의 순수 보조 함수를 테스트한다.
  - `textOrigin(rect, style, lineCount, lineHeight)`: 정렬별 x·y를 계산한다.
  - `overflowExtent(row, col, textWidth, isEmpty)`: 넘침 범위를 계산한다.
  - `numberFits(width, text)`: 숫자가 칸에 들어가지 않으면 `'###'`이 된다.
  - 실제 캔버스 그리기는 테스트하지 않는다.
- [ ] **Step 2–4**: 실패 확인 → 구현 → `npm test && npx tsc --noEmit && npm run build`
- [ ] **Step 5: 수동 확인 준비**: `.superpowers/sdd/<plan>/manual/양식.xlsx` 픽스처를 만드는 스크립트를 작성한다. 병합, 여러 줄, 채우기, 테두리, `m"월" d"일"`, 틀 고정이 들어간 주간 보고 형태다. 결과는 `npm run tauri dev`에서 `?canvas=1`로 확인할 수 있게 한다(사용자 확인은 Task 17).
- [ ] **Step 6: 커밋** — `feat(ui): 서식을 그리는 캔버스 표(읽기 전용)`

---

## 2단계: 편집

### Task 9: 선택과 키보드 이동(순수 함수)

**Files:**
- Create: `src/grid/selection.ts`, `tests/grid/selection.test.ts`, `src/ui/selectionTypes.ts`(기존 `Selection`·`ContextMenuRequest` 이동, Grid.tsx는 여기서 import)

**Interfaces:**
- Produces:
  ```ts
  export interface GridSel { anchor: { row: number; col: number }; focus: { row: number; col: number }; mode: 'cells' | 'rows' | 'cols' } // row는 화면 행
  export function expandForMerges(sel: GridSel, geo: SheetGeometry): { r1: number; c1: number; r2: number; c2: number }   // 병합이 걸치면 넓힌다(반복)
  export function move(sel: GridSel, key: 'ArrowUp'|'ArrowDown'|'ArrowLeft'|'ArrowRight'|'Home'|'End'|'PageUp'|'PageDown'|'Tab'|'ShiftTab'|'Enter'|'ShiftEnter', opts: { extend: boolean; jump: boolean /* ⌘ */; pageRows: number; geo: SheetGeometry; lastRow: number; lastCol: number }): GridSel
  export function toSelection(sel: GridSel, sheet: number, rowView: number[] | null, geo: SheetGeometry): Selection  // 1차 Selection(원래 행 번호, rows/cols, stale) 규칙 그대로
  ```
- 병합 칸 안에서 이동하면 병합 밖으로 한 번에 빠진다. 병합 칸에 들어가면 앵커를 가리킨다.
- ⌘+방향키는 Excel 규칙대로 데이터 경계로 이동한다(값이 있는 칸과 빈 칸의 경계).
- 숨김 행·열은 건너뛴다.

- [ ] **Step 1: 실패하는 테스트**: 병합 A12:A15 안에서 ↓를 누르면 A16으로 간다. 병합 안에서 → 를 누르면 B12로 간다. Shift+↓로 넓힌 범위가 병합 일부를 덮으면 `expandForMerges`로 병합 전체가 된다. 숨김 행 4를 건너뛴다. Tab/Enter로 이동한다. ⌘↓는 다음 데이터 경계로 간다.
- [ ] **Step 2–5**: 실패 확인 → 구현 → 통과 → 커밋 `feat(grid): 병합을 고려한 선택·이동`

---

### Task 10: 캔버스 표 — 마우스·스크롤·크기 조절·우클릭

**Files:**
- Modify: `src/ui/canvas/CanvasGrid.tsx`, `draw.ts`
- Test: `tests/grid/hit.test.ts`(히트 테스트 보조 함수: 머리글/경계선 판정)

**Interfaces:**
- Produces (Props 확장):
  ```ts
  onSelect(sel: Selection): void; onContextMenu(e: ContextMenuRequest): void;
  onResize?(e: { kind: 'col'; cols: number[]; px: number } | { kind: 'row'; rows: number[] /* 원래 행 */; px: number }): void;
  onStartEdit(e: { row: number; col: number; initial?: string /* 타이핑으로 시작한 글자 */ }): void;
  ```
- 다음 동작을 지원한다.
  - 클릭, 드래그 범위 선택, Shift+클릭
  - 머리글 클릭(행·열 선택)과 머리글 드래그
  - 더블클릭으로 편집 시작
  - 우클릭 메뉴. 대상은 행 머리글·열 머리글·칸이고, 선택 밖을 누르면 그 칸을 선택한다.
  - 머리글 경계선 드래그로 크기 조절. 조절 중에는 점선 가이드를 보여 주고, 놓으면 `onResize`를 호출한다. 선택된 여러 열 중 하나의 경계를 끌면 선택된 모든 열에 같은 너비를 적용한다.
  - 드래그 중 가장자리에 닿으면 자동 스크롤한다.
- `hitTest(x, y, geo, headerW, headerH, scroll) → { kind: 'cell'|'rowHeader'|'colHeader'|'corner'|'colEdge'|'rowEdge'; row?; col? }` 를 export하고 테스트한다(경계선 ±3px).
- [ ] **Step 1–4**: 히트 테스트를 먼저 작성한다 → 구현 → `npm test && npx tsc --noEmit && npm run build`
- [ ] **Step 5: 커밋** — `feat(ui): 캔버스 표 마우스 선택·크기 조절·우클릭`

---

### Task 11: 칸 편집기(IME·여러 줄)와 복사·붙여넣기

**Files:**
- Create: `src/ui/canvas/CellEditor.tsx`, `src/ui/canvas/clipboard.ts`, `tests/ui/editorKeys.test.ts`, `tests/ui/clipboard.test.ts`
- Modify: `src/ui/paste.ts`(병합 규칙), `CanvasGrid.tsx`

**Interfaces:**
- Produces:
  ```ts
  export function editorKeyAction(e: { key: string; altKey: boolean; shiftKey: boolean; metaKey: boolean; isComposing: boolean; keyCode: number }): 'commitDown'|'commitUp'|'commitRight'|'commitLeft'|'cancel'|'newline'|'none'
  // isComposing 또는 keyCode===229이면 무조건 'none'. Alt+Enter → 'newline'. Enter → commitDown, Shift+Enter → commitUp, Tab → commitRight, Shift+Tab → commitLeft, Escape → cancel.
  export function copyText(wb: WorkbookModel, sheet: number, range: CellRange, rowView: number[] | null): string // TSV, 여러 줄 칸은 "…"로 감싸고 " → ""
  export function pasteEdits(target, values, sheet, rowView, merges): { edits: CellEdit[]; dropped: number; blocked?: string }  // 기존 함수 확장
  ```
- 편집기 규칙:
  - 편집기는 병합 범위를 포함한 칸 사각형 위에 같은 글꼴·정렬로 뜬다.
  - 내용이 넘치면 아래·오른쪽으로 늘어나며, 화면 끝을 넘지 않는다.
  - 타이핑으로 시작하면 기존 값을 대체하고, 더블클릭이나 F2로 시작하면 기존 값을 유지한다.
  - 조합 중인 첫 글자로 편집을 시작하는 경우에는, 숨은 입력창에서 조합을 받아 그대로 편집기로 옮긴다(첫 글자 유실 금지).
- 붙여넣기 병합 규칙:
  - 대상이 병합이면 앵커에 붙여넣는다.
  - 붙여넣기 범위가 어떤 병합의 일부만 덮으면 `blocked = '병합된 셀의 일부는 바꿀 수 없습니다.'`를 돌려주고 아무것도 바꾸지 않는다.
  - 병합 안의 앵커가 아닌 칸에 들어가는 값은 버린다.
- `flushEdits()`(저장·닫기 전 편집 확정)를 CanvasGrid의 ref 메서드 `commitEditor()`로 바꾼다.

- [ ] **Step 1: 실패하는 테스트**: 다음을 확인한다.
  - `editorKeyAction`에 `isComposing: true, key: 'Enter'`를 주면 `'none'`이다.
  - `keyCode: 229`면 `'none'`이다.
  - Alt+Enter는 `'newline'`이다.
  - `copyText`가 여러 줄 칸을 따옴표로 감싼다.
  - `pasteEdits`가 병합 일부를 덮으면 `blocked`를 돌려준다.
  - 병합 앵커에는 붙여넣기가 들어간다.
- [ ] **Step 2–4**: 실패 확인 → 구현 → `npm test && npx tsc --noEmit && npm run build`
- [ ] **Step 5: 커밋** — `feat(ui): 칸 편집기(한글 조합·여러 줄)와 복사·붙여넣기`

---

### Task 12: Glide 제거, App을 캔버스 표로 교체

**Files:**
- Modify: `src/ui/App.tsx`, `src/ui/FormulaBar.tsx`(여러 줄 값: `textarea` 한 줄 높이, ⌥Enter 줄 바꿈, 내용이 여러 줄이면 펼침 버튼으로 3줄까지 확장), `package.json`(glide 제거), `index.html`(`#portal` 제거)
- Delete: `src/ui/Grid.tsx`

규칙:
- 1차 동작을 모두 그대로 옮긴다.
  - 필터·검색의 rowView, 행·열 삽입·삭제 메뉴와 단축키, 되돌리기 라우팅
  - 저장·닫기 흐름(편집기 확정 후 진행), 붙여넣기 결과 안내, 병합 거부 안내
- csv 창에서는 크기 조절을 `sessionSizes`에만 저장하고, 서식 도구 막대 자리는 비워 둔다.
- `?canvas=1` 개발 분기를 제거한다.
- 1차 수동 체크리스트 B항목(편집·저장·되돌리기·붙여넣기·필터)이 새 표에서도 같은 결과여야 한다(사용자 확인은 Task 17).

- [ ] **Step 1**: 구현
- [ ] **Step 2**: `npm test && npx tsc --noEmit && npm run build && cargo build --manifest-path src-tauri/Cargo.toml`를 실행한다. 번들 크기가 1차(881KB)보다 줄었는지 보고한다.
- [ ] **Step 3: 커밋** — `feat(ui): 캔버스 표로 교체하고 Glide 제거`

---

## 3단계: 서식 편집과 저장

### Task 13: 서식 연산 모델(`internXf`, `applyStylePatch`, 새 Op)

**Files:**
- Create: `src/core/styleOps.ts`, `tests/core/styleOps.test.ts`
- Modify: `src/core/types.ts`(Op 추가), `src/core/ops.ts`(applyOp가 새 Op 처리, dateStyleIds 갱신)

**Interfaces:**
- Produces:
  ```ts
  export interface StylePatch {
    font?: Partial<Pick<FontDef, 'name' | 'size' | 'bold' | 'italic' | 'underline' | 'strike'>> & { color?: ColorRef | null };
    fill?: { color: ColorRef } | null;                    // null = 채우기 없음
    border?: { preset: 'all' | 'outside' | 'none'; style?: 'thin' | 'medium' | 'thick'; color?: ColorRef };
    align?: { horizontal?: string | null; vertical?: string | null; wrapText?: boolean };
    numFmt?: string;                                       // 형식 코드('General' 포함)
    clear?: true;                                          // 서식 지우기 → xf 0
  }
  // Op 추가 (types.ts)
  | { kind: 'style'; sheet: number; rows: number[]; c1: number; c2: number; patch: StylePatch }   // rows: 원래 행 번호(필터면 보이는 행만), 오름차순
  | { kind: 'colWidth'; sheet: number; c1: number; c2: number; width: number | null /* 문자 단위, null = 기본 */ }
  | { kind: 'rowHeight'; sheet: number; rows: number[]; height: number | null /* pt, null = 기본 */ }
  | { kind: 'merge'; sheet: number; range: CellRange }
  | { kind: 'unmerge'; sheet: number; range: CellRange }
  // styleOps.ts
  export function internFont(t: StyleTable, f: FontDef): number      // 같은 정의(정규화 후 깊은 비교)가 있으면 그 번호, 없으면 끝에 추가
  export function internFill(t: StyleTable, f: FillDef): number
  export function internBorder(t: StyleTable, b: BorderDef): number
  export function internNumFmt(t: StyleTable, code: string): number   // 내장 코드면 내장 id, 아니면 numFmts에서 찾고 없으면 max(163, 기존 최대)+1
  export function internXf(t: StyleTable, xf: Xf): number
  export function applyStylePatch(t: StyleTable, baseXf: number, patch: StylePatch, pos: { top: boolean; bottom: boolean; left: boolean; right: boolean }): number
  export function styleCellsOf(sheet: SheetModel, op: Extract<Op, { kind: 'style' }>): { row: number; col: number; base: number; pos: {…} }[] // base = effectiveXf, 순서: 행 오름차순 → 열 오름차순
  ```

정규화와 비교 규칙: 비교 키는 `JSON.stringify(정렬된 키)`로 만든다. `applyFlags`는 비교에서 빼되, 새 xf에는 `applyFont/applyFill/applyBorder/applyAlignment/applyNumberFormat="1"`을 바뀐 항목에 붙인다. 원본 xf의 `xfId`·`protection`은 이어받는다.

테두리 `outside` 프리셋은 `pos`가 범위 가장자리인 변에만 선을 긋고 안쪽 변은 원래 값을 유지한다. `all`은 네 변을 모두 그린다. `none`은 네 변을 모두 지운다.

`applyOp`가 처리하는 새 Op:
- `style`: 각 칸의 styleId를 `applyStylePatch` 결과로 바꾸고, 칸이 없으면 값 없는 칸 `{value:null, styleId, edited:true}`을 만든다. 날짜 형식이면 `dateStyleIds`를 갱신한다.
- `colWidth`/`rowHeight`: `layout`을 바꾼다.
- `merge`/`unmerge`: `layout.merges`를 바꾼다. 앵커가 아닌 칸의 값을 지우는 것은 store가 같은 단계에 `set` Op로 넣는다.

- [ ] **Step 1: 실패하는 테스트**: 다음을 확인한다.
  - 같은 패치를 두 번 적용해도 `fonts/xfs` 길이가 늘지 않는다.
  - 굵게를 켜고 끈 뒤 다시 켜면 처음 만든 xf 번호를 재사용한다.
  - `outside` 테두리는 3×3 범위에서 가운데 칸을 바꾸지 않고, 모서리 칸은 해당 두 변만 바꾼다.
  - `clear`는 xf 0이 된다.
  - `internNumFmt('0.00%')`는 내장 10이다.
  - `internNumFmt('m"월" d"일"')`는 기존 176을 재사용한다.
  - 새 코드는 177이 된다.
  - replay 결정성: 같은 Op 목록을 두 번 replay하면 StyleTable이 같다.
  - `style` Op가 필터 행 목록(rows)에만 적용된다.
- [ ] **Step 2–5**: 실패 확인 → 구현 → `npm test` → 커밋 `feat(core): 서식 연산과 서식 번호 재사용`

---

### Task 14: 서식 저장(styles.xml 추가·s 적용·cols·행 높이·병합)

**Files:**
- Create: `src/xlsx/stylesWrite.ts`, `tests/xlsx/stylesWrite.test.ts`
- Modify: `src/xlsx/save.ts`(새 Op 분기), `src/xlsx/create.ts`(새 파일에 서식·크기·병합 기록)

**Interfaces:**
- Consumes: Task 13의 `applyStylePatch`, `styleCellsOf` 규칙. 저장기는 `readStyles(parts)`로 만든 StyleTable과 저장 시점의 시트 모델 정보(행 style, 열 style)가 필요하다. 그러므로 `SaveContext`에 `styles: StyleTable`(지연 생성)과 `layoutOf(sheet)`(해당 시트 DOM에서 계산)를 추가한다.
- Produces: `export function applyStyleOp(ctx: SaveContext, op: Extract<Op, {kind:'style'|'colWidth'|'rowHeight'|'merge'|'unmerge'}>): void`, `export function flushStyles(ctx: SaveContext): void` (finishSave 직전에 새 항목을 styles.xml에 추가, dirty 표시)

저장 규칙:
- styles.xml:
  - `baseCounts` 이후의 새 항목만 해당 목록 끝에 추가하고 `count`를 갱신한다.
  - `numFmts`가 없으면 styleSheet의 첫 자식으로 만든다.
  - 새 항목이 없으면 styles.xml을 dirty로 만들지 않는다.
- 칸: `setCellStyle(doc, row, col, s)`로 `s`만 바꾼다. 칸이 없으면 `<c r s/>`를 만든다(1차 `setCellInSheet`의 정렬 삽입 재사용). 원래 `s`가 없던 칸에 0이 아닌 값을 넣는다.
- colWidth:
  - `<cols>`가 없으면 `sheetFormatPr` 뒤(없으면 `sheetViews`→`dimension`→`sheetPr` 뒤, `sheetData` 앞)에 만든다.
  - 대상 열이 걸친 `<col>` 구간은 나눈다.
  - 대상 구간을 `width`, `customWidth="1"`로 쓴다. 이웃 구간과 속성이 같으면 합친다.
  - `null`이면 `width`·`customWidth`를 제거하고, 다른 속성이 없으면 구간을 지운다.
- rowHeight: `<row>`에 `ht`와 `customHeight="1"`을 쓴다. 행이 없으면 빈 `<row r>`를 정렬 위치에 만든다. `null`이면 두 속성을 제거한다.
- merge/unmerge:
  - `<mergeCells>`가 없으면 스키마 순서 자리에 만든다. 위치는 `sheetData`, `sheetCalcPr`, `sheetProtection`, `protectedRanges`, `scenarios`, `autoFilter`, `sortState`, `dataConsolidate`, `customSheetViews` 다음이다.
  - `count`를 갱신하고, unmerge로 비면 컨테이너를 제거한다.
- 새 파일(`createXlsx`): 모델의 StyleTable 전체를 styles.xml로 쓴다(기본 표에서 시작). 칸의 `s`, cols, 행 높이, 병합도 함께 쓴다.

- [ ] **Step 1: 실패하는 테스트**:
  - 굵게와 채우기를 적용해 저장한 뒤 `readXlsx`로 다시 읽으면 같은 서식이 나온다. `resolve` 결과로 비교한다.
  - 손대지 않은 칸의 `s=`는 문자열이 같다. 기존 fonts/fills/xfs 항목도 문자열이 같다.
  - 같은 패치를 10번 적용해 저장해도 `cellXfs count`는 +1만 늘어난다.
  - `numFmts`가 없던 파일에 사용자 형식을 쓰면 `<numFmts>`가 styleSheet 첫 자식으로 생긴다.
  - `<col min="1" max="5">`에서 3열만 너비를 바꾸면 세 구간으로 나뉜다.
  - mergeCells가 없는 시트에 병합을 넣으면 `sheetData` 뒤, `conditionalFormatting` 앞에 생긴다.
  - 저장한 파일을 다시 읽어 Task 13 모델로 replay한 결과와 서식·병합·크기가 같다(동일성 테스트, 무작위 30회).
  - 새 파일에 서식을 넣고 xlsx로 저장한 뒤 다시 읽어도 서식이 유지된다.
- [ ] **Step 2–5**: 실패 확인 → 구현 → `npm test` → 커밋 `feat(xlsx): 서식·크기·병합 저장`

---

### Task 15: 문서 상태 API(서식·크기·병합)

**Files:**
- Modify: `src/document/store.ts`, `tests/document/store.test.ts`

**Interfaces:**
- Produces:
  ```ts
  formatCells(sheet: number, sel: { rows: number[]; c1: number; c2: number }, patch: StylePatch): string | null;   // 차단 이유 또는 null. 20만 칸 초과 차단, csv면 '이 형식(csv)은 서식을 저장할 수 없습니다.'
  setColWidth(sheet: number, cols: number[], width: number | null): string | null;   // 연속 구간마다 Op, 한 단계
  setRowHeight(sheet: number, rows: number[], height: number | null): string | null;
  mergeCheck(sheet: number, range: CellRange): { ok: true; lossy: boolean } | { ok: false; reason: string };   // lossy = 앵커 외 값 존재
  merge(sheet: number, range: CellRange): string | null;   // 앵커 외 값 지우는 set Op + merge Op 한 단계
  unmerge(sheet: number, range: CellRange): string | null; // 범위와 겹치는 병합 모두 해제
  get hasFormatting(): boolean;   // 새 문서가 기본 서식 외 서식을 쓰는지(csv 저장 경고용)
  ```

규칙:
- 병합 차단:
  - 필터나 검색이 걸려 있을 때(UI가 넘겨주는 rows가 연속이 아니면) `'필터를 해제한 뒤 병합하세요.'`를 돌려준다.
  - 표(table) 범위와 겹치면 `'표(table) 안에서는 병합할 수 없습니다.'`를 돌려준다.
- 서식 범위:
  - `rows`는 UI가 원래 행 번호로 넘긴다.
  - 행·열 전체 선택은 사용 범위(rowCount·colCount)로 줄인다.
  - 병합 일부를 덮으면 병합 전체로 넓힌다. 열 방향은 c1·c2를 넓히고, 행 방향은 병합 행을 추가한다.
- 열 너비·행 높이(px → 단위 변환)는 UI가 하지 않고 store가 한다. 화면에서는 px로 받고, `pxToColWidth(px, mdw)`, `pxToPt(px)`를 Task 5 metrics에 추가해 쓴다.
- csv에서 `setColWidth`/`setRowHeight`를 부르면 차단 이유를 돌려준다. UI가 sessionSizes로 처리한다.

- [ ] **Step 1: 실패하는 테스트**: 다음 7가지를 확인한다.
  1. 서식 적용 후 dirty가 되고, undo로 styleId가 원래대로 돌아간다.
  2. 20만 칸 초과를 차단한다.
  3. 필터 rows(비연속)로 서식을 바꾸면 보이는 행에만 적용된다.
  4. 비연속 rows로 병합하면 차단된다.
  5. 병합 시 lossy를 판정하고, 앵커 외 값이 지워진다. 이것은 undo 한 번으로 복원된다.
  6. csv에서 서식 변경이 차단된다.
  7. 새 문서에 서식을 넣으면 `hasFormatting`이 true가 된다.
- [ ] **Step 2–5**: 실패 확인 → 구현 → `npm test` → 커밋 `feat: 서식·크기·병합 편집 상태 관리`

---

### Task 16: 서식 도구 막대와 연결

**Files:**
- Create: `src/ui/Toolbar.tsx`, `src/ui/toolbar.css`, `tests/ui/toolbarState.test.ts`
- Modify: `src/ui/App.tsx`, `src/ui/canvas/CanvasGrid.tsx`(onResize 연결), `src-tauri/src/menu.rs`(서식 메뉴: 굵게 ⌘B, 기울임 ⌘I, 밑줄 ⌘U → "menu" 이벤트 `bold`/`italic`/`underline`)

**Interfaces:**
- Produces: `export function toolbarState(wb: WorkbookModel, sheet: number, anchor: {row:number; col:number}): { font: string; size: number; bold: boolean; italic: boolean; underline: boolean; color: string; fill: string | null; hAlign: string; vAlign: string; wrap: boolean; numFmt: 'general'|'number'|'currency'|'percent'|'date'|'custom'; merged: boolean }`

도구 막대 구성(왼쪽부터):
1. 글꼴 종류 선택: 파일에 쓰인 글꼴과 `맑은 고딕`, `Calibri`, `Arial`, `굴림`, `돋움`, `바탕`
2. 크기: 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36 + 직접 입력
3. **굵게·기울임·밑줄**
4. 글자색과 채우기색: 팔레트 10색 × 6단계(테마 색 tint)와 최근 색 5개, `없음`
5. 테두리: 모든 테두리, 바깥 테두리, 테두리 없음, 선 굵기 선택
6. 가로 정렬(왼·가운데·오른)과 세로 정렬(위·가운데·아래)
7. 줄 바꿈
8. 병합하고 가운데 맞춤, 병합 풀기
9. 숫자 형식 선택: 일반, 숫자 `#,##0`, 통화 `₩#,##0`, 백분율 `0%`, 날짜 `yyyy-mm-dd`, 날짜 `m"월" d"일"`
10. 서식 지우기

동작 규칙:
- 단추 상태는 선택 앵커 칸의 서식을 보여 준다.
- 병합에서 lossy이면 다음 문구로 확인을 받는다: `"병합하면 왼쪽 위 칸의 값만 남습니다. 계속할까요?"` 단추는 `병합`과 `취소`다.
- 차단 이유는 `message(…, { kind: 'warning' })`로 보여 준다.
- csv와 저장 중에는 도구 막대 전체를 비활성화한다.
- 머리글 경계를 끌어 크기를 바꾸면 xlsx는 `store.setColWidth`/`setRowHeight`를, csv는 sessionSizes를 쓴다.
- 새 문서를 csv로 저장할 때 `store.hasFormatting`이면 다음 문구로 확인을 받는다: `"csv에는 서식이 저장되지 않습니다. 계속할까요?"`

- [ ] **Step 1: 실패하는 테스트**: `toolbarState`가 굵은 칸, 테마 채우기, `m"월" d"일"` 칸(→ 'date'), 병합 칸을 각각 맞게 판정한다.
- [ ] **Step 2–4**: 구현 → `npm test && npx tsc --noEmit && npm run build && cargo build --manifest-path src-tauri/Cargo.toml`
- [ ] **Step 5: 커밋** — `feat(ui): 서식 도구 막대`

---

### Task 17: 빌드와 수동 확인 목록

**Files:**
- Modify: `README.md`(2차 기능·제한 갱신), `.superpowers/sdd/<plan>/manual/CHECKLIST.md`(새로 작성)

- [ ] **Step 1**: `npm run tauri build`로 `.app`을 만들고 경로와 크기를 보고한다.
- [ ] **Step 2**: 수동 확인 목록(한국어)을 작성한다.
  1. 사용자의 주간 보고 xlsx 복사본이 Excel과 같은 모습인지(병합, 여러 줄, 색, 테두리, "9월 28일", 행 높이, 틀 고정)
  2. 한글 입력: 짧은 칸을 바로 입력하고, 여러 줄 칸에서 ⌥Enter로 줄을 추가하고, 조합 중 Enter를 누른다.
  3. 병합 칸 편집과 붙여넣기
  4. 서식 도구 막대 각 기능. 저장 후 Excel에서 열어 복구 경고가 없고 서식이 맞는지 본다.
  5. 열 너비·행 높이를 끌어 바꾸고 저장한 뒤 Excel에서 같은 크기인지 본다.
  6. 2만 행 csv와 큰 xlsx에서 스크롤이 매끄러운지 본다.
  7. 1차 체크리스트 B항목을 다시 확인한다.
- [ ] **Step 3**: README에 다음 제한을 적는다.
  - 조건부 서식 표시 없음
  - 그림·차트 표시 없음
  - 일부 글자만 다른 서식(rich text)은 칸 전체 서식으로 보임
  - 무늬·그라데이션 채우기는 대표 색으로 보임
  - 대체 글꼴 때문에 줄 바꿈 위치가 Excel과 조금 다를 수 있음
- [ ] **Step 4: 커밋** — `docs: 2차 기능 안내와 확인 목록`
