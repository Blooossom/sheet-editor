# Sheet Editor 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** xlsx와 csv를 더블클릭으로 바로 열어 셀 값과 행·열을 고치고, 원본 파일에 그대로 덮어써 저장하는 가벼운 맥 앱을 만든다.

**Architecture:** 파일 형식 처리(xlsx/csv 읽기·쓰기), 문서 모델과 편집 연산, 수식 참조 이동, 수식 재계산은 모두 순수 TypeScript 모듈(`src/core`, `src/csv`, `src/xlsx`, `src/formula`)이며 vitest로 검증한다. 편집은 `Op`(연산)의 목록으로 기록한다. xlsx를 저장할 때는 **원본 바이트에 연산을 차례로 다시 적용**해서, 고친 부분만 바뀌게 한다. Rust(Tauri)는 파일 입출력, 인코딩 변환, 창·메뉴·파일 열기 이벤트만 맡는다.

**Tech Stack:** Tauri 2, React + TypeScript (Vite), vitest, fflate(zip), @xmldom/xmldom(테스트 환경의 XML), @glideapps/glide-data-grid(가상 스크롤 표), fast-formula-parser(수식 계산, MIT), Rust encoding_rs·tempfile, tauri-plugin-dialog

**Spec:** `docs/superpowers/specs/2026-10-23-sheet-editor-design.md`

## Global Constraints

- 맥 데스크톱 앱, Tauri 2 + React + TypeScript. 패키지 매니저는 npm.
- 지원 형식은 `.xlsx`, `.csv` 두 가지뿐이다.
- 저장(⌘S)은 원본 파일을 바로 덮어쓴다. 다른 이름으로 저장하기는 없다. 단, 새 파일을 처음 저장할 때만 위치·이름·형식을 묻는다.
- xlsx 저장 시 고치지 않은 zip 파트는 압축 해제 내용이 원본과 바이트 단위로 같아야 한다. 고친 시트에서도 고치지 않은 셀의 XML은 그대로여야 한다(공유 수식을 일반 수식으로 푸는 경우만 예외).
- 문자열은 `t="inlineStr"`로 쓴다. 기존 `sharedStrings.xml`은 건드리지 않는다. 기존 셀의 `s=`(스타일)는 유지한다.
- 저장한 xlsx에는 `<calcPr fullCalcOnLoad="1"/>`을 설정한다.
- csv는 원래 인코딩(UTF-8 / UTF-8 BOM / EUC-KR), 구분자, 줄바꿈, 따옴표 방식을 유지한다. 새 csv는 UTF-8, BOM 없음, 쉼표, LF.
- csv 셀은 모두 문자열이다. 숫자 변환, 수식 계산을 하지 않는다.
- 검색·필터는 화면만 바꾸고 파일 내용과 행 순서는 바꾸지 않는다.
- 다른 파일은 새 창으로 연다. 이미 열린 파일은 그 창을 앞으로 가져온다. 상태는 창마다 따로 둔다.
- 화면 문구는 한국어다.
- 커밋할 때마다 `change/YYYY-MM-DD.md`에 변경 요약을 한 줄씩 추가한다(형식은 기존 파일 참고). 커밋 메시지 끝에는 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`를 붙인다.

## Review Focus

1. **숫자처럼 보이는 문자열**(`00123`, 16자리 이상 ID, 전화번호): xlsx에서도 숫자로 바뀌어 앞의 0이 사라지거나 지수 표기가 되면 안 된다. csv에서는 절대 변환하지 않는다. → Task 4 `parseInput` 테스트
2. **따옴표 안에 구분자·따옴표·줄바꿈이 있는 csv와 마지막 줄바꿈이 없는 csv**: 열고 바로 저장하면 바이트가 같아야 한다. → Task 5 왕복 테스트
3. **EUC-KR csv에 EUC-KR로 표현할 수 없는 글자(이모지 등)를 넣고 저장**: 조용히 `&#...;`나 `?`로 바뀌면 안 되고, 저장을 거부하며 이유를 알려야 한다. 한글·공백이 있는 경로도 정상 동작해야 한다. → Task 6 테스트
4. **수식이 있는 xlsx를 고친 뒤 Excel로 열 때 "복구" 경고가 뜨는 경우**: calcChain을 지우고, 공유 수식의 대표 셀을 덮어써도 나머지 셀의 수식이 살아 있어야 한다. → Task 8 테스트
5. **필터가 걸린 상태에서 편집·행 삭제**: 화면상의 행 번호가 아니라 원래 행 번호에 적용돼야 한다. → Task 12 테스트

---

## 파일 구조

```
sheet-editor/
├─ src/
│  ├─ core/
│  │  ├─ types.ts          # Cell, SheetModel, WorkbookModel, Op
│  │  ├─ refs.ts           # A1 변환, 수식·범위 참조 이동(행·열 삽입/삭제), 상대 참조 평행 이동
│  │  ├─ ops.ts            # parseInput, makeSetOp, applyOp, structuralBlockReason, cloneWorkbook, replay
│  │  ├─ display.ts        # displayText, editText, 날짜 변환
│  │  └─ view.ts           # 검색·필터 → 화면에 보일 원래 행 번호 목록
│  ├─ csv/csv.ts           # detectDialect, parseCsv, serializeCsv, csvToWorkbook, workbookToCsvRows
│  ├─ xlsx/
│  │  ├─ xml.ts            # parseXml, serializeXml (환경별 DOMParser/xmldom)
│  │  ├─ package.ts        # unzip/zip, 시트 경로 해석, rels 처리
│  │  ├─ read.ts           # readXlsx
│  │  ├─ blockers.ts       # 구조 변경 차단 이유, 표 범위
│  │  ├─ cells.ts          # 시트 DOM에서 셀 쓰기, 공유 수식 풀기
│  │  ├─ structure.ts      # 행·열 삽입/삭제를 패키지 전체에 반영
│  │  ├─ save.ts           # saveXlsx
│  │  └─ create.ts         # createXlsx (새 파일)
│  ├─ formula/recalc.ts    # recalcWorkbook
│  ├─ document/
│  │  ├─ io.ts             # DocumentIo 인터페이스와 Tauri 구현
│  │  └─ store.ts          # DocumentStore (창 하나의 문서 상태)
│  ├─ ui/                  # App, TopBar, FormulaBar, Grid, SheetTabs, SidePanel, ContextMenu
│  └─ main.tsx
├─ src-tauri/src/
│  ├─ main.rs, lib.rs
│  ├─ io.rs                # 파일 읽기/쓰기, 인코딩 판별·변환
│  ├─ windows.rs           # 문서 창 관리, 열린 파일 추적, 수정 여부
│  └─ menu.rs              # 앱 메뉴
└─ tests/                  # vitest (src 구조를 그대로 따름), tests/helpers/
```

좌표는 내부에서 모두 **0부터 시작**한다(`row 0` = 엑셀 1행, `col 0` = A열).

---

### Task 1: 프로젝트 뼈대

**Files:**
- Create: Tauri 템플릿 전체, `vitest.config.ts`, `tests/smoke.test.ts`, `README.md`
- Modify: `.gitignore`

**Interfaces:**
- Produces: `npm test`(vitest run), `npm run tauri dev`, `cargo test --manifest-path src-tauri/Cargo.toml`가 동작하는 프로젝트

- [ ] **Step 1: 기존 저장소 위에 템플릿 생성**

빈 임시 폴더에서 `npm create tauri-app@latest sheet-editor -- --template react-ts --manager npm --yes`(또는 대화형으로 react-ts, npm)로 만든 뒤, 결과물을 이 저장소 루트로 복사한다(`.git`, `docs`, `change`, `LICENSE`는 유지). `tauri.conf.json`에서 `productName`은 `Sheet Editor`, `identifier`는 `com.blooossom.sheeteditor`로 정한다. 의존성도 설치한다.
- npm: `fflate @xmldom/xmldom @glideapps/glide-data-grid fast-formula-parser @tauri-apps/plugin-dialog`
- dev: `vitest`
- Rust: `encoding_rs`, `tauri-plugin-dialog`, dev-dependency `tempfile`

`@glideapps/glide-data-grid`가 React 19와 peer 의존성이 충돌하면 `react`, `react-dom`, `@types/react*`를 18.3 버전으로 고정한다.

- [ ] **Step 2: vitest 설정과 스모크 테스트 작성**

`vitest.config.ts`에는 `environment: 'node'`, `include: ['tests/**/*.test.ts']`를 둔다. `package.json`의 `"test"`는 `vitest run`으로 설정한다.

```ts
// tests/smoke.test.ts
import { it, expect } from 'vitest';
it('vitest 동작', () => expect(1 + 1).toBe(2));
```

- [ ] **Step 3: 확인**

Run: `npm test && cargo test --manifest-path src-tauri/Cargo.toml && npm run tauri dev`
Expected: 테스트 통과. 템플릿 창이 뜬다(확인 후 종료).

- [ ] **Step 4: `.gitignore`·README 정리 후 커밋**

`.gitignore`에 `node_modules/`, `dist/`, `src-tauri/target/`, `.superpowers/`가 있는지 확인한다. README에는 앱 소개 한 줄과 위 세 명령만 적는다.

```bash
git add -A && git commit -m "chore: Tauri + React + TS 프로젝트 뼈대"
```

---

### Task 2: 셀 참조 이동 (`src/core/refs.ts`)

**Files:**
- Create: `src/core/refs.ts`, `tests/core/refs.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export type Axis = 'row' | 'col';
  export interface Shift { sheet: string; axis: Axis; at: number; count: number } // count>0 삽입, <0 삭제. at은 0부터
  export function parseA1(ref: string): { row: number; col: number };   // "C3" → {row:2,col:2}, $ 허용
  export function toA1(row: number, col: number): string;
  export function colToLetters(col: number): string;                    // 0 → "A", 27 → "AB"
  export function shiftFormula(formula: string, formulaSheet: string, shift: Shift): string;
  export function translateFormula(formula: string, dRow: number, dCol: number): string;
  export function shiftRange(range: string, axis: Axis, at: number, count: number): string | null;
  export function shiftSqref(sqref: string, axis: Axis, at: number, count: number): string | null;
  ```
  수식 문자열은 앞의 `=` 없이 다룬다.

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
import { describe, it, expect } from 'vitest';
import { parseA1, toA1, shiftFormula, translateFormula, shiftRange, shiftSqref } from '../../src/core/refs';
const S = (axis: 'row' | 'col', at: number, count: number, sheet = 'Sheet1') => ({ sheet, axis, at, count });

describe('refs', () => {
  it('A1 변환', () => {
    expect(parseA1('C3')).toEqual({ row: 2, col: 2 });
    expect(parseA1('$AB$10')).toEqual({ row: 9, col: 27 });
    expect(toA1(0, 27)).toBe('AB1');
  });
  it('행 삽입은 삽입 위치 이후 참조만 민다(절대 참조 포함)', () => {
    expect(shiftFormula('B3*C3', 'Sheet1', S('row', 1, 2))).toBe('B5*C5');
    expect(shiftFormula('A1+$B$5', 'Sheet1', S('row', 2, 1))).toBe('A1+$B$6');
  });
  it('범위 안쪽 삽입은 범위를 늘린다', () => {
    expect(shiftFormula('SUM(A1:A10)', 'Sheet1', S('row', 4, 3))).toBe('SUM(A1:A13)');
  });
  it('시트 지정 참조는 대상 시트일 때만 옮긴다', () => {
    expect(shiftFormula('Data!A5+A5', 'Sheet1', S('row', 0, 1, 'Data'))).toBe('Data!A6+A5');
    expect(shiftFormula("'My Data'!B2", 'Sheet1', S('col', 0, 1, 'My Data'))).toBe("'My Data'!C2");
  });
  it('삭제된 셀 참조는 #REF!, 범위 일부 삭제는 줄인다', () => {
    expect(shiftFormula('A3+A5', 'Sheet1', S('row', 2, -1))).toBe('#REF!+A4');
    expect(shiftFormula('SUM(A1:A10)', 'Sheet1', S('row', 2, -3))).toBe('SUM(A1:A7)');
    expect(shiftFormula('SUM(A2:A4)', 'Sheet1', S('row', 1, -3))).toBe('SUM(#REF!)');
  });
  it('문자열 상수와 함수 이름은 건드리지 않는다', () => {
    expect(shiftFormula('"A5"&LOG10(A5)', 'Sheet1', S('row', 0, 1))).toBe('"A5"&LOG10(A6)');
  });
  it('열 전체·행 전체 참조', () => {
    expect(shiftFormula('SUM(B:B)', 'Sheet1', S('col', 0, 1))).toBe('SUM(C:C)');
    expect(shiftFormula('SUM(3:3)', 'Sheet1', S('row', 0, 1))).toBe('SUM(4:4)');
  });
  it('translateFormula는 상대 부분만 옮긴다', () => {
    expect(translateFormula('A1*$B$1+B$2', 2, 1)).toBe('B3*$B$1+C$2');
  });
  it('shiftRange / shiftSqref', () => {
    expect(shiftRange('A2:C4', 'row', 0, 1)).toBe('A3:C5');
    expect(shiftRange('A2:C4', 'row', 1, -3)).toBeNull();
    expect(shiftRange('B2', 'col', 1, -1)).toBeNull();
    expect(shiftSqref('A1:A3 C5', 'row', 1, -1)).toBe('A1:A2 C4');
    expect(shiftSqref('A2', 'row', 1, -1)).toBeNull();
  });
});
```

- [ ] **Step 2: 실패 확인** — Run: `npx vitest run tests/core/refs.test.ts` / Expected: FAIL (모듈 없음)

- [ ] **Step 3: 구현**

수식은 왼쪽부터 토큰 단위로 훑는다.
- `"..."` 문자열(`""`는 이스케이프)과 `Table[...]` 구조적 참조는 건너뛴다.
- 시트 접두어(`Name!` 또는 `'Na''me'!`)를 읽은 뒤, 참조 패턴을 다음 순서로 시도한다: `셀:셀`, `열:열`, `행:행`, `셀`.
- 참조 패턴 바로 뒤에 `(`가 오거나, 앞뒤가 영문자·숫자·`_`·`.`와 붙어 있으면 참조가 아니다.
- 시트가 없으면 `formulaSheet`의 참조로 본다. 시트 이름 비교는 대소문자를 구분하지 않는다.
- 삽입 위치가 범위의 시작과 같으면 범위 전체를 옮기고, 시작보다 뒤이고 끝 이하이면 끝만 늘린다.

- [ ] **Step 4: 통과 확인** — Run: `npx vitest run tests/core/refs.test.ts` / Expected: PASS

- [ ] **Step 5: 커밋** — `git commit -m "feat(core): 셀 참조 이동"`

---

### Task 3: 문서 모델 타입과 테스트 헬퍼

**Files:**
- Create: `src/core/types.ts`, `tests/helpers/model.ts`

**Interfaces:**
- Produces:
  ```ts
  export type CellValue = string | number | boolean | null;
  export interface Cell {
    value: CellValue;      // 표시·계산용 값. 수식이면 마지막 계산값(파일의 캐시 값)
    formula?: string;      // '=' 없이
    text?: boolean;        // 원래 문자열 셀(xlsx t="s"|"inlineStr"|"str")이면 true
    styleId?: number;      // xlsx s=
    edited?: boolean;      // 마지막 저장 이후 고친 셀 (노란 표시)
  }
  export interface SheetModel {
    name: string;
    cells: Map<number, Map<number, Cell>>; // row → col → Cell
    rowCount: number;                        // 마지막 사용 행 + 1
    colCount: number;
    blockers: string[];                      // 행·열 구조 변경 차단 이유 (한국어 문구)
    tableRanges: string[];                   // 표(table) 범위 "A1:D20"
  }
  export interface WorkbookModel {
    kind: 'xlsx' | 'csv';
    sheets: SheetModel[];
    dateStyleIds: Set<number>;               // 날짜 서식인 styleId
    date1904: boolean;
  }
  export type StructuralKind = 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols';
  export type Op =
    | { kind: 'set'; sheet: number; row: number; col: number; cell: Cell | null } // null = 지우기
    | { kind: StructuralKind; sheet: number; at: number; count: number };        // count ≥ 1
  ```
  `tests/helpers/model.ts`: `makeWorkbook(kind, sheets: { name: string; grid: (string | number | null | { f: string; v?: CellValue })[][] }[]): WorkbookModel`

- [ ] **Step 1: 타입과 헬퍼 작성**
- [ ] **Step 2: 확인** — Run: `npx tsc --noEmit` / Expected: 오류 없음
- [ ] **Step 3: 커밋** — `git commit -m "feat(core): 문서 모델 타입"`

---

### Task 4: 편집 연산 (`src/core/ops.ts`)

**Files:**
- Create: `src/core/ops.ts`, `tests/core/ops.test.ts`

**Interfaces:**
- Consumes: Task 2 `shiftFormula`, `shiftRange`, `Shift`. Task 3의 모든 타입
- Produces:
  ```ts
  export function parseInput(input: string, ctx: { kind: 'xlsx' | 'csv'; prev?: Cell; isDate: boolean; date1904: boolean }): Cell | null;
  export function makeSetOp(wb: WorkbookModel, sheet: number, row: number, col: number, input: string): Op;
  export function applyOp(wb: WorkbookModel, op: Op): void;              // wb를 직접 바꾼다
  export function structuralBlockReason(wb: WorkbookModel, op: Op): string | null;
  export function cloneWorkbook(wb: WorkbookModel): WorkbookModel;
  export function replay(base: WorkbookModel, ops: Op[]): WorkbookModel; // clone 후 순서대로 applyOp
  export function getCell(sheet: SheetModel, row: number, col: number): Cell | undefined;
  ```

입력 해석 규칙(`parseInput`):
- 빈 문자열 → `null`. 이전 셀에 `styleId`가 있으면 `{ value: null, styleId, edited: true }`.
- csv → 항상 `{ value: input, text: true, edited: true }`.
- xlsx에서 `=`로 시작 → `{ formula: input.slice(1), value: null }`.
- 이전 셀이 `text`였거나, `^0\d`(앞자리 0) 또는 숫자 16자리 이상이면 → 문자열.
- `isDate`이고 `YYYY-MM-DD` 또는 `YYYY-MM-DD HH:MM(:SS)` 형식이면 → 날짜 일련번호(숫자). 1900 체계(1899-12-30 기준), `date1904`이면 1904-01-01 기준.
- `^-?\d+(\.\d+)?([eE][+-]?\d+)?$` 형식이면 → 숫자.
- 그 외 → 문자열.
- `styleId`는 이전 셀의 값을 그대로 이어받는다. `edited: true`.

구조 연산(`applyOp`):
- 셀을 옮기거나 지운다.
- **모든 시트**의 수식을 `shiftFormula`로 고친다. csv는 수식이 없으므로 셀만 옮긴다.
- 해당 시트의 `tableRanges`를 `shiftRange`로 옮긴다.
- `rowCount`/`colCount`를 갱신한다.

`structuralBlockReason`:
- 구조 연산이고 `blockers`가 있으면 `"이 시트에는 ${blockers.join(', ')}이(가) 있어 행·열을 추가/삭제할 수 없습니다."`를 반환한다.
- 열 연산이 `tableRanges`와 열 구간이 겹치면 `"표(table) 범위 안에서는 열을 추가/삭제할 수 없습니다."`를 반환한다. 열 삽입은 표의 시작 열보다 뒤이고 끝 열 이하일 때 겹친다고 본다.
- 그 외에는 `null`.

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
import { describe, it, expect } from 'vitest';
import { parseInput, makeSetOp, applyOp, structuralBlockReason, replay, getCell } from '../../src/core/ops';
import { makeWorkbook } from '../helpers/model';
const X = { kind: 'xlsx' as const, isDate: false, date1904: false };

describe('parseInput', () => {
  it('숫자처럼 보이는 문자열을 보존한다', () => {
    expect(parseInput('00123', X)).toMatchObject({ value: '00123', text: true });
    expect(parseInput('1234567890123456789', X)).toMatchObject({ value: '1234567890123456789', text: true });
    expect(parseInput('42', { ...X, prev: { value: 'x', text: true } })).toMatchObject({ value: '42', text: true });
    expect(parseInput('3.5', X)).toMatchObject({ value: 3.5 });
  });
  it('csv는 항상 문자열', () => {
    expect(parseInput('=A1', { ...X, kind: 'csv' })).toEqual({ value: '=A1', text: true, edited: true });
    expect(parseInput('007', { ...X, kind: 'csv' })).toMatchObject({ value: '007' });
  });
  it('수식, 날짜, 스타일 유지, 지우기', () => {
    expect(parseInput('=B3*C3', X)).toMatchObject({ formula: 'B3*C3', value: null });
    expect(parseInput('2026-10-23', { ...X, isDate: true })).toMatchObject({ value: 46318 });
    expect(parseInput('hi', { ...X, prev: { value: 1, styleId: 4 } })).toMatchObject({ value: 'hi', styleId: 4 });
    expect(parseInput('', { ...X, prev: { value: 1, styleId: 4 } })).toEqual({ value: null, styleId: 4, edited: true });
    expect(parseInput('', X)).toBeNull();
  });
});

describe('applyOp', () => {
  it('행 삽입: 셀과 모든 시트의 수식을 옮긴다', () => {
    const wb = makeWorkbook('xlsx', [
      { name: 'Sheet1', grid: [[1], [2], [{ f: 'A1+A2', v: 3 }]] },
      { name: 'Sum', grid: [[{ f: 'Sheet1!A3*2', v: 6 }]] },
    ]);
    applyOp(wb, { kind: 'insertRows', sheet: 0, at: 1, count: 1 });
    expect(getCell(wb.sheets[0], 1, 0)).toBeUndefined();
    expect(getCell(wb.sheets[0], 3, 0)?.formula).toBe('A1+A3');
    expect(getCell(wb.sheets[1], 0, 0)?.formula).toBe('Sheet1!A4*2');
    expect(wb.sheets[0].rowCount).toBe(4);
  });
  it('열 삭제: 셀을 당기고 지워진 참조는 #REF!', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'Sheet1', grid: [[1, 2, { f: 'A1+B1', v: 3 }]] }]);
    applyOp(wb, { kind: 'deleteCols', sheet: 0, at: 1, count: 1 });
    expect(getCell(wb.sheets[0], 0, 1)?.formula).toBe('A1+#REF!');
    expect(wb.sheets[0].colCount).toBe(2);
  });
  it('set은 edited 표시를 남기고, replay는 순차 적용과 같다', () => {
    const base = makeWorkbook('csv', [{ name: 'data', grid: [['a', 'b'], ['1', '2']] }]);
    const ops = [makeSetOp(base, 0, 1, 1, '9'), { kind: 'deleteRows' as const, sheet: 0, at: 0, count: 1 }];
    const out = replay(base, ops);
    expect(getCell(out.sheets[0], 0, 1)).toMatchObject({ value: '9', edited: true });
    expect(getCell(base.sheets[0], 1, 1)?.value).toBe('2'); // base는 그대로
  });
});

describe('structuralBlockReason', () => {
  it('차단 요소와 표 열 범위', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]);
    wb.sheets[0].blockers = ['그림·차트'];
    expect(structuralBlockReason(wb, { kind: 'insertRows', sheet: 0, at: 0, count: 1 })).toContain('그림·차트');
    wb.sheets[0].blockers = [];
    wb.sheets[0].tableRanges = ['B1:D10'];
    expect(structuralBlockReason(wb, { kind: 'deleteCols', sheet: 0, at: 2, count: 1 })).toContain('표');
    expect(structuralBlockReason(wb, { kind: 'insertCols', sheet: 0, at: 1, count: 1 })).toBeNull();
    expect(structuralBlockReason(wb, { kind: 'insertRows', sheet: 0, at: 3, count: 1 })).toBeNull();
  });
});
```

- [ ] **Step 2: 실패 확인** — Run: `npx vitest run tests/core/ops.test.ts` / Expected: FAIL
- [ ] **Step 3: 위 Interfaces와 규칙대로 구현**
- [ ] **Step 4: 통과 확인** — Run: `npx vitest run tests/core/ops.test.ts` / Expected: PASS
- [ ] **Step 5: 커밋** — `git commit -m "feat(core): 셀·행·열 편집 연산과 되돌리기용 replay"`

---

### Task 5: CSV 읽기·쓰기 (`src/csv/csv.ts`)

**Files:**
- Create: `src/csv/csv.ts`, `tests/csv/csv.test.ts`

**Interfaces:**
- Consumes: Task 3 타입
- Produces:
  ```ts
  export interface CsvDialect {
    delimiter: ',' | '\t' | ';' | '|';
    newline: '\n' | '\r\n';
    quoteAll: boolean;        // 비어 있지 않은 필드를 모두 따옴표로 감쌌는가
    quoteEmpty: boolean;      // 빈 필드를 "" 로 썼는가
    trailingNewline: boolean;
  }
  export const DEFAULT_DIALECT: CsvDialect; // ',', '\n', false, false, true
  export function detectDialect(text: string): CsvDialect;
  export function parseCsv(text: string, d: CsvDialect): string[][];
  export function serializeCsv(rows: string[][], d: CsvDialect): string;
  export function csvToWorkbook(rows: string[][], name: string): WorkbookModel; // kind 'csv', 시트 하나
  export function workbookToCsvRows(wb: WorkbookModel): string[][];
  ```

판별과 저장 규칙:
- 판별은 따옴표 밖 기준으로 앞의 최대 20줄을 본다.
  - 구분자: 줄마다 개수가 같고 1 이상인 후보 중 가장 많은 것. 없으면 `,`.
  - 줄바꿈: 처음 나온 줄바꿈의 종류.
- 저장할 때 각 행은 시트 너비(`colCount`)만큼 채운다.
- `quoteAll`이 아니면, 구분자·따옴표·`\r`·`\n`이 들어간 필드만 따옴표로 감싼다.

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
import { describe, it, expect } from 'vitest';
import { detectDialect, parseCsv, serializeCsv } from '../../src/csv/csv';

describe('csv', () => {
  it('방식 판별', () => {
    expect(detectDialect('"a","b"\r\n"1",""\r\n')).toEqual({ delimiter: ',', newline: '\r\n', quoteAll: true, quoteEmpty: true, trailingNewline: true });
    expect(detectDialect('a\tb\n1\t2')).toEqual({ delimiter: '\t', newline: '\n', quoteAll: false, quoteEmpty: false, trailingNewline: false });
    expect(detectDialect('a;b;c\n1;2;3\n').delimiter).toBe(';');
  });
  it('따옴표 안의 구분자·따옴표·줄바꿈', () => {
    const t = 'a,"b,""c""\nd"\n';
    expect(parseCsv(t, detectDialect(t))).toEqual([['a', 'b,"c"\nd']]);
  });
  it.each([
    'id,name\n001,김민지\n',
    '"id","name"\r\n"001","이준호"\r\n',
    'a,"x,y",c\n1,"say ""hi""",3',
    'a\tb\n1\t\n',
    '"a","b"\n"1",""\n',
    '',
  ])('열고 그대로 저장하면 같다: %j', (s) => {
    const d = detectDialect(s);
    expect(serializeCsv(parseCsv(s, d), d)).toBe(s);
  });
});
```

- [ ] **Step 2: 실패 확인** — Run: `npx vitest run tests/csv` / Expected: FAIL
- [ ] **Step 3: 구현**
- [ ] **Step 4: 통과 확인** — Expected: PASS
- [ ] **Step 5: 커밋** — `git commit -m "feat(csv): 방식 판별과 왕복 보존"`

---

### Task 6: Rust 파일 입출력과 인코딩 (`src-tauri/src/io.rs`)

**Files:**
- Create: `src-tauri/src/io.rs`, `src/document/io.ts`
- Modify: `src-tauri/src/lib.rs` (명령 등록)

**Interfaces:**
- Produces (Rust):
  ```rust
  pub struct Decoded { pub text: String, pub encoding: &'static str /* "utf-8" | "euc-kr" */, pub bom: bool }
  pub fn decode_text(bytes: &[u8]) -> Decoded;      // BOM → utf-8+bom, 올바른 UTF-8 → utf-8, 그 외 → euc-kr
  pub fn encode_text(text: &str, encoding: &str, bom: bool) -> Result<Vec<u8>, String>;
  pub fn write_atomic(path: &Path, bytes: &[u8]) -> Result<(), String>; // 같은 폴더 임시 파일 → rename
  #[tauri::command] fn read_file(path: String) -> Result<tauri::ipc::Response, String>;
  #[tauri::command] fn write_file(request: tauri::ipc::Request) -> Result<(), String>; // 본문=바이트, 헤더 x-path=encodeURIComponent(경로)
  #[tauri::command] fn read_text(path: String) -> Result<TextFile, String>;           // {text, encoding, bom}
  #[tauri::command] fn write_text(path: String, text: String, encoding: String, bom: bool) -> Result<(), String>;
  ```
- Produces (TS, `src/document/io.ts`):
  ```ts
  export interface TextFile { text: string; encoding: 'utf-8' | 'euc-kr'; bom: boolean }
  export interface DocumentIo {
    readFile(path: string): Promise<Uint8Array>;
    writeFile(path: string, bytes: Uint8Array): Promise<void>;
    readText(path: string): Promise<TextFile>;
    writeText(path: string, file: TextFile): Promise<void>;
  }
  export const tauriIo: DocumentIo;
  ```

`encode_text`에서 EUC-KR로 표현할 수 없는 글자가 있으면(`encoding_rs`의 `had_unmappable`) 다음 오류를 반환한다: `"EUC-KR로 저장할 수 없는 글자가 있습니다: {처음 나온 글자}"`.

- [ ] **Step 1: 실패하는 Rust 테스트 작성** (`io.rs` 아래 `#[cfg(test)] mod tests`)

```rust
#[test] fn utf8_and_bom() {
    let d = decode_text("가,b".as_bytes()); assert_eq!((d.encoding, d.bom, d.text.as_str()), ("utf-8", false, "가,b"));
    let d = decode_text(b"\xEF\xBB\xBFa"); assert_eq!((d.encoding, d.bom, d.text.as_str()), ("utf-8", true, "a"));
}
#[test] fn euc_kr_roundtrip() {
    let bytes = [0xB0, 0xA1, 0xB3, 0xAA]; // "가나"
    let d = decode_text(&bytes); assert_eq!((d.encoding, d.text.as_str()), ("euc-kr", "가나"));
    assert_eq!(encode_text("가나", "euc-kr", false).unwrap(), bytes.to_vec());
}
#[test] fn euc_kr_unmappable_is_error() {
    let e = encode_text("가😀", "euc-kr", false).unwrap_err(); assert!(e.contains("EUC-KR") && e.contains("😀"));
}
#[test] fn bom_written() { assert_eq!(encode_text("a", "utf-8", true).unwrap(), b"\xEF\xBB\xBFa".to_vec()); }
#[test] fn atomic_write_korean_path() {
    let dir = tempfile::tempdir().unwrap(); let p = dir.path().join("한글 폴더 파일.csv");
    std::fs::write(&p, b"old").unwrap(); write_atomic(&p, b"new").unwrap();
    assert_eq!(std::fs::read(&p).unwrap(), b"new");
}
```

- [ ] **Step 2: 실패 확인** — Run: `cargo test --manifest-path src-tauri/Cargo.toml io::` / Expected: FAIL (컴파일 오류)
- [ ] **Step 3: 구현하고 명령을 `invoke_handler`에 등록. `src/document/io.ts`의 `tauriIo` 작성**

`writeFile`은 `invoke('write_file', bytes, { headers: { 'x-path': encodeURIComponent(path) } })`로 보낸다. Rust 쪽에서는 헤더를 percent-decode해서 경로로 쓴다. 이 decode에 필요한 의존성으로 `percent-encoding` 크레이트를 추가한다.

- [ ] **Step 4: 통과 확인** — Run: `cargo test --manifest-path src-tauri/Cargo.toml && npx tsc --noEmit` / Expected: PASS
- [ ] **Step 5: 커밋** — `git commit -m "feat(io): 파일 입출력과 csv 인코딩 판별·보존"`

---

### Task 7: xlsx 읽기 (`src/xlsx/xml.ts`, `package.ts`, `read.ts`, `blockers.ts`)

**Files:**
- Create: `src/xlsx/xml.ts`, `src/xlsx/package.ts`, `src/xlsx/read.ts`, `src/xlsx/blockers.ts`, `tests/helpers/xlsxFixture.ts`, `tests/xlsx/read.test.ts`

**Interfaces:**
- Consumes: Task 2 `parseA1`, Task 3 타입
- Produces:
  ```ts
  // xml.ts — globalThis.DOMParser가 있으면(앱) 그것을, 없으면(vitest) @xmldom/xmldom을 쓴다
  export function parseXml(text: string): Document;
  export function serializeXml(doc: Document): string; // 기존 선언을 지우고 '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'을 앞에 붙임
  // package.ts
  export type Parts = Map<string, Uint8Array>;       // 경로는 앞의 '/' 없이 (예: 'xl/workbook.xml')
  export function unzipParts(bytes: Uint8Array): Parts;
  export function zipParts(parts: Parts): Uint8Array; // '[Content_Types].xml'을 첫 항목으로
  export function readText(parts: Parts, path: string): string | undefined;
  export function writeText(parts: Parts, path: string, text: string): void;
  export interface SheetInfo { name: string; path: string; relsPath: string }
  export function sheetInfos(parts: Parts): SheetInfo[];  // workbook.xml 순서. 절대/상대 Target 모두 처리
  export function relTargets(parts: Parts, relsPath: string, baseDir: string): { id: string; type: string; target: string }[];
  // read.ts
  export function readXlsx(bytes: Uint8Array): WorkbookModel;
  // blockers.ts
  export function findBlockers(parts: Parts, sheet: SheetInfo): string[];
  export function findTableRanges(parts: Parts, sheet: SheetInfo): string[];
  ```
  `tests/helpers/xlsxFixture.ts`:
  ```ts
  export interface FixtureSheet { name: string; xml: string /* <sheetData>…</sheetData> 및 그 뒤 요소들(mergeCells 등) */; rels?: { id: string; type: string; target: string }[] }
  export function buildXlsx(o: { sheets: FixtureSheet[]; sharedStrings?: string[]; styles?: string; workbookExtra?: string; extraParts?: Record<string, string> }): Uint8Array;
  export function partText(bytes: Uint8Array, path: string): string | undefined;
  ```
  `buildXlsx`는 `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`, workbook rels, 시트, `docProps/app.xml`, `xl/theme/theme1.xml`을 만든다. `styles`를 주지 않으면 기본 styles도 만든다. 시트의 `<worksheet>`에는 main 네임스페이스와 `xmlns:r`(relationships)를 선언한다. `workbookExtra`는 `</sheets>` 뒤(예: `<definedNames>`)에 들어간다.

읽기 규칙:
- 셀 유형별 값:
  - `t="s"` → 공유 문자열. `<r>` 안의 `<t>`는 이어 붙이고, `<rPh>` 안의 `<t>`는 제외한다.
  - `t="inlineStr"` → `<is>` 안의 텍스트.
  - `t="b"` → boolean.
  - `t="str"`·`t="e"` → 문자열.
  - 그 외 → 숫자.
- `text` 플래그는 `s`·`inlineStr`·`str`일 때 true다. 수식은 `<f>`의 텍스트다. 공유 수식의 자식(`<f t="shared" si>`의 텍스트가 빈 경우)은 대표 셀 수식을 `translateFormula`로 옮겨 채운다.
- `dateStyleIds`: `cellXfs`의 순번 중 `numFmtId`가 14–22 또는 45–47이거나, 사용자 정의 형식 코드에서 `"…"`와 `[…]`를 뺀 나머지에 `y`/`d`/`h`/`s`가 있는 것.
- `date1904`: `workbookPr date1904="1"`이면 true.
- `blockers`(이 순서·문구 그대로):
  - `'그림·차트'`: 시트 rels에 drawing이 있음
  - `'메모'`: comments, threadedComments, vmlDrawing
  - `'피벗 테이블'`: 시트 rels의 pivotTable, 또는 `pivotCacheDefinition`의 `worksheetSource sheet`가 이 시트
  - `'이 시트를 참조하는 차트'`: `xl/charts/*.xml`의 `<c:f>`에 이 시트 이름이 있음
  - `'확장 서식(스파크라인 등)'`: 시트 XML에 `<extLst>`가 있음

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
import { describe, it, expect } from 'vitest';
import { readXlsx } from '../../src/xlsx/read';
import { buildXlsx } from '../helpers/xlsxFixture';
const DATE_STYLES = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>`;

describe('readXlsx', () => {
  it('값 종류와 시트 순서', () => {
    const wb = readXlsx(buildXlsx({
      sharedStrings: ['이름', '<r><t>김</t></r><r><t>민지</t></r><rPh sb="0" eb="1"><t>キム</t></rPh>'],
      styles: DATE_STYLES,
      sheets: [
        { name: '목록', xml: `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>메모</t></is></c><c r="C1" s="1"><v>46318</v></c></row>
          <row r="3"><c r="A3" t="s"><v>1</v></c><c r="B3"><v>3.5</v></c><c r="C3" t="b"><v>1</v></c><c r="D3"><f>B3*2</f><v>7</v></c></row></sheetData>` },
        { name: 'Data 2', xml: '<sheetData/>' },
      ],
    }));
    const s = wb.sheets[0];
    expect(wb.sheets.map((x) => x.name)).toEqual(['목록', 'Data 2']);
    expect(s.cells.get(0)!.get(0)).toMatchObject({ value: '이름', text: true });
    expect(s.cells.get(2)!.get(0)!.value).toBe('김민지');
    expect(s.cells.get(0)!.get(1)!.value).toBe('메모');
    expect(s.cells.get(0)!.get(2)).toMatchObject({ value: 46318, styleId: 1 });
    expect(wb.dateStyleIds.has(1)).toBe(true);
    expect(s.cells.get(2)!.get(2)!.value).toBe(true);
    expect(s.cells.get(2)!.get(3)).toMatchObject({ formula: 'B3*2', value: 7 });
    expect([s.rowCount, s.colCount]).toEqual([3, 4]);
  });
  it('공유 수식 자식을 풀어서 읽는다', () => {
    const wb = readXlsx(buildXlsx({ sheets: [{ name: 'S', xml: `<sheetData>
      <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row>
      <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row></sheetData>` }] }));
    expect(wb.sheets[0].cells.get(1)!.get(1)!.formula).toBe('A2*2');
  });
  it('구조 변경 차단 요소', () => {
    const wb = readXlsx(buildXlsx({
      sheets: [
        { name: 'S', xml: '<sheetData/><drawing r:id="rId1"/>', rels: [{ id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing', target: '../drawings/drawing1.xml' }] },
        { name: 'T', xml: '<sheetData/>' },
      ],
      extraParts: { 'xl/drawings/drawing1.xml': '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>', 'xl/charts/chart1.xml': '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:f>T!$A$1:$A$5</c:f></c:chartSpace>' },
    }));
    expect(wb.sheets[0].blockers).toEqual(['그림·차트']);
    expect(wb.sheets[1].blockers).toEqual(['이 시트를 참조하는 차트']);
  });
});
```

- [ ] **Step 2: 실패 확인** — Run: `npx vitest run tests/xlsx/read.test.ts` / Expected: FAIL
- [ ] **Step 3: 구현** (`xlsxFixture.ts`도 이 단계에서 만든다)
- [ ] **Step 4: 통과 확인** — Expected: PASS
- [ ] **Step 5: 성능 확인 테스트 추가**

`tests/xlsx/perf.test.ts`에서 20,000행 × 10열(문자열·숫자 반반) 픽스처를 만들어 `readXlsx`가 **4초 안에** 끝나는지 확인한다. node 환경의 xmldom 기준이며, 앱에서는 네이티브 DOMParser를 쓰므로 더 빠르다. 넘으면 시트 XML 읽기를 정규식 기반 스캐너로 바꾼다.

- [ ] **Step 6: 커밋** — `git commit -m "feat(xlsx): 읽기와 구조 변경 차단 요소 판별"`

---

### Task 8: xlsx 값 저장 (`src/xlsx/cells.ts`, `src/xlsx/save.ts`)

**Files:**
- Create: `src/xlsx/cells.ts`, `src/xlsx/save.ts`, `tests/xlsx/save.test.ts`

**Interfaces:**
- Consumes: Task 7 `unzipParts`/`zipParts`/`sheetInfos`/`parseXml`/`serializeXml`, Task 2 `translateFormula`/`toA1`
- Produces:
  ```ts
  // cells.ts
  export function setCellInSheet(doc: Document, row: number, col: number, cell: Cell | null): void;
  export function expandSharedFormulas(doc: Document): void; // 모든 t="shared"를 일반 <f>로
  // save.ts
  export function saveXlsx(original: Uint8Array, ops: Op[]): Uint8Array;
  ```
  이 Task의 `saveXlsx`는 `set` 연산만 처리한다. 구조 연산을 받으면 `throw new Error('not implemented')`를 던지고, Task 9에서 채운다.

쓰기 규칙(`setCellInSheet`):
- `<row r>`와 `<c r>`이 없으면 번호 순서에 맞는 자리에 새로 만든다.
- 셀 종류별로 쓴다.
  - 숫자 → `t` 제거, `<v>`
  - 문자열 → `t="inlineStr"`, `<is><t xml:space="preserve">`
  - boolean → `t="b"`
  - 수식 → `<f>`만 쓰고 `<v>`는 제거. `t`도 제거한다.
- 이전 셀의 `s`와 다른 속성은 유지한다.
- `cell === null`이면 `<c>`를 지운다. 값이 null인데 `styleId`가 있으면 `<c r s>`만 남긴다.
- 기존 `<row>`의 다른 속성(ht, spans 등)은 건드리지 않는다.

저장 규칙(`saveXlsx`):
- 손대는 시트만 DOM으로 파싱한다. 그 시트에 `expandSharedFormulas`를 먼저 한 번 적용한다.
- `workbook.xml`의 `<calcPr>`에 `fullCalcOnLoad="1"`을 설정한다(없으면 `</workbook>` 직전의 올바른 순서 위치에 만든다).
- `xl/calcChain.xml`이 있으면 파일, workbook rels의 해당 항목, `[Content_Types].xml`의 Override를 모두 지운다.
- 나머지 파트는 원본 바이트를 그대로 둔다.

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
import { describe, it, expect } from 'vitest';
import { unzipSync } from 'fflate';
import { saveXlsx } from '../../src/xlsx/save';
import { readXlsx } from '../../src/xlsx/read';
import { buildXlsx, partText } from '../helpers/xlsxFixture';
const set = (row: number, col: number, cell: any, sheet = 0) => ({ kind: 'set' as const, sheet, row, col, cell });
const cellsXml = (xml: string) => Object.fromEntries([...xml.matchAll(/<c r="([A-Z]+\d+)"[\s\S]*?(?:\/>|<\/c>)/g)].map((m) => [m[1], m[0]]));

const RICH = () => buildXlsx({
  sharedStrings: ['a', 'b'],
  sheets: [
    { name: 'S1', xml: `<sheetData><row r="1" ht="30" customHeight="1"><c r="A1" t="s" s="0"><v>0</v></c><c r="F1"><v>1</v></c></row>
      <row r="2"><c r="A2"><v>5</v></c><c r="B2" s="0"><v>2</v></c><c r="C2"><f t="shared" ref="C2:C4" si="0">A2*2</f><v>10</v></c></row>
      <row r="3"><c r="A3"><v>6</v></c><c r="C3"><f t="shared" si="0"/><v>12</v></c></row>
      <row r="4"><c r="A4"><v>7</v></c><c r="C4"><f t="shared" si="0"/><v>14</v></c></row>
      <row r="5"><c r="A5" t="s"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells>` },
    { name: 'S2', xml: '<sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData>' },
  ],
  extraParts: { 'xl/calcChain.xml': '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C2" i="1"/></calcChain>' },
});

describe('saveXlsx 값 저장', () => {
  it('고치지 않은 파트와 셀은 그대로다', () => {
    const src = RICH();
    const out = saveXlsx(src, [set(1, 1, { value: 'hello', styleId: 0, edited: true })]);
    const a = unzipSync(src), b = unzipSync(out);
    const changed = new Set(['xl/worksheets/sheet1.xml', 'xl/workbook.xml', 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']);
    for (const p of Object.keys(a)) if (!changed.has(p)) expect(b[p], p).toEqual(a[p]);
    const before = cellsXml(partText(src, 'xl/worksheets/sheet1.xml')!), after = cellsXml(partText(out, 'xl/worksheets/sheet1.xml')!);
    for (const r of ['A1', 'F1', 'A2', 'A3', 'A5']) expect(after[r], r).toBe(before[r]);
    expect(after['B2']).toMatch(/s="0"/);
    expect(after['B2']).toMatch(/t="inlineStr"/);
    expect(partText(out, 'xl/worksheets/sheet1.xml')).toContain('<mergeCell ref="D1:E1"/>');
    expect(partText(out, 'xl/worksheets/sheet1.xml')).toMatch(/<row r="1" ht="30" customHeight="1"/);
  });
  it('calcChain 제거, fullCalcOnLoad 설정', () => {
    const out = saveXlsx(RICH(), [set(0, 5, { value: 2, edited: true })]);
    expect(partText(out, 'xl/calcChain.xml')).toBeUndefined();
    expect(partText(out, '[Content_Types].xml')).not.toContain('calcChain');
    expect(partText(out, 'xl/_rels/workbook.xml.rels')).not.toContain('calcChain');
    expect(partText(out, 'xl/workbook.xml')).toMatch(/<calcPr[^>]*fullCalcOnLoad="1"/);
  });
  it('공유 수식 대표 셀을 덮어써도 나머지 수식은 살아 있다', () => {
    const wb = readXlsx(saveXlsx(RICH(), [set(1, 2, { value: 99, edited: true })]));
    const s = wb.sheets[0];
    expect(s.cells.get(1)!.get(2)).toMatchObject({ value: 99 });
    expect(s.cells.get(1)!.get(2)!.formula).toBeUndefined();
    expect(s.cells.get(2)!.get(2)!.formula).toBe('A3*2');
    expect(s.cells.get(3)!.get(2)!.formula).toBe('A4*2');
  });
  it('없는 행·셀은 순서에 맞게 만들고, 지우기는 스타일만 남긴다', () => {
    const out = saveXlsx(RICH(), [set(0, 3, { value: 'd', edited: true }), set(5, 0, { value: 1, edited: true }), set(1, 1, { value: null, styleId: 0, edited: true }), set(4, 0, null)]);
    const xml = partText(out, 'xl/worksheets/sheet1.xml')!;
    expect(xml.indexOf('r="A1"')).toBeLessThan(xml.indexOf('r="D1"'));
    expect(xml.indexOf('r="D1"')).toBeLessThan(xml.indexOf('r="F1"'));
    expect(xml.indexOf('<row r="5"')).toBeLessThan(xml.indexOf('<row r="6"'));
    expect(cellsXml(xml)['B2']).toMatch(/^<c r="B2" s="0"\s*\/>$/);
    expect(cellsXml(xml)['A5']).toBeUndefined();
    const s = readXlsx(out).sheets[0];
    expect(s.cells.get(5)!.get(0)!.value).toBe(1);
  });
  it('수식 입력은 <f>로 쓰고 캐시 값은 비운다', () => {
    const xml = partText(saveXlsx(RICH(), [set(0, 5, { formula: 'A2+A3', value: null, edited: true })]), 'xl/worksheets/sheet1.xml')!;
    expect(cellsXml(xml)['F1']).toMatch(/<f>A2\+A3<\/f>/);
    expect(cellsXml(xml)['F1']).not.toMatch(/<v>/);
  });
});
```

- [ ] **Step 2: 실패 확인** — Run: `npx vitest run tests/xlsx/save.test.ts` / Expected: FAIL
- [ ] **Step 3: 구현**
- [ ] **Step 4: 통과 확인** — Expected: PASS
- [ ] **Step 5: 커밋** — `git commit -m "feat(xlsx): 고친 셀만 바꿔 넣는 저장"`

---

### Task 9: xlsx 행·열 추가/삭제 저장 (`src/xlsx/structure.ts`)

**Files:**
- Create: `src/xlsx/structure.ts`, `tests/xlsx/structure.test.ts`
- Modify: `src/xlsx/save.ts` (구조 연산 분기)

**Interfaces:**
- Consumes: Task 2 `shiftFormula`/`shiftRange`/`shiftSqref`, Task 8 `saveXlsx` 내부의 "열린 DOM 캐시"(경로 → Document)
- Produces:
  ```ts
  export function applyStructural(ctx: { parts: Parts; docs: Map<string, Document>; sheets: SheetInfo[] }, op: Extract<Op, { kind: StructuralKind }>): void;
  ```
  `ctx.docs`는 이미 파싱해 둔 XML 파트의 캐시다. 저장 마지막에 직렬화한다.

반영 대상과 규칙:

대상 시트에서:
- `<row r>`와 `<c r>`를 옮긴다. 삭제 범위의 행·셀은 지운다.
- `<dimension ref>`를 다시 계산한다.
- `<mergeCell>`은 `shiftRange`로 옮기고, `null`이 되면 지운 뒤 `count`를 갱신한다.
- 다음 범위는 `shiftSqref`로 옮기고, `null`이 되면 요소를 지운다:
  - `<hyperlink ref>`
  - `<conditionalFormatting sqref>` (안의 `<formula>`는 `shiftFormula`)
  - `<dataValidation sqref>` (안의 `<formula1/2>`는 `shiftFormula`)
- `<autoFilter ref>`를 옮긴다.
- 열 연산에서 `<cols><col min max>`는 다음처럼 처리한다.
  - 삽입 위치보다 뒤에 있는 구간은 민다.
  - 삽입 위치를 포함하는 구간은 `max`를 늘린다.
  - 삭제할 때는 구간을 줄이고, 빈 구간은 지운다.

모든 시트에서:
- `<f>`를 `shiftFormula(대상 시트 기준)`로 고친다.

통합 문서와 표:
- `workbook.xml`의 `<definedName>` 텍스트를 `shiftFormula`로 고친다(인쇄 영역 포함).
- 대상 시트 rels에 연결된 `xl/tables/*.xml`에서 `<table ref>`와 `<autoFilter ref>`를 `shiftRange`로 옮긴다.

요소를 지워서 `<mergeCells>`, `<hyperlinks>`, `<dataValidations>`가 비면 그 컨테이너도 지운다. `count` 속성이 있으면 갱신한다.

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
import { describe, it, expect } from 'vitest';
import { saveXlsx } from '../../src/xlsx/save';
import { buildXlsx, partText } from '../helpers/xlsxFixture';
const op = (kind: any, at: number, count: number, sheet = 0) => ({ kind, sheet, at, count });
const TABLE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const HL_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

const FIX = () => buildXlsx({
  workbookExtra: `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">S1!$A$1:$C$5</definedName><definedName name="Total">S1!$C$5</definedName></definedNames>`,
  sheets: [
    { name: 'S1', xml: `<dimension ref="A1:C5"/><cols><col min="2" max="3" width="20" customWidth="1"/></cols><sheetData>
      <row r="1"><c r="A1"><v>1</v></c></row><row r="2" ht="40" customHeight="1"><c r="A2"><v>2</v></c></row>
      <row r="5"><c r="C5"><f>SUM(A1:A2)</f><v>3</v></c></row></sheetData>
      <autoFilter ref="A1:C5"/><mergeCells count="2"><mergeCell ref="A3:B3"/><mergeCell ref="B4:C5"/></mergeCells>
      <conditionalFormatting sqref="A1:A5"><cfRule type="expression" priority="1"><formula>A1&gt;A$2</formula></cfRule></conditionalFormatting>
      <dataValidations count="1"><dataValidation type="list" sqref="B2:B5"><formula1>$A$1:$A$2</formula1></dataValidation></dataValidations>
      <hyperlinks><hyperlink ref="A2" r:id="rId9"/></hyperlinks><tableParts count="1"><tablePart r:id="rId1"/></tableParts>`,
      rels: [{ id: 'rId1', type: TABLE_REL, target: '../tables/table1.xml' }, { id: 'rId9', type: HL_REL, target: 'https://example.com' }] },
    { name: 'S2', xml: '<sheetData><row r="1"><c r="A1"><f>S1!A2*2</f><v>4</v></c></row></sheetData>' },
  ],
  extraParts: { 'xl/tables/table1.xml': '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:C5"><autoFilter ref="A1:C5"/><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>' },
});

describe('구조 저장', () => {
  it('행 2개 삽입(2행 위치)', () => {
    const out = saveXlsx(FIX(), [op('insertRows', 1, 2)]);
    const s1 = partText(out, 'xl/worksheets/sheet1.xml')!;
    expect(s1).toMatch(/<row r="4" ht="40" customHeight="1"><c r="A4"><v>2<\/v><\/c><\/row>/);
    expect(s1).toContain('<c r="C7"><f>SUM(A1:A4)</f>');
    expect(s1).toContain('<dimension ref="A1:C7"/>');
    expect(s1).toContain('<autoFilter ref="A1:C7"/>');
    expect(s1).toContain('<mergeCell ref="A5:B5"/>');
    expect(s1).toContain('<mergeCell ref="B6:C7"/>');
    expect(s1).toContain('sqref="A1:A7"');
    expect(s1).toContain('<formula>A1&gt;A$4</formula>');
    expect(s1).toContain('sqref="B4:B7"');
    expect(s1).toContain('<formula1>$A$1:$A$4</formula1>');
    expect(s1).toContain('<hyperlink ref="A4" r:id="rId9"/>');
    expect(partText(out, 'xl/worksheets/sheet2.xml')).toContain('<f>S1!A4*2</f>');
    expect(partText(out, 'xl/workbook.xml')).toContain('S1!$A$1:$C$7');
    expect(partText(out, 'xl/workbook.xml')).toContain('S1!$C$7');
    expect(partText(out, 'xl/tables/table1.xml')).toMatch(/ref="A1:C7"><autoFilter ref="A1:C7"\/>/);
  });
  it('행 삭제: 병합 일부는 줄이고 전부는 지우며, 참조는 #REF!', () => {
    const s1 = partText(saveXlsx(FIX(), [op('deleteRows', 1, 2)]), 'xl/worksheets/sheet1.xml')!;
    expect(s1).not.toContain('A3:B3');
    expect(s1).toContain('<mergeCell ref="B2:C3"/>');
    expect(s1).toContain('<mergeCells count="1">');
    expect(s1).toContain('SUM(A1:A1)');
    expect(s1).not.toContain('<hyperlink');
    expect(partText(saveXlsx(FIX(), [op('deleteRows', 1, 1)]), 'xl/worksheets/sheet2.xml')).toContain('<f>#REF!*2</f>');
  });
  it('열 삽입·삭제: 열 너비 구간', () => {
    expect(partText(saveXlsx(FIX(), [op('insertCols', 0, 1)]), 'xl/worksheets/sheet1.xml')).toContain('<col min="3" max="4" width="20" customWidth="1"/>');
    expect(partText(saveXlsx(FIX(), [op('insertCols', 2, 1)]), 'xl/worksheets/sheet1.xml')).toContain('<col min="2" max="4"');
    expect(partText(saveXlsx(FIX(), [op('deleteCols', 1, 1)]), 'xl/worksheets/sheet1.xml')).toContain('<col min="2" max="2"');
  });
  it('여러 연산을 순서대로', () => {
    const s1 = partText(saveXlsx(FIX(), [op('insertRows', 0, 1), { kind: 'set', sheet: 0, row: 0, col: 0, cell: { value: 'head', edited: true } }, op('deleteRows', 0, 1)]), 'xl/worksheets/sheet1.xml')!;
    expect(s1).toContain('<c r="C5"><f>SUM(A1:A2)</f>');
    expect(s1).not.toContain('head');
  });
});
```

- [ ] **Step 2: 실패 확인** — Run: `npx vitest run tests/xlsx/structure.test.ts` / Expected: FAIL
- [ ] **Step 3: 구현**
- [ ] **Step 4: 전체 테스트 통과 확인** — Run: `npm test` / Expected: PASS
- [ ] **Step 5: 커밋** — `git commit -m "feat(xlsx): 행·열 추가/삭제를 수식·병합·서식 범위에 반영"`

---

### Task 10: 새 xlsx 만들기 (`src/xlsx/create.ts`)

**Files:**
- Create: `src/xlsx/create.ts`, `tests/xlsx/create.test.ts`

**Interfaces:**
- Consumes: Task 7 `zipParts`, Task 8 `setCellInSheet`
- Produces: `export function createXlsx(wb: WorkbookModel): Uint8Array;` — 첫 시트만 쓴다. 시트 이름은 `Sheet1`.

`[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`(calcPr fullCalcOnLoad 포함), workbook rels, 기본 `xl/styles.xml`(셀 서식 1개), 시트를 템플릿 문자열로 만든 뒤 셀을 `setCellInSheet`로 채운다.

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
import { it, expect } from 'vitest';
import { createXlsx } from '../../src/xlsx/create';
import { readXlsx } from '../../src/xlsx/read';
import { makeWorkbook } from '../helpers/model';
it('만든 파일을 다시 읽으면 값이 같다', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'Sheet1', grid: [['이름', 3], [null, { f: 'B1*2', v: null }]] }]);
  const back = readXlsx(createXlsx(wb)).sheets[0];
  expect(back.name).toBe('Sheet1');
  expect(back.cells.get(0)!.get(0)!.value).toBe('이름');
  expect(back.cells.get(0)!.get(1)!.value).toBe(3);
  expect(back.cells.get(1)!.get(1)!.formula).toBe('B1*2');
});
```

- [ ] **Step 2–4: 실패 확인 → 구현 → 통과 확인** (`npx vitest run tests/xlsx/create.test.ts`)
- [ ] **Step 5: 커밋** — `git commit -m "feat(xlsx): 새 파일 만들기"`

---

### Task 11: 화면 표시 값과 수식 재계산 (`src/core/display.ts`, `src/formula/recalc.ts`)

**Files:**
- Create: `src/core/display.ts`, `src/formula/recalc.ts`, `tests/core/display.test.ts`, `tests/formula/recalc.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export function displayText(cell: Cell | undefined, wb: WorkbookModel): string;
  export function editText(cell: Cell | undefined, wb: WorkbookModel): string; // 수식은 '='+formula, 날짜는 ISO
  export function recalcWorkbook(wb: WorkbookModel): void; // kind==='xlsx'일 때만. 수식 셀 value 갱신
  ```

표시 규칙:
- 날짜 스타일 숫자: 정수면 `YYYY-MM-DD`, 소수면 `YYYY-MM-DD HH:MM:SS`.
- 그 외 숫자: `String(Number(v.toPrecision(15)))`.
- boolean: `TRUE`/`FALSE`.
- null: `''`.
- 날짜·백분율 외의 numFmt는 적용하지 않는다(설계 문서 6절에 따른 결정).

재계산 규칙:
- `fast-formula-parser`를 쓴다. `onCell`/`onRange`는 모델에서 값을 꺼내고, 그 칸이 수식 셀이면 먼저 계산한다(메모이즈, 순환 참조는 `#REF!`이 아니라 기존 값 유지).
- 계산 중 예외가 나거나 결과가 `#NAME?`(지원하지 않는 함수)이면 그 셀의 기존 값을 유지한다.
- 오류 값(FormulaError)은 `'#DIV/0!'` 같은 문자열로 둔다.

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
// tests/core/display.test.ts
import { it, expect } from 'vitest';
import { displayText, editText } from '../../src/core/display';
import { makeWorkbook } from '../helpers/model';
it('표시 형식', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]); wb.dateStyleIds.add(1);
  expect(displayText({ value: 46318, styleId: 1 }, wb)).toBe('2026-10-23');
  expect(displayText({ value: 46318.5, styleId: 1 }, wb)).toBe('2026-10-23 12:00:00');
  expect(editText({ value: 46318, styleId: 1 }, wb)).toBe('2026-10-23');
  expect(displayText({ value: 0.1 + 0.2 }, wb)).toBe('0.3');
  expect(displayText({ value: true }, wb)).toBe('TRUE');
  expect(editText({ value: 6, formula: 'A1*B1' }, wb)).toBe('=A1*B1');
  expect(displayText(undefined, wb)).toBe('');
});
// tests/formula/recalc.test.ts
import { it, expect } from 'vitest';
import { recalcWorkbook } from '../../src/formula/recalc';
import { applyOp, makeSetOp, getCell } from '../../src/core/ops';
import { makeWorkbook } from '../helpers/model';
it('값을 바꾸면 참조하는 수식이 다시 계산되고, 모르는 함수는 기존 값 유지', () => {
  const wb = makeWorkbook('xlsx', [
    { name: 'S', grid: [[2, 3, { f: 'A1*B1', v: 6 }, { f: 'FOOBAR(A1)', v: 7 }, { f: 'C1+Data!A1', v: 16 }]] },
    { name: 'Data', grid: [[10]] },
  ]);
  applyOp(wb, makeSetOp(wb, 0, 0, 0, '4'));
  recalcWorkbook(wb);
  expect(getCell(wb.sheets[0], 0, 2)!.value).toBe(12);
  expect(getCell(wb.sheets[0], 0, 3)!.value).toBe(7);
  expect(getCell(wb.sheets[0], 0, 4)!.value).toBe(22);
});
```

- [ ] **Step 2–4: 실패 확인 → 구현 → 통과 확인** (`npx vitest run tests/core/display.test.ts tests/formula`)
- [ ] **Step 5: 커밋** — `git commit -m "feat: 표시 형식과 수식 재계산"`

---

### Task 12: 검색·필터와 문서 상태 (`src/core/view.ts`, `src/document/store.ts`)

**Files:**
- Create: `src/core/view.ts`, `src/document/store.ts`, `tests/core/view.test.ts`, `tests/document/store.test.ts`

**Interfaces:**
- Consumes: Task 4 전부, Task 5 csv 함수, Task 6 `DocumentIo`/`TextFile`, Task 7 `readXlsx`, Task 8·9 `saveXlsx`, Task 10 `createXlsx`, Task 11 `displayText`/`recalcWorkbook`
- Produces:
  ```ts
  // view.ts
  export interface ColumnFilter { col: number; values: Set<string> } // displayText 기준
  export function buildRowView(sheet: SheetModel, wb: WorkbookModel, query: string, filter: ColumnFilter | null): number[] | null; // null = 전체
  export function distinctValues(sheet: SheetModel, wb: WorkbookModel, col: number, limit?: number /*500*/): { value: string; count: number }[]; // count 내림차순, 같으면 value 오름차순
  // store.ts
  export type Source =
    | { format: 'xlsx'; bytes: Uint8Array }
    | { format: 'csv'; dialect: CsvDialect; encoding: 'utf-8' | 'euc-kr'; bom: boolean }
    | { format: 'new' };
  export class DocumentStore {
    constructor(io: DocumentIo);
    path: string | null; source: Source; wb: WorkbookModel; version: number; // 바뀔 때마다 +1
    get dirty(): boolean;  get canUndo(): boolean;  get canRedo(): boolean;
    subscribe(fn: () => void): () => void;
    load(path: string): Promise<void>;       // 확장자로 형식 결정
    newDocument(): void;                     // 빈 xlsx 모델, source 'new'
    setCells(edits: { sheet: number; row: number; col: number; input: string }[]): void; // 한 번의 되돌리기 단위
    structural(sheet: number, kind: StructuralKind, rows: number[]): string | null; // 원래 행/열 번호 목록. 차단 시 이유 반환, 변경 없음
    undo(): void;  redo(): void;
    save(newPath?: string): Promise<void>;   // source 'new'면 newPath 필수(.xlsx/.csv)
  }
  ```

상태 규칙:
- 되돌리기 단위는 `steps: Op[][]`다. 현재 모델은 연산마다 바로 갱신한다.
- `undo`는 마지막 단위를 redo 쪽으로 옮기고, `wb = replay(base, steps.flat())`로 다시 만든 뒤 재계산한다.
- `dirty`는 `steps.length > 0`이다.
- `structural(..., rows)`는 같은 종류의 연속 구간으로 묶는다.
  - 삭제는 큰 번호부터 연산을 만든다.
  - 삽입은 `rows`의 최솟값 위치에 `rows.length`개를 넣는다.
  - 모두 한 단위다.
  - `insert*`는 "선택한 행 위", "선택한 열 왼쪽"을 기준으로 한다. 아래·오른쪽 삽입은 UI가 `at+1`로 넘긴다.
- 저장이 끝나면 다음을 처리한다.
  - `base = cloneWorkbook(wb)`로 두고, 모든 셀의 `edited`를 지운다.
  - `steps`와 redo를 비운다.
  - xlsx는 `source.bytes`를 새 바이트로 바꾼다.
  - `source 'new'`는 형식에 맞는 source로 바꾼다. csv로 저장했으면 `wb.kind`를 `'csv'`로 바꾸고 모든 셀을 `displayText` 문자열(`text: true`)로 바꾼다.
- 저장 후에는 그 이전으로 되돌릴 수 없다(설계 결정).
- 저장할 때 xlsx는 `saveXlsx(source.bytes, steps.flat())`을 `io.writeFile`로 쓴다.
- csv는 `serializeCsv(workbookToCsvRows(wb), dialect)`를 원래 encoding·bom으로 `io.writeText` 한다. 오류가 나면 그대로 throw하고 상태는 바꾸지 않는다.

- [ ] **Step 1: 실패하는 테스트 작성**

```ts
// tests/core/view.test.ts
import { it, expect } from 'vitest';
import { buildRowView, distinctValues } from '../../src/core/view';
import { makeWorkbook } from '../helpers/model';
const wb = makeWorkbook('csv', [{ name: 'd', grid: [['id', 'name', 'status'], ['1', '김민지', 'active'], ['2', '이준호', 'inactive'], ['3', '김서연', 'active']] }]);
const s = wb.sheets[0];
it('검색·필터는 원래 행 번호를 돌려준다', () => {
  expect(buildRowView(s, wb, '', null)).toBeNull();
  expect(buildRowView(s, wb, '김', null)).toEqual([1, 3]);
  expect(buildRowView(s, wb, 'ACTIVE', null)).toEqual([1, 2, 3]); // 대소문자 무시, 'inactive'도 포함
  expect(buildRowView(s, wb, '', { col: 2, values: new Set(['active']) })).toEqual([1, 3]);
  expect(buildRowView(s, wb, '서연', { col: 2, values: new Set(['active']) })).toEqual([3]);
  expect(distinctValues(s, wb, 2)).toEqual([{ value: 'active', count: 2 }, { value: 'inactive', count: 1 }, { value: 'status', count: 1 }]);
});

// tests/document/store.test.ts
import { it, expect } from 'vitest';
import { DocumentStore } from '../../src/document/store';
import type { DocumentIo, TextFile } from '../../src/document/io';
function fakeIo(files: Record<string, TextFile | Uint8Array>) {
  const written: Record<string, any> = {};
  const io: DocumentIo = {
    readFile: async (p) => files[p] as Uint8Array, readText: async (p) => files[p] as TextFile,
    writeFile: async (p, b) => { written[p] = b; }, writeText: async (p, f) => { written[p] = f; },
  };
  return { io, written };
}
const CSV = { text: '"id","name"\r\n"1","가"\r\n"2","나"\r\n"3","다"\r\n', encoding: 'euc-kr' as const, bom: false };

it('편집·되돌리기·저장이 원래 형식을 지킨다', async () => {
  const { io, written } = fakeIo({ '/t/a.csv': CSV });
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  st.setCells([{ sheet: 0, row: 1, col: 1, input: '라' }]);
  expect(st.dirty).toBe(true);
  st.undo(); expect(st.dirty).toBe(false);
  st.redo(); await st.save();
  expect(written['/t/a.csv']).toEqual({ text: '"id","name"\r\n"1","라"\r\n"2","나"\r\n"3","다"\r\n', encoding: 'euc-kr', bom: false });
  expect(st.dirty).toBe(false);
});
it('필터 상태의 여러 행 삭제는 원래 행 번호로, 한 단위로', async () => {
  const { io } = fakeIo({ '/t/a.csv': CSV });
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  st.structural(0, 'deleteRows', [1, 3]);
  expect(workbookCol(st, 1)).toEqual(['name', '나']);
  st.undo(); expect(workbookCol(st, 1)).toEqual(['name', '가', '나', '다']);
});
it('새 문서는 경로가 있어야 저장되고, 저장 후 일반 문서가 된다', async () => {
  const { io, written } = fakeIo({});
  const st = new DocumentStore(io); st.newDocument();
  st.setCells([{ sheet: 0, row: 0, col: 0, input: 'x' }]);
  await expect(st.save()).rejects.toThrow();
  await st.save('/t/new.csv');
  expect(written['/t/new.csv']).toEqual({ text: 'x\n', encoding: 'utf-8', bom: false });
  expect(st.path).toBe('/t/new.csv');
});
function workbookCol(st: DocumentStore, col: number) {
  const s = st.wb.sheets[0]; return Array.from({ length: s.rowCount }, (_, r) => s.cells.get(r)?.get(col)?.value);
}
```

- [ ] **Step 2–4: 실패 확인 → 구현 → 통과 확인** (`npx vitest run tests/core/view.test.ts tests/document`)
- [ ] **Step 5: 커밋** — `git commit -m "feat: 검색·필터와 창별 문서 상태(되돌리기·저장)"`

---

### Task 13: Rust 창·메뉴·파일 열기 (`windows.rs`, `menu.rs`)

**Files:**
- Create: `src-tauri/src/windows.rs`, `src-tauri/src/menu.rs`
- Modify: `src-tauri/src/lib.rs`, `src-tauri/tauri.conf.json`, `src-tauri/capabilities/default.json`

**Interfaces:**
- Produces (Rust 명령, 프론트엔드가 `invoke`로 호출):
  ```rust
  #[tauri::command] fn open_document(app: AppHandle, path: Option<String>) -> Result<(), String>;
  #[tauri::command] fn set_document_path(window: WebviewWindow, path: String) -> Result<(), String>;
  #[tauri::command] fn set_dirty(window: WebviewWindow, dirty: bool);
  ```
- Produces (이벤트): 포커스된 창에 `"menu"` 이벤트를 보낸다. payload는 `"new" | "open" | "save" | "close" | "undo" | "redo" | "insert" | "delete"` 중 하나다.
- 문서 창 URL: `index.html?path=<encodeURIComponent(경로)>` 또는 `index.html?new=1`. 창 라벨은 `doc-<n>`.

동작 규칙:
- `open_document(Some(p))`
  - `p`를 `canonicalize`한 경로로 이미 열린 창이 있으면 `set_focus`만 한다.
  - 열린 창이 없으면 새 창을 연다(제목은 파일 이름, 크기 1200×800).
  - 이때 "손대지 않은 새 문서 창"(path 없음, dirty 아님)이 있으면 새 창을 연 뒤 그 창을 닫는다. 그래서 앱을 처음 실행할 때 생긴 빈 창이 파일 창으로 바뀐 것처럼 보인다.
- `open_document(None)`: 새 문서 창을 연다.
- `setup`에서 새 문서 창 하나를 연다. `tauri.conf.json`의 `app.windows`는 빈 배열로 둔다.
- `RunEvent::Opened { urls }`: Finder 더블클릭으로 들어온 파일이다. `file://` URL마다 `open_document`를 호출한다.
- `RunEvent::ExitRequested`(⌘Q 포함)
  - dirty인 창이 있으면 `api.prevent_exit()` 후 모든 창에 `close()`를 호출한다.
  - 그러면 각 창의 닫기 확인이 뜨고, 모든 창이 닫히면 앱이 끝난다.
- 메뉴 구성:
  - 앱 메뉴: 기본 항목
  - 파일: 새 파일 ⌘N, 열기… ⌘O, 저장 ⌘S, 닫기 ⌘W
  - 편집: 실행 취소 ⌘Z, 다시 실행 ⌘⇧Z, 구분선, 잘라내기·복사·붙여넣기·전체 선택(predefined)
  - 시트: 행/열 삽입 ⌘⇧=, 행/열 삭제 ⌘-
- `tauri.conf.json`의 `bundle.fileAssociations`
  - `[{ "ext": ["xlsx"], "name": "Excel Workbook", "role": "Editor" }, { "ext": ["csv"], "name": "CSV", "role": "Editor" }]`
- capabilities: dialog(open/save/message) 권한과 `core:window:allow-close`, `core:window:allow-destroy`, `core:window:allow-set-title`, `core:webview:allow-internal-toggle-devtools`를 추가한다.

- [ ] **Step 1: 구현**
- [ ] **Step 2: 확인** — Run: `cargo build --manifest-path src-tauri/Cargo.toml && cargo test --manifest-path src-tauri/Cargo.toml` / Expected: 성공
- [ ] **Step 3: 커밋** — `git commit -m "feat(app): 문서별 창, 메뉴, 파일 열기 이벤트"`

---

### Task 14: 화면 — 상단 바, 입력줄, 표, 시트 탭

**Files:**
- Create: `src/ui/App.tsx`, `src/ui/TopBar.tsx`, `src/ui/FormulaBar.tsx`, `src/ui/Grid.tsx`, `src/ui/SheetTabs.tsx`, `src/ui/useStore.ts`, `src/ui/app.css`
- Modify: `src/main.tsx`(템플릿 내용 교체)

**Interfaces:**
- Consumes: Task 12 `DocumentStore`, `buildRowView`; Task 11 `displayText`/`editText`; Task 6 `tauriIo`; Task 13 명령과 `"menu"` 이벤트
- Produces:
  ```ts
  export function useStore(store: DocumentStore): number; // version을 구독해 다시 그리기
  export interface Selection { sheet: number; row: number /* 원래 행 번호 */; col: number; rows: number[]; cols: number[] } // rows/cols: 머리글로 선택된 원래 번호
  // Grid props
  { store: DocumentStore; sheet: number; rowView: number[] | null; onSelect(sel: Selection): void; onContextMenu(e: { x: number; y: number; target: 'row' | 'col' | 'cell'; sel: Selection }): void }
  ```

화면 규칙(시안 C):
- **상단 바**: 파일 이름(새 문서는 `제목 없음`), `· N행`, 수정됐으면 `● 수정됨`, 오른쪽에 `저장 ⌘S` 버튼.
- **입력줄**: 왼쪽에 셀 주소(`toA1`), 오른쪽에 `editText`. Enter를 누르면 `setCells`, Esc를 누르면 원래 값으로 돌린다.
- **표**: Glide Data Grid를 쓴다.
  - 0번 열은 "행 머리글" 열이다. 원래 행 번호+1을 보여 주고, 고정(`freezeColumns=1`)하며 편집할 수 없다. 내장 rowMarkers는 쓰지 않는다.
  - 1번 열부터 시트 열이고, 제목은 `A`, `B`…다.
  - 열 수는 `max(colCount + 5, 26)`이다.
  - 행 수는 `rowView`가 있으면 그 길이, 없으면 `max(rowCount + 50, 100)`이다.
  - `edited` 셀은 배경 `#fff6d6`이다.
  - 편집과 붙여넣기는 `onCellsEdited` → `store.setCells` 한 번으로 처리한다.
  - 열 머리글을 클릭하면 열 선택, 0번 열을 클릭하면 행 선택이다.
- **시트 탭**: xlsx이고 시트가 2개 이상일 때만 보인다. 시트를 바꾸면 검색·필터를 초기화한다.
- **창 제어**
  - 수정 여부가 바뀔 때마다 `set_dirty`를 호출하고, 창 제목을 파일 이름으로 둔다.
  - `getCurrentWindow().onCloseRequested`에서 dirty면 `message('저장하지 않은 변경 사항이 있습니다. 저장할까요?', { buttons: { yes: '저장', no: '저장 안 함', cancel: '취소' } })`로 묻는다. 저장 실패나 취소 시에는 닫지 않는다.
- **열기·저장·드래그**
  - 메뉴 `open`은 `open({ multiple: true, filters: [{ name: '스프레드시트', extensions: ['xlsx', 'csv'] }] })` 후 경로마다 `open_document`를 호출한다.
  - 메뉴 `save`에서 새 문서이면 `save({ filters: [{ name: 'Excel', extensions: ['xlsx'] }, { name: 'CSV', extensions: ['csv'] }] })`로 경로를 받는다. 이어서 `store.save(path)`, `set_document_path`를 호출한다.
  - 저장 오류는 `message(오류 문구, { kind: 'error' })`로 보여 준다.
  - 창에 파일을 끌어다 놓으면(`getCurrentWebview().onDragDropEvent`의 `drop`) 경로마다 `open_document`를 호출한다.
- **실행 취소와 다시 실행**: 메뉴 `undo`/`redo`는 포커스가 `input`/`textarea`(편집 중인 셀이나 입력줄, 검색창)에 있으면 `document.execCommand('undo'|'redo')`를 실행한다. 그렇지 않으면 `store.undo()`/`redo()`를 호출한다.

- [ ] **Step 1: 구현**
- [ ] **Step 2: 타입·테스트 확인** — Run: `npx tsc --noEmit && npm test` / Expected: 통과
- [ ] **Step 3: 앱 실행 확인** — Run: `npm run tauri dev`, 다음을 손으로 확인한다.
  - 빈 창에서 셀에 입력하면 노란색 표시와 `● 수정됨`이 나온다.
  - ⌘O로 csv를 열면 새 창에 뜨고, 빈 창은 사라진다.
  - 값을 고치고 ⌘S를 누르면 노란색이 없어지고, 파일 내용이 바뀐다(`cat`으로 확인).
  - ⌘Z로 되돌릴 수 있다.
  - 수정한 채로 ⌘W를 누르면 저장 여부를 묻는다.
  - 수식 셀의 입력줄에 `=…`가 나오고, 참조 셀을 고치면 값이 바뀐다.
  - 2만 행 csv에서 스크롤이 부드럽다.
- [ ] **Step 4: 커밋** — `git commit -m "feat(ui): 표, 입력줄, 상단 바, 시트 탭"`

---

### Task 15: 화면 — 검색·필터 패널, 행·열 메뉴

**Files:**
- Create: `src/ui/SidePanel.tsx`, `src/ui/ContextMenu.tsx`
- Modify: `src/ui/App.tsx`, `src/ui/Grid.tsx`

**Interfaces:**
- Consumes: Task 12 `buildRowView`/`distinctValues`/`ColumnFilter`, `store.structural`; Task 14 `Selection`, `onContextMenu`

규칙:
- **오른쪽 패널(폭 220px)**
  - 맨 위는 `검색` 입력이다. 입력 후 200ms가 지나면 반영한다.
  - 그 아래는 `필터` 열 선택(`A`, `B`… 드롭다운)과 값 체크박스 목록이다(`distinctValues`, 최대 500개). 500개를 넘으면 `상위 500개만 표시`라고 안내한다.
  - 체크박스는 처음에 모두 켜져 있다. 모두 켜져 있으면 필터가 없는 것으로 본다. `필터 해제` 버튼이 있다.
  - `rowView`는 검색어나 필터를 바꿀 때만 다시 계산한다. 편집하거나 행·열을 바꿔서 조건에 맞지 않게 된 행도 다음 변경 전까지는 보인다.
  - 구조 연산 뒤에는 `rowView`를 다시 계산한다(원래 행 번호가 바뀌므로).
- **우클릭 메뉴**
  - 행 머리글을 우클릭하면 `위에 행 삽입`, `아래에 행 삽입`, `행 삭제`가 나온다.
  - 열 머리글을 우클릭하면 `왼쪽에 열 삽입`, `오른쪽에 열 삽입`, `열 삭제`가 나온다.
  - 일반 셀을 우클릭하면 위 6개가 모두 나오고, 그 셀의 행·열을 대상으로 한다.
  - 여러 행·열이 선택돼 있으면 선택한 개수만큼 삽입·삭제한다.
- **단축키**: 메뉴 `insert`(⌘⇧=)와 `delete`(⌘-)는 열이 선택돼 있으면 열을, 그 밖에는 현재 선택의 행들을 대상으로 한다. 삽입은 위·왼쪽 기준이다.
- **차단**: `store.structural`이 이유를 반환하면 `message(이유, { kind: 'warning' })`로 보여 준다.

- [ ] **Step 1: 구현**
- [ ] **Step 2: 타입·테스트 확인** — Run: `npx tsc --noEmit && npm test` / Expected: 통과
- [ ] **Step 3: 앱 실행 확인** (`npm run tauri dev`)
  - 검색하면 해당 행만 보이고, 행 머리글 번호는 원래 번호다.
  - 필터로 값을 끄면 그 행이 숨겨진다.
  - 필터 상태에서 두 행을 골라 삭제하면 맞는 행이 지워진다.
  - 행·열 삽입/삭제 후 수식이 따라 바뀐다.
  - 그림이 있는 xlsx에서 행 삭제를 시도하면 안내가 뜬다.
- [ ] **Step 4: 커밋** — `git commit -m "feat(ui): 검색·필터 패널과 행·열 추가/삭제 메뉴"`

---

### Task 16: 빌드와 실제 파일 확인

**Files:**
- Modify: `README.md` (설치·기본 앱 지정 방법)

- [ ] **Step 1: 빌드** — Run: `npm run tauri build` / Expected: `src-tauri/target/release/bundle/macos/Sheet Editor.app` 생성
- [ ] **Step 2: 앱 설치와 파일 연결 확인**
  - 앱을 `/Applications`로 복사한다.
  - Finder에서 xlsx 파일을 고르고 정보 가져오기 → 다음으로 열기 → Sheet Editor → 모두 변경.
  - 더블클릭하면 앱이 열린다. 이미 열린 파일을 다시 더블클릭하면 같은 창이 앞으로 온다.
  - csv도 같은 방법으로 확인한다.
- [ ] **Step 3: 실제 Excel 파일 왕복 확인**
  - 서식, 병합, 수식, 여러 시트가 있는 실제 xlsx를 복사한다.
  - 값 수정, 행 삽입·삭제, 열 삽입을 차례로 하고 저장한다.
  - `/Applications/Microsoft Excel.app`(라이선스가 만료돼도 읽기 전용으로 열림)과 구글 스프레드시트에서 열어 본다.
  - **복구 경고가 없어야** 한다. 서식·병합·다른 시트가 그대로여야 하고, 수식 결과가 맞아야 한다.
- [ ] **Step 4: 실행 속도 확인**: 앱 아이콘 더블클릭부터 창이 뜨기까지 1초 안팎인지 확인한다.
- [ ] **Step 5: README 갱신 후 커밋·푸시** — `git commit -m "docs: 설치와 기본 앱 지정 방법" && git push`
