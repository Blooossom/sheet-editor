import { it, expect } from 'vitest';
import { DocumentStore } from '../../src/document/store';
import type { DocumentIo, TextFile } from '../../src/document/io';
import { readXlsx } from '../../src/xlsx/read';
import { buildXlsx } from '../helpers/xlsxFixture';
import { getCell } from '../../src/core/ops';
import { colWidthPx, maxDigitWidth, rowHeightPx } from '../../src/grid/metrics';
import { effectiveXf, StyleResolver } from '../../src/grid/resolve';
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
it('값이 같은 편집은 단위를 만들지 않고, 구독자에게 알린다', async () => {
  const { io } = fakeIo({ '/t/a.csv': CSV });
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  let n = 0; const off = st.subscribe(() => n++);
  const v = st.version;
  st.setCells([{ sheet: 0, row: 1, col: 1, input: '가' }]);
  expect(st.dirty).toBe(false); expect(st.version).toBe(v); expect(n).toBe(0);
  st.structural(0, 'insertRows', [1, 2]); expect(st.wb.sheets[0].rowCount).toBe(6); expect(n).toBe(1);
  off(); st.undo(); expect(n).toBe(1);
});
it('고친 것이 없으면 저장해도 파일을 다시 쓰지 않는다 (새 문서는 예외)', async () => {
  const xlsx = buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData/>' }] });
  const { io, written } = fakeIo({ '/t/a.csv': CSV, '/t/b.xlsx': xlsx });
  const csv = new DocumentStore(io); await csv.load('/t/a.csv');
  await csv.save();
  const book = new DocumentStore(io); await book.load('/t/b.xlsx');
  await book.save();
  expect(written).toEqual({});
  const fresh = new DocumentStore(io); fresh.newDocument();
  await fresh.save('/t/new.xlsx');
  expect(Object.keys(written)).toEqual(['/t/new.xlsx']);
});
it('지원하지 않는 확장자는 거부한다', async () => {
  const st = new DocumentStore(fakeIo({}).io);
  await expect(st.load('/t/a.txt')).rejects.toThrow('지원하지 않는 파일 형식입니다');
});
function workbookCol(st: DocumentStore, col: number) {
  const s = st.wb.sheets[0]; return Array.from({ length: s.rowCount }, (_, r) => s.cells.get(r)?.get(col)?.value);
}

// ---- 보완 테스트 ----
const XLSX = () => buildXlsx({ sheets: [
  { name: 'S1', xml: '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>' },
  { name: 'S2', xml: '<sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData>' },
] });
const val = (st: DocumentStore, s: number, r: number, c: number) => getCell(st.wb.sheets[s], r, c)?.value;

it('새 문서 csv 저장: 파일과 메모리가 같은 표시 문자열', async () => {
  const { io, written } = fakeIo({});
  const st = new DocumentStore(io); st.newDocument();
  st.setCells([{ sheet: 0, row: 0, col: 0, input: '3.5' }, { sheet: 0, row: 0, col: 1, input: '=1+2' }]);
  await st.save('/t/n.csv');
  expect(written['/t/n.csv'].text).toBe('3.5,3\n');
  expect(st.wb.kind).toBe('csv');
  expect(val(st, 0, 0, 0)).toBe('3.5'); expect(val(st, 0, 0, 1)).toBe('3');
  expect(getCell(st.wb.sheets[0], 0, 1)?.formula).toBeUndefined();
});
it('저장 중에는 편집·재저장이 막히고 saving이 알려진다', async () => {
  let release!: () => void;
  const { io } = fakeIo({ '/t/a.csv': CSV });
  io.writeText = () => new Promise<void>((r) => { release = r; });
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  st.setCells([{ sheet: 0, row: 1, col: 1, input: '라' }]);
  let n = 0; st.subscribe(() => n++);
  const p = st.save();
  expect(st.saving).toBe(true); expect(n).toBe(1);
  const msg = '저장 중에는 편집할 수 없습니다.';
  expect(() => st.setCells([{ sheet: 0, row: 1, col: 1, input: 'z' }])).toThrow(msg);
  expect(() => st.structural(0, 'deleteRows', [1])).toThrow(msg);
  expect(() => st.undo()).toThrow(msg); expect(() => st.redo()).toThrow(msg);
  expect(() => st.newDocument()).toThrow(msg);
  await expect(st.load('/t/a.csv')).rejects.toThrow(msg);
  await expect(st.save()).rejects.toThrow(msg);
  release(); await p;
  expect(st.saving).toBe(false); expect(st.dirty).toBe(false);
  st.setCells([{ sheet: 0, row: 1, col: 1, input: 'y' }]);
});
it('xlsx 저장은 saveXlsx 결과를 쓰고, 이후 저장은 새 바이트 기준이다', async () => {
  const { io, written } = fakeIo({ '/t/a.xlsx': XLSX() });
  const st = new DocumentStore(io); await st.load('/t/a.xlsx');
  st.setCells([{ sheet: 0, row: 0, col: 0, input: '5' }]);
  expect(val(st, 0, 0, 1)).toBe(10); // 재계산
  await st.save();
  expect(val(st, 0, 0, 0)).toBe(5);
  expect(readXlsx(written['/t/a.xlsx']).sheets[0].cells.get(0)?.get(0)?.value).toBe(5);
  expect(st.canUndo).toBe(false); st.undo(); expect(val(st, 0, 0, 0)).toBe(5);
  expect(getCell(st.wb.sheets[0], 0, 0)?.edited).toBeUndefined();
  st.setCells([{ sheet: 0, row: 0, col: 0, input: '7' }]); await st.save();
  expect(readXlsx(written['/t/a.xlsx']).sheets[0].cells.get(0)?.get(0)?.value).toBe(7);
});
it('새 편집은 redo를 지운다', async () => {
  const { io } = fakeIo({ '/t/a.csv': CSV });
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  st.setCells([{ sheet: 0, row: 1, col: 1, input: 'x' }]); st.undo(); expect(st.canRedo).toBe(true);
  st.setCells([{ sheet: 0, row: 1, col: 1, input: 'y' }]); expect(st.canRedo).toBe(false);
});
it('저장 실패(문자열 reject)는 Error로 바뀌고 상태를 지킨다', async () => {
  const { io } = fakeIo({ '/t/a.csv': CSV });
  io.writeText = () => Promise.reject('디스크 오류');
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  st.setCells([{ sheet: 0, row: 1, col: 1, input: 'x' }]);
  const err = await st.save().catch((e) => e);
  expect(err).toBeInstanceOf(Error); expect(err.message).toBe('디스크 오류');
  expect(st.dirty).toBe(true); expect(st.saving).toBe(false);
  expect(getCell(st.wb.sheets[0], 1, 1)?.edited).toBe(true);
  st.undo(); expect(val(st, 0, 1, 1)).toBe('가');
});
it('차단된 구조 변경은 이유만 돌려주고 아무것도 바꾸지 않는다', async () => {
  const { io } = fakeIo({ '/t/a.xlsx': XLSX() });
  const st = new DocumentStore(io); await st.load('/t/a.xlsx');
  st.wb.sheets[0].blockers = ['차단 요소'];
  const v = st.version;
  expect(st.structural(0, 'deleteRows', [0])).toContain('차단 요소');
  expect(st.version).toBe(v); expect(st.dirty).toBe(false);
});
it('둘째 시트 편집과 undo 후 재계산', async () => {
  const { io } = fakeIo({ '/t/a.xlsx': XLSX() });
  const st = new DocumentStore(io); await st.load('/t/a.xlsx');
  st.setCells([{ sheet: 1, row: 0, col: 0, input: '4' }, { sheet: 0, row: 0, col: 0, input: '8' }]);
  expect(val(st, 1, 0, 0)).toBe(4); expect(val(st, 0, 0, 1)).toBe(16);
  st.undo(); expect(val(st, 1, 0, 0)).toBe(9); expect(val(st, 0, 0, 1)).toBe(2);
});
it('csv 시트 이름은 파일 기본 이름, bareQuotes 방식은 왕복된다', async () => {
  const text = 'a,12" pipe\r\nb,c';
  const { io, written } = fakeIo({ '/t/My.Data.CSV': { text, encoding: 'utf-8', bom: false } });
  const st = new DocumentStore(io); await st.load('/t/My.Data.CSV');
  expect(st.wb.sheets[0].name).toBe('My.Data');
  st.setCells([{ sheet: 0, row: 1, col: 0, input: 'B' }]); await st.save();
  expect(written['/t/My.Data.CSV'].text).toBe('a,12" pipe\r\nB,c');
});
it('떨어진 구간 삭제는 한 단위, 삽입은 최솟값 위치에 개수만큼', async () => {
  const { io } = fakeIo({ '/t/a.csv': { text: 'h\na\nb\nc\nd\n', encoding: 'utf-8', bom: false } });
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  expect(st.structural(0, 'deleteRows', [1, 3, 4])).toBeNull();
  expect(workbookCol(st, 0)).toEqual(['h', 'b']);
  st.undo(); expect(workbookCol(st, 0)).toEqual(['h', 'a', 'b', 'c', 'd']); expect(st.canUndo).toBe(false);
  st.structural(0, 'insertRows', [3, 2]);
  expect(workbookCol(st, 0)).toEqual(['h', 'a', undefined, undefined, 'b', 'c', 'd']);
  st.undo(); expect(st.canUndo).toBe(false);
});

it('undo·redo는 되돌린 단계의 연산을 돌려준다', async () => {
  const { io } = fakeIo({ '/t/a.csv': CSV });
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  expect(st.undo()).toBeNull();
  st.structural(0, 'deleteRows', [1]);
  expect(st.undo()?.map((o) => o.kind)).toEqual(['deleteRows']);
  expect(st.redo()?.map((o) => o.kind)).toEqual(['deleteRows']);
  expect(st.redo()).toBeNull();
});

// ---- 서식·크기·병합 (Task 15)
function newBook(rows: string[][] = []) {
  const { io, written } = fakeIo({});
  const st = new DocumentStore(io); st.newDocument();
  st.setCells(rows.flatMap((r, ri) => r.map((input, ci) => ({ sheet: 0, row: ri, col: ci, input }))));
  return { st, written };
}
const sid = (st: DocumentStore, r: number, c: number) => getCell(st.wb.sheets[0], r, c)?.styleId ?? 0;
const BOLD = { font: { bold: true } };

it('서식 적용은 dirty가 되고 undo로 styleId가 돌아가며 redo 된다', async () => {
  const xlsx = buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row></sheetData>' }] });
  const { io } = fakeIo({ '/t/b.xlsx': xlsx });
  const st = new DocumentStore(io); await st.load('/t/b.xlsx');
  expect(st.formatCells(0, { rows: [0], c1: 0, c2: 0 }, BOLD)).toBeNull();
  expect(st.dirty).toBe(true);
  const id = sid(st, 0, 0);
  expect(id).not.toBe(0);
  st.undo(); expect(sid(st, 0, 0)).toBe(0); expect(st.dirty).toBe(false);
  st.redo(); expect(sid(st, 0, 0)).toBe(id);
});
it('20만 칸 초과를 차단한다', () => {
  const { st } = newBook([['a']]);
  const rows = Array.from({ length: 200001 }, (_, i) => i);
  expect(st.formatCells(0, { rows, c1: 0, c2: 0 }, BOLD)).toBe('한 번에 서식을 바꿀 수 있는 칸은 200,000개까지입니다.');
  expect(st.formatCells(0, { rows: rows.slice(0, 100000), c1: 0, c2: 1 }, BOLD)).toBeNull();
});
it('병합이 많아도 넓은 서식 대상의 병합 확장이 빠르다', () => {
  for (const rows of [Array.from({ length: 100000 }, (_, i) => i), Array.from({ length: 50000 }, (_, i) => i * 2)]) {
    const { st } = newBook([['a']]);
    const merges = st.wb.sheets[0].layout.merges;
    for (let i = 0; i < 3000; i++) merges.push({ r1: 300000 + i * 2, c1: 0, r2: 300001 + i * 2, c2: 1 });
    const t0 = performance.now();
    expect(st.formatCells(0, { rows, c1: 0, c2: 0 }, BOLD)).toBeNull();
    expect(performance.now() - t0).toBeLessThan(400);
  }
});
it('필터 rows(비연속)는 보이는 행에만 적용된다', () => {
  const { st } = newBook([['a'], ['b'], ['c'], ['d']]);
  st.formatCells(0, { rows: [0, 2], c1: 0, c2: 0 }, BOLD);
  expect(sid(st, 0, 0)).not.toBe(0); expect(sid(st, 1, 0)).toBe(0);
  expect(sid(st, 2, 0)).not.toBe(0); expect(sid(st, 3, 0)).toBe(0);
});
it('행·열 전체 선택은 사용 범위로 줄인다', () => {
  const { st } = newBook([['a', 'b'], ['c', 'd']]);
  expect(st.formatCells(0, { rows: 'all', c1: 0, c2: 0 }, BOLD)).toBeNull();
  expect(st.wb.sheets[0].rowCount).toBe(2);
  expect(st.formatCells(0, { rows: [0], c1: 0, c2: 'all' }, BOLD)).toBeNull();
  expect(st.wb.sheets[0].colCount).toBe(2);
  expect(sid(st, 0, 1)).not.toBe(0);
});
it('빈 시트의 전체 선택은 단계를 만들지 않는다', () => {
  const { st } = newBook();
  expect(st.formatCells(0, { rows: 'all', c1: 0, c2: 0 }, BOLD)).toBeNull();
  expect(st.dirty).toBe(false);
});
it('빈 칸도 일반 선택이면 서식이 적용된다', () => {
  const { st } = newBook();
  expect(st.formatCells(0, { rows: [0, 1], c1: 0, c2: 1 }, BOLD)).toBeNull();
  expect(st.dirty).toBe(true); expect(sid(st, 1, 1)).not.toBe(0);
});
it('병합 일부를 덮으면 열은 c1·c2, 행은 병합 행까지 넓힌다(필터가 없을 때)', () => {
  const { st } = newBook([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']]);
  st.merge(0, { r1: 0, c1: 1, r2: 1, c2: 2 }, 'all');
  st.formatCells(0, { rows: [0], c1: 2, c2: 2 }, BOLD);
  for (const [r, c] of [[0, 1], [0, 2], [1, 1], [1, 2]]) expect(sid(st, r, c)).not.toBe(0);
  expect(sid(st, 0, 0)).toBe(0); expect(sid(st, 2, 2)).toBe(0);
});
it('필터(비연속 rows)에서는 병합 행을 더하지 않고 열만 넓힌다', () => {
  const { st } = newBook([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i'], ['j', 'k', 'l']]);
  st.merge(0, { r1: 2, c1: 1, r2: 3, c2: 2 }, 'all');
  st.formatCells(0, { rows: [0, 2], c1: 2, c2: 2 }, BOLD);
  expect(sid(st, 0, 2)).not.toBe(0);
  expect(sid(st, 2, 1)).not.toBe(0); expect(sid(st, 2, 2)).not.toBe(0);
  expect(sid(st, 3, 1)).toBe(0); expect(sid(st, 3, 2)).toBe(0);
  expect(sid(st, 1, 1)).toBe(0); // 안 보이는 행은 그대로
});
it('비연속 rows로 병합하면 차단된다', () => {
  const { st } = newBook([['a'], ['b'], ['c']]);
  const r = { r1: 0, c1: 0, r2: 2, c2: 1 };
  expect(st.mergeCheck(0, r, [0, 2])).toEqual({ ok: false, reason: '필터를 해제한 뒤 병합하세요.' });
  expect(st.merge(0, r, [0, 2])).toBe('필터를 해제한 뒤 병합하세요.');
  expect(st.wb.sheets[0].layout.merges).toEqual([]);
  expect(st.merge(0, r, [0, 1, 2])).toBeNull();
});
it('병합은 lossy를 판정하고 앵커 외 값을 지우며 undo 한 번으로 복원된다', () => {
  const { st } = newBook([['a', 'b'], ['c', '=1+1']]);
  const r = { r1: 0, c1: 0, r2: 1, c2: 1 };
  expect(st.mergeCheck(0, r, 'all')).toEqual({ ok: true, lossy: true });
  st.formatCells(0, { rows: [0], c1: 1, c2: 1 }, BOLD);
  const bold = sid(st, 0, 1);
  expect(st.merge(0, r, 'all')).toBeNull();
  const sh = () => st.wb.sheets[0];
  expect(sh().layout.merges).toEqual([r]);
  expect(getCell(sh(), 0, 0)?.value).toBe('a');
  expect(getCell(sh(), 0, 1)?.value ?? null).toBeNull();
  expect(getCell(sh(), 1, 0)).toEqual({ value: null, styleId: 0, edited: true });
  expect(getCell(sh(), 1, 1)?.formula).toBeUndefined();
  expect(sid(st, 0, 1)).toBe(bold); // 지워도 서식은 남는다
  st.undo();
  expect(sh().layout.merges).toEqual([]);
  expect(getCell(sh(), 0, 1)?.value).toBe('b');
  expect(getCell(sh(), 1, 0)?.value).toBe('c');
  expect(getCell(sh(), 1, 1)?.formula).toBe('1+1');
});
it('필터 rows가 병합으로 넓힌 대상을 덮지 못하면 차단된다', () => {
  const { st } = newBook();
  st.merge(0, { r1: 0, c1: 1, r2: 2, c2: 2 }, 'all');
  // A1:B1 선택(보이는 행 [0]) → 병합 B1:C3까지 넓어지므로 행 1·2가 안 보이면 막는다
  const r = { r1: 0, c1: 0, r2: 0, c2: 1 };
  expect(st.merge(0, r, [0])).toBe('필터를 해제한 뒤 병합하세요.');
  expect(st.merge(0, r, [0, 1, 2])).toBeNull();
});
it('빈 patch·크기 한계·잘못된 크기', () => {
  const { st } = newBook([['a']]);
  st.undo();
  expect(st.formatCells(0, { rows: [0], c1: 0, c2: 0 }, {})).toBeNull();
  expect(st.dirty).toBe(false);
  st.setColWidth(0, [0], 100000);
  expect(st.wb.sheets[0].layout.colRuns[0].info.width).toBe(255);
  st.setRowHeight(0, [0], 100000);
  expect(st.wb.sheets[0].layout.rows.get(0)!.height).toBe(409.5);
  expect(st.setColWidth(0, [0], Infinity)).toBe('올바른 크기가 아닙니다.');
  expect(st.setRowHeight(0, [0], NaN)).toBe('올바른 크기가 아닙니다.');
});
it('값이 없는 병합은 lossy가 아니고, 한 칸 병합은 단계를 만들지 않는다', () => {
  const { st } = newBook();
  expect(st.mergeCheck(0, { r1: 0, c1: 0, r2: 1, c2: 1 }, 'all')).toEqual({ ok: true, lossy: false });
  expect(st.merge(0, { r1: 0, c1: 0, r2: 0, c2: 0 }, 'all')).toBeNull();
  expect(st.dirty).toBe(false);
});
it('병합을 일부만 덮는 범위는 그 병합까지 넓혀서 합친다', () => {
  const { st } = newBook();
  st.merge(0, { r1: 0, c1: 1, r2: 1, c2: 2 }, 'all');
  st.merge(0, { r1: 0, c1: 0, r2: 0, c2: 1 }, 'all'); // B1:C2의 일부와 겹친다 → A1:C2
  expect(st.wb.sheets[0].layout.merges).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 2 }]);
});
it('표(table)와 겹치면 병합을 막는다', () => {
  const { st } = newBook([['a', 'b'], ['c', 'd']]);
  st.wb.sheets[0].tables = [{ ref: 'B2:C5', headerRows: 1, totalsRows: 0 }];
  const r = { r1: 0, c1: 0, r2: 1, c2: 1 };
  expect(st.mergeCheck(0, r, 'all')).toEqual({ ok: false, reason: '표(table) 안에서는 병합할 수 없습니다.' });
  expect(st.merge(0, r, 'all')).toBe('표(table) 안에서는 병합할 수 없습니다.');
  expect(st.merge(0, { r1: 0, c1: 0, r2: 0, c2: 1 }, 'all')).toBeNull();
});
it('unmerge는 겹치는 병합을 모두 풀고, 없으면 단계가 없다', () => {
  const { st } = newBook();
  st.merge(0, { r1: 0, c1: 0, r2: 1, c2: 1 }, 'all');
  st.merge(0, { r1: 3, c1: 0, r2: 4, c2: 1 }, 'all');
  expect(st.unmerge(0, { r1: 1, c1: 1, r2: 3, c2: 1 })).toBeNull();
  expect(st.wb.sheets[0].layout.merges).toEqual([]);
  st.undo();
  expect(st.wb.sheets[0].layout.merges).toHaveLength(2);
  const n = st.version;
  expect(st.unmerge(0, { r1: 8, c1: 8, r2: 9, c2: 9 })).toBeNull();
  expect(st.version).toBe(n);
});
it('csv에서는 서식·크기·병합이 차단된다', async () => {
  const { io } = fakeIo({ '/t/a.csv': CSV });
  const st = new DocumentStore(io); await st.load('/t/a.csv');
  const msg = '이 형식(csv)은 서식을 저장할 수 없습니다.';
  expect(st.formatCells(0, { rows: [0], c1: 0, c2: 0 }, BOLD)).toBe(msg);
  expect(st.setColWidth(0, [0], 100)).toBe(msg);
  expect(st.setRowHeight(0, [0], 30)).toBe(msg);
  expect(st.merge(0, { r1: 0, c1: 0, r2: 1, c2: 1 }, 'all')).toBe(msg);
  expect(st.mergeCheck(0, { r1: 0, c1: 0, r2: 1, c2: 1 }, 'all')).toEqual({ ok: false, reason: msg });
  expect(st.unmerge(0, { r1: 0, c1: 0, r2: 1, c2: 1 })).toBe(msg);
  expect(st.dirty).toBe(false); expect(st.hasFormatting).toBe(false);
});
it('열 너비·행 높이: px를 단위로 바꿔 연속 구간마다 Op, 한 단계', () => {
  const { st } = newBook();
  const mdw = maxDigitWidth(st.wb.styles.fonts[0]);
  expect(st.setColWidth(0, [4, 0, 1, 2], 120)).toBeNull();
  const runs = st.wb.sheets[0].layout.colRuns;
  expect(runs.map((r) => [r.c1, r.c2])).toEqual([[0, 2], [4, 4]]);
  expect(colWidthPx(runs[0].info.width!, mdw)).toBe(120);
  expect(st.undo()?.map((o) => o.kind)).toEqual(['colWidth', 'colWidth']);
  expect(st.wb.sheets[0].layout.colRuns).toEqual([]);
  expect(st.setRowHeight(0, [3, 1], 40)).toBeNull();
  expect(rowHeightPx(st.wb.sheets[0].layout.rows.get(1)!.height!)).toBe(40);
  expect(st.undo()?.map((o) => o.kind)).toEqual(['rowHeight']);
  expect(st.wb.sheets[0].layout.rows.size).toBe(0);
  expect(st.setColWidth(0, [], 100)).toBeNull();
  expect(st.dirty).toBe(false);
});
it('hasFormatting: 새 문서에서 서식·크기·병합이 생기면 true, 되돌리면 false', () => {
  const { st } = newBook([['a', 'b']]);
  expect(st.hasFormatting).toBe(false);
  st.formatCells(0, { rows: [0], c1: 0, c2: 0 }, BOLD); expect(st.hasFormatting).toBe(true);
  st.undo(); expect(st.hasFormatting).toBe(false);
  st.setColWidth(0, [0], 100); expect(st.hasFormatting).toBe(true); st.undo();
  st.setRowHeight(0, [0], 40); expect(st.hasFormatting).toBe(true); st.undo();
  st.merge(0, { r1: 0, c1: 0, r2: 1, c2: 0 }, 'all'); expect(st.hasFormatting).toBe(true);
});
it('새 문서의 서식·병합·너비가 xlsx 저장 후에도 남는다', async () => {
  const { st, written } = newBook([['a', 'b']]);
  st.formatCells(0, { rows: [0], c1: 0, c2: 0 }, BOLD);
  st.merge(0, { r1: 0, c1: 0, r2: 1, c2: 1 }, 'all');
  st.setColWidth(0, [0], 120);
  await st.save('/t/n.xlsx');
  const re = readXlsx(written['/t/n.xlsx']);
  expect(re.sheets[0].layout.merges).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 1 }]);
  expect(re.sheets[0].layout.colRuns).toHaveLength(1);
  expect(getCell(re.sheets[0], 0, 0)?.styleId ?? 0).not.toBe(0);
  expect(getCell(re.sheets[0], 0, 1)?.value ?? null).toBeNull();
});
it('저장 중에는 서식 편집이 막힌다', async () => {
  const { st } = newBook([['a']]);
  const p = st.save('/t/n.xlsx');
  expect(() => st.formatCells(0, { rows: [0], c1: 0, c2: 0 }, BOLD)).toThrow('저장 중');
  await p;
});

// ---- 병합하고 가운데 맞춤 (Task 16)
it('merge에 서식을 주면 병합과 가운데 맞춤이 되돌리기 한 단계다', () => {
  const { st } = newBook([['a', 'b'], ['c', 'd']]);
  const r = { r1: 0, c1: 0, r2: 1, c2: 1 };
  expect(st.merge(0, r, 'all', { align: { horizontal: 'center' } })).toBeNull();
  expect(st.wb.sheets[0].layout.merges).toEqual([r]);
  for (const [row, col] of [[0, 0], [0, 1], [1, 0], [1, 1]]) {
    expect(st.wb.styles.xfs[sid(st, row, col)].alignment?.horizontal).toBe('center');
  }
  st.undo();
  expect(st.wb.sheets[0].layout.merges).toEqual([]);
  expect(sid(st, 0, 0)).toBe(0);
  expect(getCell(st.wb.sheets[0], 1, 1)?.value).toBe('d');
});
it('병합하고 가운데 맞춤을 저장하면 다시 읽은 병합·서식이 화면 모델과 같다', async () => {
  const xlsx = buildXlsx({ sheets: [{ name: 'S', xml:
    '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1" t="inlineStr"><is><t>b</t></is></c></row>' +
    '<row r="2"><c r="A2" t="inlineStr"><is><t>c</t></is></c></row></sheetData>' }] });
  const { io, written } = fakeIo({ '/t/m.xlsx': xlsx });
  const st = new DocumentStore(io); await st.load('/t/m.xlsx');
  expect(st.merge(0, { r1: 0, c1: 0, r2: 1, c2: 2 }, 'all', { align: { horizontal: 'center' } })).toBeNull();
  const model = st.wb;
  await st.save();
  const re = readXlsx(written['/t/m.xlsx']);
  expect(re.sheets[0].layout.merges).toEqual(model.sheets[0].layout.merges);
  const a = new StyleResolver(model);
  const b = new StyleResolver(re);
  for (let r = 0; r <= 2; r++) for (let c = 0; c <= 3; c++) {
    expect(b.get(effectiveXf(re.sheets[0], r, c))).toEqual(a.get(effectiveXf(model.sheets[0], r, c)));
    expect(getCell(re.sheets[0], r, c)?.value ?? null).toEqual(getCell(model.sheets[0], r, c)?.value ?? null);
  }
  expect(b.get(effectiveXf(re.sheets[0], 1, 2)).hAlign).toBe('center');
});
