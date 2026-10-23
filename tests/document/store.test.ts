import { it, expect } from 'vitest';
import { DocumentStore } from '../../src/document/store';
import type { DocumentIo, TextFile } from '../../src/document/io';
import { readXlsx } from '../../src/xlsx/read';
import { buildXlsx } from '../helpers/xlsxFixture';
import { getCell } from '../../src/core/ops';
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
