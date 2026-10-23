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
