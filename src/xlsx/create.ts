import type { WorkbookModel } from '../core/types';
import { toA1 } from '../core/refs';
import { strToU8 } from 'fflate';
import { serializeXml, parseXml, NS_MAIN, NS_REL } from './xml';
import { zipParts, type Parts } from './package';
import { setCellInSheet } from './cells';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const CONTENT_TYPES = `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;

const ROOT_RELS = `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${REL_BASE}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

const WORKBOOK = `${XML_HEAD}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>`;

const WORKBOOK_RELS = `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${REL_BASE}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL_BASE}/styles" Target="styles.xml"/></Relationships>`;

const STYLES = `${XML_HEAD}<styleSheet xmlns="${NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

// 첫 시트만 쓰는 최소 xlsx. 새 파일은 서식이 0번뿐이라 styleId는 쓰지 않는다.
export function createXlsx(wb: WorkbookModel): Uint8Array {
  const sheet = wb.sheets[0];
  let maxRow = 0;
  let maxCol = 0;
  const filled: { row: number; col: number }[] = [];
  if (sheet) {
    for (const [row, cols] of [...sheet.cells].sort((a, b) => a[0] - b[0])) {
      for (const [col, cell] of [...cols].sort((a, b) => a[0] - b[0])) {
        if (cell.value === null && cell.formula === undefined) continue;
        filled.push({ row, col });
        maxRow = Math.max(maxRow, row);
        maxCol = Math.max(maxCol, col);
      }
    }
  }
  const ref = filled.length ? `A1:${toA1(maxRow, maxCol)}` : 'A1';
  const doc = parseXml(`${XML_HEAD}<worksheet xmlns="${NS_MAIN}"><dimension ref="${ref}"/><sheetData/></worksheet>`);
  for (const { row, col } of filled) {
    const { value, formula, text } = sheet.cells.get(row)!.get(col)!;
    setCellInSheet(doc, row, col, { value, formula, text });
  }
  const parts: Parts = new Map();
  const put = (path: string, text: string) => parts.set(path, strToU8(text));
  put('[Content_Types].xml', CONTENT_TYPES);
  put('_rels/.rels', ROOT_RELS);
  put('xl/workbook.xml', WORKBOOK);
  put('xl/_rels/workbook.xml.rels', WORKBOOK_RELS);
  put('xl/styles.xml', STYLES);
  put('xl/worksheets/sheet1.xml', serializeXml(doc));
  return zipParts(parts);
}
