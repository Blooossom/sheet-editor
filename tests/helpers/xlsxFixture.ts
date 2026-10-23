import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

export interface FixtureSheet { name: string; xml: string /* <sheetData>…</sheetData> 및 그 뒤 요소들(mergeCells 등) */; rels?: { id: string; type: string; target: string }[] }

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const DEFAULT_STYLES = `<styleSheet xmlns="${NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;

const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements/></a:theme>`;

const APP = `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>`;

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// extraParts 경로 → Override 콘텐츠 타입 (알려진 종류만)
function overrideTypeFor(path: string): string | undefined {
  if (/^xl\/tables\/[^/]+\.xml$/.test(path)) return `${CT}.spreadsheetml.table+xml`;
  if (/^xl\/drawings\/[^/]+\.xml$/.test(path)) return `${CT}.drawing+xml`;
  if (/^xl\/charts\/[^/]+\.xml$/.test(path)) return `${CT}.drawingml.chart+xml`;
  if (path === 'xl/calcChain.xml') return `${CT}.spreadsheetml.calcChain+xml`;
  if (/^xl\/comments[^/]*\.xml$/.test(path)) return `${CT}.spreadsheetml.comments+xml`;
  if (/^xl\/pivotTables\/[^/]+\.xml$/.test(path)) return `${CT}.spreadsheetml.pivotTable+xml`;
  if (/^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(path)) return `${CT}.spreadsheetml.pivotCacheDefinition+xml`;
  return undefined;
}

export function buildXlsx(o: { sheets: FixtureSheet[]; sharedStrings?: string[]; styles?: string; workbookExtra?: string; extraParts?: Record<string, string> }): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const put = (path: string, text: string) => { files[path] = strToU8(text); };
  const extra = o.extraParts ?? {};

  const overrides: [string, string][] = [
    ['/xl/workbook.xml', `${CT}.spreadsheetml.sheet.main+xml`],
  ];
  o.sheets.forEach((_s, i) => overrides.push([`/xl/worksheets/sheet${i + 1}.xml`, `${CT}.spreadsheetml.worksheet+xml`]));
  overrides.push(['/xl/styles.xml', `${CT}.spreadsheetml.styles+xml`], ['/xl/theme/theme1.xml', `${CT}.theme+xml`]);
  if (o.sharedStrings) overrides.push(['/xl/sharedStrings.xml', `${CT}.spreadsheetml.sharedStrings+xml`]);
  overrides.push(['/docProps/app.xml', `${CT}.extended-properties+xml`]);
  for (const path of Object.keys(extra)) {
    const type = overrideTypeFor(path);
    if (type) overrides.push([`/${path}`, type]);
  }

  put('[Content_Types].xml', `${HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('')}</Types>`);
  put('_rels/.rels', `${HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`);

  const wbRels = o.sheets.map((_s, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`);
  let next = o.sheets.length + 1;
  const addRel = (type: string, target: string) => wbRels.push(`<Relationship Id="rId${next++}" Type="${REL}/${type}" Target="${target}"/>`);
  addRel('styles', 'styles.xml');
  addRel('theme', 'theme/theme1.xml');
  if (o.sharedStrings) addRel('sharedStrings', 'sharedStrings.xml');
  if ('xl/calcChain.xml' in extra) addRel('calcChain', 'calcChain.xml');
  put('xl/_rels/workbook.xml.rels', `${HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels.join('')}</Relationships>`);

  put('xl/workbook.xml', `${HEADER}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets>${o.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${o.workbookExtra ?? ''}</workbook>`);

  o.sheets.forEach((s, i) => {
    put(`xl/worksheets/sheet${i + 1}.xml`, `${HEADER}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">${s.xml}</worksheet>`);
    if (s.rels) {
      put(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, `${HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${s.rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"/>`).join('')}</Relationships>`);
    }
  });

  put('xl/styles.xml', HEADER + (o.styles ?? DEFAULT_STYLES));
  put('xl/theme/theme1.xml', HEADER + THEME);
  put('docProps/app.xml', HEADER + APP);
  if (o.sharedStrings) {
    const items = o.sharedStrings.map((s) => (s.startsWith('<') ? `<si>${s}</si>` : `<si><t xml:space="preserve">${esc(s)}</t></si>`));
    put('xl/sharedStrings.xml', `${HEADER}<sst xmlns="${NS_MAIN}" count="${items.length}" uniqueCount="${items.length}">${items.join('')}</sst>`);
  }
  for (const [path, text] of Object.entries(extra)) put(path, text);
  return zipSync(files);
}

export function partText(bytes: Uint8Array, path: string): string | undefined {
  const data = unzipSync(bytes)[path];
  return data === undefined ? undefined : strFromU8(data);
}
