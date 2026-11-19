/* 엑셀 원본 보존 저장: 지정한 셀의 값만 바꾸고 나머지(서식, 메모, 스레드 메모, 그림)는 그대로.
 * JSZip 필요.  patchCells(JSZip, bytes, sheetName, [{ref:'AA3', value:'반영'}, ...]) -> Uint8Array
 */
(function (root) {
  'use strict';
  const esc = (s) => String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colNum = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  async function sheetPath(zip, sheetName) {
    const wb = await zip.file('xl/workbook.xml').async('string');
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const sheets = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => m[0]);
    const tag = sheets.find((t) => (t.match(/\bname="([^"]*)"/) || [])[1] === esc(sheetName));
    if (!tag) throw new Error(`'${sheetName}' 시트를 찾지 못했어요`);
    const rid = tag.match(/\br:id="([^"]+)"/)[1];
    const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0])
      .find((t) => (t.match(/\bId="([^"]+)"/) || [])[1] === rid);
    let target = rel.match(/\bTarget="([^"]+)"/)[1];
    target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    return target;
  }

  function cellXml(ref, attrs, value) {
    const s = (attrs.match(/\bs="(\d+)"/) || [])[0];
    const base = `<c r="${ref}"${s ? ' ' + s : ''}`;
    if (value === null || value === undefined || value === '') return base + '/>';
    return `${base} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  }

  function patchRow(rowXml, ref, value) {
    const cellRe = new RegExp(`<c r="${reEsc(ref)}"((?:\\s[^>]*?)?)(?:/>|>[\\s\\S]*?</c>)`);
    const m = rowXml.match(cellRe);
    if (m) return rowXml.replace(cellRe, () => cellXml(ref, m[1] || '', value));
    // 셀이 없으면 열 순서에 맞게 끼워 넣기
    const want = colNum(ref.match(/^[A-Z]+/)[0]);
    const newCell = cellXml(ref, '', value);
    const cells = [...rowXml.matchAll(/<c r="([A-Z]+)\d+"/g)];
    const after = cells.find((c) => colNum(c[1]) > want);
    if (after) return rowXml.slice(0, after.index) + newCell + rowXml.slice(after.index);
    if (/\/>$/.test(rowXml)) return rowXml.replace(/\/>$/, `>${newCell}</row>`);
    return rowXml.replace(/<\/row>$/, newCell + '</row>');
  }

  async function patchCells(JSZip, bytes, sheetName, edits) {
    const zip = await JSZip.loadAsync(bytes);
    const path = await sheetPath(zip, sheetName);
    let xml = await zip.file(path).async('string');
    const byRow = new Map();
    for (const e of edits) {
      const r = Number(e.ref.match(/\d+$/)[0]);
      if (!byRow.has(r)) byRow.set(r, []);
      byRow.get(r).push(e);
    }
    for (const [r, list] of byRow) {
      const rowRe = new RegExp(`<row r="${r}"(?:\\s[^>]*?)?(?:/>|>[\\s\\S]*?</row>)`);
      const m = xml.match(rowRe);
      if (!m) throw new Error(`${r}행을 찾지 못했어요`);
      let row = m[0];
      for (const e of list) row = patchRow(row, e.ref, e.value);
      xml = xml.replace(rowRe, () => row);
    }
    zip.file(path, xml, { createFolders: false });
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  root.XlsxPatch = { patchCells };
})(typeof window !== 'undefined' ? window : globalThis);
