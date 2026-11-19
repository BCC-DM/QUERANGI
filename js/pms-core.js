/* PMS Expectedness 판정 엔진 (브라우저 + Node 공용)
 * Python 버전(meddra.py, label_parser.py, judge.py)과 같은 규칙.
 *
 * 쓰는 법
 *   const md = PMS.buildMedDRA('29.0', {enPt, enMdhier, krPt, krLlt});   // ASCII 파일 내용(문자열)
 *   const label = await PMS.readLabelPdf(pdfjsLib, pdfBytes);           // 허가사항 PDF
 *   const {terms, unmapped} = PMS.labelTerms(label, md, termMap, scope);
 *   const [expt, ref] = PMS.judgeRow(md, terms, ptName, special);
 */
(function (root) {
  'use strict';

  // ------------------------------------------------------------------ 공통
  const norm = (s) => String(s || '').replace(/\s+/g, '').toLowerCase();
  const grams = (s) => {
    const g = new Set();
    for (let i = 0; i < s.length - 1; i++) g.add(s.slice(i, i + 2));
    if (!g.size) g.add(s);
    return g;
  };
  const splitAsc = (text) => text.split(/\r?\n/).filter(Boolean).map((l) => l.split('$'));

  // ------------------------------------------------------------------ MedDRA
  /* files: {enPt, enMdhier, krPt, krLlt} — 각 ASCII 파일의 텍스트 */
  function buildMedDRA(version, files) {
    const md = {
      version, ptEn: new Map(), ptKr: new Map(), ptByEn: new Map(), lltKr: new Map(),
      ptHlts: new Map(), hltName: new Map(), ptPrimarySoc: new Map(), ptPrimaryHlt: new Map(),
      hltSize: new Map(), gramIndex: new Map(), maxLltLen: 0,
    };
    for (const r of splitAsc(files.enPt)) { md.ptEn.set(r[0], r[1]); md.ptByEn.set(norm(r[1]), r[0]); }
    for (const r of splitAsc(files.krPt)) md.ptKr.set(r[0], r[1]);
    for (const r of splitAsc(files.krLlt)) {
      const [code, name, pt] = r; const cur = r[9]; const key = norm(name);
      const old = md.lltKr.get(key);
      if (!old || (old[2] !== 'Y' && cur === 'Y')) md.lltKr.set(key, [pt, name, cur]);
    }
    for (const r of splitAsc(files.enMdhier)) {
      const pt = r[0], hlt = r[1], soc = r[3];
      if (!md.ptHlts.has(pt)) md.ptHlts.set(pt, new Set());
      md.ptHlts.get(pt).add(hlt);
      md.hltName.set(hlt, r[5]);
      if (r[11] === 'Y') { md.ptPrimarySoc.set(pt, soc); md.ptPrimaryHlt.set(pt, hlt); }
    }
    for (const hs of md.ptHlts.values()) for (const h of hs) md.hltSize.set(h, (md.hltSize.get(h) || 0) + 1);
    for (const key of md.lltKr.keys()) {
      if (key.length > md.maxLltLen) md.maxLltLen = key.length;
      for (const g of grams(key)) {
        if (!md.gramIndex.has(g)) md.gramIndex.set(g, []);
        md.gramIndex.get(g).push(key);
      }
    }
    md.ptCode = (name) => md.ptByEn.get(norm(name)) || null;
    md.krTermToPt = (term) => { const h = md.lltKr.get(norm(term)); return h ? [h[0], h[1]] : null; };
    md.suggest = (term, n = 3) => suggest(md, term, n);
    return md;
  }

  /* 사전 파일(쿼랑이 meddra/*.json)용으로 필요한 칸만 남기기. buildMedDRA가 그대로 읽을 수 있는 형식 */
  function slimMedDRA(version, files) {
    const keep = (text, cols, width) => text.split(/\r?\n/).filter(Boolean).map((l) => {
      const r = l.split('$'); const o = new Array(width).fill('');
      cols.forEach((c) => { o[c] = r[c] || ''; }); return o.join('$');
    }).join('\n');
    return {
      format: 'querangi-meddra-1', version, builtAt: new Date().toISOString().slice(0, 10),
      enPt: keep(files.enPt, [0, 1], 2), krPt: keep(files.krPt, [0, 1], 2),
      krLlt: keep(files.krLlt, [0, 1, 2, 9], 10), enMdhier: keep(files.enMdhier, [0, 1, 3, 5, 11], 12),
    };
  }

  function suggest(md, term, n) {
    const q = norm(term);
    const variants = new Set([q, q.replace(/수치/g, ''), q.replace(/혈중/g, '혈액'),
      q.replace(/비정상/g, '이상'), q.replace(/혈중/g, '')]);
    const best = new Map();
    for (const v of variants) {
      const gq = grams(v);
      const cands = new Set();
      for (const g of gq) for (const k of (md.gramIndex.get(g) || [])) cands.add(k);
      for (const key of cands) {
        const gk = grams(key);
        let inter = 0; for (const g of gq) if (gk.has(g)) inter++;
        const score = inter / (gq.size + gk.size - inter);
        const [pt, name] = md.lltKr.get(key);
        if (score > (best.has(pt) ? best.get(pt)[0] : 0)) best.set(pt, [score, name]);
      }
    }
    return [...best.entries()].sort((a, b) => b[1][0] - a[1][0]).slice(0, n)
      .map(([pt, [sc, name]]) => ({ pt, ptEn: md.ptEn.get(pt) || '', ptKr: md.ptKr.get(pt) || '',
        llt: name, score: Math.round(sc * 100) / 100 }));
  }

  // ------------------------------------------------------------------ PDF 읽기
  const NOISE = [/^\d\d\. \d+\. \d+\. 오[전후] .*약학정보원$/, /^https?:\/\//, /^[\d\s]+$/];
  const clean = (s) => String(s).replace(/&nbsp;/g, ' ').replace(/\(cid:\d+\)/g, '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/\s+/g, ' ').trim();

  /* pdf.js 텍스트 조각 -> 줄 [{y, text, chunks:[{x0,x1,text}]}] (페이지별) */
  async function pageLines(page) {
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    const items = [];
    for (const it of tc.items) {
      if (!it.str) continue;
      const x = it.transform[4], y = vp.height - it.transform[5];
      items.push({ x0: x, x1: x + it.width, y, s: it.str });
    }
    items.sort((a, b) => a.y - b.y || a.x0 - b.x0);
    const lines = [];
    for (const it of items) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - it.y) <= 3) last.items.push(it);
      else lines.push({ y: it.y, items: [it] });
    }
    for (const ln of lines) {
      ln.items.sort((a, b) => a.x0 - b.x0);
      // 공백 조각은 위치 계산에서 빼고 '띄어쓰기 있음' 표시로만 씀 (pdf.js가 칸 사이에 넓은 공백을 넣음)
      let text = '', cur = null, prevEnd = null, pendingSpace = false;
      const chunks = [];
      for (const it of ln.items) {
        if (!it.s.trim()) { pendingSpace = true; continue; }
        const gap = prevEnd === null ? 0 : it.x0 - prevEnd;
        const sp = prevEnd !== null && (pendingSpace || gap > 3) ? ' ' : '';
        text += sp + it.s;
        if (!cur || gap > 8) { cur = { x0: it.x0, x1: it.x1, text: it.s }; chunks.push(cur); }
        else { cur.text += sp + it.s; cur.x1 = it.x1; }
        prevEnd = it.x1; pendingSpace = false;
      }
      ln.text = clean(text);
      ln.chunks = chunks.map((c) => ({ x0: c.x0, x1: c.x1, text: clean(c.text) })).filter((c) => c.text);
    }
    const out = lines.filter((l) => l.text);
    out.items = items;
    return out;
  }

  /* 표 테두리(얇은 사각형) 읽기: pdf.js 연산 목록에서 사각형을 찾아 화면 좌표로 바꿈 */
  async function pageEdges(page, pdfjsLib) {
    const OPS = pdfjsLib.OPS, vp = page.getViewport({ scale: 1 });
    let ops;
    try { ops = await page.getOperatorList(); } catch (e) { return { H: [], V: [] }; }
    const mul = (m, n) => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3],
      m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
    let ctm = [1, 0, 0, 1, 0, 0]; const stack = [], H = [], V = [];
    const pt = (x, y) => [ctm[0] * x + ctm[2] * y + ctm[4], vp.height - (ctm[1] * x + ctm[3] * y + ctm[5])];
    const addRect = (x, y, w, h) => {
      const a = pt(x, y), b = pt(x + w, y + h);
      const x0 = Math.min(a[0], b[0]), x1 = Math.max(a[0], b[0]), y0 = Math.min(a[1], b[1]), y1 = Math.max(a[1], b[1]);
      const W = x1 - x0, Hh = y1 - y0;
      if (W <= 2.5 && Hh > 4) V.push({ x: (x0 + x1) / 2, y0, y1 });
      else if (Hh <= 2.5 && W > 4) H.push({ y: (y0 + y1) / 2, x0, x1 });
      else if (W > 4 && Hh > 4 && W < vp.width * 0.9) {        // 칸을 사각형 테두리로 그린 경우
        H.push({ y: y0, x0, x1 }, { y: y1, x0, x1 }); V.push({ x: x0, y0, y1 }, { x: x1, y0, y1 });
      }
    };
    for (let i = 0; i < ops.fnArray.length; i++) {
      const fn = ops.fnArray[i], args = ops.argsArray[i];
      if (fn === OPS.save) stack.push(ctm);
      else if (fn === OPS.restore) ctm = stack.pop() || ctm;
      else if (fn === OPS.transform) ctm = mul(ctm, args);
      else if (fn === OPS.paintFormXObjectBegin) { stack.push(ctm); if (args[0]) ctm = mul(ctm, args[0]); }
      else if (fn === OPS.paintFormXObjectEnd) ctm = stack.pop() || ctm;
      else if (fn === OPS.constructPath) {
        const sub = args[0], c = args[1]; let j = 0;
        for (const o of sub) {
          if (o === OPS.rectangle) { addRect(c[j], c[j + 1], c[j + 2], c[j + 3]); j += 4; }
          else if (o === OPS.moveTo || o === OPS.lineTo) j += 2;
          else if (o === OPS.curveTo) j += 6;
          else if (o === OPS.curveTo2 || o === OPS.curveTo3) j += 4;
        }
      }
    }
    return { H, V };
  }

  async function readLabelPdf(pdfjsLib, bytes) {
    const doc = await pdfjsLib.getDocument({ data: bytes, verbosity: 0 }).promise;
    const pages = [], edges = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      pages.push(await pageLines(page));
      edges.push(await pageEdges(page, pdfjsLib));
    }
    const grids = pages.map((pl, i) => findGrid(pl, edges[i]));
    const lines = [];
    pages.forEach((pl, i) => {
      const g = grids[i]; let marked = false;
      for (const l of pl) {
        if (g && l.y > g.top + 1 && l.y < g.bottom + 2) {          // 표 안의 줄은 '[표]' 한 줄로 바꿈
          if (!marked) { lines.push('[표]'); marked = true; }
          continue;
        }
        if (NOISE.some((re) => re.test(l.text))) continue;
        if (lines.length && lines[lines.length - 1] === l.text) continue;   // 웹페이지 중복 제목
        lines.push(l.text);
      }
    });
    const product = (lines.find((l) => /mg|밀리그램|정|캡슐|주|액/.test(l)) || lines[0] || '').slice(0, 60);
    const table = grids.some(Boolean) ? parseGridTables(pages, grids, lines) : parseTable(pages, lines);
    return { pages, lines, table, product };
  }

  // ------------------------------------------------------------------ 이상반응 표 (테두리 기준)
  const uniqSorted = (vals, tol) => {
    const out = [];
    for (const v of [...vals].sort((a, b) => a - b)) if (!out.length || v - out[out.length - 1] > tol) out.push(v);
    return out;
  };

  /* '기관계'로 시작하는 머리글이 있는 표의 칸 경계(xs, ys) 찾기 */
  function findGrid(pl, edges) {
    if (!edges.H.length || !edges.V.length) return null;
    const head = pl.find((l) => l.chunks.length && /^기관계/.test(l.chunks[0].text));
    if (!head) return null;
    const above = edges.H.filter((e) => e.y <= head.y - 4 && e.x0 <= head.chunks[0].x0 + 2);
    if (!above.length) return null;
    const top = Math.max(...above.map((e) => e.y));
    // 윗줄이 칸마다 따로 그려져 있으므로 같은 높이의 조각을 이어서 표 너비를 정함
    const seg = edges.H.filter((e) => Math.abs(e.y - top) < 1.5).sort((a, b) => a.x0 - b.x0);
    const start = seg.find((e) => e.x0 <= head.chunks[0].x0 + 2 && e.x1 >= head.chunks[0].x0 - 2) || seg[0];
    const topLine = [start];
    for (const e of seg) { const r = Math.max(...topLine.map((t) => t.x1)); if (e.x0 <= r + 3 && e.x1 > r) topLine.push(e); }
    const left = Math.min(...topLine.map((e) => e.x0)) - 2, right = Math.max(...topLine.map((e) => e.x1)) + 2;
    let bottom = top, grew = true;
    while (grew) {                                            // 세로선을 이어 따라가며 표 아래 끝 찾기
      grew = false;
      for (const v of edges.V) {
        if (v.x >= left && v.x <= right && v.y0 <= bottom + 2 && v.y1 > bottom + 0.5) { bottom = v.y1; grew = true; }
      }
    }
    if (bottom - top < 20) return null;
    const ys = uniqSorted(edges.H.filter((e) => e.y >= top - 1 && e.y <= bottom + 1 && e.x1 > left && e.x0 < right).map((e) => e.y), 1.5);
    const xs = uniqSorted(edges.V.filter((v) => v.x >= left && v.x <= right && v.y1 > top && v.y0 < bottom).map((v) => v.x), 1.5);
    if (ys.length < 3 || xs.length < 3) return null;
    return { top, bottom, left, right, xs, ys };
  }

  /* 칸 안 글자 조각 -> 줄 단위 글 (공백 조각은 띄어쓰기로) */
  function itemsToLines(items) {
    const sorted = [...items].sort((a, b) => a.y - b.y || a.x0 - b.x0);
    const rows = [];
    for (const it of sorted) {
      const last = rows[rows.length - 1];
      if (last && Math.abs(last.y - it.y) <= 3) last.items.push(it); else rows.push({ y: it.y, items: [it] });
    }
    return rows.map((r) => {
      r.items.sort((a, b) => a.x0 - b.x0);
      let text = '', prevEnd = null, pending = false;
      for (const it of r.items) {
        if (!it.s.trim()) { pending = true; continue; }
        text += (prevEnd !== null && (pending || it.x0 - prevEnd > 3) ? ' ' : '') + it.s;
        prevEnd = it.x1; pending = false;
      }
      return clean(text);
    }).filter(Boolean);
  }

  /* 칸 글 -> 용어 목록: 쉼표로 나누고, 쉼표 없이 줄이 바뀌면 한 용어로 이어 붙임 */
  function cellTerms(lines) {
    if (!lines.length) return [];
    const hasComma = lines.some((l) => l.includes(','));
    if (!hasComma) return lines.length === 1 ? [lines[0]] : lines;
    let joined = '';
    lines.forEach((l, i) => { joined += (i && !/,$/.test(joined) ? ' ' : (i ? ' ' : '')) + l; });
    return joined.split(',').map(clean).filter(Boolean);
  }

  // 발현 빈도 구분은 질환이 아니므로 허가사항 용어에서 항상 뺌
  const FREQ_WORD = /^(매우\s*)?(흔하게|흔함|흔하지\s*않게|흔하지\s*않음|드물게|드묾|드뭄)$|^알려지지\s*않음$|^빈도\s*불명$|^기관계/;
  const isFreqWord = (t) => FREQ_WORD.test(String(t).replace(/\*+$/, '').replace(/\s*\(.*\)$/, '').trim());

  const HEADER_RE = /기관계|SOC|발현\s?빈도|흔함|흔하게|드물게|드묾|알려지지|빈도/;

  function parseGridTables(pages, grids, lines) {
    const notes = {}, terms = [];
    let title = lines.find((l) => /^표 \d/.test(l) && l.includes('이상반응')) || '';
    if (!title) {                                            // '표 1.' 제목이 없으면 표 바로 위 소항목을 제목으로
      const k = lines.indexOf('[표]');
      for (let i = k - 1; i >= 0 && k > 0; i--) {
        if (RE_SUB.test(lines[i]) || RE_ITEM.test(lines[i])) { title = RE_SUB.test(lines[i]) ? lines[i] : ''; break; }
      }
    }
    // 각주: 이상반응 표('[표]') 바로 아래에 붙은 '*' 줄만
    const near = new Set();
    lines.forEach((l, k) => { if (l === '[표]') for (let j = k + 1; j < lines.length && j <= k + 8; j++) { if (/^\*/.test(lines[j])) near.add(j); else if (near.size && !near.has(j - 1)) break; } });
    for (let i = 0; i < lines.length; i++) {
      if (!near.has(i)) continue;
      const m = lines[i].match(/^(\*+)\s*:?\s*(.+)$/);
      if (!m) continue;
      let body = m[2];
      while (i + 1 < lines.length && lines[i].length >= 55 && !/[.)]$/.test(lines[i])
             && !/^\*/.test(lines[i + 1]) && !RE_SUB.test(lines[i + 1])) { body += lines[++i]; }
      notes[m[1]] = `${m[1]}: ${body}`;
    }
    pages.forEach((pl, pi) => {
      const g = grids[pi]; if (!g) return;
      const nr = g.ys.length - 1, nc = g.xs.length - 1;
      const cells = Array.from({ length: nr }, () => Array.from({ length: nc }, () => []));
      for (const it of pl.items) {
        const cx = (it.x0 + it.x1) / 2, cy = it.y - 3;
        if (cy < g.top || cy > g.bottom || cx < g.left || cx > g.right) continue;
        let r = g.ys.findIndex((y, k) => k < nr && cy >= y && cy < g.ys[k + 1]);
        let c = g.xs.findIndex((x, k) => k < nc && cx >= x && cx < g.xs[k + 1]);
        if (r < 0 || c < 0) continue;
        cells[r][c].push(it);
      }
      const text = cells.map((row) => row.map((items) => itemsToLines(items)));
      // 머리글 행: 빈도/기관계 말이 있는 행 (데이터 행이 시작되기 전까지)
      let h = 0;
      while (h < nr && text[h].some((ls) => HEADER_RE.test(ls.join(' '))) && h < 4) h++;
      const freq = Array.from({ length: nc }, (_, c) => {
        for (let r = h - 1; r >= 0; r--) { const t = text[r][c].join(' '); if (t) return t.split(' (')[0].replace(/\(.*$/, '').trim(); }
        return '';
      });
      for (let r = h; r < nr; r++) {
        const soc = clean(text[r][0].join(' '));
        for (let c = 1; c < nc; c++) {
          for (const t of cellTerms(text[r][c])) {
            if (isFreqWord(t)) continue;
            const mark = (t.match(/(\*+)$/) || [, ''])[1];
            terms.push({ term: t.replace(/\*+$/, '').trim(), shown: t, mark, soc, freq: freq[c] });
          }
        }
      }
    });
    return { title, terms, notes };
  }

  // ------------------------------------------------------------------ 이상반응 표
  const RE_ITEM = /^(\d{1,2})\.\s*(\S.*)$/;
  const RE_SUB = /^(\d{1,2})\)\s*(\S.*)$/;
  const RE_BULLET = /^(?:[•·\-]|\(\d{1,2}\))\s*/;   // •, -, (1) 목록

  function parseTable(pages, lines) {
    let title = '';
    const notes = {}, terms = [];
    for (const ln of lines) if (/^표 \d/.test(ln) && ln.includes('이상반응')) { title = ln; break; }
    // 각주: '*' 줄 (다음 줄로 이어지면 합침)
    for (let i = 0; i < lines.length; i++) {
      const m = lines[i].match(/^(\*+)\s*:?\s*(.+)$/);
      if (!m) continue;
      let body = m[2];
      while (i + 1 < lines.length && lines[i].length >= 55 && !/[.)]$/.test(lines[i])
             && !/^\*/.test(lines[i + 1]) && !RE_SUB.test(lines[i + 1])) { body += lines[++i]; }
      notes[m[1]] = `${m[1]}: ${body}`;
    }
    for (const pl of pages) {
      const hi = pl.findIndex((l) => l.chunks.length && l.chunks[0].text.startsWith('기관계'));
      if (hi < 0) continue;
      const header = pl[hi];
      let i = hi + 1;
      while (i < pl.length && pl[i].chunks.every((c) => c.text.startsWith('('))) i++;
      const data = [];
      for (; i < pl.length; i++) {
        const t = pl[i].text;
        if (/^\*/.test(t) || RE_SUB.test(t) || RE_ITEM.test(t) || NOISE.some((re) => re.test(t))) break;
        data.push(pl[i]);
      }
      if (!data.length) continue;
      // 칸 시작 위치 = 내용 글자의 왼쪽 x 묶음
      const xs = [...new Set(data.flatMap((l) => l.chunks.map((c) => Math.round(c.x0))))].sort((a, b) => a - b);
      const cols = [];
      for (const x of xs) if (!cols.length || x - cols[cols.length - 1] > 4) cols.push(x);
      const colOf = (x) => { let k = 0; cols.forEach((c, j) => { if (x >= c - 2) k = j; }); return k; };
      const freq = cols.map((c, j) => {
        const next = j + 1 < cols.length ? cols[j + 1] : Infinity;
        const h = header.chunks.find((hc) => { const mid = (hc.x0 + hc.x1) / 2; return mid >= c - 2 && mid < next; });
        return h ? h.text.split(' (')[0] : '';
      });
      // 행 나누기: 줄 간격이 평소보다 크면 새 행
      const gaps = []; for (let k = 1; k < data.length; k++) gaps.push(data[k].y - data[k - 1].y);
      const sorted = [...gaps].sort((a, b) => a - b);
      const lineH = sorted.length ? sorted[Math.floor(sorted.length / 4)] : 12;
      const rows = [];
      data.forEach((l, k) => {
        if (!rows.length || data[k].y - data[k - 1].y > lineH * 1.4) rows.push(cols.map(() => []));
        const row = rows[rows.length - 1];
        for (const c of l.chunks) row[colOf(c.x0)].push(c.text);
      });
      for (const row of rows) {
        const soc = clean(row[0].join(' '));
        for (let ci = 1; ci < row.length; ci++) {
          for (let t of row[ci].join('\n').split(/[,\n]/)) {
            t = clean(t);
            if (!t || isFreqWord(t)) continue;
            const mark = (t.match(/(\*+)$/) || [, ''])[1];
            terms.push({ term: t.replace(/\*+$/, '').trim(), shown: t, mark, soc, freq: freq[ci] });
          }
        }
      }
    }
    return { title, terms, notes };
  }

  // ------------------------------------------------------------------ 문단 나누기
  const SECTION_RANK = { '이상반응 표': 1, '이상반응 본문': 2, '용법용량': 3, '일반적 주의': 4, '경고': 5 };
  const DEFAULT_SCOPE = ['이상반응 표', '이상반응 본문', '용법용량', '일반적 주의', '경고'];
  const FULL_LINE = 55;

  const isHeading = (ln) => ln.length <= 25 && !/[.,)]$/.test(ln) && !RE_BULLET.test(ln)
    && !ln.startsWith('*') && !ln.includes('%');

  function sectionKind(itemTitle, block) {
    if (block === '용법용량') return '용법용량';
    const t = itemTitle.replace(/ /g, '');
    if (t.includes('이상반응')) return '이상반응 본문';
    if (t.includes('일반적주의')) return '일반적 주의';
    if (t.includes('경고')) return '경고';
    return null;
  }

  function buildUnits(lines) {
    const units = [];
    let block = null, item = '', sub = '', head = '', lead = '', cur = null, prevRaw = '', inTable = false;
    const flush = () => { if (cur && cur.text) units.push(cur); cur = null; };
    for (const ln of lines) {
      const cont = cur !== null && prevRaw.length >= FULL_LINE && !prevRaw.endsWith('.');
      prevRaw = ln;
      if (ln.startsWith('용법 · 용량') || ln === '용법·용량') { flush(); block = '용법용량'; item = sub = head = lead = ''; continue; }
      if (ln.startsWith('사용상의 주의사항')) { flush(); block = '주의사항'; item = sub = head = lead = ''; continue; }
      if (block === null) continue;
      if (/^표 \d/.test(ln) || ln === '[표]') { flush(); inTable = true; continue; }
      if (inTable) { if (RE_SUB.test(ln) || RE_ITEM.test(ln)) inTable = false; else continue; }
      if (!cont) {
        if (RE_ITEM.test(ln) && ln.length < 40) { flush(); item = ln; sub = head = lead = ''; continue; }
        if (RE_SUB.test(ln) && ln.length < 50) { flush(); sub = ln; head = lead = ''; continue; }
        if (isHeading(ln)) { flush(); head = ln; lead = ''; continue; }
      }
      if (cont && !RE_BULLET.test(ln)) { cur.text += ln; continue; }
      flush();
      const isBullet = RE_BULLET.test(ln);
      const text = ln.replace(RE_BULLET, '');
      const path = [item, sub, head].filter(Boolean);
      const first = !units.some((u) => u.path.join('\u0001') === path.join('\u0001'));
      cur = { kind: sectionKind(item, block), path, text, lead: isBullet ? lead : '', first };
      if (!isBullet) lead = (text.endsWith('.') && text.length < 60) ? text : '';
    }
    flush();
    return units;
  }

  // ------------------------------------------------------------------ 용어 찾기
  const KR_STOP = new Set(['지침', '자세', '신장', '주사', '검사', '입원', '치료', '투여', '전이', '발적', '이루', '의존', '손상', '분비물']);
  const START_OK = new Set([' ', '(', ',', '/', '·', '‘', '’', "'", '"', '[']);
  const PROSE_EXCLUDE_SOC = new Set(['10042613', '10041244']);   // 시술, 사회 환경
  const GENERIC_PT = new Set(['10067484', '10061623', '10060933', '10013710', '10022891']);
  const SOC_INVESTIGATIONS = '10022891';
  const DIRECTION = ['increase', 'decrease', 'abnormal', 'positive', 'prolong', 'shorten',
    'elevat', 'low', 'high', 'present', 'detected'];
  const isAscii = (s) => /^[\x00-\x7F]*$/.test(s);

  function spotTerms(text, md) {
    const chars = [], pos = [];
    for (let i = 0; i < text.length; i++) if (!/\s/.test(text[i])) { chars.push(text[i]); pos.push(i); }
    const s = chars.join('');
    const hits = [];
    let i = 0;
    while (i < s.length) {
      const p = pos[i];
      const wordStart = p === 0 || START_OK.has(text[p - 1]);
      let best = null;
      for (let L = Math.min(md.maxLltLen, s.length - i); L > 1; L--) {
        const key = s.slice(i, i + L).toLowerCase();
        if (!md.lltKr.has(key) || KR_STOP.has(key)) continue;
        if (!wordStart && (L < 4 || isAscii(key))) continue;
        if (isAscii(key)) {
          const end = pos[i + L - 1] + 1;
          if (L < 3 || (end < text.length && /[A-Za-z0-9]/.test(text[end]))) continue;
        }
        best = [L, md.lltKr.get(key)]; break;
      }
      if (best) {
        const [L, [pt, name]] = best;
        hits.push({ llt: name, ptCode: pt, found: text.slice(pos[i], pos[i + L - 1] + 1) });
        i += L;
      } else i++;
    }
    return hits;
  }

  function isReactionTerm(md, pt) {
    if (GENERIC_PT.has(pt)) return false;
    const soc = md.ptPrimarySoc.get(pt);
    if (PROSE_EXCLUDE_SOC.has(soc)) return false;
    if (soc === SOC_INVESTIGATIONS) {
      const name = (md.ptEn.get(pt) || '').toLowerCase();
      return DIRECTION.some((d) => name.includes(d));
    }
    return true;
  }

  const WEAK_PATTERNS = [/연구되지 않았/, /확립되지 않았/, /포함되지 않았/, /병력이 있는/, /진단된 경우/,
    /위험이 있/, /예방/, /환자에게는.*피해야/, /환자에 대한 안전성/,
    /있는 환자/, /가진 환자/, /손상된 환자/, /환자를 대상으로/, /환자에게(는)? 주의/, /주의하여 사용/,
    /환자에게(는)? (이 약을 )?투여하지/, /환자$/];
  const rfind = (t, sub, end) => (end - sub.length < 0 ? -1 : t.lastIndexOf(sub, end - sub.length));
  function sentenceOf(text, found) {
    const i = text.indexOf(found);
    if (i < 0) return text;
    const st = Math.max(rfind(text, '다.', i), rfind(text, '. ', i));
    const en = text.indexOf('다.', i);
    return text.slice(st >= 0 ? st + 2 : 0, en >= 0 ? en + 2 : text.length);
  }
  const isWeak = (sentence) => WEAK_PATTERNS.some((re) => re.test(sentence));
  const refProse = (u) => [...u.path, ...(u.lead ? [u.lead] : []), u.text]
    .map((p) => p.replace(/용법·용량/g, '용법용량')).join('\n');
  const splitNoteTerms = (note) => note.replace(/^\*+\s*:?\s*/, '').replace(/\.$/, '')
    .replace(/ 포함/g, '').split(',').map(clean).filter(Boolean);

  /* 허가사항 -> 용어 목록. overrides: {허가사항 용어: PT 영문명} (확정 매핑) */
  function extract(label, md, overrides, scope) {
    overrides = overrides || {};
    scope = new Set(scope || DEFAULT_SCOPE);
    const out = [], unmapped = [];
    const mapTerm = (t) => {
      if (Object.prototype.hasOwnProperty.call(overrides, t)) {
        const code = md.ptCode(overrides[t]);
        return code ? [code, `(확인 매핑) ${overrides[t]}`] : null;
      }
      return md.krTermToPt(t);
    };
    const { title, terms, notes } = label.table;
    const secItem = label.lines.find((l) => RE_ITEM.test(l) && l.includes('이상반응')) || '2. 약물이상반응';
    if (scope.has('이상반응 표')) {
      for (const t of terms) {
        const note = t.mark ? notes[t.mark] : null;
        const ref = [secItem, title, `- ${t.shown}`].filter(Boolean).join('\n') + (note ? `\n${note}` : '');
        const row = { term: t.term, section: '이상반응 표', rank: 1, weak: false,
          detail: `${t.soc} / ${t.freq}`, ref, how: '표' };
        const m = mapTerm(t.term);
        if (m) out.push({ ...row, ptCode: m[0], llt: m[1] }); else unmapped.push(row);
        if (note) {
          for (const nt of splitNoteTerms(note)) {
            const m2 = mapTerm(nt);
            const r2 = { ...row, term: nt, detail: `${row.detail} / 각주(${t.shown})`, how: '표 각주' };
            if (m2) out.push({ ...r2, ptCode: m2[0], llt: m2[1] }); else unmapped.push(r2);
          }
        }
      }
    }
    for (const u of buildUnits(label.lines)) {
      if (!u.kind || !scope.has(u.kind)) continue;
      let head = (u.first && u.path.length > 1) ? u.path[u.path.length - 1] : '';
      head = head.replace(/\([A-Za-z ,.\-]+\)/g, '').replace(/^\d+\)\s*/, '');
      for (const [src, txt] of [['소제목', head], ['본문', u.text]]) {
        if (!txt) continue;
        for (const h of spotTerms(txt, md)) {
          if (!isReactionTerm(md, h.ptCode)) continue;
          const weak = isWeak(src === '소제목' ? txt : sentenceOf(txt, h.found));
          out.push({ term: h.found, ptCode: h.ptCode, llt: h.llt, section: u.kind,
            rank: SECTION_RANK[u.kind], detail: u.path.join(' > '), ref: refProse(u), how: src, weak });
        }
      }
    }
    for (const o of out) { o.ptEn = md.ptEn.get(o.ptCode) || ''; o.ptKr = md.ptKr.get(o.ptCode) || ''; }
    return { out, unmapped };
  }

  // ------------------------------------------------------------------ 용어 매핑표
  /* termMap: {허가사항 용어: [[PT 영문명, '확정'|'제안'], ...]} */
  function parseTermMapCsv(text) {
    const map = {};
    const rows = text.replace(/^﻿/, '').split(/\r?\n/).filter((l) => l.trim());
    for (const line of rows.slice(1)) {
      const cells = csvSplit(line);
      if (cells.length < 3) continue;
      const [t, pt, st] = cells.map((c) => c.trim());
      (map[t] = map[t] || []).push([pt, st]);
    }
    return map;
  }
  function csvSplit(line) {
    const out = []; let cur = '', q = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (q) { if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c; }
      else if (c === '"') q = true; else if (c === ',') { out.push(cur); cur = ''; } else cur += c;
    }
    out.push(cur); return out;
  }
  function termMapToCsv(map) {
    const esc = (s) => (/[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
    const lines = ['허가사항 용어,PT,상태'];
    for (const [t, rows] of Object.entries(map)) for (const [pt, st] of rows) lines.push([t, pt, st].map(esc).join(','));
    return '﻿' + lines.join('\r\n') + '\r\n';
  }

  function labelTerms(label, md, termMap, scope) {
    termMap = termMap || {};
    const confirmed = {};
    for (const [t, rows] of Object.entries(termMap)) if (rows[0] && rows[0][1] === '확정') confirmed[t] = rows[0][0];
    const { out, unmapped } = extract(label, md, confirmed, scope);
    for (const u of unmapped) {
      let code, st;
      if (termMap[u.term]) { [code, st] = [md.ptCode(termMap[u.term][0][0]), termMap[u.term][0][1]]; }
      else { const c = md.suggest(u.term); u.candidates = c; code = c.length ? c[0].pt : null; st = '제안'; }
      if (code) out.push({ ...u, ptCode: code, llt: `(후보 ${st})`, ptEn: md.ptEn.get(code) || '',
        ptKr: md.ptKr.get(code) || '', candidate: st !== '확정' });
    }
    for (const [term, rows] of Object.entries(termMap)) {
      const base = out.find((o) => o.term === term && o.section === '이상반응 표');
      if (!base) continue;
      for (const [ptName, st] of rows.slice(1)) {
        const code = md.ptCode(ptName);
        if (code) out.push({ ...base, ptCode: code, llt: `(묶음 ${st})`, ptEn: md.ptEn.get(code) || '',
          ptKr: md.ptKr.get(code) || '', candidate: st !== '확정', how: '표 묶음' });
      }
    }
    return { terms: out, unmapped };
  }

  // ------------------------------------------------------------------ 판정
  const NO_SPECIAL = new Set(['', '해당없음', '해당 없음', 'NA', 'N/A', '없음', 'NONE']);

  function judgeRow(md, terms, ptName, special) {
    if (special !== null && special !== undefined && !NO_SPECIAL.has(String(special).trim().toUpperCase()))
      return ['확인필요', `특수상황: ${String(special).trim()}`];
    const code = md.ptCode(ptName);
    if (!code) return ['확인필요', `PT '${ptName}'를 MedDRA v${md.version}에서 찾지 못함`];
    const same = terms.filter((t) => t.ptCode === code).sort((a, b) => a.rank - b.rank);
    const strong = same.filter((t) => !t.weak && !t.candidate);
    if (strong.length) return ['반영', strong[0].ref];
    if (same.length) {
      const t = same[0];
      const why = t.candidate ? '후보 매핑 확인 필요' : '약한 근거: 환자 상태 문맥';
      return ['확인필요', `[${why}] ${t.term}\n${t.ref}`];
    }
    const mine = md.ptHlts.get(code) || new Set();
    let best = null;
    for (const t of terms) {
      for (const h of (md.ptHlts.get(t.ptCode) || [])) {
        if (!mine.has(h)) continue;
        const primary = md.ptPrimaryHlt.get(code) === h && md.ptPrimaryHlt.get(t.ptCode) === h;
        const key = [t.weak ? 1 : 0, primary ? 0 : 1, t.rank, md.hltSize.get(h)];
        if (!best || cmp(key, best.key) < 0)
          best = { key, t, h, primary };
      }
    }
    if (best) {
      const { t, h, primary } = best;
      const tag = `[같은 HLT: ${md.hltName.get(h)} · ${primary ? 'Primary' : 'Secondary'} · ${md.hltSize.get(h)}개 묶음]`;
      return ['확인필요', `${tag} 허가사항 용어 '${t.term}' (PT ${t.ptEn})\n${t.ref}`];
    }
    return ['미반영', null];
  }
  const cmp = (a, b) => { for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return a[i] < b[i] ? -1 : 1; return 0; };

  root.PMS = { isFreqWord, buildMedDRA, slimMedDRA, readLabelPdf, parseTable, buildUnits, spotTerms, extract, labelTerms,
    judgeRow, parseTermMapCsv, termMapToCsv, SECTION_RANK, DEFAULT_SCOPE, norm };
})(typeof window !== 'undefined' ? window : globalThis);
