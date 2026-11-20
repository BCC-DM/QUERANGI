/* PMS Expectedness 판정 엔진 (브라우저 + Node 공용)
 * Python 버전(meddra.py, label_parser.py, judge.py)의 규칙에서 시작.
 * 2026-11-20 개정: 허가사항 8개 검토 반영
 *   - 이상반응(이상사례) 항목 안의 테두리 표를 모두 읽음 (머리글 모양·쪽 넘김·합친 칸 대응)
 *   - 숫자 줄을 항목 제목으로 읽던 버그, 이상사례·일반적인 주의·일반의약품 항목 인식
 *   - 본문 오탐 규칙: 꾸밈말(~성·~제·~한·~상), 띄어쓰기 조각, 제외 단어 추가, LDL/HDL
 *   - 허가사항 표현 사전(LABEL_SYNONYMS), 사전에 없는 용어의 PT 후보(변형 일치·부분 일치·PT 비슷함)
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
  const DASH = /[\-\u2010-\u2015\u2212]/;
  const norm = (s) => String(s || '').replace(/[\s\-\u2010-\u2015\u2212]+/g, '').toLowerCase();
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
    addSynonyms(md, LABEL_SYNONYMS);
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

  /* 한글 PT 이름으로 연결 (2026-11-20 지후 요청)
     허가사항 표현은 LLT보다 PT 이름과 글자가 조금 다른 경우가 많음: 저체온 ↔ 저체온증, 졸음 ↔ 졸림, 횡문근융해증 ↔ 횡문근 융해
     1) 변형 일치: 끝의 '증'·'성' 붙이기/떼기, 수치·혈중 정리, 증가↔상승, 감소↔저하 등을 바꿔서 LLT·PT 이름과 똑같아지면 후보
     2) PT 비슷함: 한글 PT 이름과 글자 차이(편집 거리)·두 글자 조각이 비슷한 순서 */
  function termVariants(term) {
    const b = norm(term);
    const out = new Set();
    const add = (x) => { if (x && x !== b && x.length >= 2) out.add(x); };
    const base = [b, b.replace(/수치/g, ''), b.replace(/^혈중/, '혈액'), b.replace(/^혈중/, ''), b.replace(/^(요중|뇨중)/, '소변'), b.replace(/^(요중|뇨중)/, ''),
      b.replace(/상승$/, '증가'), b.replace(/증가$/, '상승'), b.replace(/저하$/, '감소'), b.replace(/감소$/, '저하'), b.replace(/이상$/, '비정상'), b.replace(/비정상$/, '이상')];
    for (const x of base) {
      add(x);
      if (x.endsWith('증')) add(x.slice(0, -1)); else add(x + '증');
      add(x.replace(/성(?=.)/g, ''));
      add(x.replace(/([가-힣]{2})(저혈압|고혈압|부종|통증|출혈|궤양)$/, '$1성$2'));
    }
    return [...out];
  }
  function lev(a, b) {
    if (Math.abs(a.length - b.length) > 2) return 99;
    const d = Array.from({ length: a.length + 1 }, (_, i) => [i]);
    for (let j = 1; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++)
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[a.length][b.length];
  }
  function suggestPT(md, term, n) {
    if (!md.ptKrIndex) md.ptKrIndex = [...md.ptKr.entries()].map(([c, k]) => [c, norm(k), grams(norm(k))]);
    const q = norm(term), gq = grams(q), res = [];
    for (const [code, k, gk] of md.ptKrIndex) {
      let inter = 0; for (const g of gq) if (gk.has(g)) inter++;
      const jac = inter / (gq.size + gk.size - inter);
      const ed = lev(q, k), led = ed === 99 ? 0 : 1 - ed / Math.max(q.length, k.length);
      const same1 = q[0] === k[0];
      // 짧은 단어는 글자 하나 차이로도 엉뚱한 말이 됨(졸음 ↔ 울음): 편집 거리는 첫 글자가 같거나 3글자 이상일 때만
      let sc = (same1 || q.length >= 3) ? Math.max(jac, same1 ? led : led * 0.8) : jac;
      if (same1) sc += 0.1;
      if (sc >= 0.34) res.push({ pt: code, ptEn: md.ptEn.get(code) || '', ptKr: md.ptKr.get(code) || '', llt: md.ptKr.get(code) || '', score: Math.round(sc * 100) / 100, how: 'PT 비슷함' });
    }
    return res.sort((a, b) => b.score - a.score).slice(0, n);
  }
  function variantHits(md, term) {
    const out = [];
    for (const v of termVariants(term)) {
      const h = md.lltKr.get(v);
      if (h && !out.some((x) => x.pt === h[0])) out.push({ pt: h[0], ptEn: md.ptEn.get(h[0]) || '', ptKr: md.ptKr.get(h[0]) || '', llt: h[1], score: '변형 일치', how: '변형 일치' });
    }
    return out;
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
    const pages = [], edges = [], vps = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      pages.push(await pageLines(page));
      edges.push(await pageEdges(page, pdfjsLib));
      vps.push(page.getViewport({ scale: 1 }));
    }
    // 새 방식 (2026-11-20): 테두리가 그려진 표를 모두 찾고, 이상반응(이상사례) 항목 안의 표를 이상반응 표로 읽음
    const bt = readBorderTables(pages, edges, vps);
    if (bt.table.terms.length) {
      const product = (bt.lines.find((l) => /mg|밀리그램|정|캡슐|주|액/.test(l)) || bt.lines[0] || '').slice(0, 60);
      return { pages, lines: bt.lines, table: bt.table, product, legacy: false };
    }
    // 예전 방식: '기관계' 머리글로 표 찾기
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
    return { pages, lines, table, product, legacy: true };
  }


  // ------------------------------------------------------------------ 이상반응 표 (테두리 표 전체, 2026-11-20)
  /* 1) 쪽마다 표 테두리 선을 이어진 묶음으로 나눠 표를 찾음 (쪽 테두리·머리말 상자는 뺌)
     2) 이상반응(이상사례) 항목 안에 있는 표만 이상반응 표로 씀. 다른 표의 글자는 본문으로 되돌림
     3) 머리글로 열 역할(SOC / 빈도 / 용어 / 숫자)을 정하고, 합쳐진 칸은 같은 값으로 채움
     4) 머리글 없이 다음 쪽 맨 위에서 이어지는 표는 앞 표의 열 역할을 이어받음 */
  const T_HEAD = /n\s*=|기관계|SOC|발현\s*부위|발현\s*빈도|발현\s*증상|빈도|이상\s*반응|이상\s*사례|증상|우선\s*용어|그룹\s*용어|분류|등급|위약|이\s*약|투여군|환자\s*비율|\d+\s*(명|건)|\d\s*%\s*(이상|미만|이하|초과)/i;
  const T_TERM = /우선\s*용어|증상|이상\s*반응|이상\s*사례|용어|부작용/;
  const T_SOC = /기관계|SOC|부위|분류/;
  const hasHangul = (x) => /[가-힣]/.test(x);
  const freqLike = (x) => { x = String(x).replace(/^\d\)\s*/, ''); const t = x.replace(/\s*\(.*$/, '').trim();
    return isFreqWord(t) || /\d\s*%\s*(이상|미만|이하|초과)|^[<>≥≤]\s*\d/.test(x) || /^[\d.\s%()~<>≥≤,\/명건\-]+$/.test(x) || (x.length < 30 && /흔하|흔함|흔히|드물|드묾|불명|알려지지|알 수 없|때때로/.test(x)); };

  function findTables(E, vp) {
    const W = vp.width, Hh = vp.height;
    const H = E.H.filter((h) => !(h.x0 < 45 && h.x1 > W - 45) && h.y > 30 && h.y < Hh - 30 && h.x1 - h.x0 < W * 0.95);
    const V = E.V.filter((v) => v.x > 45 && v.x < W - 45 && v.y1 > 30 && v.y0 < Hh - 30);
    const segs = H.map((h) => ({ t: 'H', ...h })).concat(V.map((v) => ({ t: 'V', ...v })));
    const par = segs.map((_, i) => i);
    const f = (i) => (par[i] === i ? i : (par[i] = f(par[i])));
    for (let i = 0; i < segs.length; i++) for (let j = i + 1; j < segs.length; j++) {
      const a = segs[i], b = segs[j]; let c;
      if (a.t !== b.t) { const h = a.t === 'H' ? a : b, v = a.t === 'H' ? b : a; c = v.x >= h.x0 - 2 && v.x <= h.x1 + 2 && h.y >= v.y0 - 2 && h.y <= v.y1 + 2; }
      else if (a.t === 'H') c = Math.abs(a.y - b.y) < 1.5 && a.x0 <= b.x1 + 2 && b.x0 <= a.x1 + 2;
      else c = Math.abs(a.x - b.x) < 1.5 && a.y0 <= b.y1 + 2 && b.y0 <= a.y1 + 2;
      if (c) par[f(i)] = f(j);
    }
    const groups = {};
    segs.forEach((sg, i) => { (groups[f(i)] = groups[f(i)] || []).push(sg); });
    const out = [];
    for (const ss of Object.values(groups)) {
      const h = ss.filter((x) => x.t === 'H'), v = ss.filter((x) => x.t === 'V');
      const ys = uniqSorted(h.map((x) => x.y), 1.5), xs = uniqSorted(v.map((x) => x.x), 1.5);
      if (ys.length < 2 || xs.length < 2) continue;
      const nr = ys.length - 1, nc = xs.length - 1;
      if (nc < 2 && nr < 3) continue;                        // 글 상자 하나짜리는 표가 아님 (본문으로 읽음)
      out.push({ H: h, V: v, ys, xs, nr, nc, x0: xs[0], x1: xs[nc], y0: ys[0], y1: ys[nr] });
    }
    return out.sort((a, b) => a.y0 - b.y0);
  }

  /* 칸 글자 읽기 + 합쳐진 칸 찾기 (칸 사이에 선이 없으면 같은 칸) */
  function gridCells(items, t) {
    const { ys, xs, nr, nc } = t;
    const sub = Array.from({ length: nr }, () => Array.from({ length: nc }, () => []));
    for (const it of items) {
      if (!it.s.trim()) { /* 띄어쓰기 조각도 위치대로 넣음 */ }
      const cx = (it.x0 + it.x1) / 2, cy = it.y - 3;
      if (cy < t.y0 || cy > t.y1 || cx < t.x0 || cx > t.x1) continue;
      const r = ys.findIndex((y, k) => k < nr && cy >= y && cy < ys[k + 1]);
      const c = xs.findIndex((x, k) => k < nc && cx >= x && cx < xs[k + 1]);
      if (r >= 0 && c >= 0) sub[r][c].push(it);
    }
    const id = (r, c) => r * nc + c, par = Array.from({ length: nr * nc }, (_, i) => i);
    const f = (i) => (par[i] === i ? i : (par[i] = f(par[i])));
    const hLine = (y, xa, xb) => t.H.some((h) => Math.abs(h.y - y) < 1.5 && h.x0 <= xa + 2 && h.x1 >= xb - 2);
    const vLine = (x, ya, yb) => t.V.some((v) => Math.abs(v.x - x) < 1.5 && v.y0 <= ya + 2 && v.y1 >= yb - 2);
    for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) {
      const mx = (xs[c] + xs[c + 1]) / 2, my = (ys[r] + ys[r + 1]) / 2;
      if (r + 1 < nr && !hLine(ys[r + 1], mx - 1, mx + 1)) par[f(id(r, c))] = f(id(r + 1, c));
      if (c + 1 < nc && !vLine(xs[c + 1], my - 1, my + 1)) par[f(id(r, c))] = f(id(r, c + 1));
    }
    const gItems = {}, gCells = {};
    for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) {
      const g = f(id(r, c));
      (gItems[g] = gItems[g] || []).push(...sub[r][c]);
      (gCells[g] = gCells[g] || []).push([r, c]);
    }
    const text = {};
    for (const g of Object.keys(gItems)) text[g] = itemsToLines(gItems[g]);
    return Array.from({ length: nr }, (_, r) => Array.from({ length: nc }, (_, c) => {
      const g = f(id(r, c)), cells = gCells[g];
      const cols = new Set(cells.map((x) => x[1])), firstRow = Math.min(...cells.map((x) => x[0]));
      return { g, lines: text[g], text: clean(text[g].join(' ')), spanAll: cols.size === nc && nc > 1, first: firstRow === r };
    }));
  }

  /* 칸 글 -> 용어: 쉼표로 나누되 괄호 안 쉼표는 그대로 (상기도 감염 (인두염, 부비동염 포함)) */
  function splitCell(lines) {
    if (!lines.length) return [];
    if (!lines.some((l) => l.includes(','))) return lines.length === 1 ? [lines[0]] : lines;
    const joined = lines.join(' ');
    const out = []; let cur = '', depth = 0;
    for (const ch of joined) {
      if (ch === '(' || ch === '[') depth++;
      if ((ch === ')' || ch === ']') && depth > 0) depth--;
      if (ch === ',' && depth === 0) { out.push(cur); cur = ''; } else cur += ch;
    }
    out.push(cur);
    return out.map(clean).filter(Boolean);
  }
  const socName = (x) => clean(String(x).replace(/^\d+\)\s*/, '').replace(/\s*\d\)$/, ''));

  /* 표 용어 다듬기: 각주 표시·윗첨자·(N명)·빈도 괄호 떼기 */
  function tidyTerm(raw) {
    let t = clean(raw).replace(/[＊∗※]/g, '*').replace(/^<\s*|\s*>$/g, '').replace(/^<|>$/g, '');
    let lead = t.match(/^\d\)\s*/); if (lead) t = t.slice(lead[0].length);
    t = t.replace(/\s*\(\s*\d\s*\)\)?\s*$/, '');
    t = t.replace(/\s*\([^()]*(포함|치명|결과)[^()]*\)\s*$/, '');
    t = t.replace(/\s*\((빈도\s*불명|빈도는?\s*알려지지\s*않음)\)\s*$/, '');
    t = t.replace(/\s*\(\d[\d,]*\s*(명|건)\)\s*$/, '').replace(/\s+\d+(\.\d+)?\s*%.*$/, '');
    let mark = '';
    const m1 = t.match(/^(?:[*†‡]+\s*)+/); if (m1) { mark = m1[0].replace(/\s/g, '').slice(0, 3); t = t.slice(m1[0].length); }
    const m2 = t.match(/\s*([*†‡]+)$/); if (m2) { mark = mark || m2[1]; t = t.slice(0, -m2[0].length); }
    const m3 = t.match(/([가-힣)])\s*([a-h]{1,2}|\d\))$/); if (m3) { mark = mark || m3[2]; t = t.slice(0, -(m3[0].length - 1)); }
    t = t.replace(/^<\s*|\s*>$/g, '').trim();
    return { term: t, mark };
  }

  function tableTitle(lines, j) {
    for (let k = j - 1; k >= Math.max(0, j - 3); k--) if (/^[<\[]?\s*표\s*\d/.test(lines[k]) || (/표/.test(lines[k]) && /이상반응|이상사례/.test(lines[k]) && lines[k].length < 90)) return lines[k];
    for (let k = j - 1; k >= Math.max(0, j - 25); k--) { if (RE_ITEM.test(lines[k]) && lines[k].length < 80) break; if (RE_SUB.test(lines[k]) && lines[k].length < 90) return lines[k]; }
    return '';
  }

  function tableNotes(lines, j) {
    const notes = {}, groups = {};
    for (let k = j + 1, n = 0; k < lines.length && n < 30; k++, n++) {
      const l = lines[k];
      const g = l.match(/^(?:[a-z]\s+)?그룹\s*용어\s+(.+?)[은는]\s+(.+?)의?\s*우선\s*용어/);
      if (g) { groups[g[1].trim()] = g[2].split(/,|\s및\s/).map(clean).filter(Boolean); continue; }
      const m = l.match(/^([*†‡]+)\s*:?\s*(.+)$/);
      if (m) { let body = m[2]; while (k + 1 < lines.length && lines[k].length >= FULL_LINE && !/[.)]$/.test(lines[k]) && !/^[*†‡]/.test(lines[k + 1])) body += lines[++k]; notes[m[1]] = `${m[1]}: ${body}`; continue; }
      if (/^(PT\s*=|[a-z]$|[a-z]\s)/.test(l)) continue;
      if (Object.keys(notes).length || Object.keys(groups).length) { if (!/^그룹|^\*|^[a-z]/.test(l)) break; }
      else if (n > 2) break;
    }
    return { notes, groups };
  }

  function readBorderTables(pages, edges, vps) {
    const lines = [], tables = [];
    pages.forEach((pl, i) => {
      const ts = findTables(edges[i], vps[i]);
      const kOf = ts.map(() => -1);
      for (const l of pl) {
        const j = ts.findIndex((t) => l.y > t.y0 + 1 && l.y < t.y1 + 2 && l.chunks.some((c) => c.x0 >= t.x0 - 3 && c.x1 <= t.x1 + 3));
        if (j >= 0) {
          if (kOf[j] < 0) { kOf[j] = tables.length; tables.push({ page: i, t: ts[j], raw: [] }); lines.push('\u0000T' + kOf[j]); }
          tables[kOf[j]].raw.push(l.text); continue;
        }
        if (NOISE.some((re) => re.test(l.text))) continue;
        if (lines.length && lines[lines.length - 1] === l.text) continue;
        lines.push(l.text);
      }
    });
    // 표가 어느 항목 안에 있는지
    let block = null, item = '';
    for (const ln of lines) {
      if (ln.startsWith('용법 · 용량') || ln === '용법·용량') { block = '용법용량'; item = ''; continue; }
      if (ln.startsWith('사용상의 주의사항')) { block = '주의사항'; item = ''; continue; }
      if (ln.startsWith('\u0000T')) { const T = tables[+ln.slice(2)]; T.item = item; T.ae = block === '주의사항' && sectionKind(item, block) === '이상반응 본문'; continue; }
      if (RE_ITEM.test(ln) && ln.length < 80) item = ln;
    }
    const outLines = [];
    for (const ln of lines) {
      if (!ln.startsWith('\u0000T')) { outLines.push(ln); continue; }
      const T = tables[+ln.slice(2)];
      if (T.ae) outLines.push('[표]'); else for (const r of T.raw) if (!NOISE.some((re) => re.test(r))) outLines.push(r);
    }
    // 이상반응 표 읽기
    const terms = [], allNotes = {}; let prev = null; const markerIdx = []; outLines.forEach((l, k) => { if (l === '[표]') markerIdx.push(k); });
    let mi = 0;
    tables.forEach((T) => {
      if (!T.ae) return;
      const j = markerIdx[mi++];
      const cells = gridCells(pages[T.page].items, T.t);
      const nc = T.t.nc;
      // 머리글 행
      let h = 0;
      while (h < Math.min(4, cells.length)) {
        const row = cells[h].filter((c) => c.first || c.spanAll);
        const vals = row.map((c) => c.text).filter(Boolean);
        if (!vals.length) { h++; continue; }
        if (cells[h][0].spanAll && !T_HEAD.test(cells[h][0].text)) break;
        if (vals.every((v) => T_HEAD.test(v) || freqLike(v) || !hasHangul(v)) && vals.some((v) => T_HEAD.test(v) || isFreqWord(v))) h++; else break;
      }
      const cont = !h && prev && T.page === prev.page + 1 && prev.nc === nc;
      let roles, colFreq, title, soc = '';
      if (cont) { roles = prev.roles; colFreq = prev.colFreq; title = prev.title; soc = prev.soc; }
      else {
        title = tableTitle(outLines, j);
        const head = Array.from({ length: nc }, (_, c) => clean(cells.slice(0, h).map((r) => r[c].text).filter((v, k, a) => v && a.indexOf(v) === k).join(' ')));
        const data = (c) => cells.slice(h).map((r) => r[c]).filter((x) => x.first && !x.spanAll && x.text).map((x) => x.text);
        roles = []; colFreq = [];
        for (let c = 0; c < nc; c++) {
          const hd = head[c], d = data(c);
          const noH = d.length ? d.filter((x) => !hasHangul(x)).length / d.length : 0;
          const fl = d.length ? d.filter(freqLike).length / d.length : 0;
          let role;
          if (d.length && noH >= 0.6) role = 'num';
          else if (hd && T_TERM.test(hd)) role = 'term';
          else if (hd && T_SOC.test(hd)) role = 'soc';
          else if (hd && /빈도/.test(hd) && fl >= 0.5) role = 'freq';
          else if (d.length && fl >= 0.7) role = 'freq';
          else if (!hd && c === 0 && nc > 1) role = 'soc';
          else role = 'term';
          roles.push(role);
          let fh = '';
          for (let r = h - 1; r >= 0 && role === 'term'; r--) { const v = cells[r][c].text; if (v && freqLike(v)) { fh = v.replace(/^\d\)\s*/, '').split(' (')[0].replace(/\(.*$/, '').trim(); break; } }
          colFreq.push(fh);
        }
        if (!roles.includes('term')) { const k = roles.lastIndexOf('soc'); if (k >= 0 && k < nc - 1) roles[k + 1] = 'term'; }
      }
      const { notes, groups } = tableNotes(outLines, j);
      for (let r = h; r < cells.length; r++) {
        const row = cells[r];
        if (row[0].spanAll) {
          const t = row[0].text;
          if (!t || !row[0].first) continue;
          const m = t.match(/^([*†‡]+)\s*:?\s*(.+)$/);
          if (m) notes[m[1]] = `${m[1]}: ${m[2]}`; else if (t.length <= 40) soc = socName(t);
          continue;
        }
        const socCol = roles.indexOf('soc');
        if (socCol >= 0 && row[socCol].text) soc = socName(row[socCol].text);
        if (roles.includes('num') && roles.every((x, c) => x !== 'num' || !row[c].text)) {
          const tc = roles.indexOf('term');
          if (tc >= 0 && row[tc].text && row[tc].text.length <= 40) { soc = socName(row[tc].text); continue; }
        }
        const fq = roles.map((x, c) => (x === 'freq' ? row[c].text : '')).filter(Boolean).join(' ').split(' (')[0];
        roles.forEach((role, c) => {
          if (role !== 'term' || !row[c].first) return;
          for (const raw0 of splitCell(row[c].lines)) {
            // 표 안 설명 문장은 용어가 아님 (예: "보고된 이상반응 중 연구자가 … 판단한 이상반응")
            if (/되었|보고됨|판단|관련성|연구자|환자군|위하여|위해/.test(raw0) || raw0.split(' ').length > 6) continue;
            const parts = /[가-힣]\s*\/\s*[가-힣]/.test(raw0) && !/\(/.test(raw0) ? raw0.split('/') : [raw0];
            for (const raw of parts) {
            const { term, mark } = tidyTerm(raw);
            if (!term || isFreqWord(term) || !(hasHangul(term) || /^[A-Za-z][A-Za-z\- ]{2,}$/.test(term))) continue;
            terms.push({ term, shown: clean(raw0), mark, soc, freq: fq || colFreq[c], title, item: T.item, notes, group: groups[term] || null });
            }
          }
        });
      }
      prev = { page: T.page, nc, roles, colFreq, title, soc };
      for (const k of Object.keys(notes)) if (!(k in allNotes)) allNotes[k] = notes[k];
    });
    return { lines: outLines, table: { title: '', terms, notes: allNotes } };
  }

  // ------------------------------------------------------------------ 이상반응 표 (테두리 기준, 예전 방식)
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
  const FREQ_WORD = /^(매우\s*)?(흔하게|흔함|흔히|흔하지\s*않게|흔하지\s*않음|드물게|드묾|드뭄)$|^(때때로|자주|가끔|이따금)$|^알려지지\s*않음$|^빈도\s*불명$|^기관계/;
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
            const t2 = t.replace(/\s*\(\d[\d,]*\s*(명|건)\)\s*$/, '');
            const mark = (t2.match(/(\*+)$/) || t2.match(/^(\*+)/) || [, ''])[1];
            terms.push({ term: t2.replace(/^\*+\s*/, '').replace(/\*+$/, '').trim(), shown: t, mark, soc, freq: freq[c] });
          }
        }
      }
    });
    return { title, terms, notes };
  }

  // ------------------------------------------------------------------ 이상반응 표
  const RE_ITEM = /^(\d{1,2})\.(?!\d)\s*(\S.*)$/;
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
            const t2 = t.replace(/\s*\(\d[\d,]*\s*(명|건)\)\s*$/, '');
            const mark = (t2.match(/(\*+)$/) || t2.match(/^(\*+)/) || [, ''])[1];
            terms.push({ term: t2.replace(/^\*+\s*/, '').replace(/\*+$/, '').trim(), shown: t, mark, soc, freq: freq[ci] });
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
    if (t.includes('이상반응') || t.includes('이상사례') || t.includes('즉각중지')) return '이상반응 본문';
    if (t.includes('일반적주의') || t.includes('일반적인주의')) return '일반적 주의';
    if (t.includes('경고')) return '경고';
    return null;
  }

  function buildUnits(lines, skipTables) {
    if (skipTables === undefined) skipTables = true;
    const units = [];
    let block = null, item = '', sub = '', head = '', lead = '', cur = null, prevRaw = '', inTable = false;
    const flush = () => { if (cur && cur.text) units.push(cur); cur = null; };
    for (const ln of lines) {
      const cont = cur !== null && prevRaw.length >= FULL_LINE && !prevRaw.endsWith('.');
      prevRaw = ln;
      if (ln.startsWith('용법 · 용량') || ln === '용법·용량') { flush(); block = '용법용량'; item = sub = head = lead = ''; continue; }
      if (ln.startsWith('사용상의 주의사항')) { flush(); block = '주의사항'; item = sub = head = lead = ''; continue; }
      if (block === null) continue;
      if (ln === '[표]' || /^[<\[]?\s*표\s*\d/.test(ln)) { flush(); if (skipTables) inTable = true; continue; }
      if (inTable) { if (RE_SUB.test(ln) || RE_ITEM.test(ln)) inTable = false; else continue; }
      if (!cont) {
        if (RE_ITEM.test(ln) && ln.length < 80) { flush(); item = ln; sub = head = lead = ''; continue; }
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
  const KR_STOP = new Set(['지침', '자세', '신장', '주사', '검사', '입원', '치료', '투여', '전이', '발적', '이루', '의존', '손상', '분비물',
    // 2026-11-20 허가사항 8개 검토에서 '질환 아님'으로 판단된 일반 단어·분류명
    '의심', '분석', '질병', '가스', '사고', '살모넬라', '과량투여', '투약오류', '국소반응', '신생물', '혈액학적장애']);
  const START_OK = new Set([' ', '(', ',', '/', '·', '‘', '’', "'", '"', '[']);
  const PROSE_EXCLUDE_SOC = new Set(['10042613', '10041244']);   // 시술, 사회 환경
  const GENERIC_PT = new Set(['10067484', '10061623', '10060933', '10013710', '10022891']);
  const SOC_INVESTIGATIONS = '10022891';
  const DIRECTION = ['increase', 'decrease', 'abnormal', 'positive', 'prolong', 'shorten',
    'elevat', 'low', 'high', 'present', 'detected'];
  const isAscii = (s) => /^[\x00-\x7F]*$/.test(s);

  // 바로 뒤에 붙으면 앞 단어를 꾸미는 말(이뇨제, 불안정성, 반점상, 무반응한)이라 그 자리의 용어로 보지 않음
  const MODIFIER_NEXT = /^(성|제|한|상(?!태))/;
  const HANGUL = /[가-힣]/;
  function spotTerms(text, md) {
    // 단어 중간의 한글 풀이 괄호 건너뛰기: 용혈성(적혈구 파괴성)빈혈 -> 용혈성빈혈
    const skip = new Set();
    const glossRe = /[가-힣]\(([가-힣\s]{1,12})\)(?=[가-힣])(?![을를이가은는의과와로으에도])/g;
    let gm;
    while ((gm = glossRe.exec(text))) { if (/포함|및|등/.test(gm[1])) continue; for (let k = gm.index + 1; k < gm.index + gm[0].length; k++) skip.add(k); }
    const chars = [], pos = [];
    for (let i = 0; i < text.length; i++) if (!/\s/.test(text[i]) && !DASH.test(text[i]) && !skip.has(i)) { chars.push(text[i]); pos.push(i); }
    const s = chars.join('');
    const hits = [];
    let i = 0;
    while (i < s.length) {
      const p = pos[i];
      const wordStart = p === 0 || START_OK.has(text[p - 1]);
      let best = null, blocked = false;
      for (let L = Math.min(md.maxLltLen, s.length - i); L > 1; L--) {
        const key = s.slice(i, i + L).toLowerCase();
        if (!md.lltKr.has(key) || KR_STOP.has(key)) continue;
        if (!wordStart && (L < 4 || isAscii(key))) continue;
        if (/[A-Za-z0-9]/.test(text[p - 1] || '') && /^[a-z0-9]/.test(key)) continue;
        const end = pos[i + L - 1] + 1;
        if (isAscii(key)) {
          if (L < 3 || (end < text.length && /[A-Za-z0-9]/.test(text[end]))) continue;
        }
        const found = text.slice(p, end);
        // 띄어쓰기를 넘어 잡혔는데 마지막 조각이 한 글자이고 바로 뒤에 글자가 붙음: 노출 증|가, 노인 환|자
        const lastSeg = found.split(/\s+/).pop();
        if (/\s/.test(found) && lastSeg.length === 1 && end < text.length && HANGUL.test(text[end])) continue;
        // 가장 긴 후보 뒤에 꾸밈 접미사가 붙으면 이 자리는 건너뜀 (더 짧은 후보로 내려가지 않음)
        if (MODIFIER_NEXT.test(text.slice(end, end + 2))) { blocked = true; break; }
        best = [L, md.lltKr.get(key), end]; break;
      }
      if (best && !blocked) {
        const [L, [pt, name], end] = best;
        hits.push({ llt: name, ptCode: pt, found: text.slice(pos[i], end), at: pos[i] });
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
      return DIRECTION.some((d) => (d === 'low' || d === 'high') ? new RegExp('\\b' + d + '$').test(name) : name.includes(d));
    }
    return true;
  }

  const WEAK_PATTERNS = [/연구되지 않았/, /확립되지 않았/, /포함되지 않았/, /병력이 있는/, /진단된 경우/,
    /위험이 있/, /예방/, /환자에게는.*피해야/, /환자에 대한 안전성/,
    /있는 환자/, /가진 환자/, /손상된 환자/, /환자를 대상으로/, /환자에게(는)? 주의/, /주의하여 사용/,
    /환자에게(는)? (이 약을 )?투여하지/, /환자$/];
  const rfind = (t, sub, end) => (end - sub.length < 0 ? -1 : t.lastIndexOf(sub, end - sub.length));
  function sentenceOf(text, found, at) {
    const i = (at !== undefined && at >= 0) ? at : text.indexOf(found);
    if (i < 0) return text;
    const st = Math.max(rfind(text, '다.', i), rfind(text, '. ', i));
    const en = text.indexOf('다.', i);
    return text.slice(st >= 0 ? st + 2 : 0, en >= 0 ? en + 2 : text.length);
  }
  const isWeak = (sentence) => WEAK_PATTERNS.some((re) => re.test(sentence));
  const refProse = (u) => [...u.path, ...(u.lead ? [u.lead] : []), u.text]
    .map((p) => p.replace(/용법·용량/g, '용법용량')).join('\n');
  /* 각주가 용어 나열(타브너스 "*: 알라닌 아미노전이효소 증가, …")일 때만 용어로 씀.
     설명 문장("** 보고된 이상반응 중 연구자가 … 판단한 이상반응")이면 빈 목록 */
  const NOTE_SENTENCE = /되었|됨|보고된|판단|관련성|연구자|경우|의미|나타낸|표시|발현율|환자|투여|조사에서/;
  const splitNoteTerms = (note) => {
    const parts = note.replace(/^[*†‡]+\s*:?\s*/, '').replace(/\.$/, '').replace(/ 포함/g, '').split(',').map(clean).filter(Boolean);
    if (parts.some((x) => NOTE_SENTENCE.test(x) || x.split(' ').length > 5)) return [];
    return parts;
  };

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
      const hit = md.krTermToPt(t);
      if (hit) return hit;
      const t2 = t.replace(/\s*\([^()]*\)\s*$/, '');
      return t2 !== t && t2 ? md.krTermToPt(t2) : null;
    };
    const { title, terms, notes } = label.table;
    const secItem = label.lines.find((l) => RE_ITEM.test(l) && l.includes('이상반응')) || '2. 약물이상반응';
    if (scope.has('이상반응 표')) {
      for (const t of terms) {
        const tn = t.notes || notes;
        const note = t.mark ? tn[t.mark] : null;
        const ref = [t.item || secItem, t.title !== undefined ? t.title : title, `- ${t.shown}`].filter(Boolean).join('\n') + (note ? `\n${note}` : '');
        const row = { term: t.term, section: '이상반응 표', rank: 1, weak: false,
          detail: `${t.soc} / ${t.freq}`, ref, how: '표' };
        const m = mapTerm(t.term);
        if (m) out.push({ ...row, ptCode: m[0], llt: m[1] }); else unmapped.push(row);
        const noteTerms = (note ? splitNoteTerms(note) : []).concat(t.group || []);
        if (noteTerms.length) {
          for (const nt of noteTerms) {
            const m2 = mapTerm(nt);
            const r2 = { ...row, term: nt, detail: `${row.detail} / 각주(${t.shown})`, how: '표 각주' };
            if (m2) out.push({ ...r2, ptCode: m2[0], llt: m2[1] }); else unmapped.push(r2);
          }
        }
      }
    }
    for (const u of buildUnits(label.lines, label.legacy !== false && label.table.terms.length > 0)) {
      if (!u.kind || !scope.has(u.kind)) continue;
      let head = (u.first && u.path.length > 1) ? u.path[u.path.length - 1] : '';
      head = head.replace(/\([A-Za-z ,.\-]+\)/g, '').replace(/^\d+\)\s*/, '');
      for (const [src, txt] of [['소제목', head], ['본문', u.text]]) {
        if (!txt) continue;
        for (const h of spotTerms(txt, md)) {
          if (!isReactionTerm(md, h.ptCode)) continue;
          const sentence = src === '소제목' ? txt : sentenceOf(txt, h.found, h.at);
          const weak = isWeak(sentence);
          out.push({ term: h.found, ptCode: h.ptCode, llt: h.llt, section: u.kind,
            rank: SECTION_RANK[u.kind], detail: u.path.join(' > '), ref: refProse(u), how: src, weak, sentence });
        }
      }
    }
    for (const o of out) { o.ptEn = md.ptEn.get(o.ptCode) || ''; o.ptKr = md.ptKr.get(o.ptCode) || ''; }
    return { out, unmapped };
  }

  // ------------------------------------------------------------------ 허가사항 표현 사전
  /* 허가사항에 쓰인 표현이 MedDRA 한글 LLT와 글자가 다를 때, 맞는 LLT로 이어 주는 목록.
     2026-11-20 허가사항 8개 검토에서 지후가 '일부만 잡힘'·'PT 틀림'으로 판단한 것 */
  // [허가사항 표현, MedDRA 한글 LLT, 출처]
  const LABEL_SYNONYMS = [
    ['불안정성 협심증', '불안정 협심증', '지후'],
    ['비소세포성 폐암', '비-소세포 폐암', '지후'], ['비소세포폐암', '비-소세포 폐암', '지후'],
    ['전신홍반루푸스', '전신 홍반성 루푸스', '지후'], ['피부홍반루푸스', '피부 홍반성 루푸스', '지후'],
    ['근위축성측삭경화증', '근위축 측삭 경화증', '지후'],
    ['미란성 역류식도염', '미란성 위 식도 역류 질환', '지후'],
    ['혈관긴장상태', '혈관 긴장 이상증', '지후'],
    ['나트륨 수치 감소', '혈액 나트륨 감소', '지후'],
    ['과질소혈증', '고질소 혈증', '지후'],
    ['만성 폐색성 폐 질환', '만성 폐쇄성 폐 질환', '지후'],
    ['흡인성탈모', '견인성 탈모', '지후'],
    ['긴장항진', '근육 긴장 항진증', '지후'],
    ['간기능검사이상', '간 기능 시험 이상', '지후'],
    ['근육골격 경직', '경직', '지후'],
    ['신동맥협착증', '신장 동맥 협착', '지후'], ['신동맥협착', '신장 동맥 협착', '지후'],
    // 아래는 Claude 제안 (지후 확인 전): '~성' 규칙 때문에 빠지는 맞는 용어를 되살림
    ['면역매개성 괴사성 근병증', '면역-매개 괴사성 근병증', 'Claude'],
    ['불안정형 협심증', '불안정 협심증', 'Claude'],
    ['불안정성 당뇨병', '불안정형 당뇨병', 'Claude'],
    ['위식도 역류성 질환', '위 식도 역류 질환', 'Claude'],
    ['류마티스성 다발성 근육통', '류마티스성 다발 근육통', 'Claude'],
    ['당뇨병성 신경병증', '당뇨 신경 병증', 'Claude'],
    ['마비성 장폐색', '무력성 장폐색증', 'Claude'],
    ['허혈성 심질환', '허혈성 심장병', 'Claude'],
    ['중독성 표피괴사용해', '독성 표피 괴사 용해', 'Claude'], ['중독성표피괴사용해', '독성 표피 괴사 용해', 'Claude'],
  ];
  function addSynonyms(md, list) {
    for (const [expr, llt] of list) {
      const h = md.lltKr.get(norm(llt));
      if (!h) continue;                          // 사전 버전에 그 LLT가 없으면 건너뜀
      const key = norm(expr);
      md.lltKr.set(key, [h[0], h[1], 'Y']);
      if (key.length > md.maxLltLen) md.maxLltLen = key.length;
    }
    return md;
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
      // 후보는 매핑표에 이미 있어도 항상 계산 (매핑 화면의 고르기 목록에 씀)
      const hits = spotTerms(u.term, md);
      const total = u.term.replace(/[\s()\/<>\-]/g, '').length || 1;
      const covered = hits.reduce((n, h) => n + h.found.replace(/\s/g, '').length, 0);
      const part = covered / total >= 0.6 ? hits.map((h) => ({ pt: h.ptCode, ptEn: md.ptEn.get(h.ptCode) || '', ptKr: md.ptKr.get(h.ptCode) || '', llt: h.llt, score: '부분 일치', how: '부분 일치' })) : [];
      // 후보 순서: 변형 일치 > 부분 일치 > PT 비슷함 > LLT 비슷함 (같은 PT는 한 번만)
      const sims = suggestPT(md, u.term, 4).concat(md.suggest(u.term, 4).map((x) => ({ ...x, how: 'LLT 비슷함', score: Math.round((x.score + (norm(x.llt)[0] === norm(u.term)[0] ? 0.1 : 0)) * 100) / 100 })))
        .sort((a, b) => b.score - a.score);
      const pool = variantHits(md, u.term).concat(part, sims);
      const c = pool.filter((x, i) => pool.findIndex((y) => y.pt === x.pt) === i).slice(0, Math.max(3, part.length));
      u.candidates = c.map((x) => ({ ...x, ptKr: x.ptKr || md.ptKr.get(x.pt) || '' }));
      let code, st;
      if (termMap[u.term]) { [code, st] = [md.ptCode(termMap[u.term][0][0]), termMap[u.term][0][1]]; }
      else {
        code = c.length ? c[0].pt : null; st = '제안';
        // 두 용어가 쉼표 없이 붙은 칸(눈의 감각이상 눈꺼풀 부종): 부분 일치가 여럿이면 각각 후보로
        for (const extra of part.slice(1)) out.push({ ...u, ptCode: extra.pt, llt: '(후보 제안)', ptEn: extra.ptEn, ptKr: extra.ptKr, candidate: true });
      }
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

  root.PMS = { termVariants, suggestPT, variantHits, pageLines, pageEdges, findTables, gridCells, readBorderTables, LABEL_SYNONYMS, addSynonyms, isFreqWord, buildMedDRA, slimMedDRA, readLabelPdf, parseTable, buildUnits, spotTerms, extract, labelTerms,
    judgeRow, parseTermMapCsv, termMapToCsv, SECTION_RANK, DEFAULT_SCOPE, norm };
})(typeof window !== 'undefined' ? window : globalThis);
