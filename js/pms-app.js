/* 쿼랑이 · PMS Expectedness 점검 화면 (pms-check.html)
 * 판정 엔진: js/pms-core.js · 엑셀 저장: js/xlsx-patch.js
 */
(function () {
  'use strict';
  var PMS = window.PMS;
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var fmt = function (n) { return Number(n).toLocaleString('ko-KR'); };
  var tick = function () { return new Promise(function (r) { setTimeout(r, 30); }); };
  var busy = function (msg) { $('busyText').textContent = msg || '처리 중…'; $('busy').classList.remove('hidden'); };
  var idle = function () { $('busy').classList.add('hidden'); };

  // ------------------------------------------------------------ 브라우저 저장소 (IndexedDB)
  var DB = {
    db: null,
    open: function () {
      var self = this;
      if (self.db) return Promise.resolve(self.db);
      return new Promise(function (res, rej) {
        var r = indexedDB.open('querangi-pms', 1);
        r.onupgradeneeded = function () { r.result.createObjectStore('meddra'); r.result.createObjectStore('termmaps'); };
        r.onsuccess = function () { self.db = r.result; res(self.db); };
        r.onerror = function () { rej(r.error); };
      });
    },
    req: function (store, mode, fn) {
      return this.open().then(function (db) {
        return new Promise(function (res, rej) {
          var t = db.transaction(store, mode), q = fn(t.objectStore(store));
          t.oncomplete = function () { res(q.result); };
          t.onerror = function () { rej(t.error); };
        });
      });
    },
    get: function (s, k) { return this.req(s, 'readonly', function (o) { return o.get(k); }); },
    put: function (s, k, v) { return this.req(s, 'readwrite', function (o) { return o.put(v, k); }); },
    keys: function (s) { return this.req(s, 'readonly', function (o) { return o.getAllKeys(); }); },
  };

  // ------------------------------------------------------------ 상태
  var S = {
    versions: [], catalog: [], mdMode: 'auto', adminFiles: [], mdCache: {}, sheet: null, label: null, termMap: {},
    scope: PMS.DEFAULT_SCOPE.slice(), view: null, results: [], ran: false, filter: 'all',
  };

  function dropZone(zone, input, onFiles) {
    zone.onclick = function () { input.click(); };
    input.onchange = function () { var f = Array.prototype.slice.call(input.files); input.value = ''; if (f.length) onFiles(f); };
    ['dragenter', 'dragover'].forEach(function (e) { zone.addEventListener(e, function (ev) { ev.preventDefault(); zone.classList.add('on'); }); });
    ['dragleave', 'drop'].forEach(function (e) { zone.addEventListener(e, function (ev) { ev.preventDefault(); zone.classList.remove('on'); }); });
    zone.addEventListener('drop', function (ev) { var f = Array.prototype.slice.call(ev.dataTransfer.files); if (f.length) onFiles(f); });
  }

  // ------------------------------------------------------------ 1. MedDRA (쿼랑이 meddra/ 폴더)
  var MODE_KEY = 'pms.meddraMode';
  var vcmp = function (a, b) { var x = a.split('.').map(Number), y = b.split('.').map(Number); return x[0] - y[0] || x[1] - y[1]; };

  async function loadCatalog() {
    var server = [], msg = '';
    try {
      var res = await fetch('meddra/versions.json', { cache: 'no-cache' });
      if (!res.ok) throw new Error(res.status);
      server = ((await res.json()).versions || []).map(function (v) { return { version: String(v.version), file: v.file, builtAt: v.builtAt, bytes: v.bytes, local: false }; });
    } catch (e) { msg = '쿼랑이의 MedDRA 목록(meddra/versions.json)을 불러오지 못했어요.'; }
    // 이 브라우저에만 추가한 버전(관리자 기능으로 만든 것)
    var keys = [];
    try { keys = ((await DB.keys('meddra')) || []).map(String); } catch (e) {}
    for (var i = 0; i < keys.length; i++) {
      if (server.some(function (x) { return x.version === keys[i]; })) continue;
      var rec = await DB.get('meddra', keys[i]);
      if (rec && rec.local) server.push({ version: keys[i], builtAt: rec.builtAt, local: true });
    }
    S.catalog = server.sort(function (a, b) { return vcmp(a.version, b.version); });
    S.versions = S.catalog.map(function (c) { return c.version; });
    var mode = localStorage.getItem(MODE_KEY) || 'auto';
    S.mdMode = mode === 'auto' || S.versions.indexOf(mode) >= 0 ? mode : 'auto';
    $('mdSel').innerHTML = '<option value="auto">코딩시트 DICVER 따르기 (권장)</option>' + S.catalog.map(function (c) {
      return '<option value="' + esc(c.version) + '">MedDRA v' + esc(c.version) + '로 모두 비교' + (c.local ? ' (이 PC에만 있음)' : '') + '</option>';
    }).join('');
    $('mdSel').value = S.mdMode;
    $('mdMsg').innerHTML = msg ? '<span style="color:var(--err)">' + esc(msg) + '</span>' : '';
    await renderMdInfo();
    updateRun();
  }

  async function renderMdInfo() {
    if (!S.catalog.length) { $('mdInfo').classList.add('hidden'); return; }
    var parts = [];
    for (var i = 0; i < S.catalog.length; i++) {
      var c = S.catalog[i], rec = null;
      try { rec = await DB.get('meddra', c.version); } catch (e) {}
      var cached = rec && (c.local || rec.builtAt === c.builtAt);
      parts.push('v' + esc(c.version) + ' ' + (cached ? '(이 PC에 저장됨)' : '(처음 쓸 때 내려받아요)'));
    }
    $('mdInfo').innerHTML = '<span><b>쿼랑이 MedDRA</b></span><span class="muted">' + parts.join(' · ') + ' · 영문 + 한글</span>';
    $('mdInfo').classList.remove('hidden');
  }

  async function getMd(v) {
    if (S.mdCache[v]) return S.mdCache[v];
    var c = S.catalog.filter(function (x) { return x.version === v; })[0];
    if (!c) throw new Error('MedDRA v' + v + ' 사전이 쿼랑이에 없어요.');
    var rec = null;
    try { rec = await DB.get('meddra', v); } catch (e) {}
    if (!rec || (!c.local && rec.builtAt !== c.builtAt)) {
      busy('MedDRA v' + v + ' 내려받는 중… (처음 한 번)'); await tick();
      var res = await fetch('meddra/' + c.file, { cache: 'no-cache' });
      if (!res.ok) throw new Error('MedDRA v' + v + ' 사전 파일을 내려받지 못했어요 (' + res.status + ').');
      rec = await res.json();
      try { await DB.put('meddra', v, rec); } catch (e) {}
      renderMdInfo();
    }
    busy('MedDRA v' + v + ' 준비 중…'); await tick();
    S.mdCache[v] = PMS.buildMedDRA(v, rec);
    return S.mdCache[v];
  }

  /* 관리자: MedDRA 폴더 -> 쿼랑이 사전 파일 만들기 */
  async function buildFromFolder(files) {
    var byDir = {};
    files.forEach(function (f) {
      var rel = f.webkitRelativePath || f.name, dir = rel.slice(0, rel.lastIndexOf('/') + 1);
      (byDir[dir] = byDir[dir] || {})[f.name.toLowerCase()] = f;
    });
    var found = {};
    for (var dir in byDir) {
      var fs = byDir[dir];
      if (!fs['meddra_release.asc']) continue;
      var rel = (await fs['meddra_release.asc'].text()).split('$');
      var v = (rel[0] || '').trim(), lang = (rel[1] || '').trim().toLowerCase();
      if (!/^\d+\.\d+$/.test(v)) continue;
      found[v] = found[v] || {};
      if (lang.indexOf('english') === 0 && fs['pt.asc'] && fs['mdhier.asc']) found[v].en = fs;
      if (lang.indexOf('korean') === 0 && fs['pt.asc'] && fs['llt.asc']) found[v].kr = fs;
    }
    var made = [], partial = [];
    for (var ver in found) {
      var p = found[ver];
      if (!p.en || !p.kr) { partial.push('v' + ver + '(' + (p.en ? '한글' : '영문') + ' 없음)'); continue; }
      busy('MedDRA v' + ver + ' 사전 파일 만드는 중…'); await tick();
      var bundle = PMS.slimMedDRA(ver, { enPt: await p.en['pt.asc'].text(), enMdhier: await p.en['mdhier.asc'].text(),
        krPt: await p.kr['pt.asc'].text(), krLlt: await p.kr['llt.asc'].text() });
      var text = JSON.stringify(bundle);
      var onServer = S.catalog.some(function (c) { return c.version === ver && !c.local; });
      try { await DB.put('meddra', ver, Object.assign({}, bundle, { local: !onServer })); } catch (e) {}
      delete S.mdCache[ver];
      made.push({ version: ver, file: 'meddra_' + ver + '.json', builtAt: bundle.builtAt, bytes: new Blob([text]).size, text: text });
    }
    idle();
    if (!made.length) {
      $('mdAdminMsg').innerHTML = '<span style="color:var(--err)">' + (partial.length ? esc(partial.join(', ')) + ' — 영문과 한글이 모두 있어야 해요.' :
        'MedDRA ASCII 파일(meddra_release.asc가 있는 폴더)을 찾지 못했어요.') + '</span>';
      return;
    }
    // 업로드할 versions.json (쿼랑이 목록 + 새로 만든 버전)
    var list = S.catalog.filter(function (c) { return !c.local && !made.some(function (m) { return m.version === c.version; }); })
      .map(function (c) { return { version: c.version, file: c.file, builtAt: c.builtAt, bytes: c.bytes }; })
      .concat(made.map(function (m) { return { version: m.version, file: m.file, builtAt: m.builtAt, bytes: m.bytes }; }))
      .sort(function (a, b) { return vcmp(a.version, b.version); });
    S.adminFiles = made.map(function (m) { return { name: m.file, text: m.text, type: 'application/json' }; })
      .concat([{ name: 'versions.json', text: JSON.stringify({ versions: list }, null, 2), type: 'application/json' }]);
    $('mdAdminMsg').innerHTML = '<div class="dictInfo" style="flex-direction:column;align-items:flex-start;gap:8px"><b>✓ ' +
      made.map(function (m) { return 'v' + esc(m.version); }).join(', ') + ' 사전 파일을 만들었어요. 이 브라우저에서는 바로 쓸 수 있어요.</b>' +
      '<span class="muted">아래 파일을 내려받아 GitHub 저장소의 <b>meddra</b> 폴더에 올리면(같은 이름은 덮어쓰기) 모든 팀원이 쓸 수 있어요.</span><div class="row" style="margin:0">' +
      S.adminFiles.map(function (f, i) { return '<button class="ghost" data-admin="' + i + '">' + esc(f.name) + ' 내려받기</button>'; }).join('') + '</div>' +
      (partial.length ? '<span style="color:var(--err)">' + esc(partial.join(', ')) + '는 영문·한글이 모두 있어야 해요.</span>' : '') + '</div>';
    await loadCatalog(); prepare();
  }

  // ------------------------------------------------------------ 2. 파일
  function normVer(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v.toFixed(1);
    var m = String(v).match(/(\d+)\.(\d+)/); if (m) return (+m[1]) + '.' + m[2];
    var n = String(v).match(/\d+/); return n ? n[0] + '.0' : null;
  }

  async function loadSheet(files) {
    var file = files[0]; busy('코딩시트 읽는 중…'); await tick();
    try {
      var bytes = new Uint8Array(await file.arrayBuffer());
      var wb = XLSX.read(bytes, { type: 'array' });
      var names = wb.SheetNames.slice().sort(function (a, b) { return (b === 'Adverse Event') - (a === 'Adverse Event'); });
      var pick = null;
      for (var i = 0; i < names.length && !pick; i++) {
        var ws = wb.Sheets[names[i]]; if (!ws || !ws['!ref']) continue;
        var rng = XLSX.utils.decode_range(ws['!ref']);
        var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, range: rng, blankrows: true });
        var hi = rows.slice(0, 8).findIndex(function (r) { return r && r.indexOf('PT') >= 0 && r.indexOf('SHEET_EXPT') >= 0; });
        if (hi >= 0) pick = { name: names[i], rng: rng, rows: rows, hi: hi };
      }
      if (!pick) throw new Error('PT와 SHEET_EXPT 변수명이 있는 시트를 찾지 못했어요.');
      var h = pick.rows[pick.hi], col = function (k) { return h.indexOf(k); };
      if (col('REF') < 0) throw new Error('REF 컬럼을 찾지 못했어요.');
      var special = ['AESPECIAL', 'AESS'].filter(function (k) { return col(k) >= 0; })[0] || null;
      var keyCol = col('UNIQUE_KEY'), list = [];
      pick.rows.forEach(function (r, i) {
        if (i <= pick.hi || !r) return;
        var key = keyCol >= 0 ? r[keyCol] : r[col('PT')];
        if (key == null || key === '') return;
        list.push({ excelRow: pick.rng.s.r + i + 1, key: key, subj: col('SUBJID') >= 0 ? r[col('SUBJID')] : '',
          pt: r[col('PT')], ptkor: col('PTKOR') >= 0 ? r[col('PTKOR')] : '', aeterm: col('AETERM') >= 0 ? r[col('AETERM')] : '',
          special: special ? r[col(special)] : null, ver: normVer(col('DICVER') >= 0 ? r[col('DICVER')] : null),
          existing: { expt: r[col('SHEET_EXPT')], ref: r[col('REF')] } });
      });
      var letter = function (c) { return XLSX.utils.encode_col(c + pick.rng.s.c); };
      S.sheet = { fileName: file.name, bytes: bytes, sheetName: pick.name, rows: list, special: special,
        exptLetter: letter(col('SHEET_EXPT')), refLetter: letter(col('REF')), error: null };
    } catch (e) { S.sheet = { fileName: file.name, error: e.message || String(e) }; }
    idle(); S.ran = false; renderFiles(); prepare();
  }

  async function loadPdf(files) {
    var file = files[0]; busy('허가사항 PDF 읽는 중…'); await tick();
    try {
      var label = await PMS.readLabelPdf(window.pdfjsLib, new Uint8Array(await file.arrayBuffer()));
      if (!label.lines.some(function (l) { return l.indexOf('사용상의 주의사항') === 0; }))
        throw new Error('"사용상의 주의사항"을 찾지 못했어요. 약학정보원 허가사항 PDF가 맞는지 확인해 주세요.');
      label.fileName = file.name; S.label = label; S.termMap = {};
      try { var saved = await DB.get('termmaps', label.product); if (saved && saved.map) S.termMap = saved.map; } catch (e) {}
      $('mapSaved').textContent = Object.keys(S.termMap).length ? '"' + label.product + '" 매핑표를 이 브라우저에서 불러왔어요.' : '';
    } catch (e) { S.label = { fileName: file.name, error: e.message || String(e) }; }
    idle(); S.ran = false; renderFiles(); prepare();
  }

  function renderFiles() {
    var rows = [];
    if (S.sheet) {
      var st = S.sheet.error ? '<span class="tag t-제외">읽기 불가</span> <span class="muted">' + esc(S.sheet.error) + '</span>'
        : '<span class="tag t-정상">읽음</span> <span class="muted">' + esc(S.sheet.sheetName) + ' 시트 · 이상사례 ' + fmt(S.sheet.rows.length) + '건 · MedDRA ' +
          esc(uniq(S.sheet.rows.map(function (r) { return r.ver ? 'v' + r.ver : 'DICVER 없음'; })).join(', ')) + ' · 특수상황 컬럼 ' + esc(S.sheet.special || '없음') + '</span>';
      rows.push('<tr><td>코딩시트</td><td><b>' + esc(S.sheet.fileName) + '</b></td><td>' + st + '</td></tr>');
    }
    if (S.label) {
      var st2 = S.label.error ? '<span class="tag t-제외">읽기 불가</span> <span class="muted">' + esc(S.label.error) + '</span>'
        : '<span class="tag t-정상">읽음</span> <span class="muted">' + esc(S.label.product) + ' · 이상반응 표 용어 ' + S.label.table.terms.length + '개 · 각주 ' + Object.keys(S.label.table.notes).length + '개</span>';
      rows.push('<tr><td>허가사항</td><td><b>' + esc(S.label.fileName) + '</b></td><td>' + st2 + '</td></tr>');
    }
    $('fileRows').innerHTML = rows.join('');
    $('fileBox').classList.toggle('hidden', !rows.length);
    $('dropSheet').classList.toggle('done', !!(S.sheet && !S.sheet.error));
    $('dropPdf').classList.toggle('done', !!(S.label && !S.label.error));
    if (S.sheet && !S.sheet.error) $('dropSheet').innerHTML = '다른 코딩시트로 바꾸려면 <b>클릭해서 선택</b>';
    if (S.label && !S.label.error) $('dropPdf').innerHTML = '다른 허가사항으로 바꾸려면 <b>클릭해서 선택</b>';
  }
  var uniq = function (a) { return a.filter(function (x, i) { return a.indexOf(x) === i; }); };

  // ------------------------------------------------------------ 3. 반영 범위
  var SCOPE_NOTE = { '이상반응 표': '각주 포함', '이상반응 본문': '주요 이상반응 등', '용법용량': '투여 중단 기준 등',
    '일반적 주의': '환자 상태 문맥은 약한 근거', '경고': '임시 사용 (사례 확인 전)' };
  function renderScope() {
    $('scope').innerHTML = Object.keys(PMS.SECTION_RANK).map(function (k) {
      var on = S.scope.indexOf(k) >= 0;
      return '<label class="scope' + (on ? '' : ' off') + '"><input type="checkbox" data-scope="' + esc(k) + '"' + (on ? ' checked' : '') + '><span><b>' +
        PMS.SECTION_RANK[k] + '. ' + esc(k) + '</b><small>' + esc(SCOPE_NOTE[k]) + '</small></span></label>';
    }).join('');
  }

  // ------------------------------------------------------------ 준비 (매핑 화면)
  var ready = function () { return S.versions.length && S.label && !S.label.error; };
  function sheetVersions() {
    var latest = S.versions[S.versions.length - 1];
    if (S.mdMode && S.mdMode !== 'auto') return { main: S.mdMode, all: [S.mdMode], latest: S.mdMode, fixed: true };
    if (!S.sheet || S.sheet.error) return { main: latest, all: [latest], latest: latest };
    var cnt = {};
    S.sheet.rows.forEach(function (r) { var v = r.ver || latest; cnt[v] = (cnt[v] || 0) + 1; });
    var all = Object.keys(cnt), main = all.sort(function (a, b) { return cnt[b] - cnt[a]; })[0] || latest;
    return { main: main, all: all, latest: latest };
  }

  async function prepare() {
    updateRun();
    if (!ready()) { renderMap(); return; }
    var vs = sheetVersions();
    if (S.versions.indexOf(vs.main) < 0) { renderMap(); return; }
    var md;
    try { md = await getMd(vs.main); } catch (e) { idle(); $('mdMsg').innerHTML = '<span style="color:var(--err)">' + esc(e.message) + '</span>'; return; }
    var r = PMS.labelTerms(S.label, md, S.termMap, S.scope);
    idle();
    var added = false;
    r.unmapped.forEach(function (u) {
      if (!S.termMap[u.term] && u.candidates && u.candidates.length) { S.termMap[u.term] = [[u.candidates[0].ptEn, '제안']]; added = true; }
    });
    if (added) { saveMap(); r = PMS.labelTerms(S.label, md, S.termMap, S.scope); }
    S.view = { md: md, ver: vs.main, terms: r.terms, unmapped: r.unmapped };
    renderMap(); renderTermList();
    if (S.ran) run();
  }

  // ------------------------------------------------------------ 4. 매핑
  var saveTimer = null;
  function saveMap() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      if (!S.label || S.label.error) return;
      DB.put('termmaps', S.label.product, { map: S.termMap, savedAt: new Date().toISOString() }).then(function () {
        $('mapSaved').textContent = '"' + S.label.product + '" 매핑표가 이 브라우저에 자동 저장돼요.';
      }, function () { $('mapSaved').textContent = '브라우저에 저장하지 못했어요. CSV로 저장해 두세요.'; });
    }, 200);
  }

  function renderMap() {
    var v = S.view;
    $('mapExport').disabled = !v;
    if (!v || !ready()) { $('mapBody').innerHTML = '<p class="muted" style="margin:0">MedDRA 사전과 허가사항 PDF가 준비되면 여기에 표시돼요.</p>'; $('mapState').innerHTML = ''; return; }
    var md = v.md, where = {};
    v.terms.forEach(function (t) { if (t.section === '이상반응 표' && !where[t.term]) where[t.term] = t; });
    v.unmapped.forEach(function (u) { where[u.term] = u; });
    var terms = Object.keys(S.termMap);
    if (!terms.length) {
      $('mapBody').innerHTML = '<div class="dictInfo"><b>확인할 매핑이 없어요</b><span class="muted">허가사항 용어가 모두 MedDRA에 바로 연결돼요.</span></div>';
      $('mapState').innerHTML = '<span class="tag t-반영">확인 완료</span>'; return;
    }
    var pending = terms.filter(function (t) { return S.termMap[t].some(function (r) { return r[1] !== '확정'; }); }).length;
    $('mapState').innerHTML = pending ? '<span class="tag t-확인필요">확인 필요 ' + pending + '개</span>' : '<span class="tag t-반영">모두 확정</span>';
    var krOf = function (pt) { return md.ptKr.get(md.ptCode(pt)) || 'MedDRA v' + v.ver + '에 없는 PT'; };
    var rows = terms.map(function (t) {
      var list = S.termMap[t], u = where[t], main = list[0];
      var cands = (u && u.candidates) || md.suggest(t);
      var opts = cands.map(function (c) { return [c.ptEn, c.ptEn + ' / ' + c.ptKr + ' (유사도 ' + c.score + ')']; });
      if (main[0] && !opts.some(function (o) { return o[0] === main[0]; })) opts.unshift([main[0], main[0] + ' / ' + krOf(main[0])]);
      var sel = '<select data-act="main" data-term="' + esc(t) + '">' + opts.map(function (o) {
        return '<option value="' + esc(o[0]) + '"' + (o[0] === main[0] ? ' selected' : '') + '>' + esc(o[1]) + '</option>'; }).join('') +
        '<option value="__custom">직접 입력…</option></select>';
      var extra = list.slice(1).map(function (r, k) {
        return '<div class="bundle"><span class="tag t-' + esc(r[1]) + '">' + esc(r[1]) + '</span><span>' + esc(r[0]) + '</span><span class="muted">' + esc(krOf(r[0])) +
          '</span><button class="ghost mini" data-act="rm" data-term="' + esc(t) + '" data-k="' + (k + 1) + '">빼기</button></div>';
      }).join('');
      var pos = !u ? '이전에 저장한 매핑' : u.how === '표 각주' ? '이상반응 표 ' + u.detail.split(' / ').slice(-1)[0] : u.section + ' · ' + u.detail;
      return '<tr><td style="width:22%"><b>' + esc(t) + '</b><div class="muted">' + esc(pos) + '</div></td>' +
        '<td style="width:34%">' + sel + '<div class="muted" style="margin-top:4px">' + esc(krOf(main[0])) + '</div></td>' +
        '<td style="width:10%;white-space:nowrap"><span class="tag t-' + esc(main[1]) + '">' + esc(main[1]) + '</span><br><button class="ghost mini" style="margin-top:6px" data-act="toggle" data-term="' + esc(t) + '">' +
        (main[1] === '확정' ? '확정 취소' : '확정') + '</button></td>' +
        '<td>' + (extra || '<span class="muted">없음</span>') + '<div class="addrow"><input data-act="addpt" data-term="' + esc(t) + '" placeholder="PT 영문명 (예: Aspartate aminotransferase increased)"><button class="ghost mini" data-act="add" data-term="' + esc(t) + '">추가</button></div></td></tr>';
    });
    $('mapBody').innerHTML = '<div class="tableWrap" style="max-height:none"><table class="maptbl"><thead><tr><th>허가사항 용어</th><th>대표 PT</th><th>상태</th><th>묶음 PT (같은 근거로 반영)</th></tr></thead><tbody>' +
      rows.join('') + '</tbody></table></div>' +
      (pending ? '<div class="row" style="margin-top:10px"><button class="ghost" data-act="confirm-all">후보 전체 확정</button><span class="muted">후보를 하나씩 확인한 뒤에 눌러 주세요.</span></div>' : '');
  }

  function onMap(e) {
    var el = e.target.closest('[data-act]'); if (!el || !S.view) return;
    var t = el.dataset.term, act = el.dataset.act, md = S.view.md;
    if (e.type === 'change' && act === 'main') {
      var v = el.value;
      if (v === '__custom') {
        v = (prompt('"' + t + '"의 PT 영문명을 입력해 주세요', S.termMap[t][0][0]) || '').trim();
        if (!v) { renderMap(); return; }
        if (!md.ptCode(v)) { alert('MedDRA v' + S.view.ver + '에 "' + v + '" PT가 없어요. 철자를 확인해 주세요.'); renderMap(); return; }
      }
      S.termMap[t][0] = [md.ptEn.get(md.ptCode(v)) || v, '제안'];
    } else if (e.type === 'click' && act === 'toggle') {
      S.termMap[t][0][1] = S.termMap[t][0][1] === '확정' ? '제안' : '확정';
    } else if (e.type === 'click' && act === 'rm') {
      S.termMap[t].splice(Number(el.dataset.k), 1);
    } else if (e.type === 'click' && act === 'add') {
      var inp = el.parentElement.querySelector('[data-act=addpt]'), val = inp.value.trim();
      if (!val) return;
      var code = md.ptCode(val);
      if (!code) { inp.classList.add('bad'); inp.title = 'MedDRA에 없는 PT예요'; return; }
      S.termMap[t].push([md.ptEn.get(code), '확정']);
    } else if (e.type === 'click' && act === 'confirm-all') {
      Object.keys(S.termMap).forEach(function (k) { S.termMap[k].forEach(function (r) { r[1] = '확정'; }); });
    } else return;
    saveMap(); prepare();
  }

  function renderTermList() {
    var seen = {}, rows = [];
    S.view.terms.slice().sort(function (a, b) { return a.rank - b.rank; }).forEach(function (t) {
      var k = [t.ptCode, t.section, t.detail, t.term].join('|'); if (seen[k]) return; seen[k] = 1;
      var strength = t.candidate ? '후보 매핑' : t.weak ? '약함(환자 상태 문맥)' : '강함';
      rows.push('<tr><td>' + esc(t.section) + '</td><td class="muted">' + esc(t.detail) + '</td><td>' + esc(t.term) + '</td><td>' + esc(t.ptEn) + '</td><td>' + esc(t.ptKr) + '</td><td>' + esc(strength) + '</td></tr>');
    });
    $('termListBody').innerHTML = '<div class="tableWrap"><table><thead><tr><th>항목</th><th>위치</th><th>허가사항 용어</th><th>PT</th><th>PT(한글)</th><th>근거 강도</th></tr></thead><tbody>' + rows.join('') + '</tbody></table></div>';
  }

  // ------------------------------------------------------------ 5. 점검
  function updateRun() {
    var okSheet = S.sheet && !S.sheet.error, okPdf = S.label && !S.label.error;
    var missing = [];
    if (S.versions.length && okSheet) {
      var vs = sheetVersions(); missing = vs.all.filter(function (v) { return S.versions.indexOf(v) < 0; });
    }
    $('run').disabled = !(S.versions.length && okSheet && okPdf && !missing.length);
    $('runMsg').innerHTML = !S.versions.length ? 'MedDRA 사전 목록을 불러오지 못했어요. 인터넷 연결을 확인하거나 1번 관리자 기능으로 추가해 주세요.'
      : !okSheet ? '코딩시트를 넣어 주세요.' : !okPdf ? '허가사항 PDF를 넣어 주세요.'
      : missing.length ? '<span style="color:var(--err)">코딩시트가 MedDRA v' + esc(missing.join(', v')) + '를 쓰는데 쿼랑이에 없는 버전이에요. 1번에서 비교할 버전을 고르거나, 관리자 기능으로 추가해 주세요.</span>'
      : S.ran ? (sheetVersions().fixed ? 'MedDRA v' + S.mdMode + '로 비교했어요. ' : '') + '점검을 마쳤어요. 매핑이나 반영 범위를 바꾸면 결과가 바로 다시 계산돼요.' : '준비됐어요. 이상사례 ' + fmt(S.sheet.rows.length) + '건';
  }

  async function run() {
    if ($('run').disabled) return;
    var vs = sheetVersions(), byVer = {};
    try {
      for (var i = 0; i < vs.all.length; i++) {
        var md = await getMd(vs.all[i]);
        byVer[vs.all[i]] = { md: md, terms: PMS.labelTerms(S.label, md, S.termMap, S.scope).terms };
      }
    } catch (e) { idle(); $('runMsg').innerHTML = '<span style="color:var(--err)">' + esc(e.message) + '</span>'; return; }
    idle();
    S.results = S.sheet.rows.map(function (r) {
      var b = byVer[vs.fixed ? vs.main : (r.ver || vs.latest)], j = PMS.judgeRow(b.md, b.terms, r.pt, r.special);
      return Object.assign({}, r, { expt: j[0], ref: j[1] });
    });
    S.ran = true;
    $('result').classList.remove('hidden');
    renderCards(); renderResults(); updateRun();
  }

  var hasExisting = function () { return S.results.some(function (r) { return r.existing.expt; }); };
  var isFlip = function (r) { return (r.existing.expt === '반영' && r.expt === '미반영') || (r.existing.expt === '미반영' && r.expt === '반영'); };
  var isDiff = function (r) { return r.existing.expt && r.existing.expt !== r.expt; };

  function renderCards() {
    var R = S.results, n = function (f) { return R.filter(f).length; };
    var defs = [['반영', '반영', '허가사항에 같은 PT가 있음'], ['확인필요', '확인필요', 'REF 첫 줄의 이유를 확인'], ['미반영', '미반영', '허가사항에서 못 찾음']];
    var html = defs.map(function (d) {
      return '<div class="card c-' + d[1] + (S.filter === d[0] ? ' sel' : '') + '" data-f="' + d[0] + '"><div class="v">' + fmt(n(function (r) { return r.expt === d[0]; })) + '</div><div class="k">' + d[0] + ' · ' + d[2] + '</div></div>';
    }).join('');
    if (hasExisting()) {
      var f = n(isFlip);
      html += '<div class="card c-반대' + (f ? '' : ' zero') + (S.filter === 'flip' ? ' sel' : '') + '" data-f="flip"><div class="v">' + fmt(f) + '</div><div class="k">기존 값과 반대 · 반영↔미반영 (0이 정상)</div></div>';
    } else {
      html += '<div class="card' + (S.filter === 'all' ? ' sel' : '') + '" data-f="all"><div class="v">' + fmt(R.length) + '</div><div class="k">전체 이상사례</div></div>';
    }
    $('cards').innerHTML = html;
    Array.prototype.forEach.call($('cards').children, function (c) {
      c.onclick = function () { S.filter = S.filter === c.dataset.f ? 'all' : c.dataset.f; renderCards(); renderResults(); };
    });
  }

  function renderResults() {
    var q = $('q').value.trim().toUpperCase(), ex = hasExisting(), keep = $('keep').checked;
    var list = S.results.filter(function (r) {
      if (S.filter === 'flip') { if (!isFlip(r)) return false; }
      else if (S.filter !== 'all' && r.expt !== S.filter) return false;
      if (!q) return true;
      return [r.key, r.subj, r.pt, r.ptkor, r.aeterm, r.special, r.ref].join(' ').toUpperCase().indexOf(q) >= 0;
    });
    $('count').textContent = fmt(list.length) + '건 표시';
    $('resHead').innerHTML = '<tr><th style="width:5%">행</th><th style="width:11%">UNIQUE_KEY</th><th style="width:19%">PT</th><th style="width:11%">특수상황</th><th style="width:8%">판정</th><th>REF</th>' + (ex ? '<th style="width:10%">기존 값</th>' : '') + '</tr>';
    $('resRows').innerHTML = list.map(function (r) {
      var ref = r.ref || '', long = ref.split('\n').length > 5 || ref.length > 320;
      var refHtml = long ? '<span class="clip">' + esc(ref.split('\n').slice(0, 4).join('\n').slice(0, 300)) + '…</span><span class="full hidden">' + esc(ref) + '</span>\n<span class="more" data-more>더 보기</span>' : esc(ref);
      var exCell = !ex ? '' : '<td>' + (r.existing.expt ? '<span class="tag t-' + esc(r.existing.expt) + '">' + esc(r.existing.expt) + '</span> ' +
        (isFlip(r) ? '<span class="flip">반대</span>' : isDiff(r) ? '<span class="muted">다름</span>' : '<span class="muted">같음</span>') +
        (keep ? '<div class="muted">내려받을 때 기존 값 유지</div>' : '') : '<span class="muted">비어 있음</span>') + '</td>';
      var sp = r.special && !/^해당\s?없음$/.test(String(r.special).trim()) ? r.special : '';
      return '<tr><td class="muted">' + r.excelRow + '</td><td>' + esc(r.key) + '</td><td><b>' + esc(r.pt) + '</b><div class="muted">' + esc(r.ptkor || '') +
        (r.aeterm ? ' · ' + esc(r.aeterm) : '') + '</div></td><td>' + esc(sp) + '</td><td><span class="tag t-' + esc(r.expt) + '">' + esc(r.expt) + '</span></td><td class="ref">' + refHtml + '</td>' + exCell + '</tr>';
    }).join('') || '<tr><td colspan="7" class="muted">해당하는 행이 없어요.</td></tr>';
  }

  async function download() {
    busy('결과 엑셀 만드는 중…'); await tick();
    try {
      var keep = $('keep').checked, edits = [], filled = 0;
      S.results.forEach(function (r) {
        if (keep && r.existing.expt) return;
        filled++;
        edits.push({ ref: S.sheet.exptLetter + r.excelRow, value: r.expt });
        edits.push({ ref: S.sheet.refLetter + r.excelRow, value: r.ref });
      });
      var out = await XlsxPatch.patchCells(JSZip, S.sheet.bytes, S.sheet.sheetName, edits);
      var name = S.sheet.fileName.replace(/\.xlsx$/i, '') + '_Expectedness.xlsx';
      saveBlob(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
      $('dlMsg').innerHTML = '<span style="color:var(--ok)">✓ ' + fmt(filled) + '개 행을 채운 "' + esc(name) + '" 파일을 내려받았어요.' +
        (filled < S.results.length ? ' 기존 값이 있던 ' + fmt(S.results.length - filled) + '개 행은 그대로 두었어요.' : '') + '</span>';
    } catch (e) { $('dlMsg').innerHTML = '<span style="color:var(--err)">엑셀을 만들지 못했어요: ' + esc(e.message || e) + '</span>'; }
    idle();
  }

  function saveBlob(blob, name) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ------------------------------------------------------------ 연결
  dropZone($('dropMd'), $('pickMd'), buildFromFolder);
  $('mdSel').onchange = function () {
    S.mdMode = $('mdSel').value;
    try { localStorage.setItem(MODE_KEY, S.mdMode); } catch (e) {}
    S.view = null; prepare();
  };
  $('mdAdminMsg').addEventListener('click', function (e) {
    var b = e.target.closest('[data-admin]'); if (!b) return;
    var f = S.adminFiles[Number(b.dataset.admin)];
    saveBlob(new Blob([f.text], { type: f.type }), f.name);
  });
  dropZone($('dropSheet'), $('pickSheet'), loadSheet);
  dropZone($('dropPdf'), $('pickPdf'), loadPdf);
  renderScope();
  $('scope').addEventListener('change', function (e) {
    var k = e.target.dataset.scope; if (!k) return;
    S.scope = Object.keys(PMS.SECTION_RANK).filter(function (x) { return x === k ? e.target.checked : S.scope.indexOf(x) >= 0; });
    renderScope(); prepare();
  });
  $('mapBody').addEventListener('click', onMap);
  $('mapBody').addEventListener('change', onMap);
  $('mapExport').onclick = function () {
    saveBlob(new Blob([PMS.termMapToCsv(S.termMap)], { type: 'text/csv;charset=utf-8' }),
      ('term_map_' + ((S.label && S.label.product) || 'label') + '.csv').replace(/[\\/:*?"<>|]/g, '_'));
  };
  $('mapImport').onclick = function () { $('pickMap').click(); };
  $('pickMap').onchange = async function () {
    var f = $('pickMap').files[0]; $('pickMap').value = ''; if (!f) return;
    S.termMap = PMS.parseTermMapCsv(await f.text());
    $('mapSaved').textContent = f.name + '을 불러왔어요.'; saveMap(); prepare();
  };
  $('run').onclick = run;
  $('dl').onclick = download;
  $('q').oninput = renderResults;
  $('keep').onchange = renderResults;
  $('resRows').addEventListener('click', function (e) {
    var m = e.target.closest('[data-more]'); if (!m) return;
    var td = m.parentElement, full = td.querySelector('.full'), open = full.classList.contains('hidden');
    full.classList.toggle('hidden', !open); td.querySelector('.clip').classList.toggle('hidden', open); m.textContent = open ? '접기' : '더 보기';
  });
  loadCatalog().then(prepare);
})();
