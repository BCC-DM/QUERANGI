/* CM Validation – 학습(기준사전) 핵심 로직
 * 브라우저와 Node 양쪽에서 동작. 입력: SheetJS workbook, 출력: 기준사전 객체 */
(function (root) {
  'use strict';

  // ---------- 1. 컬럼 자동 인식 ----------
  var ALIAS = {
    subj:  ['SUBJID', 'SUBJ_ID', 'USUBJID', 'SUBJECT'],
    trt:   ['CMTRT'],
    dose:  ['CMDOSE', 'CMDOSTOT', 'CMDSTXT', 'DOSE'],
    unit:  ['CMDOSU', 'CMUNIT'],
    freq:  ['CMDOSFRQ', 'CMFREQ', 'CMFEQ', 'CMFRQ'],
    route: ['CMROUTE'],
    code:  ['KIMSMFDS'],
    generic: ['KIMSGENI', 'GENERIC', 'ATCLV5', 'LEVEL5'],
    visit: ['VISIT', 'VISIT_NM', 'REPORT', 'VISITNUM'],
    seq:   ['SEQ', 'CMSEQ', 'ITEMGROUP_REPEATKEY_1']
  };
  var LABEL = { subj: '피험자', trt: '제품명', dose: '용량', unit: '단위', freq: '빈도', route: '경로', code: '품목코드', generic: '성분명', visit: '방문', seq: '순번' };

  function isBlank(v) {
    if (v === null || v === undefined) return true;
    var s = String(v).trim().toUpperCase();
    return s === '' || s === 'NULL' || s === 'NAN' || s === 'NONE' || s === 'N/A';
  }

  // 시트 중 CMTRT가 있는 시트를 찾고, 머리글 줄(처음 10줄 안)을 찾는다
  function findTable(wb, XLSX) {
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var ws = wb.Sheets[wb.SheetNames[i]];
      var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
      var aoaRaw = null; // 셀 실제 값 (숫자 칸이 "1,000"처럼 서식만 붙은 경우 구분용)
      for (var r = 0; r < Math.min(10, aoa.length); r++) {
        var hdr = (aoa[r] || []).map(function (h) { return h == null ? '' : String(h).trim().toUpperCase(); });
        if (hdr.indexOf('CMTRT') >= 0) {
          var map = {};
          Object.keys(ALIAS).forEach(function (k) {
            for (var j = 0; j < ALIAS[k].length; j++) {
              var idx = hdr.indexOf(ALIAS[k][j]);
              if (idx >= 0) { map[k] = { idx: idx, name: ALIAS[k][j] }; break; }
            }
          });
          // DOSE 같은 합친 값("20 mg")보다 숫자 전용 컬럼(CMDOSE)을 우선
          aoaRaw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
          return { sheet: wb.SheetNames[i], header: hdr, map: map, rows: aoa.slice(r + 1), rawRows: aoaRaw.slice(r + 1), headerRow: r + 1 };
        }
      }
    }
    return null;
  }

  // ---------- 2. 표기 정리 ----------
  var UNIT = {
    tab: 'tab', tablet: 'tab', tabs: 'tab', t: 'tab', '정': 'tab',
    cap: 'cap', capsule: 'cap', caps: 'cap', '캡슐': 'cap',
    amp: 'amp', ampule: 'amp', ampoule: 'amp', ample: 'amp', '앰플': 'amp',
    ml: 'ml', cc: 'ml', mcg: 'mcg', microgram: 'mcg', micrograms: 'mcg', 'μg': 'mcg', 'µg': 'mcg', ug: 'mcg', '㎍': 'mcg',
    mg: 'mg', '㎎': 'mg', g: 'g', iu: 'IU', unit: 'IU', units: 'IU', u: 'IU',
    // 풀어 쓴 단위 (11-23 지후 결정)
    milligram: 'mg', milligrams: 'mg', '밀리그램': 'mg', gram: 'g', grams: 'g', '그램': 'g', '마이크로그램': 'mcg',
    milliliter: 'ml', milliliters: 'ml', millilitre: 'ml', millilitres: 'ml', '밀리리터': 'ml', '㎖': 'ml',
    liter: 'l', liters: 'l', litre: 'l', litres: 'l', '리터': 'l',
    'international unit': 'IU', 'international units': 'IU', '국제단위': 'IU', '단위': 'IU',
    drop: 'drop', drops: 'drop', gtt: 'drop', pack: 'pack', pk: 'pack', pkg: 'pack', '포': 'pack',
    vial: 'vial', bag: 'bag', patch: 'patch', puff: 'puff', '분무': 'puff', spray: 'puff',
    syrn: 'syringe', syringe: 'syringe', bottle: 'bottle', btl: 'bottle', ea: 'ea',
    fingerspoon: 'FTU', ftu: 'FTU', fingertip: 'FTU', 'finger tip': 'FTU', 'finger tip unit': 'FTU', 'fingertip unit': 'FTU',
    syr: 'syringe', '나잘스프레이': 'puff',
    supp: 'supp', 'supp.': 'supp', suppository: 'supp', suppositories: 'supp', '좌약': 'supp', '좌제': 'supp', uk: 'UK', unk: 'UK', unknown: 'UK', '모름': 'UK', dose: 'dose'
  };
  var FREQ = {
    qd: 'QD', od: 'QD', '1': 'QD', daily: 'QD', bid: 'BID', tid: 'TID', qid: 'QID', prn: 'PRN',
    once: 'ONCE', '1회': 'ONCE', stat: 'ONCE',
    'every week': 'QW', qw: 'QW', weekly: 'QW', 'every month': 'QM', 'q month': 'QM', monthly: 'QM',
    'every 2month': 'Q2M', 'every 2 months': 'Q2M', 'every 6months': 'Q6M', 'every 6 months': 'Q6M',
    'q 6months': 'Q6M', 'q 6 months': 'Q6M', 'q6m': 'Q6M', '6개월간격': 'Q6M', '6개월 간격': 'Q6M',
    '3주 마다': 'Q3W', '3주마다': 'Q3W', '3일마다': 'Q3D', '3일 마다': 'Q3D',
    'every 48 hours': 'Q48H', 'q48h': 'Q48H'
  };
  var ROUTE = {
    po: 'PO', oral: 'PO', iv: 'IV', im: 'IM', sc: 'SC', subcutaneous: 'SC',
    inhalation: 'INHALATION', inhal: 'INHALATION', inh: 'INHALATION',
    eye: 'OPHTHALMIC', 'eye drop': 'OPHTHALMIC', ophthalmic: 'OPHTHALMIC', opthalmic: 'OPHTHALMIC', oph: 'OPHTHALMIC',
    ou: 'OPHTHALMIC', od: 'OPHTHALMIC', os: 'OPHTHALMIC', '점안': 'OPHTHALMIC',
    topical: 'TOPICAL', derma: 'TOPICAL', dermal: 'TOPICAL', epidermal: 'TOPICAL', cutaneous: 'TOPICAL', skin: 'TOPICAL',
    ext: 'TOPICAL', '외용': 'TOPICAL', '국소': 'TOPICAL', '피부에 도포': 'TOPICAL', '피부': 'TOPICAL',
    td: 'TRANSDERMAL', transdermal: 'TRANSDERMAL', '피부 부착': 'TRANSDERMAL', '피부부착': 'TRANSDERMAL',
    nasal: 'NASAL', '코세척': 'NASAL', 'intra-articular': 'INTRA-ARTICULAR', ia: 'INTRA-ARTICULAR', '관절강내': 'INTRA-ARTICULAR',
    vaginal: 'VAGINAL', pv: 'VAGINAL', '질내': 'VAGINAL', '질': 'VAGINAL', sublingual: 'SUBLINGUAL', sl: 'SUBLINGUAL', '설하': 'SUBLINGUAL',
    gargle: 'GARGLE', '가글': 'GARGLE', '구강': 'GARGLE', ivt: 'INTRAVITREAL', ear: 'OTIC', otic: 'OTIC',
    '방광내주입': 'INTRAVESICAL', intravesical: 'INTRAVESICAL', rectal: 'RECTAL', pr: 'RECTAL', anus: 'RECTAL', '항문': 'RECTAL', '직장': 'RECTAL'
  };
  function stripOther(s) { var m = /^other\((.*)\)$/i.exec(s); return (m ? m[1] : s).trim(); }
  function normUnit(v) { if (isBlank(v)) return null; var s = stripOther(String(v).trim()); return UNIT[s.toLowerCase()] || s.toLowerCase(); }
  function normFreq(v) {
    if (isBlank(v)) return null;
    var s = String(v).trim().toLowerCase().replace(/\s+/g, ' ').replace(/\(.*\)$/, '').trim();
    if (s === 'uk' || s === 'unk' || s === 'unknown') return null;
    return FREQ[s] || s.toUpperCase();
  }
  function normRoute(v) { if (isBlank(v)) return null; var s = stripOther(String(v).trim()); return ROUTE[s.toLowerCase()] || s.toUpperCase(); }
  function normName(v) { return isBlank(v) ? null : String(v).replace(/\s+/g, '').toUpperCase(); }

  // 용량: CMDOSE에는 숫자만 (11-23 지후 결정). 허용: 1, 12, 0.5, 12.5 / 오류: .5, 5., 0,5, 1,000, 05, 20mg
  var DOSE_RE = /^-?(0|[1-9][0-9]*)(\.[0-9]+)?$/;
  function parseDose(v, unit) {
    if (isBlank(v)) return { dose: null, unit: unit };
    if (typeof v === 'number') return { dose: isFinite(v) ? v : null, unit: unit };
    var s = String(v).trim();
    return { dose: DOSE_RE.test(s) ? parseFloat(s) : null, unit: unit };
  }
  // 용량이 숫자 형식이 아닐 때 이유와 안내
  function doseProblem(raw) {
    var s = String(raw).trim();
    if (/월|일|\d{4}-\d{2}/.test(s)) return ['날짜 형식', '엑셀이 분수(예: 1/2)를 날짜로 바꾼 것으로 보여요. 0.5처럼 소수로 입력'];
    if (/^[0-9.,\s]*[0-9][0-9.,\s]*[^0-9.,\s\/~\-].*$/.test(s)) return ['용량 칸에 단위·글자가 같이 있어요', 'CMDOSE에는 숫자만, 단위는 CMDOSU에 입력'];
    if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) return ['천 단위 쉼표가 있어요', '쉼표 없이 입력 (예: 1000)'];
    if (/,/.test(s)) return ['소수점 자리에 쉼표가 있어요', '소수점은 점(.)으로 입력 (예: 0.5)'];
    if (/\..*\./.test(s)) return ['소수점이 두 개 이상이에요', '숫자를 다시 확인 (예: 1.5)'];
    if (/^\.\d/.test(s)) return ['소수점 앞에 숫자가 없어요', '0.5처럼 0을 붙여 입력'];
    if (/^\d+\.$/.test(s)) return ['소수점 뒤에 숫자가 없어요', '5 또는 5.0처럼 입력'];
    if (/^0\d/.test(s)) return ['0으로 시작하는 숫자예요', '0.5를 잘못 쓴 건지 확인 (예: 05 → 0.5 또는 5)'];
    return ['숫자가 아니에요', '숫자만 입력 (예: 0.5)'];
  }

  // ---------- 3. 함량 추출과 mg 환산 ----------
  // 제품명의 한글 함량 단위 → 영문 (11-23 지후 결정)
  var KOR_STR = { '밀리그램': 'mg', '마이크로그램': 'mcg', '그램': 'g', '국제단위': 'iu' };
  var MASS = { mg: 1, mcg: 0.001, g: 1000 };
  // 제품명에서 함량 하나만 있을 때 반환. "아모잘탄 5/50mg"처럼 복합제면 null
  function strengthFromName(name) {
    if (!name) return null;
    var s = String(name);
    if (/\d\s*\/\s*\d/.test(s)) return null;
    var re = /(\d+(?:\.\d+)?)\s*(mg|mcg|μg|µg|㎍|㎎|ug|g|iu|IU|밀리그램|마이크로그램|그램|국제단위)(?![a-z])(?!\s*\/)/gi, m, found = [];
    while ((m = re.exec(s)) !== null) {
      var u = m[2].toLowerCase(); u = KOR_STR[u] || UNIT[u] || u;
      found.push({ v: parseFloat(m[1]), u: u === 'iu' ? 'IU' : u });
    }
    if (!found.length) return null;
    var key = found[0].v + found[0].u;
    for (var i = 1; i < found.length; i++) if (found[i].v + found[i].u !== key) return null;
    return found[0];
  }
  // 질량 단위는 모두 mg로 통일. tab/cap은 함량을 알면 곱해서 환산
  function canonical(dose, unit, strength) {
    if (dose == null || !unit) return null;
    var d = dose, u = unit, conv = false;
    if ((u === 'tab' || u === 'cap' || u === 'puff') && strength) { d = dose * strength.v; u = strength.u; conv = true; }
    if (MASS[u] != null) { d = d * MASS[u]; u = 'mg'; }
    d = Math.round(d * 1e6) / 1e6;
    return { text: d + ' ' + u, conv: conv };
  }

  // 제품명 끝의 제형: 함량·용량·괄호를 떼어낸 뒤 마지막 글자로 판단
  // 예: '레토프라 정 10mg' → 정, '엘도스 캡슐' → 캡슐, '타조락탐 주 4.5g' → 주
  function formOf(name) {
    if (!name) return null;
    var s = String(name).trim(), prev;
    do {
      prev = s;
      s = s.replace(/\s*\([^()]*\)\s*$/, '')
           .replace(/\s*[\d.,\/]+\s*(mg|mcg|μg|µg|㎍|㎎|g|ml|㎖|l|iu|units?|%|밀리그램|마이크로그램|그램|국제단위|밀리리터|mg\/ml|mg\/g)?\s*$/i, '')
           .replace(/[\s.\-_,]+$/, '');
    } while (s !== prev && s.length);
    var m;
    if ((m = /(캡슐|캡)$/.exec(s))) return { form: 'cap', word: m[1] };
    if (/정$/.test(s)) return { form: 'tab', word: '정' };
    if ((m = /(주사액|주사제|주사|주)$/.exec(s))) return { form: 'inj', word: m[1] };
    // 영문 제형 표기: 'Pletaal tab. 100mg', 'CELEBREX CAP 100mg', 'xxx inj.'
    if ((m = /(?:^|[\s.\-])(tabs?|tablets?)$/i.exec(s))) return { form: 'tab', word: m[1] };
    if ((m = /(?:^|[\s.\-])(caps?|capsules?)$/i.exec(s))) return { form: 'cap', word: m[1] };
    if ((m = /(?:^|[\s.\-])(inj|injection)$/i.exec(s))) return { form: 'inj', word: m[1] };
    return null;
  }

  // ---------- 4. 파일 한 개 읽기 ----------
  function readFile(wb, XLSX, label) {
    var t = findTable(wb, XLSX);
    if (!t) return { label: label, ok: false, reason: 'CMTRT 컬럼을 찾지 못했어요', rows: [] };
    var need = ['subj', 'trt', 'dose', 'unit', 'route'], miss = need.filter(function (k) { return !t.map[k]; });
    if (miss.length) return { label: label, ok: false, reason: '필수 컬럼 없음: ' + miss.map(function (k) { return LABEL[k]; }).join(', '), rows: [], table: t };
    function get(r, k) { return t.map[k] ? r[t.map[k].idx] : null; }
    var data = [];
    t.rows.forEach(function (r, i) { if (r && !isBlank(get(r, 'trt'))) { r.__row = t.headerRow + 1 + i; r.__raw = t.rawRows ? t.rawRows[i] : null; data.push(r); } });
    // 용량: 엑셀 숫자 칸은 실제 값을 씀 (서식 때문에 "1,000"으로 보여도 값이 1000이면 정상). 날짜로 바뀐 칸은 글자 그대로 둬서 오류로 잡음
    function getDose(r) {
      var txt = get(r, 'dose'), rv = (r.__raw && t.map.dose) ? r.__raw[t.map.dose.idx] : null;
      if (typeof rv === 'number' && txt != null && /^-?[0-9,]*\.?[0-9]+$/.test(String(txt).trim())) return rv;
      return txt;
    }
    // 숫자 코드 단위 감지 (예: 1, 4, 10) → 코드표 없이는 학습 불가
    var units = data.map(function (r) { return get(r, 'unit'); }).filter(function (v) { return !isBlank(v); });
    var numeric = units.filter(function (v) { return /^\d+$/.test(String(v).trim()); }).length;
    if (units.length && numeric / units.length > 0.8)
      return { label: label, ok: false, reason: '단위·빈도·경로가 숫자 코드로 되어 있어요 (코드표 필요)', rows: [], table: t };
    var rows = data.map(function (r) {
      var unit = normUnit(get(r, 'unit')), pd = parseDose(getDose(r), unit), code = get(r, 'code');
      code = (!isBlank(code) && /^\d+$/.test(String(code).trim())) ? String(code).trim() : null;
      return {
        subj: label + '|' + String(get(r, 'subj')), raw: String(get(r, 'trt')).trim(), name: normName(get(r, 'trt')),
        code: code, generic: isBlank(get(r, 'generic')) ? null : String(get(r, 'generic')).trim(),
        dose: pd.dose, unit: pd.unit, freq: t.map.freq ? normFreq(get(r, 'freq')) : null, route: normRoute(get(r, 'route')),
        rowNo: r.__row, subjRaw: get(r, 'subj'), visit: get(r, 'visit'), seq: get(r, 'seq'),
        rawDose: get(r, 'dose'), rawUnit: get(r, 'unit'), rawFreq: get(r, 'freq'), rawRoute: get(r, 'route')
      };
    });
    return { label: label, ok: true, rows: rows, table: t, hasFreq: !!t.map.freq };
  }

  // ---------- 5. 기준사전 만들기 ----------
  function build(files, params) {
    params = params || {};
    var minSubj = params.minSubj || 5, rareCut = params.rareCut || 0.05;
    var all = [];
    files.forEach(function (f) { if (f.ok) all = all.concat(f.rows.map(function (r) { r.study = f.label; return r; })); });

    // 제품명 → 품목코드 연결표 (코드가 있는 행에서 가장 많이 쓰인 코드)
    var nc = {};
    all.forEach(function (r) { if (r.code && r.name) { nc[r.name] = nc[r.name] || {}; nc[r.name][r.code] = (nc[r.name][r.code] || 0) + 1; } });
    var nameToCode = {};
    Object.keys(nc).forEach(function (n) { nameToCode[n] = Object.keys(nc[n]).sort(function (a, b) { return nc[n][b] - nc[n][a]; })[0]; });
    function keyOf(r) { return r.code ? 'K:' + r.code : (r.name && nameToCode[r.name] ? 'K:' + nameToCode[r.name] : (r.name ? 'N:' + r.name : null)); }

    var groups = {};
    all.forEach(function (r) { var k = keyOf(r); if (!k) return; r.key = k; (groups[k] = groups[k] || []).push(r); });

    var products = {}, stats = { rows: all.length, byCode: 0, byLink: 0, byName: 0, converted: 0 };
    all.forEach(function (r) { if (!r.key) return; if (r.code) stats.byCode++; else if (r.key[0] === 'K') stats.byLink++; else stats.byName++; });

    Object.keys(groups).forEach(function (k) {
      var g = groups[k], names = {}, generics = {}, studies = {}, subj = {};
      g.forEach(function (r) { names[r.raw] = (names[r.raw] || 0) + 1; if (r.generic) generics[r.generic] = 1; studies[r.study] = 1; subj[r.subj] = 1; });
      // 함량: 이 제품에 쓰인 이름들에서 같은 함량이 하나로만 나오면 사용
      var str = null, ok = true;
      Object.keys(names).forEach(function (n) { var s = strengthFromName(n); if (!s) return; if (!str) str = s; else if (str.v !== s.v || str.u !== s.u) ok = false; });
      if (!ok) str = null;
      function tally(fn) {
        var seen = {}, out = {};
        g.forEach(function (r) { var v = fn(r); if (v == null) return; var sk = r.subj + '\u0001' + v; if (seen[sk]) return; seen[sk] = 1; out[v] = (out[v] || 0) + 1; });
        return out;
      }
      var dose = tally(function (r) { var c = canonical(r.dose, r.unit, str); if (c && c.conv) r._conv = 1; return c ? c.text : null; });
      g.forEach(function (r) { if (r._conv) stats.converted++; });
      var topName = Object.keys(names).sort(function (a, b) { return names[b] - names[a]; })[0];
      products[k] = {
        name: topName, names: Object.keys(names).length > 1 ? Object.keys(names) : undefined,
        generic: Object.keys(generics)[0] || null, strength: str, nSubj: Object.keys(subj).length,
        studies: Object.keys(studies).sort(), dose: dose, freq: tally(function (r) { return r.freq; }), route: tally(function (r) { return r.route; })
      };
    });

    var nameToKey = {};
    all.forEach(function (r) { if (r.name && r.key) nameToKey[r.name] = r.key; });
    var keys = Object.keys(products);
    var judged = keys.filter(function (k) { return products[k].nSubj >= minSubj; });
    var judgedSet = {}; judged.forEach(function (k) { judgedSet[k] = 1; });
    stats.products = keys.length; stats.judgedProducts = judged.length;
    stats.judgedRows = all.filter(function (r) { return judgedSet[r.key]; }).length;
    stats.withStrength = keys.filter(function (k) { return products[k].strength; }).length;

    return {
      meta: {
        version: 1, created: new Date().toISOString(), params: { minSubj: minSubj, rareCut: rareCut },
        files: files.map(function (f) { return { label: f.label, used: f.ok, rows: f.rows.length, reason: f.reason || null }; }),
        stats: stats
      },
      nameToKey: nameToKey, products: products
    };
  }

  // ---------- 6. 검증 ----------
  var UNIT_OK = ['mg', 'mcg', 'g', 'ml', 'IU', 'tab', 'cap', 'amp', 'vial', 'drop', 'pack', 'puff', 'patch', 'bag', 'syringe', 'bottle', 'FTU', 'ea', 'kit', 'l', 'meq', 'mmol', 'dose', 'UK', 'box', 'apply', '회', 'supp'];
  var FREQ_OK = ['QD', 'BID', 'TID', 'QID', 'PRN', 'ONCE', 'QW', 'QM', 'Q2M', 'Q3M', 'Q6M', 'Q2W', 'Q3W', 'Q3D', 'Q48H', 'QOD', 'Q4H', 'Q6H', 'Q8H', 'Q12H'];
  var ROUTE_OK = ['PO', 'IV', 'IM', 'SC', 'INHALATION', 'OPHTHALMIC', 'TOPICAL', 'TRANSDERMAL', 'NASAL', 'INTRA-ARTICULAR', 'VAGINAL', 'SUBLINGUAL', 'GARGLE', 'INTRAVITREAL', 'OTIC', 'RECTAL', 'INTRADERMAL', 'INTRAVESICAL'];
  var LEVEL = { '오류': 3, '확인 필요': 2, '판정 불가': 1, '정상': 0 };

  function pct(n, d) { return Math.round(n / d * 100) + '%'; }
  function topVals(obj, n) {
    var tot = 0; Object.keys(obj).forEach(function (k) { tot += obj[k]; });
    return Object.keys(obj).sort(function (a, b) { return obj[b] - obj[a]; }).slice(0, n || 3)
      .map(function (k) { return k + ' ' + obj[k] + '명(' + pct(obj[k], tot) + ')'; }).join(', ');
  }

  // 이슈 리스트 엑셀에서 노란색으로 칠할 칸 (trt=CMTRT, dose=CMDOSE, unit=CMDOSU, freq=CMFREQ, route=CMROUTE)
  function fieldsFor(rule, msg) {
    switch (rule) {
      case 'STRUCT-01': return ['dose'];
      case 'STRUCT-02': {
        var f = [];
        if (msg.indexOf('용량') >= 0) f.push('dose');
        if (msg.indexOf('단위') >= 0) f.push('unit');
        if (msg.indexOf('빈도') >= 0) f.push('freq');
        if (msg.indexOf('경로') >= 0) f.push('route');
        return f;
      }
      case 'STRUCT-03': return ['unit'];
      case 'STRUCT-04': return msg.indexOf('빈도') === 0 ? ['freq'] : ['route'];
      case 'FORM-01': case 'FORM-02': return ['trt', 'unit'];
      case 'FORM-03': return ['trt', 'route'];
      case 'FORM-04': return ['trt', 'unit'];
      case 'UNIT-01': case 'UNIT-02': case 'UNIT-03': case 'UNIT-04': case 'UNIT-05': return ['unit', 'route'];
      case 'DOSE-01': case 'DICT-DOSE': case 'DICT-UNIT': return ['dose', 'unit'];
      case 'DICT-FREQ': return ['freq'];
      case 'DICT-ROUTE': return ['route'];
      default: return rule.indexOf('NAME-') === 0 ? ['trt', 'route'] : [];
    }
  }

  function checkRow(r, dic) {
    var issues = [], P = dic.products, minSubj = dic.meta.params.minSubj, cut = dic.meta.params.rareCut;
    function add(level, rule, msg, ref) { issues.push({ level: level, rule: rule, msg: msg, ref: ref || '', fields: fieldsFor(rule, msg) }); }
    var raw = String(r.rawDose == null ? '' : r.rawDose).trim(), doseUK = /^(UK|UNK|UNKNOWN|모름)$/i.test(raw);
    var doseBad = false;

    // 1단계: 형식 검사
    if (!isBlank(r.rawDose) && !doseUK && r.dose == null) { var dp = doseProblem(raw); add('오류', 'STRUCT-01', '용량 형식 오류(' + dp[0] + '): "' + raw + '"', dp[1]); doseBad = true; }
    else if (r.dose != null && r.dose <= 0) { add('오류', 'STRUCT-01', '용량이 0 이하예요: ' + raw); doseBad = true; }
    var miss = [];
    if (isBlank(r.rawDose)) miss.push('용량');
    if (isBlank(r.rawUnit) && !(r.unit)) miss.push('단위');
    if (r.hasFreq && isBlank(r.rawFreq)) miss.push('빈도');
    if (isBlank(r.rawRoute)) miss.push('경로');
    if (miss.length) { add('오류', 'STRUCT-02', '필수값이 비어 있어요: ' + miss.join(', ')); if (miss.indexOf('용량') >= 0) doseBad = true; }
    if (r.unit && UNIT_OK.indexOf(r.unit) < 0) { add('오류', 'STRUCT-03', '허용 단위가 아니에요: "' + r.rawUnit + '"', '허용: ' + UNIT_OK.join(', ')); doseBad = true; }
    var freqBad = !!(r.freq && FREQ_OK.indexOf(r.freq) < 0), routeBad = !!(r.route && ROUTE_OK.indexOf(r.route) < 0);
    if (freqBad) add('오류', 'STRUCT-04', '빈도가 표준 표기가 아니에요: "' + r.rawFreq + '"', '예: QD, BID, TID, QID, PRN, ONCE, QW, QM');
    if (routeBad) add('오류', 'STRUCT-04', '경로가 표준 표기가 아니에요: "' + r.rawRoute + '"', '예: PO, IV, IM, SC, INHALATION, OPHTHALMIC, TOPICAL');

    // 제형 규칙 (11-18 지후 결정)
    var fm = formOf(r.raw);
    if (fm && fm.form === 'tab' && r.unit === 'cap')
      add('오류', 'FORM-01', '제품명이 "' + fm.word + '"(으)로 끝나는데 단위가 캡슐이에요: "' + r.rawUnit + '"', '정제는 tab으로 입력');
    if (fm && fm.form === 'cap' && r.unit === 'tab')
      add('오류', 'FORM-02', '제품명이 "' + fm.word + '"(으)로 끝나는데 단위가 정이에요: "' + r.rawUnit + '"', '캡슐은 cap으로 입력');
    if (fm && fm.form === 'inj' && r.route && !routeBad && ['IV', 'IM', 'SC'].indexOf(r.route) < 0) {
      add('오류', 'FORM-03', '제품명이 "' + fm.word + '"(으)로 끝나는데 경로가 IV·IM·SC가 아니에요: "' + r.rawRoute + '"', '주사제는 IV, IM, SC 중 하나로 입력');
      routeBad = true;
    }

    // 11-19 추가: 단위·제품명과 경로가 모순되는 경우
    if (fm && (fm.form === 'tab' || fm.form === 'cap') && ['ml', 'drop', 'puff'].indexOf(r.unit) >= 0)
      add('오류', 'FORM-04', '제품명이 "' + fm.word + '"(으)로 끝나는 고형제인데 단위가 액체·분무 단위예요: "' + r.rawUnit + '"', '정제·캡슐은 tab, cap 또는 mg 등으로 입력');
    var rt = (r.route && !routeBad) ? r.route : null;
    function routeErr(rule, level, msg, ref) { add(level, rule, msg, ref); routeBad = true; rt = null; }
    if (rt) {
      var TAB_OK = ['PO', 'SUBLINGUAL', 'VAGINAL'], CAP_OK = ['PO', 'VAGINAL', 'INHALATION'];
      if (r.unit === 'tab' && TAB_OK.indexOf(rt) < 0)
        routeErr('UNIT-01', '오류', '단위가 정(tab)인데 경로가 ' + r.rawRoute + '이에요', 'tab은 PO, SUBLINGUAL(설하), VAGINAL(질정)만 가능');
      else if (r.unit === 'cap' && CAP_OK.indexOf(rt) < 0)
        routeErr('UNIT-01', '오류', '단위가 캡슐(cap)인데 경로가 ' + r.rawRoute + '이에요', 'cap은 PO, VAGINAL(질캡슐), INHALATION(흡입용 캡슐)만 가능');
      else if (['amp', 'vial', 'syringe', 'bag'].indexOf(r.unit) >= 0 && (rt === 'PO' || rt === 'TOPICAL'))
        routeErr('UNIT-02', '오류', '주사 단위(' + r.rawUnit + ')인데 경로가 ' + r.rawRoute + '이에요', '주사 단위는 IV, IM, SC 등 주사·흡입·세척 경로로 입력');
      else if (r.unit === 'puff' && ['INHALATION', 'NASAL', 'SUBLINGUAL'].indexOf(rt) < 0)
        routeErr('UNIT-03', '오류', '단위가 puff(흡입·분무)인데 경로가 ' + r.rawRoute + '이에요', 'puff는 INHALATION, NASAL, SUBLINGUAL(설하 스프레이)만 가능');
      else if (r.unit === 'drop' && ['IV', 'IM', 'SC'].indexOf(rt) >= 0)
        routeErr('UNIT-04', '오류', '단위가 drop인데 경로가 ' + r.rawRoute + '이에요', 'drop은 OPHTHALMIC, OTIC, NASAL, PO(경구 드롭) 등으로 입력');
      else if (r.unit === 'supp' && ['RECTAL', 'VAGINAL'].indexOf(rt) < 0)
        routeErr('UNIT-05', '오류', '단위가 좌약(supp)인데 경로가 ' + r.rawRoute + '이에요', '좌약은 RECTAL(항문·직장) 또는 VAGINAL(질)로 입력');
    }
    if (rt && r.raw) {
      var nm = String(r.raw);
      var isEye = /(점안|안약|안연고|eye ?drop|eye ?oint|opht)/i.test(nm);
      var isNasal = /(나잘|nasal|비강|비액)/i.test(nm);
      var isInj = /(주사|inj)/i.test(nm) || (fm && fm.form === 'inj');
      if (isEye && rt !== 'OPHTHALMIC')
        routeErr('NAME-01', '오류', '점안제인데 경로가 ' + r.rawRoute + '이에요', '점안제는 OPHTHALMIC으로 입력');
      else if (isNasal && rt !== 'NASAL')
        routeErr('NAME-02', '확인 필요', '비강 제품(나잘·nasal)인데 경로가 ' + r.rawRoute + '이에요', '비강 스프레이는 보통 NASAL로 입력');
      else if (!isNasal && /(흡입|할러|엘립타|터부헬러|디스커스|네뷸|레스퓰|inhal|ellipta|handihaler|respimat|turbuhaler|diskus|nebul)/i.test(nm) && rt !== 'INHALATION')
        routeErr('NAME-03', '오류', '흡입제인데 경로가 ' + r.rawRoute + '이에요', '흡입제는 INHALATION으로 입력');
      else if (/(패취|패치|플라스타|카타플라스마|patch|plaster)/i.test(nm) && ['TRANSDERMAL', 'TOPICAL'].indexOf(rt) < 0)
        routeErr('NAME-04', '오류', '패취제인데 경로가 ' + r.rawRoute + '이에요', '패취제는 TRANSDERMAL 또는 TOPICAL로 입력');
      else if (!isEye && !isNasal && !isInj && /(시럽|syrup|현탁|suspension|내복액)/i.test(nm) && rt !== 'PO')
        routeErr('NAME-05', '오류', '시럽·현탁액(경구 액제)인데 경로가 ' + r.rawRoute + '이에요', '경구 액제는 PO로 입력');
      else if (!isEye && /(연고|크림|로션|cream|lotion|ointment|oint)/i.test(nm) && rt === 'PO')
        routeErr('NAME-06', '오류', '연고·크림·로션(외용제)인데 경로가 PO예요', '외용제는 TOPICAL 등으로 입력');
      else if (/(좌제|좌약|suppos)/i.test(nm) && ['RECTAL', 'VAGINAL'].indexOf(rt) < 0)
        routeErr('NAME-07', '오류', '좌제인데 경로가 ' + r.rawRoute + '이에요', '좌제는 RECTAL 또는 VAGINAL로 입력');
      else if (/(질정|(^|[^연])질캡슐|질연질캡슐|질좌|질크림|vaginal)/i.test(nm) && rt !== 'VAGINAL')
        routeErr('NAME-08', '오류', '질 제품인데 경로가 ' + r.rawRoute + '이에요', '질정·질캡슐은 VAGINAL로 입력');
    }

    // 제품 찾기
    var key = r.code ? 'K:' + r.code : (r.name && dic.nameToKey[r.name] ? dic.nameToKey[r.name] : (r.name ? 'N:' + r.name : null));
    var p = key ? P[key] : null;
    var str = strengthFromName(r.raw) || (p && p.strength) || null;

    // 2단계: 제품 함량 대조
    if (!doseBad && r.dose != null && str && str.u !== 'IU') {
      var sMg = str.v * MASS[str.u];
      if (MASS[r.unit] != null && sMg) {
        var dMg = r.dose * MASS[r.unit], ratio = dMg / sMg;
        if (ratio >= 10) { doseBad = true; } if (ratio >= 10) add('오류', 'DOSE-01', '1회 용량이 함량의 ' + Math.round(ratio) + '배예요 (' + r.dose + ' ' + r.unit + ', 함량 ' + str.v + ' ' + str.u + ')', '자릿수·단위 오류 의심. 정 수로 환산하면 ' + Math.round(ratio * 10) / 10 + '정');
        else if (ratio < 0.1) { doseBad = true; add('오류', 'DOSE-01', '1회 용량이 함량의 ' + (Math.round(ratio * 1000) / 1000) + '배예요 (' + r.dose + ' ' + r.unit + ', 함량 ' + str.v + ' ' + str.u + ')', '단위(mg·g·mcg)를 잘못 골랐거나, 정·캡슐·바이알 수로 써야 했는지 확인'); }
      } else if ((r.unit === 'tab' || r.unit === 'cap') && r.dose >= 10) {
        doseBad = true;
        add('오류', 'DOSE-01', '한 번에 ' + r.dose + ' ' + r.unit + '예요 (함량 ' + str.v + ' ' + str.u + ')', r.dose === str.v ? '함량 숫자를 정 수 자리에 넣은 것으로 보여요. 단위를 ' + str.u + '로 확인' : '정·캡슐 수 확인');
      }
    }

    // 3단계: 학습 사전 대조
    var judged = !!(p && p.nSubj >= minSubj);
    if (judged) {
      function rare(obj, val, label, skip) {
        if (skip || val == null) return;
        var tot = 0; Object.keys(obj).forEach(function (k) { tot += obj[k]; });
        if (tot < minSubj) return;               // 그 항목은 학습 자료가 적음
        var n = obj[val] || 0;
        if (n / tot < cut) add('확인 필요', 'DICT-' + label.code, '이 제품에서 드문 ' + label.ko + ': ' + val + ' (' + n + '/' + tot + '명, ' + pct(n, tot) + ')', '자주 쓰인 값: ' + topVals(obj));
      }
      var c = canonical(r.dose, r.unit, p.strength || str), unitMismatch = false;
      if (c && !(p.strength || str) && !doseBad) {
        // 함량을 몰라 mg와 정·캡슐 수(또는 ml와 포 등)를 서로 바꿀 수 없는 제품
        var byUnit = {}, totU = 0;
        Object.keys(p.dose).forEach(function (k) { var u = k.split(' ').slice(1).join(' '); byUnit[u] = (byUnit[u] || 0) + p.dose[k]; totU += p.dose[k]; });
        var mainU = Object.keys(byUnit).sort(function (a, b) { return byUnit[b] - byUnit[a]; })[0];
        var myU = c.text.split(' ').slice(1).join(' ');
        if (totU >= minSubj && mainU && myU !== mainU) {
          unitMismatch = true;
          add('확인 필요', 'DICT-UNIT', '제품명에 함량이 없어 ' + myU + ' 입력을 학습값과 비교할 수 없어요: ' + c.text,
            '이 제품은 주로 ' + mainU + '로 입력됐어요 (' + byUnit[mainU] + '/' + totU + '명). 자주 쓰인 값: ' + topVals(p.dose));
        }
      }
      rare(p.dose, c ? c.text : null, { code: 'DOSE', ko: '용량이에요' }, doseBad || unitMismatch);
      rare(p.freq, r.freq, { code: 'FREQ', ko: '빈도예요' }, freqBad);
      rare(p.route, r.route, { code: 'ROUTE', ko: '경로예요' }, routeBad);
    }
    var level = issues.reduce(function (m, x) { return LEVEL[x.level] > LEVEL[m] ? x.level : m; }, '정상');
    if (level === '정상' && !judged) level = '판정 불가';
    return { level: level, issues: issues, key: key, product: p, judged: judged,
      note: !judged ? (p ? '학습 피험자 ' + p.nSubj + '명 (' + minSubj + '명 미만)' : '학습 데이터에 없는 제품') : '' };
  }

  function validate(file, dic) {
    var out = file.rows.map(function (r) { r.hasFreq = file.hasFreq; var res = checkRow(r, dic); res.row = r; return res; });
    var sum = { '오류': 0, '확인 필요': 0, '판정 불가': 0, '정상': 0 };
    out.forEach(function (x) { sum[x.level]++; });
    return { results: out, summary: sum };
  }

  var api = { findTable: findTable, readFile: readFile, build: build, validate: validate, checkRow: checkRow, normUnit: normUnit, normFreq: normFreq, normRoute: normRoute,
    normName: normName, parseDose: parseDose, formOf: formOf, strengthFromName: strengthFromName, canonical: canonical, LABEL: LABEL };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.CMCore = api;
})(this);
