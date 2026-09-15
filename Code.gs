/**
 * 매매 트래커 — Google Apps Script 버전
 *
 * 사용 방법:
 * 1. 새 Google Sheet를 만들고 확장 프로그램 > Apps Script를 엽니다.
 * 2. 이 파일 내용을 Code.gs에 붙여넣습니다.
 * 3. 같은 프로젝트에 HTML 파일을 하나 추가하고 이름을 정확히 "Index" 로 지정한 뒤,
 *    Index.html 파일의 내용을 붙여넣습니다.
 * 4. Apps Script 편집기 상단에서 함수 선택을 "seedData"로 바꾸고 ▶ 실행을 눌러
 *    기존에 저장돼 있던 매물 데이터를 한 번 채워 넣습니다. (권한 승인 필요)
 * 5. 배포 > 새 배포 > 웹 앱으로 배포:
 *    - 실행할 사용자: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 *    배포 후 나오는 URL이 "링크가 있는 모든 사람"에게 공유 가능한 주소입니다.
 * 6. (선택) 함수 선택을 "setupDailyTrigger"로 바꾸고 한 번 실행하면
 *    매일 오전 9시(한국시간)에 등록된 모든 단지의 KB시세를 자동으로 새로고침합니다.
 */

var LISTINGS_SHEET_NAME = 'Listings';
var META_SHEET_NAME = 'Meta';

var COLUMNS = [
  'id', 'dong', 'floorText', 'direction', 'priceRaw', 'priceMin', 'priceMax',
  'latestPriceRaw', 'latestPriceMin', 'latestPriceMax', 'confirmedDate', 'brokerCount',
  'tags', 'highlightedTags', 'link', 'memo', 'status', 'createdAt', 'updatedAt'
];

/**
 * 기본으로 내장된 단지 목록입니다. 첫 번째 항목(gurodusan)은 기존에 쓰던
 * 시트 이름(Listings/Meta)을 그대로 사용하는 기본 단지입니다.
 * areaSqm(㎡)만 넣으면 평수 표기·붙여넣기 평형 매칭 문구는 자동으로 계산됩니다.
 *
 * 사용자가 앱의 "단지 추가" 버튼으로 추가한 단지는 Complexes 시트에 저장되고,
 * "KB시세 연결" 버튼으로 지정/변경한 kbUrl도 Complexes 시트에 저장되어
 * 여기 있는 기본값을 덮어씁니다(resolvedComplexList_ 참고).
 *
 * kbComplexNo(단지기본일련번호)/kbAreaNo(면적일련번호)가 있으면 kbland.kr의
 * 실제 시세 API(BasePrcInfoNew)를 직접 호출해 정확한 평형의 시세를 가져옵니다.
 * 브라우저 개발자도구 Network 탭에서 평형을 선택했을 때 호출되는
 * https://api.kbland.kr/land-price/price/BasePrcInfoNew?단지기본일련번호=...&면적일련번호=... 요청을 보면 값을 확인할 수 있습니다.
 * 이 값이 없는 단지는 kbUrl 페이지를 통째로 가져와 텍스트에서 추출합니다
 * (단지에 평형이 하나뿐일 때만 안정적으로 동작해요).
 */
var COMPLEXES = [
  { id: 'gurodusan', name: '구로두산', areaSqm: 66, kbUrl: 'https://kbland.kr/se/c/766' },
  { id: 'hanyangmarkview', name: '한양수자인성남마크뷰', areaSqm: 56.66, kbUrl: 'https://kbland.kr/se/c/42671', kbComplexNo: 42671, kbAreaNo: 41439 },
  { id: 'byeoksanlivepark', name: '벽산라이브파크', areaSqm: 74, kbUrl: '' }
];
var DEFAULT_COMPLEX_ID = COMPLEXES[0].id;
var COMPLEXES_SHEET_NAME = 'Complexes';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('매매 트래커')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 단지 helpers ---------- */

function getComplexesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(COMPLEXES_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(COMPLEXES_SHEET_NAME);
    sh.appendRow(['id', 'name', 'areaSqm', 'kbUrl', 'kbComplexNo', 'kbAreaNo', 'createdAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readComplexRows_() {
  var sh = getComplexesSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, 7).getValues();
  return values.filter(function (r) { return r[0]; }).map(function (r) {
    return {
      id: r[0], name: r[1], areaSqm: r[2],
      kbUrl: r[3] || '', kbComplexNo: r[4] || '', kbAreaNo: r[5] || '', createdAt: r[6] || ''
    };
  });
}

/**
 * 내장 단지(COMPLEXES)와 Complexes 시트에 저장된 값을 합칩니다.
 * 시트에 같은 id가 있으면 kbUrl/kbComplexNo/kbAreaNo를 그 값으로 덮어쓰고
 * (KB시세 연결/변경), 시트에만 있는 id는 사용자가 새로 추가한 단지로 취급합니다.
 */
function resolvedComplexList_() {
  var rows = readComplexRows_();
  var overrides = {};
  rows.forEach(function (r) { overrides[r.id] = r; });

  var result = [];
  var seen = {};
  COMPLEXES.forEach(function (base) {
    var o = overrides[base.id];
    result.push({
      id: base.id,
      name: (o && o.name) || base.name,
      areaSqm: (o && o.areaSqm) || base.areaSqm,
      kbUrl: o ? o.kbUrl : (base.kbUrl || ''),
      kbComplexNo: o ? o.kbComplexNo : (base.kbComplexNo || ''),
      kbAreaNo: o ? o.kbAreaNo : (base.kbAreaNo || '')
    });
    seen[base.id] = true;
  });
  rows.forEach(function (r) {
    if (seen[r.id]) return;
    result.push(r);
  });
  return result;
}

function complexById_(complexId) {
  var list = resolvedComplexList_();
  var found = null;
  list.forEach(function (c) { if (c.id === complexId) found = c; });
  return found || list[0];
}

function pyeongOf_(areaSqm) {
  return Math.round(Number(areaSqm) / 3.305785);
}

function areaLabelOf_(areaSqm) {
  return '매매 · 전용 ' + areaSqm + '㎡(~' + pyeongOf_(areaSqm) + '평)';
}

function complexViewOf_(c) {
  return {
    id: c.id,
    name: c.name,
    areaSqm: c.areaSqm,
    areaLabel: areaLabelOf_(c.areaSqm),
    areaMatch: String(c.areaSqm),
    exampleArea: c.areaSqm + '㎡',
    hasKbLink: !!c.kbUrl
  };
}

function slugifyComplexId_(name) {
  var s = String(name).toLowerCase().replace(/[^a-z0-9가-힣]+/g, '-').replace(/^-+|-+$/g, '');
  return (s || 'complex').slice(0, 40);
}

function listingsSheetName_(complexId) {
  return complexId === DEFAULT_COMPLEX_ID ? LISTINGS_SHEET_NAME : (LISTINGS_SHEET_NAME + '_' + complexId);
}

function metaSheetName_(complexId) {
  return complexId === DEFAULT_COMPLEX_ID ? META_SHEET_NAME : (META_SHEET_NAME + '_' + complexId);
}

/* ---------- sheet helpers ---------- */

function getListingsSheet_(complexId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = listingsSheetName_(complexId);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(COLUMNS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getMetaSheet_(complexId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = metaSheetName_(complexId);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(['key', 'value']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function numOrNull_(v) {
  return (v === '' || v === null || v === undefined) ? null : Number(v);
}

function parseJsonArray_(v) {
  try {
    var a = JSON.parse(v || '[]');
    return Array.isArray(a) ? a : [];
  } catch (e) {
    return [];
  }
}

function rowToObject_(row) {
  return {
    id: row[0],
    dong: row[1] || '',
    floorText: row[2] || '',
    direction: row[3] || '',
    priceRaw: row[4] || '',
    priceMin: numOrNull_(row[5]),
    priceMax: numOrNull_(row[6]),
    latestPriceRaw: row[7] || '',
    latestPriceMin: numOrNull_(row[8]),
    latestPriceMax: numOrNull_(row[9]),
    confirmedDate: row[10] || '',
    brokerCount: numOrNull_(row[11]),
    tags: parseJsonArray_(row[12]),
    highlightedTags: parseJsonArray_(row[13]),
    link: row[14] || '',
    memo: row[15] || '',
    status: row[16] || 'new',
    createdAt: row[17] || '',
    updatedAt: row[18] || ''
  };
}

function objectToRow_(obj) {
  return [
    obj.id, obj.dong || '', obj.floorText || '', obj.direction || '', obj.priceRaw || '',
    obj.priceMin, obj.priceMax, obj.latestPriceRaw || '', obj.latestPriceMin, obj.latestPriceMax,
    obj.confirmedDate || '', obj.brokerCount,
    JSON.stringify(obj.tags || []), JSON.stringify(obj.highlightedTags || []),
    obj.link || '', obj.memo || '', obj.status || 'new',
    obj.createdAt || '', obj.updatedAt || ''
  ];
}

function readAllListings_(complexId) {
  var sh = getListingsSheet_(complexId);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  return values.filter(function (r) { return r[0]; }).map(rowToObject_);
}

function findRowIndexById_(sh, id) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  return -1;
}

function keyOf_(dong, floorText, priceRaw) {
  return [dong, floorText, priceRaw].join('__');
}

function readKbPrice_(complexId) {
  var sh = getMetaSheet_(complexId);
  var lastRow = sh.getLastRow();
  var obj = {};
  if (lastRow >= 2) {
    var values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    values.forEach(function (r) { if (r[0]) obj[r[0]] = r[1]; });
  }
  if (!obj.generalPrice) return null;
  return {
    generalPrice: obj.generalPrice || '',
    generalDate: obj.generalDate || '',
    dealPrice: obj.dealPrice || '',
    dealDate: obj.dealDate || '',
    dealFloor: obj.dealFloor || '',
    fetchedAt: obj.fetchedAt || ''
  };
}

function writeKbPrice_(complexId, kb) {
  var sh = getMetaSheet_(complexId);
  sh.clearContents();
  sh.appendRow(['key', 'value']);
  Object.keys(kb).forEach(function (k) { sh.appendRow([k, kb[k]]); });
}

/* ---------- client-callable functions ---------- */

function gsGetComplexes() {
  return resolvedComplexList_().map(complexViewOf_);
}

/**
 * 사용자가 "단지 추가" 버튼으로 새 단지를 등록합니다. KB시세 링크는 비워둔
 * 채로 생성되며, 나중에 gsSetComplexKbLink로 연결/변경할 수 있습니다.
 */
function gsAddComplex(name, areaSqm) {
  name = (name || '').trim();
  areaSqm = Number(areaSqm);
  if (!name) throw new Error('단지 이름을 입력해주세요.');
  if (!areaSqm || isNaN(areaSqm) || areaSqm <= 0) throw new Error('평수(㎡)를 올바르게 입력해주세요.');

  var existingIds = {};
  resolvedComplexList_().forEach(function (c) { existingIds[c.id] = true; });

  var baseId = slugifyComplexId_(name);
  var uniqueId = baseId;
  var suffix = 2;
  while (existingIds[uniqueId]) { uniqueId = baseId + '-' + suffix; suffix++; }

  var sh = getComplexesSheet_();
  sh.appendRow([uniqueId, name, areaSqm, '', '', '', new Date().toISOString()]);
  return { complexes: gsGetComplexes(), newId: uniqueId };
}

/**
 * 단지의 KB부동산 시세 페이지 링크를 연결하거나 바꿉니다. 내장 단지(COMPLEXES)의
 * 값을 Complexes 시트에 override로 저장하는 방식이라, 기본 단지든 사용자가
 * 추가한 단지든 동일하게 동작합니다. API 방식(kbComplexNo/kbAreaNo)이 아니라
 * 페이지 전체를 가져와 추출하는 방식으로 전환됩니다.
 */
function gsSetComplexKbLink(complexId, kbUrl) {
  kbUrl = (kbUrl || '').trim();
  if (!kbUrl) throw new Error('링크를 입력해주세요.');
  if (!/^https?:\/\//i.test(kbUrl)) kbUrl = 'https://' + kbUrl;

  var sh = getComplexesSheet_();
  var lastRow = sh.getLastRow();
  var rowIdx = -1;
  if (lastRow >= 2) {
    var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === complexId) { rowIdx = i + 2; break; }
    }
  }

  if (rowIdx === -1) {
    var base = null;
    COMPLEXES.forEach(function (c) { if (c.id === complexId) base = c; });
    if (!base) throw new Error('알 수 없는 단지예요.');
    sh.appendRow([complexId, base.name, base.areaSqm, kbUrl, '', '', new Date().toISOString()]);
  } else {
    sh.getRange(rowIdx, 4, 1, 3).setValues([[kbUrl, '', '']]);
  }
  return gsGetComplexes();
}

function gsGetData(complexId) {
  return { listings: readAllListings_(complexId), kbPrice: readKbPrice_(complexId) };
}

function gsAddListing(complexId, entry) {
  var sh = getListingsSheet_(complexId);
  var existing = readAllListings_(complexId);
  var key = keyOf_(entry.dong, entry.floorText, entry.priceRaw);
  var dup = existing.some(function (it) {
    return keyOf_(it.dong, it.floorText, it.priceRaw) === key;
  });
  if (dup) return { added: false };

  var now = new Date().toISOString();
  var obj = {
    id: Utilities.getUuid(),
    dong: entry.dong, floorText: entry.floorText || '', direction: entry.direction || '',
    priceRaw: entry.priceRaw, priceMin: entry.priceMin, priceMax: entry.priceMax,
    latestPriceRaw: entry.latestPriceRaw || '', latestPriceMin: entry.latestPriceMin, latestPriceMax: entry.latestPriceMax,
    confirmedDate: entry.confirmedDate || '', brokerCount: entry.brokerCount,
    tags: entry.tags || [], highlightedTags: [], link: entry.link || '', memo: entry.memo || '',
    status: 'new', createdAt: now, updatedAt: now
  };
  sh.appendRow(objectToRow_(obj));
  return { added: true, item: obj };
}

function gsAddListingsBulk(complexId, entries) {
  var sh = getListingsSheet_(complexId);
  var existing = readAllListings_(complexId);
  var existingKeys = {};
  existing.forEach(function (it) { existingKeys[keyOf_(it.dong, it.floorText, it.priceRaw)] = true; });

  var added = 0, skipped = 0, newOnes = [];
  entries.forEach(function (entry) {
    var key = keyOf_(entry.dong, entry.floorText, entry.priceRaw);
    if (existingKeys[key]) { skipped++; return; }
    existingKeys[key] = true;
    var now = new Date().toISOString();
    var obj = {
      id: Utilities.getUuid(),
      dong: entry.dong, floorText: entry.floorText || '', direction: entry.direction || '',
      priceRaw: entry.priceRaw, priceMin: entry.priceMin, priceMax: entry.priceMax,
      latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null,
      confirmedDate: entry.confirmedDate || '', brokerCount: entry.brokerCount,
      tags: [], highlightedTags: [], link: '', memo: entry.memo || '', status: 'new',
      createdAt: now, updatedAt: now
    };
    sh.appendRow(objectToRow_(obj));
    added++; newOnes.push(entry);
  });
  return { added: added, skipped: skipped, newOnes: newOnes };
}

function gsUpdateListing(complexId, id, patch) {
  var sh = getListingsSheet_(complexId);
  var rowIdx = findRowIndexById_(sh, id);
  if (rowIdx === -1) return { ok: false };
  var row = sh.getRange(rowIdx, 1, 1, COLUMNS.length).getValues()[0];
  var obj = rowToObject_(row);
  Object.keys(patch).forEach(function (k) { obj[k] = patch[k]; });
  obj.updatedAt = new Date().toISOString();
  sh.getRange(rowIdx, 1, 1, COLUMNS.length).setValues([objectToRow_(obj)]);
  return { ok: true, item: obj };
}

/* ---------- KB시세 ---------- */

/**
 * HTML을 사람이 페이지를 보고 복사한 것과 비슷한 일반 텍스트로 변환합니다.
 * (debugKbPriceBlocks에서 문맥을 읽기 쉽게 출력하기 위한 용도)
 */
function htmlToPlainText_(html) {
  var noScripts = html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ');
  var withBreaks = noScripts
    .replace(/<(br|li|tr|p|div|h[1-6])[^>]*>/gi, '\n')
    .replace(/<\/(li|tr|p|div|h[1-6])>/gi, '\n');
  var stripped = withBreaks.replace(/<[^>]+>/g, ' ');
  var decoded = stripped
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
  var lines = decoded.split('\n')
    .map(function (l) { return l.replace(/[ \t]+/g, ' ').trim(); })
    .filter(Boolean);
  return lines.join('\n');
}

/**
 * 라벨(예: 'KB시세 일반가') 바로 뒤에 오는 두 개의 span 값을 읽어옵니다.
 * 한 페이지에 여러 평형의 시세가 나오는 단지가 있어서, areaHint(예: '56.66')를
 * 넘기면 그 문자열이 앞쪽 근처(최대 3000자 이내)에 등장하는 라벨을 우선적으로
 * 선택합니다. 못 찾으면 페이지에서 가장 먼저 나오는 라벨 값으로 대체합니다.
 */
function extractLabelValue_(html, label, areaHint) {
  function parseAt(idx) {
    var after = html.slice(idx, idx + 1200);
    var spans = after.match(/<span[^>]*>([^<]+)<\/span>/g);
    if (!spans || spans.length < 2) return null;
    function textOf(tag) {
      var m = tag.match(/>([^<]+)</);
      return m ? m[1] : '';
    }
    return { first: textOf(spans[0]), second: textOf(spans[1]) };
  }

  var searchFrom = 0;
  var firstMatch = null;
  while (true) {
    var idx = html.indexOf('>' + label + '<', searchFrom);
    if (idx === -1) break;
    var parsed = parseAt(idx);
    if (parsed) {
      if (!firstMatch) firstMatch = parsed;
      if (areaHint) {
        var windowStart = Math.max(0, idx - 3000);
        if (html.slice(windowStart, idx).indexOf(areaHint) !== -1) {
          return parsed;
        }
      }
    }
    searchFrom = idx + label.length + 2;
  }
  return firstMatch;
}

function normalizeKbDate_(s) {
  if (!s) return '';
  var parts = s.split(/[.\-]/).filter(Boolean);
  if (parts.length < 3) return s;
  var y = parts[0].length === 2 ? ('20' + parts[0]) : parts[0];
  var m = ('' + parts[1]).padStart(2, '0');
  var d = ('' + parts[2]).padStart(2, '0');
  return y + '.' + m + '.' + d;
}

function withThousandsCommas_(n) {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatManwonPrice_(v) {
  if (v === null || v === undefined || v === '' || isNaN(v)) return '';
  v = Math.round(Number(v));
  var eok = Math.floor(v / 10000);
  var man = v % 10000;
  if (eok === 0) return withThousandsCommas_(man) + '만';
  if (man === 0) return eok + '억';
  return eok + '억 ' + withThousandsCommas_(man);
}

function formatYyyymmdd_(s) {
  if (!s) return '';
  s = String(s);
  if (s.length !== 8) return s;
  return s.slice(0, 4) + '.' + s.slice(4, 6) + '.' + s.slice(6, 8);
}

/**
 * kbland.kr의 실제 시세 API(BasePrcInfoNew)를 단지기본일련번호+면적일련번호로
 * 직접 호출해 정확한 평형의 매매 시세를 가져옵니다. HTML을 통째로 가져와
 * 텍스트로 추출하는 방식과 달리, 페이지에 여러 평형이 섞여 있어도 정확합니다.
 */
function refreshKbPriceFromApi_(complex) {
  var url = 'https://api.kbland.kr/land-price/price/BasePrcInfoNew'
    + '?' + encodeURIComponent('단지기본일련번호') + '=' + complex.kbComplexNo
    + '&' + encodeURIComponent('면적일련번호') + '=' + complex.kbAreaNo;
  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36'
    }
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('KB시세 API 응답 오류 (코드 ' + res.getResponseCode() + ')');
  }
  var json = JSON.parse(res.getContentText());
  var body = json && json.dataBody && json.dataBody.data;
  var sise = body && body['시세'] && body['시세'][0];
  if (!sise) {
    throw new Error('KB시세 API 응답에서 시세 정보를 찾지 못했어요.');
  }

  var kb = {
    generalPrice: formatManwonPrice_(sise['매매일반거래가']),
    generalDate: formatYyyymmdd_(sise['시세기준년월일']),
    dealPrice: formatManwonPrice_(sise['매매거래금액']),
    dealDate: formatYyyymmdd_(sise['매매계약종료년월일']),
    dealFloor: sise['매매해당층수'] ? (sise['매매해당층수'] + '층') : '',
    fetchedAt: new Date().toISOString()
  };
  writeKbPrice_(complex.id, kb);
  return kb;
}

/**
 * 단지 상세 페이지를 통째로 가져와 텍스트에서 KB시세를 추출합니다.
 * 단지에 평형이 하나뿐인 경우에만 안정적으로 동작합니다(여러 평형이 섞여
 * 있으면 어느 평형 값을 가져올지 페이지 기본값에 의존하게 됩니다).
 */
function refreshKbPriceFromHtml_(complex) {
  var res = UrlFetchApp.fetch(complex.kbUrl, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('kbland.kr 응답 오류 (코드 ' + res.getResponseCode() + ')');
  }
  var html = res.getContentText();
  var areaMatch = String(complex.areaSqm);

  var general = extractLabelValue_(html, 'KB시세 일반가', areaMatch);
  var deal = extractLabelValue_(html, '최근 실거래가', areaMatch);
  if (!general || !deal) {
    throw new Error('시세 페이지 구조를 인식하지 못했어요. kbland.kr 화면 구성이 바뀌었을 수 있어요.');
  }

  var dealParts = deal.second.split('/');
  var kb = {
    generalPrice: general.first,
    generalDate: normalizeKbDate_(general.second),
    dealPrice: deal.first,
    dealDate: normalizeKbDate_(dealParts[0]),
    dealFloor: dealParts[1] || '',
    fetchedAt: new Date().toISOString()
  };
  writeKbPrice_(complex.id, kb);
  return kb;
}

function gsRefreshKbPrice(complexId) {
  var complex = complexById_(complexId);
  if (!complex.kbUrl) {
    throw new Error('이 단지는 KB시세 링크가 연결되어 있지 않아요. 먼저 링크를 연결해주세요.');
  }
  if (complex.kbComplexNo && complex.kbAreaNo) {
    return refreshKbPriceFromApi_(complex);
  }
  return refreshKbPriceFromHtml_(complex);
}

/* ---------- 일회성 유틸리티 (Apps Script 편집기에서 직접 실행) ---------- */

/**
 * KB시세 오매칭 문제를 진단하기 위한 함수입니다(2차: 평형/거래유형 원인 좁히기).
 * 1) 페이지에 구조화된 데이터(JSON)가 통째로 심어져 있는지 확인하고
 * 2) '56.66', '102.2', '매매', '전세' 같은 키워드가 어디에 몇 번 등장하는지,
 * 3) 'KB시세 일반가' 라벨 각각의 앞쪽 문맥(최대 2000자)에 그 키워드들이
 *    포함돼 있는지를 실행 로그에 출력합니다.
 *
 * 사용법: Apps Script 편집기 상단 함수 선택을 "debugKbPriceBlocks"로 바꾸고
 * ▶ 실행 → 실행 로그 전체를 복사해 공유해주세요.
 */
function debugKbPriceBlocks() {
  var complex = complexById_('hanyangmarkview');
  var res = UrlFetchApp.fetch(complex.kbUrl, {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36'
    }
  });
  Logger.log('응답 코드: ' + res.getResponseCode());
  var html = res.getContentText();
  Logger.log('HTML 길이: ' + html.length);

  Logger.log('----- 구조화 데이터(JSON) 마커 위치 -----');
  ['__NEXT_DATA__', '__INITIAL_STATE__', '__PRELOADED_STATE__', 'application/ld+json', 'application/json'].forEach(function (marker) {
    Logger.log('"' + marker + '" 위치: ' + html.indexOf(marker));
  });

  Logger.log('----- 평형/거래유형 키워드 위치 -----');
  ['56.66', '102.2', '40.95', '84.89', '매매', '전세', '월세'].forEach(function (needle) {
    var positions = [];
    var i = 0;
    while (positions.length < 6) {
      var found = html.indexOf(needle, i);
      if (found === -1) break;
      positions.push(found);
      i = found + needle.length;
    }
    Logger.log('"' + needle + '" 위치: ' + JSON.stringify(positions));
  });

  var label = 'KB시세 일반가';
  var idx = 0, count = 0;
  while (count < 5) {
    idx = html.indexOf('>' + label + '<', idx);
    if (idx === -1) break;
    count++;
    var wideStart = Math.max(0, idx - 6000);
    var beforeWide = htmlToPlainText_(html.slice(wideStart, idx));
    Logger.log('===== 매치 #' + count + ' (offset ' + idx + ') 앞쪽 문맥(라벨 바로 앞 2000자) =====');
    Logger.log(beforeWide.slice(-2000));
    idx += label.length;
  }
  if (count === 0) {
    Logger.log('라벨을 하나도 찾지 못했어요.');
  } else {
    Logger.log('총 ' + count + '개 매치 발견');
  }
}

/**
 * 등록된 모든 단지의 KB시세를 순서대로 새로고침합니다.
 * 한 단지에서 오류가 나도 나머지 단지는 계속 새로고침을 시도합니다.
 */
function refreshAllKbPrices() {
  resolvedComplexList_().forEach(function (c) {
    if (!c.kbUrl) return;
    try {
      gsRefreshKbPrice(c.id);
    } catch (e) {
      Logger.log(c.name + ' KB시세 새로고침 실패: ' + e.message);
    }
  });
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'gsRefreshKbPrice' || fn === 'refreshAllKbPrices') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshAllKbPrices')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .inTimezone('Asia/Seoul')
    .create();
}

/**
 * Claude 아티팩트에 저장돼 있던 기존 매물 데이터를 그대로 옮겨 심습니다.
 * Apps Script 편집기에서 이 함수를 한 번만 실행하세요 (이미 데이터가 있으면 덮어씁니다).
 * 구로두산(기본 단지) 데이터만 채웁니다.
 */
function seedData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName(listingsSheetName_(DEFAULT_COMPLEX_ID));
  if (old) ss.deleteSheet(old);
  var oldMeta = ss.getSheetByName(metaSheetName_(DEFAULT_COMPLEX_ID));
  if (oldMeta) ss.deleteSheet(oldMeta);

  var sh = getListingsSheet_(DEFAULT_COMPLEX_ID);
  var seedListings = [
    { id: '1j1ob6pwukk2iec943k9', dong: '103동', floorText: '25/25층', direction: '북동향', priceRaw: '7억 3,000', priceMin: 73000, priceMax: 73000, latestPriceRaw: '7억 2,000', latestPriceMin: 72000, latestPriceMax: 72000, confirmedDate: '2026.08.29', brokerCount: 10, tags: ['옛수리', '12월입주협의'], highlightedTags: [], link: 'https://naver.me/FetV0tXp', memo: '옛 수리', status: 'hold', createdAt: '2026-09-14T16:22:41.779Z', updatedAt: '2026-09-14T16:56:18.560Z' },
    { id: '3hxnyidi5bri7ajkilxp', dong: '101동', floorText: '19/25층', direction: '남동향', priceRaw: '7억 6,000', priceMin: 76000, priceMax: 76000, latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null, confirmedDate: '2026.08.31', brokerCount: 4, tags: ['올수리', '12월입주협의'], highlightedTags: [], link: 'https://naver.me/F5spVFdz', memo: '샷시 , 중문, 욕실 포함 최근 올수리', status: 'hold', createdAt: '2026-09-14T16:11:59.953Z', updatedAt: '2026-09-14T17:07:15.251Z' },
    { id: '5q7z8o6wxt57pqaueyf8', dong: '110동', floorText: '8/23층', direction: '북동향', priceRaw: '7억 3,000', priceMin: 73000, priceMax: 73000, latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null, confirmedDate: '2026.09.03', brokerCount: 8, tags: ['깨끗', '0912방문', '가격협의가능'], highlightedTags: ['0912방문', '가격협의가능'], link: 'https://naver.me/F7F3Mye1', memo: '앞뒤트여 밝고 깨끗함 입주협의', status: 'interested', createdAt: '2026-09-14T16:14:24.212Z', updatedAt: '2026-09-14T17:40:44.933Z' },
    { id: '6gp11ukwk3210xu3yejs', dong: '101동', floorText: '2/25층', direction: '남서향', priceRaw: '7억 3,000', priceMin: 73000, priceMax: 73000, latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null, confirmedDate: '2026.09.12', brokerCount: 4, tags: ['특올수리', '12월입주협의'], highlightedTags: [], link: 'https://naver.me/GuCVRCRe', memo: '특 올수리', status: 'hold', createdAt: '2026-09-14T16:23:33.322Z', updatedAt: '2026-09-14T16:55:03.215Z' },
    { id: 'duwov259q6jy7me5rn3t', dong: '109동', floorText: '20/25층', direction: '남서향', priceRaw: '7억 7,000', priceMin: 77000, priceMax: 77000, latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null, confirmedDate: '2026.09.11', brokerCount: 1, tags: ['특올수리'], highlightedTags: [], link: 'https://naver.me/xrzV4SJZ', memo: '샷시 전체 포함 특 올수리', status: 'hold', createdAt: '2026-09-14T16:18:30.867Z', updatedAt: '2026-09-14T16:58:15.787Z' },
    { id: 'e2lt5ahwaj29sxxtscj7', dong: '103동', floorText: '3/25층', direction: '남동향', priceRaw: '7억 3,000', priceMin: 73000, priceMax: 73000, latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null, confirmedDate: '2026.09.09', brokerCount: 2, tags: ['일부수리', '깨끗', '가격협의가능'], highlightedTags: [], link: 'https://naver.me/G7Ni8VyF', memo: '화장실, 바닥 등 일부수리, 가격 협의 가능', status: 'hold', createdAt: '2026-09-14T16:16:24.494Z', updatedAt: '2026-09-14T17:14:18.079Z' },
    { id: 'iac5hutz6nahf72l93rj', dong: '101동', floorText: '7/25층', direction: '남동향', priceRaw: '7억 8,000', priceMin: 78000, priceMax: 78000, latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null, confirmedDate: '2026.09.08', brokerCount: 2, tags: ['특올수리', '12월입주협의'], highlightedTags: [], link: 'https://naver.me/FwGh6jLQ', memo: '특 올수리, 12월 입주협의', status: 'hold', createdAt: '2026-09-14T16:21:14.970Z', updatedAt: '2026-09-14T17:42:21.775Z' },
    { id: 'jkzzakzniip4igcrryen', dong: '101동', floorText: '고/25층', direction: '남서향', priceRaw: '7억 7,000', priceMin: 77000, priceMax: 77000, latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null, confirmedDate: '2026.09.08', brokerCount: 1, tags: ['즉시입주가능'], highlightedTags: [], link: 'https://naver.me/5zUmpJ7t', memo: '즉시 입주가능 전망좋고 일조량 좋음', status: 'hold', createdAt: '2026-09-14T16:00:27.795Z', updatedAt: '2026-09-14T17:08:29.098Z' },
    { id: 'yv4nczlujqhtxs7fgh4i', dong: '109동', floorText: '10/25층', direction: '남서향', priceRaw: '7억 4,000', priceMin: 74000, priceMax: 74000, latestPriceRaw: '', latestPriceMin: null, latestPriceMax: null, confirmedDate: '2026.09.08', brokerCount: 7, tags: ['옛수리', '1월입주협의'], highlightedTags: [], link: 'https://naver.me/GFsUwCME', memo: '옛 수리, 1월 입주협의', status: 'hold', createdAt: '2026-09-14T16:20:01.295Z', updatedAt: '2026-09-14T16:57:49.689Z' }
  ];
  seedListings.forEach(function (obj) { sh.appendRow(objectToRow_(obj)); });

  writeKbPrice_(DEFAULT_COMPLEX_ID, {
    generalPrice: '6억 8,000',
    generalDate: '2026.09.11',
    dealPrice: '7억 1,500',
    dealDate: '2026.09.07',
    dealFloor: '13층',
    fetchedAt: '2026-09-15T12:00:00.000Z'
  });

  Logger.log('시드 완료: 매물 ' + seedListings.length + '건 + KB시세');
}
