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
 * 페이지 우측 상단의 단지 선택 버튼에 표시되는 단지 목록입니다.
 * 단지를 추가하려면 이 배열에 항목을 추가하세요. 첫 번째 항목(gurodusan)은
 * 기존에 쓰던 시트 이름(Listings/Meta)을 그대로 사용하는 기본 단지입니다.
 */
var COMPLEXES = [
  { id: 'gurodusan', name: '구로두산', areaLabel: '매매 · 전용 66㎡(전용44.64)', areaMatch: '66', exampleArea: '66㎡ (전용44.64)', kbUrl: 'https://kbland.kr/se/c/766' },
  { id: 'hanyangmarkview', name: '한양수자인성남마크뷰', areaLabel: '매매 · 전용 56.66㎡(전용40.95)', areaMatch: '56.66', exampleArea: '56.66㎡ (전용40.95)', kbUrl: 'https://kbland.kr/se/c/42671' },
  { id: 'byeoksanlivepark', name: '벽산라이브파크', areaLabel: '매매 · 전용 102.49㎡(전용84.89)', areaMatch: '102.49', exampleArea: '102.49㎡ (전용84.89)', kbUrl: 'https://kbland.kr/se/c/422' }
];
var DEFAULT_COMPLEX_ID = COMPLEXES[0].id;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('매매 트래커')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 단지 helpers ---------- */

function complexById_(complexId) {
  var found = null;
  COMPLEXES.forEach(function (c) { if (c.id === complexId) found = c; });
  return found || COMPLEXES[0];
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
  return COMPLEXES.map(function (c) {
    return { id: c.id, name: c.name, areaLabel: c.areaLabel, areaMatch: c.areaMatch, exampleArea: c.exampleArea };
  });
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
      tags: [], highlightedTags: [], link: '', memo: '', status: 'new',
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

/* ---------- 네이버 부동산 링크로 매물 정보 가져오기 ---------- */

/**
 * HTML을 사람이 페이지를 보고 복사-붙여넣기 했을 때와 비슷한 일반 텍스트로 변환합니다.
 * (붙여넣기 탭에서 쓰는 것과 같은 방식으로 파싱하기 위한 전처리)
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

function extractDongFromText_(text) {
  var m = text.match(/(\d{1,4}동)(?!\d)/);
  return m ? m[1] : '';
}

function extractFloorTextFromText_(text) {
  var m = text.match(/(\d{1,3})\s*\/\s*(\d{1,3})\s*층/);
  if (m) return m[1] + '/' + m[2] + '층';
  m = text.match(/(\d{1,3})\s*층\s*\/\s*(\d{1,3})\s*층/);
  if (m) return m[1] + '/' + m[2] + '층';
  m = text.match(/(고)\s*\/\s*(\d{1,3})\s*층/);
  if (m) return m[1] + '/' + m[2] + '층';
  return '';
}

function extractDirectionFromText_(text) {
  var m = text.match(/(남동향|남서향|북동향|북서향|남향|북향|동향|서향)/);
  return m ? m[1] : '';
}

function extractPriceRawFromText_(text) {
  var m = text.match(/매매\s*([\d][\d,.\s~억만]*)/);
  if (!m) return '';
  return m[1].trim().replace(/\s+/g, ' ').replace(/[~,.\s]+$/, '');
}

function extractConfirmedDateFromText_(text) {
  var m = text.match(/확인[^\d\n]{0,8}(\d{4}[.\-]\d{1,2}[.\-]\d{1,2})/);
  if (!m) m = text.match(/(\d{4}[.\-]\d{1,2}[.\-]\d{1,2})[^\d\n]{0,4}확인/);
  return m ? normalizeKbDate_(m[1]) : '';
}

function extractMemoFromText_(text) {
  var m = text.match(/(?:특징|매물\s*특징|중개사\s*코멘트|상세\s*설명)[:\s]*([^\n]{5,80})/);
  return m ? m[1].trim() : '';
}

/**
 * 네이버 부동산(naver.me 단축링크, land.naver.com 등) 매물 상세 링크에서
 * 동/층/방향/가격/확인일자/메모를 최대한 자동으로 추출합니다.
 * 사이트 화면 구성이 자주 바뀌기 때문에 항목이 비어 있을 수 있고,
 * 그런 경우 "직접 입력" 탭에서 나머지를 채워 넣어야 합니다.
 */
function gsFetchListingFromLink(rawUrl) {
  var url = (rawUrl || '').trim();
  if (!url) throw new Error('링크를 입력해주세요.');
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
  if (!/naver\.(me|com)/i.test(url)) {
    throw new Error('네이버 부동산(naver.me, land.naver.com) 링크만 지원해요.');
  }

  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36'
    }
  });
  var code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('링크 페이지 응답 오류 (코드 ' + code + ')');
  }
  var text = htmlToPlainText_(res.getContentText());

  var entry = {
    dong: extractDongFromText_(text),
    floorText: extractFloorTextFromText_(text),
    direction: extractDirectionFromText_(text),
    priceRaw: extractPriceRawFromText_(text),
    confirmedDate: extractConfirmedDateFromText_(text),
    memo: extractMemoFromText_(text),
    link: url
  };

  var foundAny = entry.dong || entry.priceRaw || entry.floorText || entry.direction || entry.confirmedDate;
  if (!foundAny) {
    throw new Error('페이지에서 매물 정보를 인식하지 못했어요. 링크를 다시 확인하거나 직접 입력해 주세요.');
  }
  return entry;
}

/* ---------- KB시세 ---------- */

function extractLabelValue_(html, label) {
  var idx = html.indexOf('>' + label + '<');
  if (idx === -1) return null;
  var after = html.slice(idx, idx + 1200);
  var spans = after.match(/<span[^>]*>([^<]+)<\/span>/g);
  if (!spans || spans.length < 2) return null;
  function textOf(tag) {
    var m = tag.match(/>([^<]+)</);
    return m ? m[1] : '';
  }
  return { first: textOf(spans[0]), second: textOf(spans[1]) };
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

function gsRefreshKbPrice(complexId) {
  var complex = complexById_(complexId);
  var res = UrlFetchApp.fetch(complex.kbUrl, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('kbland.kr 응답 오류 (코드 ' + res.getResponseCode() + ')');
  }
  var html = res.getContentText();

  var general = extractLabelValue_(html, 'KB시세 일반가');
  var deal = extractLabelValue_(html, '최근 실거래가');
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

/* ---------- 일회성 유틸리티 (Apps Script 편집기에서 직접 실행) ---------- */

/**
 * 등록된 모든 단지의 KB시세를 순서대로 새로고침합니다.
 * 한 단지에서 오류가 나도 나머지 단지는 계속 새로고침을 시도합니다.
 */
function refreshAllKbPrices() {
  COMPLEXES.forEach(function (c) {
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
