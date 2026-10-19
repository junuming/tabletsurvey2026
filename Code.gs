/**
 * 스마트CM 태블릿PC 연말 사용 보고 — 서버 코드 (Google Apps Script)
 *
 * 응답은 이 스크립트가 연결된 구글 시트의 '응답' 탭에,
 * 사진은 구글 드라이브의 '스마트CM 태블릿PC 연말보고 사진' 폴더에 현장별로 저장됩니다.
 *
 * ▼ 처음 한 번: 아래 ADMIN_PASSWORD를 바꾸고, 위쪽 함수 선택에서 setup 을 골라 [실행]하세요.
 */

// 관리자 대시보드 비밀번호 — 반드시 바꾸세요. 현장에는 알려주지 마세요.
const ADMIN_PASSWORD = '8888';

const APP_TITLE = '스마트CM 태블릿PC 사용 보고';
const SHEET_NAME = '응답';
const FOLDER_NAME = '스마트CM 태블릿PC 연말보고 사진';
const TEMP_FOLDER_NAME = '_업로드 대기';
const JSON_HEADER = '원본 데이터(JSON)';
const BASE_HEADERS = ['응답ID', '제출일시', '현장명', '단장', '기기', '점검자', '점검일', '사용 시작 월', '공종', '사용 인원', '연간 활용도 점수', '등급', '검측·도면 점수', '회의·소통 점수', '불량 항목 수', '첨부 사진 수'];

const DEFAULT_CONFIG = {
  sites: [
    { name: '대구국가산단 A7-1BL 1공구 시공단계', leader: '채규숙' },
    { name: '아이센스 마곡 R&D', leader: '백승탁' },
    { name: '인천지방검찰청 북부지청 신축공사', leader: '김운용' },
    { name: 'LH_양산사송 A-6BL 및 B-1BL 민간참여 아파트 건설공사', leader: '윤태주' },
    { name: 'LH_도청이전신도시 RH12BL 건설공사 5공구', leader: '이종실' },
    { name: '일산병원 건강 통합관리센터 건립공사', leader: '양기영' },
    { name: 'LH_울산다운2 A-1, 2, 10BL 아파트', leader: '류재구' },
    { name: '충남형 도시리브투게더 첫사업', leader: '김두한' },
    { name: '조선대 호남권 감염병원', leader: '김영철' },
    { name: 'LH_남양주왕숙 A-3,B-1,B-2BL(민간참여)', leader: '정재현' },
    { name: '충남 국제전시컨벤션센터 건립공사', leader: '최정익' },
    { name: '고려대학교 안암병원 본관 1,2병동외벽 리모델링 및 수술부 증축공사', leader: '임채영' },
    { name: 'LH_의정부우정 A-1BL 아파트 건설공사 시공단계', leader: '김태익' },
    { name: '양산부산대학교병원 스마트 진료지원교육물류센터', leader: '이병택' },
    { name: '서남병원 증축 및 리모델링공사', leader: '덩종표' },
    { name: '평택 브레인시티 공동5BL 공동주택 신축공사', leader: '조영호' },
    { name: '곤지암역세권 A1-1BL 공동주택 신축공사', leader: '최종근' },
    { name: '광명제5R구역 주택재개발정비사업', leader: '김충현' },
    { name: '화성동탄2 A76-2, A78블록 주택건설사업', leader: '이광주' },
    { name: '24-A-00부대 건설사업관리용역(A137)', leader: '김수철' },
    { name: 'LH_충남도청이전도시 RH15BL 아파트 건설공사', leader: '이장무' },
    { name: '판교 제2테크노벨리 스마트모빌리티 실증허브 조성사업', leader: '박민우' },
    { name: '도시정비본부_잠실5단지', leader: '최배근' }
  ],
  devices: ['A8', 'Active5'],
  gradeA: 70,
  gradeB: 40,
  deadline: ''
};

/* ---------------- 웹앱 진입점 ---------------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------- 처음 한 번 실행 ---------------- */
function setup() {
  if (ADMIN_PASSWORD === '여기에-관리자-비밀번호' || ADMIN_PASSWORD.length < 6) {
    throw new Error('맨 위의 ADMIN_PASSWORD를 6자 이상 새 비밀번호로 바꾼 뒤 다시 실행하세요.');
  }
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('구글 시트의 [확장 프로그램 > Apps Script]에서 만든 스크립트여야 합니다.');
  props.setProperty('SHEET_ID', ss.getId());
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(BASE_HEADERS.concat([JSON_HEADER]));
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, BASE_HEADERS.length + 1).setFontWeight('bold').setBackground('#e7ebf8');
  }
  let root;
  const rootId = props.getProperty('FOLDER_ID');
  if (rootId) { try { root = DriveApp.getFolderById(rootId); } catch (e) { root = null; } }
  if (!root) { root = DriveApp.createFolder(FOLDER_NAME); props.setProperty('FOLDER_ID', root.getId()); }
  childFolder_(root, TEMP_FOLDER_NAME);
  if (!props.getProperty('CONFIG')) props.setProperty('CONFIG', JSON.stringify(DEFAULT_CONFIG));
  Logger.log('준비 완료');
  Logger.log('응답 시트: ' + ss.getUrl());
  Logger.log('사진 폴더: ' + root.getUrl());
}

/* ---------------- 현장(누구나) ---------------- */
function apiPublic() {
  const rows = readAll_();
  const sites = {};
  rows.forEach(function (r) { if (r.info && r.info.site) sites[r.info.site] = 1; });
  return { config: getConfig_(), total: rows.length, sites: Object.keys(sites).length };
}

function apiUploadPhoto(p) {
  if (!p || typeof p.data !== 'string' || p.data.length > 14000000) throw new Error('사진 용량이 너무 큽니다.');
  const mime = /^image\/(jpeg|png|webp)$/.test(p.mime || '') ? p.mime : 'image/jpeg';
  const blob = Utilities.newBlob(Utilities.base64Decode(p.data), mime, 'upload_' + Date.now() + '.jpg');
  const file = tempFolder_().createFile(blob);
  return file.getId();
}

function apiDeletePhoto(id) {
  const file = DriveApp.getFileById(String(id));
  if (!isIn_(file, tempFolder_())) throw new Error('삭제할 수 없는 사진입니다.');
  file.setTrashed(true);
  return true;
}

function apiSubmit(resp, flat) {
  if (!resp || !resp.rid || !resp.info || !String(resp.info.site || '').trim() || !String(resp.info.inspector || '').trim()) {
    throw new Error('현장명과 점검자 이름은 필수입니다.');
  }
  const json = JSON.stringify(resp);
  if (json.length > 45000) throw new Error('응답 내용이 너무 깁니다.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const headers = ensureHeaders_(sh, (flat || []).map(function (kv) { return kv[0]; }));
    const idCol = headers.indexOf('응답ID');
    if (sh.getLastRow() > 1) {
      const ids = sh.getRange(2, idCol + 1, sh.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(resp.rid)) return { ok: true, duplicate: true };
    }
    // 사진: 대기 폴더 → 현장 폴더로 옮기고 이름 정리
    const root = rootFolder_(), temp = tempFolder_();
    const siteFolder = childFolder_(root, safeName_(resp.info.site));
    const date = String(resp.info.date || '').replace(/-/g, '');
    Object.keys(resp.photos || {}).forEach(function (k) {
      (resp.photos[k] || []).forEach(function (id, n) {
        try {
          const f = DriveApp.getFileById(id);
          if (!isIn_(f, temp)) return;
          f.setName(safeName_(resp.info.site) + '_' + k + '_' + date + '_' + ('0' + (n + 1)).slice(-2) + '.jpg');
          f.moveTo(siteFolder);
        } catch (e) { /* 이미 옮겨졌거나 없는 파일은 건너뜀 */ }
      });
    });
    const row = new Array(headers.length).fill('');
    (flat || []).forEach(function (kv) { const i = headers.indexOf(kv[0]); if (i >= 0) row[i] = kv[1] == null ? '' : kv[1]; });
    row[headers.indexOf(JSON_HEADER)] = json;
    sh.appendRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- 관리자(비밀번호) ---------------- */
function apiAdminData(pw) {
  check_(pw);
  return {
    responses: readAll_(),
    sheetUrl: SpreadsheetApp.openById(prop_('SHEET_ID')).getUrl(),
    folderUrl: rootFolder_().getUrl(),
    webUrl: ScriptApp.getService().getUrl() || ''
  };
}

function apiAdminDelete(pw, rid) {
  check_(pw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idCol = headers.indexOf('응답ID'), jsonCol = headers.indexOf(JSON_HEADER);
    const n = sh.getLastRow() - 1;
    if (n < 1) return { ok: false };
    const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][idCol]) !== String(rid)) continue;
      try {
        const r = JSON.parse(vals[i][jsonCol]);
        Object.keys(r.photos || {}).forEach(function (k) { (r.photos[k] || []).forEach(function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }); });
      } catch (e) {}
      sh.deleteRow(i + 2);
      return { ok: true };
    }
    return { ok: false };
  } finally {
    lock.releaseLock();
  }
}

function apiAdminSaveConfig(pw, cfg) {
  check_(pw);
  const c = {
    sites: (cfg.sites || []).filter(function (s) { return s && s.name; }).map(function (s) { return { name: String(s.name).slice(0, 120), leader: String(s.leader || '').slice(0, 40) }; }),
    devices: (cfg.devices || []).map(String).filter(String).slice(0, 30),
    gradeA: Number(cfg.gradeA) || 70,
    gradeB: Number(cfg.gradeB) || 0,
    deadline: /^\d{4}-\d{2}-\d{2}$/.test(cfg.deadline || '') ? cfg.deadline : ''
  };
  if (!c.devices.length) c.devices = DEFAULT_CONFIG.devices;
  PropertiesService.getScriptProperties().setProperty('CONFIG', JSON.stringify(c));
  return c;
}

/** 사진 미리보기: 작은 썸네일 (최대 12장씩) */
function apiAdminThumbs(pw, ids) {
  check_(pw);
  const out = {};
  (ids || []).slice(0, 12).forEach(function (id) {
    try {
      const f = DriveApp.getFileById(id);
      let b = f.getThumbnail();
      if (!b) b = f.getBlob();
      out[id] = 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
    } catch (e) { out[id] = ''; }
  });
  return out;
}

/** 사진 크게 보기 */
function apiAdminPhoto(pw, id) {
  check_(pw);
  const b = DriveApp.getFileById(id).getBlob();
  return 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
}

/* ---------------- 내부 함수 ---------------- */
function check_(pw) {
  if (String(pw || '') !== ADMIN_PASSWORD || ADMIN_PASSWORD === '여기에-관리자-비밀번호') {
    Utilities.sleep(800);
    throw new Error('AUTH');
  }
}
function prop_(k) {
  const v = PropertiesService.getScriptProperties().getProperty(k);
  if (!v) throw new Error('setup 함수를 먼저 실행해 주세요.');
  return v;
}
function getConfig_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('CONFIG')) || DEFAULT_CONFIG; } catch (e) { return DEFAULT_CONFIG; }
}
function sheet_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
}
function rootFolder_() { return DriveApp.getFolderById(prop_('FOLDER_ID')); }
function tempFolder_() { return childFolder_(rootFolder_(), TEMP_FOLDER_NAME); }
function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function isIn_(file, folder) {
  const it = file.getParents(), id = folder.getId();
  while (it.hasNext()) if (it.next().getId() === id) return true;
  return false;
}
function safeName_(s) {
  return String(s || '').replace(/[\\\/:*?"<>|#%]+/g, '_').replace(/\s+/g, ' ').trim().slice(0, 80) || '현장';
}
function ensureHeaders_(sh, keys) {
  let headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
  const want = BASE_HEADERS.concat(keys).concat([JSON_HEADER]);
  const missing = want.filter(function (k, i) { return k && headers.indexOf(k) < 0 && want.indexOf(k) === i; });
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#e7ebf8');
    headers = headers.concat(missing);
  }
  return headers;
}
function readAll_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const jsonCol = headers.indexOf(JSON_HEADER), idCol = headers.indexOf('응답ID');
  if (jsonCol < 0) return [];
  const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
  const out = [];
  vals.forEach(function (v) {
    try {
      const r = JSON.parse(v[jsonCol]);
      if (r && r.info) { r.rid = String(v[idCol] || r.rid); out.push(r); }
    } catch (e) {}
  });
  return out;
}
