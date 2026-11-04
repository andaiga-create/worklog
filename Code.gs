/**
 * 작업일지 백엔드 (Google Apps Script 웹앱)
 *
 * 사전 준비
 *  1) 작업일지용 Google Sheets에서 [확장 프로그램 > Apps Script]로 열어 이 코드를 붙여넣는다.
 *  2) [프로젝트 설정 > 스크립트 속성]에 PIN 이름으로 접근코드를 추가한다. (예: 4829)
 *  3) 웹앱으로 배포: 실행 계정 = 나, 액세스 = 모든 사용자
 */

/* ===== 설정 ===== */
var SHEET_LOG = "작업일지";
var PHOTO_FOLDER_NAME = "작업일지_사진";
var MY_LOG_LIMIT = 50;

// 마스터 탭: 시트 이름 -> { 응답 키: 헤더(1행) 이름 }
var MASTER_TABS = {
  // 진행 중인 [조선소-호선-장비] 한 줄이 프로젝트 하나
  // 진행 중인 프로젝트 한 줄 = "조선소-호선-장비" (예: HHI-8123-크레인). 앱이 앞에서 두 번째 '-'까지만 나눠 인식한다
  projects:{ tab: "프로젝트", cols: { label: "프로젝트" } },
  staff:   { tab: "작업자",   cols: { name: "이름" } },                          // 자사 작업자
  vendorRows: { tab: "외주업체", cols: { vendor: "업체명", name: "작업자명" } },  // 업체명 - 작업자명 (한 줄에 한 명)
  items:   { tab: "작업항목", cols: { type: "작업구분", name: "항목명" } }
};

// 작업일지 시트 헤더 (마지막 '클라이언트ID'는 중복 제출 방지용)
var LOG_HEADERS = ["일지ID","제출시각","작성일","작성자","소속","외주업체","함께작업자",
  "조선소","호선","장비명","작업구분","작업항목","시작","종료","작업시간",
  "진행상태","세부작업","특이사항","사진링크","클라이언트ID"];

/* ===== 진입점 ===== */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!checkPin_(p.pin)) return json_({ error: "unauthorized" });
  try {
    if (p.action === "master") return json_(getMaster_());
    if (p.action === "mylogs") return json_({ logs: getMyLogs_(p.writer) });
    if (p.action === "last") return json_({ log: getLastLog_(p.yard, p.hull, p.equip) });
    return json_({ error: "unknown action" });
  } catch (err) {
    return json_({ error: String(err) });
  }
}

function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: "bad request" }); }
  if (!checkPin_(data.pin)) return json_({ error: "unauthorized" });

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (data.action === "itemops") { applyItemOps_(data.ops || []); return json_({ ok: true }); }
    return json_(saveLog_(data));
  } catch (err) {
    return json_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ===== 인증 ===== */
function checkPin_(pin) {
  var real = PropertiesService.getScriptProperties().getProperty("PIN");
  return !!real && String(pin) === String(real);
}

/* ===== 마스터 ===== */
function getMaster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {};
  Object.keys(MASTER_TABS).forEach(function (key) {
    var def = MASTER_TABS[key];
    var sh = ss.getSheetByName(def.tab);
    out[key] = [];
    if (!sh || sh.getLastRow() < 2) return;
    var values = sh.getDataRange().getDisplayValues();
    var head = values[0];
    var idx = {};
    Object.keys(def.cols).forEach(function (k) { idx[k] = head.indexOf(def.cols[k]); });
    if (key === "projects" && idx.label < 0) idx.label = 0;   // 헤더 이름이 달라도 A열을 사용
    for (var r = 1; r < values.length; r++) {
      var row = {}, empty = true;
      Object.keys(idx).forEach(function (k) {
        var v = idx[k] >= 0 ? String(values[r][idx[k]]).trim() : "";
        if (v) empty = false;
        row[k] = v;
      });
      if (!empty) out[key].push(row);
    }
  });
  // 프로젝트: "조선소-호선-장비" 문자열을 조선소/호선/장비로 분리 (장비명 안의 '-'는 그대로 유지)
  var seen = {};
  out.projects = out.projects.map(function (r) {
    var m = String(r.label || "").match(/^\s*([^-]+?)\s*-\s*([^-]+?)\s*-\s*(.+?)\s*$/);
    if (!m) return null;
    return { yard: m[1], yardName: m[1], hull: m[2], owner: "", name: m[3], code: "", serial: "" };
  }).filter(function (p) {
    if (!p) return false;
    var k = p.yard + "|" + p.hull + "|" + p.name;
    if (seen[k]) return false; seen[k] = true; return true;
  });
  // 앱이 쓰는 형태로 변환: workers(자사+외주 전원), vendors(업체 목록)
  var vendorNames = [];
  out.vendorRows.forEach(function (r) { if (r.vendor && vendorNames.indexOf(r.vendor) < 0) vendorNames.push(r.vendor); });
  out.workers = out.staff.filter(function (s) { return s.name; })
    .map(function (s) { return { name: s.name, affil: "자사", vendor: "" }; })
    .concat(out.vendorRows.filter(function (r) { return r.name; })
      .map(function (r) { return { name: r.name, affil: "외주", vendor: r.vendor }; }));
  out.vendors = vendorNames.map(function (n) { return { name: n }; });
  delete out.staff;
  delete out.vendorRows;
  return out;
}

/* ===== 작업항목 추가/수정/삭제 (A열=작업구분, B열=항목명) ===== */
function applyItemOps_(ops) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MASTER_TABS.items.tab);
  if (!sh) {
    sh = ss.insertSheet(MASTER_TABS.items.tab);
    sh.getRange(1, 1, 1, 2).setValues([["작업구분", "항목명"]]).setFontWeight("bold");
  }
  ops.forEach(function (o) {
    var last = sh.getLastRow();
    var vals = last > 1 ? sh.getRange(2, 1, last - 1, 2).getDisplayValues() : [];
    var hits = [];
    for (var i = 0; i < vals.length; i++) if (vals[i][0] === o.type && vals[i][1] === o.name) hits.push(i + 2);
    if (o.op === "add" && !hits.length) sh.appendRow([o.type, o.name]);
    else if (o.op === "del") for (var k = hits.length - 1; k >= 0; k--) sh.deleteRow(hits[k]);
    else if (o.op === "rename" && hits.length) sh.getRange(hits[0], 2).setValue(o.newName);
  });
}

/* ===== 일지 저장 ===== */
function getLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveLog_(d) {
  var req = ["date", "writer", "affil", "yard", "hull", "equip", "type", "start", "end", "status"];
  for (var i = 0; i < req.length; i++) if (!d[req[i]]) return { error: "missing " + req[i] };

  var sh = getLogSheet_();
  var last = sh.getLastRow();
  var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var cids = last > 1 ? sh.getRange(2, LOG_HEADERS.length, last - 1, 1).getValues() : [];

  // 중복 제출 방지 (오프라인 재전송 등)
  if (d.clientId) {
    for (var c = 0; c < cids.length; c++) {
      if (cids[c][0] === d.clientId) return { ok: true, id: ids[c][0], duplicate: true };
    }
  }

  // 일지ID: {조선소}-{호선}-{장비명}-{YYYYMMDD}-{작성자}-{순번}
  var prefix = [d.yard, d.hull, d.equip, String(d.date).replace(/-/g, ""), d.writer].join("-");
  var count = 0;
  for (var j = 0; j < ids.length; j++) if (String(ids[j][0]).indexOf(prefix + "-") === 0) count++;
  var id = prefix + "-" + ("0" + (count + 1)).slice(-2);

  var links = savePhotos_(id, d.photos || []);

  var row = [id, Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"), d.date, d.writer, d.affil,
    d.vendor || "", (d.workers || []).join(", "), d.yard, d.hull, d.equip, d.type, (d.items || []).join(", "),
    d.start, d.end, d.hours, d.status, d.detail || "", d.notes || "", links.join("\n"), d.clientId || ""];

  var rng = sh.getRange(last + 1, 1, 1, LOG_HEADERS.length);
  rng.setNumberFormat("@");              // 날짜/숫자 자동 변환 방지
  rng.setValues([row.map(String)]);
  return { ok: true, id: id, photos: links };
}

function savePhotos_(id, photos) {
  if (!photos.length) return [];
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
  var links = [];
  photos.slice(0, 5).forEach(function (dataUrl, i) {
    var b64 = String(dataUrl).split(",")[1];
    if (!b64) return;
    var blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/jpeg", id + "_" + (i + 1) + ".jpg");
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    links.push(file.getUrl());
  });
  return links;
}

/* ===== 조회 ===== */
function rowToLog_(r) {
  var split = function (s) { return s ? String(s).split(/\s*,\s*/) : []; };
  return {
    id: r[0], date: r[2], writer: r[3], affil: r[4], vendor: r[5], workers: split(r[6]),
    yard: r[7], hull: String(r[8]), equip: r[9], type: r[10], items: split(r[11]),
    start: r[12], end: r[13], hours: r[14], status: r[15], detail: r[16], notes: r[17]
  };
}

function readLogs_() {
  var sh = getLogSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getDisplayValues().map(rowToLog_);
}

function getMyLogs_(writer) {
  if (!writer) return [];
  return readLogs_().filter(function (l) { return l.writer === writer; }).reverse().slice(0, MY_LOG_LIMIT);
}

function getLastLog_(yard, hull, equip) {
  var logs = readLogs_();
  for (var i = logs.length - 1; i >= 0; i--) {
    var l = logs[i];
    if (l.yard === yard && l.hull === String(hull) && l.equip === equip) return l;
  }
  return null;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
