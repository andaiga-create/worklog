/* =====================================================================
   검사신청: 신청 저장 → 구글 문서 틀에 채워 PDF 생성 → 지정 메일로 발송
   - PDF는 드라이브에 저장하지 않고 메일 첨부로만 보낸다 (재발송이 필요하면 시트의 요청원본으로 다시 만든다)
   - 틀: 드라이브의 '검사신청서_틀_성능검사', '검사신청서_틀_출하검사' (구글 문서)
   - 이력: 시트 '검사신청' 탭 (한 신청 = 한 줄 = PDF 한 장)
   - 받는 사람: 시트 '검사메일' 탭 A열 (한 줄에 주소 하나)
   - 부품 목록: 시트 '검사항목' 탭 (검사구분 / 부품명), 앱에서 추가·수정·삭제
   - 처음 한 번 setupInsp 를 편집기에서 ▶ 실행한다 (탭 생성, 권한 승인, 재발송 트리거)
   ===================================================================== */
var INSP_TPL = { perf: "1NeZSczEOhVMApItX_t2M1sLImHVyE4FmoAKhbVWxOIo", ship: "1sOXrQsurPM8dwiLIt4HfYjS0NRdsGaXeyhNNXuDZ1Gg" };
var INSP_LABEL = { perf: "성능검사", ship: "출하검사" };
var INSP_PREFIX = { perf: "PT", ship: "SH" };
var INSP_TAB = "검사신청";
var INSP_MAIL_TAB = "검사메일";
var INSP_ITEM_TAB = "검사항목";
var INSP_FOLDER = "검사신청서_PDF";
var INSP_MAIL_SEED = "erai77@techflower.co.kr";
var INSP_HEADERS = ["관리번호", "신청시각", "신청일자", "작성자", "검사구분", "조선소", "호선", "장비명", "품명",
  "검사요구일시", "부품내역", "확인사항", "PDF링크", "발송상태", "클라이언트ID", "요청원본"];
var INSP_TZ = "Asia/Seoul";

/* ---------- 설치 ---------- */
function setupInsp() {
  getInspSheet_();
  getInspRecipients_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(INSP_ITEM_TAB)) {
    var sh = ss.insertSheet(INSP_ITEM_TAB);
    sh.getRange(1, 1, 1, 2).setValues([["검사구분", "부품명"]]).setFontWeight("bold");
  }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "retryInspMail") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("retryInspMail").timeBased().everyHours(1).create();
  // 권한 승인을 한 번에 받기 위해 문서/드라이브/메일 기능을 한 번씩 건드린다
  DocumentApp.openById(INSP_TPL.perf).getName();
  DocumentApp.openById(INSP_TPL.ship).getName();
  MailApp.getRemainingDailyQuota();
}

/* ---------- 시트 ---------- */
function getInspSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(INSP_TAB);
  if (!sh) {
    sh = ss.insertSheet(INSP_TAB);
    sh.getRange(1, 1, 1, INSP_HEADERS.length).setValues([INSP_HEADERS]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function getInspRecipients_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(INSP_MAIL_TAB);
  if (!sh) {
    sh = ss.insertSheet(INSP_MAIL_TAB);
    sh.getRange("A1:B2").setValues([["이메일", "← 한 줄에 주소 하나. 받는 사람 추가·삭제는 이 열에서 하세요."], [INSP_MAIL_SEED, ""]]);
    sh.setColumnWidth(1, 260);
  }
  var last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
    var m = String(r[0]).trim();
    if (m.indexOf("@") > 0 && out.indexOf(m) < 0) out.push(m);
  });
  return out;
}

function inspFolder_() {
  var it = DriveApp.getFoldersByName(INSP_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(INSP_FOLDER);
}

/* ---------- 신청 저장 ---------- */
function saveInsp_(d) {
  if (!INSP_TPL[d.type]) return { error: "missing type" };
  var req = ["writer", "yard", "hull", "equip", "reqDate", "reqTime"];
  for (var i = 0; i < req.length; i++) if (!d[req[i]]) return { error: "missing " + req[i] };
  if (!d.parts || !d.parts.length) return { error: "missing parts" };

  var today = Utilities.formatDate(new Date(), INSP_TZ, "yyyy-MM-dd");
  if (!(String(d.reqDate) > today)) return { error: "date" };   // 검사 요구일은 내일부터

  var sh = getInspSheet_();
  var last = sh.getLastRow();
  var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var cids = last > 1 ? sh.getRange(2, 15, last - 1, 1).getValues() : [];
  if (d.clientId) {
    for (var c = 0; c < cids.length; c++) {
      if (String(cids[c][0]) === d.clientId) return { ok: true, id: ids[c][0], duplicate: true };
    }
  }

  // 관리번호: PT-YYYYMMDD-NN (성능), SH-YYYYMMDD-NN (출하)
  var prefix = INSP_PREFIX[d.type] + "-" + today.replace(/-/g, "");
  var n = 0;
  for (var j = 0; j < ids.length; j++) if (String(ids[j][0]).indexOf(prefix + "-") === 0) n++;
  var id = prefix + "-" + ("0" + (n + 1)).slice(-2);

  // 품명: 프로젝트 탭의 품명
  var pn = d.pn || "";
  try {
    getMaster_().projects.forEach(function (p) {
      if (p.yard === d.yard && p.hull === String(d.hull) && p.name === d.equip && p.parts && p.parts.length) pn = p.parts.join(", ");
    });
  } catch (e) {}

  var checkTotal = 0, checkDone = 0;
  (d.checks || []).forEach(function (sec) { sec.forEach(function (v) { checkTotal++; if (v) checkDone++; }); });
  var snap = { id: id, type: d.type, date: today, writer: d.writer, yard: d.yard, hull: d.hull, equip: d.equip, pn: pn,
    parts: d.parts, reqDate: d.reqDate, reqTime: d.reqTime, checks: d.checks || [] };
  var partsTxt = d.parts.map(function (p) { return p.name + " × " + p.qty; }).join(", ");

  var row = [id, Utilities.formatDate(new Date(), INSP_TZ, "yyyy-MM-dd HH:mm:ss"), today, d.writer, INSP_LABEL[d.type],
    d.yard, d.hull, d.equip, pn, d.reqDate + " " + d.reqTime, partsTxt, checkDone + "/" + checkTotal, "", "처리중", d.clientId || "", JSON.stringify(snap)];
  var rng = sh.getRange(last + 1, 1, 1, INSP_HEADERS.length);
  rng.setNumberFormat("@");
  rng.setValues([row.map(String)]);

  var res = processInsp_(last + 1);
  return { ok: true, id: id, pdf: res.pdf, mail: res.mail };
}

/* ---------- PDF 만들기 + 메일 ---------- */
function processInsp_(r) {
  var sh = getInspSheet_();
  var row = sh.getRange(r, 1, 1, INSP_HEADERS.length).getValues()[0].map(String);
  var d = JSON.parse(row[15]);
  var out = { pdf: "", mail: row[13] };
  var blob;
  try {
    blob = makeInspPdf_(d);   // 드라이브에 저장하지 않고 메일 첨부용으로만 만든다
  } catch (e) {
    Logger.log("PDF 생성 실패: " + e);
    sh.getRange(r, 14).setValue("PDF실패");
    out.mail = "PDF실패";
    return out;
  }
  try {
    var to = getInspRecipients_();
    if (!to.length) { sh.getRange(r, 14).setValue("수신자없음"); out.mail = "수신자없음"; return out; }
    MailApp.sendEmail({
      to: to.join(","),
      subject: inspSubject_(d),
      htmlBody: inspMailHtml_(d, row, ""),
      body: inspSubject_(d),
      attachments: [blob],
      name: "작업일지"
    });
    sh.getRange(r, 14).setValue("발송완료");
    out.mail = "발송완료";
  } catch (e2) {
    Logger.log("메일 발송 실패: " + e2);
    sh.getRange(r, 14).setValue("발송대기");
    out.mail = "발송대기";
  }
  return out;
}

function inspEsc_(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }

function inspDow_(dateStr) {
  var p = String(dateStr).split("-");
  return "일월화수목금토".charAt(new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getDay());
}
function inspReqLabel_(d) {
  var p = String(d.reqDate).split("-");
  return Number(p[1]) + "/" + Number(p[2]) + "(" + inspDow_(d.reqDate) + ") " + d.reqTime;
}
function inspSubject_(d) {
  return "[검사신청] " + d.id + " · " + d.hull + " " + INSP_LABEL[d.type] + " · 검사요구 " + inspReqLabel_(d);
}
function inspMailHtml_(d, row, pdfUrl) {
  var td = 'style="padding:6px 10px;border-bottom:1px solid #eef0f3;font-size:13px"';
  var th = 'style="padding:6px 10px;border-bottom:1px solid #eef0f3;font-size:13px;color:#6b7280;width:110px;white-space:nowrap"';
  var tr = function (k, v) { return "<tr><th " + th + " align=left>" + k + "</th><td " + td + ">" + inspEsc_(v) + "</td></tr>"; };
  return '<div style="font-family:-apple-system,\'Malgun Gothic\',sans-serif;max-width:640px;color:#1f2937">' +
    '<div style="font-size:18px;font-weight:bold;margin-bottom:8px">' + inspEsc_(INSP_LABEL[d.type]) + ' 신청</div>' +
    '<table style="border-collapse:collapse;width:100%">' +
    tr("관리번호", d.id) + tr("프로젝트", d.yard + "-" + d.hull + "-" + d.equip) + tr("품명", d.pn) +
    tr("검사 요구", d.reqDate + " (" + inspDow_(d.reqDate) + ") " + d.reqTime) + tr("신청자", d.writer) +
    tr("검사 부품", row[10]) + tr("확인 사항", row[11] + " 확인") + '</table>' +
    '<p style="font-size:12.5px;color:#6b7280;margin-top:12px">신청서 PDF를 첨부했습니다.' + (pdfUrl ? ' <a href="' + pdfUrl + '">드라이브에서 열기</a>' : '') + '</p></div>';
}

function makeInspPdf_(d) {
  var tmp = DriveApp.getFileById(INSP_TPL[d.type]).makeCopy("tmp_" + d.id);   // 구글 문서는 저장 용량에 잡히지 않고, 끝나면 휴지통으로 보낸다
  try {
    var doc = DocumentApp.openById(tmp.getId());
    fillInspDoc_(doc, d);
    doc.saveAndClose();
    return tmp.getAs("application/pdf").setName(d.id + ".pdf");
  } finally {
    tmp.setTrashed(true);
  }
}

/* ---------- 문서 틀 채우기 (칸 위치가 아니라 칸의 글자를 보고 찾는다) ---------- */
function norm_(s) { return String(s).replace(/\s+/g, ""); }

function setCellText_(cell, text, bold) {
  var t = cell.editAsText();
  t.setText(text);
  if (String(text).length) { t.setFontSize(10); t.setBold(!!bold); }
}

function setAfterLabel_(table, label, text) {
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var c = 0; c < row.getNumCells() - 1; c++) {
      if (norm_(row.getCell(c).getText()) === norm_(label)) { setCellText_(row.getCell(c + 1), text); return true; }
    }
  }
  return false;
}

function fillInspDoc_(doc, d) {
  var T = doc.getBody().getTables();
  var head = T[0];
  setAfterLabel_(head, "호선번호", d.yard + "-" + d.hull + "-" + d.equip);
  setAfterLabel_(head, "품명", d.pn || "");
  setAfterLabel_(head, "검사요구일시", d.reqDate + " (" + inspDow_(d.reqDate) + ") " + d.reqTime);
  setAfterLabel_(head, "신청일자", d.date);

  // 1. 검사신청 내용 (NO, 부품명, 도면번호, 수량, 검사종류, 수검담당자/Tel)
  var pt = T[1], hdr = pt.getRow(0), iName = -1, iQty = -1, iMgr = hdr.getNumCells() - 1;
  for (var c = 0; c < hdr.getNumCells(); c++) {
    var h = norm_(hdr.getCell(c).getText());
    if (h === "부품명") iName = c;
    if (h === "수량") iQty = c;
    if (h.indexOf("수검담당자") === 0) iMgr = c;
  }
  var lastData = 0;
  for (var r = 1; r < pt.getNumRows(); r++) {
    if (norm_(pt.getRow(r).getCell(0).getText()) === "검사종류") break;
    lastData = r;
  }
  var need = d.parts.length;
  while (lastData < need) {   // 부품이 양식의 칸보다 많으면 마지막 데이터 줄을 복사해 늘린다
    pt.insertTableRow(lastData + 1, pt.getRow(lastData).copy());
    lastData++;
  }
  for (var i = 0; i < need; i++) {
    var row = pt.getRow(1 + i);
    setCellText_(row.getCell(0), String(i + 1));
    if (iName >= 0) setCellText_(row.getCell(iName), d.parts[i].name);
    if (iQty >= 0) setCellText_(row.getCell(iQty), String(d.parts[i].qty));
    setCellText_(row.getCell(iMgr), d.writer);
  }

  // 확인 사항 표: 체크한 줄의 '확인' 칸에 ✓
  var idx = d.type === "perf" ? [2, 3, 4] : [2];
  (d.checks || []).forEach(function (sec, si) {
    var tb = T[idx[si]];
    if (!tb) return;
    sec.forEach(function (on, j) {
      if (!on || 1 + j >= tb.getNumRows()) return;
      var row = tb.getRow(1 + j);
      setCellText_(row.getCell(row.getNumCells() - 1), "✓", true);
    });
  });
}

/* ---------- 재발송 (1시간마다, 실패·미발송 건) ---------- */
function retryInspMail() {
  var sh = getInspSheet_(), last = sh.getLastRow();
  if (last < 2) return;
  var st = sh.getRange(2, 14, last - 1, 1).getValues();
  for (var i = 0; i < st.length; i++) {
    var s = String(st[i][0]);
    if (s === "발송대기" || s === "수신자없음" || s === "PDF실패") processInsp_(i + 2);
  }
}

/* ---------- 내 검사신청 목록 ---------- */
function getMyInsp_(writer) {
  if (!writer) return [];
  var sh = getInspSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, INSP_HEADERS.length).getDisplayValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0 && out.length < 100; i--) {
    var r = rows[i];
    if (r[3] !== writer) continue;
    out.push({ id: r[0], date: r[2], writer: r[3], type: r[4], yard: r[5], hull: r[6], equip: r[7], pn: r[8],
      req: r[9], parts: r[10], checks: r[11], pdf: r[12], status: r[13] });
  }
  return out;
}
