/* =====================================================================
   외주 정산 집계 (매월 25일 기준)
   - 정산 기간: 전월 26일 ~ 당월 25일. 매월 26일 08시에 메일로 보낸다.
   - 받는 사람: 시트의 '정산수신' 탭 A열 (한 줄에 주소 하나)
   - 집계 결과는 '외주정산' 탭에도 기간별로 남긴다 (같은 기간을 다시 돌리면 그 기간 줄만 새로 쓴다)
   - 메일에 날짜별 상세 CSV(엑셀에서 열림)를 첨부한다
   - 처음 한 번 setupSettlement 를 ▶ 실행한다 (탭 생성 + 매월 26일 트리거)
   - Report.gs 의 함수(netMins_, esc_, td_, th_, sec_, kp_)를 같이 쓴다. 웹앱 재배포 불필요
   ===================================================================== */
var SETTLE_DAY = 25;              // 정산 기준일
var SETTLE_MAIL_TAB = "정산수신";
var SETTLE_TAB = "외주정산";
var SETTLE_SEED = "erai77@techflower.co.kr";
var SETTLE_TZ = "Asia/Seoul";

/* ---------- 설치 / 실행 ---------- */
function setupSettlement() {
  getSettleRecipients_();
  getSettleSheet_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "sendSettlement") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendSettlement").timeBased().onMonthDay(SETTLE_DAY + 1).atHour(8).nearMinute(0).inTimezone(SETTLE_TZ).create();
}

// 트리거가 매월 26일에 호출: 방금 끝난 정산 기간을 보낸다
function sendSettlement() {
  var p = settlePeriod_(Utilities.formatDate(new Date(), SETTLE_TZ, "yyyy-MM-dd"));
  return sendSettlementFor_(p.start, p.end, getSettleRecipients_());
}

// 시험: 가장 최근에 끝난 정산 기간을 '정산수신' 첫 번째 주소로만 보낸다. '외주정산' 탭에는 쓰지 않는다
function testSettlement() {
  var p = settlePeriod_(Utilities.formatDate(new Date(), SETTLE_TZ, "yyyy-MM-dd"));
  Logger.log(sendSettlementFor_(p.start, p.end, getSettleRecipients_().slice(0, 1), true) + " (" + p.start + " ~ " + p.end + ")");
}

// 지난 기간을 다시 보내고 싶을 때: 편집기에서 아래 날짜를 바꿔 ▶ 실행
function resendSettlement() {
  sendSettlementFor_("2026-10-26", "2026-11-25", getSettleRecipients_());
}

/* ---------- 기간 계산 ---------- */
function stShift_(d, n) {
  var p = d.split("-"), x = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + n));
  return x.getUTCFullYear() + "-" + ("0" + (x.getUTCMonth() + 1)).slice(-2) + "-" + ("0" + x.getUTCDate()).slice(-2);
}
function stDay_(y, m, d) {   // m: 1~12, 넘치거나 모자라면 앞뒤 해로 넘긴다
  var x = new Date(Date.UTC(y, m - 1, d));
  return x.getUTCFullYear() + "-" + ("0" + (x.getUTCMonth() + 1)).slice(-2) + "-" + ("0" + x.getUTCDate()).slice(-2);
}
// today 이전에 마지막으로 끝난 정산 기간 (26일에 부르면 전월 26일 ~ 어제(25일))
function settlePeriod_(today) {
  var y = Number(today.slice(0, 4)), m = Number(today.slice(5, 7)), d = Number(today.slice(8, 10));
  var endM = d > SETTLE_DAY ? m : m - 1;
  var end = stDay_(y, endM, SETTLE_DAY);
  var start = stShift_(stDay_(y, endM - 1, SETTLE_DAY), 1);
  return { start: start, end: end };
}

/* ---------- 시트 ---------- */
function getSettleRecipients_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SETTLE_MAIL_TAB);
  if (!sh) {
    sh = ss.insertSheet(SETTLE_MAIL_TAB);
    sh.getRange("A1:B2").setValues([["이메일", "← 외주 정산 메일 받는 사람. 한 줄에 주소 하나."], [SETTLE_SEED, ""]]);
    sh.setColumnWidth(1, 260);
  }
  var last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
    var x = String(r[0]).trim();
    if (x.indexOf("@") > 0 && out.indexOf(x) < 0) out.push(x);
  });
  return out;
}

var SETTLE_HEAD = ["정산기간", "업체", "작업자", "작업일수", "일지", "MH", "작업 날짜", "집계시각"];
function getSettleSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SETTLE_TAB);
  if (!sh) {
    sh = ss.insertSheet(SETTLE_TAB);
    sh.getRange(1, 1, 1, SETTLE_HEAD.length).setValues([SETTLE_HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 130); sh.setColumnWidth(7, 420);
  }
  return sh;
}

/* ---------- 집계 ---------- */
function settleBuild_(logs, start, end) {
  var ext = logs.filter(function (l) { return l.date >= start && l.date <= end && (l.affil === "외주" || l.vendor); });
  var V = {}, vk = [];
  ext.forEach(function (l) {
    var v = l.vendor || "(업체 미기재)";
    if (!V[v]) { V[v] = { w: {}, wk: [], m: 0, n: 0, dates: {} }; vk.push(v); }
    var o = V[v], m = netMins_(l.start, l.end);
    if (!o.w[l.writer]) { o.w[l.writer] = { m: 0, n: 0, dates: {}, proj: {} }; o.wk.push(l.writer); }
    var w = o.w[l.writer];
    w.m += m; w.n++; w.dates[l.date] = (w.dates[l.date] || 0) + m;
    var pk = l.yard + "-" + l.hull + "-" + l.equip; w.proj[pk] = (w.proj[pk] || 0) + m;
    o.m += m; o.n++; o.dates[l.date] = 1;
  });
  vk.sort();
  var vendors = vk.map(function (v) {
    var o = V[v];
    o.wk.sort();
    return { name: v, mins: o.m, logs: o.n, days: Object.keys(o.dates).length,
      workers: o.wk.map(function (k) {
        var w = o.w[k], ds = Object.keys(w.dates).sort();
        var proj = Object.keys(w.proj).sort(function (a, b) { return w.proj[b] - w.proj[a]; });
        return { name: k, mins: w.m, logs: w.n, days: ds.length, dates: ds, dayMins: w.dates, proj: proj };
      }) };
  });
  ext.sort(function (a, b) {
    return String(a.vendor).localeCompare(String(b.vendor)) || String(a.writer).localeCompare(String(b.writer)) ||
      String(a.date).localeCompare(String(b.date)) || String(a.start).localeCompare(String(b.start));
  });
  return { vendors: vendors, rows: ext };
}

function stH_(m) { return Math.round(m / 60 * 10) / 10; }
function stDates_(ds) {   // 2026-11-03,2026-11-04 → 11/3, 11/4
  return ds.map(function (d) { return Number(d.slice(5, 7)) + "/" + Number(d.slice(8)); }).join(", ");
}

var SETTLE_OVER_MIN = 480;   // 하루 8시간(식사시간 제외) 초과면 청색
// '공휴일' 탭(현황 탭과 같이 씀)의 날짜 → 이름. 탭이 없으면 빈 목록
function stHolidays_() {
  var out = {};
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("공휴일");
    if (!sh || sh.getLastRow() < 2) return out;
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var v = r[0], d = Object.prototype.toString.call(v) === "[object Date]" ? Utilities.formatDate(v, tz, "yyyy-MM-dd") : String(v).trim().slice(0, 10);
      if (/^\d{4}-\d{2}-\d{2}$/.test(d)) out[d] = String(r[1] || "공휴일");
    });
  } catch (e) { Logger.log("공휴일 읽기 실패: " + e); }
  return out;
}
function stIsRed_(d, hol) {
  var p = d.split("-");
  return hol[d] !== undefined || new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]))).getUTCDay() === 0;
}
// 작업 날짜 목록(HTML): 일요일·공휴일은 적색, 8시간 초과는 청색(둘 다면 적색 굵게 + 밑줄)
function stDatesHtml_(w, hol) {
  return w.dates.map(function (d) {
    var red = stIsRed_(d, hol), over = (w.dayMins[d] || 0) > SETTLE_OVER_MIN;
    var st = "white-space:nowrap";
    if (red) st += ";color:#dc2626;font-weight:bold";
    else if (over) st += ";color:#2563eb;font-weight:bold";
    if (red && over) st += ";text-decoration:underline";
    var tip = (hol[d] ? hol[d] + " · " : "") + stH_(w.dayMins[d] || 0) + "h";
    return '<span title="' + esc_(tip) + '" style="' + st + '">' + Number(d.slice(5, 7)) + "/" + Number(d.slice(8)) + "</span>";
  }).join(", ");
}

function settleHtml_(t, start, end, hol) {
  hol = hol || {};
  var tm = 0, tn = 0, tw = 0;
  t.vendors.forEach(function (v) { tm += v.mins; tn += v.logs; tw += v.workers.length; });
  var h = '<div style="max-width:1300px;font-family:-apple-system,\'Malgun Gothic\',sans-serif;color:#1f2937;border:1px solid #e5e7eb;border-radius:10px;padding:22px 24px">';
  h += '<table style="width:100%;border-collapse:collapse"><tr><td style="font-size:20px;font-weight:bold">외주 정산 집계 · ' + start + ' ~ ' + end + '</td>' +
    '<td style="text-align:right;font-size:12.5px;color:#6b7280">' + REPORT_SITE + ' · 매월 ' + SETTLE_DAY + '일 기준 · 자동 발송</td></tr></table>';
  h += '<table style="width:100%;border-collapse:collapse;margin-top:14px"><tr>' + kp_("업체", t.vendors.length + "곳") + kp_("작업자", tw + "명") +
    kp_("일지", tn + "건") + kp_("총 MH", stH_(tm) + "h") + '</tr></table>';
  if (!t.vendors.length) return h + '<p style="margin-top:18px">이 기간에는 외주 작업 일지가 없습니다.</p></div>';

  h += sec_("업체별 요약") + '<table style="width:100%;border-collapse:collapse"><tr>' + th_("업체", ";width:200px") + th_("작업자", ";text-align:right;width:80px") +
    th_("작업일수", ";text-align:right;width:90px") + th_("일지", ";text-align:right;width:80px") + th_("MH", ";text-align:right;width:90px") + th_("") + '</tr>';
  t.vendors.forEach(function (v) {
    h += '<tr>' + td_("<b>" + esc_(v.name) + "</b>") + td_(v.workers.length + "명", ";text-align:right") + td_(v.days + "일", ";text-align:right") +
      td_(v.logs + "건", ";text-align:right") + td_("<b>" + stH_(v.mins) + "h</b>", ";text-align:right") + td_("") + '</tr>';
  });
  h += '<tr>' + td_("<b>합계</b>", ";border-bottom:0") + td_("<b>" + tw + "명</b>", ";text-align:right;border-bottom:0") + td_("-", ";text-align:right;border-bottom:0") +
    td_("<b>" + tn + "건</b>", ";text-align:right;border-bottom:0") + td_("<b>" + stH_(tm) + "h</b>", ";text-align:right;border-bottom:0") + td_("", ";border-bottom:0") + '</tr></table>';

  h += sec_("작업자별 상세 <span style=\"font-size:12px;font-weight:normal;color:#6b7280\">작업 날짜: " +
    "<b style=\"color:#2563eb\">청색</b> 하루 8시간 초과 · <b style=\"color:#dc2626\">적색</b> 일요일·공휴일 · " +
    "<b style=\"color:#dc2626;text-decoration:underline\">적색 밑줄</b> 둘 다</span>");
  t.vendors.forEach(function (v) {
    h += '<div style="border:1px solid #d1d5db;border-radius:6px;margin-bottom:14px;overflow:hidden">' +
      '<div style="background:#f9fafb;padding:8px 12px;font-size:14px"><b>' + esc_(v.name) + '</b> <span style="color:#6b7280;font-size:12px">· ' +
      v.workers.length + '명 · ' + stH_(v.mins) + 'h</span></div><table style="width:100%;border-collapse:collapse"><tr>' +
      th_("작업자", ";width:110px") + th_("작업일수", ";text-align:right;width:80px") + th_("일지", ";text-align:right;width:60px") + th_("MH", ";text-align:right;width:70px") +
      th_("작업 날짜", ";width:45%") + th_("주요 프로젝트") + '</tr>';
    v.workers.forEach(function (w) {
      h += '<tr>' + td_(esc_(w.name)) + td_(w.days + "일", ";text-align:right") + td_(w.logs + "건", ";text-align:right") + td_("<b>" + stH_(w.mins) + "h</b>", ";text-align:right") +
        td_(stDatesHtml_(w, hol), ";white-space:normal;color:#374151;font-size:12.5px") + td_(esc_(w.proj.slice(0, 3).join(", ") + (w.proj.length > 3 ? " 외 " + (w.proj.length - 3) : "")), ";white-space:normal;color:#6b7280;font-size:12.5px") + '</tr>';
    });
    h += '</table></div>';
  });
  h += '<div style="font-size:12.5px;color:#6b7280;border-top:1px solid #e5e7eb;padding-top:10px;margin-top:18px">MH는 시작~종료 시간에서 식사시간(12:00~13:00, 17:30~18:00)을 뺀 값입니다. ' +
    '날짜별 상세는 첨부 CSV(엑셀에서 열림)에 있습니다. 받는 사람은 시트의 \'' + SETTLE_MAIL_TAB + '\' 탭에서 바꿀 수 있습니다.</div></div>';
  return h;
}

function settleCsv_(t) {
  var q = function (s) { s = String(s == null ? "" : s); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  var lines = [["업체", "작업자", "날짜", "프로젝트", "작업구분", "작업항목", "시작", "종료", "MH", "상태"].join(",")];
  t.rows.forEach(function (l) {
    lines.push([l.vendor || "(업체 미기재)", l.writer, l.date, l.yard + "-" + l.hull + "-" + l.equip, l.type, (l.items || []).join(" / "),
      l.start, l.end, stH_(netMins_(l.start, l.end)), l.status].map(q).join(","));
  });
  return "﻿" + lines.join("\r\n");   // BOM: 엑셀에서 한글이 깨지지 않게
}

/* ---------- 발송 ---------- */
function sendSettlementFor_(start, end, to, test) {
  if (!to.length) return "skip: 받는 사람 없음";
  var t = settleBuild_(readLogs_(), start, end);
  var period = start + " ~ " + end;
  if (!test) {
    var sh = getSettleSheet_(), last = sh.getLastRow();
    if (last > 1) {   // 같은 기간의 이전 집계는 지우고 새로 쓴다
      var col = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (var i = col.length - 1; i >= 0; i--) if (col[i][0] === period) sh.deleteRow(i + 2);
    }
    var now = Utilities.formatDate(new Date(), SETTLE_TZ, "yyyy-MM-dd HH:mm");
    var out = [];
    t.vendors.forEach(function (v) {
      v.workers.forEach(function (w) { out.push([period, v.name, w.name, w.days, w.logs, stH_(w.mins), stDates_(w.dates), now]); });
    });
    if (out.length) {
      var r = sh.getRange(sh.getLastRow() + 1, 1, out.length, SETTLE_HEAD.length);
      r.setNumberFormat("@"); r.setValues(out.map(function (x) { return x.map(String); }));
    }
  }
  var tm = 0; t.vendors.forEach(function (v) { tm += v.mins; });
  var subject = (test ? "[시험] " : "") + "[외주 정산] " + period + " · 업체 " + t.vendors.length + "곳 / " + stH_(tm) + "h";
  var csv = Utilities.newBlob(settleCsv_(t), "text/csv", "외주정산_" + start + "_" + end + ".csv");
  MailApp.sendEmail({ to: to.join(","), subject: subject, htmlBody: settleHtml_(t, start, end, stHolidays_()), body: subject, name: REPORT_SENDER, attachments: [csv] });
  return "sent " + t.rows.length + "건";
}
