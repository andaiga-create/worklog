/* =====================================================================
   매일 08시 전일 작업일지 보고 메일
   - 받는 사람: 시트의 '보고수신' 탭 A열 (한 줄에 주소 하나, 추가/삭제는 거기서)
   - 전일에 일지가 한 건도 없으면(휴일 등) 보내지 않는다
   - 처음 한 번 setupDailyReport 를 실행해 매일 08시 트리거를 만든다
   - 이 파일은 Code.gs 와 같은 프로젝트에 '새 파일'로 추가한다(웹앱 재배포 불필요)
   ===================================================================== */
var REPORT_TAB = "보고수신";
var REPORT_SEED = "erai77@techflower.co.kr";
var REPORT_SENDER = "작업일지";
var REPORT_SITE = "김해공장";
var REPORT_TZ = "Asia/Seoul";

/* ---------- 설치 / 실행 진입점 ---------- */
// 1) Apps Script 편집기에서 이 함수를 선택해 ▶ 한 번만 실행: 보고수신 탭 생성 + 매일 08시 트리거
function setupDailyReport() {
  getReportSheet_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "sendDailyReport") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendDailyReport").timeBased().everyDays(1).atHour(8).nearMinute(0).inTimezone(REPORT_TZ).create();
}

// 2) 트리거가 매일 호출: 어제 일지를 보고서로 만들어 보낸다 (일지가 없으면 보내지 않음)
function sendDailyReport() {
  var y = Utilities.formatDate(new Date(new Date().getTime() - 86400000), REPORT_TZ, "yyyy-MM-dd");
  return sendReportForDate_(y, getRecipients_());
}

// 3) 시험 발송: 일지가 있는 가장 최근 날짜의 보고서를 '보고수신' 첫 번째 주소로만 보낸다
function testReport() {
  var logs = readLogs_(), last = "";
  logs.forEach(function (l) { if (l.date > last) last = l.date; });
  if (!last) { Logger.log("일지가 없습니다."); return; }
  var to = getRecipients_().slice(0, 1);
  Logger.log(sendReportForDate_(last, to) + " (" + last + " → " + to.join(",") + ")");
}

/* ---------- 받는 사람 ---------- */
function getReportSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(REPORT_TAB);
  if (!sh) {
    sh = ss.insertSheet(REPORT_TAB);
    sh.getRange("A1:B2").setValues([["이메일", "← 한 줄에 주소 하나. 받는 사람 추가·삭제는 이 열에서 하세요."], [REPORT_SEED, ""]]);
    sh.setColumnWidth(1, 260);
  }
  return sh;
}

function getRecipients_() {
  var sh = getReportSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var out = [];
  sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
    var m = String(r[0]).trim();
    if (m.indexOf("@") > 0 && out.indexOf(m) < 0) out.push(m);
  });
  return out;
}

/* ---------- 발송 ---------- */
function sendReportForDate_(d, to) {
  var all = readLogs_();
  var day = all.filter(function (l) { return l.date === d; });
  if (!day.length) return "skip: 일지 없음";
  if (!to.length) return "skip: 받는 사람 없음";
  var rep = buildReport_(d, day, all);
  MailApp.sendEmail({ to: to.join(","), subject: rep.subject, htmlBody: rep.html, body: rep.subject, name: REPORT_SENDER, inlineImages: rep.images });
  return "sent " + day.length + "건";
}

/* ---------- 계산 도우미 ---------- */
function mins_(l) {
  var a = String(l.start).split(":"), b = String(l.end).split(":");
  var m = (Number(b[0]) * 60 + Number(b[1])) - (Number(a[0]) * 60 + Number(a[1]));
  return m > 0 ? m : 0;
}
function mh_(m) { return (Math.round(m / 60 * 10) / 10) + "h"; }
function hh_(m) { return (Math.round(m / 60 * 10) / 10) + "h"; }
function esc_(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function isExt_(l) { return l.affil === "외주" || !!l.vendor; }
function who_(l) { return esc_(l.writer) + ' <span style="color:#6b7280;font-size:12px">(' + esc_(isExt_(l) ? (l.vendor || "외주") : "직영") + ')</span>'; }
function chip_(status) {
  var s = String(status || "-");
  var c = /^보류/.test(s) ? ["#fef3c7", "#92400e"] : /^완료/.test(s) ? ["#dcfce7", "#166534"] : ["#dbeafe", "#1e40af"];
  return '<span style="display:inline-block;font-size:11.5px;padding:1px 8px;border-radius:9px;white-space:nowrap;background:' + c[0] + ';color:' + c[1] + '">' + esc_(s) + '</span>';
}
function projKey_(l) { return l.yard + "-" + l.hull + "-" + l.equip; }
function uniq_(arr) { var o = []; arr.forEach(function (x) { if (x && o.indexOf(x) < 0) o.push(x); }); return o; }

/* ---------- 원형 그래프(이미지, Apps Script 내장 차트 사용) ---------- */
var PIE_COLORS = ["#3b82f6", "#f59e0b", "#10b981", "#ef4444", "#8b5cf6", "#06b6d4", "#ec4899", "#9ca3af"];
function makePie_(cid, pairs) {
  try {
    var dt = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "구분").addColumn(Charts.ColumnType.NUMBER, "MH");
    pairs.forEach(function (p) { dt.addRow([p[0], Math.round(p[1] / 60 * 10) / 10]); });
    var colors = pairs.map(function (p, i) { return PIE_COLORS[i % PIE_COLORS.length]; });
    var chart = Charts.newPieChart().setDataTable(dt.build()).setDimensions(280, 200).setColors(colors)
      .setLegendPosition(Charts.Position.NONE).setOption("pieSliceText", "percentage").setBackgroundColor("#fcfcfd").build();
    return chart.getAs("image/png").setName(cid + ".png");
  } catch (e) { Logger.log("원형 그래프 실패: " + e); return null; }   // 그래프가 실패해도 표만으로 보고서는 나간다
}

/* ---------- HTML 조각 ---------- */
var TD = "padding:6px 9px;font-size:13px;white-space:nowrap;border-bottom:1px solid #eef0f3;text-align:left;vertical-align:middle;color:#1f2937";
var TH = "padding:6px 9px;font-size:12.5px;background:#f3f4f6;color:#374151;font-weight:normal;text-align:left;white-space:nowrap;border-bottom:1px solid #d1d5db";
function td_(c, ex) { return '<td style="' + TD + (ex || "") + '">' + c + '</td>'; }
function th_(c, ex) { return '<th style="' + TH + (ex || "") + '">' + c + '</th>'; }
function sec_(t) { return '<div style="font-size:15px;font-weight:bold;border-bottom:2px solid #1f2937;padding-bottom:4px;margin:24px 0 8px">' + t + '</div>'; }
function kp_(label, val, bg, fg) {
  return '<td style="background:' + (bg || "#f3f4f6") + ';text-align:center;padding:9px 4px;border-radius:8px;width:11%"><div style="font-size:11.5px;color:' + (fg || "#6b7280") + '">' + label + '</div><div style="font-size:20px;font-weight:bold;color:' + (fg || "#1f2937") + '">' + val + '</div></td><td style="width:8px"></td>';
}

/* ---------- 보고서 본문 ---------- */
function buildReport_(d, day, all) {
  var p = d.split("-");
  var dow = ["일", "월", "화", "수", "목", "금", "토"][new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getDay()];
  var dlabel = d + "(" + dow + ")";

  // 품명(검색용 정보) 가져오기: 실패해도 보고서는 보낸다
  var parts = {};
  try { getMaster_().projects.forEach(function (pr) { parts[pr.yard + "-" + pr.hull + "-" + pr.name] = pr.parts || []; }); } catch (e) {}

  // 요약
  var totMin = 0, st = { ok: 0, pg: 0, hd: 0 };
  day.forEach(function (l) {
    totMin += mins_(l);
    if (/^보류/.test(l.status)) st.hd++; else if (/^완료/.test(l.status)) st.ok++; else st.pg++;
  });
  var writers = uniq_(day.map(function (l) { return l.writer; }));
  var nExt = uniq_(day.filter(isExt_).map(function (l) { return l.writer; })).length;
  var nOwn = writers.length - nExt;
  var projKeys = uniq_(day.map(projKey_)).sort();

  var subject = "[생산부서 전일 작업보고] " + dlabel + " · 일지 " + day.length + "건 / 작업자 " + writers.length + "명 / " + hh_(totMin) + (st.hd ? " · 보류 " + st.hd + "건" : "");

  var h = '<div style="max-width:1500px;font-family:-apple-system,\'Malgun Gothic\',sans-serif;color:#1f2937;border:1px solid #e5e7eb;border-radius:10px;padding:22px 24px">';
  h += '<table style="width:100%;border-collapse:collapse"><tr><td style="font-size:20px;font-weight:bold">' + dlabel + ' 작업일지 보고</td><td style="text-align:right;font-size:12.5px;color:#6b7280">' + REPORT_SITE + ' · 자동 발송 · 작성일 ' + d + ' 전체</td></tr></table>';

  h += '<table style="width:100%;border-collapse:collapse;margin-top:14px"><tr>' +
    kp_("일지", day.length + "건") + kp_("작업자", writers.length + "명") + kp_("직영", nOwn + "명") + kp_("외주", nExt + "명") +
    kp_("총 작업시간", hh_(totMin)) + kp_("프로젝트", projKeys.length + "개") +
    kp_("완료", st.ok, "#dcfce7", "#166534") + kp_("진행중", st.pg, "#dbeafe", "#1e40af") + kp_("보류", st.hd, "#fef3c7", "#92400e") + '</tr></table>';

  // 확인 필요 사항 (보류)
  var holds = day.filter(function (l) { return /^보류/.test(l.status); });
  if (holds.length) {
    h += sec_("확인 필요 사항") + '<table style="width:100%;border-collapse:collapse">';
    holds.forEach(function (l) {
      h += '<tr>' + td_(chip_(l.status), ";width:110px") + td_("<b>" + esc_(projKey_(l)) + "</b>", ";width:210px") +
        td_(esc_(l.type) + " · " + who_(l), ";width:240px") + td_(esc_(l.detail || "-"), ";white-space:normal;color:#6b7280;font-size:12.5px") + '</tr>';
    });
    h += '</table>';
  }

  // 프로젝트별 작업 상세
  h += sec_("프로젝트별 작업 상세");
  var images = {}, pieN = 0;
  projKeys.forEach(function (k) {
    var rows = day.filter(function (l) { return projKey_(l) === k; });
    rows.sort(function (a, b) {
      return (isExt_(a) - isExt_(b)) || String(a.vendor).localeCompare(String(b.vendor)) || String(a.writer).localeCompare(String(b.writer)) || String(a.start).localeCompare(String(b.start));
    });
    var rMin = 0, anyHold = false, allDone = true;
    rows.forEach(function (l) { rMin += mins_(l); if (/^보류/.test(l.status)) anyHold = true; if (!/^완료/.test(l.status)) allDone = false; });
    var pst = anyHold ? "보류 포함" : (allDone ? "완료" : "진행중");
    var pcolor = anyHold ? "보류" : (allDone ? "완료" : "진행중");
    var pn = (parts[k] || []).slice(0, 3).join(", ");

    h += '<div style="border:1px solid #d1d5db;border-radius:6px;margin-bottom:14px;overflow:hidden">';
    h += '<table style="width:100%;border-collapse:collapse;background:#f9fafb"><tr><td style="padding:8px 12px;white-space:nowrap"><b style="font-size:14px">' + esc_(k) + '</b> &nbsp;<span style="color:#6b7280;font-size:12px">' + esc_(pn) + '</span></td>' +
      '<td style="padding:8px 12px;text-align:right;white-space:nowrap;color:#6b7280;font-size:12px">전일 작업자 ' + uniq_(rows.map(function (l) { return l.writer; })).length + '명 · 일지 ' + rows.length + '건 · <b style="color:#1f2937">' + hh_(rMin) + '</b> &nbsp;' + chip_(pcolor).replace(esc_(pcolor), pst) + '</td></tr></table>';

    // 전일 상세 (카드 위쪽)
    h += '<table style="width:100%;border-collapse:collapse"><tr>' + th_("작업자", ";width:150px") + th_("시간", ";width:120px") + th_("시수", ";text-align:right;width:54px") +
      th_("작업구분", ";width:90px") + th_("작업항목", ";width:27%") + th_("메모", ";width:38%") + th_("상태", ";width:120px") + '</tr>';
    rows.forEach(function (l) {
      h += '<tr>' + td_(who_(l)) + td_(esc_(l.start) + "~" + esc_(l.end)) + td_(hh_(mins_(l)), ";text-align:right") + td_(esc_(l.type)) +
        td_(esc_((l.items || []).join(", ")), ";white-space:normal") + td_(esc_(l.detail || "-"), ";white-space:normal;color:#374151") + td_(chip_(l.status)) + '</tr>';
    });
    h += '</table>';

    // 작업구분별 누적 MH (착수일부터 보고일까지, 해당 프로젝트 전체 일지) - 카드 아래쪽, 표 + 원형 그래프
    var cum = {}, order = [], first = "9999", total = 0;
    all.forEach(function (l) {
      if (projKey_(l) !== k || l.date > d) return;
      var t = l.type || "미분류";
      if (!cum[t]) { cum[t] = { m: 0, day: 0, dates: {}, who: {} }; order.push(t); }
      var m = mins_(l);
      cum[t].m += m; total += m;
      cum[t].dates[l.date] = 1; cum[t].who[l.writer] = 1;
      if (l.date === d) cum[t].day += m;
      if (l.date < first) first = l.date;
    });
    order.sort(function (a2, b2) { return cum[b2].m - cum[a2].m; });
    var allWho = {}, dayTot = 0;
    var cid = "pie" + (++pieN), pie = makePie_(cid, order.map(function (t) { return [t, cum[t].m]; }));
    if (pie) images[cid] = pie;

    h += '<div style="padding:8px 12px;border-top:1px solid #d1d5db;background:#fcfcfd"><div style="font-size:12px;color:#6b7280;margin-bottom:4px">작업구분별 누적 MH · ' + d + ' 기준 (최초 일지 ' + first + ' ~)</div>' +
      '<table style="width:100%;border-collapse:collapse"><tr><td style="vertical-align:top;padding:0">' +
      '<table style="width:100%;border-collapse:collapse"><tr>' + th_("작업구분", ";width:140px") + th_("누적 MH", ";text-align:right;width:100px") + th_("비중", ";text-align:right;width:70px") +
      th_("전일 MH", ";text-align:right;width:90px") + th_("작업일수", ";text-align:right;width:80px") + th_("참여인원", ";text-align:right;width:80px") + '</tr>';
    order.forEach(function (t, i) {
      var c = cum[t], pct = total ? Math.round(c.m / total * 100) : 0;
      Object.keys(c.who).forEach(function (w) { allWho[w] = 1; });
      dayTot += c.day;
      h += '<tr>' + td_('<span style="display:inline-block;width:10px;height:10px;border-radius:5px;background:' + PIE_COLORS[i % PIE_COLORS.length] + ';margin-right:6px"></span>' + esc_(t)) +
        td_("<b>" + mh_(c.m) + "</b>", ";text-align:right") + td_(pct + "%", ";text-align:right") + td_(c.day ? "+" + mh_(c.day) : "-", ";text-align:right;color:#1e40af") +
        td_(Object.keys(c.dates).length + "일", ";text-align:right") + td_(Object.keys(c.who).length + "명", ";text-align:right") + '</tr>';
    });
    h += '<tr>' + td_("<b>합계</b>", ";border-bottom:0") + td_("<b>" + mh_(total) + "</b>", ";text-align:right;border-bottom:0") + td_("100%", ";text-align:right;border-bottom:0") +
      td_("<b>+" + mh_(dayTot) + "</b>", ";text-align:right;color:#1e40af;border-bottom:0") + td_("-", ";text-align:right;border-bottom:0") + td_(Object.keys(allWho).length + "명", ";text-align:right;border-bottom:0") + '</tr></table>' +
      '</td>' + (pie ? '<td style="width:300px;vertical-align:middle;text-align:center;padding:0 0 0 12px"><img src="cid:' + cid + '" width="280" height="200" alt="작업구분별 비중" style="display:block"></td>' : '') + '</tr></table></div>';
    h += '</div>';
  });

  // 소속별 작업시간 (맨 아래)
  var g = {}, gk = [];
  day.forEach(function (l) {
    var key = isExt_(l) ? "외주|" + (l.vendor || "-") : "직영|-";
    if (!g[key]) { g[key] = { w: {}, n: 0, m: 0 }; gk.push(key); }
    g[key].w[l.writer] = 1; g[key].n++; g[key].m += mins_(l);
  });
  gk.sort(function (a, b) { return (a.charAt(0) === "직" ? 0 : 1) - (b.charAt(0) === "직" ? 0 : 1) || a.localeCompare(b); });
  h += sec_("소속별 작업시간") + '<table style="width:100%;border-collapse:collapse"><tr>' + th_("구분", ";width:70px") + th_("업체", ";width:130px") + th_("인원", ";text-align:right;width:70px") +
    th_("일지", ";text-align:right;width:70px") + th_("작업시간", ";text-align:right;width:90px") + th_("작업자") + '</tr>';
  gk.forEach(function (k2) {
    var q = k2.split("|"), o = g[k2];
    h += '<tr>' + td_(q[0]) + td_(esc_(q[1])) + td_(Object.keys(o.w).length + "명", ";text-align:right") + td_(o.n + "건", ";text-align:right") +
      td_(hh_(o.m), ";text-align:right") + td_(esc_(Object.keys(o.w).join(", ")), ";white-space:normal") + '</tr>';
  });
  h += '<tr>' + td_("<b>합계</b>", ";border-bottom:0") + td_("", ";border-bottom:0") + td_("<b>" + writers.length + "명</b>", ";text-align:right;border-bottom:0") + td_("<b>" + day.length + "건</b>", ";text-align:right;border-bottom:0") +
    td_("<b>" + hh_(totMin) + "</b>", ";text-align:right;border-bottom:0") + td_("", ";border-bottom:0") + '</tr></table>';

  var url = ""; try { url = SpreadsheetApp.getActiveSpreadsheet().getUrl(); } catch (e2) {}
  h += '<div style="font-size:12.5px;color:#6b7280;border-top:1px solid #e5e7eb;padding-top:10px;margin-top:18px">시수와 MH는 시작~종료 시간 기준 인시(쉬는 시간 포함)입니다. 받는 사람은 시트의 \'' + REPORT_TAB + '\' 탭에서 바꿀 수 있습니다.' + (url ? ' · <a href="' + url + '" style="color:#1e40af">시트 열기</a>' : '') + '</div>';
  h += '</div>';
  return { subject: subject, html: h, images: images };
}
