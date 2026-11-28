/* =====================================================================
   '현황' 탭: 프로젝트 × 작업구분 × 날짜별 시수(MH)
   - 처음 한 번 setupDashboard 를 ▶ 실행 → '현황', '공휴일' 탭 생성
   - '현황' 탭의 [새로고침] 체크박스를 누르거나, 상단 메뉴 '작업일지 > 현황 새로고침'
   - 시작일/종료일은 B1, B2 에 입력 (비우면 최근 2주 / 오늘)
   - 날짜는 최근 날짜가 가장 오른쪽에, 일지가 없는 날과 작업구분은 표시하지 않는다
   - 작업구분 순서는 '작업항목' 탭 A열에 처음 나오는 순서로 고정
   - 시수는 식사시간(12:00~13:00, 17:30~18:00)을 뺀 값
   - 이 파일은 Code.gs 와 같은 프로젝트에 새 파일로 추가한다 (웹앱 재배포 불필요)
   ===================================================================== */
var DB_TAB = "현황";
var DB_HOL_TAB = "공휴일";
var DB_TZ = "Asia/Seoul";
var DB_HEAD_ROW = 4;      // 표 머리글 행
var DB_BREAKS = [[720, 780], [1050, 1080]];

// '공휴일' 탭을 처음 만들 때 넣는 목록. 임시공휴일·회사 휴무일은 탭에서 직접 추가/삭제한다.
var DB_HOL_SEED = [
  ["2026-10-05", "대체공휴일(개천절)"], ["2026-10-09", "한글날"], ["2026-12-25", "성탄절"],
  ["2027-01-01", "신정"], ["2027-02-06", "설날 연휴"], ["2027-02-07", "설날"], ["2027-02-08", "설날 연휴"], ["2027-02-09", "대체공휴일(설날)"],
  ["2027-03-01", "삼일절"], ["2027-05-05", "어린이날"], ["2027-05-13", "부처님오신날"], ["2027-06-06", "현충일"],
  ["2027-08-15", "광복절"], ["2027-08-16", "대체공휴일(광복절)"], ["2027-09-14", "추석 연휴"], ["2027-09-15", "추석"], ["2027-09-16", "추석 연휴"],
  ["2027-10-03", "개천절"], ["2027-10-04", "대체공휴일(개천절)"], ["2027-10-09", "한글날"], ["2027-10-11", "대체공휴일(한글날)"],
  ["2027-12-25", "성탄절"], ["2027-12-27", "대체공휴일(성탄절)"]
];

/* ---------- 설치 / 실행 ---------- */
function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(DB_HOL_TAB)) {
    var h = ss.insertSheet(DB_HOL_TAB);
    h.getRange(1, 1, 1, 2).setValues([["날짜", "이름 (임시공휴일·회사 휴무일은 여기에 추가)"]]).setFontWeight("bold");
    var r = h.getRange(2, 1, DB_HOL_SEED.length, 2);
    r.setNumberFormat("@");
    r.setValues(DB_HOL_SEED);
    h.setColumnWidth(1, 110); h.setColumnWidth(2, 300);
  }
  getDashSheet_();
  refreshDashboard();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("작업일지").addItem("현황 새로고침", "refreshDashboard").addToUi();
}

// '현황' 탭의 새로고침 체크박스를 누르면 갱신한다
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== DB_TAB || e.range.getA1Notation() !== "D1" || e.range.getValue() !== true) return;
  try { refreshDashboard(); } finally { e.range.setValue(false); }
}

function getDashSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DB_TAB);
  if (!sh) {
    sh = ss.insertSheet(DB_TAB, 0);
    sh.getRange("A1:A2").setValues([["시작일 (비우면 최근 2주)"], ["종료일 (비우면 오늘)"]]).setFontWeight("bold");
    sh.getRange("B1:B2").setBackground("#fff2cc").setNumberFormat("yyyy-mm-dd").setHorizontalAlignment("center");
    sh.getRange("C1").setValue("새로고침 ▶").setFontWeight("bold").setHorizontalAlignment("right");
    sh.getRange("D1").insertCheckboxes().setBackground("#d2e3fc");
  }
  return sh;
}

/* ---------- 계산 도우미 ---------- */
function dbMins_(l) {
  var a = String(l.start).split(":"), b = String(l.end).split(":");
  var s = Number(a[0]) * 60 + Number(a[1]), e = Number(b[0]) * 60 + Number(b[1]);
  if (!(e > s)) return 0;
  var m = e - s;
  DB_BREAKS.forEach(function (x) { m -= Math.max(0, Math.min(e, x[1]) - Math.max(s, x[0])); });
  return m > 0 ? m : 0;
}
function dbH_(m) { return m ? Math.round(m / 60 * 100) / 100 : ""; }
function dbDate_(v, tz) {
  if (Object.prototype.toString.call(v) === "[object Date]") return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  var m = String(v || "").trim().match(/^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  return m ? m[1] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[3]).slice(-2) : "";
}
function dbShift_(d, days) {
  var p = d.split("-");
  var x = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + days));
  return x.getUTCFullYear() + "-" + ("0" + (x.getUTCMonth() + 1)).slice(-2) + "-" + ("0" + x.getUTCDate()).slice(-2);
}
function dbDow_(d) {
  var p = d.split("-");
  return new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]))).getUTCDay();
}
function dbHeat_(m) {
  var h = m / 60;
  return h <= 0 ? null : h <= 4 ? "#e8f0fe" : h <= 8 ? "#d2e3fc" : h <= 16 ? "#aecbfa" : "#8ab4f8";
}
function dbStatusColor_(s) { return /^보류/.test(s) ? "#fde68a" : /^완료/.test(s) ? "#d9ead3" : s ? "#d2e3fc" : null; }

/* ---------- 표 만들기 (시트에 쓰기 전의 값/색/메모 계산) ---------- */
function dbBuild_(logs, typeOrder, parts, holidays, start, end) {
  var key = function (l) { return l.yard + "-" + l.hull + "-" + l.equip; };
  var typeOf = function (l) { return l.type || "미분류"; };
  var upto = logs.filter(function (l) { return l.date && l.date <= end; });
  var inRange = upto.filter(function (l) { return l.date >= start; });

  var dates = [], projects = [];
  inRange.forEach(function (l) {
    if (dates.indexOf(l.date) < 0) dates.push(l.date);
    if (projects.indexOf(key(l)) < 0) projects.push(key(l));
  });
  dates.sort();                    // 오래된 날짜부터, 최근 날짜가 가장 오른쪽
  projects.sort();

  // 작업구분 고정 순서: '작업항목' 탭 순서 → 탭에 없는 것(가나다순) → 미분류
  var order = typeOrder.slice(), extra = [];
  upto.forEach(function (l) { var t = typeOf(l); if (t !== "미분류" && order.indexOf(t) < 0 && extra.indexOf(t) < 0) extra.push(t); });
  order = order.concat(extra.sort(), ["미분류"]);

  var nCols = 3 + dates.length + 3;
  var values = [], bg = [], bold = [], notes = [], merges = [];
  function blank(v) { var a = []; for (var i = 0; i < nCols; i++) a.push(v); return a; }
  function push(row, color, weight, note) { values.push(row); bg.push(color); bold.push(weight); notes.push(note); }

  var head = ["프로젝트", "품명", "작업구분"].concat(dates.map(function (d) {
    return Number(d.slice(5, 7)) + "/" + Number(d.slice(8)) + "(" + "일월화수목금토".charAt(dbDow_(d)) + ")";
  }), ["기간 합계", "누적 MH", "최근 상태"]);
  var headColor = head.map(function (_, i) {
    var d = dates[i - 3];
    if (!d) return "#202124";
    if (holidays[d] !== undefined || dbDow_(d) === 0) return "#d93025";   // 일요일·공휴일: 적색
    return dbDow_(d) === 6 ? "#1a73e8" : "#202124";                      // 토요일: 청색
  });
  var headNote = head.map(function (_, i) { var d = dates[i - 3]; return d && holidays[d] !== undefined ? holidays[d] || "공휴일" : ""; });

  var dayTotal = {}, grand = 0, grandCum = 0;
  projects.forEach(function (p) {
    var pLogs = upto.filter(function (l) { return key(l) === p; });
    var types = order.filter(function (t) { return pLogs.some(function (l) { return typeOf(l) === t; }); });
    var first = values.length;
    var sub = {}, subPeriod = 0, subCum = 0;
    types.forEach(function (t, ti) {
      var tl = pLogs.filter(function (l) { return typeOf(l) === t; });
      var row = blank(""), color = blank(null), weight = blank("normal"), note = blank("");
      row[0] = ti === 0 ? p : ""; row[1] = ti === 0 ? (parts[p] || "") : ""; row[2] = t;
      var period = 0, cum = 0, latest = null;
      tl.forEach(function (l) {
        cum += dbMins_(l);
        if (!latest || (l.date + l.start) >= (latest.date + latest.start)) latest = l;
      });
      dates.forEach(function (d, di) {
        var dl = tl.filter(function (l) { return l.date === d; });
        if (!dl.length) return;
        var m = 0, hold = false, lines = [];
        dl.forEach(function (l) {
          var lm = dbMins_(l); m += lm;
          if (/^보류/.test(l.status)) hold = true;
          lines.push(l.writer + " " + l.start + "~" + l.end + " (" + dbH_(lm) + "h) · " + l.status +
            ((l.items && l.items.length) ? "\n  " + l.items.join(", ") : "") + (l.detail ? "\n  메모: " + l.detail : ""));
        });
        var c = 3 + di;
        row[c] = dbH_(m); color[c] = hold ? "#fde68a" : dbHeat_(m); note[c] = lines.join("\n");
        period += m; sub[d] = (sub[d] || 0) + m; dayTotal[d] = (dayTotal[d] || 0) + m;
      });
      row[nCols - 3] = dbH_(period) || 0; row[nCols - 2] = dbH_(cum) || 0;
      row[nCols - 1] = latest ? latest.status : "";
      color[nCols - 1] = dbStatusColor_(row[nCols - 1]);
      subPeriod += period; subCum += cum;
      push(row, color, weight, note);
    });
    var sr = blank(""), sc = blank("#f1f3f4"), sw = blank("bold"), sn = blank("");
    sr[2] = "소계";
    dates.forEach(function (d, di) { sr[3 + di] = dbH_(sub[d] || 0); });
    sr[nCols - 3] = dbH_(subPeriod) || 0; sr[nCols - 2] = dbH_(subCum) || 0;
    sc[0] = null; sc[1] = null;
    push(sr, sc, sw, sn);
    merges.push({ from: first, count: types.length + 1 });
    grand += subPeriod; grandCum += subCum;
  });

  if (projects.length) {
    var tr = blank(""), tc = blank("#e8eaed"), tw = blank("bold"), tn = blank("");
    tr[0] = "일 합계";
    dates.forEach(function (d, di) { tr[3 + di] = dbH_(dayTotal[d] || 0); });
    tr[nCols - 3] = dbH_(grand) || 0; tr[nCols - 2] = dbH_(grandCum) || 0;
    push(tr, tc, tw, tn);
  }
  return { nCols: nCols, nDates: dates.length, head: head, headColor: headColor, headNote: headNote,
    values: values, bg: bg, bold: bold, notes: notes, merges: merges, projects: projects.length };
}

/* ---------- 시트에 쓰기 ---------- */
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getDashSheet_();
  var tz = ss.getSpreadsheetTimeZone();
  var today = Utilities.formatDate(new Date(), DB_TZ, "yyyy-MM-dd");
  var end = dbDate_(sh.getRange("B2").getValue(), tz) || today;
  var start = dbDate_(sh.getRange("B1").getValue(), tz) || dbShift_(end, -13);
  if (start > end) { var t = start; start = end; end = t; }

  var typeOrder = [];
  var it = ss.getSheetByName("작업항목");
  if (it && it.getLastRow() > 1) it.getRange(2, 1, it.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) {
    var v = String(r[0]).trim(); if (v && typeOrder.indexOf(v) < 0) typeOrder.push(v);
  });
  var parts = {};
  try { getMaster_().projects.forEach(function (p) { parts[p.yard + "-" + p.hull + "-" + p.name] = (p.parts || []).join(", "); }); } catch (e) {}
  var holidays = {};
  var hs = ss.getSheetByName(DB_HOL_TAB);
  if (hs && hs.getLastRow() > 1) hs.getRange(2, 1, hs.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var d = dbDate_(r[0], tz); if (d) holidays[d] = String(r[1]);
  });

  var t = dbBuild_(readLogs_(), typeOrder, parts, holidays, start, end);

  // 이전 표 지우기 (1~2행의 입력칸은 그대로)
  sh.setFrozenRows(0); sh.setFrozenColumns(0);
  var needRows = DB_HEAD_ROW + Math.max(t.values.length, 1);
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());
  if (sh.getMaxColumns() < t.nCols) sh.insertColumnsAfter(sh.getMaxColumns(), t.nCols - sh.getMaxColumns());
  var area = sh.getRange(3, 1, sh.getMaxRows() - 2, sh.getMaxColumns());
  area.breakApart(); area.clear(); area.clearNote();

  sh.getRange("A3").setValue("기간 " + start + " ~ " + end + " · 갱신 " + Utilities.formatDate(new Date(), DB_TZ, "yyyy-MM-dd HH:mm") +
    " · 시수는 식사시간 제외 · 칸에 마우스를 올리면 작업 내용").setFontColor("#5f6368").setFontSize(9);

  var hr = sh.getRange(DB_HEAD_ROW, 1, 1, t.nCols);
  hr.setValues([t.head]).setFontColors([t.headColor]).setNotes([t.headNote])
    .setFontWeight("bold").setBackground("#dadce0").setHorizontalAlignment("center");

  if (!t.projects) {
    sh.getRange(DB_HEAD_ROW + 1, 1).setValue("이 기간에는 작성된 일지가 없습니다.");
  } else {
    var body = sh.getRange(DB_HEAD_ROW + 1, 1, t.values.length, t.nCols);
    body.setValues(t.values).setBackgrounds(t.bg).setFontWeights(t.bold).setNotes(t.notes).setVerticalAlignment("middle");
    sh.getRange(DB_HEAD_ROW + 1, 3, t.values.length, t.nCols - 2).setHorizontalAlignment("center");
    sh.getRange(DB_HEAD_ROW, 1, t.values.length + 1, t.nCols).setBorder(true, true, true, true, true, true, "#c7c9cc", SpreadsheetApp.BorderStyle.SOLID);
    t.merges.forEach(function (m) {
      var top = DB_HEAD_ROW + 1 + m.from;
      if (m.count > 1) { sh.getRange(top, 1, m.count, 1).merge(); sh.getRange(top, 2, m.count, 1).merge(); }
      sh.getRange(top, 1, m.count, 2).setVerticalAlignment("middle").setWrap(true);
      sh.getRange(top, 1, m.count, 1).setFontWeight("bold");
      sh.getRange(top + m.count - 1, 1, 1, t.nCols).setBorder(null, null, true, null, null, null, "#5f6368", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    });
  }

  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 100);
  if (t.nDates) sh.setColumnWidths(4, t.nDates, 64);
  sh.setColumnWidths(4 + t.nDates, 2, 80); sh.setColumnWidth(6 + t.nDates, 130);
  sh.setFrozenRows(DB_HEAD_ROW); sh.setFrozenColumns(3);
  ss.toast("현황을 갱신했습니다. (" + start + " ~ " + end + ")");
}
