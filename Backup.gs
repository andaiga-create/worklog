/* =====================================================================
   작업일지 시트 주간 자동 백업
   - 매주 일요일 새벽 3시에 이 스프레드시트 전체의 사본을 드라이브 '작업일지_백업' 폴더에 만든다
   - 최근 4개만 남기고 오래된 사본은 휴지통으로 보낸다 (휴지통에서 30일 뒤 자동 삭제)
   - 처음 한 번 setupBackup 을 ▶ 실행한다 (트리거 생성 + 첫 백업). 웹앱 재배포 불필요
   - 복구: 백업 폴더의 사본을 열어 필요한 탭을 원본으로 복사하면 된다
   - 사진 파일은 드라이브에 따로 있으므로 백업 대상이 아니다 (사본에도 사진 링크는 그대로 남는다)
   ===================================================================== */
var BACKUP_FOLDER = "작업일지_백업";
var BACKUP_KEEP = 4;
var BACKUP_PREFIX = "작업일지_백업_";

function setupBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "backupSheet") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("backupSheet").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).inTimezone("Asia/Seoul").create();
  backupSheet();
}

function backupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var it = DriveApp.getFoldersByName(BACKUP_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
  var name = BACKUP_PREFIX + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd_HHmm");
  DriveApp.getFileById(ss.getId()).makeCopy(name, folder);

  // 오래된 사본 정리: 이름이 BACKUP_PREFIX 로 시작하는 파일만, 날짜 이름순으로 최근 BACKUP_KEEP 개 유지
  var files = [], fi = folder.getFiles();
  while (fi.hasNext()) { var f = fi.next(); if (f.getName().indexOf(BACKUP_PREFIX) === 0) files.push(f); }
  files.sort(function (a, b) { return b.getName().localeCompare(a.getName()); });
  files.slice(BACKUP_KEEP).forEach(function (f) { f.setTrashed(true); });
  Logger.log("백업 완료: " + name + " (보관 " + Math.min(files.length, BACKUP_KEEP) + "개)");
}
