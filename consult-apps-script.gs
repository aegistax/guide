/**
 * 이지스 고객 가이드 - 상담 신청 접수용 Apps Script
 * 시트에 연결되지 않은 독립 프로젝트에서도 동작합니다(시트는 자동 생성).
 */
var TO = 'mail@aegistax.kr,ctangch@gmail.com';
var RETENTION_YEARS = 2;      // 보유 기간(년) - 개인정보 동의서 문구와 일치
var PURGE_GMAIL = false;      // true로 바꾸면 이 계정 Gmail의 오래된 "[상담 신청]" 메일도 휴지통으로 이동(추가 권한 필요)
var METHODS = ['전화상담', '방문상담', '카카오톡 상담'];
var TYPES = ['기장/신고대리', '법인설립/전환', '법인전환', '양도소득세', '상속·증여세', '가족간 부동산 거래'];  // '법인전환'은 이전 페이지 호환용, '기타: 직접입력'은 별도 허용

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var d = JSON.parse(e.postData.contents);
    if (d.website) return out_();
    var name = clean_(d.name, 30), phone = clean_(d.phone, 20),
        type = clean_(d.type, 60), message = clean_(d.message, 2000),
        method = METHODS.indexOf(clean_(d.method, 20)) >= 0 ? clean_(d.method, 20) : '미선택';   // 이전 페이지 호환
    if (!name || !/^[0-9\-]{9,13}$/.test(phone) || (TYPES.indexOf(type) < 0 && !/^기타: .+/.test(type)) ||
        message.length < 5 || d.agree !== true) return out_();

    // 1) 이메일 먼저 발송
    MailApp.sendEmail({
      to: TO,
      name: '이지스 고객 가이드',
      subject: '[상담 신청] ' + type + ' / ' + method + ' - ' + name,
      body: '새 상담 신청이 접수되었습니다.\n\n' +
            '성함: ' + name + '\n연락처: ' + phone + '\n상담 유형: ' + type + '\n상담 방식: ' + method +
            (method === '방문상담' ? '\n※ 방문 상담 - 일정 조율을 위해 안내 전화 필요' : '') + '\n\n문의 내용:\n' + message
    });

    // 2) 시트 기록 (실패해도 메일에는 영향 없음)
    try {
      var sh = sheet_();
      if (sh.getLastRow() === 0) sh.appendRow(['접수일시', '성함', '연락처', '상담 유형', '상담 방식', '문의 내용']);
      sh.appendRow([new Date(), safe_(name), safe_(phone), safe_(type), method, safe_(message)]);
    } catch (e2) { console.error(e2); }
  } catch (err) {
    console.error(err);
  } finally {
    lock.releaseLock();
  }
  return out_();
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('SHEET_ID');
    if (id) { ss = SpreadsheetApp.openById(id); }
    else { ss = SpreadsheetApp.create('이지스 상담접수'); props.setProperty('SHEET_ID', ss.getId()); }
  }
  return ss.getSheetByName('상담접수') || ss.insertSheet('상담접수');
}

function clean_(v, max) { return String(v || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max); }
function safe_(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function out_() { return ContentService.createTextOutput('ok'); }

// 권한 승인용: 편집기에서 이 함수를 한 번 실행하면 메일·시트 권한 승인 창이 뜹니다.
function authorize() {
  MailApp.getRemainingDailyQuota();
  sheet_();
}


/* ===== 보유 기간 경과 자료 자동 삭제 ===== */

// 매일 새벽 자동 실행되는 함수: 접수일시가 보유 기간(2년)을 넘은 시트 행을 삭제합니다.
function purgeOld() {
  var cutoff = new Date();
  cutoff.setFullYear(cutoff.getFullYear() - RETENTION_YEARS);

  var sh = sheet_();
  var last = sh.getLastRow();
  var removed = 0;
  if (last >= 2) {
    var dates = sh.getRange(2, 1, last - 1, 1).getValues();   // A열(접수일시), 머리글 제외
    for (var i = dates.length - 1; i >= 0; i--) {              // 아래에서 위로 삭제해야 행 번호가 밀리지 않음
      var d = dates[i][0];
      if (d instanceof Date && d < cutoff) { sh.deleteRow(i + 2); removed++; }
    }
  }

  var trashed = 0;
  if (PURGE_GMAIL) {
    var threads = GmailApp.search('subject:"[상담 신청]" older_than:' + (RETENTION_YEARS * 365) + 'd', 0, 100);
    threads.forEach(function (th) { th.moveToTrash(); trashed++; });
  }
  console.log('삭제된 시트 행: ' + removed + ' / 휴지통 이동 메일: ' + trashed);
}

// 편집기에서 이 함수를 한 번만 실행하면 매일 03시경 purgeOld가 자동 실행되도록 등록됩니다.
function installPurgeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'purgeOld') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('purgeOld').timeBased().everyDays(1).atHour(3).create();
}
