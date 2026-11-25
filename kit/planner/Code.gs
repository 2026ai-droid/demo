/**
 * 예시 A. 계획–집중–성찰 미니 플래너 (따라 만들기 키트)
 *
 * 쓰는 법
 *  1) 새 Google 스프레드시트 > 확장 프로그램 > Apps Script 를 엽니다.
 *  2) 이 파일 내용을 Code.gs 에, Index.html 내용을 새 HTML 파일(이름: Index)에 붙여넣습니다.
 *  3) 위쪽 함수 목록에서 setup 을 골라 한 번 실행하고 권한을 승인합니다.
 *  4) 배포 > 새 배포 > 웹 앱 (실행 사용자: 나 / 액세스: 우리 학교 도메인 사용자) 로 배포합니다.
 *
 * 개인정보 원칙
 *  - 이름 대신 반·번호·별칭만 받습니다. 별칭은 학생이 정한 간단한 말이면 됩니다.
 *  - 비밀번호, API 키, 다른 문서의 ID를 코드에 직접 쓰지 마세요.
 *  - 학년이 끝나면 '기록' 시트를 내려받아 보관 기준에 따라 정리하세요.
 */

const SHEET = '기록';
const SUMMARY = '반별 요약';
const HEAD = ['저장 시각', '반', '번호', '별칭', '오늘의 계획', '목표', '집중(분)', '잘된 점', '어려웠던 점', '다음 계획'];

/** 처음 한 번만 실행: 기록 시트와 반별 요약 시트를 만듭니다. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#eef2f8');
  }
  const sum = ss.getSheetByName(SUMMARY) || ss.insertSheet(SUMMARY);
  sum.clear();
  sum.getRange('A1').setFormula(
    `=QUERY('${SHEET}'!A:J,"select B, count(A), sum(G) where B is not null group by B ` +
    `label count(A) '기록 수', sum(G) '집중(분) 합계'",1)`
  );
}

/** 웹 앱 화면 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('계획–집중–성찰 플래너')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 입력값 정리: 길이 제한 + 수식으로 해석되지 않도록 처리 */
/** 반·번호는 시트가 숫자로 바꿔 저장하므로 앞자리 0을 떼어 맞춥니다. */
function num_(v) {
  return clean_(v, 10).replace(/^0+(?=\d)/, '');
}

function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max || 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

/** 학생 기록 저장 후, 그 학생의 최근 기록을 돌려줍니다. */
function saveRecord(r) {
  r = r || {};
  const cls = num_(r.cls), no = num_(r.no), nick = clean_(r.nick, 20);
  if (!cls || !no || !nick) throw new Error('반, 번호, 별칭을 모두 입력하세요.');
  if (!clean_(r.plan)) throw new Error('오늘의 계획을 먼저 적어 주세요.');
  const minutes = Math.max(0, Math.min(600, Math.round(Number(r.minutes) || 0)));

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActive().getSheetByName(SHEET).appendRow([
      new Date(), cls, no, nick, clean_(r.plan), clean_(r.goal), minutes,
      clean_(r.good), clean_(r.hard), clean_(r.next)
    ]);
  } finally {
    lock.releaseLock();
  }
  return getMyRecords(cls, no, nick);
}

/** 반·번호·별칭이 모두 같은 기록만 최근 5개 돌려줍니다. */
function getMyRecords(cls, no, nick) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  const tz = Session.getScriptTimeZone();
  return sh.getDataRange().getValues().slice(1)
    .filter(v => String(v[1]) === num_(cls) && String(v[2]) === num_(no) && String(v[3]) === String(nick))
    .slice(-5).reverse()
    .map(v => ({
      at: Utilities.formatDate(new Date(v[0]), tz, 'M/d HH:mm'),
      plan: String(v[4]), minutes: Number(v[6]) || 0, next: String(v[9])
    }));
}
