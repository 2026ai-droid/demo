/**
 * 예시 B. 먼저 생각하고, AI 답과 비교하기 (따라 만들기 키트)
 *
 * AID 흐름의 Independence → Deepen 장면을 담은 활동지입니다.
 *  - 학생은 자기 생각을 먼저 저장해야 AI 참고 답변을 볼 수 있습니다.
 *  - AI 참고 답변은 교사가 미리 만들어 '설정' 시트에 붙여넣습니다.
 *    그래서 이 예시는 API 키나 사용료가 필요 없습니다.
 *  - 최초 생각, 최종 생각, 고친 이유가 함께 남아 과정 중심 평가의 근거가 됩니다.
 *
 * 쓰는 법: 예시 A와 같습니다. setup 을 한 번 실행한 뒤 '설정' 시트에서 질문을 바꾸고 배포하세요.
 */

const SET = '설정';
const RES = '응답';
const SET_HEAD = ['질문 번호', '질문', 'AI 참고 답변', '공개(O/X)'];
const RES_HEAD = ['처음 저장', '반', '번호', '별칭', '질문 번호', '최초 생각', '최종 생각', '선택', '고친 이유', '마지막 저장'];
const CHOICES = ['그대로 둠(KEEP)', '일부 가져와 씀(USE)', '내 생각을 고침(CHANGE)'];

/** 처음 한 번만 실행: 설정·응답 시트와 예시 질문을 만듭니다. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let set = ss.getSheetByName(SET);
  if (!set) {
    set = ss.insertSheet(SET);
    set.appendRow(SET_HEAD);
    set.appendRow(['Q1',
      '압력이 일정할 때 기체의 온도를 높이면 부피는 어떻게 될까요? 이유와 함께 예측해 보세요.',
      '기체의 부피는 온도에 비례합니다. 그래서 온도가 10℃에서 20℃로 2배가 되면 부피도 항상 2배가 됩니다.',
      'O']);
    set.setFrozenRows(1);
    set.getRange(1, 1, 1, SET_HEAD.length).setFontWeight('bold').setBackground('#eef2f8');
  }
  let res = ss.getSheetByName(RES);
  if (!res) {
    res = ss.insertSheet(RES);
    res.appendRow(RES_HEAD);
    res.setFrozenRows(1);
    res.getRange(1, 1, 1, RES_HEAD.length).setFontWeight('bold').setBackground('#eef2f8');
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('먼저 생각하고 AI와 비교하기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 반·번호는 시트가 숫자로 바꿔 저장하므로 앞자리 0을 떼어 맞춥니다. */
function num_(v) {
  return clean_(v, 10).replace(/^0+(?=\d)/, '');
}

function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max || 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function settings_() {
  return SpreadsheetApp.getActive().getSheetByName(SET).getDataRange().getValues().slice(1)
    .filter(v => v[0] && String(v[3]).trim().toUpperCase() === 'O');
}

/** 공개된 질문 목록 (AI 답변은 보내지 않습니다) */
function getQuestions() {
  return settings_().map(v => ({ id: String(v[0]), text: String(v[1]) }));
}

function findRow_(sh, who, qid) {
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][1]) === who.cls && String(vals[i][2]) === who.no &&
        String(vals[i][3]) === who.nick && String(vals[i][4]) === qid) return { row: i + 1, v: vals[i] };
  }
  return null;
}

/** 1단계: 최초 생각을 저장해야 AI 참고 답변을 돌려줍니다. 한 번 저장한 최초 생각은 바뀌지 않습니다. */
function saveFirst(p) {
  p = p || {};
  const who = { cls: num_(p.cls), no: num_(p.no), nick: clean_(p.nick, 20) };
  const qid = clean_(p.qid, 20), first = clean_(p.first);
  if (!who.cls || !who.no || !who.nick) throw new Error('반, 번호, 별칭을 모두 입력하세요.');
  if (first.length < 10) throw new Error('내 생각을 조금 더 자세히 써 주세요(10자 이상).');
  const q = settings_().find(v => String(v[0]) === qid);
  if (!q) throw new Error('지금 열려 있는 질문이 아니에요.');

  const sh = SpreadsheetApp.getActive().getSheetByName(RES);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const found = findRow_(sh, who, qid);
    if (!found) sh.appendRow([new Date(), who.cls, who.no, who.nick, qid, first, '', '', '', '']);
    return { ai: String(q[2]), first: found ? String(found.v[5]) : first };
  } finally {
    lock.releaseLock();
  }
}

/** 2단계: 비교한 뒤 최종 생각과 고친 이유를 저장합니다. */
function saveFinal(p) {
  p = p || {};
  const who = { cls: num_(p.cls), no: num_(p.no), nick: clean_(p.nick, 20) };
  const qid = clean_(p.qid, 20);
  if (CHOICES.indexOf(p.choice) < 0) throw new Error('세 가지 중 하나를 골라 주세요.');
  if (clean_(p.reason).length < 5) throw new Error('왜 그렇게 정했는지 이유를 써 주세요.');
  const sh = SpreadsheetApp.getActive().getSheetByName(RES);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const found = findRow_(sh, who, qid);
    if (!found) throw new Error('먼저 내 생각을 저장해 주세요.');
    sh.getRange(found.row, 7, 1, 4).setValues([[clean_(p.final), p.choice, clean_(p.reason), new Date()]]);
  } finally {
    lock.releaseLock();
  }
  return true;
}
