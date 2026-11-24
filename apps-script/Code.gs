// 2026 안성 학교관리자 직무연수 - 참여 결과 저장용 Apps Script
// '2026 안성 학교관리자 직무연수 · 참여 결과' 구글 시트에서
// 확장 프로그램 > Apps Script 를 열고 이 코드를 통째로 붙여넣은 뒤 웹 앱으로 배포하세요.
//
// 시트 구성
//   실시간 점수판 : 참여 인원·평균·만점, 이름별 ○/× 와 순위 (수식으로 자동)
//   문항별 정답률 : 문항마다 응답 수·정답률과 막대 그래프 (수식으로 자동)
//   퀴즈기록      : 연수 페이지에서 답을 누를 때마다 한 줄씩 쌓임
//   해볼업무      : 마무리 장에서 '제출'을 누르면 쌓임
//   집계          : 점수판 계산용 (손대지 마세요)

const QUIZ_SHEET = '퀴즈기록';
const TASK_SHEET = '해볼업무';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const isTask = d.type === 'task';
    const sh = sheet_(isTask ? TASK_SHEET : QUIZ_SHEET);
    const who = String(d.name || '이름 없음').slice(0, 40);
    if (isTask) {
      sh.appendRow([new Date(), who, String(d.text || '').slice(0, 1000), d.session || '']);
    } else {
      sh.appendRow([new Date(), who, Number(d.q) || 0, d.a === 'O' ? 'O' : 'X', d.ok === true, d.session || '']);
    }
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// 발표자 화면(결과 버튼)이 읽어 가는 부분
function doGet(e) {
  const session = (e && e.parameter && e.parameter.session) || '';
  const read = n => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().filter(r => r[1] !== '');
  };
  const quiz = read(QUIZ_SHEET)
    .filter(r => !session || r[5] === session)
    .map(r => ({ t: new Date(r[0]).getTime(), name: r[1], q: r[2], a: r[3], ok: r[4] === true || r[4] === 'TRUE' }));
  const tasks = read(TASK_SHEET)
    .filter(r => !session || r[3] === session)
    .map(r => ({ t: new Date(r[0]).getTime(), name: r[1], text: r[2] }));
  return out_({ quiz: quiz, tasks: tasks });
}

// ▶ 리허설용: 가상 참여자 8명의 답을 넣어 점수판이 채워지는 모습을 미리 봅니다.
function 시연데이터넣기() {
  const sh = sheet_(QUIZ_SHEET);
  const answer = ['O', 'X', 'X', 'O', 'O'];
  const names = ['가온초 교감', '나래초 교감', '다솜초 교장', '라온초 교감', '마루초 교감', '바다초 교장', '사랑초 교감', '아름초 교감'];
  names.forEach((n, i) => {
    for (let q = 1; q <= 5; q++) {
      const ok = Math.random() < 0.55 + i * 0.05;
      const a = ok ? answer[q - 1] : (answer[q - 1] === 'O' ? 'X' : 'O');
      sh.appendRow([new Date(), n, q, a, ok, 'rehearsal']);
    }
  });
}

// ▶ 연수 시작 전: 퀴즈기록과 해볼업무를 모두 지웁니다. (점수판은 자동으로 비워져요)
function 기록지우기() {
  [QUIZ_SHEET, TASK_SHEET].forEach(n => {
    const sh = sheet_(n);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  });
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(name === QUIZ_SHEET
      ? ['시각', '이름', '문항', '답', '정답여부', '회차']
      : ['시각', '이름', '내용', '회차']);
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
