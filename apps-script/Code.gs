// 관리자 AI 실습 키트 - 참여 결과 저장용 Apps Script
// 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣고 웹 앱으로 배포하세요.

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = d.type === 'task' ? '해볼업무' : '퀴즈';
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(name === '퀴즈'
        ? ['시각', '이름', '문항', '답', '정답여부', '회차']
        : ['시각', '이름', '내용', '회차']);
    }
    const who = String(d.name || '').slice(0, 40);
    if (d.type === 'task') {
      sh.appendRow([new Date(), who, String(d.text || '').slice(0, 1000), d.session || '']);
    } else {
      sh.appendRow([new Date(), who, Number(d.q) || 0, d.a === 'O' ? 'O' : 'X', d.ok === true, d.session || '']);
    }
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const session = (e && e.parameter && e.parameter.session) || '';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const read = n => {
    const sh = ss.getSheetByName(n);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  };
  const quiz = read('퀴즈')
    .filter(r => !session || r[5] === session)
    .map(r => ({ t: new Date(r[0]).getTime(), name: r[1], q: r[2], a: r[3], ok: r[4] === true || r[4] === 'TRUE' }));
  const tasks = read('해볼업무')
    .filter(r => !session || r[3] === session)
    .map(r => ({ t: new Date(r[0]).getTime(), name: r[1], text: r[2] }));
  return out({ quiz: quiz, tasks: tasks });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
