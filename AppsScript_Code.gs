/**
 * BRL Perio Chart - Google Apps Script 수신용 웹앱
 * --------------------------------------------------
 * 역할: 차트에서 보낸 데이터를 받아 지정한 구글 시트의
 *       회차별 탭(Baseline / Mid-point / End-point)에 한 줄씩 누적한다.
 *
 * 설정 순서 (한 번만):
 *  1) 공유 구글 계정으로 새 구글 시트 생성 → 주소창 URL에서 시트 ID 복사
 *     예) docs.google.com/spreadsheets/d/[이 부분이 ID]/edit
 *  2) 아래 SHEET_ID 에 붙여넣기
 *  3) 시트 메뉴: 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기 → 저장
 *  4) 배포 → 새 배포 → 유형: 웹 앱
 *       - 실행 사용자: 나(공유 계정)
 *       - 액세스 권한: 모든 사용자
 *     → 배포 → 웹앱 URL 복사 (이 URL을 차트에 넣습니다)
 */

// ===== 여기만 본인 시트 ID로 바꾸세요 =====
const SHEET_ID = '여기에_구글시트_ID_붙여넣기';

// (선택) 간단한 보안용 키. 차트와 동일하게 맞추면 외부 무단 전송을 막습니다.
// 빈 문자열('')로 두면 키 검사를 하지 않습니다.
const SECRET = '';

// 회차값 → 탭 이름 매핑
const VISIT_TABS = {
  baseline: 'Baseline',
  mid: 'Mid-point',
  end: 'End-point'
};

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (SECRET && body.secret !== SECRET) {
      return json({ ok: false, error: 'invalid secret' });
    }

    const visit = (body.visit || 'baseline').toLowerCase();
    const tabName = VISIT_TABS[visit] || 'Baseline';
    const headers = body.headers;   // 열 이름 배열 (ID, Visit, ExamDate, GR_..., ...)
    const row = body.row;           // 값 배열 (headers와 같은 순서)

    if (!Array.isArray(headers) || !Array.isArray(row)) {
      return json({ ok: false, error: 'headers/row missing' });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) sheet = ss.insertSheet(tabName);

    // 탭이 비어 있으면 헤더(+ 저장시각) 먼저 기록
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['SavedAt'].concat(headers));
      sheet.setFrozenRows(1);
    }

    // 같은 ID가 이미 있으면 그 줄을 덮어쓰고, 없으면 새 줄 추가
    const idCol = 2; // SavedAt(1), ID(2)
    const id = String(row[0]);
    const savedAt = new Date();
    const fullRow = [savedAt].concat(row);

    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow > 1) {
      const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === id) { targetRow = i + 2; break; }
      }
    }

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, fullRow.length).setValues([fullRow]);
      return json({ ok: true, action: 'updated', tab: tabName, row: targetRow });
    } else {
      sheet.appendRow(fullRow);
      return json({ ok: true, action: 'appended', tab: tabName, row: sheet.getLastRow() });
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, message: 'BRL Perio Chart endpoint is running.' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
