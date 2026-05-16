/**
 * 주간 계획 - Apps Script Backend (JSON API)
 *
 * 프론트엔드는 GitHub Pages에서 호스팅하고, 이 스크립트는 JSON API만 제공.
 *
 * 설치:
 *   1) 새 Google Sheet 생성 (이름 아무거나)
 *   2) "확장 프로그램 → Apps Script" 클릭
 *   3) 기본 Code.gs 내용을 모두 지우고 이 파일을 붙여넣기 → 💾 저장
 *   4) "배포 → 새 배포" → 톱니바퀴 → 유형: "웹 앱"
 *        - 다음 사용자로 실행: 본인
 *        - 액세스 권한: **모든 사용자** ← 외부(GitHub Pages)에서 호출하려면 필수
 *   5) 권한 승인 → 발급된 URL을 GitHub Pages 프론트엔드 index.html의
 *      WEB_APP_URL 상수에 붙여넣기
 *
 * 보안 주의:
 *   - URL을 아는 사람은 누구나 시트 읽기/쓰기 가능 (URL이 곧 비밀번호)
 *   - 개인용 데이터에 한해 사용
 *
 * 코드 수정 후 재배포: "배포 → 배포 관리 → 연필 아이콘 → 버전: 새 버전"
 *   (URL은 그대로 유지됨)
 */

const SHEET_NAME = 'Data';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let result;
  try {
    if (action === 'getAllData') {
      result = getAllData_();
    } else {
      result = { ok: true, hint: 'Weekly Planner API. GET ?action=getAllData or POST JSON.' };
    }
  } catch (err) {
    result = { error: String((err && err.message) || err) };
  }
  return json_(result);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'saveWeek') {
      result = { ok: true, ts: saveWeek_(body.weekKey, body.data) };
    } else if (action === 'replaceAllData') {
      result = { ok: true, ts: replaceAllData_(body.state) };
    } else {
      result = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { error: String((err && err.message) || err) };
  }
  return json_(result);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['week_key', 'data_json', 'updated_at']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 700);
    sheet.setColumnWidth(3, 170);
  }
  return sheet;
}

function getAllData_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { weeks: {}, lastModified: 0 };

  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const weeks = {};
  let lastModified = 0;

  for (const [key, json, updated] of rows) {
    if (!key || !json) continue;
    try {
      weeks[String(key)] = JSON.parse(json);
      const ts = updated instanceof Date ? updated.getTime() : 0;
      if (ts > lastModified) lastModified = ts;
    } catch (e) {
      console.error('parse error for', key, e);
    }
  }
  return { weeks, lastModified };
}

function saveWeek_(weekKey, weekRows) {
  if (!weekKey) throw new Error('weekKey required');
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const arr = Array.isArray(weekRows) ? weekRows : [];
  const now = new Date();

  let foundRow = -1;
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(weekKey)) { foundRow = i + 2; break; }
    }
  }

  if (arr.length === 0) {
    if (foundRow > 0) sheet.deleteRow(foundRow);
    return now.getTime();
  }

  const json = JSON.stringify(arr);
  if (foundRow > 0) {
    sheet.getRange(foundRow, 2, 1, 2).setValues([[json, now]]);
  } else {
    sheet.appendRow([weekKey, json, now]);
  }
  return now.getTime();
}

function replaceAllData_(state) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();

  const weeks = (state && state.weeks) || {};
  const now = new Date();
  const rows = [];
  for (const key in weeks) {
    if (!Array.isArray(weeks[key]) || weeks[key].length === 0) continue;
    rows.push([key, JSON.stringify(weeks[key]), now]);
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  return now.getTime();
}
