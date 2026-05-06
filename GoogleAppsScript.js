// Google Apps Script 코드
// Google Sheets 메뉴 > 확장 프로그램 > Apps Script에서 이 코드를 모두 복사하세요
// SHEET_ID를 입력할 필요 없습니다. HTML에서 자동으로 전달됩니다.

function doGet(e) {
  const action = e.parameter.action || 'getData';
  const sheetId = e.parameter.sheetId;

  if (action === 'getData') {
    return ContentService.createTextOutput(JSON.stringify(getData(sheetId)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput('Invalid action');
}

function doPost(e) {
  const action = e.parameter.action;
  const sheetId = e.parameter.sheetId;

  try {
    switch(action) {
      case 'updateAttendance':
        updateAttendance(sheetId, e.parameter.grade, e.parameter.className,
                        e.parameter.number, e.parameter.attendance);
        break;
    }
    return ContentService.createTextOutput('OK');
  } catch(error) {
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

function getData(sheetId) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName('데이터');
    
    if (!sheet) {
      return {};
    }

    const data = sheet.getDataRange().getValues();
    const result = {};

    // 첫 번째 행은 헤더 (학년, 반, 번호, 이름, 출석상태)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const grade = row[0];
      const className = row[1];
      const number = row[2];
      const name = row[3];
      let attendance = row[4];

      // 공백이면 0, 1은 1로
      if (attendance === 1 || attendance === '1') {
        attendance = 1;
      } else {
        attendance = 0;
      }

      if (!grade) continue;

      if (!result[grade]) result[grade] = {};
      if (!result[grade][className]) result[grade][className] = [];

      result[grade][className].push({
        number: number,
        name: name,
        attendance: attendance
      });
    }

    return result;
  } catch(error) {
    Logger.log('getData 오류: ' + error);
    return { error: error.toString() };
  }
}

function updateAttendance(sheetId, grade, className, number, attendance) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName('데이터');

    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === grade && data[i][1] === className && data[i][2] == number) {
        sheet.getRange(i + 1, 5).setValue(attendance); // 5번째 열 (출석상태)
        return;
      }
    }
  } catch(error) {
    Logger.log('updateAttendance 오류: ' + error);
  }
}
