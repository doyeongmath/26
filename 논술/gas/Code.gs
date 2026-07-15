// =============================================
// 논술 독서 감상 기록 앱 - GAS 백엔드
// =============================================
// 스크립트 속성 설정 (프로젝트 설정 > 스크립트 속성):
//   ANTHROPIC_API_KEY : sk-ant-...
//   SPREADSHEET_ID    : 구글 시트 URL의 /d/XXXXXX/edit 에서 XXXXXX 부분

function getSheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  const ss = id
    ? SpreadsheetApp.openById(id)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트를 찾을 수 없습니다. 스크립트 속성에서 SPREADSHEET_ID를 설정해주세요.');
  return ss.getSheets()[0];
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('독서감상기록')
    .setTitle('논술 독서 감상 기록')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getStudentRecords(hakbun, name) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowHakbun = String(Math.floor(Number(row[2])));
    const rowName = row[3] ? row[3].toString().trim() : '';
    if (rowHakbun === String(hakbun).trim() && rowName === name.trim()) {
      records.push({
        timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Seoul', 'yyyy.MM.dd') : '',
        ban: row[1],
        hakbun: rowHakbun,
        name: rowName,
        bookName: row[4] ? row[4].toString() : '',
        author: row[5] ? row[5].toString() : '',
        reflection: row[6] ? row[6].toString() : '',
        memorableSentence: row[7] ? row[7].toString() : ''
      });
    }
  }

  return records;
}

function submitGibub(hakbun, name, ban, text) {
  if (!text || text.trim().length === 0) {
    return { success: false, error: '내용을 입력해주세요.' };
  }

  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  const ss = id
    ? SpreadsheetApp.openById(id)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트를 찾을 수 없습니다.');

  const SUBMIT_SHEET = '생기부_초안_제출';
  let sheet = ss.getSheetByName(SUBMIT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(SUBMIT_SHEET);
    sheet.appendRow(['제출시각', '반', '학번', '이름', '작성내용', '글자수']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#d9ead3');
  }

  const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const charCount = text.trim().length;
  sheet.appendRow([timestamp, ban, hakbun, name, text.trim(), charCount]);

  return { success: true };
}
