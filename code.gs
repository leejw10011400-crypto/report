/**
 * Google 스프레드시트 ID 설정
 */
const SPREADSHEET_ID = '1eRfRtOAkfyQVGa14o6YP6YQxVgQBZH27MtAUbrG4x-E';

/**
 * 웹 앱 접속 시 HTML 페이지 렌더링
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('보고서 수정 사항 기록기')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 시트 가져오기 및 초기 헤더 자동 설정
 */
function getTargetSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheets()[0]; // 첫 번째 시트 사용
  
  // 시트가 비어있으면 헤더 행 자동 생성
  if (sheet.getLastRow() === 0) {
    const headers = ['기록 일시', '수정 위치 (어느 부분)', '수정 내용 (뭘로 수정할지)', '추가 내용 (뭘 추가할지)', '처리 상태'];
    sheet.appendRow(headers);
    
    // 헤더 스타일링
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1e293b')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 새로운 보고서 수정 사항을 시트에 저장
 * @param {Object} formData - 클라이언트에서 전달받은 폼 데이터
 */
function saveRevision(formData) {
  try {
    const sheet = getTargetSheet();
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    
    const rowData = [
      now,
      formData.location || '-',
      formData.modification || '-',
      formData.addition || '-',
      '대기중'
    ];
    
    sheet.appendRow(rowData);
    
    // 가독성을 위해 열 너비 자동 맞춤 및 줄바꿈 설정
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 5).setWrap(true).setVerticalAlignment('top');
    
    return { success: true, message: '수정 사항이 성공적으로 기록되었습니다.' };
  } catch (error) {
    return { success: false, message: '저장 중 오류가 발생했습니다: ' + error.toString() };
  }
}

/**
 * 최근 기록된 수정 사항 목록 불러오기 (최신순 최대 20건)
 */
function getRecentRevisions() {
  try {
    const sheet = getTargetSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    // 최신 항목이 위로 오도록 역순 정렬 후 최대 20개 반환
    return data.reverse().slice(0, 20).map((row, index) => ({
      rowNumber: lastRow - index,
      timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : String(row[0]),
      location: String(row[1]),
      modification: String(row[2]),
      addition: String(row[3]),
      status: String(row[4] || '대기중')
    }));
  } catch (error) {
    throw new Error('데이터를 불러오는 데 실패했습니다: ' + error.toString());
  }
}

/**
 * 항목의 처리 상태(대기중 / 완료) 변경
 */
function toggleStatus(rowNumber, newStatus) {
  try {
    const sheet = getTargetSheet();
    sheet.getRange(rowNumber, 5).setValue(newStatus);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
