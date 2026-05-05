// Google Apps Script 코드 (확장 프로그램 → Apps Script에 붙여넣기)
// 배포: [배포] → [새 배포] → 유형: 웹 앱 → 액세스: 모든 사용자

const SPREADSHEET_ID = "1CLfo4ck7LOVhoszNnArqCe_5qoKLNyx-z-Soa755X20";

function doGet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();

  const result = {
    sheets: sheets
      .map(function(sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return null;

        // 헤더 행 감지 (번호/단어/의미)
        let startRow = 0;
        const firstRow = data[0].map(function(c) { return String(c).toLowerCase().trim(); });
        const hasHeader = firstRow.some(function(c) {
          return /word|단어|의미|meaning|번호|no/.test(c);
        });
        if (hasHeader) startRow = 1;

        const words = data.slice(startRow)
          .filter(function(row) { return row[1] && String(row[1]).trim(); })
          .map(function(row) {
            const sec = row[0];
            return {
              s: sec === "" || sec === null ? 1
                 : isNaN(Number(sec)) ? String(sec).trim()
                 : Number(sec),
              w: String(row[1]).trim(),
              m: String(row[2] || "").trim()
            };
          });

        return { name: sheet.getName(), words: words };
      })
      .filter(function(tab) { return tab && tab.words.length > 0; })
  };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
