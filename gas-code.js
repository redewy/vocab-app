// Google Apps Script 코드 (확장 프로그램 → Apps Script에 붙여넣기)
// 배포: [배포] → [새 배포] → 유형: 웹 앱 → 액세스: 모든 사용자

const SPREADSHEET_ID = "1CLfo4ck7LOVhoszNnArqCe_5qoKLNyx-z-Soa755X20";

// 스프레드시트 컬럼 순서:
// 지문번호 | 영단어 | 한글뜻 | 유의어 | 반의어 | 영어오답1 | 영어오답2 | 영어오답3 | 영어오답4 | 한글오답1 | 한글오답2 | 한글오답3 | 한글오답4
// 유의어/반의어가 없는 단어는 해당 칸을 비워두면 됨. 여러 개면 쉼표(,)로 구분.

function doGet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();

  const result = {
    sheets: sheets
      .map(function(sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return null;

        // 헤더 행 감지
        const firstRow = data[0].map(function(c) { return String(c).toLowerCase().trim(); });
        const hasHeader = firstRow.some(function(c) {
          return /word|단어|영단어|의미|한글뜻|meaning|번호|지문번호|no/.test(c);
        });
        const startRow = hasHeader ? 1 : 0;

        // 헤더로 컬럼 인덱스 찾기 (헤더 없으면 고정 순서 사용)
        function colIdx(keywords) {
          for (var i = 0; i < firstRow.length; i++) {
            for (var k = 0; k < keywords.length; k++) {
              if (firstRow[i].includes(keywords[k])) return i;
            }
          }
          return -1;
        }

        var iSec = hasHeader ? colIdx(["지문번호","번호","no","section"]) : 0;
        var iW   = hasHeader ? colIdx(["영단어","word","단어"])           : 1;
        var iM   = hasHeader ? colIdx(["한글뜻","뜻","meaning","의미"])   : 2;
        var iSyn = hasHeader ? colIdx(["유의어","synonym"])              : 3;
        var iAnt = hasHeader ? colIdx(["반의어","antonym"])              : 4;
        var iEW  = hasHeader ? [colIdx(["영어오답1"]), colIdx(["영어오답2"]), colIdx(["영어오답3"]), colIdx(["영어오답4"])] : [-1,-1,-1,-1];
        var iKW  = hasHeader ? [colIdx(["한글오답1"]), colIdx(["한글오답2"]), colIdx(["한글오답3"]), colIdx(["한글오답4"])] : [-1,-1,-1,-1];

        if (iSec < 0) iSec = 0;
        if (iW   < 0) iW   = 1;
        if (iM   < 0) iM   = 2;

        function str(row, idx) {
          return idx >= 0 && row[idx] != null ? String(row[idx]).trim() : "";
        }

        // 유의어/반의어는 없을 수도 있고, 있으면 쉼표(,)로 여러 개 구분
        function list(row, idx) {
          var v = str(row, idx);
          return v ? v.split(/[,、\/]/).map(function(s) { return s.trim(); }).filter(Boolean) : [];
        }

        const words = data.slice(startRow)
          .filter(function(row) { return str(row, iW); })
          .map(function(row) {
            const sec = row[iSec];
            return {
              s: (sec === "" || sec === null) ? 1
                 : isNaN(Number(sec)) ? String(sec).trim()
                 : Number(sec),
              w:   str(row, iW),
              m:   str(row, iM),
              syn: list(row, iSyn),
              ant: list(row, iAnt),
              ew:  iEW.map(function(i) { return str(row, i); }).filter(Boolean),
              kw:  iKW.map(function(i) { return str(row, i); }).filter(Boolean),
              // 프론트에서도 컬럼명으로 접근할 수 있도록 원본 키도 포함
              "유의어": str(row, iSyn), "반의어": str(row, iAnt),
              "영어오답1": str(row, iEW[0]), "영어오답2": str(row, iEW[1]),
              "영어오답3": str(row, iEW[2]), "영어오답4": str(row, iEW[3]),
              "한글오답1": str(row, iKW[0]), "한글오답2": str(row, iKW[1]),
              "한글오답3": str(row, iKW[2]), "한글오답4": str(row, iKW[3]),
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
