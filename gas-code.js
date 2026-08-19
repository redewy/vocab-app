// Google Apps Script 코드 (확장 프로그램 → Apps Script에 붙여넣기)
// 배포: [배포] → [새 배포] → 유형: 웹 앱 → 액세스: 모든 사용자

const SPREADSHEET_ID = "1CLfo4ck7LOVhoszNnArqCe_5qoKLNyx-z-Soa755X20";

// 스프레드시트 컬럼 순서:
// 지문번호 | 영단어 | 한글뜻 | 유의어 | 반의어 | 영어오답1 | 영어오답2 | 영어오답3 | 영어오답4 | 한글오답1 | 한글오답2 | 한글오답3 | 한글오답4
// 유의어/반의어가 없는 단어는 해당 칸을 비워두면 됨. 여러 개면 쉼표(,)로 구분.

// 단어장이 아닌 내부 데이터 시트 (doGet 단어 목록에서 제외)
const DATA_SHEETS = ["회원", "학습기록"];

function doGet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();

  const result = {
    sheets: sheets
      .map(function(sheet) {
        if (DATA_SHEETS.indexOf(sheet.getName()) !== -1) return null;
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

/* ───────────────── 회원 / 학습기록 (doPost) ─────────────────
 * 요청: fetch(URL, { method: "POST", body: JSON.stringify({ action, ... }) })
 * action:
 *  - signup    { id, pw, name }
 *  - login     { id, pw }                      → { ok, name, stars }
 *  - getUser   { id }                          → { ok, name, stars, role } (세션 복원용)
 *  - getDashboard { id }                       → { ok, members, records } (teacher 전용)
 *  - getRecords { id }                         → { ok, records } (본인 진도 조회)
 *  - saveStars { id, stars: ["word", ...] }
 *  - saveResult{ id, result: { tab, type, direction, score, total, wrong: [...], sections: [...] } }
 */

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 회원 시트: 아이디 | 비밀번호 | 이름 | 가입일 | 즐겨찾기(쉼표구분) | 역할
// 역할(F열)에 "teacher"를 입력하면 그 계정은 앱에서 학습현황 대시보드를 볼 수 있음
function getMemberSheet(ss) {
  var sh = ss.getSheetByName("회원");
  if (!sh) {
    sh = ss.insertSheet("회원");
    sh.appendRow(["아이디", "비밀번호", "이름", "가입일", "즐겨찾기", "역할"]);
  }
  if (String(sh.getRange(1, 6).getValue()) === "") sh.getRange(1, 6).setValue("역할");
  return sh;
}

function memberRole(rowData) {
  return String(rowData[5] || "").trim().toLowerCase();
}

// 학습기록 시트: 날짜 | 아이디 | 이름 | 탭 | 유형 | 점수 | 총문제 | 정답률 | 틀린단어 | 범위
function getRecordSheet(ss) {
  var sh = ss.getSheetByName("학습기록");
  if (!sh) {
    sh = ss.insertSheet("학습기록");
    sh.appendRow(["날짜", "아이디", "이름", "탭", "유형", "점수", "총문제", "정답률", "틀린단어", "범위"]);
    return sh;
  }
  // 기존 시트 마이그레이션: J열(범위) 헤더가 없으면 추가 (기존 행 데이터는 그대로 유지됨)
  if (String(sh.getRange(1, 10).getValue()).trim() === "") sh.getRange(1, 10).setValue("범위");
  return sh;
}

function findMemberRow(sh, id) {
  var data = sh.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === id) return { row: r + 1, data: data[r] };
  }
  return null;
}

function doPost(e) {
  // 동시 쓰기 충돌 방지
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var members = getMemberSheet(ss);

    if (action === "signup") {
      var id = String(body.id || "").trim();
      var pw = String(body.pw || "").trim();
      var name = String(body.name || "").trim();
      if (!id || !pw || !name) return jsonOut({ ok: false, error: "아이디, 비밀번호, 이름을 모두 입력해주세요." });
      if (findMemberRow(members, id)) return jsonOut({ ok: false, error: "이미 사용 중인 아이디입니다." });
      members.appendRow([id, pw, name, new Date(), ""]);
      return jsonOut({ ok: true, name: name, stars: [] });
    }

    if (action === "login") {
      var m = findMemberRow(members, String(body.id || "").trim());
      if (!m || String(m.data[1]).trim() !== String(body.pw || "").trim()) {
        return jsonOut({ ok: false, error: "아이디 또는 비밀번호가 올바르지 않습니다." });
      }
      var stars = String(m.data[4] || "").split(",").map(function(s) { return s.trim(); }).filter(Boolean);
      return jsonOut({ ok: true, name: String(m.data[2]), stars: stars, role: memberRole(m.data) });
    }

    // 저장된 세션 복원용 (비밀번호 없이 이름/즐겨찾기만 반환)
    if (action === "getUser") {
      var mg = findMemberRow(members, String(body.id || "").trim());
      if (!mg) return jsonOut({ ok: false, error: "회원을 찾을 수 없습니다." });
      var gStars = String(mg.data[4] || "").split(",").map(function(s) { return s.trim(); }).filter(Boolean);
      return jsonOut({ ok: true, name: String(mg.data[2]), stars: gStars, role: memberRole(mg.data) });
    }

    // 선생님 대시보드: 전체 회원 + 학습기록 (역할이 teacher인 계정만)
    if (action === "getDashboard") {
      var mt = findMemberRow(members, String(body.id || "").trim());
      if (!mt || memberRole(mt.data) !== "teacher") {
        return jsonOut({ ok: false, error: "선생님 계정만 조회할 수 있습니다." });
      }
      var mData = members.getDataRange().getValues();
      var memberList = [];
      for (var mi = 1; mi < mData.length; mi++) {
        var mid = String(mData[mi][0]).trim();
        if (!mid) continue;
        memberList.push({
          id: mid,
          name: String(mData[mi][2]),
          joined: mData[mi][3] instanceof Date ? mData[mi][3].toISOString() : String(mData[mi][3]),
          role: memberRole(mData[mi]),
        });
      }
      var rData = getRecordSheet(ss).getDataRange().getValues();
      var records = [];
      for (var ri = 1; ri < rData.length; ri++) {
        if (!String(rData[ri][1]).trim()) continue;
        records.push({
          date: rData[ri][0] instanceof Date ? rData[ri][0].toISOString() : String(rData[ri][0]),
          id: String(rData[ri][1]).trim(),
          tab: String(rData[ri][3]),
          type: String(rData[ri][4]),
          score: Number(rData[ri][5]) || 0,
          total: Number(rData[ri][6]) || 0,
          wrong: String(rData[ri][8] || ""),
          sections: String(rData[ri][9] || ""),
        });
      }
      return jsonOut({ ok: true, members: memberList, records: records });
    }

    // 학생 본인 학습기록 조회 (깜빡이 진도 표시용)
    if (action === "getRecords") {
      var myId = String(body.id || "").trim();
      if (!findMemberRow(members, myId)) return jsonOut({ ok: false, error: "회원을 찾을 수 없습니다." });
      var qData = getRecordSheet(ss).getDataRange().getValues();
      var mine = [];
      for (var qi = 1; qi < qData.length; qi++) {
        if (String(qData[qi][1]).trim() !== myId) continue;
        mine.push({
          date: qData[qi][0] instanceof Date ? qData[qi][0].toISOString() : String(qData[qi][0]),
          tab: String(qData[qi][3]),
          type: String(qData[qi][4]),
          score: Number(qData[qi][5]) || 0,
          total: Number(qData[qi][6]) || 0,
          sections: String(qData[qi][9] || ""),
        });
      }
      return jsonOut({ ok: true, records: mine });
    }

    if (action === "saveStars") {
      var m2 = findMemberRow(members, String(body.id || "").trim());
      if (!m2) return jsonOut({ ok: false, error: "회원을 찾을 수 없습니다." });
      var list = Array.isArray(body.stars) ? body.stars : [];
      members.getRange(m2.row, 5).setValue(list.join(","));
      return jsonOut({ ok: true });
    }

    if (action === "saveResult") {
      var m3 = findMemberRow(members, String(body.id || "").trim());
      if (!m3) return jsonOut({ ok: false, error: "회원을 찾을 수 없습니다." });
      var r = body.result || {};
      var total = Number(r.total) || 0;
      // 깜빡이테스트처럼 채점이 없는 유형은 score를 보내지 않음 → 점수/정답률 칸을 비워둠
      var hasScore = r.score !== undefined && r.score !== null && r.score !== "";
      var score = hasScore ? (Number(r.score) || 0) : "";
      var pct = (hasScore && total > 0) ? Math.round(Number(r.score) / total * 100) + "%" : "";
      getRecordSheet(ss).appendRow([
        new Date(), String(body.id).trim(), String(m3.data[2]),
        String(r.tab || ""), String(r.type || ""),
        score, total, pct,
        (Array.isArray(r.wrong) ? r.wrong : []).join(", "),
        (Array.isArray(r.sections) ? r.sections : []).join(", "),
      ]);
      return jsonOut({ ok: true });
    }

    return jsonOut({ ok: false, error: "알 수 없는 요청입니다." });
  } catch (err) {
    return jsonOut({ ok: false, error: "서버 오류: " + err });
  } finally {
    lock.releaseLock();
  }
}
