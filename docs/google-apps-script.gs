/**
 * 가브리엘라 천사의 집 축복미사 — 참석 접수 수신기
 *
 * 설치 방법은 docs/참석폼-설치.md 참고.
 * 구글 시트에 붙여넣고(확장 프로그램 > Apps Script) 웹 앱으로 배포한다.
 */

var SHEET_NAME = '참석명단';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return out({ ok: false, error: 'busy' });
  }

  try {
    var d = JSON.parse(e.postData.contents);

    // 자동 등록 봇 차단 — 숨김 필드에 값이 채워져 있으면 조용히 무시
    if (d.website) return out({ ok: true, skipped: 'trap' });

    var sh = sheet();

    // 중복 차단 — 같은 sid가 이미 있으면 저장하지 않음
    if (d.sid) {
      var sids = sh.getLastRow() > 1
        ? sh.getRange(2, 7, sh.getLastRow() - 1, 1).getValues()
        : [];
      for (var i = 0; i < sids.length; i++) {
        if (sids[i][0] === d.sid) return out({ ok: true, skipped: 'duplicate' });
      }
    }

    var guests = Math.max(0, Math.min(30, parseInt(d.guests, 10) || 0));

    sh.appendRow([
      new Date(),                       // A 접수시각
      cut(d.org, 60),                   // B 소속
      cut(d.name, 30),                  // C 이름
      cut(d.phone, 20),                 // D 연락처
      guests,                           // E 동행인
      1 + guests,                       // F 총인원
      cut(d.sid, 40)                    // G 접수키(중복 차단용)
    ]);

    return out({ ok: true });

  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 배포 확인용 — 브라우저로 웹앱 주소를 열면 보인다 */
function doGet() {
  var sh = sheet();
  var n = Math.max(0, sh.getLastRow() - 1);
  return out({ ok: true, message: '참석 접수 준비됨', count: n });
}

function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['접수시각', '소속', '이름', '연락처', '동행인', '총인원', '접수키']);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 160);
    sh.setColumnWidth(4, 130);
    sh.hideColumns(7); // 접수키는 내부용이라 숨김
  }
  return sh;
}

function cut(v, n) {
  return String(v == null ? '' : v).slice(0, n);
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
