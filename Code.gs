/**
 * 스마트 주보 - 구글 시트 통합 데이터 연동 엔진 스크립트
 * 처리 대상: 접속 로그, 마음으로 드리는 고백 제출, 목장 나눔, 심방 및 양육 자율 신청서, 성경 통독 체크
 */

function doPost(e) {
  if (e && e.parameter && e.parameter.type === '특별기록') return saveSpecialRecord_(e);
  if (e && e.parameter && e.parameter.type === 'reading') return saveReading_(e);   // 성경 통독 체크
  // CORS 처리 헤더 구성
  var output = ContentService.createTextOutput();
  
  try {
    var param = e.parameter;
    var type = param.type; // 데이터 전송 구분을 위한 식별 키
    
    // 현재 활성화된 구글 스프레드시트 가져오기
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "";
    var headers = [];
    var rowData = [];
    
    // 1. 요청 타입별 적치할 시트 이름 및 헤더 필드 매핑 구조 분기
    if (type === "접속자") {
      sheetName = "1_접속자로그";
      headers = ["기록일시", "주보설정날짜", "활동내용"];
      rowData = [param.date, param.bulletinDate, param.content];
      
    } else if (type === "마드고") {
      sheetName = "2_마음의고백";
      headers = ["제출일시", "작성자", "성경본문", "설교제목", "받은은혜와 깨달음", "기도제목"];
      rowData = [param.date, param.author, param.biblePassage, param.sermonTitle, param.grace, param.prayer];
      
    } else if (type === "목장나눔") {
      sheetName = "3_목장나눔기록";
      headers = ["작성일자", "소속목장", "작성성도", "나눔내용", "공유범위"];
      rowData = [param.date, param.mokjang, param.author, param.text, param.공유여부];
      
    } else if (type === "심방양육신청") {
      sheetName = "4_심방양육신청서";
      headers = ["신청일자", "소속목장", "신청성도", "희망일정 및 기입내용", "공유범위"];
      rowData = [param.date, param.mokjang, param.author, param.text, param.공유여부];
    } else {
      return output.setContent(JSON.stringify({status: "SUCCESS", message: "Unknown type bypassed"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 2. 대상 시트가 존재하지 않을 경우 자동 생성 및 헤더 배치 인프라 가동
    var targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(sheetName);
      targetSheet.appendRow(headers);
      // 첫 행 디자인 가독성 확보 (볼드 처리 및 배경색 하이라이트)
      targetSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#F1F5F9");
    }
    
    // 3. 최하단 행에 데이터 자동 추가 적치
    targetSheet.appendRow(rowData);
    
    // 최종 가공 완료 응답 반환
    return output.setContent(JSON.stringify({status: "SUCCESS", sheet: sheetName}))
                 .setMimeType(ContentService.MimeType.JSON);
                 
  } catch(error) {
    // 예외 발생 시 에러 추적 결과 반환
    return output.setContent(JSON.stringify({status: "ERROR", message: error.toString()}))
                 .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET 요청에 대한 기본 보안 확인 테스트 핑
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'reading') return readingStats_(e);   // 성경 통독 현황
  return ContentService.createTextOutput("스마트 주보 클라우드 서버 엔진 작동 중.");
}
function saveSpecialRecord_(e) {
  var p = e.parameter;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var name = String(p.recordTitle || '특별기록').replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90) || '특별기록';
    var sheet = ss.getSheetByName(name) || ss.insertSheet(name);

    var fields = [];
    try { fields = JSON.parse(p.fieldsJson || '[]'); } catch (err) { fields = []; }

    var fixed = ['제출일시', '이름', '목장/셀', '제출구분', '작성 항목 수'];
    var header = sheet.getLastRow() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String)
      : [];
    if (header.length === 0) header = fixed.slice();
    // 기존 열의 순서는 절대 바꾸지 않고, 새로 생긴 항목만 오른쪽 끝에 덧붙입니다.
    fields.forEach(function (f) { if (f && f.label && header.indexOf(f.label) === -1) header.push(f.label); });
    if (header.indexOf('전체 내용') === -1) header.push('전체 내용');
    if (header.indexOf('기록번호') === -1) header.push('기록번호');   // 수정·삭제할 줄을 찾는 번호

    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#FEFCBF');
    sheet.setFrozenRows(1);

    // 같은 성도(같은 기록번호)가 이미 낸 줄 찾기
    var idCol = header.indexOf('기록번호') + 1;
    var foundRow = 0;
    var entryId = String(p.entryId || '');
    if (entryId && sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
      for (var i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === entryId) { foundRow = i + 2; break; } }
    }

    // 삭제 요청
    if (p.action === 'delete') {
      if (foundRow) sheet.deleteRow(foundRow);
      return ContentService.createTextOutput('deleted');
    }

    var row = header.map(function () { return ''; });
    var put = function (col, val) { var k = header.indexOf(col); if (k !== -1) row[k] = val; };
    put('제출일시', p.date || new Date());
    put('이름', p.author || '');
    put('목장/셀', p.mokjang || '');
    put('제출구분', foundRow ? '수정 제출' : (p.submitKind || '최초 제출'));
    put('작성 항목 수', p.filled || '');
    fields.forEach(function (f) { if (f && f.label) put(f.label, f.text || ''); });
    put('전체 내용', p.content || '');
    put('기록번호', entryId);

    if (foundRow) {
      sheet.getRange(foundRow, 1, 1, header.length).setValues([row]);   // 원래 줄을 새 내용으로 교체
    } else {
      sheet.appendRow(row);
      foundRow = sheet.getLastRow();
    }
    sheet.getRange(foundRow, 1, 1, header.length).setWrap(true).setVerticalAlignment('top');
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}


/* ============================================================
 * 성경 통독 (v8)
 * 시트 '5_통독체크'에 한 줄씩 쌓입니다. (읽은 날짜·일차·별명·목장·기기번호)
 * - 같은 사람(기기)이 같은 일차를 다시 누르면 줄을 새로 만들지 않고 고칩니다.
 * - '읽었어요 취소'를 누르면 그 줄이 지워집니다.
 * - 앱 홈에 보이는 '오늘 읽은 사람 수·별명, 요일별 인원'은 아래 readingStats_ 가 계산해서 돌려줍니다.
 * ============================================================ */
var READING_SHEET = '5_통독체크';
var READING_HEADER = ['체크일시', '읽은날짜', '일차', '별명', '목장/셀', '별명숨김', '기기번호'];

function getReadingSheet_(create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(READING_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(READING_SHEET);
    sh.appendRow(READING_HEADER);
    sh.getRange(1, 1, 1, READING_HEADER.length).setFontWeight('bold').setBackground('#FBEFD3');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');   // 날짜가 달력 서식으로 바뀌지 않게 글자로 고정
    sh.getRange('G:G').setNumberFormat('@');
  }
  return sh;
}

function saveReading_(e) {
  var p = e.parameter;
  var device = String(p.device || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 40);
  var day = String(p.day || '');
  if (!device || !/^\d{1,3}$/.test(day)) return ContentService.createTextOutput('bad');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getReadingSheet_(true);
    var n = sh.getLastRow(), row = 0;
    if (n > 1) {
      var v = sh.getRange(2, 3, n - 1, 5).getValues();   // C~G: 일차, 별명, 목장, 숨김, 기기
      for (var i = v.length - 1; i >= 0; i--) {
        if (String(v[i][0]) === day && String(v[i][4]) === device) { row = i + 2; break; }
      }
    }
    var readDate = String(p.readDate || '').slice(0, 10);
    if (p.action === 'uncheck') {
      if (row) { sh.deleteRow(row); }
      CacheService.getScriptCache().remove('rd_' + readDate);
      return ContentService.createTextOutput('deleted');
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(readDate)) return ContentService.createTextOutput('bad date');
    var hidden = p.hidden ? '숨김' : '';
    var nick = hidden ? '' : String(p.nick || '').replace(/[\r\n\t]+/g, ' ').slice(0, 12);
    var rowValues = [[p.date || new Date(), readDate, Number(day), nick, String(p.mokjang || '').slice(0, 30), hidden, device]];
    if (row) {
      sh.getRange(row, 1, 1, 7).setValues(rowValues);
    } else {
      sh.appendRow(rowValues[0]);
    }
    CacheService.getScriptCache().remove('rd_' + readDate);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function readingStats_(e) {
  var date = String((e.parameter && e.parameter.date) || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) date = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var cache = CacheService.getScriptCache(), key = 'rd_' + date, hit = cache.get(key);
  var json = ContentService.MimeType.JSON;
  if (hit) return ContentService.createTextOutput(hit).setMimeType(json);

  var result = { status: 'OK', date: date, count: 0, names: [], week: [] };
  var sh = getReadingSheet_(false);
  if (sh && sh.getLastRow() > 1) {
    var d = new Date(date + 'T00:00:00Z');
    var offset = (d.getUTCDay() + 6) % 7;                       // 월요일부터 시작
    var days = [];
    for (var k = 0; k < 7; k++) {
      days.push(Utilities.formatDate(new Date(d.getTime() + (k - offset) * 86400000), 'UTC', 'yyyy-MM-dd'));
    }
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getDisplayValues();
    var perDay = {};                                           // 날짜 → { 기기번호: true }
    days.forEach(function (x) { perDay[x] = {}; });
    var seenName = {};
    for (var i = v.length - 1; i >= 0; i--) {                  // 최근에 읽은 사람이 먼저 나오게 아래에서 위로
      var rd = String(v[i][1]), dev = String(v[i][6]);
      if (!perDay[rd]) continue;
      perDay[rd][dev] = true;
      if (rd === date && v[i][3] && !seenName[dev] && result.names.length < 40) {
        seenName[dev] = true;
        result.names.push(String(v[i][3]));
      }
    }
    days.forEach(function (x) { result.week.push({ date: x, count: Object.keys(perDay[x]).length }); });
    result.count = Object.keys(perDay[date] || {}).length;
  }
  var text = JSON.stringify(result);
  cache.put(key, text, 30);                                    // 30초 동안은 같은 결과를 재사용해 시트 부담을 줄임
  return ContentService.createTextOutput(text).setMimeType(json);
}
