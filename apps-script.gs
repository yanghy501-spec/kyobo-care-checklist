/**
 * 맞춤관리 체크리스트 - 구글시트 응답 저장용 Apps Script
 *
 * 설치 방법
 * 1. 구글시트 새로 만들기 (이름 예: 맞춤관리 체크리스트 응답)
 * 2. 확장 프로그램 > Apps Script > 기존 코드 지우고 이 코드 전체 붙여넣기 > 저장
 * 3. 배포 > 새 배포 > 유형 선택(톱니) > 웹 앱
 *    - 다음 사용자 인증 정보로 실행: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 * 4. 배포 > 권한 승인 (경고 화면이 나오면 고급 > 이동 > 허용)
 * 5. 발급된 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 복사해서 전달
 *
 * 응답은 "맞춤관리체크" 탭에 한 줄씩 쌓입니다. 탭과 머리글은 첫 응답 때 자동으로 만들어집니다.
 */

const SHEET_NAME = '맞춤관리체크';
const HEADERS = ['시각', '유입경로', '성함', '생년월일', '연락처', '병원/진료과',
                 '자산관리', '사업장·노무관리', '보장관리', '기타 관심사항', '궁금한 점',
                 '안내 방식', '연락 방법', '시간대', '마케팅동의'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type !== 'checklist') return out('unknown type');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }

    const row = [new Date(), d.src, d.name, d.birth, d.phone, d.hosp,
                 d.asset, d.biz, d.cover, d.etc, d.question,
                 d.guide, d.method, d.time, d.mkt];
    sh.appendRow(row.map(safe));
    return out('ok');
  } finally {
    lock.releaseLock();
  }
}

// 브라우저로 웹 앱 주소를 열었을 때 연결 확인용
function doGet() {
  return out('맞춤관리 체크리스트 수집 서버 작동 중');
}

// 시트 수식 주입 방지: =, +, -, @ 로 시작하는 값 앞에 ' 추가
function safe(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v === undefined ? '' : v;
}

function out(msg) {
  return ContentService.createTextOutput(msg);
}
