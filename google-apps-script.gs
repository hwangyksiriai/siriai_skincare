/**
 * 시리아이 캠페인 신청서 → 구글 시트 수집용 Apps Script
 *
 * [설치 방법]
 * 1) 신청 내용을 받을 구글 시트를 새로 만든다.
 * 2) 상단 메뉴 [확장 프로그램] → [Apps Script] 클릭.
 * 3) 기존 코드를 모두 지우고 이 파일 내용을 통째로 붙여넣고 저장.
 * 4) 우측 상단 [배포] → [새 배포] → 유형 선택(톱니바퀴) → [웹 앱].
 *    - 설명: 아무거나
 *    - 실행 계정: 나
 *    - 액세스 권한: "모든 사용자"  (★ 중요: 익명도 제출 가능해야 함)
 * 5) [배포] 클릭 → 권한 승인 → 생성된 "웹 앱 URL"을 복사.
 * 6) index.html / index-50000.html / index-200000.html 세 파일 각각의
 *    SHEET_ENDPOINT 값에 같은 URL을 붙여넣는다. (3개 모두 동일 URL → 한 시트로 모임)
 *
 * 코드를 수정하면 반드시 [배포] → [배포 관리] → 기존 배포 [편집] → 새 버전으로 다시 배포해야 반영됨.
 */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    // 첫 행에 헤더가 없으면 자동 생성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '신청시각', '고료', '지원유형', '선택브랜드', '이름',
        '인스타그램', '휴대폰', '이메일', '우편번호', '주소', '상세주소', '요청사항', '원가청구동의'
      ]);
    }

    sheet.appendRow([
      new Date(),
      data.amount || '',
      data.support_type || '',
      data.brands || '',
      data.name || '',
      data.instagram || '',
      data.phone || '',
      data.email || '',
      data.zonecode || '',
      data.address || '',
      data.address_detail || '',
      data.notes || '',
      data.agree || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
