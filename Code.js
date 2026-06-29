const TIMEZONE = "Asia/Seoul";

function onOpen() {
SpreadsheetApp.getUi()
.createMenu("한국의집 예약관리")
.addItem("DB 초기 세팅", "setupKoreaHouseDB")
.addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("한국의집 스마트 예약관리")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupKoreaHouseDB() {
const ss = SpreadsheetApp.getActiveSpreadsheet();

setupSheet_(ss, "Reservations", [
"예약번호",
"등록일시",
"예약일",
"예약시간",
"고객명",
"전화번호",
"인원",
"좌석",
"행사구분",
"메뉴",
"예약금",
"입금상태",
"예약상태",
"담당자",
"요청사항",
"예약경로",
"예약확정문자",
"전날안내문자",
"방문완료",
"방문완료시간",
"수정일시"
]);

setupSheet_(ss, "Customers", [
"전화번호",
"고객명",
"방문횟수",
"최근방문일",
"VIP",
"노쇼횟수",
"예약취소횟수",
"선호좌석",
"알레르기",
"메모",
"최근수정일시"
]);

setupSheet_(ss, "Seats", [
"예약일",
"예약시간",
"좌석",
"예약번호",
"고객명",
"상태"
]);

setupSheet_(ss, "MessageHistory", [
"발송일시",
"예약번호",
"고객명",
"전화번호",
"문자종류",
"발송자",
"발송결과",
"메시지내용"
]);

setupSheet_(ss, "ReservationHistory", [
"기록일시",
"예약번호",
"작업",
"변경항목",
"변경전",
"변경후",
"작업자"
]);

setupSheet_(ss, "Settings", [
"구분",
"항목",
"값",
"비고"
]);

setupSheet_(ss, "Users", [
"이름",
"권한",
"사용여부",
"비고"
]);

setupSheet_(ss, "Logs", [
"기록일시",
"작업자",
"작업",
"상세내용"
]);

setupDefaultSettings_();

SpreadsheetApp.getUi().alert("한국의집 예약관리 DB 세팅 완료");
}

function setupSheet_(ss, sheetName, headers) {
let sheet = ss.getSheetByName(sheetName);
if (!sheet) sheet = ss.insertSheet(sheetName);

sheet.clear();
sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
sheet.setFrozenRows(1);

sheet.getRange(1, 1, 1, headers.length)
.setFontWeight("bold")
.setHorizontalAlignment("center")
.setBackground("#18252d")
.setFontColor("#ffffff");

sheet.autoResizeColumns(1, headers.length);
}

function setupDefaultSettings_() {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");

const values = [
["매장", "브랜드", "한국의집", ""],
["매장", "매장명", "롯데월드몰점", ""],
["매장", "예약번호코드", "KH-LWM", ""],
["문자", "발신번호", "0232134560", ""],
["계좌", "은행명", "하나은행", ""],
["계좌", "계좌번호", "602-910043-80104", ""],
["계좌", "예금주", "한국의집 롯데월드몰점", ""],
["좌석", "좌석목록", "홀,청우정,룸,별실,기타", ""],
["예약", "운영시작시간", "10:00", ""],
["예약", "운영종료시간", "20:00", ""],
["예약", "시간간격", "30", "분 단위"],
["정책", "자동문자발송", "사용안함", "직원 수동 선택 발송"]
];

sheet.getRange(2, 1, values.length, 4).setValues(values);
}
