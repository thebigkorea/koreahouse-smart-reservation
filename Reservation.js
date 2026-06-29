function createReservation(data) {
validateReservationData_(data);

const ss = SpreadsheetApp.getActiveSpreadsheet();
const reservationSheet = ss.getSheetByName("Reservations");
const seatsSheet = ss.getSheetByName("Seats");

const reservationNo = generateReservationNo_();
const now = new Date();

const row = [
reservationNo,
now,
data.reserveDate || "",
data.reserveTime || "",
data.customerName || "",
normalizePhone_(data.phone || ""),
Number(data.people || 0),
data.seat || "",
data.eventType || "",
data.menu || "",
Number(data.deposit || 0),
data.depositStatus || "미입금",
data.reserveStatus || "예약접수",
data.staff || "",
data.memo || "",
data.channel || "전화",
"",
"",
"",
"",
now
];

reservationSheet.appendRow(row);

seatsSheet.appendRow([
data.reserveDate || "",
data.reserveTime || "",
data.seat || "",
reservationNo,
data.customerName || "",
"예약"
]);

upsertCustomerFromReservation(data);
sortReservations_();
writeReservationLog_(reservationNo, "예약등록", "", "", "", data.staff || "");

return {
ok: true,
reservationNo: reservationNo,
message: "예약이 등록되었습니다."
};
}

function getAllReservations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  return values
    .filter(function(row){
      return row[12] !== "예약취소";
    })
    .map(row => ({
      reservationNo: row[0],
      createdAt: formatDateTime_(row[1]),
      reserveDate: formatDate_(row[2]),
      reserveTime: formatTime_(row[3]),
      customerName: row[4],
      phone: formatPhoneDisplay_(row[5]),
      people: row[6],
      seat: row[7],
      eventType: row[8],
      menu: row[9],
      deposit: row[10],
      depositStatus: row[11],
      reserveStatus: row[12],
      staff: row[13],
      memo: row[14],
      channel: row[15],
      confirmSms: row[16],
      beforeSms: row[17],
      visitDone: row[18],
      visitDoneAt: formatDateTime_(row[19]),
      updatedAt: formatDateTime_(row[20])
    }));
}

function getTodayReservationData() {
const today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
const list = getAllReservations().filter(r =>
r.reserveDate === today &&
r.reserveStatus !== "예약취소"
);

const people = list.reduce((sum, r) => sum + Number(r.people || 0), 0);
const depositWait = list.filter(r => r.depositStatus !== "입금완료").length;
const confirmed = list.filter(r => r.reserveStatus === "예약확정").length;

return {
today,
count: list.length,
people,
depositWait,
confirmed,
reservations: list
};
}

function updateReservation(reservationNo, data) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
const row = findReservationRow_(reservationNo);

if (!row) {
return { ok: false, message: "예약을 찾을 수 없습니다." };
}

const oldData = sheet.getRange(row, 1, 1, 21).getValues()[0];

const newValues = [
reservationNo,
oldData[1],
data.reserveDate || oldData[2],
data.reserveTime || oldData[3],
data.customerName || oldData[4],
normalizePhone_(data.phone || oldData[5]),
Number(data.people || oldData[6] || 0),
data.seat || oldData[7],
data.eventType || oldData[8],
data.menu || oldData[9],
Number(data.deposit || oldData[10] || 0),
data.depositStatus || oldData[11],
data.reserveStatus || oldData[12],
data.staff || oldData[13],
data.memo || oldData[14],
data.channel || oldData[15],
oldData[16],
oldData[17],
oldData[18],
oldData[19],
new Date()
];

sheet.getRange(row, 1, 1, 21).setValues([newValues]);

refreshSeatForReservation_(reservationNo, newValues);
updateCustomerFromReservation_(data);
sortReservations_();

writeReservationLog_(reservationNo, "예약수정", "전체", JSON.stringify(oldData), JSON.stringify(newValues), data.staff || "");

return {
ok: true,
message: "예약이 수정되었습니다."
};
}

function cancelReservation(reservationNo, staff, reason) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
const row = findReservationRow_(reservationNo);

if (!row) {
return { ok: false, message: "예약을 찾을 수 없습니다." };
}

sheet.getRange(row, 13).setValue("예약취소");
sheet.getRange(row, 21).setValue(new Date());

const seatsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
const lastRow = seatsSheet.getLastRow();

if (lastRow > 1) {
const values = seatsSheet.getRange(2, 1, lastRow - 1, 6).getValues();

for (let i = 0; i < values.length; i++) {
  if (String(values[i][3]) === String(reservationNo)) {
    seatsSheet.getRange(i + 2, 6).setValue("취소");
  }
}

}

writeReservationLog_(reservationNo, "예약취소", "예약상태", "", reason || "취소", staff || "");

return {
ok: true,
message: "예약이 취소되었습니다."
};
}

function markDepositDone(reservationNo, staff){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if(!row){
    return {
      ok:false,
      message:"예약을 찾을 수 없습니다."
    };
  }

  sheet.getRange(row,12).setValue("입금완료");
  sheet.getRange(row,21).setValue(new Date());

  writeReservationLog_(
    reservationNo,
    "입금완료",
    "입금상태",
    "미입금",
    "입금완료",
    staff || ""
  );

  return {
    ok:true,
    message:"입금완료 처리되었습니다."
  };
}

function confirmReservation(reservationNo, staff){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if(!row){
    return {
      ok:false,
      message:"예약을 찾을 수 없습니다."
    };
  }

  // 입금 확인
  const depositStatus = sheet.getRange(row,11).getValue();

  if(depositStatus != "입금완료"){
    return {
      ok:false,
      message:"예약금을 먼저 입금완료 처리하세요."
    };
  }

  sheet.getRange(row,13).setValue("예약확정");

  writeReservationLog_(
    reservationNo,
    "예약확정",
    "예약상태",
    "",
    "",
    staff || ""
  );

  return {
    ok:true,
    message:"예약이 확정되었습니다."
  };
}

function findReservationRow_(reservationNo) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
const lastRow = sheet.getLastRow();

if (lastRow <= 1) return null;

const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

for (let i = 0; i < values.length; i++) {
if (String(values[i][0]) === String(reservationNo)) {
return i + 2;
}
}

return null;
}

function generateReservationNo_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd");
  const prefix = "KH-LWM-" + today + "-";

  const lastRow = sheet.getLastRow();
  let maxNo = 0;

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    values.forEach(function(row) {
      const no = row[0] ? row[0].toString() : "";

      if (no.indexOf(prefix) === 0) {
        const tail = Number(no.replace(prefix, ""));

        if (!isNaN(tail) && tail > maxNo) {
          maxNo = tail;
        }
      }
    });
  }

  return prefix + (maxNo + 1).toString().padStart(6, "0");
}

function validateReservationData_(data) {
if (!data) throw new Error("예약 데이터가 없습니다.");
if (!data.reserveDate) throw new Error("예약일은 필수입니다.");
if (!data.reserveTime) throw new Error("예약시간은 필수입니다.");
if (!data.customerName) throw new Error("고객명은 필수입니다.");
if (!data.phone) throw new Error("전화번호는 필수입니다.");
if (!data.people) throw new Error("인원은 필수입니다.");
}

function updateCustomerFromReservation_(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const phone = normalizePhone_(data.phone || "");

  if (!phone) return;

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

    for (let i = 0; i < values.length; i++) {
      const savedPhone = values[i][0] ? values[i][0].toString() : "";

      if (savedPhone === phone) {
        const row = i + 2;
        const count = Number(values[i][2] || 0) + 1;

        let grade = "일반";
        if (count >= 20) grade = "VIP";
        else if (count >= 10) grade = "Gold";
        else if (count >= 5) grade = "Silver";

        sheet.getRange(row, 2).setValue(data.customerName || values[i][1]);
        sheet.getRange(row, 3).setValue(count);
        sheet.getRange(row, 4).setValue(data.reserveDate || "");
        sheet.getRange(row, 5).setValue(grade);
        sheet.getRange(row, 11).setValue(new Date());
        return;
      }
    }
  }

  sheet.appendRow([
    phone,
    data.customerName || "",
    1,
    data.reserveDate || "",
    "일반",
    0,
    0,
    data.seat || "",
    "",
    "",
    new Date()
  ]);
}

function refreshSeatForReservation_(reservationNo, reservationValues) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
const lastRow = sheet.getLastRow();

if (lastRow > 1) {
const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

```
for (let i = 0; i < values.length; i++) {
  if (String(values[i][3]) === String(reservationNo)) {
    sheet.getRange(i + 2, 1, 1, 6).setValues([[
      reservationValues[2],
      reservationValues[3],
      reservationValues[7],
      reservationNo,
      reservationValues[4],
      reservationValues[12] === "예약취소" ? "취소" : "예약"
    ]]);
    return;
  }
}
```

}

sheet.appendRow([
reservationValues[2],
reservationValues[3],
reservationValues[7],
reservationNo,
reservationValues[4],
"예약"
]);
}

function sortReservations_() {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
const lastRow = sheet.getLastRow();

if (lastRow <= 2) return;

sheet.getRange(2, 1, lastRow - 1, 21).sort([
{ column: 3, ascending: true },
{ column: 4, ascending: true }
]);
}

function writeReservationLog_(reservationNo, action, field, beforeValue, afterValue, staff) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ReservationHistory");

sheet.appendRow([
new Date(),
reservationNo,
action,
field || "",
beforeValue || "",
afterValue || "",
staff || "시스템"
]);
}

function normalizePhone_(phone) {
  let p = (phone || "").toString().replace(/[^0-9]/g, "");

  if (p.length === 10 && p.charAt(0) !== "0") {
    p = "0" + p;
  }

  return p;
}

function formatPhoneDisplay_(phone) {
  let p = (phone || "").toString().replace(/[^0-9]/g, "");

  if (p.length === 10 && p.charAt(0) !== "0") {
    p = "0" + p;
  }

  if (p.length === 11) {
    return p.slice(0, 3) + "-" + p.slice(3, 7) + "-" + p.slice(7);
  }

  return p;
}

function formatDate_(value) {
if (!value) return "";

if (value instanceof Date) {
return Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd");
}

return String(value).trim();
}

function formatTime_(value) {
if (!value) return "";

if (value instanceof Date) {
return Utilities.formatDate(value, "Asia/Seoul", "HH:mm");
}

return String(value).trim();
}

function formatDateTime_(value) {
if (!value) return "";

if (value instanceof Date) {
return Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
}

return String(value).trim();
}
function testCreateReservation() {
  const result = createReservation({
    reserveDate: "2026-06-29",
    reserveTime: "11:00",
    customerName: "홍길동",
    phone: "010-1234-5678",
    people: 4,
    seat: "홀",
    eventType: "일반식사",
    menu: "청우정식",
    deposit: 50000,
    staff: "한경란",
    memo: "테스트 예약",
    channel: "전화"
  });

  Logger.log(result);
}
function updateReservation(data){

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Reservations");

  const values = sheet.getDataRange().getValues();

  for(let i=1;i<values.length;i++){

    if(values[i][0] == data.reservationNo){

      sheet.getRange(i+1,3).setValue(data.reserveDate);
      sheet.getRange(i+1,4).setValue(data.reserveTime);
      sheet.getRange(i+1,5).setValue(data.customerName);
      sheet.getRange(i+1,6).setValue(data.phone);
      sheet.getRange(i+1,7).setValue(data.people);
      sheet.getRange(i+1,8).setValue(data.seat);
      sheet.getRange(i+1,9).setValue(data.eventType);
      sheet.getRange(i+1,10).setValue(data.menu);
      sheet.getRange(i+1,11).setValue(data.deposit);
      sheet.getRange(i+1,14).setValue(data.staff);
      sheet.getRange(i+1,15).setValue(data.memo);
      sheet.getRange(i+1,21).setValue(new Date());

      return {
        ok:true,
        message:"예약이 수정되었습니다."
      };
    }
  }

  return {
    ok:false,
    message:"예약번호를 찾을 수 없습니다."
  };
}
function markNoShow(reservationNo, staff, reason){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if(!row){
    return {
      ok:false,
      message:"예약을 찾을 수 없습니다."
    };
  }

  const phone = sheet.getRange(row, 6).getValue();
  const oldStatus = sheet.getRange(row, 13).getValue();

  sheet.getRange(row, 13).setValue("노쇼");
  sheet.getRange(row, 21).setValue(new Date());

  increaseCustomerNoShowCount(phone);

  writeReservationLog_(
    reservationNo,
    "노쇼",
    "예약상태",
    oldStatus,
    "노쇼",
    (staff || "") + " / " + (reason || "")
  );

  return {
    ok:true,
    message:"노쇼 처리되었습니다."
  };
}