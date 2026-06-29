function setupCustomerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Customers");

  if (!sheet) {
    sheet = ss.insertSheet("Customers");
  }

  const headers = [
    "고객명",
    "전화번호",
    "첫예약일",
    "최근예약일",
    "예약횟수",
    "취소횟수",
    "노쇼횟수",
    "고객등급",
    "메모",
    "최근수정일시"
  ];

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

function searchCustomerByPhone(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  if (!sheet) return { found:false, message:"Customers 시트가 없습니다." };

  const targetPhone = normalizePhone_(phone);
  const lastRow = sheet.getLastRow();

  if (!targetPhone || lastRow <= 1) {
    return { found:false, message:"신규 고객입니다." };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  for (let i = 0; i < values.length; i++) {
    const savedPhone = normalizePhone_(values[i][1]);

    if (savedPhone === targetPhone) {
      return {
        found:true,
        customerName: values[i][0],
        phone: formatPhoneDisplay_(values[i][1]),
        firstDate: formatDate_(values[i][2]),
        recentDate: formatDate_(values[i][3]),
        reservationCount: values[i][4] || 0,
        cancelCount: values[i][5] || 0,
        noShowCount: values[i][6] || 0,
        grade: values[i][7] || "신규",
        memo: values[i][8] || ""
      };
    }
  }

  return { found:false, message:"신규 고객입니다." };
}

function upsertCustomerFromReservation(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  if (!sheet) return;

  const phone = normalizePhone_(data.phone);
  if (!phone) return;

  const now = new Date();
  const reserveDate = data.reserveDate || "";
  const customerName = data.customerName || "";
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

    for (let i = 0; i < values.length; i++) {
      const savedPhone = normalizePhone_(values[i][1]);

      if (savedPhone === phone) {
        const row = i + 2;
        const count = Number(values[i][4] || 0) + 1;
        const grade = getCustomerGrade_(count, Number(values[i][6] || 0));

        sheet.getRange(row, 1).setValue(customerName || values[i][0]);
        sheet.getRange(row, 4).setValue(reserveDate);
        sheet.getRange(row, 5).setValue(count);
        sheet.getRange(row, 8).setValue(grade);
        sheet.getRange(row, 10).setValue(now);
        return;
      }
    }
  }

  sheet.appendRow([
    customerName,
    phone,
    reserveDate,
    reserveDate,
    1,
    0,
    0,
    "신규",
    "",
    now
  ]);
}

function increaseCustomerCancelCount(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  if (!sheet) return;

  const targetPhone = normalizePhone_(phone);
  const lastRow = sheet.getLastRow();

  if (!targetPhone || lastRow <= 1) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  for (let i = 0; i < values.length; i++) {
    if (normalizePhone_(values[i][1]) === targetPhone) {
      const row = i + 2;
      const cancelCount = Number(values[i][5] || 0) + 1;
      const noShowCount = Number(values[i][6] || 0);
      const reservationCount = Number(values[i][4] || 0);

      sheet.getRange(row, 6).setValue(cancelCount);
      sheet.getRange(row, 8).setValue(getCustomerGrade_(reservationCount, noShowCount));
      sheet.getRange(row, 10).setValue(new Date());
      return;
    }
  }
}

function increaseCustomerNoShowCount(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  if (!sheet) return;

  const targetPhone = normalizePhone_(phone);
  const lastRow = sheet.getLastRow();

  if (!targetPhone || lastRow <= 1) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  for (let i = 0; i < values.length; i++) {
    if (normalizePhone_(values[i][1]) === targetPhone) {
      const row = i + 2;
      const noShowCount = Number(values[i][6] || 0) + 1;
      const reservationCount = Number(values[i][4] || 0);

      sheet.getRange(row, 7).setValue(noShowCount);
      sheet.getRange(row, 8).setValue(getCustomerGrade_(reservationCount, noShowCount));
      sheet.getRange(row, 10).setValue(new Date());
      return;
    }
  }
}

function getCustomerGrade_(reservationCount, noShowCount) {
  if (noShowCount >= 2) return "예약주의";
  if (reservationCount >= 20) return "VIP";
  if (reservationCount >= 5) return "단골";
  return "신규";
}