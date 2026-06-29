function saveReservation(data) {
  try {
    return createReservation(data);
  } catch (err) {
    return {
      ok: false,
      message: "서버 오류: " + err.message + "\n" + err.stack
    };
  }
}

function getReservationList() {
  return getAllReservations();
}

function getTodayDashboard() {
  return getTodayReservationData();
}

function checkSeatConflictApi(data) {
  const list = getAllReservations();

  const conflicts = list.filter(r =>
    r.reserveDate === data.reserveDate &&
    r.reserveTime === data.reserveTime &&
    r.seat === data.seat &&
    r.reserveStatus !== "예약취소"
  );

  return {
    conflict: conflicts.length > 0,
    count: conflicts.length,
    list: conflicts
  };
}

function cancelReservationApi(reservationNo, staff, reason) {
  return cancelReservation(reservationNo, staff, reason);
}

function updateReservationApi(data){
  return updateReservation(data);
}
function markDepositDoneApi(reservationNo, staff){
  return markDepositDone(reservationNo, staff);
}

function confirmReservationApi(reservationNo, staff){
  return confirmReservation(reservationNo, staff);
}
function markNoShowApi(reservationNo, staff, reason){
  return markNoShow(reservationNo, staff, reason);
}