// SiM-Guru — domain validation/authorization helpers
function _validateEventDefinition(payload) {
  var name =
    payload && typeof payload.name === "string" ? payload.name.trim() : "";
  if (!name) {
    return { ok: false, message: "Nama acara tidak boleh kosong." };
  }
  var timeStart =
    payload && typeof payload.time_start === "string"
      ? payload.time_start.trim()
      : "";
  var timeEnd =
    payload && typeof payload.time_end === "string"
      ? payload.time_end.trim()
      : "";
  var HH_MM = /^([01]\d|2[0-3]):[0-5]\d$/;
  if (!HH_MM.test(timeStart) || !HH_MM.test(timeEnd)) {
    return {
      ok: false,
      message: "Jam mulai harus lebih awal dari jam selesai.",
    };
  }
  if (timeStart >= timeEnd) {
    return {
      ok: false,
      message: "Jam mulai harus lebih awal dari jam selesai.",
    };
  }
  var raw = payload ? payload.jtm_val : undefined;
  var jtm;
  if (typeof raw === "number") {
    jtm = raw;
  } else if (typeof raw === "string") {
    var trimmed = raw.trim();
    jtm = trimmed === "" ? NaN : Number(trimmed);
  } else {
    jtm = NaN;
  }
  if (!Number.isInteger(jtm) || jtm < 1 || jtm > 99) {
    return {
      ok: false,
      message: "Nilai JTM harus berupa bilangan bulat antara 1 dan 99.",
    };
  }
  return { ok: true };
}

function _authorizeEventWriter(user) {
  return !!(user && typeof user.role === "string" && user.role === "admin");
}

function _computeEventJtm(attendanceRecords) {
  if (!Array.isArray(attendanceRecords)) return 0;
  return attendanceRecords.reduce(function (sum, r) {
    if (!r) return sum;
    var js = r.journal_submitted;
    if (js !== true && String(js).toLowerCase() !== "true") return sum;
    var val = Number(r.jtm_val);
    if (!isFinite(val) || val <= 0) return sum;
    return sum + val;
  }, 0);
}

function _computeTransportDates(regularDates, eventAttendanceRecords) {
  var regularArr = Array.isArray(regularDates) ? regularDates : [];
  var eventArr = Array.isArray(eventAttendanceRecords)
    ? eventAttendanceRecords
    : [];
  var eventDates = eventArr.reduce(function (acc, r) {
    if (r && r.journal_submitted === true && r.date) {
      acc.push(r.date);
    }
    return acc;
  }, []);
  var uniqueSet = new Set(regularArr.concat(eventDates));
  return Array.from(uniqueSet);
}
var _lastEventDaysCount = 0;

function _authorizeAttendanceWriter(user, isPicketConfirmed) {
  if (!user) return false;
  if (typeof user.role === "string" && user.role === "admin") return true;
  return isPicketConfirmed === true;
}
var JTM_ERROR_CODES = {
  NOT_AUTHORIZED: "NOT_AUTHORIZED",
  NOT_DAY_OF_OCCURRENCE: "NOT_DAY_OF_OCCURRENCE",
  ADJUSTED_EXCEEDS_SCHEDULED: "ADJUSTED_EXCEEDS_SCHEDULED",
  ADJUSTED_INVALID: "ADJUSTED_INVALID",
  REASON_REQUIRED: "REASON_REQUIRED",
  REASON_TOO_LONG: "REASON_TOO_LONG",
  ALLOCATION_INVALID: "ALLOCATION_INVALID",
  CONSERVATION_VIOLATION: "CONSERVATION_VIOLATION",
  SUBSTITUTE_IS_ORIGINAL: "SUBSTITUTE_IS_ORIGINAL",
  SUBSTITUTE_IS_ADMIN: "SUBSTITUTE_IS_ADMIN",
  NO_ACTIVE_CONFIRMATION: "NO_ACTIVE_CONFIRMATION",
  PERSISTENCE_FAILED: "PERSISTENCE_FAILED",
};
var JTM_REASON_MAX_LENGTH = 500;
