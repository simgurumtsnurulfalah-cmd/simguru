const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const EXT_USERS_SS_ID = "1cnMeA7xinjvilVloS7A0_SFxzOBMON7DIiT7_cnnYUg";
const EXT_USERS_SHEET = "Users";
const SHEET_NAME = {
  USERS: "Users",
  SESSIONS: "Sessions",
  SCHEDULES: "Schedules",
  LOGS: "Teaching_Logs",
  CONFIG: "Config",
  CALENDAR: "Academic_Calendar",
  HONOR_HISTORY: "Honor_History",
  ALLOWANCES: "Allowances",
  SUBJECTS: "Subjects",
  RESET_REQUESTS: "Reset_Requests",
  ATTENDANCE: "Daily_Attendance",
  PICKET_SCHEDULES: "Picket_Schedules",
  SUBSTITUTES: "Substitutes",
  CEREMONY_SCHEDULES: "Ceremony_Schedules",
  ANNOUNCEMENTS: "Announcements",
  ATTENDANCE_SCHED_TEMPLATES: "Attendance_Sched_Templates",
  ATTENDANCE_LEAVES: "Attendance_Leaves",
  STUDENT_ATTENDANCE: "Student_Attendance",
};
const EXAM_SHEET = {
  PERIODS: "Exam_Periods",
  SESSIONS: "Exam_Sessions",
  ROOMS: "Exam_Rooms",
  SUPERVISORS: "Exam_Supervisors",
  COMMITTEE: "Exam_Committee",
  BAP: "Exam_BAP",
};
const EVENT_SHEET = {
  DEFINITIONS: "Event_Definitions",
  ATTENDANCE: "Event_Attendance",
  JOURNALS: "Event_Journals",
};
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
function _jtmValidateAdjusted(scheduledJtm, rawAdjusted) {
  var invalid = {
    ok: false,
    code: JTM_ERROR_CODES.ADJUSTED_INVALID,
    message:
      "Adjusted JTM harus berupa bilangan bulat antara 0 dan " +
      scheduledJtm +
      ".",
  };
  var value;
  if (typeof rawAdjusted === "number") {
    value = rawAdjusted;
  } else if (typeof rawAdjusted === "string") {
    var trimmed = rawAdjusted.trim();
    if (trimmed === "") return invalid;
    value = Number(trimmed);
  } else {
    return invalid;
  }
  if (typeof value !== "number" || isNaN(value) || !isFinite(value))
    return invalid;
  if (!Number.isInteger(value) || value < 0) return invalid;
  if (value > scheduledJtm) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.ADJUSTED_EXCEEDS_SCHEDULED,
      message:
        "Adjusted JTM tidak boleh melebihi Scheduled JTM (" +
        scheduledJtm +
        ").",
    };
  }
  return { ok: true, value: value };
}
function _jtmComputeDifference(scheduledJtm, adjustedJtm) {
  var difference = Math.round(scheduledJtm - adjustedJtm);
  return difference < 0 ? 0 : difference;
}
function _jtmValidateReason(adjustedJtm, scheduledJtm, rawReason) {
  var trimmed = "";
  if (typeof rawReason === "string") {
    trimmed = rawReason.trim();
  } else if (
    rawReason !== null &&
    rawReason !== undefined &&
    typeof rawReason !== "object"
  ) {
    trimmed = String(rawReason).trim();
  }
  if (!(adjustedJtm < scheduledJtm)) return { ok: true, reason: trimmed };
  if (trimmed.length === 0) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.REASON_REQUIRED,
      message: "Alasan penyesuaian wajib diisi ketika JTM dikurangi.",
    };
  }
  if (trimmed.length > JTM_REASON_MAX_LENGTH) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.REASON_TOO_LONG,
      message:
        "Alasan penyesuaian tidak boleh melebihi " +
        JTM_REASON_MAX_LENGTH +
        " karakter.",
    };
  }
  return { ok: true, reason: trimmed };
}
function _jtmValidateAllocation(
  rawAllocation,
  originalUserId,
  substituteUserId,
  isSubstituteAdmin,
) {
  var invalid = {
    ok: false,
    code: JTM_ERROR_CODES.ALLOCATION_INVALID,
    message: "Alokasi JTM harus berupa angka lebih besar dari 0.",
  };
  var value;
  if (typeof rawAllocation === "number") {
    value = rawAllocation;
  } else if (typeof rawAllocation === "string") {
    var trimmed = rawAllocation.trim();
    if (trimmed === "") return invalid;
    value = Number(trimmed);
  } else {
    return invalid;
  }
  if (typeof value !== "number" || isNaN(value) || !isFinite(value))
    return invalid;
  if (value <= 0) return invalid;
  if (substituteUserId === originalUserId) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.SUBSTITUTE_IS_ORIGINAL,
      message:
        "Guru pengganti tidak boleh sama dengan guru/pengawas asli untuk kegiatan ini.",
    };
  }
  if (isSubstituteAdmin) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.SUBSTITUTE_IS_ADMIN,
      message: "Guru pengganti tidak boleh seorang Admin.",
    };
  }
  return { ok: true, value: value };
}
function _jtmCheckConservation(
  scheduledJtm,
  adjustedJtm,
  existingAllocations,
  newAllocation,
) {
  var sumExisting = 0;
  if (existingAllocations && typeof existingAllocations.length === "number") {
    for (var i = 0; i < existingAllocations.length; i++) {
      var entry = Number(existingAllocations[i]);
      if (!isNaN(entry) && isFinite(entry)) sumExisting += entry;
    }
  }
  var addition = Number(newAllocation);
  if (isNaN(addition) || !isFinite(addition)) addition = 0;
  var base = Number(adjustedJtm);
  if (isNaN(base) || !isFinite(base)) base = 0;
  var total = base + sumExisting + addition;
  if (total <= scheduledJtm) return { ok: true, total: total };
  return {
    ok: false,
    code: JTM_ERROR_CODES.CONSERVATION_VIOLATION,
    total: total,
    message:
      "Total JTM (" +
      total +
      ") melebihi Scheduled JTM (" +
      scheduledJtm +
      "). Alokasi ditolak agar tidak melampaui batas.",
  };
}
function _jtmAuthorizeKbm(user, isPicketToday, todayDate, occurrenceDate) {
  var role =
    user && typeof user.role === "string" ? user.role.toLowerCase() : "";
  var isAdmin = role === "admin";
  var isAuthorizedRole = isAdmin || !!isPicketToday;
  if (!isAuthorizedRole) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.NOT_AUTHORIZED,
      message:
        "Anda tidak berwenang melakukan penyesuaian JTM untuk kegiatan KBM ini.",
    };
  }
  if (todayDate !== occurrenceDate) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.NOT_DAY_OF_OCCURRENCE,
      message:
        "Penyesuaian JTM hanya dapat dilakukan pada hari pelaksanaan (hari H) kegiatan.",
    };
  }
  return { ok: true };
}
function _jtmAuthorizeExam(authResult, todayDate, sessionDate) {
  var isAuthorizedRole = !!(authResult && authResult.ok);
  if (!isAuthorizedRole) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.NOT_AUTHORIZED,
      message:
        "Anda tidak berwenang melakukan penyesuaian JTM untuk pengawas ujian ini.",
    };
  }
  if (todayDate !== sessionDate) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.NOT_DAY_OF_OCCURRENCE,
      message:
        "Penyesuaian JTM hanya dapat dilakukan pada hari pelaksanaan (hari H) sesi ujian.",
    };
  }
  return { ok: true };
}
function doSetup() {
  const ss = SpreadsheetApp.openById(SS_ID);
  createSheetIfNeeded_(ss, SHEET_NAME.USERS, [
    "id",
    "username",
    "password",
    "role",
    "full_name",
    "nip",
    "base_salary",
    "photo_url",
    "security_question",
    "security_answer",
    "phone",
    "email",
    "km_distance",
    "additional_role",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.SESSIONS, [
    "token",
    "user_id",
    "expiry_time",
    "created_at",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.SCHEDULES, [
    "id",
    "user_id",
    "day_name",
    "day_index",
    "time_start",
    "time_end",
    "subject",
    "class_name",
    "jtm_val",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.LOGS, [
    "log_id",
    "schedule_id",
    "user_id",
    "date",
    "materi",
    "siswa_hadir",
    "siswa_absen",
    "notes",
    "jtm_val",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.CALENDAR, [
    "event_id",
    "date",
    "description",
    "is_holiday",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.CONFIG, ["key", "value", "description"]);
  createSheetIfNeeded_(ss, SHEET_NAME.HONOR_HISTORY, [
    "id",
    "user_id",
    "periode",
    "total_jtm",
    "total_honor",
    "tanggal_simpan",
    "details_json",
    "trx_id",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.ALLOWANCES, [
    "id",
    "duty_name",
    "user_id",
    "amount",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.SUBJECTS, ["id", "name"]);
  createSheetIfNeeded_(ss, SHEET_NAME.RESET_REQUESTS, [
    "request_id",
    "username",
    "full_name",
    "status",
    "created_at",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.ATTENDANCE, [
    "id",
    "date",
    "user_id",
    "status",
    "confirmed_at",
    "schedule_id",
    "sched_time_in",
    "sched_time_out",
    "time_in",
    "time_out",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.ATTENDANCE_SCHED_TEMPLATES, [
    "id",
    "day_index",
    "sched_type",
    "sched_time_in",
    "sched_time_out",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.ATTENDANCE_LEAVES, [
    "id",
    "date",
    "user_id",
    "leave_time",
    "return_time",
    "reason",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.PICKET_SCHEDULES, [
    "id",
    "day_index",
    "user_id",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.SUBSTITUTES, [
    "id",
    "date",
    "original_user_id",
    "substitute_user_id",
    "schedule_id",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.CEREMONY_SCHEDULES, [
    "id",
    "date",
    "user_id",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.ANNOUNCEMENTS, [
    "id",
    "title",
    "body",
    "severity",
    "is_dismissible",
    "is_active",
    "starts_at",
    "ends_at",
    "target_role",
    "cta_text",
    "cta_url",
    "created_by",
    "created_at",
    "updated_at",
    "target_user_ids",
    "dismissible_user_ids",
    "nondismissible_user_ids",
  ]);
  createSheetIfNeeded_(ss, "System_Locks", [
    "lock_id",
    "status",
    "timestamp",
    "created_by",
    "metadata",
  ]);
  createSheetIfNeeded_(ss, EXAM_SHEET.PERIODS, [
    "id",
    "name",
    "date_start",
    "date_end",
    "jtm_committee_per_day",
    "description",
    "created_by",
    "created_at",
  ]);
  createSheetIfNeeded_(ss, EXAM_SHEET.SESSIONS, [
    "id",
    "period_id",
    "date",
    "session_name",
    "time_start",
    "time_end",
    "jtm_val",
  ]);
  createSheetIfNeeded_(ss, EXAM_SHEET.ROOMS, [
    "id",
    "session_id",
    "room_name",
    "subject",
    "class_name",
    "total_siswa",
  ]);
  createSheetIfNeeded_(ss, EXAM_SHEET.SUPERVISORS, [
    "id",
    "room_id",
    "user_id",
    "status",
    "is_substitute",
    "original_user_id",
    "confirmed_by",
    "confirmed_at",
    "substitution_cancelled_by",
    "substitution_cancelled_at",
    "jtm_val",
  ]);
  createSheetIfNeeded_(ss, EXAM_SHEET.COMMITTEE, [
    "id",
    "period_id",
    "date",
    "user_id",
    "jtm_val",
    "status",
    "is_substitute",
    "original_user_id",
    "confirmed_by",
    "confirmed_at",
    "substitution_cancelled_by",
    "substitution_cancelled_at",
  ]);
  createSheetIfNeeded_(ss, EXAM_SHEET.BAP, [
    "id",
    "supervisor_id",
    "peserta_hadir",
    "peserta_absen",
    "catatan",
    "submitted_at",
    "updated_at",
  ]);
  createSheetIfNeeded_(ss, JTM_SHEET.ADJUSTMENTS, JTM_ADJUSTMENTS_HEADERS);
  createSheetIfNeeded_(ss, JTM_SHEET.REALLOCATIONS, JTM_REALLOCATIONS_HEADERS);
  createSheetIfNeeded_(ss, EVENT_SHEET.DEFINITIONS, [
    "id",
    "name",
    "type",
    "recurrence_day_index",
    "dates_json",
    "time_start",
    "time_end",
    "jtm_val",
    "description",
    "created_by",
    "created_at",
  ]);
  createSheetIfNeeded_(ss, EVENT_SHEET.ATTENDANCE, [
    "id",
    "event_id",
    "date",
    "user_id",
    "time_in",
    "confirmed_by",
    "confirmed_at",
    "journal_submitted",
    "jtm_val",
  ]);
  createSheetIfNeeded_(ss, EVENT_SHEET.JOURNALS, [
    "journal_id",
    "event_id",
    "attendance_id",
    "user_id",
    "date",
    "description",
    "submitted_at",
  ]);
  createSheetIfNeeded_(ss, SHEET_NAME.STUDENT_ATTENDANCE, [
    "id", "date", "class_name", "hadir", "sakit", "izin", "alpa", "submitted_by", "submitted_at"
  ]);
}
function getStudentAttendancePageData(token) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const isAdmin = String(user.role).toLowerCase() === "admin";
    const isPicket = _isPicketOfficer(user);
    if (!isAdmin && !isPicket) return { status: "error", message: "Akses Ditolak: Halaman ini hanya dapat diakses oleh guru piket atau guru piket pengganti yang kehadirannya sudah dikonfirmasi admin hari ini." };
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const cfg = _getConfigMap();
    const activeTP = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    const allSchedules = getData("Schedules");
    const classSet = new Set();
    allSchedules.forEach(function(s) {
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      if (sTP === activeTP && sSem === activeSem) {
        const cn = String(s.class_name || "").trim();
        if (cn) classSet.add(cn);
      }
    });
    const classList = Array.from(classSet).sort();
    const saData = getData("Student_Attendance");
    const todayRekap = saData.filter(function(r) {
      return safeDate(r.date) === todayStr;
    });
    const classes = classList.map(function(cn) {
      const rekap = todayRekap.find(function(r) { return String(r.class_name).trim() === cn; });
      const stdVal = parseInt(cfg["std_kelas_" + cn]);
      return {
        class_name: cn,
        hadir: rekap ? (parseInt(rekap.hadir) || 0) : 0,
        sakit: rekap ? (parseInt(rekap.sakit) || 0) : 0,
        izin: rekap ? (parseInt(rekap.izin) || 0) : 0,
        alpa: rekap ? (parseInt(rekap.alpa) || 0) : 0,
        submitted_by: rekap ? String(rekap.submitted_by || "") : "",
        submitted_at: rekap ? String(rekap.submitted_at || "") : "",
        has_data: !!rekap,
        std_siswa: isNaN(stdVal) ? null : stdVal
      };
    });
    return {
      status: "success",
      is_admin: isAdmin,
      is_picket_today: isPicket,
      date_str: todayStr,
      active_tp: activeTP,
      active_sem: activeSem,
      classes: classes
    };
  } catch (e) {
    return { status: "error", message: "Server Error (SA): " + e.toString() };
  }
}
function saveStudentAttendance(token, payload) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const isAdmin = String(user.role).toLowerCase() === "admin";
    const isPicket = _isPicketOfficer(user);
    if (!isAdmin && !isPicket) return { status: "error", message: "Akses Ditolak." };
    const className = payload && String(payload.class_name || "").trim();
    if (!className) return { status: "error", message: "Nama kelas wajib diisi." };
    const hadir = parseInt(payload.hadir) || 0;
    const sakit = parseInt(payload.sakit) || 0;
    const izin  = parseInt(payload.izin)  || 0;
    const alpa  = parseInt(payload.alpa)  || 0;
    if (hadir < 0 || sakit < 0 || izin < 0 || alpa < 0) return { status: "error", message: "Nilai tidak boleh negatif." };
    if ((hadir + sakit + izin + alpa) === 0) return { status: "error", message: "Total siswa tidak boleh nol." };
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const timeStr = Utilities.formatDate(new Date(), tz, "HH:mm:ss");
    const sheet = getSheet("Student_Attendance");
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(function(h) { return String(h).toLowerCase().trim().replace(/\s+/g,"_"); });
    const dateCol   = headers.indexOf("date");
    const classCol  = headers.indexOf("class_name");
    const hadirCol  = headers.indexOf("hadir");
    const sakitCol  = headers.indexOf("sakit");
    const izinCol   = headers.indexOf("izin");
    const alpaCol   = headers.indexOf("alpa");
    const subByCol  = headers.indexOf("submitted_by");
    const subAtCol  = headers.indexOf("submitted_at");
    let existingRow = -1;
    for (let i = 1; i < values.length; i++) {
      const rowDate  = safeDate(values[i][dateCol]);
      const rowClass = String(values[i][classCol] || "").trim();
      if (rowDate === todayStr && rowClass === className) { existingRow = i; break; }
    }
    if (existingRow > 0) {
      const r = existingRow + 1;
      sheet.getRange(r, hadirCol + 1).setValue(hadir);
      sheet.getRange(r, sakitCol + 1).setValue(sakit);
      sheet.getRange(r, izinCol  + 1).setValue(izin);
      sheet.getRange(r, alpaCol  + 1).setValue(alpa);
      sheet.getRange(r, subByCol + 1).setValue(user.id);
      sheet.getRange(r, subAtCol + 1).setValue(timeStr);
      _invalidateDataSnapshot();
      return { status: "success", id: String(values[existingRow][0]) };
    } else {
      const newId = generateId("SA");
      sheet.appendRow([newId, todayStr, className, hadir, sakit, izin, alpa, user.id, timeStr]);
      _invalidateDataSnapshot();
      return { status: "success", id: newId };
    }
  } catch (e) {
    return { status: "error", message: "Server Error (SaveSA): " + e.toString() };
  }
}
function checkStudentAttendanceToday(token) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const isAdmin = String(user.role).toLowerCase() === "admin";
    const isPicket = _isPicketOfficer(user);
    if (!isAdmin && !isPicket) return { status: "success", redirect_needed: false };
    if (isAdmin) return { status: "success", redirect_needed: false, is_admin: true, is_picket_today: false };
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const dayIdx = _getIndoDayIndex();
    const cfg = _getConfigMap();
    const activeTP  = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    const allSchedules = getData("Schedules");
    const classSet = new Set();
    allSchedules.forEach(function(s) {
      const sTP  = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      const isToday = String(s.day_index) === String(dayIdx);
      if (isToday && sTP === activeTP && sSem === activeSem) {
        const cn = String(s.class_name || "").trim();
        if (cn) classSet.add(cn);
      }
    });
    const totalCount = classSet.size;
    const saData = getData("Student_Attendance");
    const submittedCount = saData.filter(function(r) { return safeDate(r.date) === todayStr; }).length;
    const redirectNeeded = totalCount > 0 && submittedCount < totalCount;
    return {
      status: "success",
      redirect_needed: redirectNeeded,
      submitted_count: submittedCount,
      total_count: totalCount,
      is_admin: isAdmin,
      is_picket_today: isPicket
    };
  } catch (e) {
    return { status: "success", redirect_needed: false };
  }
}
function getStudentAttendanceHistory(token, params) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const isAdmin = String(user.role).toLowerCase() === "admin";
    let records = getData("Student_Attendance");
    const month = params && params.month ? String(params.month).padStart(2, "0") : null;
    const year  = params && params.year  ? String(params.year)  : null;
    if (month && year) {
      const prefix = year + "-" + month;
      records = records.filter(function(r) {
        return String(r.date || "").startsWith(prefix);
      });
    }
    if (!isAdmin) {
      records = records.filter(function(r) {
        return String(r.submitted_by) === String(user.id);
      });
    }
    const users = getData("Users");
    const userMap = {};
    users.forEach(function(u) {
      userMap[String(u.id)] = u.full_name || u.username || String(u.id);
    });
    const mapped = records.map(function(r) {
      // submitted_at is already normalised to "HH:mm" by getData()
      // which formats any Date value via Utilities.formatDate(val, tz, "HH:mm").
      // If it comes through as a plain "HH:mm:ss" string, trim to 5 chars.
      var submittedAt = r.submitted_at;
      if (typeof submittedAt === "string" && submittedAt.length > 5) {
        submittedAt = submittedAt.substring(0, 5);
      }
      return {
        id: r.id,
        date: r.date,
        class_name: r.class_name,
        hadir: parseInt(r.hadir) || 0,
        sakit: parseInt(r.sakit) || 0,
        izin:  parseInt(r.izin)  || 0,
        alpa:  parseInt(r.alpa)  || 0,
        submitted_by: r.submitted_by,
        submitted_by_name: userMap[String(r.submitted_by)] || String(r.submitted_by || "-"),
        submitted_at: submittedAt || "-"
      };
    });
    mapped.sort(function(a, b) {
      if (b.date !== a.date) return b.date.localeCompare(a.date);
      return String(a.class_name).localeCompare(String(b.class_name));
    });
    return { status: "success", is_admin: isAdmin, records: mapped };
  } catch(e) {
    return { status: "error", message: "Server Error (SAH): " + e.toString() };
  }
}
function createSheetIfNeeded_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
  } else {
    const currentHeaders = sheet
      .getRange(1, 1, 1, sheet.getLastColumn() || 1)
      .getValues()[0];
    if (headers.length > currentHeaders.length) {
      const missingHeaders = headers.slice(currentHeaders.length);
      if (missingHeaders.length > 0) {
        sheet
          .getRange(1, currentHeaders.length + 1, 1, missingHeaders.length)
          .setValues([missingHeaders]);
      }
    }
  }
}
function getDbId() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getId();
  } catch (e) {
    throw new Error(
      "Script tidak terhubung ke Spreadsheet. Pastikan script ini berada di dalam file Google Sheets (Container-bound) atau masukkan ID secara manual di Server_Database.gs",
    );
  }
}
let _cachedSS = null;
function getSheet(name) {
  if (!_cachedSS) {
    const id = getDbId();
    _cachedSS = SpreadsheetApp.openById(id);
  }
  let sheet = _cachedSS.getSheetByName(name);
  if (!sheet) {
    sheet = _cachedSS.insertSheet(name);
  }
  return sheet;
}
function getData(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0].map((h) =>
      String(h).toLowerCase().trim().replace(/\s+/g, "_"),
    );
    const rows = data.slice(1);
    var tz = Session.getScriptTimeZone();
    return rows.map((row) => {
      let obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (h.includes("date") && val instanceof Date) {
          obj[h] = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        } else if (h === "periode" && val instanceof Date) {
          obj[h] = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        } else if (
          (h === "time_in" ||
            h === "time_out" ||
            h === "sched_time_in" ||
            h === "sched_time_out" ||
            h === "leave_time" ||
            h === "return_time" ||
            h === "time_start" ||
            h === "time_end" ||
            h === "submitted_at") &&
          val instanceof Date
        ) {
          obj[h] = Utilities.formatDate(val, tz, "HH:mm");
        } else if (val instanceof Date) {
          obj[h] = val.toISOString();
        } else {
          obj[h] = val;
        }
      });
      return obj;
    });
  } catch (e) {
    console.error("Error in getData: " + e.toString());
    return [];
  }
}
function findData(sheetName, key, value) {
  const all = getData(sheetName);
  return all.find((item) => String(item[key]) === String(value));
}
function generateId(prefix) {
  return prefix + "-" + new Date().getTime();
}
function formatRupiah(num) {
  let nominal = Number(num);
  if (isNaN(nominal)) nominal = 0;
  return new Intl.NumberFormat("id-ID", {
    style: "currency",
    currency: "IDR",
    minimumFractionDigits: 0,
  }).format(nominal);
}
function formatDateIndo(dateInput) {
  if (!dateInput) return "-";
  const d = new Date(dateInput);
  if (isNaN(d.getTime())) return String(dateInput);
  const months = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ];
  const day = String(d.getDate()).padStart(2, "0");
  const month = months[d.getMonth()];
  const year = d.getFullYear();
  return `${day} ${month} ${year}`;
}
function loginUser(username, password) {
  const extPassword = _extGetPassword_(username);
  if (extPassword === null || String(extPassword) !== String(password)) {
    return { status: "error", message: "Username atau Password salah!" };
  }
  const users = getData("Users");
  const user = users.find(
    (u) => String(u.username).trim() === String(username).trim(),
  );
  if (!user) {
    return {
      status: "error",
      message: "Profil pengguna tidak ditemukan. Hubungi Administrator.",
    };
  }
  // --- Maintenance Mode Check ---
  const configData = getData("Config");
  const _cfgVal = (key) => { const item = configData.find((c) => c.key === key); return item ? item.value : ""; };
  let isMaintenance = String(_cfgVal("maintenance_mode")).toLowerCase() === "true";
  
  if (!isMaintenance && String(_cfgVal("maintenance_scheduled")).toLowerCase() === "true") {
    const startStr = _cfgVal("maintenance_start");
    const endStr = _cfgVal("maintenance_end");
    if (startStr && endStr) {
      const start = new Date(startStr);
      const end = new Date(endStr);
      const now = new Date();
      if (!isNaN(start.getTime()) && !isNaN(end.getTime()) && now >= start && now <= end) {
        isMaintenance = true;
      }
    }
  }

  if (isMaintenance) {
    if (String(user.role).toLowerCase() !== "admin") {
      const maintenanceMsg = _cfgVal("maintenance_message") || "";
      return {
        status: "maintenance",
        message: maintenanceMsg,
      };
    }
  }
  // --- End Maintenance Mode Check ---
  const token = "TKN-" + Utilities.getUuid();
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + 24);
  try {
    const sessionsSheet = getSheet("Sessions");
    const lastRow = sessionsSheet.getLastRow();
    if (lastRow > 1) {
      const range = sessionsSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      const now = new Date();
      for (let i = range.length - 1; i >= 0; i--) {
        const row = range[i];
        const expiryRaw = row[2];
        let expDate = null;
        if (expiryRaw instanceof Date) expDate = expiryRaw;
        else if (expiryRaw) {
          const tryD = new Date(expiryRaw);
          if (!isNaN(tryD.getTime())) expDate = tryD;
        }
        const isExpired = !expDate || expDate <= now;
        if (isExpired) {
          sessionsSheet.deleteRow(i + 2);
        }
      }
    }
  } catch (e) {}
  getSheet("Sessions").appendRow([token, user.id, expiry, new Date()]);
  const isDefaultPass = String(password) === "123456";
  const hasSecurityQuestion = !!(
    user.security_question && String(user.security_question).trim()
  );
  const hasEmailVerified = !!(
    user.email &&
    String(user.email).trim() &&
    _isValidEmail_(String(user.email))
  );
  return {
    status: "success",
    token,
    user: {
      id: user.id,
      name: user.full_name,
      role: user.role,
      is_default_pass: isDefaultPass,
      has_security_question: hasSecurityQuestion,
      has_email_verified: hasEmailVerified,
      email: String(user.email || ""),
    },
  };
}
function logoutUser(token) {
  const sheet = getSheet("Sessions");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { status: "success" };
}
function verifySession(token) {
  const sheet = getSheet("Sessions");
  const data = sheet.getDataRange().getValues();
  const session = data.find((row) => row[0] === token);
  if (!session) return null;
  const expiry = new Date(session[2]);
  if (new Date() > expiry) return null;
  const userId = session[1];
  const user = findData("Users", "id", userId);
  return user;
}
function checkSessionStatus(token) {
  try {
    if (!token) return { status: "expired" };
    const user = verifySession(token);
    return user ? { status: "ok" } : { status: "expired" };
  } catch (e) {
    console.error("checkSessionStatus error: " + e);
    return { status: "expired" };
  }
}
function updateUserProfile(token, payload) {
  const user = verifySession(token);
  if (!user)
    return { status: "error", message: "Sesi habis, silakan login kembali" };
  const sheetData = getData("Users");
  const userRowIndex = sheetData.findIndex(
    (u) => String(u.id) === String(user.id),
  );
  if (userRowIndex === -1)
    return { status: "error", message: "User tidak ditemukan" };
  const sheetObj = getSheet("Users");
  const sheetRowNumber = userRowIndex + 2;
  if (payload.full_name && !payload.phone && payload.phone !== "") {
    sheetObj.getRange(sheetRowNumber, 5).setValue(payload.full_name);
    return { status: "success", message: "Nama berhasil diubah" };
  }
  if (typeof payload.phone === "string") {
    const cleaned = _normalizePhoneNumber_(payload.phone);
    if (payload.phone.trim() !== "" && !cleaned.valid) {
      return {
        status: "error",
        message: cleaned.message || "Format nomor HP tidak valid.",
      };
    }
    const phoneCol = _ensureUserColumn_(sheetObj, "phone", 11);
    sheetObj.getRange(sheetRowNumber, phoneCol).setValue(cleaned.normalized);
    if (payload.full_name)
      sheetObj.getRange(sheetRowNumber, 5).setValue(payload.full_name);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Data berhasil disimpan",
      phone: cleaned.normalized,
    };
  }
  if (payload.new_password && payload.old_password) {
    const storedUsername = sheetData[userRowIndex].username;
    const currentExtPass = _extGetPassword_(storedUsername);
    if (
      currentExtPass === null ||
      String(currentExtPass) !== String(payload.old_password)
    ) {
      return {
        status: "error",
        message: "Password lama yang Anda masukkan salah.",
      };
    }
    if (String(payload.new_password) === String(payload.old_password)) {
      return {
        status: "error",
        message: "Password baru tidak boleh sama dengan password lama.",
      };
    }
    const complexityCheck = _checkPasswordComplexity(
      payload.new_password,
      user.full_name,
      storedUsername,
    );
    if (!complexityCheck.valid) {
      return { status: "error", message: complexityCheck.message };
    }
    const extOk = _extSetPassword_(storedUsername, payload.new_password);
    if (!extOk) {
      return {
        status: "error",
        message: "Gagal memperbarui password. Hubungi Administrator.",
      };
    }
    sheetObj.getRange(sheetRowNumber, 3).setValue(payload.new_password);
    return { status: "success", message: "Password berhasil diubah" };
  }
  return { status: "error", message: "Data tidak lengkap" };
}
function updateGuruProfileDetails(token, payload) {
  const user = verifySession(token);
  if (!user)
    return { status: "error", message: "Sesi habis, silakan login kembali" };
  const sheetData = getData("Users");
  const userRowIndex = sheetData.findIndex(
    (u) => String(u.id) === String(user.id),
  );
  if (userRowIndex === -1)
    return { status: "error", message: "User tidak ditemukan" };
  const sheetObj = getSheet("Users");
  const sheetRowNumber = userRowIndex + 2;
  sheetObj
    .getRange(sheetRowNumber, 15)
    .setValue(payload.status_kepegawaian || "");
  sheetObj.getRange(sheetRowNumber, 16).setValue(payload.golongan || "");
  sheetObj.getRange(sheetRowNumber, 17).setValue(payload.alamat || "");
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  return { status: "success", message: "Profil guru berhasil disimpan" };
}
function getMyProfile(token) {
  const user = verifySession(token);
  if (!user)
    return { status: "error", message: "Sesi habis, silakan login kembali" };
  const sheetData = getData("Users");
  const me = sheetData.find((u) => String(u.id) === String(user.id));
  if (!me) return { status: "error", message: "User tidak ditemukan" };
  const configRaw = getData("Config");
  let config = {};
  (configRaw || []).forEach((c) => {
    config[c.key] = c.value;
  });
  const activeTP = config.tahun_pelajaran || "";
  const activeSem = config.semester || "";
  const allowancesData = getData("Allowances") || [];
  const dutyNames = allowancesData
    .filter((a) => String(a.user_id).trim() === String(user.id).trim())
    .map((a) => a.duty_name)
    .filter(Boolean);
  const jabatan = [...new Set(dutyNames)];
  const schedulesData = getData("Schedules") || [];
  const mySubjects = schedulesData
    .filter((s) => {
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return (
        String(s.user_id).trim() === String(user.id).trim() &&
        sTP === activeTP &&
        sSem === activeSem
      );
    })
    .map((s) => `${s.subject || ""} (${s.class_name || ""})`.trim());
  const mapel_diampu = [...new Set(mySubjects)].filter(Boolean);
  const userRowIndex = sheetData.findIndex(
    (u) => String(u.id) === String(user.id),
  );
  const sheetObj = getSheet("Users");
  const userRowNumber = userRowIndex + 2;
  let profileDetails = ["", "", ""];
  if (userRowIndex !== -1) {
    try {
      profileDetails = sheetObj
        .getRange(userRowNumber, 15, 1, 3)
        .getValues()[0];
    } catch (e) {}
  }
  return {
    status: "success",
    profile: {
      id: String(me.id || ""),
      username: String(me.username || ""),
      full_name: String(me.full_name || ""),
      role: String(me.role || ""),
      nip: String(me.nip || ""),
      phone: String(me.phone || ""),
      email: String(me.email || ""),
      km_distance: String(me.km_distance || ""),
      jabatan: jabatan,
      mapel_diampu: mapel_diampu,
      status_kepegawaian: String(profileDetails[0] || ""),
      golongan: String(profileDetails[1] || ""),
      alamat: String(profileDetails[2] || ""),
    },
  };
}
function _ensureUserColumn_(sheetObj, headerName, defaultPos) {
  const lastCol = sheetObj.getLastColumn();
  const headers = sheetObj.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (
      String(headers[i]).trim().toLowerCase() ===
      String(headerName).toLowerCase()
    ) {
      return i + 1;
    }
  }
  const newCol = lastCol + 1;
  sheetObj.getRange(1, newCol).setValue(headerName);
  return newCol;
}
function _normalizePhoneNumber_(input) {
  let s = String(input || "").trim();
  if (!s) return { valid: true, normalized: "" };
  s = s.replace(/[\s\-\(\)\.]/g, "");
  if (s.indexOf("+") === 0) s = s.substring(1);
  if (!/^\d+$/.test(s)) {
    return {
      valid: false,
      normalized: "",
      message: "Nomor HP hanya boleh mengandung angka.",
    };
  }
  if (s.charAt(0) === "0") {
    s = "62" + s.substring(1);
  } else if (s.charAt(0) !== "6" || s.charAt(1) !== "2") {
    if (s.length >= 8 && s.length <= 13) {
      s = "62" + s;
    }
  }
  if (s.length < 10 || s.length > 15) {
    return {
      valid: false,
      normalized: "",
      message: "Nomor HP harus 10–15 digit (termasuk kode negara).",
    };
  }
  if (s.indexOf("62") !== 0) {
    return {
      valid: false,
      normalized: "",
      message: "Nomor HP harus menggunakan kode negara Indonesia (62).",
    };
  }
  return { valid: true, normalized: s };
}
function requestEmailOTP(token, newEmail) {
  try {
    const user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali" };
    const email = String(newEmail || "")
      .trim()
      .toLowerCase();
    if (!email) return { status: "error", message: "Email wajib diisi." };
    if (!_isValidEmail_(email)) {
      return { status: "error", message: "Format email tidak valid." };
    }
    const sheetData = getData("Users");
    const otherWithEmail = sheetData.find(
      (u) =>
        String(u.email || "").toLowerCase() === email &&
        String(u.id) !== String(user.id),
    );
    if (otherWithEmail) {
      return {
        status: "error",
        message: "Email ini sudah digunakan oleh akun lain.",
      };
    }
    const props = PropertiesService.getScriptProperties();
    const rateKey = "otp_rate_" + user.id;
    const lastReq = Number(props.getProperty(rateKey) || 0);
    const now = Date.now();
    const SECS = 60;
    if (lastReq && now - lastReq < SECS * 1000) {
      const wait = Math.ceil((SECS * 1000 - (now - lastReq)) / 1000);
      return {
        status: "error",
        message: "Mohon tunggu " + wait + " detik sebelum meminta kode lagi.",
      };
    }
    const code = String(Math.floor(100000 + Math.random() * 900000));
    const dataKey = "otp_data_" + user.id;
    const payload = {
      code: code,
      email: email,
      expires: now + 10 * 60 * 1000,
      attempts: 0,
    };
    props.setProperty(dataKey, JSON.stringify(payload));
    props.setProperty(rateKey, String(now));
    try {
      const subject = "Kode Verifikasi Email — SiM-Guru";
      const htmlBody =
        '<div style="font-family:Inter,Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px;color:#0f172a;">' +
        '<div style="background:linear-gradient(135deg,#4338CA,#4F46E5);color:white;padding:20px 24px;border-radius:14px 14px 0 0;">' +
        '<div style="font-size:13px;opacity:0.85;letter-spacing:0.05em;">SiM-GURU</div>' +
        '<div style="font-size:20px;font-weight:800;margin-top:4px;">Verifikasi Alamat Email</div>' +
        "</div>" +
        '<div style="border:1px solid #e2e8f0;border-top:0;border-radius:0 0 14px 14px;padding:24px;background:#fff;">' +
        '<p style="font-size:14px;color:#334155;margin:0 0 16px;">Halo <strong>' +
        _escHtml_(user.full_name || "Pengguna") +
        "</strong>,</p>" +
        '<p style="font-size:14px;color:#334155;margin:0 0 16px;">Anda telah meminta untuk mengubah alamat email akun SiM-Guru. Gunakan kode verifikasi berikut:</p>' +
        '<div style="background:#EEF2FF;border:1px dashed #6366F1;border-radius:12px;padding:18px;text-align:center;margin:16px 0;">' +
        "<div style=\"font-family:'JetBrains Mono',monospace;font-size:32px;font-weight:800;letter-spacing:0.4em;color:#4338CA;\">" +
        code +
        "</div>" +
        "</div>" +
        '<p style="font-size:13px;color:#64748b;margin:16px 0 8px;">⏱️ Kode berlaku <strong>10 menit</strong>.</p>' +
        '<p style="font-size:13px;color:#64748b;margin:0 0 16px;">Jika Anda tidak meminta perubahan ini, abaikan email ini.</p>' +
        '<hr style="border:0;border-top:1px solid #e2e8f0;margin:20px 0;">' +
        '<p style="font-size:11px;color:#94a3b8;margin:0;">Email otomatis dari SiM-Guru. Jangan balas email ini.</p>' +
        "</div>" +
        "</div>";
      const textBody =
        "Kode verifikasi email SiM-Guru Anda: " +
        code +
        "\n\nKode berlaku 10 menit. Jika Anda tidak meminta perubahan ini, abaikan email ini.";
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: textBody,
        htmlBody: htmlBody,
        name: "SiM-Guru",
      });
    } catch (mailErr) {
      props.deleteProperty(dataKey);
      props.deleteProperty(rateKey);
      return {
        status: "error",
        message:
          "Gagal mengirim email verifikasi: " +
          (mailErr && mailErr.message ? mailErr.message : mailErr),
      };
    }
    return {
      status: "success",
      message:
        "Kode OTP telah dikirim ke " + email + ". Kode berlaku 10 menit.",
      expires_in: 600,
      cooldown: SECS,
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}
function verifyEmailOTP(token, newEmail, otpCode) {
  try {
    const user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali" };
    const email = String(newEmail || "")
      .trim()
      .toLowerCase();
    const code = String(otpCode || "").trim();
    if (!_isValidEmail_(email))
      return { status: "error", message: "Format email tidak valid." };
    if (!/^\d{6}$/.test(code))
      return { status: "error", message: "Kode OTP harus 6 digit angka." };
    const props = PropertiesService.getScriptProperties();
    const dataKey = "otp_data_" + user.id;
    const raw = props.getProperty(dataKey);
    if (!raw)
      return {
        status: "error",
        message: "Tidak ada permintaan OTP aktif. Minta kode baru.",
      };
    let data;
    try {
      data = JSON.parse(raw);
    } catch (e) {
      props.deleteProperty(dataKey);
      return { status: "error", message: "Data OTP rusak. Minta kode baru." };
    }
    if (Date.now() > Number(data.expires || 0)) {
      props.deleteProperty(dataKey);
      return {
        status: "error",
        message: "Kode OTP sudah kedaluwarsa. Minta kode baru.",
      };
    }
    if (String(data.email).toLowerCase() !== email) {
      return {
        status: "error",
        message:
          "Email yang Anda masukkan berbeda dengan permintaan OTP terakhir. Minta kode baru.",
      };
    }
    if (Number(data.attempts || 0) >= 5) {
      props.deleteProperty(dataKey);
      return {
        status: "error",
        message: "Terlalu banyak percobaan salah. Minta kode baru.",
      };
    }
    if (String(data.code) !== code) {
      data.attempts = Number(data.attempts || 0) + 1;
      props.setProperty(dataKey, JSON.stringify(data));
      const remain = 5 - data.attempts;
      return {
        status: "error",
        message:
          "Kode OTP salah. " +
          (remain > 0 ? "Sisa percobaan: " + remain : "Percobaan habis."),
      };
    }
    const sheetData = getData("Users");
    const userRowIndex = sheetData.findIndex(
      (u) => String(u.id) === String(user.id),
    );
    if (userRowIndex === -1)
      return { status: "error", message: "User tidak ditemukan" };
    const otherWithEmail = sheetData.find(
      (u) =>
        String(u.email || "").toLowerCase() === email &&
        String(u.id) !== String(user.id),
    );
    if (otherWithEmail) {
      props.deleteProperty(dataKey);
      return {
        status: "error",
        message: "Email ini sudah digunakan oleh akun lain.",
      };
    }
    const sheetObj = getSheet("Users");
    const emailCol = _ensureUserColumn_(sheetObj, "email", 12);
    sheetObj.getRange(userRowIndex + 2, emailCol).setValue(email);
    props.deleteProperty(dataKey);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Email berhasil diverifikasi & disimpan.",
      email: email,
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}
function _isValidEmail_(s) {
  if (!s) return false;
  return /^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$/.test(String(s));
}
function _escHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
function _notifGetEmail_(userId) {
  if (!userId) return null;
  try {
    var users = getData("Users");
    var u = users.find(function (x) {
      return String(x.id) === String(userId);
    });
    if (!u) return null;
    var email = String(u.email || "")
      .trim()
      .toLowerCase();
    if (!email || !_isValidEmail_(email)) return null;
    return email;
  } catch (_) {
    return null;
  }
}
function _notifGetName_(userId) {
  try {
    var u = getData("Users").find(function (x) {
      return String(x.id) === String(userId);
    });
    return u ? String(u.full_name || u.username || "Pengguna") : "Pengguna";
  } catch (_) {
    return "Pengguna";
  }
}
function _notifBuildHtml_(opts) {
  opts = opts || {};
  var color = opts.accent || "#4F46E5";
  var color2 = opts.accent2 || "#4338CA";
  var schoolName = "";
  try {
    var cfg = _getConfigMap();
    schoolName = String(cfg["school_name"] || cfg["app_name"] || "");
  } catch (_) {}
  var badgesHtml = "";
  if (opts.badges && opts.badges.length) {
    badgesHtml =
      '<div style="display:flex;flex-wrap:wrap;gap:6px;margin:12px 0 4px;">' +
      opts.badges
        .map(function (b) {
          return (
            '<span style="display:inline-block;background:#EEF2FF;color:#4338CA;border:1px solid #C7D2FE;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:600;">' +
            _escHtml_(b) +
            "</span>"
          );
        })
        .join("") +
      "</div>";
  }
  var parasHtml = "";
  if (opts.paragraphs && opts.paragraphs.length) {
    parasHtml = opts.paragraphs
      .map(function (p) {
        return (
          '<p style="font-size:14px;color:#334155;line-height:1.55;margin:10px 0;">' +
          p +
          "</p>"
        );
      })
      .join("");
  }
  var ctaHtml = "";
  if (opts.ctaText && opts.ctaUrl) {
    ctaHtml =
      '<div style="text-align:center;margin:18px 0 4px;">' +
      '<a href="' +
      _escHtml_(opts.ctaUrl) +
      '" style="display:inline-block;background:linear-gradient(135deg,' +
      color2 +
      "," +
      color +
      ');color:white;text-decoration:none;padding:10px 20px;border-radius:10px;font-weight:700;font-size:13px;">' +
      _escHtml_(opts.ctaText) +
      "</a></div>";
  }
  return (
    "" +
    '<div style="font-family:Inter,Arial,sans-serif;max-width:560px;margin:0 auto;padding:0;color:#0f172a;background:#F8FAFC;">' +
    '<div style="padding:24px 16px;">' +
    '<div style="background:linear-gradient(135deg,' +
    color2 +
    "," +
    color +
    ');color:white;padding:20px 24px;border-radius:14px 14px 0 0;">' +
    '<div style="font-size:11px;opacity:0.85;letter-spacing:0.08em;font-weight:600;text-transform:uppercase;">SiM-GURU' +
    (schoolName ? " · " + _escHtml_(schoolName) : "") +
    "</div>" +
    '<div style="font-size:18px;font-weight:800;margin-top:6px;line-height:1.25;">' +
    _escHtml_(opts.title || "Notifikasi") +
    "</div>" +
    "</div>" +
    '<div style="border:1px solid #E2E8F0;border-top:0;border-radius:0 0 14px 14px;padding:22px 24px;background:#fff;">' +
    (opts.name
      ? '<p style="font-size:14px;color:#334155;margin:0 0 14px;">Halo <strong>' +
        _escHtml_(opts.name) +
        "</strong>,</p>"
      : "") +
    (opts.intro
      ? '<p style="font-size:14px;color:#334155;line-height:1.55;margin:0 0 12px;">' +
        opts.intro +
        "</p>"
      : "") +
    badgesHtml +
    parasHtml +
    ctaHtml +
    '<hr style="border:0;border-top:1px solid #E2E8F0;margin:18px 0 12px;">' +
    '<p style="font-size:11px;color:#94A3B8;margin:0;">Email otomatis dari SiM-Guru. Mohon tidak membalas email ini. Jika ada pertanyaan, hubungi admin sekolah.</p>' +
    "</div>" +
    "</div>" +
    "</div>"
  );
}
function _notifSend_(toEmail, subject, htmlBody, textBody) {
  if (!toEmail || !_isValidEmail_(toEmail)) return false;
  try {
    MailApp.sendEmail({
      to: toEmail,
      subject: subject,
      body:
        textBody || subject + "\n\n" + "Buka aplikasi SiM-Guru untuk detail.",
      htmlBody: htmlBody,
      name: "SiM-Guru",
    });
    return true;
  } catch (e) {
    try {
      console.warn("[notif] gagal kirim ke " + toEmail + ": " + e);
    } catch (_) {}
    return false;
  }
}
function _notifToUser_(userId, subject, htmlBody, textBody) {
  var email = _notifGetEmail_(userId);
  if (!email) return false;
  return _notifSend_(email, subject, htmlBody, textBody);
}
function _notifIsHoliday_(dateStr) {
  try {
    var ds =
      dateStr || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    var holidays = getData("Academic_Calendar") || [];
    return holidays.some(function (h) {
      try {
        var hStr = Utilities.formatDate(
          new Date(h.date),
          "Asia/Jakarta",
          "yyyy-MM-dd",
        );
        return hStr === ds;
      } catch (_) {
        return false;
      }
    });
  } catch (_) {
    return false;
  }
}
function _notifIsKbmEnabled_(dateStr) {
  try {
    var ds =
      dateStr || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    if (_notifIsHoliday_(ds)) return false;
    return true;
  } catch (e) {
    return true;
  }
}
function _notifIsTeachingDay_(dateStr) {
  try {
    var ds =
      dateStr || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    if (_notifIsHoliday_(ds)) return false;
    if (_notifIsExamEnabled_(ds)) return false;
    return true;
  } catch (e) {
    return true;
  }
}
function _notifIsExamEnabled_(dateStr) {
  try {
    var ds =
      dateStr || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    if (_notifIsHoliday_(ds)) return false;
    var periods = getData(EXAM_SHEET.PERIODS) || [];
    return periods.some(function (p) {
      var ps = String(p.date_start || "");
      var pe = String(p.date_end || "");
      return ps && pe && ds >= ps && ds <= pe;
    });
  } catch (_) {
    return false;
  }
}
function _notifPicketConfirmed_(userId, opts) {
  if (!_notifIsKbmEnabled_()) return;
  opts = opts || {};
  var byRole = String(opts.byRole || "Admin");
  var byName = String(opts.byName || "").trim();
  var byLabel = byRole + (byName ? " (" + byName + ")" : "");
  var name = _notifGetName_(userId);
  var dateStr = _notifTodayID_();
  var html = _notifBuildHtml_({
    title: "✅ Kehadiran Piket Anda Dikonfirmasi",
    accent: "#3B82F6",
    accent2: "#1D4ED8",
    name: name,
    intro:
      "Kehadiran Anda sebagai <strong>Guru Piket</strong> hari ini telah dikonfirmasi oleh <strong>" +
      _escHtml_(byLabel) +
      "</strong>.",
    badges: ["Tugas: Guru Piket", "Tanggal: " + dateStr],
    paragraphs: [
      "JTM piket sudah masuk ke sistem honorarium Anda. Tetap pantau dashboard untuk update lainnya.",
    ],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Piket Dikonfirmasi · " + dateStr,
    html,
  );
}
function _notifPicketRevoked_(userId) {
  if (!_notifIsKbmEnabled_()) return;
  var name = _notifGetName_(userId);
  var dateStr = _notifTodayID_();
  var html = _notifBuildHtml_({
    title: "⚠️ Konfirmasi Piket Dibatalkan",
    accent: "#F59E0B",
    accent2: "#B45309",
    name: name,
    intro:
      "Konfirmasi kehadiran Anda sebagai <strong>Guru Piket</strong> hari ini telah <strong>dibatalkan</strong> oleh admin.",
    badges: ["Tugas: Guru Piket", "Tanggal: " + dateStr],
    paragraphs: [
      "JTM piket Anda telah dihapus dari sistem honorarium. Hubungi admin jika ada pertanyaan.",
    ],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Konfirmasi Piket Dibatalkan · " + dateStr,
    html,
  );
}
function _notifPicketSubstituted_(originalUserId, substituteUserId) {
  if (!_notifIsKbmEnabled_()) return;
  var dateStr = _notifTodayID_();
  var subName = _notifGetName_(substituteUserId);
  var origName = _notifGetName_(originalUserId);
  var htmlSub = _notifBuildHtml_({
    title: "📌 Anda Ditugaskan sebagai Guru Piket Pengganti",
    accent: "#8B5CF6",
    accent2: "#6D28D9",
    name: subName,
    intro:
      "Admin menugaskan Anda untuk menggantikan <strong>" +
      _escHtml_(origName) +
      "</strong> sebagai Guru Piket hari ini.",
    badges: ["Pengganti dari: " + origName, "Tanggal: " + dateStr],
    paragraphs: [
      "Mohon segera hadir di pos piket. JTM akan dihitung setelah kehadiran Anda dikonfirmasi.",
    ],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Piket Pengganti · " + dateStr,
    htmlSub,
  );
  var htmlOrig = _notifBuildHtml_({
    title: "ℹ️ Tugas Piket Digantikan",
    accent: "#64748B",
    accent2: "#334155",
    name: origName,
    intro:
      "Tugas piket Anda hari ini telah dialihkan ke <strong>" +
      _escHtml_(subName) +
      "</strong> oleh admin.",
    badges: ["Diganti oleh: " + subName, "Tanggal: " + dateStr],
  });
  _notifToUser_(
    originalUserId,
    "[SiM-Guru] Tugas Piket Digantikan · " + dateStr,
    htmlOrig,
  );
}
function _notifPicketSubCancelled_(substituteUserId) {
  if (!_notifIsKbmEnabled_()) return;
  var name = _notifGetName_(substituteUserId);
  var dateStr = _notifTodayID_();
  var html = _notifBuildHtml_({
    title: "🚫 Tugas Piket Pengganti Dibatalkan",
    accent: "#F97316",
    accent2: "#C2410C",
    name: name,
    intro:
      "Tugas Anda sebagai <strong>Guru Piket Pengganti</strong> hari ini telah <strong>dibatalkan</strong> oleh admin.",
    badges: ["Tanggal: " + dateStr],
    paragraphs: [
      "JTM yang sebelumnya tercatat sudah dihapus. Tidak ada tindakan yang perlu Anda lakukan.",
    ],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengganti Piket Dibatalkan · " + dateStr,
    html,
  );
}
function _notifKbmConfirmed_(userId, opts) {
  if (!_notifIsTeachingDay_()) return;
  opts = opts || {};
  var name = _notifGetName_(userId);
  var dateStr = _notifTodayID_();
  var byName = opts.byName || "Guru Piket / Admin";
  var html = _notifBuildHtml_({
    title: "✅ Kehadiran Mengajar Dikonfirmasi",
    accent: "#10B981",
    accent2: "#047857",
    name: name,
    intro:
      "Kehadiran Anda sebagai Guru Mata Pelajaran hari ini telah dikonfirmasi oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Mengajar Dikonfirmasi · " + dateStr,
    html,
  );
}
function _notifKbmRevoked_(userId, opts) {
  if (!_notifIsTeachingDay_()) return;
  opts = opts || {};
  var name = _notifGetName_(userId);
  var dateStr = _notifTodayID_();
  var byName = opts.byName || "Guru Piket / Admin";
  var html = _notifBuildHtml_({
    title: "⚠️ Konfirmasi Kehadiran Mengajar Dibatalkan",
    accent: "#F59E0B",
    accent2: "#B45309",
    name: name,
    intro:
      "Konfirmasi kehadiran mengajar Anda hari ini telah <strong>dibatalkan</strong> oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr],
    paragraphs: ["Mohon segera tindak lanjuti."],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Mengajar Dibatalkan · " + dateStr,
    html,
  );
}
function _notifKbmSubstituted_(
  originalUserId,
  substituteUserId,
  scheduleInfo,
  byName,
) {
  if (!_notifIsTeachingDay_()) return;
  var dateStr = _notifTodayID_();
  var origName = _notifGetName_(originalUserId);
  var subName = _notifGetName_(substituteUserId);
  byName = byName || "Guru Piket / Admin";
  var info = scheduleInfo ? " · " + scheduleInfo : "";
  var htmlSub = _notifBuildHtml_({
    title: "📌 Anda Ditugaskan sebagai Guru Pengganti KBM",
    accent: "#8B5CF6",
    accent2: "#6D28D9",
    name: subName,
    intro:
      "<strong>" +
      _escHtml_(byName) +
      "</strong> menugaskan Anda untuk menggantikan <strong>" +
      _escHtml_(origName) +
      "</strong> dalam jadwal mengajar.",
    badges: ["Pengganti dari: " + origName, "Tanggal: " + dateStr + info],
    paragraphs: ["JTM akan dihitung setelah Anda mengisi jurnal mengajar."],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengganti KBM · " + dateStr,
    htmlSub,
  );
  var htmlOrig = _notifBuildHtml_({
    title: "ℹ️ Jadwal Mengajar Digantikan",
    accent: "#64748B",
    accent2: "#334155",
    name: origName,
    intro:
      "Salah satu jadwal mengajar Anda hari ini dialihkan ke <strong>" +
      _escHtml_(subName) +
      "</strong> oleh " +
      _escHtml_(byName) +
      ".",
    badges: ["Diganti oleh: " + subName, "Tanggal: " + dateStr + info],
  });
  _notifToUser_(
    originalUserId,
    "[SiM-Guru] Jadwal Mengajar Digantikan · " + dateStr,
    htmlOrig,
  );
}
function _notifKbmSubCancelled_(substituteUserId, scheduleInfo, byName) {
  if (!_notifIsTeachingDay_()) return;
  var name = _notifGetName_(substituteUserId);
  var dateStr = _notifTodayID_();
  byName = byName || "Guru Piket / Admin";
  var info = scheduleInfo ? " · " + scheduleInfo : "";
  var html = _notifBuildHtml_({
    title: "🚫 Tugas Pengganti KBM Dibatalkan",
    accent: "#F97316",
    accent2: "#C2410C",
    name: name,
    intro:
      "Tugas Anda sebagai pengganti KBM telah <strong>dibatalkan</strong> oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr + info],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengganti KBM Dibatalkan · " + dateStr,
    html,
  );
}
function _notifAutoJurnalLiburGenerated_(
  generatedLogs,
  dateStr,
  holidayDesc,
  actorName,
) {
  var summary = { sent: 0, failed: 0, recipients: 0 };
  if (!Array.isArray(generatedLogs) || generatedLogs.length === 0)
    return summary;
  var byUser = {};
  generatedLogs.forEach(function (l) {
    var uid = String(l.user_id || "").trim();
    if (!uid) return;
    if (!byUser[uid]) byUser[uid] = [];
    byUser[uid].push(l);
  });
  var displayDate = _notifFmtDateLong_(dateStr);
  var holidayLabel = String(holidayDesc || "Hari Libur").trim();
  var actor = String(actorName || "Sistem").trim();
  Object.keys(byUser).forEach(function (uid) {
    var items = byUser[uid];
    if (!items || !items.length) return;
    var name = _notifGetName_(uid) || "Bapak/Ibu Guru";
    var teachingItems = [];
    var picketItems = [];
    var ceremonyItems = [];
    items.forEach(function (it) {
      var sid = String(it.schedule_id || "");
      if (sid === "PICKET-DUTY") picketItems.push(it);
      else if (sid === "CEREMONY-DUTY") ceremonyItems.push(it);
      else teachingItems.push(it);
    });
    var totalJtm = items.reduce(function (s, it) {
      return s + (Number(it.jtm_val) || 0);
    }, 0);
    var paragraphs = [];
    if (teachingItems.length) {
      var rowsHtml = teachingItems
        .map(function (it) {
          return (
            '<li style="margin:4px 0;color:#334155;font-size:13px;">' +
            "<strong>" +
            _escHtml_(it.subject || "-") +
            "</strong>" +
            " · Kelas <strong>" +
            _escHtml_(it.class_name || "-") +
            "</strong>" +
            ' · <span style="color:#6366F1;font-weight:700;">' +
            (Number(it.jtm_val) || 0) +
            " JTM</span>" +
            "</li>"
          );
        })
        .join("");
      paragraphs.push(
        "📚 <strong>Jurnal Mengajar Otomatis (" +
          teachingItems.length +
          " jadwal):</strong>" +
          '<ul style="margin:8px 0 4px 18px;padding:0;">' +
          rowsHtml +
          "</ul>",
      );
    }
    if (picketItems.length) {
      paragraphs.push(
        "🛡️ <strong>Tugas Piket:</strong> Kehadiran Anda sebagai <strong>Guru Piket</strong> hari ini " +
          "telah dikonfirmasi otomatis oleh sistem " +
          '(<span style="color:#10B981;font-weight:700;">+' +
          4 * picketItems.length +
          " JTM</span>).",
      );
    }
    if (ceremonyItems.length) {
      paragraphs.push(
        "🎌 <strong>Tugas Pembina Upacara:</strong> Anda terdaftar sebagai Pembina Upacara dan " +
          "kehadirannya telah dikonfirmasi otomatis " +
          '(<span style="color:#F59E0B;font-weight:700;">+' +
          5 * ceremonyItems.length +
          " JTM</span>).",
      );
    }
    paragraphs.push(
      "✨ Semua entri di atas sudah masuk ke <strong>jurnal mengajar</strong> dan akan terhitung pada " +
        "<strong>honorarium bulan ini</strong>. Anda tidak perlu mengisi jurnal manual untuk tanggal ini.",
    );
    paragraphs.push(
      '<span style="color:#64748B;font-size:12px;">Dibuat oleh: ' +
        _escHtml_(actor) +
        "</span>",
    );
    var html = _notifBuildHtml_({
      title: "🎉 Bonus JTM Libur — Jurnal Otomatis",
      accent: "#7C3AED",
      accent2: "#5B21B6",
      name: name,
      intro:
        'Hari ini adalah <strong>hari libur bonus</strong> ("' +
        _escHtml_(holidayLabel) +
        '"). ' +
        "Sistem telah <strong>otomatis membuat jurnal</strong> untuk Anda sehingga tidak perlu input manual.",
      badges: ["Tanggal: " + displayDate, "Total JTM Bonus: " + totalJtm],
      paragraphs: paragraphs,
    });
    var textBody =
      "Halo " +
      name +
      ",\n\n" +
      'Hari ini hari libur bonus ("' +
      holidayLabel +
      '"). Sistem otomatis membuat jurnal untuk Anda.\n\n' +
      "Tanggal: " +
      displayDate +
      "\n" +
      "Total JTM Bonus: " +
      totalJtm +
      "\n" +
      (teachingItems.length
        ? "Jadwal mengajar otomatis: " + teachingItems.length + "\n"
        : "") +
      (picketItems.length
        ? "Tugas Piket: dikonfirmasi otomatis (+" +
          4 * picketItems.length +
          " JTM)\n"
        : "") +
      (ceremonyItems.length
        ? "Pembina Upacara: dikonfirmasi otomatis (+" +
          5 * ceremonyItems.length +
          " JTM)\n"
        : "") +
      "\nDibuat oleh: " +
      actor +
      "\n\nSemua entri ini sudah masuk jurnal dan akan terhitung pada honorarium bulan ini.";
    summary.recipients++;
    var ok = _notifToUser_(
      uid,
      "[SiM-Guru] 🎉 Bonus JTM Libur · " + displayDate,
      html,
      textBody,
    );
    if (ok) summary.sent++;
    else summary.failed++;
  });
  return summary;
}
function _notifCommitteeConfirmed_(userId, dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  var name = _notifGetName_(userId);
  dateStr = _notifFmtDateLong_(dateStr);
  var html = _notifBuildHtml_({
    title: "✅ Kehadiran Panitia Ujian Dikonfirmasi",
    accent: "#0891B2",
    accent2: "#0369A1",
    name: name,
    intro:
      "Kehadiran Anda sebagai <strong>Panitia Ujian</strong> telah dikonfirmasi oleh admin.",
    badges: ["Tanggal: " + dateStr],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Panitia Ujian Dikonfirmasi · " + dateStr,
    html,
  );
}
function _notifCommitteeRevoked_(userId, dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  var name = _notifGetName_(userId);
  dateStr = _notifFmtDateLong_(dateStr);
  var html = _notifBuildHtml_({
    title: "⚠️ Konfirmasi Panitia Ujian Dibatalkan",
    accent: "#F59E0B",
    accent2: "#B45309",
    name: name,
    intro:
      "Konfirmasi kehadiran Anda sebagai Panitia Ujian telah <strong>dibatalkan</strong> oleh admin.",
    badges: ["Tanggal: " + dateStr],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Konfirmasi Panitia Ujian Dibatalkan · " + dateStr,
    html,
  );
}
function _notifCommitteeSubstituted_(
  originalUserId,
  substituteUserId,
  dateStr,
) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  dateStr = _notifFmtDateLong_(dateStr);
  var origName = _notifGetName_(originalUserId);
  var subName = _notifGetName_(substituteUserId);
  var htmlSub = _notifBuildHtml_({
    title: "📌 Anda Ditugaskan sebagai Panitia Ujian Pengganti",
    accent: "#8B5CF6",
    accent2: "#6D28D9",
    name: subName,
    intro:
      "Admin menugaskan Anda menggantikan <strong>" +
      _escHtml_(origName) +
      "</strong> sebagai Panitia Ujian.",
    badges: ["Pengganti dari: " + origName, "Tanggal: " + dateStr],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Panitia Ujian Pengganti · " + dateStr,
    htmlSub,
  );
  var htmlOrig = _notifBuildHtml_({
    title: "ℹ️ Tugas Panitia Ujian Digantikan",
    accent: "#64748B",
    accent2: "#334155",
    name: origName,
    intro:
      "Tugas Panitia Ujian Anda telah dialihkan ke <strong>" +
      _escHtml_(subName) +
      "</strong> oleh admin.",
    badges: ["Diganti oleh: " + subName, "Tanggal: " + dateStr],
  });
  _notifToUser_(
    originalUserId,
    "[SiM-Guru] Tugas Panitia Ujian Digantikan · " + dateStr,
    htmlOrig,
  );
}
function _notifCommitteeSubCancelled_(substituteUserId, dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  dateStr = _notifFmtDateLong_(dateStr);
  var name = _notifGetName_(substituteUserId);
  var html = _notifBuildHtml_({
    title: "🚫 Tugas Panitia Ujian Pengganti Dibatalkan",
    accent: "#F97316",
    accent2: "#C2410C",
    name: name,
    intro:
      "Tugas Anda sebagai Panitia Ujian Pengganti telah <strong>dibatalkan</strong> oleh admin.",
    badges: ["Tanggal: " + dateStr],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Panitia Pengganti Dibatalkan · " + dateStr,
    html,
  );
}
function _notifSupervisorConfirmed_(userId, dateStr, byName, sessionInfo) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  var name = _notifGetName_(userId);
  dateStr = _notifFmtDateLong_(dateStr);
  byName = byName || "Panitia Ujian / Admin";
  var html = _notifBuildHtml_({
    title: "✅ Kehadiran Pengawas Ruang Ujian Dikonfirmasi",
    accent: "#10B981",
    accent2: "#047857",
    name: name,
    intro:
      "Kehadiran Anda sebagai <strong>Pengawas Ruang Ujian</strong> telah dikonfirmasi oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr].concat(sessionInfo ? [sessionInfo] : []),
    paragraphs: ["Mohon segera mengisi BAP setelah sesi berakhir."],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Pengawas Dikonfirmasi · " + dateStr,
    html,
  );
}
function _notifSupervisorRevoked_(userId, dateStr, byName, sessionInfo) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  var name = _notifGetName_(userId);
  dateStr = _notifFmtDateLong_(dateStr);
  byName = byName || "Panitia Ujian / Admin";
  var html = _notifBuildHtml_({
    title: "⚠️ Konfirmasi Pengawas Ruang Ujian Dibatalkan",
    accent: "#F59E0B",
    accent2: "#B45309",
    name: name,
    intro:
      "Konfirmasi kehadiran Anda sebagai Pengawas Ruang Ujian telah <strong>dibatalkan</strong> oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr].concat(sessionInfo ? [sessionInfo] : []),
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Konfirmasi Pengawas Dibatalkan · " + dateStr,
    html,
  );
}
function _notifSupervisorSubstituted_(
  originalUserId,
  substituteUserId,
  dateStr,
  byName,
  sessionInfo,
) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  dateStr = _notifFmtDateLong_(dateStr);
  byName = byName || "Panitia Ujian / Admin";
  var origName = _notifGetName_(originalUserId);
  var subName = _notifGetName_(substituteUserId);
  var htmlSub = _notifBuildHtml_({
    title: "📌 Anda Ditugaskan sebagai Pengawas Ruang Ujian Pengganti",
    accent: "#8B5CF6",
    accent2: "#6D28D9",
    name: subName,
    intro:
      "<strong>" +
      _escHtml_(byName) +
      "</strong> menugaskan Anda menggantikan <strong>" +
      _escHtml_(origName) +
      "</strong> sebagai Pengawas Ruang Ujian.",
    badges: ["Pengganti dari: " + origName, "Tanggal: " + dateStr].concat(
      sessionInfo ? [sessionInfo] : [],
    ),
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengawas Pengganti · " + dateStr,
    htmlSub,
  );
  var htmlOrig = _notifBuildHtml_({
    title: "ℹ️ Tugas Pengawas Ruang Ujian Digantikan",
    accent: "#64748B",
    accent2: "#334155",
    name: origName,
    intro:
      "Tugas Pengawas Ruang Ujian Anda dialihkan ke <strong>" +
      _escHtml_(subName) +
      "</strong> oleh " +
      _escHtml_(byName) +
      ".",
    badges: ["Diganti oleh: " + subName, "Tanggal: " + dateStr].concat(
      sessionInfo ? [sessionInfo] : [],
    ),
  });
  _notifToUser_(
    originalUserId,
    "[SiM-Guru] Tugas Pengawas Digantikan · " + dateStr,
    htmlOrig,
  );
}
function _notifSupervisorSubCancelled_(
  substituteUserId,
  dateStr,
  byName,
  sessionInfo,
) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  dateStr = _notifFmtDateLong_(dateStr);
  byName = byName || "Panitia Ujian / Admin";
  var name = _notifGetName_(substituteUserId);
  var html = _notifBuildHtml_({
    title: "🚫 Tugas Pengawas Pengganti Dibatalkan",
    accent: "#F97316",
    accent2: "#C2410C",
    name: name,
    intro:
      "Tugas Anda sebagai Pengawas Ruang Ujian Pengganti telah <strong>dibatalkan</strong> oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr].concat(sessionInfo ? [sessionInfo] : []),
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengawas Pengganti Dibatalkan · " + dateStr,
    html,
  );
}
function _notifTodayID_() {
  return _notifFmtDateLong_(new Date());
}
function _notifFmtTime_(val) {
  if (val == null || val === "") return "";
  var tz = "Asia/Jakarta";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, tz, "HH:mm");
  }
  var s = String(val).trim();
  if (!s) return "";
  var m = s.match(/(\d{1,2}):(\d{1,2})/);
  if (m) {
    return ("0" + m[1]).slice(-2) + ":" + ("0" + m[2]).slice(-2);
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, "HH:mm");
  return s;
}
function _notifResolveDate_(value) {
  var tz = "Asia/Jakarta";
  if (value == null || value === "") return null;
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : { d: value, hasTime: true };
  }
  var s = String(value).trim();
  if (!s) return null;
  var dateOnly = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dateOnly) {
    var d = new Date(
      parseInt(dateOnly[1], 10),
      parseInt(dateOnly[2], 10) - 1,
      parseInt(dateOnly[3], 10),
    );
    return { d: d, hasTime: false };
  }
  var dateTime = s.match(
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/,
  );
  if (dateTime) {
    var d2 = new Date(
      parseInt(dateTime[1], 10),
      parseInt(dateTime[2], 10) - 1,
      parseInt(dateTime[3], 10),
      parseInt(dateTime[4], 10),
      parseInt(dateTime[5], 10),
      dateTime[6] ? parseInt(dateTime[6], 10) : 0,
    );
    return { d: d2, hasTime: true };
  }
  var fallback = new Date(s);
  if (!isNaN(fallback.getTime())) {
    return { d: fallback, hasTime: true };
  }
  return null;
}
function _notifFmtDate_(value) {
  var tz = "Asia/Jakarta";
  var info = _notifResolveDate_(value);
  if (!info) return String(value || "");
  var months = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ];
  var dd = Utilities.formatDate(info.d, tz, "d");
  var monthIdx = parseInt(Utilities.formatDate(info.d, tz, "M"), 10) - 1;
  var yyyy = Utilities.formatDate(info.d, tz, "yyyy");
  return ("0" + dd).slice(-2) + " " + (months[monthIdx] || "") + " " + yyyy;
}
function _notifFmtDateLong_(value) {
  var tz = "Asia/Jakarta";
  var info = _notifResolveDate_(value);
  if (!info) return String(value || "");
  var days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  var months = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ];
  var dayIdx = parseInt(Utilities.formatDate(info.d, tz, "u"), 10) % 7;
  var dd = Utilities.formatDate(info.d, tz, "d");
  var monthIdx = parseInt(Utilities.formatDate(info.d, tz, "M"), 10) - 1;
  var yyyy = Utilities.formatDate(info.d, tz, "yyyy");
  return (
    (days[dayIdx] || "") +
    ", " +
    ("0" + dd).slice(-2) +
    " " +
    (months[monthIdx] || "") +
    " " +
    yyyy
  );
}
function _notifFmtDateTime_(value) {
  var tz = "Asia/Jakarta";
  var info = _notifResolveDate_(value);
  if (!info) return String(value || "");
  var datePart = _notifFmtDate_(value);
  if (!info.hasTime) return datePart;
  var hhmm = Utilities.formatDate(info.d, tz, "HH:mm");
  return datePart + " " + hhmm;
}
function _notifFmtDateLongTime_(value) {
  var info = _notifResolveDate_(value);
  if (!info) return String(value || "");
  var datePart = _notifFmtDateLong_(value);
  if (!info.hasTime) return datePart;
  var hhmm = Utilities.formatDate(info.d, "Asia/Jakarta", "HH:mm");
  return datePart + " " + hhmm;
}
function _checkPasswordComplexity(password, fullName, username) {
  if (password.length < 8) {
    return { valid: false, message: "Password baru minimal harus 8 karakter." };
  }
  if (!/[A-Z]/.test(password)) {
    return { valid: false, message: "Password harus mengandung huruf besar." };
  }
  if (!/[a-z]/.test(password)) {
    return { valid: false, message: "Password harus mengandung huruf kecil." };
  }
  if (!/[0-9]/.test(password)) {
    return { valid: false, message: "Password harus mengandung angka." };
  }
  if (!/[!@#$%^&*(),.?":{}|<>]/.test(password)) {
    return {
      valid: false,
      message: "Password harus mengandung karakter khusus (simbol).",
    };
  }
  const lowerPass = password.toLowerCase();
  if (username && lowerPass.includes(username.toLowerCase())) {
    return {
      valid: false,
      message: "Password tidak boleh mengandung username Anda.",
    };
  }
  if (fullName) {
    const nameParts = fullName.split(" ").filter((part) => part.length > 2);
    for (let part of nameParts) {
      if (lowerPass.includes(part.toLowerCase())) {
        return {
          valid: false,
          message:
            "Password tidak boleh mengandung unsur nama Anda (" + part + ").",
        };
      }
    }
  }
  return { valid: true };
}
function _extUsersSheet_() {
  return SpreadsheetApp.openById(EXT_USERS_SS_ID).getSheetByName(
    EXT_USERS_SHEET,
  );
}
function _extFindUserRow_(username) {
  const sheet = _extUsersSheet_();
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(username).trim()) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}
function _extGetPassword_(username) {
  const found = _extFindUserRow_(username);
  return found ? String(found.row[2]) : null;
}
function _extSetPassword_(username, newPassword) {
  const sheet = _extUsersSheet_();
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(username).trim()) {
      sheet.getRange(i + 1, 3).setValue(newPassword);
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}
function _extAddUser_(username, password) {
  const sheet = _extUsersSheet_();
  if (!sheet) return false;
  sheet.appendRow(["", String(username), String(password)]);
  SpreadsheetApp.flush();
  return true;
}
function _extUpdateUsername_(oldUsername, newUsername) {
  const sheet = _extUsersSheet_();
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(oldUsername).trim()) {
      sheet.getRange(i + 1, 2).setValue(String(newUsername));
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}
function _extDeleteUser_(username) {
  const sheet = _extUsersSheet_();
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim() === String(username).trim()) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}
function getAllSchedulesAdmin(token) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") return [];
  const rawSchedules = getData("Schedules");
  const teachers = getData("Users");
  return rawSchedules.map((s) => {
    const teacher = teachers.find(
      (t) => String(t.id).trim() === String(s.user_id).trim(),
    );
    let tStart = s.time_start;
    let tEnd = s.time_end;
    if (tStart instanceof Date)
      tStart = Utilities.formatDate(tStart, "Asia/Jakarta", "HH:mm");
    if (tEnd instanceof Date)
      tEnd = Utilities.formatDate(tEnd, "Asia/Jakarta", "HH:mm");
    const days = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    return {
      id: s.id,
      user_id: s.user_id,
      guru_name: teacher ? teacher.full_name : "Unknown",
      day_name: days[Number(s.day_index)] || "-",
      day_index: s.day_index,
      time_start: String(tStart).substring(0, 5),
      time_end: String(tEnd).substring(0, 5),
      subject: s.subject,
      class_name: s.class_name,
      jtm_val: s.jtm_val,
      tahun_pelajaran: s.tahun_pelajaran || "",
      semester: s.semester || "",
    };
  });
}
function saveScheduleAdmin(token, data) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Akses ditolak" };
  const sheet = getSheet("Schedules");
  if (data.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 2).setValue(data.user_id);
        sheet.getRange(i + 1, 3).setValue(getDayName(data.day_index));
        sheet.getRange(i + 1, 4).setValue(data.day_index);
        sheet.getRange(i + 1, 5).setValue(data.time_start);
        sheet.getRange(i + 1, 6).setValue(data.time_end);
        sheet.getRange(i + 1, 7).setValue(data.subject);
        sheet.getRange(i + 1, 8).setValue(data.class_name);
        sheet.getRange(i + 1, 9).setValue(data.jtm_val);
        sheet.getRange(i + 1, 10).setValue(data.tahun_pelajaran || "");
        sheet.getRange(i + 1, 11).setValue(data.semester || "");
        return { status: "success" };
      }
    }
  } else {
    const newId = generateId("SCH");
    sheet.appendRow([
      newId,
      data.user_id,
      getDayName(data.day_index),
      data.day_index,
      data.time_start,
      data.time_end,
      data.subject,
      data.class_name,
      data.jtm_val,
      data.tahun_pelajaran || "",
      data.semester || "",
    ]);
  }
  return { status: "success" };
}
function deleteScheduleAdmin(token, id) {
  return deleteRowById(token, "Schedules", id);
}
function deleteMultipleSchedulesAdmin(token, ids) {
  const user = verifySession(token);
  if (!user || user.role !== "admin") {
    return { status: "error", message: "Akses ditolak." };
  }
  if (!Array.isArray(ids) || ids.length === 0) {
    return { status: "error", message: "Tidak ada data yang dipilih." };
  }
  let successCount = 0;
  for (const id of ids) {
    const res = deleteRowById(token, "Schedules", id);
    if (res && res.status === "success") {
      successCount++;
    }
  }
  return {
    status: "success",
    message: `${successCount} dari ${ids.length} jadwal berhasil dihapus.`,
  };
}
function getMyWeeklyData(token) {
  const user = verifySession(token);
  if (!user) return { status: "error" };
  const allSchedules = getData("Schedules");
  const config = {};
  (getData("Config") || []).forEach((c) => {
    config[c.key] = c.value;
  });
  const activeTP = config.tahun_pelajaran || "";
  const activeSem = config.semester || "";
  const mySchedules = allSchedules.filter((s) => {
    const isMe = String(s.user_id).trim() === String(user.id).trim();
    const sTP = s.tahun_pelajaran || activeTP;
    const sSem = s.semester || activeSem;
    return isMe && sTP === activeTP && sSem === activeSem;
  });
  const formattedSchedules = mySchedules.map((s) => {
    let tStart = s.time_start;
    let tEnd = s.time_end;
    if (tStart instanceof Date)
      tStart = Utilities.formatDate(tStart, "Asia/Jakarta", "HH:mm");
    if (tEnd instanceof Date)
      tEnd = Utilities.formatDate(tEnd, "Asia/Jakarta", "HH:mm");
    const days = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    return {
      ...s,
      day_name: days[Number(s.day_index)] || "-",
      time_start: String(tStart).substring(0, 5),
      time_end: String(tEnd).substring(0, 5),
    };
  });
  formattedSchedules.sort((a, b) => {
    if (Number(a.day_index) !== Number(b.day_index))
      return Number(a.day_index) - Number(b.day_index);
    return a.time_start.localeCompare(b.time_start);
  });
  const pickets = getData("Picket_Schedules").filter(
    (p) => String(p.user_id).trim() === String(user.id).trim(),
  );
  const picketDays = pickets.map((p) => {
    const days = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    return days[Number(p.day_index)];
  });
  const allowances = getData("Allowances").filter(
    (a) => String(a.user_id).trim() === String(user.id).trim(),
  );
  const cleanAllowances = allowances.map((a) => ({
    duty_name: a.duty_name,
    amount: formatRupiah(a.amount),
    raw_amount: Number(a.amount) || 0
  }));
  const allCeremonies = getData("Ceremony_Schedules");
  const todayDate = new Date();
  const todayStr = Utilities.formatDate(
    todayDate,
    "Asia/Jakarta",
    "yyyy-MM-dd",
  );
  const allLogs = getData("Teaching_Logs");
  const myCeremonies = allCeremonies
    .filter((c) => {
      const isMe = String(c.user_id).trim() === String(user.id).trim();
      const cTP = c.tahun_pelajaran || activeTP;
      const cSem = c.semester || activeSem;
      return isMe && cTP === activeTP && cSem === activeSem;
    })
    .map((c) => {
      const dateStr = safeDate(c.date);
      const cDate = new Date(dateStr + "T00:00:00");
      const tDate = new Date(todayStr + "T00:00:00");
      const isPast = cDate < tDate;
      const isToday = dateStr === todayStr;
      const isConfirmed = allLogs.some(
        (l) =>
          String(l.schedule_id) === "CEREMONY-DUTY" &&
          safeDate(l.date) === dateStr &&
          String(l.user_id) === String(user.id),
      );
      return {
        id: String(c.id),
        date: dateStr,
        date_formatted: formatDateIndo(dateStr),
        is_past: isPast,
        is_today: isToday,
        is_confirmed: isConfirmed,
      };
    })
    .sort((a, b) => new Date(a.date) - new Date(b.date));
  return {
    status: "success",
    schedules: formattedSchedules,
    picket_days: picketDays,
    allowances: cleanAllowances,
    ceremony_assignments: myCeremonies,
  };
}
function getDayName(index) {
  const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  return days[index] || "";
}
function saveTeachingJournal(token, data) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const scheduleId = data.schedule_id;
  const schedules = getData("Schedules");
  const targetSchedule = schedules.find(
    (s) => String(s.id) === String(scheduleId),
  );
  let expectedStd = null;
  if (targetSchedule) {
    const className = targetSchedule.class_name;
    const configData = getData("Config");
    let configMap = {};
    configData.forEach((c) => (configMap[c.key] = c.value));
    const configKey = "std_kelas_" + className;
    const dbTotal = parseInt(configMap[configKey]);
    const inputHadir = parseInt(data.siswa_hadir) || 0;
    const inputAbsen = parseInt(data.siswa_absen) || 0;
    const inputTotal = inputHadir + inputAbsen;
    if (!isNaN(dbTotal)) {
      expectedStd = dbTotal;
      if (inputTotal !== dbTotal) {
        return {
          status: "error",
          code: "STUDENT_COUNT_MISMATCH",
          expected: dbTotal,
          actual: inputTotal,
          message:
            "Validasi Gagal: Total siswa (" +
            inputTotal +
            ") tidak sesuai dengan standar kelas " +
            className +
            " (" +
            dbTotal +
            " siswa). Harap sesuaikan jumlah hadir dan absen.",
        };
      }
    }
  }
  // Validasi rekap absensi siswa dari guru piket
  if (targetSchedule) {
    const _saClassName = String(targetSchedule.class_name || "").trim();
    const _saTz = Session.getScriptTimeZone();
    const _saTodayStr = Utilities.formatDate(new Date(), _saTz, "yyyy-MM-dd");
    const _saAllData = getData("Student_Attendance");
    const _saRekap = _saAllData.find(function(r) {
      return String(r.class_name || "").trim() === _saClassName && safeDate(r.date) === _saTodayStr;
    });
    if (_saRekap) {
      const _saExpHadir = parseInt(_saRekap.hadir) || 0;
      const _saExpAbsen = (parseInt(_saRekap.sakit) || 0) + (parseInt(_saRekap.izin) || 0) + (parseInt(_saRekap.alpa) || 0);
      const _saActHadir = parseInt(data.siswa_hadir) || 0;
      const _saActAbsen = parseInt(data.siswa_absen) || 0;
      if (_saActHadir !== _saExpHadir || _saActAbsen !== _saExpAbsen) {
        return {
          status: "error",
          code: "STUDENT_ATTENDANCE_MISMATCH",
          class_name: _saClassName,
          expected_hadir: _saExpHadir,
          expected_absen: _saExpAbsen,
          actual_hadir: _saActHadir,
          actual_absen: _saActAbsen,
          rekap_detail: {
            sakit: parseInt(_saRekap.sakit) || 0,
            izin:  parseInt(_saRekap.izin)  || 0,
            alpa:  parseInt(_saRekap.alpa)  || 0
          },
          message: "Data absensi tidak sesuai rekap guru piket untuk kelas " + _saClassName +
            ". Rekap piket — Hadir: " + _saExpHadir + ", Sakit+Izin+Alpa: " + _saExpAbsen +
            ". Anda mengisi — Hadir: " + _saActHadir + ", Absen: " + _saActAbsen + "."
        };
      }
    }
  }
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  let targetDateStr = todayStr;
  if (data.date && /^\d{4}-\d{2}-\d{2}$/.test(String(data.date))) {
    targetDateStr = String(data.date);
  }
  const existingLogs = getData("Teaching_Logs");
  const isDuplicate = existingLogs.some((log) => {
    const logDate = Utilities.formatDate(new Date(log.date), tz, "yyyy-MM-dd");
    return (
      String(log.schedule_id) === String(scheduleId) &&
      String(log.user_id) === String(user.id) &&
      logDate === targetDateStr
    );
  });
  if (isDuplicate) {
    return {
      status: "error",
      code: "DUPLICATE",
      message:
        "Jurnal untuk jadwal ini sudah pernah diisi pada tanggal yang sama. Anda tidak dapat mengisi jurnal yang sama dua kali.",
    };
  }
  const sheet = getSheet("Teaching_Logs");
  const newId = generateId("LOG");
  const now = new Date();
  const timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");
  let jtmVal = targetSchedule ? targetSchedule.jtm_val : 0;
  const jtmAdjustment = _jtmReadAdjustment_("KBM", scheduleId);
  if (jtmAdjustment) {
    const adjDateStr =
      jtmAdjustment.date instanceof Date
        ? Utilities.formatDate(jtmAdjustment.date, tz, "yyyy-MM-dd")
        : String(jtmAdjustment.date);
    const storedAdjusted = jtmAdjustment.adjusted_jtm;
    if (
      adjDateStr === targetDateStr &&
      storedAdjusted !== null &&
      storedAdjusted !== undefined &&
      storedAdjusted !== "" &&
      !isNaN(Number(storedAdjusted))
    ) {
      jtmVal = Number(storedAdjusted);
    }
  }
  sheet.appendRow([
    newId,
    data.schedule_id,
    user.id,
    targetDateStr,
    data.material,
    data.siswa_hadir,
    data.siswa_absen,
    data.notes,
    jtmVal,
    timeStr,
  ]);
  _invalidateDataSnapshot();
  return {
    status: "success",
    log_id: newId,
    expected_std: expectedStd,
  };
}
function getClassExpectedStudents(token, scheduleId) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const schedules = getData("Schedules");
    const sched = schedules.find((s) => String(s.id) === String(scheduleId));
    if (!sched) return { status: "success", expected: null, class_name: null };
    const configData = getData("Config");
    let configMap = {};
    configData.forEach((c) => (configMap[c.key] = c.value));
    const dbTotal = parseInt(configMap["std_kelas_" + sched.class_name]);
    return {
      status: "success",
      expected: isNaN(dbTotal) ? null : dbTotal,
      class_name: sched.class_name,
    };
  } catch (e) {
    return { status: "error", message: "Server Error: " + e.toString() };
  }
}
function updateTeachingLog(token, payload) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const logId = payload && payload.log_id;
    if (!logId) return { status: "error", message: "Log ID wajib diisi." };
    const sheet = getSheet("Teaching_Logs");
    const range = sheet.getDataRange();
    const values = range.getValues();
    const header = values[0].map((h) => String(h).trim());
    const idCol = header.indexOf("log_id");
    const schedCol = header.indexOf("schedule_id");
    const userCol = header.indexOf("user_id");
    const dateCol = header.indexOf("date");
    const matCol = header.indexOf("materi");
    const hadirCol = header.indexOf("siswa_hadir");
    const absenCol = header.indexOf("siswa_absen");
    const notesCol = header.indexOf("notes");
    if (
      [idCol, schedCol, userCol, matCol, hadirCol, absenCol, notesCol].some(
        (c) => c < 0,
      )
    ) {
      return {
        status: "error",
        message: "Struktur sheet Teaching_Logs tidak sesuai.",
      };
    }
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol]).trim() === String(logId).trim()) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1)
      return { status: "error", message: "Jurnal tidak ditemukan." };
    const row = values[rowIndex];
    const isOwner = String(row[userCol]).trim() === String(user.id).trim();
    const isAdmin = String(user.role).toLowerCase() === "admin";
    if (!isOwner && !isAdmin) {
      return {
        status: "error",
        message: "Anda tidak memiliki akses untuk mengubah jurnal ini.",
      };
    }
    const schedId = row[schedCol];
    const schedules = getData("Schedules");
    const sched = schedules.find((s) => String(s.id) === String(schedId));
    const inputHadir = parseInt(payload.siswa_hadir) || 0;
    const inputAbsen = parseInt(payload.siswa_absen) || 0;
    const inputTotal = inputHadir + inputAbsen;
    if (sched) {
      const configData = getData("Config");
      let configMap = {};
      configData.forEach((c) => (configMap[c.key] = c.value));
      const dbTotal = parseInt(configMap["std_kelas_" + sched.class_name]);
      if (!isNaN(dbTotal) && inputTotal !== dbTotal) {
        return {
          status: "error",
          code: "STUDENT_COUNT_MISMATCH",
          expected: dbTotal,
          actual: inputTotal,
          message:
            "Validasi Gagal: Total siswa (" +
            inputTotal +
            ") tidak sesuai dengan standar kelas " +
            sched.class_name +
            " (" +
            dbTotal +
            " siswa). Harap sesuaikan jumlah hadir dan absen.",
        };
      }
    }
    // Validasi rekap absensi siswa dari guru piket
    if (sched) {
      const _saClassName2 = String(sched.class_name || "").trim();
      const _saTz2 = Session.getScriptTimeZone();
      const _saLogDate = safeDate(values[rowIndex][dateCol]);
      const _saAllData2 = getData("Student_Attendance");
      const _saRekap2 = _saAllData2.find(function(r) {
        return String(r.class_name || "").trim() === _saClassName2 && safeDate(r.date) === _saLogDate;
      });
      if (_saRekap2) {
        const _saExpHadir2 = parseInt(_saRekap2.hadir) || 0;
        const _saExpAbsen2 = (parseInt(_saRekap2.sakit) || 0) + (parseInt(_saRekap2.izin) || 0) + (parseInt(_saRekap2.alpa) || 0);
        const _saActHadir2 = parseInt(payload.siswa_hadir) || 0;
        const _saActAbsen2 = parseInt(payload.siswa_absen) || 0;
        if (_saActHadir2 !== _saExpHadir2 || _saActAbsen2 !== _saExpAbsen2) {
          return {
            status: "error",
            code: "STUDENT_ATTENDANCE_MISMATCH",
            class_name: _saClassName2,
            expected_hadir: _saExpHadir2,
            expected_absen: _saExpAbsen2,
            actual_hadir: _saActHadir2,
            actual_absen: _saActAbsen2,
            rekap_detail: {
              sakit: parseInt(_saRekap2.sakit) || 0,
              izin:  parseInt(_saRekap2.izin)  || 0,
              alpa:  parseInt(_saRekap2.alpa)  || 0
            },
            message: "Data absensi tidak sesuai rekap guru piket untuk kelas " + _saClassName2 +
              ". Rekap piket — Hadir: " + _saExpHadir2 + ", Sakit+Izin+Alpa: " + _saExpAbsen2 +
              ". Anda mengisi — Hadir: " + _saActHadir2 + ", Absen: " + _saActAbsen2 + "."
          };
        }
      }
    }
    const materi = String(payload.material || "").trim();
    if (!materi)
      return { status: "error", message: "Materi pembelajaran wajib diisi." };
    const rowNumber = rowIndex + 1;
    sheet.getRange(rowNumber, matCol + 1).setValue(materi);
    sheet.getRange(rowNumber, hadirCol + 1).setValue(inputHadir);
    sheet.getRange(rowNumber, absenCol + 1).setValue(inputAbsen);
    sheet
      .getRange(rowNumber, notesCol + 1)
      .setValue(String(payload.notes || ""));
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return { status: "success", log_id: logId };
  } catch (e) {
    return { status: "error", message: "Server Error: " + e.toString() };
  }
}
function getSettingsMasterData(token) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Akses ditolak" };
    const configRaw = getData("Config");
    let config = {};
    configRaw.forEach((c) => {
      config[c.key] = c.value;
    });
    const holidays = getData("Academic_Calendar").map((h) => ({
      id: h.event_id,
      date: h.date,
      desc: h.description,
      is_holiday: h.is_holiday,
    }));
    const users = getData("Users");
    const subjects = getData("Subjects").map((s) => {
      return {
        id: s.id,
        name: s.name,
      };
    });
    const allowancesRaw = getData("Allowances");
    const allowances = allowancesRaw.map((a) => {
      const u = users.find((usr) => String(usr.id) === String(a.user_id));
      return {
        id: a.id,
        name: a.duty_name,
        user_id: a.user_id,
        guru_name: u ? u.full_name : "Unknown",
        amount: a.amount,
      };
    });
    const result = {
      status: "success",
      config: config,
      holidays: holidays,
      subjects: subjects,
      allowances: allowances,
    };
    return JSON.parse(JSON.stringify(result));
  } catch (e) {
    return {
      status: "error",
      message: "Server Error in getSettingsMasterData: " + e.toString(),
    };
  }
}
function saveSystemConfig(token, data) {
  try {
    const user = verifySession(token);
    if (!user || user.role !== "admin")
      return { status: "error", message: "Akses hanya untuk Admin" };
    const sheet = getSheet("Config");
    const existingData = sheet.getDataRange().getValues();
    const updateKey = (key, val) => {
      let rowIndex = -1;
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][0] === key) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex > 0) {
        const cell = sheet.getRange(rowIndex, 2);
        if (key === "app_version" || key === "maintenance_start" || key === "maintenance_end") cell.setNumberFormat("@");
        cell.setValue(val);
      } else {
        sheet.appendRow([key, ""]);
        const cell = sheet.getRange(sheet.getLastRow(), 2);
        if (key === "app_version" || key === "maintenance_start" || key === "maintenance_end") cell.setNumberFormat("@");
        cell.setValue(val);
      }
    };
    for (const [key, value] of Object.entries(data)) {
      updateKey(key, value);
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return { status: "success" };
  } catch (e) {
    console.error("saveSystemConfig error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function saveSubject(token, data) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const nameTrimmed = (data.name || "").trim();
  if (!nameTrimmed)
    return { status: "error", message: "Nama mapel tidak boleh kosong." };
  const sheet = getSheet("Subjects");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toLowerCase() === nameTrimmed.toLowerCase()) {
      if (!data.id || String(rows[i][0]) !== String(data.id)) {
        return {
          status: "error",
          message: "Mata pelajaran dengan nama tersebut sudah ada.",
        };
      }
    }
  }
  if (data.id) {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2, 1, 1).setValue(nameTrimmed);
        found = true;
        break;
      }
    }
    if (!found)
      return { status: "error", message: "Data Mapel tidak ditemukan." };
  } else {
    const newId = generateId("SBJ");
    sheet.appendRow([newId, nameTrimmed]);
  }
  return { status: "success" };
}
function deleteSubject(token, id) {
  return deleteRowById(token, "Subjects", id);
}
function saveAllowance(token, data) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const sheet = getSheet("Allowances");
  if (data.id) {
    const rows = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2).setValue(data.duty_name);
        sheet.getRange(i + 1, 3).setValue(data.user_id);
        sheet.getRange(i + 1, 4).setValue(data.amount);
        found = true;
        break;
      }
    }
    if (!found)
      return { status: "error", message: "Tunjangan tidak ditemukan." };
  } else {
    const newId = generateId("ALW");
    sheet.appendRow([newId, data.duty_name, data.user_id, data.amount]);
  }
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  return { status: "success" };
}
function deleteAllowance(token, id) {
  return deleteRowById(token, "Allowances", id);
}
function addHoliday(token, date, desc, isHoliday, id) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const sheet = getSheet("Academic_Calendar");
  const isHolStr = isHoliday ? "True" : "False";
  if (id) {
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        let existingDate = String(data[i][1]).trim();
        let todayStr = Utilities.formatDate(
          new Date(),
          "Asia/Jakarta",
          "yyyy-MM-dd",
        );
        if (existingDate < todayStr) {
          return {
            status: "error",
            message: "Hari libur yang sudah lewat tidak dapat diubah",
          };
        }
        sheet.getRange(i + 1, 2).setValue(date);
        sheet.getRange(i + 1, 3).setValue(desc);
        sheet.getRange(i + 1, 4).setValue(isHolStr);
        found = true;
        break;
      }
    }
    if (!found) return { status: "error", message: "Data tidak ditemukan" };
  } else {
    const newId = generateId("HOL");
    sheet.appendRow([newId, date, desc, isHolStr]);
  }
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  return { status: "success" };
}
function deleteHoliday(token, id) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const sheet = getSheet("Academic_Calendar");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      let existingDate = String(data[i][1]).trim();
      let todayStr = Utilities.formatDate(
        new Date(),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (existingDate < todayStr) {
        return {
          status: "error",
          message: "Hari libur yang sudah lewat tidak dapat dihapus",
        };
      }
      sheet.deleteRow(i + 1);
      try {
        _invalidateDataSnapshot();
      } catch (_) {}
      return { status: "success" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}
function deleteRowById(token, sheetName, id) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      try {
        _invalidateDataSnapshot();
      } catch (_) {}
      return { status: "success" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}
function _annBool_(v) {
  if (v === true || v === 1) return true;
  if (v === false || v === 0 || v == null) return false;
  var s = String(v).trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "1";
}
function _annDateOnly_(v) {
  if (v == null) return "";
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return "";
    return Utilities.formatDate(v, "Asia/Jakarta", "yyyy-MM-dd");
  }
  var s = String(v).trim();
  if (!s) return "";
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
  }
  return "";
}
function _annParseIdList_(v) {
  if (v == null) return [];
  var s = String(v).trim();
  if (!s) return [];
  if (s.charAt(0) === "[" || s.charAt(0) === "{") {
    try {
      var j = JSON.parse(s);
      if (Array.isArray(j)) {
        return j
          .map(function (x) {
            return String(x || "").trim();
          })
          .filter(function (x) {
            return x.length > 0;
          });
      }
    } catch (_) {}
  }
  return s
    .split(/[,;\n]+/)
    .map(function (x) {
      return x.trim();
    })
    .filter(function (x) {
      return x.length > 0;
    });
}
function _annStringifyIdList_(arr) {
  if (!arr || !arr.length) return "";
  var seen = {};
  var clean = [];
  for (var i = 0; i < arr.length; i++) {
    var v = String(arr[i] || "").trim();
    if (!v || seen[v]) continue;
    seen[v] = true;
    clean.push(v);
  }
  return clean.length ? JSON.stringify(clean) : "";
}
function _annEnsureSheet_() {
  var headers = [
    "id",
    "title",
    "body",
    "severity",
    "is_dismissible",
    "is_active",
    "starts_at",
    "ends_at",
    "target_role",
    "cta_text",
    "cta_url",
    "created_by",
    "created_at",
    "updated_at",
    "target_user_ids",
    "dismissible_user_ids",
    "nondismissible_user_ids",
    "send_email_notification",
    "email_sent",
    "reminder_3_sent",
    "reminder_2_sent",
    "reminder_1_sent",
  ];
  var ss = SpreadsheetApp.openById(getDbId());
  var sheet = ss.getSheetByName(SHEET_NAME.ANNOUNCEMENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME.ANNOUNCEMENTS);
    sheet.appendRow(headers);
    return sheet;
  }
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.appendRow(headers);
    return sheet;
  }
  var current = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h || "")
        .toLowerCase()
        .trim()
        .replace(/\s+/g, "_");
    });
  if (headers.length > current.length) {
    var missing = headers.slice(current.length);
    sheet
      .getRange(1, current.length + 1, 1, missing.length)
      .setValues([missing]);
  }
  return sheet;
}
function getActiveAnnouncementsForUser(token) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized", items: [] };
    _annEnsureSheet_();
    const today = Utilities.formatDate(
      new Date(),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const role = String(user.role || "").toLowerCase();
    const myId = String(user.id || "");
    const all = getData(SHEET_NAME.ANNOUNCEMENTS) || [];
    const items = all
      .filter((a) => {
        if (!_annBool_(a.is_active)) return false;
        const sa = _annDateOnly_(a.starts_at);
        const ea = _annDateOnly_(a.ends_at);
        if (sa && today < sa) return false;
        if (ea && today > ea) return false;
        const targetUsers = _annParseIdList_(a.target_user_ids);
        if (targetUsers.length > 0) {
          return targetUsers.indexOf(myId) >= 0;
        }
        const tgt = String(a.target_role || "all").toLowerCase();
        if (tgt !== "all" && tgt !== role) return false;
        return true;
      })
      .map((a) => {
        const ndList = _annParseIdList_(a.nondismissible_user_ids);
        const dList = _annParseIdList_(a.dismissible_user_ids);
        let resolvedDismissible = _annBool_(a.is_dismissible);
        if (ndList.indexOf(myId) >= 0) resolvedDismissible = false;
        else if (dList.indexOf(myId) >= 0) resolvedDismissible = true;
        if (role === "admin") resolvedDismissible = true;
        return {
          id: String(a.id),
          title: String(a.title || ""),
          body: String(a.body || ""),
          severity: String(a.severity || "info").toLowerCase(),
          is_dismissible: resolvedDismissible,
          starts_at: _annDateOnly_(a.starts_at),
          ends_at: _annDateOnly_(a.ends_at),
          target_role: String(a.target_role || "all").toLowerCase(),
          cta_text: String(a.cta_text || ""),
          cta_url: String(a.cta_url || ""),
          updated_at: String(a.updated_at || a.created_at || ""),
        };
      });
    const sevRank = { critical: 0, warning: 1, success: 2, info: 3 };
    items.sort((a, b) => {
      const ra = sevRank[a.severity] != null ? sevRank[a.severity] : 4;
      const rb = sevRank[b.severity] != null ? sevRank[b.severity] : 4;
      if (ra !== rb) return ra - rb;
      return String(b.updated_at).localeCompare(String(a.updated_at));
    });
    return { status: "success", items: items };
  } catch (e) {
    return { status: "error", message: e.toString(), items: [] };
  }
}
function listAnnouncementsAdmin(token) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return { status: "error", message: "Hanya admin yang dapat mengakses." };
  }
  try {
    _annEnsureSheet_();
    const all = (getData(SHEET_NAME.ANNOUNCEMENTS) || []).map((a) => ({
      id: String(a.id),
      title: String(a.title || ""),
      body: String(a.body || ""),
      severity: String(a.severity || "info").toLowerCase(),
      is_dismissible: _annBool_(a.is_dismissible),
      is_active: _annBool_(a.is_active),
      starts_at: _annDateOnly_(a.starts_at),
      ends_at: _annDateOnly_(a.ends_at),
      target_role: String(a.target_role || "all").toLowerCase(),
      cta_text: String(a.cta_text || ""),
      cta_url: String(a.cta_url || ""),
      created_by: String(a.created_by || ""),
      created_at: String(a.created_at || ""),
      updated_at: String(a.updated_at || ""),
      target_user_ids: _annParseIdList_(a.target_user_ids),
      dismissible_user_ids: _annParseIdList_(a.dismissible_user_ids),
      nondismissible_user_ids: _annParseIdList_(a.nondismissible_user_ids),
      send_email_notification: _annBool_(a.send_email_notification),
    }));
    all.sort((a, b) => {
      if (a.is_active !== b.is_active) return a.is_active ? -1 : 1;
      return String(b.updated_at || b.created_at).localeCompare(
        String(a.updated_at || a.created_at),
      );
    });
    return { status: "success", items: all };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}
function saveAnnouncement(token, data) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return {
      status: "error",
      message: "Hanya admin yang dapat menyimpan pengumuman.",
    };
  }
  if (!data || !String(data.title || "").trim()) {
    return { status: "error", message: "Judul pengumuman wajib diisi." };
  }
  if (!String(data.body || "").trim()) {
    return { status: "error", message: "Isi pengumuman wajib diisi." };
  }
  const validSeverity = ["info", "success", "warning", "critical"];
  const validRole = ["all", "guru", "admin"];
  const severity =
    validSeverity.indexOf(String(data.severity || "").toLowerCase()) >= 0
      ? String(data.severity).toLowerCase()
      : "info";
  const targetRole =
    validRole.indexOf(String(data.target_role || "").toLowerCase()) >= 0
      ? String(data.target_role).toLowerCase()
      : "guru";
  const starts = String(data.starts_at || "").trim();
  const ends = String(data.ends_at || "").trim();
  if (starts && !/^\d{4}-\d{2}-\d{2}$/.test(starts)) {
    return {
      status: "error",
      message: "Format tanggal mulai tidak valid (yyyy-MM-dd).",
    };
  }
  if (ends && !/^\d{4}-\d{2}-\d{2}$/.test(ends)) {
    return {
      status: "error",
      message: "Format tanggal selesai tidak valid (yyyy-MM-dd).",
    };
  }
  if (starts && ends && ends < starts) {
    return {
      status: "error",
      message: "Tanggal selesai harus setelah tanggal mulai.",
    };
  }
  const sheet = _annEnsureSheet_();
  const now = Utilities.formatDate(
    new Date(),
    "Asia/Jakarta",
    "yyyy-MM-dd HH:mm:ss",
  );
  const isDismissible = _annBool_(data.is_dismissible);
  const isActive =
    data.is_active === undefined ? true : _annBool_(data.is_active);
  const ctaText = String(data.cta_text || "")
    .trim()
    .slice(0, 120);
  const ctaUrl = String(data.cta_url || "").trim();
  if (ctaUrl && !/^https?:\/\//i.test(ctaUrl)) {
    return {
      status: "error",
      message: "CTA URL harus diawali http:// atau https://",
    };
  }
  const allUsers = getData("Users") || [];
  const validIds = {};
  allUsers.forEach((u) => {
    if (u && u.id != null) validIds[String(u.id)] = true;
  });
  function _validateIds(arr, label) {
    var clean = [];
    var invalid = [];
    (arr || []).forEach((id) => {
      var s = String(id || "").trim();
      if (!s) return;
      if (validIds[s]) clean.push(s);
      else invalid.push(s);
    });
    return { clean: clean, invalid: invalid };
  }
  const targetCheck = _validateIds(data.target_user_ids, "sasaran");
  const dismCheck = _validateIds(data.dismissible_user_ids, "boleh tutup");
  const ndCheck = _validateIds(
    data.nondismissible_user_ids,
    "tidak boleh tutup",
  );
  if (
    targetCheck.invalid.length ||
    dismCheck.invalid.length ||
    ndCheck.invalid.length
  ) {
    return {
      status: "error",
      message:
        "Beberapa ID pengguna tidak ditemukan: " +
        []
          .concat(targetCheck.invalid, dismCheck.invalid, ndCheck.invalid)
          .join(", "),
    };
  }
  var conflict = dismCheck.clean.filter((id) => ndCheck.clean.indexOf(id) >= 0);
  if (conflict.length) {
    return {
      status: "error",
      message:
        'Pengguna tidak boleh berada di kedua daftar "boleh tutup" dan "tidak boleh tutup": ' +
        conflict.join(", "),
    };
  }
  const targetIdsStr = _annStringifyIdList_(targetCheck.clean);
  const dismIdsStr = _annStringifyIdList_(dismCheck.clean);
  const ndIdsStr = _annStringifyIdList_(ndCheck.clean);
  const sendEmailNotification = _annBool_(data.send_email_notification);
  if (data.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        // Cek apakah konten signifikan berubah — jika ya, reset email_sent
        // agar email dikirim ulang ke penerima yang sesuai.
        const titleChanged = String(rows[i][1] || "") !== String(data.title || "").trim();
        const bodyChanged  = String(rows[i][2] || "") !== String(data.body  || "");
        const shouldResetEmail = titleChanged || bodyChanged;
        const prevEmailSent = rows[i][18]; // col 19 (0-indexed: 18) = email_sent
        const newEmailSent  = shouldResetEmail ? false : prevEmailSent;
        // Jika email di-reset, juga reset semua flag reminder
        const prevRem3 = shouldResetEmail ? false : rows[i][19];
        const prevRem2 = shouldResetEmail ? false : rows[i][20];
        const prevRem1 = shouldResetEmail ? false : rows[i][21];

        // Kolom 2–22 (1-indexed), total 21 kolom:
        // title, body, severity, is_dismissible, is_active, starts_at, ends_at,
        // target_role, cta_text, cta_url, created_by, created_at, updated_at,
        // target_user_ids, dismissible_user_ids, nondismissible_user_ids,
        // send_email_notification, email_sent, reminder_3_sent, reminder_2_sent, reminder_1_sent
        sheet
          .getRange(i + 1, 2, 1, 21)
          .setValues([
            [
              String(data.title).trim(),       // col 2: title
              String(data.body),               // col 3: body
              severity,                        // col 4: severity
              isDismissible,                   // col 5: is_dismissible
              isActive,                        // col 6: is_active
              starts,                          // col 7: starts_at
              ends,                            // col 8: ends_at
              targetRole,                      // col 9: target_role
              ctaText,                         // col 10: cta_text
              ctaUrl,                          // col 11: cta_url
              rows[i][11] || user.full_name || user.id, // col 12: created_by (preserve)
              rows[i][12] || now,              // col 13: created_at (preserve)
              now,                             // col 14: updated_at
              targetIdsStr,                    // col 15: target_user_ids
              dismIdsStr,                      // col 16: dismissible_user_ids
              ndIdsStr,                        // col 17: nondismissible_user_ids
              sendEmailNotification,           // col 18: send_email_notification
              newEmailSent,                    // col 19: email_sent (reset jika konten berubah)
              prevRem3,                        // col 20: reminder_3_sent
              prevRem2,                        // col 21: reminder_2_sent
              prevRem1,                        // col 22: reminder_1_sent
            ],
          ]);
        try {
          _invalidateDataSnapshot();
        } catch (_) {}
        if (sendEmailNotification) {
          checkAndSendAnnouncementEmails();
        }
        return {
          status: "success",
          message: "Pengumuman diperbarui.",
          id: String(data.id),
        };
      }
    }
    return { status: "error", message: "Pengumuman tidak ditemukan." };
  }
  const newId = "ANN-" + new Date().getTime();
  sheet.appendRow([
    newId,
    String(data.title).trim(),  // col 2: title
    String(data.body),          // col 3: body
    severity,                   // col 4: severity
    isDismissible,              // col 5: is_dismissible
    isActive,                   // col 6: is_active
    starts,                     // col 7: starts_at
    ends,                       // col 8: ends_at
    targetRole,                 // col 9: target_role
    ctaText,                    // col 10: cta_text
    ctaUrl,                     // col 11: cta_url
    user.full_name || user.id,  // col 12: created_by
    now,                        // col 13: created_at
    now,                        // col 14: updated_at
    targetIdsStr,               // col 15: target_user_ids
    dismIdsStr,                 // col 16: dismissible_user_ids
    ndIdsStr,                   // col 17: nondismissible_user_ids
    sendEmailNotification,      // col 18: send_email_notification
    "",                         // col 19: email_sent
    "",                         // col 20: reminder_3_sent
    "",                         // col 21: reminder_2_sent
    "",                         // col 22: reminder_1_sent
  ]);
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  if (sendEmailNotification) {
    checkAndSendAnnouncementEmails();
  }
  return { status: "success", message: "Pengumuman dibuat.", id: newId };
}
function toggleAnnouncement(token, id, isActive) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return { status: "error", message: "Hanya admin." };
  }
  const sheet = _annEnsureSheet_();
  const rows = sheet.getDataRange().getValues();
  const now = Utilities.formatDate(
    new Date(),
    "Asia/Jakarta",
    "yyyy-MM-dd HH:mm:ss",
  );
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue(_annBool_(isActive));
      sheet.getRange(i + 1, 14).setValue(now);
      try {
        _invalidateDataSnapshot();
      } catch (_) {}
      return { status: "success" };
    }
  }
  return { status: "error", message: "Pengumuman tidak ditemukan." };
}
function deleteAnnouncement(token, id) {
  return deleteRowById(token, SHEET_NAME.ANNOUNCEMENTS, id);
}
function checkAndSendAnnouncementEmails() {
  const sheet = _annEnsureSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];
  const todayDate = new Date();
  todayDate.setHours(0, 0, 0, 0);
  const today = Utilities.formatDate(todayDate, "Asia/Jakarta", "yyyy-MM-dd");
  const idxId = headers.indexOf("id");
  const idxTitle = headers.indexOf("title");
  const idxBody = headers.indexOf("body");
  const idxIsActive = headers.indexOf("is_active");
  const idxStartsAt = headers.indexOf("starts_at");
  const idxEndsAt = headers.indexOf("ends_at");
  const idxTargetRole = headers.indexOf("target_role");
  const idxTargetUserIds = headers.indexOf("target_user_ids");
  const idxSendEmail = headers.indexOf("send_email_notification");
  const idxEmailSent = headers.indexOf("email_sent");
  const idxRem3 = headers.indexOf("reminder_3_sent");
  const idxRem2 = headers.indexOf("reminder_2_sent");
  const idxRem1 = headers.indexOf("reminder_1_sent");
  if (idxSendEmail === -1) return;
  const allUsers = getData("Users") || [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const isActive = _annBool_(row[idxIsActive]);
    const sendEmail = _annBool_(row[idxSendEmail]);
    if (!isActive || !sendEmail) continue;
    const emailSent =
      idxEmailSent !== -1 ? _annBool_(row[idxEmailSent]) : false;
    const startsAt = _annDateOnly_(row[idxStartsAt]);
    const endsAt = idxEndsAt !== -1 ? _annDateOnly_(row[idxEndsAt]) : "";
    const rem3Sent = idxRem3 !== -1 ? _annBool_(row[idxRem3]) : false;
    const rem2Sent = idxRem2 !== -1 ? _annBool_(row[idxRem2]) : false;
    const rem1Sent = idxRem1 !== -1 ? _annBool_(row[idxRem1]) : false;
    let emailToSend = null; 
    let colToUpdate = -1;
    if (!emailSent && (!startsAt || startsAt <= today)) {
      emailToSend = "start";
      colToUpdate = idxEmailSent;
    } else if (endsAt && (!startsAt || startsAt <= today)) {
      const endDate = new Date(endsAt);
      endDate.setHours(0, 0, 0, 0);
      const diffTime = endDate.getTime() - todayDate.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays === 3 && !rem3Sent) {
        emailToSend = "rem3";
        colToUpdate = idxRem3;
      } else if (diffDays === 2 && !rem2Sent) {
        emailToSend = "rem2";
        colToUpdate = idxRem2;
      } else if (diffDays === 1 && !rem1Sent) {
        emailToSend = "rem1";
        colToUpdate = idxRem1;
      }
    }
    if (emailToSend) {
      const title = String(row[idxTitle] || "");
      const body = String(row[idxBody] || "");
      const targetRole = String(row[idxTargetRole] || "all").toLowerCase();
      const targetUserIdsStr = String(row[idxTargetUserIds] || "");
      let targetUserIds = [];
      try {
        if (targetUserIdsStr) targetUserIds = JSON.parse(targetUserIdsStr);
      } catch (e) {}
      const targetEmails = [];
      allUsers.forEach((u) => {
        if (!u.email || !_isValidEmail_(u.email)) return;
        let include = false;
        if (targetUserIds.length > 0) {
          if (targetUserIds.indexOf(String(u.id)) >= 0) include = true;
        } else {
          if (targetRole === "all") {
            include = true;
          } else {
            const uRole = String(u.role || "").toLowerCase();
            if (targetRole === uRole) include = true;
          }
        }
        if (include && targetEmails.indexOf(u.email) === -1) {
          targetEmails.push(u.email);
        }
      });
      if (targetEmails.length > 0) {
        // Helper escape HTML khusus untuk email server-side (escapeHtml hanya ada di frontend)
        function _annEscapeHtml_(str) {
          return String(str || "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
        }
        let subject = "";
        let prefix = "";
        if (emailToSend === "start") {
          subject = "Pengumuman Baru: " + title;
          const dateStr = startsAt
            ? formatDateIndo(startsAt)
            : formatDateIndo(today);
          prefix = `Berlaku mulai: ${_annEscapeHtml_(dateStr)}`;
        } else {
          const daysLeft =
            emailToSend === "rem3" ? 3 : emailToSend === "rem2" ? 2 : 1;
          subject = `Peringatan (${daysLeft} hari lagi): ${title}`;
          const dateStr = formatDateIndo(endsAt);
          prefix = `Pengingat! Pengumuman ini akan berakhir pada: ${_annEscapeHtml_(dateStr)} (${daysLeft} hari lagi)`;
        }
        // Body disanitasi: tag HTML diizinkan (admin yang menulis), tapi
        // escape karakter berbahaya di luar konteks tag agar aman dikirim via email.
        // Konversi newline ke <br/> untuk plaintext fallback.
        const safeBody = String(body).replace(/\r\n/g, "\n").replace(/\n/g, "<br/>");
        const htmlBody = `
          <div style="font-family:sans-serif; max-width:600px; margin:0 auto; padding:20px; color:#334155;">
            <h2 style="color:#1e293b; margin-top:0;">${_annEscapeHtml_(title)}</h2>
            <p style="font-size:13px; color:#e11d48; margin-bottom:20px; font-weight:bold;">${prefix}</p>
            <div style="background:#f8fafc; padding:15px; border-radius:8px; line-height:1.6; color:#334155;">
              ${safeBody}
            </div>
            <p style="font-size:12px; color:#94a3b8; margin-top:30px; border-top:1px solid #e2e8f0; padding-top:15px;">
              Email otomatis dari SiM-Guru. Harap periksa aplikasi untuk detail lebih lanjut. Jangan membalas email ini.
            </p>
          </div>
        `;
        const chunk = 50;
        for (let j = 0; j < targetEmails.length; j += chunk) {
          const bccChunk = targetEmails.slice(j, j + chunk).join(",");
          try {
            MailApp.sendEmail({
              to: Session.getActiveUser().getEmail() || "noreply@simguru.local",
              bcc: bccChunk,
              subject: subject,
              htmlBody: htmlBody,
              body: "Silakan aktifkan tampilan HTML untuk melihat pengumuman ini.",
            });
          } catch (e) {
            console.error("Gagal mengirim email pengumuman: " + e.message);
          }
        }
      }
      if (colToUpdate !== -1) {
        sheet.getRange(i + 1, colToUpdate + 1).setValue(true);
      }
    }
  }
}
function installAnnouncementTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkAndSendAnnouncementEmails") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("checkAndSendAnnouncementEmails")
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone("Asia/Jakarta")
    .create();
  return {
    status: "success",
    message: "Trigger berhasil diinstall pada jam 08:00 pagi.",
  };
}
