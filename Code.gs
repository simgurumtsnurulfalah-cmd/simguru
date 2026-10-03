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
function validateDocument(trxId) {
  if (!trxId || String(trxId).trim() === "") {
    return { status: "invalid", message: "ID Transaksi tidak boleh kosong." };
  }
  const cleanId = String(trxId).trim().toUpperCase();
  const allHistory = getData("Honor_History");
  const history = allHistory.find(
    (h) =>
      String(h.trx_id || "")
        .trim()
        .toUpperCase() === cleanId,
  );
  if (history) {
    const user = findData("Users", "id", history.user_id);
    const config = getData("Config");
    const schoolNameCfg = config.find((c) => String(c.key) === "school_name");
    const kepalaSekolahCfg = config.find(
      (c) => String(c.key) === "kepala_sekolah",
    );
    const totalJtm = Number(history.total_jtm) || 0;
    const totalHonor = Number(history.total_honor) || 0;
    const baseSalary =
      Number(
        (config.find((c) => String(c.key) === "base_salary") || {}).value,
      ) || 0;
    let periodeStr = String(history.periode || "").trim();
    if (history.periode instanceof Date) {
      periodeStr = Utilities.formatDate(
        history.periode,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
    }
    let tglSimpan = history.tanggal_simpan;
    if (tglSimpan instanceof Date) {
      tglSimpan = formatDateIndo(tglSimpan);
    } else if (tglSimpan) {
      tglSimpan = formatDateIndo(new Date(tglSimpan));
    } else {
      tglSimpan = "-";
    }
    return {
      status: "valid",
      data: {
        trx_id: String(history.trx_id),
        nama_guru: user ? String(user.full_name) : "Unknown",
        nip: user ? String(user.nip || "-") : "-",
        periode: formatPeriode(periodeStr),
        total_jtm: totalJtm,
        total_honor: formatRupiah(totalHonor),
        tgl_simpan: tglSimpan,
        tgl_verifikasi: formatDateIndo(new Date()),
        school_name: schoolNameCfg
          ? String(schoolNameCfg.value)
          : "MTs Nurul Falah",
        kepala_sekolah: kepalaSekolahCfg ? String(kepalaSekolahCfg.value) : "-",
      },
    };
  }
  return {
    status: "invalid",
    message:
      "Dokumen tidak ditemukan. Pastikan ID Transaksi yang Anda masukkan benar dan sesuai dengan slip yang diterima.",
  };
}
function getGuruJournalHistoryPaginated(
  token,
  page,
  limit,
  filterDateFrom,
  filterDateTo,
  searchQuery,
) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    page = parseInt(page) || 1;
    limit = parseInt(limit) || 10;
    // BUG FIX: tz harus dideklarasikan SEBELUM getTimestamp agar tidak undefined saat closure dibuat
    const tz = Session.getScriptTimeZone();
    const allUserLogs = getData("Teaching_Logs").filter(
      (l) => String(l.user_id).trim() === String(user.id).trim(),
    );
    const schedules = getData("Schedules");
    const getTimestamp = (entry) => {
      let dStr = entry.date;
      if (dStr instanceof Date) {
        dStr = Utilities.formatDate(dStr, tz, "yyyy-MM-dd");
      } else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
        let dObj = new Date(dStr);
        if (!isNaN(dObj.getTime()))
          dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
      }
      let tStr = entry.waktu_submit;
      if (tStr instanceof Date) {
        tStr = Utilities.formatDate(tStr, tz, "HH:mm:ss");
      } else if (typeof tStr === "string") {
        if (tStr.indexOf("T") !== -1) {
          let tObj = new Date(tStr);
          if (!isNaN(tObj.getTime()))
            tStr = Utilities.formatDate(tObj, tz, "HH:mm:ss");
        } else if (tStr.length === 5) {
          tStr = tStr + ":00";
        }
      }
      let finalDateStr = dStr + "T" + (tStr || "00:00:00");
      let ts = new Date(finalDateStr).getTime();
      return isNaN(ts) ? 0 : ts;
    };
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    let logsFiltered = allUserLogs;
    if (filterDateFrom) {
      logsFiltered = logsFiltered.filter((l) => {
        let dStr = l.date;
        if (l.date instanceof Date) {
          dStr = Utilities.formatDate(l.date, tz, "yyyy-MM-dd");
        } else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
          let dObj = new Date(dStr);
          if (!isNaN(dObj.getTime()))
            dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
        }
        return String(dStr) >= String(filterDateFrom);
      });
    }
    if (filterDateTo) {
      logsFiltered = logsFiltered.filter((l) => {
        let dStr = l.date;
        if (l.date instanceof Date) {
          dStr = Utilities.formatDate(l.date, tz, "yyyy-MM-dd");
        } else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
          let dObj = new Date(dStr);
          if (!isNaN(dObj.getTime()))
            dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
        }
        return String(dStr) <= String(filterDateTo);
      });
    }
    if (searchQuery) {
      const q = String(searchQuery).toLowerCase();
      logsFiltered = logsFiltered.filter((l) => {
        let subj = "",
          cls = "";
        const sid = String(l.schedule_id).toUpperCase();
        if (sid === "PICKET-DUTY") {
          subj = "tugas piket";
        } else if (sid === "CEREMONY-DUTY") {
          subj = "pembina upacara";
        } else if (sid === "EXAM-SUPERVISOR") {
          subj = "pengawas ujian";
        } else if (sid === "EXAM-COMMITTEE") {
          subj = "panitia ujian";
        } else if (sid.indexOf("PARTIAL-SUB-") === 0) {
          subj = "substitusi parsial";
        } else {
          const s = schedules.find(
            (sc) => String(sc.id) === String(l.schedule_id),
          );
          if (s) {
            subj = String(s.subject || "").toLowerCase();
            cls = String(s.class_name || "").toLowerCase();
          }
        }
        const mat = String(l.materi || "").toLowerCase();
        return (
          subj.indexOf(q) !== -1 ||
          cls.indexOf(q) !== -1 ||
          mat.indexOf(q) !== -1
        );
      });
    }
    // BUG FIX: Stats dihitung dari logsFiltered (setelah filter aktif), bukan dari semua data
    let totalJtmAll = 0;
    let totalThisMonth = 0;
    let sumHadirAll = 0;
    let sumAbsenAll = 0;
    logsFiltered.forEach((l) => {
      const jtm = Number(l.jtm_val || 0);
      totalJtmAll += jtm;
      sumHadirAll += Number(l.siswa_hadir || 0);
      sumAbsenAll += Number(l.siswa_absen || 0);
      let dStr = l.date;
      if (l.date instanceof Date) {
        dStr = Utilities.formatDate(l.date, tz, "yyyy-MM-dd");
      } else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
        let dObj = new Date(dStr);
        if (!isNaN(dObj.getTime()))
          dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
      }
      const d = new Date(dStr);
      if (d.getMonth() === currentMonth && d.getFullYear() === currentYear) {
        totalThisMonth++;
      }
    });
    const totalAllSiswa = sumHadirAll + sumAbsenAll;
    const attRateAll =
      totalAllSiswa > 0
        ? Math.round((sumHadirAll * 100) / totalAllSiswa)
        : null;
    const sortedLogs = logsFiltered.sort(
      (a, b) => getTimestamp(b) - getTimestamp(a),
    );
    const totalItems = sortedLogs.length;
    const totalPages = Math.ceil(totalItems / limit) || 1;
    const startIndex = (page - 1) * limit;
    const pagedLogs = sortedLogs.slice(startIndex, startIndex + limit);
    const history = pagedLogs.map((log) => {
      let sched = null;
      const sid = String(log.schedule_id).toUpperCase();
      if (sid !== "PICKET-DUTY" && sid !== "CEREMONY-DUTY") {
        sched = schedules.find((s) => String(s.id) === String(log.schedule_id));
      }
      let subj = "N/A",
        cls = "N/A";
      if (sid === "PICKET-DUTY") {
        subj = "Tugas Piket";
        cls = "-";
      } else if (sid === "CEREMONY-DUTY") {
        subj = "Pembina Upacara";
        cls = "-";
      } else if (sid === "EXAM-SUPERVISOR") {
        subj = "Pengawas Ujian";
        cls = "-";
      } else if (sid === "EXAM-COMMITTEE") {
        subj = "Panitia Ujian";
        cls = "-";
      } else if (sid.indexOf("PARTIAL-SUB-") === 0) {
        subj = "Substitusi Parsial";
        cls = "-";
      } else if (sched) {
        subj = sched.subject;
        cls = sched.class_name;
      }
      let safeDate = log.date;
      if (safeDate instanceof Date) {
        safeDate = Utilities.formatDate(safeDate, tz, "yyyy-MM-dd");
      } else if (typeof safeDate === "string" && safeDate.indexOf("T") !== -1) {
        let dObj = new Date(safeDate);
        if (!isNaN(dObj.getTime()))
          safeDate = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
      }
      let safeTime = log.waktu_submit;
      if (safeTime instanceof Date) {
        safeTime = Utilities.formatDate(safeTime, tz, "HH:mm");
      } else if (typeof safeTime === "string") {
        if (safeTime.indexOf("T") !== -1) {
          let tObj = new Date(safeTime);
          if (!isNaN(tObj.getTime()))
            safeTime = Utilities.formatDate(tObj, tz, "HH:mm");
        } else if (safeTime.length >= 5) {
          safeTime = safeTime.substring(0, 5);
        }
      }
      if (!safeTime) safeTime = "-";
      return {
        log_id: log.log_id,
        date: String(safeDate),
        time: String(safeTime),
        class_name: cls,
        subject: subj,
        materi: log.materi,
        students: log.siswa_hadir,
        students_absent: log.siswa_absen || 0,
        notes: log.notes || "",
        clean_notes: log.notes || "-",
        jtm_val: Number(log.jtm_val || 0),
      };
    });
    return {
      status: "success",
      data: history,
      stats: {
        total_this_month: totalThisMonth,
        total_jtm_all: totalJtmAll,
        att_rate_all: attRateAll,
      },
      pagination: {
        current_page: page,
        total_pages: totalPages,
        total_items: totalItems,
        items_per_page: limit,
      },
    };
  } catch (error) {
    return { status: "error", message: "Server Error: " + error.toString() };
  }
}
function generateAutoJurnalForHoliday(token, targetDate) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    return _runAutoJurnalForHoliday_(user.id, user.full_name, targetDate);
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}
function _runAutoJurnalForHoliday_(actorUserId, actorName, targetDate) {
  try {
    const checkDate = targetDate ? new Date(targetDate) : new Date();
    const dateStr = Utilities.formatDate(
      checkDate,
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const dayIndex = checkDate.getDay();
    const lockId = `AUTO_JURNAL_${dateStr}`;
    const lockSheet = getSheet("System_Locks");
    const lockData = lockSheet.getDataRange().getValues();
    const existingLock = lockData.find(
      (row) => row[0] === lockId && row[1] === "COMPLETED",
    );
    if (existingLock) {
      return {
        status: "info",
        message: "Auto-jurnal untuk hari ini sudah pernah berhasil dibuat",
        date: dateStr,
        locked: true,
        locked_at: existingLock[2],
      };
    }
    const holidays = getData("Academic_Calendar");
    const holiday = holidays.find((h) => {
      const hDate = new Date(h.date);
      const hDateStr = Utilities.formatDate(
        hDate,
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        hDateStr === dateStr && String(h.is_holiday).toLowerCase() === "true"
      );
    });
    if (!holiday) {
      return {
        status: "error",
        message: "Hari ini bukan hari libur",
        date: dateStr,
      };
    }
    const holidayDesc = holiday.description || "Hari Libur";
    const lockTimestamp = Utilities.formatDate(
      new Date(),
      "Asia/Jakarta",
      "yyyy-MM-dd HH:mm:ss",
    );
    lockSheet.appendRow([
      lockId,
      "RUNNING",
      lockTimestamp,
      actorUserId || "SYSTEM",
      "",
    ]);
    const sheet = getSheet("Teaching_Logs");
    const sheetCols = Math.max(sheet.getLastColumn(), 10);
    // Ambil konfigurasi aktif (tahun pelajaran & semester)
    const configRaw = getData("Config");
    const cfg = {};
    (configRaw || []).forEach((c) => { cfg[c.key] = c.value; });
    const activeTP = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    // Filter jadwal: hanya hari yang sama DAN semester/tahun pelajaran yang aktif
    const allSchedules = getData("Schedules").filter((s) => {
      if (String(s.day_index).trim() !== String(dayIndex)) return false;
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return sTP === activeTP && sSem === activeSem;
    });
    const allLogsToday = getData("Teaching_Logs").filter((l) => {
      try {
        return (
          Utilities.formatDate(
            new Date(l.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          ) === dateStr
        );
      } catch (_) {
        return false;
      }
    });
    const autoLogKeys = new Set();
    const manualLogKeys = new Set();
    allLogsToday.forEach((l) => {
      const k = String(l.user_id) + "|" + String(l.schedule_id);
      const notes = String(l.notes || "");
      if (notes.includes("Auto-generated: Libur Bonus")) {
        autoLogKeys.add(k);
      } else if (!notes.includes("Auto-generated")) {
        manualLogKeys.add(k);
      }
    });
    const generatedLogs = [];
    const skippedEntries = [];
    const newRowsToWrite = [];
    function _bufferLog(row, summary) {
      newRowsToWrite.push(row);
      autoLogKeys.add(String(row[2]) + "|" + String(row[1]));
      generatedLogs.push(summary);
    }
    allSchedules.forEach((schedule) => {
      try {
        if (
          schedule.status &&
          String(schedule.status).toLowerCase() !== "active"
        ) {
          skippedEntries.push({
            user_id: schedule.user_id,
            schedule_id: schedule.id,
            reason: "SCHEDULE_NOT_ACTIVE",
          });
          return;
        }
        const key = String(schedule.user_id) + "|" + String(schedule.id);
        if (autoLogKeys.has(key)) {
          skippedEntries.push({
            user_id: schedule.user_id,
            schedule_id: schedule.id,
            reason: "EXACT_DUPLICATE",
          });
          return;
        }
        if (manualLogKeys.has(key)) {
          skippedEntries.push({
            user_id: schedule.user_id,
            schedule_id: schedule.id,
            reason: "HAS_MANUAL_LOG",
          });
          return;
        }
        const newLogId = generateId("LOG-AUTO");
        const timeStr = Utilities.formatDate(
          new Date(),
          "Asia/Jakarta",
          "HH:mm:ss",
        );
        const materiOtomatis =
          `[LIBUR: ${holidayDesc}] - Bonus JTM Otomatis | ` +
          `Jadwal: ${schedule.subject} Kelas ${schedule.class_name} | ` +
          `JTM: ${schedule.jtm_val}`;
        _bufferLog(
          [
            newLogId,
            schedule.id,
            schedule.user_id,
            dateStr,
            materiOtomatis,
            0,
            0,
            "Auto-generated: Libur Bonus | Idempotent: " + lockId,
            Number(schedule.jtm_val) || 0,
            timeStr,
          ],
          {
            log_id: newLogId,
            user_id: schedule.user_id,
            schedule_id: schedule.id,
            subject: schedule.subject,
            class_name: schedule.class_name,
            jtm_val: Number(schedule.jtm_val) || 0,
          },
        );
      } catch (errSchedule) {
        skippedEntries.push({
          user_id: schedule && schedule.user_id,
          schedule_id: schedule && schedule.id,
          reason: "ERROR: " + errSchedule.toString(),
        });
      }
    });
    const todayPickets = getData("Picket_Schedules").filter(
      (p) => Number(p.day_index) === dayIndex,
    );
    todayPickets.forEach((picket) => {
      try {
        const picketUserId = String(picket.user_id || "").trim();
        if (!picketUserId) return;
        const key = picketUserId + "|PICKET-DUTY";
        if (autoLogKeys.has(key) || manualLogKeys.has(key)) {
          skippedEntries.push({
            user_id: picketUserId,
            schedule_id: "PICKET-DUTY",
            reason: "ALREADY_CONFIRMED",
          });
          return;
        }
        const newLogId = generateId("LOG-PCK-AUTO");
        const timeStr = Utilities.formatDate(
          new Date(),
          "Asia/Jakarta",
          "HH:mm:ss",
        );
        _bufferLog(
          [
            newLogId,
            "PICKET-DUTY",
            picketUserId,
            dateStr,
            "Melaksanakan Tugas Piket — Auto-confirmed (" + holidayDesc + ")",
            0,
            0,
            "Auto-generated: Libur Bonus | Idempotent: " + lockId,
            4,
            timeStr,
          ],
          {
            log_id: newLogId,
            user_id: picketUserId,
            schedule_id: "PICKET-DUTY",
            subject: "Tugas Piket",
            class_name: "-",
            jtm_val: 4,
          },
        );
      } catch (errPicket) {
        skippedEntries.push({
          user_id: picket && picket.user_id,
          schedule_id: "PICKET-DUTY",
          reason: "ERROR: " + errPicket.toString(),
        });
      }
    });
    if (dayIndex === 1) {
      const todayCeremonies = getData("Ceremony_Schedules").filter((c) => {
        try {
          return (
            Utilities.formatDate(
              new Date(c.date),
              "Asia/Jakarta",
              "yyyy-MM-dd",
            ) === dateStr
          );
        } catch (_) {
          return false;
        }
      });
      todayCeremonies.forEach((cer) => {
        try {
          const cerUserId = String(cer.user_id || "").trim();
          if (!cerUserId) return;
          const key = cerUserId + "|CEREMONY-DUTY";
          if (autoLogKeys.has(key) || manualLogKeys.has(key)) {
            skippedEntries.push({
              user_id: cerUserId,
              schedule_id: "CEREMONY-DUTY",
              reason: "ALREADY_CONFIRMED",
            });
            return;
          }
          const newLogId = generateId("LOG-CER-AUTO");
          const timeStr = Utilities.formatDate(
            new Date(),
            "Asia/Jakarta",
            "HH:mm:ss",
          );
          _bufferLog(
            [
              newLogId,
              "CEREMONY-DUTY",
              cerUserId,
              dateStr,
              "Melaksanakan Tugas Pembina Upacara — Auto-confirmed (" +
                holidayDesc +
                ")",
              0,
              0,
              "Auto-generated: Libur Bonus | Idempotent: " + lockId,
              5,
              timeStr,
            ],
            {
              log_id: newLogId,
              user_id: cerUserId,
              schedule_id: "CEREMONY-DUTY",
              subject: "Pembina Upacara",
              class_name: "-",
              jtm_val: 5,
            },
          );
        } catch (errCer) {
          skippedEntries.push({
            user_id: cer && cer.user_id,
            schedule_id: "CEREMONY-DUTY",
            reason: "ERROR: " + errCer.toString(),
          });
        }
      });
    }
    if (newRowsToWrite.length > 0) {
      const normalized = newRowsToWrite.map((r) => {
        const out = r.slice(0, sheetCols);
        while (out.length < sheetCols) out.push("");
        return out;
      });
      const startRow = sheet.getLastRow() + 1;
      sheet
        .getRange(startRow, 1, normalized.length, sheetCols)
        .setValues(normalized);
      SpreadsheetApp.flush();
    }
    const lockMetadata = {
      generated_by: actorName || "SYSTEM",
      generated_count: generatedLogs.length,
      total_jtm: generatedLogs.reduce((sum, l) => sum + l.jtm_val, 0),
      holiday_desc: holidayDesc,
    };
    const lockRowIndex = lockData.findIndex((row) => row[0] === lockId);
    if (lockRowIndex > 0) {
      lockSheet.getRange(lockRowIndex + 1, 2).setValue("COMPLETED");
      lockSheet.getRange(lockRowIndex + 1, 4).setValue(actorUserId || "SYSTEM");
      lockSheet
        .getRange(lockRowIndex + 1, 5)
        .setValue(JSON.stringify(lockMetadata));
    } else {
      const newLockData = lockSheet.getDataRange().getValues();
      const runningRowIndex = newLockData.findIndex(
        (row) => row[0] === lockId && row[1] === "RUNNING",
      );
      if (runningRowIndex > 0) {
        lockSheet.getRange(runningRowIndex + 1, 2).setValue("COMPLETED");
        lockSheet
          .getRange(runningRowIndex + 1, 5)
          .setValue(JSON.stringify(lockMetadata));
      }
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    var notifSummary = { sent: 0, failed: 0, recipients: 0 };
    try {
      notifSummary =
        _notifAutoJurnalLiburGenerated_(
          generatedLogs,
          dateStr,
          holidayDesc,
          actorName || "Sistem (Otomatis)",
        ) || notifSummary;
    } catch (errNotif) {
      try {
        console.warn("[autoJurnal] gagal kirim notifikasi: " + errNotif);
      } catch (_) {}
    }
    return {
      status: "success",
      message: `Berhasil membuat ${generatedLogs.length} jurnal otomatis`,
      date: dateStr,
      lock_id: lockId,
      generated: generatedLogs.length,
      skipped: skippedEntries.length,
      details: generatedLogs,
      holiday_description: holidayDesc,
      notification: notifSummary,
    };
  } catch (error) {
    try {
      const lockSheet = getSheet("System_Locks");
      const lockId = `AUTO_JURNAL_${Utilities.formatDate(targetDate ? new Date(targetDate) : new Date(), "Asia/Jakarta", "yyyy-MM-dd")}`;
      const lockData = lockSheet.getDataRange().getValues();
      const lockRowIndex = lockData.findIndex((row) => row[0] === lockId);
      if (lockRowIndex > 0) {
        lockSheet.getRange(lockRowIndex + 1, 2).setValue("FAILED");
        lockSheet.getRange(lockRowIndex + 1, 5).setValue(error.toString());
      }
    } catch (e) {}
    return { status: "error", message: error.toString() };
  }
}
function checkAutoJurnalStatus(token, targetDate) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error" };
    const checkDate = targetDate ? new Date(targetDate) : new Date();
    const dateStr = Utilities.formatDate(
      checkDate,
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const dayIndex = checkDate.getDay();
    const lockId = `AUTO_JURNAL_${dateStr}`;
    const holidays = getData("Academic_Calendar");
    const holiday = holidays.find((h) => {
      const hDate = new Date(h.date);
      const hDateStr = Utilities.formatDate(
        hDate,
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        hDateStr === dateStr && String(h.is_holiday).toLowerCase() === "true"
      );
    });
    if (!holiday) {
      return { status: "info", is_holiday: false, date: dateStr };
    }
    const lockSheet = getSheet("System_Locks");
    const lockData = lockSheet.getDataRange().getValues();
    const lockRecord = lockData.find((row) => String(row[0]) === lockId);
    let hasAutoJurnal = false;
    let generatedBy = null;
    let generatedAt = null;
    let existingCount = 0;
    const autoLogs = getData("Teaching_Logs").filter((l) => {
      const logDate = Utilities.formatDate(
        new Date(l.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        logDate === dateStr &&
        String(l.notes).includes("Auto-generated: Libur Bonus")
      );
    });
    if (autoLogs.length > 0) {
      hasAutoJurnal = true;
      existingCount = autoLogs.length;
      if (lockRecord) {
        const metadata = lockRecord[4] ? JSON.parse(String(lockRecord[4])) : {};
        generatedBy = metadata.generated_by || "Unknown";
        generatedAt = lockRecord[2]
          ? Utilities.formatDate(
              new Date(lockRecord[2]),
              "Asia/Jakarta",
              "yyyy-MM-dd HH:mm",
            )
          : "Unknown";
      }
    }
    if (String(user.role).toLowerCase() !== "admin") {
      const configRaw2 = getData("Config");
      const cfg2 = {};
      (configRaw2 || []).forEach((c) => { cfg2[c.key] = c.value; });
      const activeTP2 = cfg2["tahun_pelajaran"] || "";
      const activeSem2 = cfg2["semester"] || "";
      const mySchedules = getData("Schedules").filter((s) => {
        if (String(s.user_id) !== String(user.id)) return false;
        if (String(s.day_index) !== String(dayIndex)) return false;
        const sTP = s.tahun_pelajaran || activeTP2;
        const sSem = s.semester || activeSem2;
        return sTP === activeTP2 && sSem === activeSem2;
      });
      const myAutoLogs = autoLogs.filter(
        (l) => String(l.user_id) === String(user.id),
      );
      return {
        status: "success",
        is_holiday: true,
        date: dateStr,
        holiday_description: holiday.description,
        has_auto_jurnal: hasAutoJurnal,
        my_schedule_count: mySchedules.length,
        my_bonus_jtm: mySchedules.reduce(
          (sum, s) => sum + Number(s.jtm_val),
          0,
        ),
        my_auto_logs: myAutoLogs.map((l) => ({
          log_id: l.log_id,
          materi: l.materi,
          jtm_val: l.jtm_val,
        })),
      };
    }
    return {
      status: "success",
      is_holiday: true,
      date: dateStr,
      holiday_description: holiday.description,
      has_auto_jurnal: hasAutoJurnal,
      can_generate: !hasAutoJurnal,
      can_cancel: hasAutoJurnal,
      existing_count: existingCount,
      lock_id: lockId,
      generated_by: generatedBy,
      generated_at: generatedAt,
    };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}
function cancelAutoJurnalForHoliday(token, targetDate) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Unauthorized: Hanya admin yang dapat membatalkan",
      };
    }
    const checkDate = targetDate ? new Date(targetDate) : new Date();
    const dateStr = Utilities.formatDate(
      checkDate,
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const lockId = `AUTO_JURNAL_${dateStr}`;
    const allLogs = getData("Teaching_Logs");
    const autoLogs = allLogs.filter((l) => {
      const logDate = Utilities.formatDate(
        new Date(l.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        logDate === dateStr &&
        String(l.notes).includes("Auto-generated: Libur Bonus")
      );
    });
    if (autoLogs.length === 0) {
      return {
        status: "error",
        message:
          "Tidak ada jurnal otomatis yang dapat dibatalkan untuk hari ini",
        date: dateStr,
      };
    }
    const logsSheet = getSheet("Teaching_Logs");
    const logsData = logsSheet.getDataRange().getValues();
    const deletedLogs = [];
    let totalJtmRemoved = 0;
    const affectedTeacherIds = new Set();
    const users = getData("Users");
    const schedules = getData("Schedules");
    for (let i = logsData.length - 1; i >= 1; i--) {
      const row = logsData[i];
      const rowDate = Utilities.formatDate(
        new Date(row[3]),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      const notes = String(row[7] || "");
      if (
        rowDate === dateStr &&
        notes.includes("Auto-generated: Libur Bonus")
      ) {
        const userId = String(row[2]);
        const scheduleId = String(row[1]);
        const jtmVal = Number(row[8] || 0);
        const teacher = users.find((u) => String(u.id) === userId);
        const schedule = schedules.find((s) => String(s.id) === scheduleId);
        deletedLogs.push({
          log_id: row[0],
          user_id: userId,
          teacher_name: teacher ? teacher.full_name : "Unknown",
          schedule_id: scheduleId,
          subject: schedule ? schedule.subject : "Unknown",
          class_name: schedule ? schedule.class_name : "-",
          jtm_val: jtmVal,
          materi: String(row[4] || ""),
        });
        totalJtmRemoved += jtmVal;
        affectedTeacherIds.add(userId);
        logsSheet.deleteRow(i + 1);
      }
    }
    const lockSheet = getSheet("System_Locks");
    const lockData = lockSheet.getDataRange().getValues();
    for (let i = 1; i < lockData.length; i++) {
      if (String(lockData[i][0]) === lockId) {
        lockSheet.getRange(i + 1, 2).setValue("CANCELLED");
        lockSheet
          .getRange(i + 1, 5)
          .setValue(
            `Cancelled by ${user.full_name} at ${Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss")}`,
          );
        break;
      }
    }
    const auditSheet = getSheet("System_Locks");
    const cancelId = `CANCEL_${dateStr}_${new Date().getTime()}`;
    auditSheet.appendRow([
      cancelId,
      "CANCELLED",
      Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss"),
      user.id,
      JSON.stringify({
        deleted_count: deletedLogs.length,
        total_jtm_removed: totalJtmRemoved,
        affected_teachers: affectedTeacherIds.size,
        lock_id: lockId,
      }),
    ]);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: `Berhasil membatalkan ${deletedLogs.length} jurnal otomatis`,
      date: dateStr,
      deleted_count: deletedLogs.length,
      total_jtm_removed: totalJtmRemoved,
      affected_teachers: affectedTeacherIds.size,
      details: deletedLogs,
      cancelled_by: user.full_name,
      cancelled_at: Utilities.formatDate(
        new Date(),
        "Asia/Jakarta",
        "yyyy-MM-dd HH:mm:ss",
      ),
    };
  } catch (error) {
    return {
      status: "error",
      message: "Server error saat pembatalan: " + error.toString(),
    };
  }
}
function runAutoJurnalScheduler() {
  try {
    const now = new Date();
    const dateStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    const holidays = getData("Academic_Calendar");
    const holiday = holidays.find((h) => {
      const hDate = new Date(h.date);
      const hDateStr = Utilities.formatDate(
        hDate,
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        hDateStr === dateStr && String(h.is_holiday).toLowerCase() === "true"
      );
    });
    if (!holiday) {
      console.log(
        "[runAutoJurnalScheduler] " + dateStr + " bukan hari libur. Skip.",
      );
      return { status: "skip", reason: "NOT_HOLIDAY", date: dateStr };
    }
    console.log(
      "[runAutoJurnalScheduler] " +
        dateStr +
        ' adalah hari libur ("' +
        holiday.description +
        '"). Menjalankan auto-jurnal...',
    );
    const result = _runAutoJurnalForHoliday_(
      "SYSTEM",
      "Sistem (Otomatis)",
      dateStr,
    );
    console.log(
      "[runAutoJurnalScheduler] Hasil: " +
        JSON.stringify({
          status: result.status,
          generated: result.generated,
          skipped: result.skipped,
          message: result.message,
          notification: result.notification || null,
        }),
    );
    return result;
  } catch (e) {
    console.error("[runAutoJurnalScheduler] Error: " + e.toString());
    return { status: "error", message: e.toString() };
  }
}
function installAutoJurnalScheduler() {
  try {
    uninstallAutoJurnalScheduler();
    const trigger = ScriptApp.newTrigger("runAutoJurnalScheduler")
      .timeBased()
      .atHour(16)
      .nearMinute(0)
      .everyDays(1)
      .inTimezone("Asia/Jakarta")
      .create();
    console.log(
      "[installAutoJurnalScheduler] Trigger terpasang. ID: " +
        trigger.getUniqueId(),
    );
    return {
      status: "success",
      message: "Trigger otomatis terpasang. Akan jalan setiap hari ~16.00 WIB.",
      trigger_id: trigger.getUniqueId(),
    };
  } catch (e) {
    console.error("[installAutoJurnalScheduler] Error: " + e.toString());
    return { status: "error", message: e.toString() };
  }
}
function uninstallAutoJurnalScheduler() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = 0;
    triggers.forEach((t) => {
      if (t.getHandlerFunction() === "runAutoJurnalScheduler") {
        ScriptApp.deleteTrigger(t);
        removed++;
      }
    });
    console.log(
      "[uninstallAutoJurnalScheduler] " + removed + " trigger dihapus.",
    );
    return { status: "success", removed: removed };
  } catch (e) {
    console.error("[uninstallAutoJurnalScheduler] Error: " + e.toString());
    return { status: "error", message: e.toString() };
  }
}
function getAutoJurnalSchedulerStatus(token) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Hanya admin yang dapat melihat status trigger.",
      };
    }
    const triggers = ScriptApp.getProjectTriggers().filter(
      (t) => t.getHandlerFunction() === "runAutoJurnalScheduler",
    );
    if (triggers.length === 0) {
      return { status: "success", installed: false, count: 0 };
    }
    return {
      status: "success",
      installed: true,
      count: triggers.length,
      trigger_ids: triggers.map((t) => t.getUniqueId()),
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}
function installAutoJurnalSchedulerByAdmin(token) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return {
      status: "error",
      message: "Hanya admin yang dapat memasang trigger.",
    };
  }
  return installAutoJurnalScheduler();
}
function uninstallAutoJurnalSchedulerByAdmin(token) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return {
      status: "error",
      message: "Hanya admin yang dapat menghapus trigger.",
    };
  }
  return uninstallAutoJurnalScheduler();
}
function runReminderScheduler() {
  try {
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    var hour = Number(Utilities.formatDate(now, "Asia/Jakarta", "HH"));
    console.log("[runReminderScheduler] Berjalan: " + dateStr + " jam " + hour);
    var stats = {
      missing_journal: 0,
      missing_bap: 0,
      pending_kbm: 0,
      pending_supervisor: 0,
    };
    try {
      stats.missing_journal = _reminderMissingJournal_(dateStr);
    } catch (e) {
      console.warn("reminder journal: " + e);
    }
    try {
      stats.missing_bap = _reminderMissingBap_(dateStr);
    } catch (e) {
      console.warn("reminder bap: " + e);
    }
    try {
      stats.pending_kbm = _reminderPendingKbm_(dateStr);
    } catch (e) {
      console.warn("reminder pending kbm: " + e);
    }
    try {
      stats.pending_supervisor = _reminderPendingSupervisor_(dateStr);
    } catch (e) {
      console.warn("reminder pending supervisor: " + e);
    }
    console.log(
      "[runReminderScheduler] Selesai. Stats: " + JSON.stringify(stats),
    );
    return { status: "success", date: dateStr, stats: stats };
  } catch (e) {
    console.error("[runReminderScheduler] Error: " + e);
    return { status: "error", message: e && e.message ? e.message : String(e) };
  }
}
function _reminderMissingJournal_(dateStr) {
  // isTeachingDay menentukan apakah ada jadwal KBM reguler hari ini
  var isTeachingDay = _notifIsTeachingDay_(dateStr);
  // Jurnal acara tetap diperiksa meski bukan hari KBM biasa (misal: hari ujian)
  // kecuali hari libur (hari libur berarti tidak ada kegiatan apapun)
  var isHoliday = _notifIsHoliday_(dateStr);
  if (!isTeachingDay && isHoliday) return 0;

  var d = new Date(dateStr + "T00:00:00");
  var dayIdx = d.getDay();
  var cfg = _getConfigMap();
  var activeTP = cfg["tahun_pelajaran"] || "";
  var activeSem = cfg["semester"] || "";

  // --- Ambil jadwal KBM aktif hari ini (hanya jika hari KBM reguler) ---
  var schedules = [];
  if (isTeachingDay) {
    schedules = getData("Schedules").filter(function (s) {
      if (String(s.day_index).trim() !== String(dayIdx)) return false;
      var sTP = s.tahun_pelajaran || activeTP;
      var sSem = s.semester || activeSem;
      return sTP === activeTP && sSem === activeSem;
    });
  }

  // --- Ambil substitusi hari ini ---
  var subs = isTeachingDay ? getData("Substitutes").filter(function (s) {
    try {
      var sd = Utilities.formatDate(
        new Date(s.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return sd === dateStr;
    } catch (_) {
      return false;
    }
  }) : [];

  // Kumpulkan schedule_id yang sudah punya pengganti hari ini
  // agar guru asli tidak diingatkan untuk jadwal yang sudah digantikan
  var substitutedScheduleIds = new Set();
  subs.forEach(function (sub) {
    substitutedScheduleIds.add(String(sub.schedule_id));
  });

  // --- Bangun set log yang sudah masuk hari ini ---
  var logs = (isTeachingDay && schedules.length > 0) ? getData("Teaching_Logs") : [];
  var loggedSet = {};
  logs.forEach(function (l) {
    try {
      var ld = Utilities.formatDate(
        new Date(l.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (ld !== dateStr) return;
      loggedSet[String(l.user_id) + "|" + String(l.schedule_id)] = true;
    } catch (_) {}
  });

  // --- byUser: map userId → daftar item jurnal KBM yang belum diisi ---
  var byUser = {};

  // 1) Jadwal reguler (guru asli, tidak ada pengganti)
  schedules.forEach(function (s) {
    var schedId = String(s.id);
    if (substitutedScheduleIds.has(schedId)) {
      // Jadwal ini sudah digantikan — tangani via blok guru pengganti di bawah
      return;
    }
    var originalUid = String(s.user_id);
    var key = originalUid + "|" + schedId;
    if (loggedSet[key]) return;
    if (!byUser[originalUid]) byUser[originalUid] = { kbm: [], events: [] };
    byUser[originalUid].kbm.push({
      subject: String(s.subject || "-"),
      class_name: String(s.class_name || "-"),
      time: _notifFmtTime_(s.time_start) + "–" + _notifFmtTime_(s.time_end),
    });
  });

  // 2) Guru pengganti — untuk setiap substitusi hari ini,
  //    cek apakah guru pengganti sudah mengisi jurnal
  subs.forEach(function (sub) {
    var schedId = String(sub.schedule_id);
    var substituteUid = String(sub.substitute_user_id);
    // Cari data jadwal yang bersangkutan
    var s = schedules.find(function (sc) {
      return String(sc.id) === schedId;
    });
    if (!s) return; // Jadwal tidak ada di semester aktif — abaikan
    var key = substituteUid + "|" + schedId;
    if (loggedSet[key]) return;
    if (!byUser[substituteUid]) byUser[substituteUid] = { kbm: [], events: [] };
    byUser[substituteUid].kbm.push({
      subject: String(s.subject || "-"),
      class_name: String(s.class_name || "-"),
      time: _notifFmtTime_(s.time_start) + "–" + _notifFmtTime_(s.time_end),
      is_substitute: true,
    });
  });

  // --- Jurnal Acara/Kegiatan yang belum diisi ---
  // Guru perlu mengisi jurnal acara jika:
  //   - sudah dikonfirmasi hadir (confirmed_by terisi)
  //   - belum mengisi jurnal (journal_submitted != true)
  //   - tanggal kehadiran = hari ini
  var eventAttendances = getData(EVENT_SHEET.ATTENDANCE).filter(function (a) {
    try {
      var ad = Utilities.formatDate(
        new Date(a.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (ad !== dateStr) return false;
      // Hanya yang sudah dikonfirmasi hadir
      if (!a.confirmed_by || String(a.confirmed_by).trim() === "") return false;
      // Belum isi jurnal
      var js = a.journal_submitted;
      if (js === true || String(js).toLowerCase() === "true") return false;
      return true;
    } catch (_) {
      return false;
    }
  });

  // Ambil definisi event untuk mendapat nama acara
  var eventDefs = {};
  if (eventAttendances.length > 0) {
    getData(EVENT_SHEET.DEFINITIONS).forEach(function (ev) {
      eventDefs[String(ev.id)] = ev;
    });
  }

  eventAttendances.forEach(function (a) {
    var uid = String(a.user_id || "").trim();
    if (!uid) return;
    var evDef = eventDefs[String(a.event_id)] || {};
    var evName = String(evDef.name || "Acara/Kegiatan");
    var evTime = evDef.time_start
      ? _notifFmtTime_(evDef.time_start) + "–" + _notifFmtTime_(evDef.time_end)
      : "-";
    if (!byUser[uid]) byUser[uid] = { kbm: [], events: [] };
    byUser[uid].events.push({
      event_name: evName,
      time: evTime,
    });
  });

  // --- Kirim notifikasi per guru ---
  var sent = 0;
  Object.keys(byUser).forEach(function (uid) {
    var data = byUser[uid];
    var kbmItems = data.kbm || [];
    var eventItems = data.events || [];
    if (!kbmItems.length && !eventItems.length) return;

    var paragraphs = [];

    // Bagian jurnal KBM
    if (kbmItems.length > 0) {
      paragraphs.push(
        "📚 <strong>Jurnal Mengajar (" + kbmItems.length + " jadwal):</strong>",
      );
      kbmItems.forEach(function (it) {
        var label = "• <strong>" + _escHtml_(it.subject) + "</strong>" +
          " — Kelas " + _escHtml_(it.class_name) +
          " (" + _escHtml_(it.time) + ")";
        if (it.is_substitute) {
          label += ' <span style="color:#8B5CF6;font-size:11px;">(Guru Pengganti)</span>';
        }
        paragraphs.push(label);
      });
    }

    // Bagian jurnal acara/kegiatan
    if (eventItems.length > 0) {
      paragraphs.push(
        "🎌 <strong>Jurnal Acara/Kegiatan (" + eventItems.length + " kegiatan):</strong>",
      );
      eventItems.forEach(function (it) {
        paragraphs.push(
          "• <strong>" + _escHtml_(it.event_name) + "</strong>" +
          " (" + _escHtml_(it.time) + ")",
        );
      });
    }

    paragraphs.push(
      '<span style="color:#94a3b8;font-size:12px;">Mohon segera lengkapi jurnal melalui menu ' +
      '<strong>Jurnal Mengajar</strong>' +
      (eventItems.length > 0 ? " dan <strong>Jurnal Acara</strong>" : "") +
      ".</span>",
    );

    var name = _notifGetName_(uid);
    var totalMissing = kbmItems.length + eventItems.length;
    var html = _notifBuildHtml_({
      title: "⏰ Pengingat: Jurnal Belum Diisi",
      accent: "#F59E0B",
      accent2: "#B45309",
      name: name,
      intro:
        "Kami mendeteksi <strong>" + totalMissing + " jurnal</strong> yang belum Anda lengkapi untuk hari <strong>" +
        _escHtml_(_notifFmtDateLong_(dateStr)) +
        "</strong>:",
      badges: [
        "Tanggal: " + _notifFmtDateLong_(dateStr),
        "Total Belum Diisi: " + totalMissing,
      ],
      paragraphs: paragraphs,
    });
    if (
      _notifToUser_(
        uid,
        "[SiM-Guru] Pengingat Jurnal · " + _notifFmtDate_(dateStr),
        html,
      )
    )
      sent++;
  });
  return sent;
}
function _reminderMissingBap_(dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return 0;
  var period = getData(EXAM_SHEET.PERIODS).find(function (p) {
    return dateStr >= String(p.date_start) && dateStr <= String(p.date_end);
  });
  if (!period) return 0;
  var sessions = getData(EXAM_SHEET.SESSIONS).filter(function (s) {
    return (
      String(s.period_id) === String(period.id) && String(s.date) === dateStr
    );
  });
  if (sessions.length === 0) return 0;
  var rooms = getData(EXAM_SHEET.ROOMS);
  var supervisors = getData(EXAM_SHEET.SUPERVISORS);
  var baps = getData(EXAM_SHEET.BAP);
  var nowTime = Utilities.formatDate(new Date(), "Asia/Jakarta", "HH:mm");
  var byUser = {};
  sessions.forEach(function (session) {
    var te = String(session.time_end || "");
    if (!te) return;
    var teH = Number(te.substring(0, 2));
    var teM = Number(te.substring(3, 5));
    var bapMin = teH * 60 + teM + 5;
    var nowMin =
      Number(nowTime.substring(0, 2)) * 60 + Number(nowTime.substring(3, 5));
    if (nowMin < bapMin) return;
    var roomsInSession = rooms.filter(function (r) {
      return String(r.session_id) === String(session.id);
    });
    roomsInSession.forEach(function (room) {
      var sups = supervisors.filter(function (sup) {
        return (
          String(sup.room_id) === String(room.id) &&
          String(sup.status) === "active"
        );
      });
      sups.forEach(function (sup) {
        var isConf = !!(
          sup.confirmed_by && String(sup.confirmed_by).trim() !== ""
        );
        if (!isConf) return;
        var bap = baps.find(function (b) {
          return String(b.supervisor_id) === String(sup.id);
        });
        var hasBap = !!(
          bap &&
          bap.submitted_at &&
          String(bap.submitted_at).trim() !== ""
        );
        if (hasBap) return;
        var uid = String(sup.user_id);
        if (!byUser[uid]) byUser[uid] = [];
        byUser[uid].push({
          session: String(session.session_name || ""),
          room: String(room.room_name || ""),
          subject: String(room.subject || ""),
          time:
            _notifFmtTime_(session.time_start) +
            "–" +
            _notifFmtTime_(session.time_end),
        });
      });
    });
  });
  var sent = 0;
  Object.keys(byUser).forEach(function (uid) {
    var items = byUser[uid];
    if (!items.length) return;
    var paragraphs = items.map(function (it) {
      return (
        "• <strong>" +
        _escHtml_(it.session) +
        "</strong> di Ruang " +
        _escHtml_(it.room) +
        " — " +
        _escHtml_(it.subject) +
        " (" +
        _escHtml_(it.time) +
        ")"
      );
    });
    var name = _notifGetName_(uid);
    var html = _notifBuildHtml_({
      title: "⏰ Pengingat: BAP Pengawas Ujian Belum Diisi",
      accent: "#F59E0B",
      accent2: "#B45309",
      name: name,
      intro:
        "Sesi ujian yang Anda awasi pada <strong>" +
        _escHtml_(_notifFmtDateLong_(dateStr)) +
        "</strong> sudah berakhir, namun BAP-nya belum disubmit:",
      badges: ["Tanggal: " + _notifFmtDateLong_(dateStr)],
      paragraphs: paragraphs.concat([
        '<span style="color:#94a3b8;font-size:12px;">Mohon segera lengkapi BAP melalui halaman <strong>Jadwal Ujian</strong> atau panel <strong>Panitia Ujian</strong>.</span>',
      ]),
    });
    if (
      _notifToUser_(
        uid,
        "[SiM-Guru] Pengingat BAP Pengawas · " + _notifFmtDate_(dateStr),
        html,
      )
    )
      sent++;
  });
  return sent;
}
function _reminderPendingKbm_(dateStr) {
  if (!_notifIsTeachingDay_(dateStr)) return 0;
  var d = new Date(dateStr + "T00:00:00");
  var dayIdx = d.getDay();
  var cfg = _getConfigMap();
  var activeTP = cfg["tahun_pelajaran"] || "";
  var activeSem = cfg["semester"] || "";
  var picketSchedules = getData("Picket_Schedules").filter(function (p) {
    return String(p.day_index).trim() === String(dayIdx);
  });
  if (picketSchedules.length === 0) return 0;
  var logs = getData("Teaching_Logs");
  var picketIds = picketSchedules.map(function (p) {
    return String(p.user_id);
  });
  var picketLogsToday = logs.filter(function (l) {
    try {
      var ld = Utilities.formatDate(
        new Date(l.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return ld === dateStr && String(l.schedule_id) === "PICKET-DUTY";
    } catch (_) {
      return false;
    }
  });
  picketLogsToday.forEach(function (l) {
    var uid = String(l.user_id);
    if (picketIds.indexOf(uid) < 0) picketIds.push(uid);
  });
  var schedules = getData("Schedules").filter(function (s) {
    if (String(s.day_index).trim() !== String(dayIdx)) return false;
    var sTP = s.tahun_pelajaran || activeTP;
    var sSem = s.semester || activeSem;
    return sTP === activeTP && sSem === activeSem;
  });
  var subs = getData("Substitutes").filter(function (s) {
    var sd = Utilities.formatDate(
      new Date(s.date),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    return sd === dateStr;
  });
  var attendance = getData("Daily_Attendance").filter(function (a) {
    var ad = Utilities.formatDate(
      new Date(a.date),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    return ad === dateStr;
  });
  var pendingUserIds = {};
  schedules.forEach(function (s) {
    var sub = subs.find(function (x) {
      return String(x.schedule_id) === String(s.id);
    });
    var effectiveUid = sub ? String(sub.substitute_user_id) : String(s.user_id);
    var confirmed = attendance.some(function (a) {
      return (
        String(a.user_id) === effectiveUid &&
        String(a.schedule_id || "") === String(s.id)
      );
    });
    if (!confirmed) pendingUserIds[effectiveUid] = true;
  });
  var pendingCount = Object.keys(pendingUserIds).length;
  if (pendingCount === 0) return 0;
  var users = getData("Users");
  var pendingNames = Object.keys(pendingUserIds)
    .slice(0, 8)
    .map(function (uid) {
      var u = users.find(function (x) {
        return String(x.id) === uid;
      });
      return u ? String(u.full_name || u.username) : "Guru";
    });
  var sent = 0;
  picketIds.forEach(function (picketUid) {
    var name = _notifGetName_(picketUid);
    var paragraphs = [
      "<strong>" +
        pendingCount +
        "</strong> guru mata pelajaran masih menunggu konfirmasi kehadiran Anda hari ini.",
      "Daftar singkat: " +
        pendingNames
          .map(function (n) {
            return _escHtml_(n);
          })
          .join(", ") +
        (pendingCount > pendingNames.length
          ? ", dan " + (pendingCount - pendingNames.length) + " lainnya"
          : ""),
      '<span style="color:#94a3b8;font-size:12px;">Buka halaman <strong>Piket &amp; Upacara</strong> untuk segera memprosesnya.</span>',
    ];
    var html = _notifBuildHtml_({
      title: "⏰ Pengingat: Konfirmasi Kehadiran Guru",
      accent: "#3B82F6",
      accent2: "#1D4ED8",
      name: name,
      intro:
        "Sebagai Guru Piket pada <strong>" +
        _escHtml_(_notifFmtDateLong_(dateStr)) +
        "</strong>, Anda perlu mengonfirmasi kehadiran rekan-rekan guru.",
      badges: ["Tanggal: " + _notifFmtDateLong_(dateStr)],
      paragraphs: paragraphs,
    });
    if (
      _notifToUser_(
        picketUid,
        "[SiM-Guru] Pengingat Konfirmasi Kehadiran Guru · " +
          _notifFmtDate_(dateStr),
        html,
      )
    )
      sent++;
  });
  return sent;
}
function _reminderPendingSupervisor_(dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return 0;
  var period = getData(EXAM_SHEET.PERIODS).find(function (p) {
    return dateStr >= String(p.date_start) && dateStr <= String(p.date_end);
  });
  if (!period) return 0;
  var committee = getData(EXAM_SHEET.COMMITTEE).filter(function (c) {
    return (
      String(c.date) === dateStr &&
      String(c.status) === "active" &&
      c.confirmed_by &&
      String(c.confirmed_by).trim() !== ""
    );
  });
  if (committee.length === 0) return 0;
  var sessions = getData(EXAM_SHEET.SESSIONS).filter(function (s) {
    return (
      String(s.period_id) === String(period.id) && String(s.date) === dateStr
    );
  });
  if (sessions.length === 0) return 0;
  var rooms = getData(EXAM_SHEET.ROOMS);
  var supervisors = getData(EXAM_SHEET.SUPERVISORS);
  var users = getData("Users");
  var pendingNames = [];
  sessions.forEach(function (session) {
    var roomsInSession = rooms.filter(function (r) {
      return String(r.session_id) === String(session.id);
    });
    roomsInSession.forEach(function (room) {
      supervisors
        .filter(function (sup) {
          return (
            String(sup.room_id) === String(room.id) &&
            String(sup.status) === "active"
          );
        })
        .forEach(function (sup) {
          var isConf = !!(
            sup.confirmed_by && String(sup.confirmed_by).trim() !== ""
          );
          if (isConf) return;
          var u = users.find(function (x) {
            return String(x.id) === String(sup.user_id);
          });
          var name = u ? String(u.full_name || u.username) : "Pengawas";
          pendingNames.push(
            name +
              " — " +
              (session.session_name || "") +
              " / " +
              (room.room_name || ""),
          );
        });
    });
  });
  if (pendingNames.length === 0) return 0;
  var sent = 0;
  committee.forEach(function (c) {
    var picketUid = String(c.user_id);
    var name = _notifGetName_(picketUid);
    var paragraphs = [
      "<strong>" +
        pendingNames.length +
        "</strong> pengawas ruang ujian masih menunggu konfirmasi kehadiran.",
      pendingNames
        .slice(0, 8)
        .map(function (n) {
          return "• " + _escHtml_(n);
        })
        .join("<br>") +
        (pendingNames.length > 8
          ? '<br><span style="color:#94a3b8">… dan ' +
            (pendingNames.length - 8) +
            " lainnya</span>"
          : ""),
      '<span style="color:#94a3b8;font-size:12px;">Buka halaman <strong>Panitia Ujian</strong> untuk segera memprosesnya.</span>',
    ];
    var html = _notifBuildHtml_({
      title: "⏰ Pengingat: Konfirmasi Pengawas Ruang Ujian",
      accent: "#0891B2",
      accent2: "#0369A1",
      name: name,
      intro:
        "Sebagai Panitia Ujian pada <strong>" +
        _escHtml_(_notifFmtDateLong_(dateStr)) +
        "</strong>, Anda perlu mengonfirmasi kehadiran para pengawas ruang ujian.",
      badges: ["Tanggal: " + _notifFmtDateLong_(dateStr)],
      paragraphs: paragraphs,
    });
    if (
      _notifToUser_(
        picketUid,
        "[SiM-Guru] Pengingat Konfirmasi Pengawas Ujian · " +
          _notifFmtDate_(dateStr),
        html,
      )
    )
      sent++;
  });
  return sent;
}
function formatPeriode(periodeValue) {
  if (!periodeValue) return "Unknown";
  const monthNames = [
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
  const parts = String(periodeValue).split(" ");
  if (
    parts.length === 2 &&
    monthNames.includes(parts[0]) &&
    !isNaN(parseInt(parts[1]))
  ) {
    return periodeValue;
  }
  let date;
  if (periodeValue instanceof Date) {
    date = periodeValue;
  } else {
    date = new Date(periodeValue);
  }
  if (!isNaN(date.getTime())) {
    try {
      const monthIndex =
        parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), "M")) -
        1;
      const year = Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "yyyy",
      );
      if (monthIndex >= 0 && monthIndex < 12) {
        return `${monthNames[monthIndex]} ${year}`;
      }
    } catch (e) {
      const month = monthNames[date.getMonth()];
      const year = date.getFullYear();
      return `${month} ${year}`;
    }
  }
  return String(periodeValue);
}
function getPeriodeSortKey(periodeValue) {
  const monthNames = [
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
  let date;
  if (periodeValue instanceof Date) {
    date = periodeValue;
  } else {
    date = new Date(periodeValue);
  }
  if (!isNaN(date.getTime())) {
    return date.getFullYear() * 12 + date.getMonth();
  }
  const parts = String(periodeValue).split(" ");
  if (parts.length === 2) {
    const monthIndex = monthNames.indexOf(parts[0]);
    const year = parseInt(parts[1]);
    if (monthIndex !== -1 && !isNaN(year)) {
      return year * 12 + monthIndex;
    }
  }
  return 0;
}
let _cachedData = null;
let _cacheTime = 0;
const CACHE_TTL = 300000;
function _invalidateDataSnapshot() {
  _cachedData = null;
  _cacheTime = 0;
}
function _getDataSnapshot(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && _cachedData && now - _cacheTime < CACHE_TTL) {
    return _cachedData;
  }
  const snapshot = {
    _cache: {},
    get users() { if (!this._cache.users) this._cache.users = getData("Users"); return this._cache.users; },
    get schedules() { if (!this._cache.schedules) this._cache.schedules = getData("Schedules"); return this._cache.schedules; },
    get logs() { if (!this._cache.logs) this._cache.logs = getData("Teaching_Logs"); return this._cache.logs; },
    get allowances() { if (!this._cache.allowances) this._cache.allowances = getData("Allowances"); return this._cache.allowances; },
    get config() { if (!this._cache.config) this._cache.config = _getConfigMap(); return this._cache.config; },
    get holidays() { if (!this._cache.holidays) this._cache.holidays = getData("Academic_Calendar"); return this._cache.holidays; },
    get honorHistory() { if (!this._cache.honorHistory) this._cache.honorHistory = getData("Honor_History"); return this._cache.honorHistory; },
    get pickets() { if (!this._cache.pickets) this._cache.pickets = getData("Picket_Schedules"); return this._cache.pickets; },
    get ceremonies() { if (!this._cache.ceremonies) this._cache.ceremonies = getData("Ceremony_Schedules"); return this._cache.ceremonies; },
    get eventAttendance() { if (!this._cache.eventAttendance) this._cache.eventAttendance = getData("Event_Attendance"); return this._cache.eventAttendance; },
    get attendance() { if (!this._cache.attendance) this._cache.attendance = getData("Daily_Attendance"); return this._cache.attendance; },
    get substitutes() { if (!this._cache.substitutes) this._cache.substitutes = getData("Substitutes"); return this._cache.substitutes; },
    get attendanceLeaves() { if (!this._cache.attendanceLeaves) this._cache.attendanceLeaves = getData("Attendance_Leaves"); return this._cache.attendanceLeaves; },

    get usersById() {
      if (!this._cache.usersById) {
        this._cache.usersById = {};
        this.users.forEach((u) => { this._cache.usersById[u.id] = u; });
      }
      return this._cache.usersById;
    },
    get schedulesById() {
      if (!this._cache.schedulesById) {
        this._cache.schedulesById = {};
        this.schedules.forEach((s) => { this._cache.schedulesById[s.id] = s; });
      }
      return this._cache.schedulesById;
    },
    get logsByUserId() {
      if (!this._cache.logsByUserId) {
        this._cache.logsByUserId = {};
        this.logs.forEach((l) => {
          if (!this._cache.logsByUserId[l.user_id]) this._cache.logsByUserId[l.user_id] = [];
          this._cache.logsByUserId[l.user_id].push(l);
        });
      }
      return this._cache.logsByUserId;
    },
    get allowancesByUserId() {
      if (!this._cache.allowancesByUserId) {
        this._cache.allowancesByUserId = {};
        this.allowances.forEach((a) => {
          if (!this._cache.allowancesByUserId[a.user_id]) this._cache.allowancesByUserId[a.user_id] = [];
          this._cache.allowancesByUserId[a.user_id].push(a);
        });
      }
      return this._cache.allowancesByUserId;
    },
    get eventAttendanceByUserId() {
      if (!this._cache.eventAttendanceByUserId) {
        this._cache.eventAttendanceByUserId = {};
        this.eventAttendance.forEach((a) => {
          if (!this._cache.eventAttendanceByUserId[a.user_id]) this._cache.eventAttendanceByUserId[a.user_id] = [];
          this._cache.eventAttendanceByUserId[a.user_id].push(a);
        });
      }
      return this._cache.eventAttendanceByUserId;
    },
    get attendanceByUserId() {
      if (!this._cache.attendanceByUserId) {
        this._cache.attendanceByUserId = {};
        this.attendance.forEach((a) => {
          if (!this._cache.attendanceByUserId[a.user_id]) this._cache.attendanceByUserId[a.user_id] = [];
          this._cache.attendanceByUserId[a.user_id].push(a);
        });
      }
      return this._cache.attendanceByUserId;
    },
    get attendanceLeavesByUserId() {
      if (!this._cache.attendanceLeavesByUserId) {
        this._cache.attendanceLeavesByUserId = {};
        this.attendanceLeaves.forEach((lv) => {
          if (!this._cache.attendanceLeavesByUserId[lv.user_id]) this._cache.attendanceLeavesByUserId[lv.user_id] = [];
          this._cache.attendanceLeavesByUserId[lv.user_id].push(lv);
        });
      }
      return this._cache.attendanceLeavesByUserId;
    },
    get substitutesByUserId() {
      if (!this._cache.substitutesByUserId) {
        this._cache.substitutesByUserId = {};
        this._cache.substitutesByScheduleId = {};
        this.substitutes.forEach((sub) => {
          if (!this._cache.substitutesByUserId[sub.substitute_user_id]) this._cache.substitutesByUserId[sub.substitute_user_id] = [];
          this._cache.substitutesByUserId[sub.substitute_user_id].push(sub);
          
          if (!this._cache.substitutesByScheduleId[sub.schedule_id]) this._cache.substitutesByScheduleId[sub.schedule_id] = [];
          this._cache.substitutesByScheduleId[sub.schedule_id].push(sub);
        });
      }
      return this._cache.substitutesByUserId;
    },
    get substitutesByScheduleId() {
      if (!this._cache.substitutesByScheduleId) {
        let dummy = this.substitutesByUserId; 
      }
      return this._cache.substitutesByScheduleId;
    }
  };
  _cachedData = snapshot;
  _cacheTime = now;
  return snapshot;
}
function getHonorariumSlipByHistory(token, periode, userId) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return { status: "error", message: "Unauthorized" };
  }
  const allHistory = getData("Honor_History");
  const normalizedPeriode = formatPeriode(String(periode || "").trim());
  const normalizedUserId = String(userId || "").trim();
  const history = allHistory.find((h) => {
    if (!h || !h.periode || !h.user_id) return false;
    const hPeriode = formatPeriode(String(h.periode).trim());
    const hUserId = String(h.user_id).trim();
    return hPeriode === normalizedPeriode && hUserId === normalizedUserId;
  });
  if (!history) {
    const sample = allHistory.slice(0, 3).map((h) => ({
      p: String(h.periode),
      u: String(h.user_id),
    }));
    console.error(
      "[getHonorariumSlipByHistory] Tidak ditemukan. " +
        "normalizedPeriode=" +
        normalizedPeriode +
        ", normalizedUserId=" +
        normalizedUserId +
        ", sampel allHistory=" +
        JSON.stringify(sample),
    );
    return {
      status: "error",
      message:
        "Data honorarium tidak ditemukan untuk periode: " + normalizedPeriode,
    };
  }
  const teacher = findData("Users", "id", userId);
  const config = _getConfigMap();
  const baseSalary = Number(config["base_salary"] || 0);
  const schoolName = config["app_school_name"] || "MTs Nurul Falah";
  const kepalaSekolah = config["kepala_sekolah"] || "[Nama Kepala Sekolah]";
  let allowanceItems = [];
  try {
    const parsed = JSON.parse(history.details_json || "{}");
    if (
      parsed.allowances &&
      Array.isArray(parsed.allowances) &&
      parsed.allowances.length > 0
    ) {
      const firstItem = parsed.allowances[0];
      if (
        typeof firstItem === "object" &&
        firstItem !== null &&
        "duty_name" in firstItem
      ) {
        allowanceItems = parsed.allowances.map((item) => ({
          desc: String(item.duty_name),
          qty: 1,
          rate: formatRupiah(Number(item.amount || 0)),
          amount: formatRupiah(Number(item.amount || 0)),
        }));
      } else {
        const allAllowances = getData("Allowances");
        allowanceItems = parsed.allowances.map((dutyName) => {
          const found = allAllowances.find(
            (a) =>
              String(a.user_id).trim() === normalizedUserId &&
              String(a.duty_name).trim().toLowerCase() ===
                String(dutyName).trim().toLowerCase(),
          );
          return {
            desc: String(dutyName),
            qty: 1,
            rate: found ? formatRupiah(Number(found.amount)) : "-",
            amount: found
              ? formatRupiah(Number(found.amount))
              : formatRupiah(0),
          };
        });
      }
    } else if (parsed.allow && Number(parsed.allow) > 0) {
      allowanceItems = [
        {
          desc: "Tunjangan Tambahan",
          qty: 1,
          rate: formatRupiah(Number(parsed.allow)),
          amount: formatRupiah(Number(parsed.allow)),
        },
      ];
    }
  } catch (e) {
    console.error(
      "[getHonorariumSlipByHistory] Gagal parse details_json: " + e.toString(),
    );
  }
  if (allowanceItems.length === 0) {
    const currentAllowances = getData("Allowances").filter(
      (a) => String(a.user_id).trim() === normalizedUserId,
    );
    allowanceItems = currentAllowances.map((a) => ({
      desc: String(a.duty_name),
      qty: 1,
      rate: formatRupiah(Number(a.amount)),
      amount: formatRupiah(Number(a.amount)),
    }));
  }
  const trxIdStr = String(history.trx_id || "-");
  const isFinal =
    trxIdStr && trxIdStr.indexOf("ESTIMASI") === -1 && trxIdStr !== "-";
  const transportEnabled =
    String(config["transport_allowance_enabled"] || "false")
      .toLowerCase()
      .trim() === "true";
  let transportTotal = 0;
  let transportStr = "";
  if (isFinal) {
    try {
      const parsedForTransport = JSON.parse(history.details_json || "{}");
      if (
        parsedForTransport.transport_total &&
        Number(parsedForTransport.transport_total) > 0
      ) {
        transportTotal = Math.round(Number(parsedForTransport.transport_total));
        transportStr = formatRupiah(transportTotal);
      }
    } catch (e) {}
  } else {
    if (transportEnabled) {
      transportTotal = _getMonthlyTransportSum(
        normalizedUserId,
        normalizedPeriode,
      );
      if (transportTotal > 0) {
        transportStr = formatRupiah(transportTotal);
      }
    }
  }
  if (transportTotal > 0) {
    allowanceItems.push({
      desc: "Tunjangan Transportasi",
      qty: 1,
      rate: transportStr,
      amount: transportStr,
    });
  }
  var eventJtm = 0;
  try {
    var _monthNames = [
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
    var _pParts = String(normalizedPeriode || "")
      .trim()
      .split(" ");
    var _targetMonthIdx = _monthNames.indexOf(_pParts[0]);
    var _targetYear = parseInt(_pParts[1], 10);
    if (_targetMonthIdx !== -1 && !isNaN(_targetYear)) {
      var _mm =
        _targetMonthIdx + 1 < 10
          ? "0" + (_targetMonthIdx + 1)
          : String(_targetMonthIdx + 1);
      var _firstDay = _targetYear + "-" + _mm + "-01";
      var _lastDd = new Date(_targetYear, _targetMonthIdx + 1, 0).getDate();
      var _lastDay =
        _targetYear +
        "-" +
        _mm +
        "-" +
        (_lastDd < 10 ? "0" + _lastDd : String(_lastDd));
      var _evtAtt = getData(EVENT_SHEET.ATTENDANCE).filter(function (r) {
        return (
          String(r.user_id || "") === normalizedUserId &&
          String(r.date || "") >= _firstDay &&
          String(r.date || "") <= _lastDay
        );
      });
      eventJtm = _computeEventJtm(_evtAtt);
    }
  } catch (_evtErr) {
    console.error("getHonorariumSlipByHistory eventJtm error: " + _evtErr);
  }
  const totalJtm = Number(history.total_jtm) || 0;
  let totalHonor = Number(history.total_honor) || 0;
  // Ambil event_jtm dari details_json jika final, agar baris tabel tidak dobel
  let savedEventJtm = 0;
  if (isFinal) {
    try {
      const parsedDetails = JSON.parse(history.details_json || "{}");
      savedEventJtm = Number(parsedDetails.event_jtm || 0);
    } catch (e) {}
  }
  const kbmJtm = isFinal ? (totalJtm - savedEventJtm) : totalJtm;
  const displayEventJtm = isFinal ? savedEventJtm : eventJtm;
  const honorFromJtm = kbmJtm * baseSalary;
  if (!isFinal) {
    if (transportTotal > 0) totalHonor += transportTotal;
    if (eventJtm > 0) totalHonor += eventJtm * baseSalary;
  }
  return {
    status: "success",
    periode: formatPeriode(String(history.periode).trim()),
    guru_nama: teacher ? String(teacher.full_name) : "Unknown",
    guru_nip: teacher ? String(teacher.nip || "-") : "-",
    tgl_cetak: formatDateIndo(new Date()),
    total_real: formatRupiah(honorFromJtm),
    qty_real: kbmJtm,
    jtm_event: displayEventJtm,
    jtm_event_amount: formatRupiah(displayEventJtm * baseSalary),
    tarif: formatRupiah(baseSalary),
    grand_total: formatRupiah(totalHonor),
    allowance_list: allowanceItems,
    trx_id: trxIdStr,
    kepala_sekolah: kepalaSekolah,
    school_name: schoolName,
    printed_by: String(user.full_name),
  };
}
function addHolidayRange(token, dateFrom, dateTo, desc, isHoliday) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  if (!dateFrom || !dateTo || !desc)
    return { status: "error", message: "Data tidak lengkap" };
  const dFrom = new Date(dateFrom);
  const dTo = new Date(dateTo);
  if (dTo < dFrom)
    return {
      status: "error",
      message: "Tanggal selesai tidak boleh kurang dari tanggal mulai",
    };
  const sheet = getSheet("Academic_Calendar");
  const isHolStr = isHoliday ? "True" : "False";
  let count = 0;
  const current = new Date(dFrom);
  while (current <= dTo) {
    const dateStr = Utilities.formatDate(
      current,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    const newId = generateId("HOL");
    sheet.appendRow([newId, dateStr, desc, isHolStr]);
    count++;
    current.setDate(current.getDate() + 1);
  }
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  return {
    status: "success",
    message: `Berhasil menambahkan ${count} hari libur`,
    count: count,
  };
}
function getAdminDashboardExtendedStats(token) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin")
      return { status: "error" };
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    const dayIndex = _getIndoDayIndex();
    const allLogs = getData("Teaching_Logs");
    const allSchedules = getData("Schedules");
    const allUsers = getData("Users");
    const allSubstitutes = getData("Substitutes");
    const attendanceData = getData("Daily_Attendance");
    const cfg = _getConfigMap();
    const activeTP = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    const todayLogs = allLogs.filter((l) => {
      const d = new Date(l.date);
      const dStr = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
      return dStr === todayStr;
    });
    const piketAdjMapToday = _jtmBuildPiketAdjustmentMap_();
    let totalJtmToday = 0;
    todayLogs.forEach((l) => {
      const sid = String(l.schedule_id);
      const sidUpper = sid.toUpperCase();
      if (sid === "PICKET-DUTY" || sid === "PIKET") {
        // Gunakan adjusted JTM jika ada, fallback ke 4
        totalJtmToday += _jtmResolvePiketLogJtm_(piketAdjMapToday, l.user_id, l.date, 4);
      } else if (sid === "CEREMONY-DUTY" || sidUpper.includes("UPACARA")) {
        totalJtmToday += 5;
      } else {
        const jtm = Number(l.jtm_val || 0);
        if (!isNaN(jtm) && jtm > 0) totalJtmToday += jtm;
        else if (sidUpper.includes("PICKET") || sidUpper.includes("PIKET")) totalJtmToday += 4;
        else if (sidUpper.includes("CEREMONY") || sidUpper.includes("UPACARA")) totalJtmToday += 5;
      }
    });
    const allEventAtt = getData("Event_Attendance");
    const todayEventAtt = allEventAtt.filter((a) => {
      if (
        a.journal_submitted !== true &&
        String(a.journal_submitted).toLowerCase() !== "true"
      )
        return false;
      const d = new Date(a.date);
      const dStr = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
      return dStr === todayStr;
    });
    todayEventAtt.forEach((a) => {
      const jtm = Number(a.jtm_val || 0);
      if (!isNaN(jtm) && jtm > 0) totalJtmToday += jtm;
    });
    const todaySchedules = allSchedules.filter((s) => {
      const isToday = String(s.day_index).trim() === String(dayIndex);
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return isToday && sTP === activeTP && sSem === activeSem;
    });
    const todayTeacherIds = [
      ...new Set(todaySchedules.map((s) => String(s.user_id))),
    ];
    const todaySubs = allSubstitutes.filter((s) => {
      const sd = Utilities.formatDate(
        new Date(s.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return sd === todayStr;
    });
    const subTeacherIds = todaySubs.map((s) => String(s.substitute_user_id));
    const allTodayIds = [...new Set([...todayTeacherIds, ...subTeacherIds])];
    const confirmedTeacherIds = new Set();
    attendanceData.forEach((a) => {
      const ad = Utilities.formatDate(
        new Date(a.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (ad === todayStr && a.status === "Present" && a.schedule_id) {
        confirmedTeacherIds.add(String(a.user_id));
      }
    });
    const subsMapToday = {};
    todaySubs.forEach((sub) => {
      subsMapToday[String(sub.schedule_id)] = sub;
    });
    const uniqueTeacherMap = {};
    todaySchedules.forEach((s) => {
      const schedId = String(s.id);
      const subEntry = subsMapToday[schedId];
      if (subEntry) {
        const subUid = String(subEntry.substitute_user_id);
        const subUser = allUsers.find((u) => String(u.id) === subUid);
        const hasDone = allLogs.some((l) => {
          const ld = Utilities.formatDate(
            new Date(l.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          );
          return (
            ld === todayStr &&
            String(l.user_id) === subUid &&
            String(l.schedule_id) === schedId
          );
        });
        if (!uniqueTeacherMap[subUid]) {
          uniqueTeacherMap[subUid] = {
            user_id: subUid,
            name: subUser ? String(subUser.full_name) : "Unknown",
            is_substitute: true,
            schedules: [],
          };
        }
        uniqueTeacherMap[subUid].schedules.push({
          subject: s.subject,
          class_name: s.class_name,
          has_journal: hasDone,
        });
      } else {
        const origUid = String(s.user_id);
        const origUser = allUsers.find((u) => String(u.id) === origUid);
        const hasDone = allLogs.some((l) => {
          const ld = Utilities.formatDate(
            new Date(l.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          );
          return (
            ld === todayStr &&
            String(l.user_id) === origUid &&
            String(l.schedule_id) === schedId
          );
        });
        if (!uniqueTeacherMap[origUid]) {
          uniqueTeacherMap[origUid] = {
            user_id: origUid,
            name: origUser ? String(origUser.full_name) : "Unknown",
            is_substitute: false,
            schedules: [],
          };
        }
        uniqueTeacherMap[origUid].schedules.push({
          subject: s.subject,
          class_name: s.class_name,
          has_journal: hasDone,
        });
      }
    });
    const todayAllEvents = allEventAtt.filter((a) => {
      try {
        return (
          Utilities.formatDate(
            new Date(a.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          ) === todayStr
        );
      } catch (e) {
        return false;
      }
    });
    todayAllEvents.forEach((e) => {
      const uid = String(e.user_id);
      const isSubmitted = String(e.journal_submitted).toLowerCase() === "true";
      if (!uniqueTeacherMap[uid]) {
        const uObj = allUsers.find((x) => String(x.id) === uid);
        uniqueTeacherMap[uid] = {
          user_id: uid,
          name: uObj ? String(uObj.full_name) : "Unknown",
          is_substitute: false,
          schedules: [],
        };
      }
      uniqueTeacherMap[uid].schedules.push({
        subject: e.event_name || "Acara Sekolah",
        class_name: "Kegiatan",
        has_journal: isSubmitted,
      });
    });
    const teacherList = Object.values(uniqueTeacherMap);
    const doneTeachers = teacherList.filter((t) =>
      t.schedules.every((s) => s.has_journal),
    );
    const notDoneTeachers = teacherList.filter((t) =>
      t.schedules.some((s) => !s.has_journal),
    );
    const pendingResets = getData("Reset_Requests").filter(
      (r) => String(r.status).toLowerCase() === "pending",
    ).length;
    const activePeriod = getData(EXAM_SHEET.PERIODS).find(
      (p) => todayStr >= String(p.date_start) && todayStr <= String(p.date_end),
    );
    let is_exam_period = !!activePeriod;
    let exam_bap_done = 0;
    let exam_bap_total = 0;
    let exam_bap_list = [];
    let exam_hadir_count = 0;
    let exam_hadir_total = 0;
    let exam_substitutes = 0;
    if (is_exam_period) {
      const allExamSessions = getData(EXAM_SHEET.SESSIONS);
      const allExamRooms = getData(EXAM_SHEET.ROOMS);
      const allExamSups = getData(EXAM_SHEET.SUPERVISORS);
      const allExamBaps = getData(EXAM_SHEET.BAP);
      const todayExamSessions = allExamSessions.filter(
        (s) =>
          String(s.period_id) === String(activePeriod.id) &&
          String(s.date) === todayStr,
      );
      const todaySessionIds = new Set(
        todayExamSessions.map((s) => String(s.id)),
      );
      const todayExamRooms = allExamRooms.filter((r) =>
        todaySessionIds.has(String(r.session_id)),
      );
      const todayRoomIds = new Set(todayExamRooms.map((r) => String(r.id)));
      const todaySups = allExamSups.filter(
        (sup) =>
          todayRoomIds.has(String(sup.room_id)) &&
          String(sup.status) === "active",
      );
      exam_bap_total = todaySups.length;
      todaySups.forEach((sup) => {
        const bap = allExamBaps.find(
          (b) => String(b.supervisor_id) === String(sup.id),
        );
        const hasBap = !!(
          bap &&
          bap.submitted_at &&
          String(bap.submitted_at).trim() !== ""
        );
        if (hasBap) exam_bap_done++;
        const userObj = allUsers.find(
          (u) => String(u.id) === String(sup.user_id),
        );
        const userName = userObj ? String(userObj.full_name) : "Unknown";
        const roomObj = todayExamRooms.find(
          (r) => String(r.id) === String(sup.room_id),
        );
        const roomName = roomObj ? String(roomObj.room_name) : "Unknown Room";
        const sesObj = todayExamSessions.find(
          (s) => String(s.id) === String(roomObj ? roomObj.session_id : ""),
        );
        const sesName = sesObj
          ? String(sesObj.session_name)
          : "Unknown Session";
        exam_bap_list.push({
          name: userName,
          room: roomName,
          session: sesName,
          has_bap: hasBap,
          is_substitute: !!sup.is_substitute,
        });
      });
      const allExamCom = getData(EXAM_SHEET.COMMITTEE);
      const todayCom = allExamCom.filter(
        (c) =>
          String(c.period_id) === String(activePeriod.id) &&
          String(c.date) === todayStr &&
          String(c.status) === "active",
      );
      const presentSups = todaySups.filter(
        (sup) => sup.confirmed_at && String(sup.confirmed_at).trim() !== "",
      ).length;
      const presentComs = todayCom.filter(
        (c) => c.confirmed_at && String(c.confirmed_at).trim() !== "",
      ).length;
      const subsSups = todaySups.filter(
        (sup) =>
          sup.is_substitute === true || String(sup.is_substitute) === "true",
      ).length;
      const subsComs = todayCom.filter(
        (c) => c.is_substitute === true || String(c.is_substitute) === "true",
      ).length;
      exam_hadir_count = presentSups + presentComs;
      exam_hadir_total = todaySups.length + todayCom.length;
      exam_substitutes = subsSups + subsComs;
    }
    return {
      status: "success",
      total_jtm_today: totalJtmToday,
      hadir_count: is_exam_period ? exam_hadir_count : confirmedTeacherIds.size,
      hadir_total: is_exam_period ? exam_hadir_total : allTodayIds.length,
      substitutes_today: is_exam_period ? exam_substitutes : todaySubs.length,
      pending_resets: pendingResets,
      journal_done: is_exam_period ? exam_bap_done : doneTeachers.length,
      journal_total: is_exam_period ? exam_bap_total : teacherList.length,
      teachers_not_done: notDoneTeachers.map((t) => ({
        name: t.name,
        is_substitute: !!t.is_substitute,
        subjects_remaining: t.schedules
          .filter((s) => !s.has_journal)
          .map((s) => s.subject + " " + s.class_name),
      })),
      is_exam_period: is_exam_period,
      exam_bap_done: exam_bap_done,
      exam_bap_total: exam_bap_total,
      exam_bap_list: exam_bap_list,
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

// ============================================================
// PIKET DASHBOARD ADMIN — fungsi khusus untuk manajemen piket
// langsung dari halaman Dashboard Admin
// ============================================================

/**
 * Mengambil status piket hari ini beserta daftar guru piket,
 * guru piket pengganti, dan data jam masuk/pulang mereka.
 * Dipakai oleh panel piket di Dashboard Admin.
 */
function getAdminDashboardPicketStatus(token) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin")
      return { status: "error", message: "Akses Ditolak." };

    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    const dayIndex = _getIndoDayIndex();

    const snap = _getDataSnapshot();
    const picketSchedules = getData("Picket_Schedules");
    const substitutes = snap.substitutes;
    const users = snap.users;
    const logs = snap.logs;
    const attendance = snap.attendance;

    // Jadwal piket rutin hari ini
    const todayPicketSchedules = picketSchedules.filter(
      p => String(p.day_index).trim() === String(dayIndex)
    );

    // Cari guru piket pengganti hari ini (dari sheet Substitutes yang terkait Picket_Schedules)
    const todayPicketSubs = substitutes.filter(s => {
      if (safeDate(s.date) !== todayStr) return false;
      return picketSchedules.some(p => String(p.id) === String(s.schedule_id));
    });

    // Bangun daftar piket hari ini (asli + pengganti)
    const picketItems = [];

    todayPicketSchedules.forEach(ps => {
      const guru = users.find(u => String(u.id) === String(ps.user_id));
      const name = guru ? guru.full_name : "Unknown";

      // Cek apakah sudah ada pengganti
      const subRecord = todayPicketSubs.find(s => String(s.schedule_id) === String(ps.id));

      // Cek konfirmasi kehadiran piket (dari Teaching_Logs PICKET-DUTY)
      const isConfirmed = logs.some(
        l => String(l.user_id) === String(ps.user_id) &&
             String(l.schedule_id) === "PICKET-DUTY" &&
             safeDate(l.date) === todayStr
      );

      // Ambil record Daily_Attendance untuk jam masuk/pulang
      const attRecord = attendance.find(
        a => String(a.user_id) === String(ps.user_id) &&
             safeDate(a.date) === todayStr &&
             (!a.schedule_id || a.schedule_id === "" || a.schedule_id === "PICKET-DUTY")
      );
      const timeIn  = attRecord ? safeTime(attRecord.time_in  || "").substring(0, 5) : null;
      const timeOut = attRecord ? safeTime(attRecord.time_out || "").substring(0, 5) : null;
      const attId   = attRecord ? String(attRecord.id) : null;

      // Jika ada pengganti yang sudah dikonfirmasi, guru asli tidak perlu dikonfirmasi lagi
      let subGuruName = null;
      let subUserId = null;
      let subConfirmed = false;
      let subTimeIn = null;
      let subTimeOut = null;
      let subAttId = null;
      let subId = null;

      if (subRecord) {
        const subGuru = users.find(u => String(u.id) === String(subRecord.substitute_user_id));
        subGuruName = subGuru ? subGuru.full_name : "Unknown";
        subUserId = String(subRecord.substitute_user_id);
        subId = String(subRecord.id);

        subConfirmed = logs.some(
          l => String(l.user_id) === subUserId &&
               String(l.schedule_id) === "PICKET-DUTY" &&
               safeDate(l.date) === todayStr
        );

        const subAttRecord = attendance.find(
          a => String(a.user_id) === subUserId &&
               safeDate(a.date) === todayStr &&
               (!a.schedule_id || a.schedule_id === "" || a.schedule_id === "PICKET-DUTY")
        );
        subTimeIn  = subAttRecord ? safeTime(subAttRecord.time_in  || "").substring(0, 5) : null;
        subTimeOut = subAttRecord ? safeTime(subAttRecord.time_out || "").substring(0, 5) : null;
        subAttId   = subAttRecord ? String(subAttRecord.id) : null;
      }

      picketItems.push({
        picket_schedule_id: String(ps.id),
        user_id:            String(ps.user_id),
        name:               name,
        is_confirmed:       isConfirmed,
        time_in:            timeIn,
        time_out:           timeOut,
        att_id:             attId,
        has_substitute:     !!subRecord,
        substitute: subRecord ? {
          sub_record_id: subId,
          user_id:       subUserId,
          name:          subGuruName,
          is_confirmed:  subConfirmed,
          time_in:       subTimeIn,
          time_out:      subTimeOut,
          att_id:        subAttId,
        } : null,
      });
    });

    // Ambil daftar semua guru (bukan admin) untuk dropdown penunjukan pengganti
    const teacherList = users
      .filter(u => String(u.role).toLowerCase() === "guru")
      .map(u => ({ id: String(u.id), name: u.full_name }))
      .sort((a, b) => a.name.localeCompare(b.name));

    return {
      status:       "success",
      today_str:    todayStr,
      day_index:    dayIndex,
      picket_items: picketItems,
      teacher_list: teacherList,
    };
  } catch (e) {
    return { status: "error", message: "Server Error (PicketDash): " + e.toString() };
  }
}

/**
 * Input atau update jam masuk / jam pulang guru piket / pengganti
 * langsung dari Dashboard Admin (hanya admin yang boleh).
 *
 * payload: { user_id, time_in?, time_out?, att_id? }
 * - Jika att_id ada, update baris yang sudah ada.
 * - Jika tidak ada, buat baris baru di Daily_Attendance.
 */
function adminInputPicketTime(token, payload) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin")
      return { status: "error", message: "Akses Ditolak. Hanya admin yang dapat melakukan ini." };

    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    const targetUserId = String((payload && payload.user_id) || "").trim();
    if (!targetUserId) return { status: "error", message: "user_id wajib diisi." };

    const timeIn  = String((payload && payload.time_in)  || "").trim();
    const timeOut = String((payload && payload.time_out) || "").trim();
    const attId   = String((payload && payload.att_id)   || "").trim();

    // Validasi format jam
    const HH_MM = /^([01]\d|2[0-3]):[0-5]\d$/;
    if (timeIn  && !HH_MM.test(timeIn))  return { status: "error", message: "Format Jam Masuk tidak valid (HH:mm)." };
    if (timeOut && !HH_MM.test(timeOut)) return { status: "error", message: "Format Jam Pulang tidak valid (HH:mm)." };

    // Validasi jam masuk < jam pulang jika keduanya diisi
    if (timeIn && timeOut) {
      const [ih, im] = timeIn.split(":").map(Number);
      const [oh, om] = timeOut.split(":").map(Number);
      if (ih * 60 + im >= oh * 60 + om)
        return { status: "error", message: "Jam Masuk harus lebih awal dari Jam Pulang." };
    }

    const sheet = getSheet(SHEET_NAME.ATTENDANCE);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim().replace(/\s+/g, "_"));
    const col = {};
    headers.forEach((h, i) => { col[h] = i; });

    // Pastikan kolom time_in dan time_out ada
    const ensureCol = (colName) => {
      if (col[colName] !== undefined) return;
      const newColNum = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColNum).setValue(colName);
      col[colName] = newColNum - 1;
    };
    ensureCol("time_in");
    ensureCol("time_out");

    // Cari baris existing: dahulukan attId jika ada
    let targetRowNum = -1;
    if (attId) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][col["id"] !== undefined ? col["id"] : 0]) === attId) {
          targetRowNum = i + 1;
          break;
        }
      }
    }
    // Fallback: cari berdasarkan user_id + date + schedule kosong/PICKET
    if (targetRowNum === -1) {
      for (let i = 1; i < rows.length; i++) {
        const rowUserId = String(rows[i][col["user_id"]] || "");
        const rowDate   = safeDate(rows[i][col["date"]]);
        const rowSched  = col["schedule_id"] !== undefined ? String(rows[i][col["schedule_id"]] || "") : "";
        if (rowUserId === targetUserId && rowDate === todayStr &&
            (rowSched === "" || rowSched === "PICKET-DUTY" || rowSched === "PIKET")) {
          targetRowNum = i + 1;
          break;
        }
      }
    }

    if (targetRowNum !== -1) {
      // Update baris yang ada
      if (timeIn  !== "") sheet.getRange(targetRowNum, col["time_in"]  + 1).setValue(timeIn);
      if (timeOut !== "") sheet.getRange(targetRowNum, col["time_out"] + 1).setValue(timeOut);
      try { _invalidateDataSnapshot(); } catch (_) {}
      return { status: "success", message: "Jam berhasil diperbarui.", att_id: String(rows[targetRowNum - 1][col["id"] !== undefined ? col["id"] : 0]) };
    } else {
      // Buat baris baru.
      // Hitung panjang baris SETELAH ensureCol menambah kolom baru, agar
      // indeks col["time_in"] / col["time_out"] tidak melebihi panjang array.
      const newId  = generateId("ATT");
      const maxIdx = Object.values(col).reduce(function(a, b) { return Math.max(a, b); }, 0);
      const rowLen = maxIdx + 1;
      const newRow = new Array(rowLen).fill("");
      if (col["id"]          !== undefined) newRow[col["id"]]          = newId;
      if (col["date"]        !== undefined) newRow[col["date"]]        = todayStr;
      if (col["user_id"]     !== undefined) newRow[col["user_id"]]     = targetUserId;
      if (col["status"]      !== undefined) newRow[col["status"]]      = "Present";
      if (col["confirmed_at"]!== undefined) newRow[col["confirmed_at"]]= new Date();
      if (col["schedule_id"] !== undefined) newRow[col["schedule_id"]] = "";
      if (col["time_in"]     !== undefined) newRow[col["time_in"]]     = timeIn;
      if (col["time_out"]    !== undefined) newRow[col["time_out"]]    = timeOut;
      sheet.appendRow(newRow);
      try { _invalidateDataSnapshot(); } catch (_) {}
      return { status: "success", message: "Data jam berhasil disimpan.", att_id: newId };
    }
  } catch (e) {
    return { status: "error", message: "Server Error (PicketTime): " + e.toString() };
  }
}

function getGuruTodayScheduleProgress(token) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error" };
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    const dayIndex = _getIndoDayIndex();
    const allSchedules = getData("Schedules");
    const allLogs = getData("Teaching_Logs");
    const allUsers = getData("Users");
    const cfg = _getConfigMap();
    const activeTP = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    const mySchedules = allSchedules.filter((s) => {
      const isMe = String(s.user_id).trim() === String(user.id).trim();
      const isToday = String(s.day_index).trim() === String(dayIndex);
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return isMe && isToday && sTP === activeTP && sSem === activeSem;
    });
    const todaySubs = getData("Substitutes").filter((s) => {
      try {
        return (
          Utilities.formatDate(
            new Date(s.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          ) === todayStr
        );
      } catch (e) {
        return false;
      }
    });
    const replacedSchedIds = new Set(
      todaySubs
        .filter(
          (sub) =>
            String(sub.original_user_id).trim() === String(user.id).trim(),
        )
        .map((sub) => String(sub.schedule_id)),
    );
    const mySubEntries = todaySubs.filter(
      (sub) => String(sub.substitute_user_id).trim() === String(user.id).trim(),
    );
    const progress = [];
    mySchedules.forEach((s) => {
      let tStart = s.time_start;
      if (tStart instanceof Date)
        tStart = Utilities.formatDate(tStart, "Asia/Jakarta", "HH:mm");
      tStart = String(tStart).substring(0, 5);
      if (replacedSchedIds.has(String(s.id))) {
        const subEntry = todaySubs.find(
          (sub) =>
            String(sub.schedule_id) === String(s.id) &&
            String(sub.original_user_id).trim() === String(user.id).trim(),
        );
        const subUser = subEntry
          ? allUsers.find(
              (u) => String(u.id) === String(subEntry.substitute_user_id),
            )
          : null;
        progress.push({
          schedule_id: s.id,
          subject: s.subject,
          class_name: s.class_name,
          time_start: tStart,
          jtm_val: s.jtm_val,
          has_journal: true,
          is_replaced: true,
          substitute_name: subUser
            ? String(subUser.full_name)
            : "Guru Pengganti",
        });
      } else {
        const hasDone = allLogs.some((l) => {
          const ld = Utilities.formatDate(
            new Date(l.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          );
          return (
            ld === todayStr &&
            String(l.schedule_id) === String(s.id) &&
            String(l.user_id) === String(user.id)
          );
        });
        progress.push({
          schedule_id: s.id,
          subject: s.subject,
          class_name: s.class_name,
          time_start: tStart,
          jtm_val: s.jtm_val,
          has_journal: hasDone,
          is_replaced: false,
        });
      }
    });
    mySubEntries.forEach((sub) => {
      const origSched = allSchedules.find(
        (s) => String(s.id) === String(sub.schedule_id),
      );
      if (!origSched) return;
      if (
        progress.some(
          (p) =>
            String(p.schedule_id) === String(sub.schedule_id) && !p.is_replaced,
        )
      )
        return;
      let tStart = origSched.time_start;
      if (tStart instanceof Date)
        tStart = Utilities.formatDate(tStart, "Asia/Jakarta", "HH:mm");
      tStart = String(tStart).substring(0, 5);
      const hasDone = allLogs.some((l) => {
        const ld = Utilities.formatDate(
          new Date(l.date),
          "Asia/Jakarta",
          "yyyy-MM-dd",
        );
        return (
          ld === todayStr &&
          String(l.schedule_id) === String(sub.schedule_id) &&
          String(l.user_id) === String(user.id)
        );
      });
      const origUser = allUsers.find(
        (u) => String(u.id) === String(sub.original_user_id),
      );
      progress.push({
        schedule_id: sub.schedule_id,
        subject: origSched.subject,
        class_name: origSched.class_name,
        time_start: tStart,
        jtm_val: origSched.jtm_val,
        has_journal: hasDone,
        is_substitute: true,
        is_replaced: false,
        original_teacher: origUser ? String(origUser.full_name) : "Guru Asli",
      });
    });
    const myEventsToday = getData("Event_Attendance").filter((a) => {
      try {
        return (
          String(a.user_id).trim() === String(user.id).trim() &&
          Utilities.formatDate(
            new Date(a.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          ) === todayStr
        );
      } catch (e) {
        return false;
      }
    });
    myEventsToday.forEach((e) => {
      const hasDone = String(e.journal_submitted).toLowerCase() === "true";
      progress.push({
        schedule_id: "EVENT-" + e.id,
        subject: e.event_name || "Acara Sekolah",
        class_name: "Kegiatan/Acara",
        time_start: "Acara",
        jtm_val: e.jtm_val || 0,
        has_journal: hasDone,
        is_replaced: false,
      });
    });
    progress.sort((a, b) =>
      (a.time_start || "").localeCompare(b.time_start || ""),
    );
    const activeItems = progress.filter((p) => !p.is_replaced);
    const done = activeItems.filter((p) => p.has_journal).length;
    const total = activeItems.length;
    return {
      status: "success",
      items: progress,
      done: done,
      total: total,
      pct: total > 0 ? Math.round((done / total) * 100) : 0,
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}
function _examNow_() {
  return Utilities.formatDate(
    new Date(),
    "Asia/Jakarta",
    "yyyy-MM-dd'T'HH:mm:ss",
  );
}
function _examToday_() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
}
function _examGetSession_(session_id) {
  return (
    getData(EXAM_SHEET.SESSIONS).find(
      (s) => String(s.id) === String(session_id),
    ) || null
  );
}
function _examGetRoom_(room_id) {
  return (
    getData(EXAM_SHEET.ROOMS).find((r) => String(r.id) === String(room_id)) ||
    null
  );
}
function _examGetRoomsBySession_(session_id) {
  return getData(EXAM_SHEET.ROOMS).filter(
    (r) => String(r.session_id) === String(session_id),
  );
}
function _examGetSessionsByPeriod_(period_id) {
  return getData(EXAM_SHEET.SESSIONS).filter(
    (s) => String(s.period_id) === String(period_id),
  );
}
function _examDeleteBySet_(sheetName, colIndex, valueSet) {
  if (!valueSet || valueSet.size === 0) return;
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (valueSet.has(String(data[i][colIndex]))) {
      sheet.deleteRow(i + 1);
    }
  }
}
function _examDeleteById_(sheetName, id) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
function _examFormatTime_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Jakarta", "HH:mm");
  }
  return String(val || "").substring(0, 5);
}
function _examCheckSupervisorConflict_(user_id, session_id, exclude_sup_id) {
  const session = _examGetSession_(session_id);
  if (!session) return "Sesi ujian tidak ditemukan.";
  const targetDate = String(session.date);
  const sessionRoomIds = new Set(
    _examGetRoomsBySession_(session_id).map((r) => String(r.id)),
  );
  const conflictSameSesi = getData(EXAM_SHEET.SUPERVISORS).some(
    (sup) =>
      String(sup.user_id) === String(user_id) &&
      String(sup.status) === "active" &&
      sessionRoomIds.has(String(sup.room_id)) &&
      (!exclude_sup_id || String(sup.id) !== String(exclude_sup_id)),
  );
  if (conflictSameSesi) {
    return "Guru ini sudah menjadi pengawas di ruang lain pada sesi yang sama.";
  }
  return null;
}
function _examCheckCommitteeConflict_(user_id, date, exclude_com_id) {
  const targetDate = String(date);
  const alreadyCommittee = getData(EXAM_SHEET.COMMITTEE).some(
    (c) =>
      String(c.user_id) === String(user_id) &&
      String(c.date) === targetDate &&
      String(c.status) === "active" &&
      (!exclude_com_id || String(c.id) !== String(exclude_com_id)),
  );
  if (alreadyCommittee) {
    return "Guru ini sudah terdaftar sebagai Panitia pada tanggal tersebut.";
  }
  return null;
}
function getExamPeriods(token) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    const today = _examToday_();
    const data = getData(EXAM_SHEET.PERIODS)
      .map((p) => {
        const ds = String(p.date_start);
        const de = String(p.date_end);
        let periodStatus;
        if (today < ds) periodStatus = "Akan Datang";
        else if (today > de) periodStatus = "Selesai";
        else periodStatus = "Berlangsung";
        return {
          id: String(p.id),
          name: String(p.name || ""),
          date_start: ds,
          date_end: de,
          jtm_committee_per_day: Number(p.jtm_committee_per_day) || 0,
          description: String(p.description || ""),
          created_by: String(p.created_by || ""),
          created_at: String(p.created_at || ""),
          status: periodStatus,
        };
      })
      .sort((a, b) => b.date_start.localeCompare(a.date_start));
    return { status: "success", data };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function addExamPeriod(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat membuat periode ujian.",
      };
    }
    if (!data.name || !String(data.name).trim()) {
      return { status: "error", message: "Nama periode ujian wajib diisi." };
    }
    if (!data.date_start) {
      return { status: "error", message: "Tanggal mulai wajib diisi." };
    }
    if (!data.date_end) {
      return { status: "error", message: "Tanggal selesai wajib diisi." };
    }
    if (String(data.date_start) > String(data.date_end)) {
      return {
        status: "error",
        message: "Tanggal mulai tidak boleh setelah tanggal selesai.",
      };
    }
    const jtm = Number(data.jtm_committee_per_day);
    if (isNaN(jtm) || jtm < 0) {
      return {
        status: "error",
        message: "Nilai JTM Panitia per hari harus berupa angka (minimal 0).",
      };
    }
    const newId = "EXP-" + new Date().getTime();
    getSheet(EXAM_SHEET.PERIODS).appendRow([
      newId,
      String(data.name).trim(),
      String(data.date_start),
      String(data.date_end),
      jtm,
      String(data.description || "").trim(),
      String(user.id),
      _examNow_(),
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Periode ujian berhasil dibuat.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function editExamPeriod(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat mengedit periode ujian.",
      };
    }
    if (!data.id)
      return { status: "error", message: "ID periode tidak ditemukan." };
    if (!data.name || !String(data.name).trim()) {
      return { status: "error", message: "Nama periode wajib diisi." };
    }
    if (!data.date_start || !data.date_end) {
      return {
        status: "error",
        message: "Tanggal mulai dan selesai wajib diisi.",
      };
    }
    if (String(data.date_start) > String(data.date_end)) {
      return {
        status: "error",
        message: "Tanggal mulai tidak boleh setelah tanggal selesai.",
      };
    }
    const jtm = Number(data.jtm_committee_per_day);
    if (isNaN(jtm) || jtm < 0) {
      return {
        status: "error",
        message: "Nilai JTM Panitia per hari harus berupa angka (minimal 0).",
      };
    }
    const existing = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(data.id),
    );
    if (!existing)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    if (_examToday_() >= String(existing.date_start)) {
      return {
        status: "error",
        message: "Periode yang sudah dimulai atau selesai tidak dapat diedit.",
      };
    }
    const sheet = getSheet(EXAM_SHEET.PERIODS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 2).setValue(String(data.name).trim());
        sheet.getRange(i + 1, 3).setValue(String(data.date_start));
        sheet.getRange(i + 1, 4).setValue(String(data.date_end));
        sheet.getRange(i + 1, 5).setValue(jtm);
        sheet
          .getRange(i + 1, 6)
          .setValue(String(data.description || "").trim());
        return {
          status: "success",
          message: "Periode ujian berhasil diperbarui.",
        };
      }
    }
    return { status: "error", message: "Periode ujian tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function deleteExamPeriod(token, id, force) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat menghapus periode ujian.",
      };
    }
    if (!id) return { status: "error", message: "ID periode tidak ditemukan." };
    const existing = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(id),
    );
    if (!existing)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const sessions = _examGetSessionsByPeriod_(id);
    const sessionIds = new Set(sessions.map((s) => String(s.id)));
    const allRooms = getData(EXAM_SHEET.ROOMS);
    const rooms = allRooms.filter((r) => sessionIds.has(String(r.session_id)));
    const roomIds = new Set(rooms.map((r) => String(r.id)));
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const supervisors = allSups.filter((sup) =>
      roomIds.has(String(sup.room_id)),
    );
    const supIds = new Set(supervisors.map((s) => String(s.id)));
    const allCom = getData(EXAM_SHEET.COMMITTEE);
    const committee = allCom.filter((c) => String(c.period_id) === String(id));
    if (!force) {
      const hasConfirmedSup = supervisors.some(
        (sup) => sup.confirmed_by && String(sup.confirmed_by).trim() !== "",
      );
      const hasConfirmedCom = committee.some(
        (c) => c.confirmed_by && String(c.confirmed_by).trim() !== "",
      );
      if (hasConfirmedSup || hasConfirmedCom) {
        return {
          status: "confirm_required",
          message:
            "Periode ini sudah memiliki data konfirmasi dan JTM terhitung. " +
            "Penghapusan akan menghapus semua data terkait secara permanen. Lanjutkan?",
        };
      }
    } else {
      supIds.forEach(function (supId) {
        _examDeleteLogByRef_(supId);
      });
      committee.forEach(function (c) {
        _examDeleteLogByRef_(String(c.id));
      });
    }
    _examDeleteBySet_(EXAM_SHEET.BAP, 1, supIds);
    _examDeleteBySet_(EXAM_SHEET.SUPERVISORS, 1, roomIds);
    _examDeleteBySet_(EXAM_SHEET.ROOMS, 1, sessionIds);
    const sesSheet = getSheet(EXAM_SHEET.SESSIONS);
    const sesData = sesSheet.getDataRange().getValues();
    for (let i = sesData.length - 1; i >= 1; i--) {
      if (String(sesData[i][1]) === String(id)) sesSheet.deleteRow(i + 1);
    }
    const comSheet = getSheet(EXAM_SHEET.COMMITTEE);
    const comData = comSheet.getDataRange().getValues();
    for (let i = comData.length - 1; i >= 1; i--) {
      if (String(comData[i][1]) === String(id)) comSheet.deleteRow(i + 1);
    }
    _examDeleteById_(EXAM_SHEET.PERIODS, id);
    return {
      status: "success",
      message: "Periode ujian dan semua data terkait berhasil dihapus.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function addExamSession(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menambah sesi ujian.",
      };
    }
    if (!data.period_id)
      return { status: "error", message: "ID periode wajib diisi." };
    if (!data.date)
      return { status: "error", message: "Tanggal sesi wajib diisi." };
    if (!data.session_name || !String(data.session_name).trim()) {
      return { status: "error", message: "Nama sesi wajib diisi." };
    }
    if (!data.time_start)
      return { status: "error", message: "Jam mulai sesi wajib diisi." };
    if (!data.time_end)
      return { status: "error", message: "Jam selesai sesi wajib diisi." };
    const ts = String(data.time_start).substring(0, 5);
    const te = String(data.time_end).substring(0, 5);
    if (ts >= te) {
      return {
        status: "error",
        message: "Jam mulai harus sebelum jam selesai.",
      };
    }
    const jtm = Number(data.jtm_val);
    if (isNaN(jtm) || jtm < 0) {
      return {
        status: "error",
        message: "Nilai JTM per sesi harus berupa angka (minimal 0).",
      };
    }
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(data.period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const sesDate = String(data.date);
    if (
      sesDate < String(period.date_start) ||
      sesDate > String(period.date_end)
    ) {
      return {
        status: "error",
        message:
          "Tanggal sesi harus berada dalam rentang periode ujian (" +
          period.date_start +
          " s.d. " +
          period.date_end +
          ").",
      };
    }
    const newId = "EXS-" + new Date().getTime();
    getSheet(EXAM_SHEET.SESSIONS).appendRow([
      newId,
      String(data.period_id),
      sesDate,
      String(data.session_name).trim(),
      ts,
      te,
      jtm,
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Sesi ujian berhasil ditambahkan.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function editExamSession(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat mengedit sesi ujian.",
      };
    }
    if (!data.id)
      return { status: "error", message: "ID sesi tidak ditemukan." };
    if (!data.session_name || !String(data.session_name).trim()) {
      return { status: "error", message: "Nama sesi wajib diisi." };
    }
    if (!data.time_start || !data.time_end) {
      return {
        status: "error",
        message: "Jam mulai dan jam selesai wajib diisi.",
      };
    }
    const ts = String(data.time_start).substring(0, 5);
    const te = String(data.time_end).substring(0, 5);
    if (ts >= te) {
      return {
        status: "error",
        message: "Jam mulai harus sebelum jam selesai.",
      };
    }
    const jtm = Number(data.jtm_val);
    if (isNaN(jtm) || jtm < 0) {
      return {
        status: "error",
        message: "Nilai JTM per sesi harus berupa angka (minimal 0).",
      };
    }
    const roomIds = new Set(
      _examGetRoomsBySession_(data.id).map((r) => String(r.id)),
    );
    const hasConfirmed = getData(EXAM_SHEET.SUPERVISORS).some(
      (sup) =>
        roomIds.has(String(sup.room_id)) &&
        sup.confirmed_by &&
        String(sup.confirmed_by).trim() !== "",
    );
    if (hasConfirmed) {
      return {
        status: "error",
        message:
          "Sesi ini sudah memiliki pengawas yang dikonfirmasi kehadirannya dan tidak dapat diedit.",
      };
    }
    const sheet = getSheet(EXAM_SHEET.SESSIONS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 4).setValue(String(data.session_name).trim());
        sheet.getRange(i + 1, 5).setValue(ts);
        sheet.getRange(i + 1, 6).setValue(te);
        sheet.getRange(i + 1, 7).setValue(jtm);
        return {
          status: "success",
          message: "Sesi ujian berhasil diperbarui.",
        };
      }
    }
    return { status: "error", message: "Sesi ujian tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function deleteExamSession(token, id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menghapus sesi ujian.",
      };
    }
    if (!id) return { status: "error", message: "ID sesi tidak ditemukan." };
    const rooms = _examGetRoomsBySession_(id);
    const roomIds = new Set(rooms.map((r) => String(r.id)));
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const supervisors = allSups.filter((sup) =>
      roomIds.has(String(sup.room_id)),
    );
    const supIds = new Set(supervisors.map((s) => String(s.id)));
    const hasConfirmed = supervisors.some(
      (sup) => sup.confirmed_by && String(sup.confirmed_by).trim() !== "",
    );
    if (hasConfirmed) {
      return {
        status: "error",
        message:
          "Sesi ini sudah memiliki pengawas yang dikonfirmasi kehadirannya dan tidak dapat dihapus.",
      };
    }
    _examDeleteBySet_(EXAM_SHEET.BAP, 1, supIds);
    _examDeleteBySet_(EXAM_SHEET.SUPERVISORS, 1, roomIds);
    _examDeleteBySet_(EXAM_SHEET.ROOMS, 1, new Set([String(id)]));
    _examDeleteById_(EXAM_SHEET.SESSIONS, id);
    return { status: "success", message: "Sesi ujian berhasil dihapus." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function addExamRoom(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menambah ruang ujian.",
      };
    }
    if (!data.session_id)
      return { status: "error", message: "ID sesi wajib diisi." };
    if (!data.room_name || !String(data.room_name).trim())
      return { status: "error", message: "Nama/nomor ruang wajib diisi." };
    if (!data.subject || !String(data.subject).trim())
      return { status: "error", message: "Mata ujian wajib diisi." };
    if (!data.class_name || !String(data.class_name).trim())
      return { status: "error", message: "Nama kelas wajib diisi." };
    const totalSiswa = Number(data.total_siswa) || 0;
    if (totalSiswa < 0) {
      return { status: "error", message: "Total siswa tidak boleh negatif." };
    }
    if (!_examGetSession_(data.session_id)) {
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    }
    const newId = "EXR-" + new Date().getTime();
    getSheet(EXAM_SHEET.ROOMS).appendRow([
      newId,
      String(data.session_id),
      String(data.room_name).trim(),
      String(data.subject).trim(),
      String(data.class_name).trim(),
      totalSiswa,
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Ruang ujian berhasil ditambahkan.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function editExamRoom(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat mengedit ruang ujian.",
      };
    }
    if (!data.id)
      return { status: "error", message: "ID ruang tidak ditemukan." };
    if (!data.room_name || !String(data.room_name).trim())
      return { status: "error", message: "Nama/nomor ruang wajib diisi." };
    if (!data.subject || !String(data.subject).trim())
      return { status: "error", message: "Mata ujian wajib diisi." };
    if (!data.class_name || !String(data.class_name).trim())
      return { status: "error", message: "Nama kelas wajib diisi." };
    const totalSiswa = Number(data.total_siswa) || 0;
    if (totalSiswa < 0) {
      return { status: "error", message: "Total siswa tidak boleh negatif." };
    }
    const hasConfirmed = getData(EXAM_SHEET.SUPERVISORS).some(
      (sup) =>
        String(sup.room_id) === String(data.id) &&
        sup.confirmed_by &&
        String(sup.confirmed_by).trim() !== "",
    );
    if (hasConfirmed) {
      return {
        status: "error",
        message:
          "Ruang ini sudah memiliki pengawas yang dikonfirmasi dan tidak dapat diedit.",
      };
    }
    const sheet = getSheet(EXAM_SHEET.ROOMS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 3).setValue(String(data.room_name).trim());
        sheet.getRange(i + 1, 4).setValue(String(data.subject).trim());
        sheet.getRange(i + 1, 5).setValue(String(data.class_name).trim());
        sheet.getRange(i + 1, 6).setValue(totalSiswa);
        return {
          status: "success",
          message: "Ruang ujian berhasil diperbarui.",
        };
      }
    }
    return { status: "error", message: "Ruang ujian tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function deleteExamRoom(token, id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menghapus ruang ujian.",
      };
    }
    if (!id) return { status: "error", message: "ID ruang tidak ditemukan." };
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const supsInRoom = allSups.filter(
      (sup) => String(sup.room_id) === String(id),
    );
    const supIds = new Set(supsInRoom.map((s) => String(s.id)));
    const hasConfirmed = supsInRoom.some(
      (sup) => sup.confirmed_by && String(sup.confirmed_by).trim() !== "",
    );
    if (hasConfirmed) {
      return {
        status: "error",
        message:
          "Ruang ini sudah memiliki pengawas yang dikonfirmasi dan tidak dapat dihapus.",
      };
    }
    _examDeleteBySet_(EXAM_SHEET.BAP, 1, supIds);
    _examDeleteBySet_(EXAM_SHEET.SUPERVISORS, 1, new Set([String(id)]));
    _examDeleteById_(EXAM_SHEET.ROOMS, id);
    return { status: "success", message: "Ruang ujian berhasil dihapus." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function assignExamSupervisor(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menetapkan pengawas.",
      };
    }
    if (!data.room_id)
      return { status: "error", message: "ID ruang wajib diisi." };
    if (!data.user_id)
      return { status: "error", message: "ID guru wajib diisi." };
    const room = _examGetRoom_(data.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const targetGuru = getData("Users").find(
      (u) => String(u.id) === String(data.user_id),
    );
    if (!targetGuru)
      return { status: "error", message: "Guru tidak ditemukan." };
    const conflict = _examCheckSupervisorConflict_(
      data.user_id,
      room.session_id,
    );
    if (conflict) return { status: "error", message: conflict };
    const session = _examGetSession_(room.session_id);
    let jtmVal = Number(data.jtm_val);
    if (isNaN(jtmVal) || jtmVal < 0) {
      jtmVal = session ? Number(session.jtm_val) || 0 : 0;
    }
    const newId = "EXSUP-" + new Date().getTime();
    getSheet(EXAM_SHEET.SUPERVISORS).appendRow([
      newId,
      String(data.room_id),
      String(data.user_id),
      "active",
      false,
      "",
      "",
      "",
      "",
      "",
      jtmVal,
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Pengawas berhasil ditetapkan.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function editExamSupervisor(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat mengedit pengawas.",
      };
    }
    if (!data || !data.id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    if (!data.user_id)
      return { status: "error", message: "ID guru wajib diisi." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(data.id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message:
          "Pengawas yang sudah dikonfirmasi kehadirannya tidak dapat diedit. " +
          "Cabut konfirmasi kehadiran terlebih dahulu.",
      };
    }
    var isSub =
      target.is_substitute === true ||
      String(target.is_substitute).toLowerCase() === "true";
    if (isSub) {
      return {
        status: "error",
        message:
          "Pengawas pengganti tidak dapat diedit langsung. Gunakan tombol Batal Ganti terlebih dahulu.",
      };
    }
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return {
        status: "error",
        message: "Ruang ujian tidak ditemukan untuk pengawas ini.",
      };
    const targetGuru = getData("Users").find(
      (u) => String(u.id) === String(data.user_id),
    );
    if (!targetGuru)
      return { status: "error", message: "Guru tidak ditemukan." };
    var userChanged = String(target.user_id) !== String(data.user_id);
    if (userChanged) {
      const conflict = _examCheckSupervisorConflict_(
        data.user_id,
        room.session_id,
        data.id,
      );
      if (conflict) return { status: "error", message: conflict };
    }
    const session = _examGetSession_(room.session_id);
    let jtmVal = Number(data.jtm_val);
    if (isNaN(jtmVal) || jtmVal < 0) {
      jtmVal = session
        ? Number(session.jtm_val) || 0
        : Number(target.jtm_val) || 0;
    }
    const sheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const rows = sheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 3).setValue(String(data.user_id));
        sheet.getRange(i + 1, 11).setValue(jtmVal);
        updated = true;
        break;
      }
    }
    if (!updated)
      return {
        status: "error",
        message: "Gagal memperbarui data pengawas (baris tidak ditemukan).",
      };
    SpreadsheetApp.flush();
    return { status: "success", message: "Data pengawas berhasil diperbarui." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function removeExamSupervisor(token, supervisor_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menghapus pengawas.",
      };
    }
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message:
          "Pengawas yang sudah dikonfirmasi kehadirannya tidak dapat dihapus langsung. " +
          "Cabut konfirmasi kehadiran terlebih dahulu.",
      };
    }
    _examDeleteBySet_(EXAM_SHEET.BAP, 1, new Set([String(supervisor_id)]));
    _examDeleteById_(EXAM_SHEET.SUPERVISORS, supervisor_id);
    return {
      status: "success",
      message: "Pengawas berhasil dihapus dari jadwal.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function assignExamCommittee(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menetapkan panitia.",
      };
    }
    if (!data.period_id)
      return { status: "error", message: "ID periode wajib diisi." };
    if (!data.date) return { status: "error", message: "Tanggal wajib diisi." };
    if (!data.user_id)
      return { status: "error", message: "ID guru wajib diisi." };
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(data.period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const comDate = String(data.date);
    if (
      comDate < String(period.date_start) ||
      comDate > String(period.date_end)
    ) {
      return {
        status: "error",
        message:
          "Tanggal panitia harus berada dalam rentang periode ujian (" +
          period.date_start +
          " s.d. " +
          period.date_end +
          ").",
      };
    }
    const targetGuru = getData("Users").find(
      (u) => String(u.id) === String(data.user_id),
    );
    if (!targetGuru)
      return { status: "error", message: "Guru tidak ditemukan." };
    const conflict = _examCheckCommitteeConflict_(data.user_id, comDate);
    if (conflict) return { status: "error", message: conflict };
    const jtmVal = Number(period.jtm_committee_per_day) || 0;
    const newId = "EXCOM-" + new Date().getTime();
    getSheet(EXAM_SHEET.COMMITTEE).appendRow([
      newId,
      String(data.period_id),
      comDate,
      String(data.user_id),
      jtmVal,
      "active",
      false,
      "",
      "",
      "",
      "",
      "",
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Panitia berhasil ditetapkan.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function removeExamCommittee(token, committee_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menghapus panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message:
          "Panitia yang sudah dikonfirmasi kehadirannya tidak dapat dihapus langsung. " +
          "Cabut konfirmasi kehadiran terlebih dahulu.",
      };
    }
    _examDeleteById_(EXAM_SHEET.COMMITTEE, committee_id);
    return {
      status: "success",
      message: "Panitia berhasil dihapus dari jadwal.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function getExamAdminMasterData(token) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return { status: "error", message: "Akses ditolak." };
    }
    const periodsResult = getExamPeriods(token);
    const periods =
      periodsResult.status === "success" ? periodsResult.data : [];
    const teachers = getData("Users")
      .filter((u) => String(u.role).toLowerCase() !== "admin")
      .map((u) => ({
        id: String(u.id),
        name: String(u.full_name || ""),
        nip: String(u.nip || ""),
        username: String(u.username || ""),
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
    return { status: "success", periods, teachers };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function getExamPeriodDetail(token, period_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return { status: "error", message: "Akses ditolak." };
    }
    if (!period_id)
      return { status: "error", message: "ID periode wajib diisi." };
    const periods = getData(EXAM_SHEET.PERIODS);
    const period = periods.find((p) => String(p.id) === String(period_id));
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const today = _examToday_();
    const ds = String(period.date_start);
    const de = String(period.date_end);
    const periodStatus =
      today < ds ? "Akan Datang" : today > de ? "Selesai" : "Berlangsung";
    const periodObj = {
      id: String(period.id),
      name: String(period.name || ""),
      date_start: ds,
      date_end: de,
      jtm_committee_per_day: Number(period.jtm_committee_per_day) || 0,
      description: String(period.description || ""),
      status: periodStatus,
    };
    const allUsers = getData("Users");
    const allSessions = getData(EXAM_SHEET.SESSIONS).filter(
      (s) => String(s.period_id) === String(period_id),
    );
    const allRooms = getData(EXAM_SHEET.ROOMS);
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const allBaps = getData(EXAM_SHEET.BAP);
    const sessions = allSessions
      .sort((a, b) => {
        const da = String(a.date),
          db = String(b.date);
        if (da !== db) return da.localeCompare(db);
        return _examFormatTime_(a.time_start).localeCompare(
          _examFormatTime_(b.time_start),
        );
      })
      .map((s) => {
        const rooms = allRooms
          .filter((r) => String(r.session_id) === String(s.id))
          .map((r) => {
            const supervisors = allSups
              .filter((sup) => String(sup.room_id) === String(r.id))
              .map((sup) => {
                const g = allUsers.find(
                  (u) => String(u.id) === String(sup.user_id),
                );
                const bapEntry = allBaps.find(
                  (b) =>
                    String(b.supervisor_id) === String(sup.id) &&
                    b.submitted_at &&
                    String(b.submitted_at).trim() !== "",
                );
                return {
                  id: String(sup.id),
                  user_id: String(sup.user_id),
                  guru_name: g ? String(g.full_name) : "Tidak Dikenal",
                  status: String(sup.status || "active"),
                  is_substitute:
                    sup.is_substitute === true ||
                    String(sup.is_substitute) === "true",
                  original_user_id: String(sup.original_user_id || ""),
                  confirmed_by: String(sup.confirmed_by || ""),
                  confirmed_at: String(sup.confirmed_at || ""),
                  substitution_cancelled_by: String(
                    sup.substitution_cancelled_by || "",
                  ),
                  substitution_cancelled_at: String(
                    sup.substitution_cancelled_at || "",
                  ),
                  jtm_val: Number(sup.jtm_val) || 0,
                  has_bap: !!bapEntry,
                };
              });
            return {
              id: String(r.id),
              room_name: String(r.room_name || ""),
              subject: String(r.subject || ""),
              class_name: String(r.class_name || ""),
              total_siswa: Number(r.total_siswa) || 0,
              supervisors,
            };
          })
          .sort((a, b) => a.room_name.localeCompare(b.room_name));
        return {
          id: String(s.id),
          date: String(s.date),
          session_name: String(s.session_name || ""),
          time_start: _examFormatTime_(s.time_start),
          time_end: _examFormatTime_(s.time_end),
          jtm_val: Number(s.jtm_val) || 0,
          rooms,
        };
      });
    const committee = getData(EXAM_SHEET.COMMITTEE)
      .filter((c) => String(c.period_id) === String(period_id))
      .map((c) => {
        const g = allUsers.find((u) => String(u.id) === String(c.user_id));
        return {
          id: String(c.id),
          date: String(c.date),
          user_id: String(c.user_id),
          guru_name: g ? String(g.full_name) : "Tidak Dikenal",
          jtm_val: Number(c.jtm_val) || 0,
          status: String(c.status || "active"),
          is_substitute:
            c.is_substitute === true || String(c.is_substitute) === "true",
          original_user_id: String(c.original_user_id || ""),
          confirmed_by: String(c.confirmed_by || ""),
          confirmed_at: String(c.confirmed_at || ""),
          substitution_cancelled_by: String(c.substitution_cancelled_by || ""),
          substitution_cancelled_at: String(c.substitution_cancelled_at || ""),
        };
      })
      .sort((a, b) => a.date.localeCompare(b.date));
    return { status: "success", period: periodObj, sessions, committee };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function _examAddJtmLog_(
  scheduleId,
  userId,
  date,
  materi,
  hadir,
  absen,
  jtmVal,
  refId,
) {
  const logId = "LOG-EXAM-" + new Date().getTime();
  getSheet(SHEET_NAME.LOGS).appendRow([
    logId,
    String(scheduleId),
    String(userId),
    String(date),
    String(materi),
    Number(hadir) || 0,
    Number(absen) || 0,
    "REF:" + String(refId),
    Number(jtmVal) || 0,
    _examNow_(),
  ]);
  return logId;
}
function _examDeleteLogByRef_(refId) {
  const sheet = getSheet(SHEET_NAME.LOGS);
  const rows = sheet.getDataRange().getValues();
  const REF = "REF:" + String(refId);
  let deleted = false;
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][7]).indexOf(REF) !== -1) {
      sheet.deleteRow(i + 1);
      deleted = true;
    }
  }
  return deleted;
}
function _examUpdateLogByRef_(refId, hadir, absen) {
  const sheet = getSheet(SHEET_NAME.LOGS);
  const rows = sheet.getDataRange().getValues();
  const REF = "REF:" + String(refId);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][7]).indexOf(REF) !== -1) {
      sheet.getRange(i + 1, 6).setValue(Number(hadir) || 0);
      sheet.getRange(i + 1, 7).setValue(Number(absen) || 0);
      return true;
    }
  }
  return false;
}
function _examGetBap_(supervisor_id) {
  return (
    getData(EXAM_SHEET.BAP).find(
      (b) => String(b.supervisor_id) === String(supervisor_id),
    ) || null
  );
}
function _examIsConfirmedPanitia_(user_id, date) {
  return getData(EXAM_SHEET.COMMITTEE).some(
    (c) =>
      String(c.user_id) === String(user_id) &&
      String(c.date) === String(date) &&
      String(c.status) === "active" &&
      c.confirmed_by &&
      String(c.confirmed_by).trim() !== "",
  );
}
function _examCheckAdminOrConfirmedPanitia_(user, date) {
  if (String(user.role).toLowerCase() === "admin") {
    return { ok: true, isAdmin: true };
  }
  if (_examIsConfirmedPanitia_(user.id, date)) {
    return { ok: true, isAdmin: false };
  }
  return {
    ok: false,
    isAdmin: false,
    error:
      "Akses ditolak. Hanya Admin atau Panitia yang sudah dikonfirmasi yang dapat melakukan aksi ini pada hari H.",
  };
}
function _examCheckBapTiming_(sessionDate, timeEnd) {
  const today = _examToday_();
  if (today < sessionDate) {
    return "Sesi ujian belum berlangsung.";
  }
  if (today > sessionDate) {
    return null;
  }
  const [teH, teM] = String(timeEnd).split(":").map(Number);
  const activeTotal = teH * 60 + teM + 5;
  const activeTime =
    String(Math.floor(activeTotal / 60)).padStart(2, "0") +
    ":" +
    String(activeTotal % 60).padStart(2, "0");
  const nowTime = Utilities.formatDate(new Date(), "Asia/Jakarta", "HH:mm");
  if (nowTime < activeTime) {
    return (
      "BAP belum bisa diisi. Sesi berakhir pukul " +
      timeEnd +
      ", tombol aktif mulai pukul " +
      activeTime +
      "."
    );
  }
  return null;
}
function confirmCommitteeAttendance(token, committee_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat mengonfirmasi kehadiran Panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    if (_examToday_() !== String(target.date)) {
      return {
        status: "error",
        message:
          "Konfirmasi hanya dapat dilakukan pada hari pelaksanaan (hari H).",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord panitia ini tidak aktif." };
    }
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message: "Panitia ini sudah dikonfirmasi kehadirannya.",
      };
    }
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(target.period_id),
    );
    const periodName = period ? String(period.name) : "Ujian";
    const now = _examNow_();
    const sheet = getSheet(EXAM_SHEET.COMMITTEE);
    const rows = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(committee_id)) {
        sheet.getRange(i + 1, 9).setValue(String(user.id));
        sheet.getRange(i + 1, 10).setValue(now);
        found = true;
        break;
      }
    }
    if (!found)
      return { status: "error", message: "Gagal memperbarui data konfirmasi." };
    _examAddJtmLog_(
      "EXAM-COMMITTEE",
      target.user_id,
      String(target.date),
      "Panitia Ujian: " + periodName,
      0,
      0,
      Number(target.jtm_val) || 0,
      committee_id,
    );
    try {
      _notifCommitteeConfirmed_(target.user_id, target.date);
    } catch (_) {}
    return {
      status: "success",
      message: "Kehadiran Panitia berhasil dikonfirmasi. JTM telah dicatat.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function revokeCommitteeAttendance(token, committee_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat mencabut konfirmasi Panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    if (_examToday_() !== String(target.date)) {
      return {
        status: "error",
        message:
          "Konfirmasi hanya dapat dicabut pada hari pelaksanaan (hari H).",
      };
    }
    if (!target.confirmed_by || String(target.confirmed_by).trim() === "") {
      return {
        status: "error",
        message: "Panitia ini belum dikonfirmasi kehadirannya.",
      };
    }
    _examDeleteLogByRef_(committee_id);
    const sheet = getSheet(EXAM_SHEET.COMMITTEE);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(committee_id)) {
        sheet.getRange(i + 1, 9).setValue("");
        sheet.getRange(i + 1, 10).setValue("");
        break;
      }
    }
    try {
      _notifCommitteeRevoked_(target.user_id, target.date);
    } catch (_) {}
    return {
      status: "success",
      message:
        "Konfirmasi kehadiran Panitia berhasil dicabut. JTM telah dihapus dari honorarium.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function confirmSupervisorAttendance(token, supervisor_id) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const sessionDate = String(session.date);
    if (_examToday_() !== sessionDate) {
      return {
        status: "error",
        message:
          "Konfirmasi hanya dapat dilakukan pada hari pelaksanaan (hari H).",
      };
    }
    const auth = _examCheckAdminOrConfirmedPanitia_(user, sessionDate);
    if (!auth.ok) return { status: "error", message: auth.error };
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord pengawas ini tidak aktif." };
    }
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message: "Pengawas ini sudah dikonfirmasi kehadirannya.",
      };
    }
    const sheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(supervisor_id)) {
        sheet.getRange(i + 1, 7).setValue(String(user.id));
        sheet.getRange(i + 1, 8).setValue(_examNow_());
        try {
          var byName = String(
            user.full_name ||
              user.name ||
              user.username ||
              (auth.isAdmin ? "Admin" : "Panitia Ujian"),
          );
          var sessionInfo =
            String(session.session_name || "") +
            " · " +
            String(room.room_name || "") +
            " (" +
            String(room.subject || "") +
            ")";
          _notifSupervisorConfirmed_(
            target.user_id,
            sessionDate,
            byName,
            sessionInfo,
          );
        } catch (_) {}
        return {
          status: "success",
          message:
            "Kehadiran Pengawas berhasil dikonfirmasi. Pengawas dapat mengisi BAP setelah sesi berakhir + 5 menit.",
        };
      }
    }
    return { status: "error", message: "Gagal memperbarui data konfirmasi." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function revokeSupervisorAttendance(token, supervisor_id, force) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const sessionDate = String(session.date);
    if (_examToday_() !== sessionDate) {
      return {
        status: "error",
        message:
          "Konfirmasi hanya dapat dicabut pada hari pelaksanaan (hari H).",
      };
    }
    const auth = _examCheckAdminOrConfirmedPanitia_(user, sessionDate);
    if (!auth.ok) return { status: "error", message: auth.error };
    if (!target.confirmed_by || String(target.confirmed_by).trim() === "") {
      return {
        status: "error",
        code: JTM_ERROR_CODES.NO_ACTIVE_CONFIRMATION,
        message: "Pengawas ini belum dikonfirmasi kehadirannya.",
      };
    }
    const bap = _examGetBap_(supervisor_id);
    const hasBap = !!(
      bap &&
      bap.submitted_at &&
      String(bap.submitted_at).trim() !== ""
    );
    if (hasBap && !force) {
      return {
        status: "confirm_required",
        message:
          "Pencabutan konfirmasi akan mengunci BAP Pengawas dan menghapus JTM-nya dari honorarium. Lanjutkan?",
      };
    }
    if (hasBap) {
      _examDeleteLogByRef_(supervisor_id);
    }
    const sheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(supervisor_id)) {
        sheet.getRange(i + 1, 7).setValue("");
        sheet.getRange(i + 1, 8).setValue("");
        break;
      }
    }
    try {
      _jtmReverseOccurrence_("EXAM", supervisor_id, sessionDate);
    } catch (_) {}
    try {
      var byName = String(
        user.full_name ||
          user.name ||
          user.username ||
          (auth.isAdmin ? "Admin" : "Panitia Ujian"),
      );
      var sessionInfo =
        String(session.session_name || "") +
        " · " +
        String(room.room_name || "") +
        " (" +
        String(room.subject || "") +
        ")";
      _notifSupervisorRevoked_(
        target.user_id,
        sessionDate,
        byName,
        sessionInfo,
      );
    } catch (_) {}
    return {
      status: "success",
      message: hasBap
        ? "Konfirmasi dicabut. BAP terkunci dan JTM telah dihapus dari honorarium."
        : "Konfirmasi kehadiran Pengawas berhasil dicabut.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function submitExamBAP(token, supervisor_id, bapData) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (String(target.user_id) !== String(user.id)) {
      return {
        status: "error",
        message: "Anda tidak terdaftar sebagai pengawas pada jadwal ini.",
      };
    }
    if (String(target.status) === "cancelled") {
      return {
        status: "error",
        message: "Substitusi Anda pada jadwal ini telah dibatalkan.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord pengawas ini tidak aktif." };
    }
    if (!target.confirmed_by || String(target.confirmed_by).trim() === "") {
      return {
        status: "error",
        message:
          "Kehadiran Anda belum dikonfirmasi Panitia. BAP belum dapat diisi.",
      };
    }
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(session.period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    if (_examToday_() > String(period.date_end)) {
      return {
        status: "error",
        message: "Periode ujian sudah selesai. BAP tidak dapat diubah.",
      };
    }
    const timingError = _examCheckBapTiming_(
      String(session.date),
      _examFormatTime_(session.time_end),
    );
    if (timingError) return { status: "error", message: timingError };
    const existing = _examGetBap_(supervisor_id);
    if (
      existing &&
      existing.submitted_at &&
      String(existing.submitted_at).trim() !== ""
    ) {
      return {
        status: "error",
        message:
          "BAP sudah pernah disubmit. Gunakan fitur Edit BAP untuk mengubah isian.",
      };
    }
    const hadir = Number(bapData.peserta_hadir);
    const absen = Number(bapData.peserta_absen);
    if (isNaN(hadir) || hadir < 0 || isNaN(absen) || absen < 0) {
      return {
        status: "error",
        message:
          "Jumlah peserta hadir dan tidak hadir wajib diisi dan tidak boleh negatif.",
      };
    }
    const totalSiswaAdmin = Number(room.total_siswa) || 0;
    const totalSiswaBAP = hadir + absen;
    if (totalSiswaAdmin > 0 && totalSiswaBAP !== totalSiswaAdmin) {
      return {
        status: "error",
        message: `Jumlah total siswa tidak sesuai. Total siswa di ruang ini: ${totalSiswaAdmin} (Hadir: ${hadir} + Alpa: ${absen} = ${totalSiswaBAP}). Silakan periksa kembali data kehadiran.`,
      };
    }
    const now = _examNow_();
    const catatan = String(bapData.catatan || "").trim();
    const bapId = "BAP-" + new Date().getTime();
    getSheet(EXAM_SHEET.BAP).appendRow([
      bapId,
      String(supervisor_id),
      hadir,
      absen,
      catatan,
      now,
      now,
    ]);
    let jtmToLog =
      Number(target.jtm_val) > 0
        ? Number(target.jtm_val)
        : Number(session.jtm_val) || 0;
    const jtmAdjustment = _jtmReadAdjustment_("EXAM", supervisor_id);
    if (jtmAdjustment) {
      const tz = Session.getScriptTimeZone();
      const sessionDateStr =
        session.date instanceof Date
          ? Utilities.formatDate(session.date, tz, "yyyy-MM-dd")
          : String(session.date);
      const adjDateStr =
        jtmAdjustment.date instanceof Date
          ? Utilities.formatDate(jtmAdjustment.date, tz, "yyyy-MM-dd")
          : String(jtmAdjustment.date);
      const storedAdjusted = jtmAdjustment.adjusted_jtm;
      const dateMatches = !adjDateStr || adjDateStr === sessionDateStr;
      if (
        dateMatches &&
        storedAdjusted !== null &&
        storedAdjusted !== undefined &&
        storedAdjusted !== "" &&
        !isNaN(Number(storedAdjusted))
      ) {
        jtmToLog = Number(storedAdjusted);
      }
    }
    _examAddJtmLog_(
      "EXAM-SUPERVISOR",
      target.user_id,
      String(session.date),
      "Pengawas Ujian: " +
        String(session.session_name || "") +
        " — " +
        String(room.room_name || ""),
      hadir,
      absen,
      jtmToLog,
      supervisor_id,
    );
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message:
        "BAP berhasil dikirim. JTM Pengawas telah dicatat ke sistem honorarium.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function editExamBAP(token, supervisor_id, bapData) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (String(target.user_id) !== String(user.id)) {
      return {
        status: "error",
        message: "Anda tidak terdaftar sebagai pengawas pada jadwal ini.",
      };
    }
    if (String(target.status) === "cancelled") {
      return {
        status: "error",
        message:
          "Substitusi Anda pada jadwal ini telah dibatalkan. BAP bersifat read-only.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord pengawas ini tidak aktif." };
    }
    if (!target.confirmed_by || String(target.confirmed_by).trim() === "") {
      return {
        status: "error",
        message: "Kehadiran Anda belum dikonfirmasi. BAP tidak dapat diedit.",
      };
    }
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(session.period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    if (_examToday_() > String(period.date_end)) {
      return {
        status: "error",
        message: "Periode ujian sudah selesai. BAP tidak dapat diubah.",
      };
    }
    const timingError = _examCheckBapTiming_(
      String(session.date),
      _examFormatTime_(session.time_end),
    );
    if (timingError) return { status: "error", message: timingError };
    const existing = _examGetBap_(supervisor_id);
    if (
      !existing ||
      !existing.submitted_at ||
      String(existing.submitted_at).trim() === ""
    ) {
      return {
        status: "error",
        message:
          "BAP belum pernah disubmit. Gunakan Submit BAP terlebih dahulu.",
      };
    }
    const hadir = Number(bapData.peserta_hadir);
    const absen = Number(bapData.peserta_absen);
    if (isNaN(hadir) || hadir < 0 || isNaN(absen) || absen < 0) {
      return {
        status: "error",
        message:
          "Jumlah peserta hadir dan tidak hadir wajib diisi dan tidak boleh negatif.",
      };
    }
    const totalSiswaAdmin = Number(room.total_siswa) || 0;
    const totalSiswaBAP = hadir + absen;
    if (totalSiswaAdmin > 0 && totalSiswaBAP !== totalSiswaAdmin) {
      return {
        status: "error",
        message: `Jumlah total siswa tidak sesuai. Total siswa di ruang ini: ${totalSiswaAdmin} (Hadir: ${hadir} + Alpa: ${absen} = ${totalSiswaBAP}). Silakan periksa kembali data kehadiran.`,
      };
    }
    const now = _examNow_();
    const catatan = String(bapData.catatan || "").trim();
    const bapSheet = getSheet(EXAM_SHEET.BAP);
    const bapRows = bapSheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < bapRows.length; i++) {
      if (String(bapRows[i][1]) === String(supervisor_id)) {
        bapSheet.getRange(i + 1, 3).setValue(hadir);
        bapSheet.getRange(i + 1, 4).setValue(absen);
        bapSheet.getRange(i + 1, 5).setValue(catatan);
        bapSheet.getRange(i + 1, 7).setValue(now);
        updated = true;
        break;
      }
    }
    if (!updated)
      return { status: "error", message: "Gagal memperbarui data BAP." };
    _examUpdateLogByRef_(supervisor_id, hadir, absen);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return { status: "success", message: "BAP berhasil diperbarui." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function substituteExamSupervisor(token, supervisor_id, new_user_id) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    if (!new_user_id)
      return { status: "error", message: "ID guru pengganti tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const sessionDate = String(session.date);
    if (_examToday_() !== sessionDate) {
      return {
        status: "error",
        message:
          "Batas waktu substitusi sudah terlewat. Substitusi hanya dapat dilakukan pada hari H.",
      };
    }
    const auth = _examCheckAdminOrConfirmedPanitia_(user, sessionDate);
    if (!auth.ok) return { status: "error", message: auth.error };
    if (String(target.status) !== "active") {
      return {
        status: "error",
        message: "Rekord pengawas ini tidak aktif dan tidak dapat digantikan.",
      };
    }
    const newGuru = getData("Users").find(
      (u) => String(u.id) === String(new_user_id),
    );
    if (!newGuru)
      return { status: "error", message: "Guru pengganti tidak ditemukan." };
    if (String(new_user_id) === String(target.user_id)) {
      return {
        status: "error",
        message: "Guru pengganti tidak boleh sama dengan guru yang diganti.",
      };
    }
    const conflict = _examCheckSupervisorConflict_(
      new_user_id,
      room.session_id,
      supervisor_id,
    );
    if (conflict)
      return { status: "error", message: "Guru pengganti: " + conflict };
    const supSheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const supRows = supSheet.getDataRange().getValues();
    for (let i = 1; i < supRows.length; i++) {
      if (String(supRows[i][0]) === String(supervisor_id)) {
        supSheet.getRange(i + 1, 4).setValue("substituted");
        break;
      }
    }
    const newId = "EXSUP-" + new Date().getTime();
    supSheet.appendRow([
      newId,
      String(target.room_id),
      String(new_user_id),
      "active",
      true,
      String(target.user_id),
      "",
      "",
      "",
      "",
      Number(target.jtm_val) || 0,
    ]);
    try {
      var byName = String(
        user.full_name ||
          user.name ||
          user.username ||
          (auth.isAdmin ? "Admin" : "Panitia Ujian"),
      );
      var sessionInfo =
        String(session.session_name || "") +
        " · " +
        String(room.room_name || "") +
        " (" +
        String(room.subject || "") +
        ")";
      _notifSupervisorSubstituted_(
        target.user_id,
        new_user_id,
        sessionDate,
        byName,
        sessionInfo,
      );
    } catch (_) {}
    return {
      status: "success",
      message:
        String(newGuru.full_name) + " kini menggantikan posisi pengawas.",
      new_supervisor_id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function cancelExamSupervisorSubstitution(token, supervisor_id, force) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (!(
      target.is_substitute === true || String(target.is_substitute) === "true"
    )) {
      return {
        status: "error",
        message: "Rekord ini bukan pengganti. Pembatalan tidak berlaku.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Substitusi ini sudah tidak aktif." };
    }
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const sessionDate = String(session.date);
    if (_examToday_() !== sessionDate) {
      return {
        status: "error",
        message: "Pembatalan substitusi hanya dapat dilakukan pada hari H.",
      };
    }
    const auth = _examCheckAdminOrConfirmedPanitia_(user, sessionDate);
    if (!auth.ok) return { status: "error", message: auth.error };
    const isConfirmed = !!(
      target.confirmed_by && String(target.confirmed_by).trim() !== ""
    );
    const bap = _examGetBap_(supervisor_id);
    const hasBap = !!(
      bap &&
      bap.submitted_at &&
      String(bap.submitted_at).trim() !== ""
    );
    if (isConfirmed && hasBap && !force) {
      return {
        status: "confirm_required",
        message:
          "Guru pengganti sudah mengisi dan menyerahkan BAP. " +
          "Pembatalan akan mengunci BAP pengganti dan menghapus JTM-nya dari honorarium. " +
          "Guru asli akan dikembalikan ke jadwal. Lanjutkan?",
      };
    }
    const now = _examNow_();
    const supSheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const supRows = supSheet.getDataRange().getValues();
    if (isConfirmed && hasBap) {
      for (let i = 1; i < supRows.length; i++) {
        if (String(supRows[i][0]) === String(supervisor_id)) {
          supSheet.getRange(i + 1, 4).setValue("cancelled");
          supSheet.getRange(i + 1, 9).setValue(String(user.id));
          supSheet.getRange(i + 1, 10).setValue(now);
          break;
        }
      }
      _examDeleteLogByRef_(supervisor_id);
    } else {
      for (let i = supRows.length - 1; i >= 1; i--) {
        if (String(supRows[i][0]) === String(supervisor_id)) {
          supSheet.deleteRow(i + 1);
          break;
        }
      }
    }
    const originalUserId = String(target.original_user_id);
    const freshRows = getSheet(EXAM_SHEET.SUPERVISORS)
      .getDataRange()
      .getValues();
    for (let i = 1; i < freshRows.length; i++) {
      if (
        String(freshRows[i][1]) === String(target.room_id) &&
        String(freshRows[i][2]) === originalUserId &&
        String(freshRows[i][3]) === "substituted"
      ) {
        const s = getSheet(EXAM_SHEET.SUPERVISORS);
        s.getRange(i + 1, 4).setValue("active");
        s.getRange(i + 1, 7).setValue("");
        s.getRange(i + 1, 8).setValue("");
        break;
      }
    }
    try {
      var byName = String(
        user.full_name ||
          user.name ||
          user.username ||
          (auth.isAdmin ? "Admin" : "Panitia Ujian"),
      );
      var sessionInfo =
        String(session.session_name || "") +
        " · " +
        String(room.room_name || "") +
        " (" +
        String(room.subject || "") +
        ")";
      _notifSupervisorSubCancelled_(
        target.user_id,
        sessionDate,
        byName,
        sessionInfo,
      );
    } catch (_) {}
    return {
      status: "success",
      message:
        "Substitusi pengawas berhasil dibatalkan. Guru asli telah dikembalikan ke jadwal.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function substituteExamCommittee(token, committee_id, new_user_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat mengganti Panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    if (!new_user_id)
      return { status: "error", message: "ID guru pengganti tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    const comDate = String(target.date);
    if (_examToday_() !== comDate) {
      return {
        status: "error",
        message: "Substitusi hanya dapat dilakukan pada hari H.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord panitia ini tidak aktif." };
    }
    const newGuru = getData("Users").find(
      (u) => String(u.id) === String(new_user_id),
    );
    if (!newGuru)
      return { status: "error", message: "Guru pengganti tidak ditemukan." };
    if (String(new_user_id) === String(target.user_id)) {
      return {
        status: "error",
        message: "Guru pengganti tidak boleh sama dengan guru yang diganti.",
      };
    }
    const conflict = _examCheckCommitteeConflict_(
      new_user_id,
      comDate,
      committee_id,
    );
    if (conflict)
      return { status: "error", message: "Guru pengganti: " + conflict };
    const comSheet = getSheet(EXAM_SHEET.COMMITTEE);
    const comRows = comSheet.getDataRange().getValues();
    for (let i = 1; i < comRows.length; i++) {
      if (String(comRows[i][0]) === String(committee_id)) {
        comSheet.getRange(i + 1, 6).setValue("substituted");
        break;
      }
    }
    const newId = "EXCOM-" + new Date().getTime();
    const jtmVal = Number(target.jtm_val) || 0;
    comSheet.appendRow([
      newId,
      String(target.period_id),
      comDate,
      String(new_user_id),
      jtmVal,
      "active",
      true,
      String(target.user_id),
      "",
      "",
      "",
      "",
    ]);
    try {
      _notifCommitteeSubstituted_(target.user_id, new_user_id, comDate);
    } catch (_) {}
    return {
      status: "success",
      message: String(newGuru.full_name) + " kini menggantikan posisi panitia.",
      new_committee_id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function cancelExamCommitteeSubstitution(token, committee_id, force) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat membatalkan substitusi Panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    if (!(
      target.is_substitute === true || String(target.is_substitute) === "true"
    )) {
      return {
        status: "error",
        message: "Rekord ini bukan pengganti. Pembatalan tidak berlaku.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Substitusi ini sudah tidak aktif." };
    }
    const comDate = String(target.date);
    if (_examToday_() !== comDate) {
      return {
        status: "error",
        message: "Pembatalan substitusi hanya dapat dilakukan pada hari H.",
      };
    }
    const isConfirmed = !!(
      target.confirmed_by && String(target.confirmed_by).trim() !== ""
    );
    if (isConfirmed && !force) {
      return {
        status: "confirm_required",
        message:
          "Guru pengganti sudah dikonfirmasi kehadirannya dan JTM-nya sudah terhitung. " +
          "Pembatalan akan menghapus JTM pengganti dari honorarium. " +
          "Guru asli akan dikembalikan ke jadwal. Lanjutkan?",
      };
    }
    const now = _examNow_();
    const comSheet = getSheet(EXAM_SHEET.COMMITTEE);
    const comRows = comSheet.getDataRange().getValues();
    if (isConfirmed) {
      for (let i = 1; i < comRows.length; i++) {
        if (String(comRows[i][0]) === String(committee_id)) {
          comSheet.getRange(i + 1, 6).setValue("cancelled");
          comSheet.getRange(i + 1, 11).setValue(String(user.id));
          comSheet.getRange(i + 1, 12).setValue(now);
          break;
        }
      }
      _examDeleteLogByRef_(committee_id);
    } else {
      for (let i = comRows.length - 1; i >= 1; i--) {
        if (String(comRows[i][0]) === String(committee_id)) {
          comSheet.deleteRow(i + 1);
          break;
        }
      }
    }
    const originalUserId = String(target.original_user_id);
    const freshRows = getSheet(EXAM_SHEET.COMMITTEE).getDataRange().getValues();
    for (let i = 1; i < freshRows.length; i++) {
      if (
        String(freshRows[i][1]) === String(target.period_id) &&
        String(freshRows[i][2]) === comDate &&
        String(freshRows[i][3]) === originalUserId &&
        String(freshRows[i][5]) === "substituted"
      ) {
        const s = getSheet(EXAM_SHEET.COMMITTEE);
        s.getRange(i + 1, 6).setValue("active");
        s.getRange(i + 1, 9).setValue("");
        s.getRange(i + 1, 10).setValue("");
        break;
      }
    }
    try {
      _notifCommitteeSubCancelled_(target.user_id, comDate);
    } catch (_) {}
    return {
      status: "success",
      message:
        "Substitusi panitia berhasil dibatalkan. Guru asli telah dikembalikan ke jadwal.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function getExamPageData(token, dateStr) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    const targetDate = dateStr ? String(dateStr) : _examToday_();
    const activePeriod = getData(EXAM_SHEET.PERIODS).find(
      (p) =>
        targetDate >= String(p.date_start) && targetDate <= String(p.date_end),
    );
    if (!activePeriod) {
      return {
        status: "success",
        is_exam_period: false,
        supervisor_duties: [],
        committee_duties: [],
        role: "none",
        is_confirmed_panitia: false,
        supervisors_to_confirm: [],
      };
    }
    const periodObj = {
      id: String(activePeriod.id),
      name: String(activePeriod.name || ""),
      date_start: String(activePeriod.date_start),
      date_end: String(activePeriod.date_end),
      jtm_committee_per_day: Number(activePeriod.jtm_committee_per_day) || 0,
      description: String(activePeriod.description || ""),
    };
    const allUsers = getData("Users");
    const allSessions = getData(EXAM_SHEET.SESSIONS).filter(
      (s) =>
        String(s.period_id) === String(activePeriod.id) &&
        String(s.date) === targetDate,
    );
    const allRooms = getData(EXAM_SHEET.ROOMS);
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const allBaps = getData(EXAM_SHEET.BAP);
    const allCom = getData(EXAM_SHEET.COMMITTEE);
    const supervisorDuties = [];
    allSessions.forEach((s) => {
      const te = _examFormatTime_(s.time_end);
      const [teH, teM] = te.split(":").map(Number);
      const activeTotal = teH * 60 + teM + 5;
      const bapActiveTime =
        String(Math.floor(activeTotal / 60)).padStart(2, "0") +
        ":" +
        String(activeTotal % 60).padStart(2, "0");
      allRooms
        .filter((r) => String(r.session_id) === String(s.id))
        .forEach((r) => {
          const mySup = allSups.find(
            (sup) =>
              String(sup.room_id) === String(r.id) &&
              String(sup.user_id) === String(user.id) &&
              String(sup.status) === "active",
          );
          if (!mySup) return;
          const bap = allBaps.find(
            (b) => String(b.supervisor_id) === String(mySup.id),
          );
          const hasBap = !!(
            bap &&
            bap.submitted_at &&
            String(bap.submitted_at).trim() !== ""
          );
          supervisorDuties.push({
            supervisor_id: String(mySup.id),
            session_id: String(s.id),
            session_name: String(s.session_name || ""),
            date: String(s.date),
            time_start: _examFormatTime_(s.time_start),
            time_end: te,
            bap_active_time: bapActiveTime,
            jtm_val: Number(mySup.jtm_val) || 0,
            room_id: String(r.id),
            room_name: String(r.room_name || ""),
            subject: String(r.subject || ""),
            class_name: String(r.class_name || ""),
            is_substitute:
              mySup.is_substitute === true ||
              String(mySup.is_substitute) === "true",
            confirmed_by: String(mySup.confirmed_by || ""),
            confirmed_at: String(mySup.confirmed_at || ""),
            has_bap: hasBap,
            bap: hasBap
              ? {
                  id: String(bap.id),
                  peserta_hadir: Number(bap.peserta_hadir) || 0,
                  peserta_absen: Number(bap.peserta_absen) || 0,
                  catatan: String(bap.catatan || ""),
                  submitted_at: String(bap.submitted_at || ""),
                  updated_at: String(bap.updated_at || ""),
                }
              : null,
          });
        });
    });
    const committeeDuties = allCom
      .filter(
        (c) =>
          String(c.period_id) === String(activePeriod.id) &&
          String(c.date) === targetDate &&
          String(c.user_id) === String(user.id) &&
          String(c.status) === "active",
      )
      .map((c) => ({
        committee_id: String(c.id),
        date: String(c.date),
        jtm_val: Number(c.jtm_val) || 0,
        is_substitute:
          c.is_substitute === true || String(c.is_substitute) === "true",
        confirmed_by: String(c.confirmed_by || ""),
        confirmed_at: String(c.confirmed_at || ""),
      }));
    const isConfirmedPanitia = committeeDuties.some(
      (c) => c.confirmed_by && String(c.confirmed_by).trim() !== "",
    );
    let role = "none";
    if (supervisorDuties.length > 0) role = "supervisor";
    else if (committeeDuties.length > 0) role = "committee";
    const supervisorsToConfirm = [];
    if (isConfirmedPanitia) {
      allSessions.forEach((s) => {
        const roomsInSession = allRooms.filter(
          (r) => String(r.session_id) === String(s.id),
        );
        roomsInSession.forEach((r) => {
          const supsInRoom = allSups
            .filter((sup) => String(sup.room_id) === String(r.id))
            .map((sup) => {
              const g = allUsers.find(
                (u) => String(u.id) === String(sup.user_id),
              );
              const bap = allBaps.find(
                (b) => String(b.supervisor_id) === String(sup.id),
              );
              return {
                supervisor_id: String(sup.id),
                user_id: String(sup.user_id),
                guru_name: g ? String(g.full_name) : "Tidak Dikenal",
                status: String(sup.status || "active"),
                is_substitute:
                  sup.is_substitute === true ||
                  String(sup.is_substitute) === "true",
                original_user_id: String(sup.original_user_id || ""),
                confirmed_by: String(sup.confirmed_by || ""),
                confirmed_at: String(sup.confirmed_at || ""),
                has_bap: !!(
                  bap &&
                  bap.submitted_at &&
                  String(bap.submitted_at).trim() !== ""
                ),
              };
            });
          if (supsInRoom.length > 0) {
            supervisorsToConfirm.push({
              session_id: String(s.id),
              session_name: String(s.session_name || ""),
              time_start: _examFormatTime_(s.time_start),
              time_end: _examFormatTime_(s.time_end),
              room_id: String(r.id),
              room_name: String(r.room_name || ""),
              subject: String(r.subject || ""),
              supervisors: supsInRoom,
            });
          }
        });
      });
    }
    return {
      status: "success",
      is_exam_period: true,
      period: periodObj,
      role,
      supervisor_duties: supervisorDuties,
      committee_duties: committeeDuties,
      is_confirmed_panitia: isConfirmedPanitia,
      supervisors_to_confirm: supervisorsToConfirm,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function getSupervisorBapDetail(token, supervisorId) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisorId)
      return { status: "error", message: "Supervisor ID kosong." };
    const sups = getData(EXAM_SHEET.SUPERVISORS);
    const sup = sups.find((s) => String(s.id) === String(supervisorId));
    if (!sup)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    const isOwner = String(sup.user_id) === String(user.id);
    const isAdmin = String(user.role).toLowerCase() === "admin";
    if (!isOwner && !isAdmin) {
      return {
        status: "error",
        message: "Anda tidak memiliki akses ke BAP ini.",
      };
    }
    const rooms = getData(EXAM_SHEET.ROOMS);
    const room = rooms.find((r) => String(r.id) === String(sup.room_id));
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const sessions = getData(EXAM_SHEET.SESSIONS);
    const session = sessions.find(
      (s) => String(s.id) === String(room.session_id),
    );
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const periods = getData(EXAM_SHEET.PERIODS);
    const period = periods.find(
      (p) => String(p.id) === String(session.period_id),
    );
    const baps = getData(EXAM_SHEET.BAP);
    const bap = baps.find(
      (b) => String(b.supervisor_id) === String(supervisorId),
    );
    const hasBap = !!(
      bap &&
      bap.submitted_at &&
      String(bap.submitted_at).trim() !== ""
    );
    const today = _examToday_();
    const periodOver = !!(period && today > String(period.date_end));
    const isCancelled = String(sup.status) === "cancelled";
    const te = _examFormatTime_(session.time_end);
    return {
      status: "success",
      supervisor_id: String(sup.id),
      session_name: String(session.session_name || ""),
      date: String(session.date),
      time_start: _examFormatTime_(session.time_start),
      time_end: te,
      jtm_val: Number(session.jtm_val) || 0,
      room_name: String(room.room_name || ""),
      subject: String(room.subject || ""),
      class_name: String(room.class_name || ""),
      is_substitute:
        sup.is_substitute === true || String(sup.is_substitute) === "true",
      status: String(sup.status || ""),
      is_cancelled: isCancelled,
      period_over: periodOver,
      period: period
        ? {
            id: String(period.id),
            name: String(period.name || ""),
            date_start: String(period.date_start),
            date_end: String(period.date_end),
          }
        : null,
      has_bap: hasBap,
      bap: hasBap
        ? {
            id: String(bap.id),
            peserta_hadir: Number(bap.peserta_hadir) || 0,
            peserta_absen: Number(bap.peserta_absen) || 0,
            catatan: String(bap.catatan || ""),
            submitted_at: String(bap.submitted_at || ""),
            updated_at: String(bap.updated_at || ""),
          }
        : null,
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}
function getExamRecapData(token, period_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat melihat rekap.",
      };
    }
    if (!period_id)
      return { status: "error", message: "ID periode wajib diisi." };
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const today = _examToday_();
    const ds = String(period.date_start);
    const de = String(period.date_end);
    const pStatus =
      today < ds ? "Akan Datang" : today > de ? "Selesai" : "Berlangsung";
    const allUsers = getData("Users");
    const allSessions = getData(EXAM_SHEET.SESSIONS).filter(
      (s) => String(s.period_id) === String(period_id),
    );
    const allRooms = getData(EXAM_SHEET.ROOMS);
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const allCom = getData(EXAM_SHEET.COMMITTEE).filter(
      (c) => String(c.period_id) === String(period_id),
    );
    const allBaps = getData(EXAM_SHEET.BAP);
    const sessionIds = new Set(allSessions.map((s) => String(s.id)));
    const sessionMap = {};
    allSessions.forEach((s) => {
      sessionMap[String(s.id)] = s;
    });
    const roomsInPeriod = allRooms.filter((r) =>
      sessionIds.has(String(r.session_id)),
    );
    const roomMap = {};
    roomsInPeriod.forEach((r) => {
      roomMap[String(r.id)] = r;
    });
    const roomIds = new Set(roomsInPeriod.map((r) => String(r.id)));
    const supsInPeriod = allSups.filter((sup) =>
      roomIds.has(String(sup.room_id)),
    );
    const jtmMap = {};
    const initGuru = (userId) => {
      if (!jtmMap[userId]) {
        const g = allUsers.find((u) => String(u.id) === userId);
        jtmMap[userId] = {
          user_id: userId,
          guru_name: g ? String(g.full_name) : "Tidak Dikenal",
          jtm_supervisor: 0,
          jtm_committee: 0,
          sessions_done: 0,
          sessions_confirmed: 0,
          bap_submitted: 0,
          duty_detail: [],
        };
      }
    };
    supsInPeriod.forEach((sup) => {
      if (String(sup.status) === "cancelled") return;
      const room = roomMap[String(sup.room_id)];
      const session = room ? sessionMap[String(room.session_id)] : null;
      if (!session) return;
      const userId = String(sup.user_id);
      initGuru(userId);
      const isConfirmed = !!(
        sup.confirmed_by && String(sup.confirmed_by).trim() !== ""
      );
      const bap = allBaps.find(
        (b) => String(b.supervisor_id) === String(sup.id),
      );
      const hasBap = !!(
        bap &&
        bap.submitted_at &&
        String(bap.submitted_at).trim() !== ""
      );
      let supJtm =
        Number(sup.jtm_val) > 0
          ? Number(sup.jtm_val)
          : Number(session.jtm_val) || 0;
      const jtmAdjustment = _jtmReadAdjustment_("EXAM", String(sup.id));
      if (jtmAdjustment) {
        const tz = Session.getScriptTimeZone();
        const sessionDateStr =
          session.date instanceof Date
            ? Utilities.formatDate(session.date, tz, "yyyy-MM-dd")
            : String(session.date);
        const adjDateStr =
          jtmAdjustment.date instanceof Date
            ? Utilities.formatDate(jtmAdjustment.date, tz, "yyyy-MM-dd")
            : String(jtmAdjustment.date);
        const storedAdjusted = jtmAdjustment.adjusted_jtm;
        const dateMatches = !adjDateStr || adjDateStr === sessionDateStr;
        if (
          dateMatches &&
          storedAdjusted !== null &&
          storedAdjusted !== undefined &&
          storedAdjusted !== "" &&
          !isNaN(Number(storedAdjusted))
        ) {
          supJtm = Number(storedAdjusted);
        }
      }
      if (isConfirmed) jtmMap[userId].sessions_confirmed++;
      if (hasBap) {
        jtmMap[userId].bap_submitted++;
        jtmMap[userId].sessions_done++;
        jtmMap[userId].jtm_supervisor += supJtm;
      }
      jtmMap[userId].duty_detail.push({
        type: "supervisor",
        date: String(session.date),
        session_name: String(session.session_name || ""),
        room_name: String(room.room_name || ""),
        is_substitute:
          sup.is_substitute === true || String(sup.is_substitute) === "true",
        is_confirmed: isConfirmed,
        has_bap: hasBap,
        jtm_val: hasBap ? supJtm : 0,
      });
    });
    allCom.forEach((c) => {
      if (String(c.status) === "cancelled") return;
      const userId = String(c.user_id);
      initGuru(userId);
      const isConfirmed = !!(
        c.confirmed_by && String(c.confirmed_by).trim() !== ""
      );
      if (isConfirmed) {
        jtmMap[userId].jtm_committee += Number(c.jtm_val) || 0;
      }
      jtmMap[userId].duty_detail.push({
        type: "committee",
        date: String(c.date),
        is_substitute:
          c.is_substitute === true || String(c.is_substitute) === "true",
        is_confirmed: isConfirmed,
        jtm_val: isConfirmed ? Number(c.jtm_val) || 0 : 0,
      });
    });
    const recapJtm = Object.values(jtmMap)
      .map((g) => ({
        ...g,
        total_jtm: g.jtm_supervisor + g.jtm_committee,
      }))
      .sort((a, b) => a.guru_name.localeCompare(b.guru_name));
    const recapBap = [];
    supsInPeriod.forEach((sup) => {
      const bap = allBaps.find(
        (b) => String(b.supervisor_id) === String(sup.id),
      );
      if (!bap || !bap.submitted_at || String(bap.submitted_at).trim() === "")
        return;
      const room = roomMap[String(sup.room_id)];
      const session = room ? sessionMap[String(room.session_id)] : null;
      const g = allUsers.find((u) => String(u.id) === String(sup.user_id));
      recapBap.push({
        supervisor_id: String(sup.id),
        user_id: String(sup.user_id),
        guru_name: g ? String(g.full_name) : "Tidak Dikenal",
        is_substitute:
          sup.is_substitute === true || String(sup.is_substitute) === "true",
        status: String(sup.status || ""),
        date: session ? String(session.date) : "",
        session_name: session ? String(session.session_name || "") : "",
        time_start: session ? _examFormatTime_(session.time_start) : "",
        time_end: session ? _examFormatTime_(session.time_end) : "",
        jtm_val:
          Number(sup.jtm_val) > 0
            ? Number(sup.jtm_val)
            : session
              ? Number(session.jtm_val) || 0
              : 0,
        room_name: room ? String(room.room_name || "") : "",
        subject: room ? String(room.subject || "") : "",
        class_name: room ? String(room.class_name || "") : "",
        peserta_hadir: Number(bap.peserta_hadir) || 0,
        peserta_absen: Number(bap.peserta_absen) || 0,
        catatan: String(bap.catatan || ""),
        submitted_at: String(bap.submitted_at || ""),
        updated_at: String(bap.updated_at || ""),
      });
    });
    recapBap.sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      return a.session_name.localeCompare(b.session_name);
    });
    return {
      status: "success",
      period: {
        id: String(period.id),
        name: String(period.name || ""),
        date_start: ds,
        date_end: de,
        jtm_committee_per_day: Number(period.jtm_committee_per_day) || 0,
        status: pStatus,
      },
      recap_jtm: recapJtm,
      recap_bap: recapBap,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}
function _cleanup(token, periodId) {
  if (!periodId) return;
  try {
    deleteExamPeriod(token, periodId, true);
    Logger.log("🧹 Cleanup: Data uji [" + periodId + "] berhasil dihapus.");
  } catch (e) {
    Logger.log(
      "⚠️  Cleanup gagal: " +
        e.message +
        " — Hapus manual via Admin > Jadwal Ujian.",
    );
  }
}
var JTM_SHEET = {
  ADJUSTMENTS: "JTM_Adjustments",
  REALLOCATIONS: "JTM_Reallocations",
};
var JTM_ADJUSTMENTS_HEADERS = [
  "id",
  "occurrence_type",
  "occurrence_ref",
  "date",
  "original_user_id",
  "scheduled_jtm",
  "adjusted_jtm",
  "jtm_difference",
  "reason",
  "adjusted_by",
  "adjusted_at",
];
var JTM_REALLOCATIONS_HEADERS = [
  "id",
  "adjustment_id",
  "occurrence_type",
  "occurrence_ref",
  "date",
  "substitute_user_id",
  "allocated_jtm",
  "created_by",
  "created_at",
  "log_id",
];
function _jtmNowIso_() {
  return Utilities.formatDate(
    new Date(),
    "Asia/Jakarta",
    "yyyy-MM-dd'T'HH:mm:ss",
  );
}
function _jtmGetAdjustmentsSheet_() {
  var ss = SpreadsheetApp.openById(getDbId());
  createSheetIfNeeded_(ss, JTM_SHEET.ADJUSTMENTS, JTM_ADJUSTMENTS_HEADERS);
  var sheet = ss.getSheetByName(JTM_SHEET.ADJUSTMENTS);
  if (!sheet)
    throw new Error(
      "Sheet JTM_Adjustments tidak ditemukan setelah createSheetIfNeeded_.",
    );
  return sheet;
}
function _jtmGetReallocationsSheet_() {
  var ss = SpreadsheetApp.openById(getDbId());
  createSheetIfNeeded_(ss, JTM_SHEET.REALLOCATIONS, JTM_REALLOCATIONS_HEADERS);
  var sheet = ss.getSheetByName(JTM_SHEET.REALLOCATIONS);
  if (!sheet)
    throw new Error(
      "Sheet JTM_Reallocations tidak ditemukan setelah createSheetIfNeeded_.",
    );
  return sheet;
}
function _jtmColMap_(sheet) {
  var lastCol = sheet.getLastColumn() || 1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[String(headers[i]).trim().toLowerCase()] = i;
  }
  return map;
}
function _jtmRowToObject_(row, headers) {
  var tz = Session.getScriptTimeZone();
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    var val = row[i];
    if (val instanceof Date) {
      try {
        var formatted = Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
        if (formatted.slice(11) === "00:00:00") {
          formatted = formatted.slice(0, 10);
        }
        val = formatted;
      } catch (_) {
        val = val.toISOString ? val.toISOString() : String(val);
      }
    }
    obj[headers[i]] = val;
  }
  return obj;
}
function _jtmBuildAdjustmentRecord_(params) {
  var scheduledJtm = Number(params.scheduledJtm) || 0;
  var adjustedJtm = Number(params.adjustedJtm) || 0;
  var difference = _jtmComputeDifference(scheduledJtm, adjustedJtm);
  var reasonRes = _jtmValidateReason(adjustedJtm, scheduledJtm, params.reason);
  var reason = reasonRes.ok
    ? reasonRes.reason
    : String(params.reason == null ? "" : params.reason).trim();
  return {
    id: params.id || generateId("JADJ"),
    occurrence_type: String(params.occurrenceType || ""),
    occurrence_ref: String(params.occurrenceRef || ""),
    date: String(params.date || ""),
    original_user_id: String(params.originalUserId || ""),
    scheduled_jtm: scheduledJtm,
    adjusted_jtm: adjustedJtm,
    jtm_difference: difference,
    reason: reason,
    adjusted_by: String(params.adjustedBy || ""),
    adjusted_at: params.adjustedAt || _jtmNowIso_(),
  };
}
function _jtmAdjustmentRecordToRow_(record) {
  return JTM_ADJUSTMENTS_HEADERS.map(function (h) {
    var v = record[h];
    return v === undefined || v === null ? "" : v;
  });
}
function _jtmFindAdjustmentRow_(occurrenceType, occurrenceRef) {
  var sheet = _jtmGetAdjustmentsSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var col = _jtmColMap_(sheet);
  var typeIdx = col["occurrence_type"];
  var refIdx = col["occurrence_ref"];
  var type = String(occurrenceType);
  var ref = String(occurrenceRef);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][typeIdx]) === type && String(data[i][refIdx]) === ref) {
      return {
        rowNumber: i + 1,
        record: _jtmRowToObject_(data[i], JTM_ADJUSTMENTS_HEADERS),
      };
    }
  }
  return null;
}
function _jtmReadAdjustment_(occurrenceType, occurrenceRef) {
  var found = _jtmFindAdjustmentRow_(occurrenceType, occurrenceRef);
  return found ? found.record : null;
}
function _jtmBuildPiketAdjustmentMap_() {
  var map = {};
  try {
    var sheet = _jtmGetAdjustmentsSheet_();
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return map;
    var col = _jtmColMap_(sheet);
    var typeIdx = col["occurrence_type"];
    var refIdx  = col["occurrence_ref"];   // format: "userId|yyyy-MM-dd"
    var adjIdx  = col["adjusted_jtm"];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][typeIdx]) !== "PIKET") continue;
      var adj = data[i][adjIdx];
      if (adj === null || adj === undefined || adj === "" || isNaN(Number(adj))) continue;
      // key sudah berupa "userId|date" — simpan langsung
      map[String(data[i][refIdx])] = Number(adj);
    }
  } catch (e) {}
  return map;
}
function _jtmResolvePiketLogJtm_(map, logUserId, logDate, fallback) {
  if (!map) return fallback;
  var tz = Session.getScriptTimeZone();
  var dStr = logDate instanceof Date
    ? Utilities.formatDate(logDate, tz, "yyyy-MM-dd")
    : String(logDate);
  var key = String(logUserId) + "|" + dStr;
  if (Object.prototype.hasOwnProperty.call(map, key)) {
    return map[key];
  }
  return fallback;
}
function _jtmBuildKbmAdjustmentMap_() {
  var map = {};
  try {
    var sheet = _jtmGetAdjustmentsSheet_();
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return map;
    var col = _jtmColMap_(sheet);
    var typeIdx = col["occurrence_type"];
    var refIdx = col["occurrence_ref"];
    var dateIdx = col["date"];
    var adjIdx = col["adjusted_jtm"];
    var tz = Session.getScriptTimeZone();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][typeIdx]) !== "KBM") continue;
      var adj = data[i][adjIdx];
      if (adj === null || adj === undefined || adj === "" || isNaN(Number(adj)))
        continue;
      var dVal = data[i][dateIdx];
      var dStr =
        dVal instanceof Date
          ? Utilities.formatDate(dVal, tz, "yyyy-MM-dd")
          : String(dVal);
      map[String(data[i][refIdx]) + "|" + dStr] = Number(adj);
    }
  } catch (e) {}
  return map;
}
function _jtmResolveKbmLogJtm_(map, scheduleId, logDate, fallback) {
  if (!map) return fallback;
  var tz = Session.getScriptTimeZone();
  var dStr =
    logDate instanceof Date
      ? Utilities.formatDate(logDate, tz, "yyyy-MM-dd")
      : String(logDate);
  var key = String(scheduleId) + "|" + dStr;
  if (Object.prototype.hasOwnProperty.call(map, key)) {
    return map[key];
  }
  return fallback;
}
function _jtmUpsertAdjustment_(record) {
  var sheet = _jtmGetAdjustmentsSheet_();
  var existing = _jtmFindAdjustmentRow_(
    record.occurrence_type,
    record.occurrence_ref,
  );
  if (existing) {
    record.id = existing.record.id;
    var row = _jtmAdjustmentRecordToRow_(record);
    sheet.getRange(existing.rowNumber, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(_jtmAdjustmentRecordToRow_(record));
  }
  return record;
}
function _jtmDeleteAdjustment_(occurrenceType, occurrenceRef) {
  var sheet = _jtmGetAdjustmentsSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  var col = _jtmColMap_(sheet);
  var typeIdx = col["occurrence_type"];
  var refIdx = col["occurrence_ref"];
  var type = String(occurrenceType);
  var ref = String(occurrenceRef);
  var deleted = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][typeIdx]) === type && String(data[i][refIdx]) === ref) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted;
}
function _jtmReallocationRecordToRow_(record) {
  return JTM_REALLOCATIONS_HEADERS.map(function (h) {
    var v = record[h];
    return v === undefined || v === null ? "" : v;
  });
}
function _jtmCreateReallocation_(params) {
  var record = {
    id: params.id || generateId("JRA"),
    adjustment_id: String(params.adjustmentId || ""),
    occurrence_type: String(params.occurrenceType || ""),
    occurrence_ref: String(params.occurrenceRef || ""),
    date: String(params.date || ""),
    substitute_user_id: String(params.substituteUserId || ""),
    allocated_jtm: Number(params.allocatedJtm) || 0,
    created_by: String(params.createdBy || ""),
    created_at: params.createdAt || _jtmNowIso_(),
    log_id: params.logId == null ? "" : String(params.logId),
  };
  _jtmGetReallocationsSheet_().appendRow(_jtmReallocationRecordToRow_(record));
  return record;
}
function _jtmReadReallocations_(occurrenceType, occurrenceRef) {
  var sheet = _jtmGetReallocationsSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var col = _jtmColMap_(sheet);
  var typeIdx = col["occurrence_type"];
  var refIdx = col["occurrence_ref"];
  var type = String(occurrenceType);
  var ref = String(occurrenceRef);
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][typeIdx]) === type && String(data[i][refIdx]) === ref) {
      out.push(_jtmRowToObject_(data[i], JTM_REALLOCATIONS_HEADERS));
    }
  }
  return out;
}
function _jtmReadReallocationById_(reallocationId) {
  var sheet = _jtmGetReallocationsSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var col = _jtmColMap_(sheet);
  var idIdx = col["id"];
  var id = String(reallocationId);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === id) {
      return _jtmRowToObject_(data[i], JTM_REALLOCATIONS_HEADERS);
    }
  }
  return null;
}
function _jtmDeleteReallocation_(reallocationId) {
  var sheet = _jtmGetReallocationsSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;
  var col = _jtmColMap_(sheet);
  var idIdx = col["id"];
  var id = String(reallocationId);
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idIdx]) === id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
function _jtmDeleteReallocationsForOccurrence_(occurrenceType, occurrenceRef) {
  var sheet = _jtmGetReallocationsSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var col = _jtmColMap_(sheet);
  var typeIdx = col["occurrence_type"];
  var refIdx = col["occurrence_ref"];
  var idIdx = col["id"];
  var type = String(occurrenceType);
  var ref = String(occurrenceRef);
  var deletedIds = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][typeIdx]) === type && String(data[i][refIdx]) === ref) {
      deletedIds.push(String(data[i][idIdx]));
      sheet.deleteRow(i + 1);
    }
  }
  return deletedIds;
}
function adjustOccurrenceJtm(token, payload) {
  var user = verifySession(token);
  if (!user) {
    return {
      status: "error",
      message: "Sesi tidak valid, silakan login kembali.",
    };
  }
  payload = payload || {};
  var occurrenceType = String(payload.occurrence_type || "").toUpperCase();
  var todayDate = _examToday_(); 
  var occurrenceRef = "";
  var occurrenceDate = "";
  var originalUserId = "";
  var scheduledJtm = 0;
  var authResult;
  if (occurrenceType === "KBM") {
    occurrenceRef = String(payload.schedule_id || "");
    if (!occurrenceRef) {
      return {
        status: "error",
        message: "ID jadwal (schedule_id) tidak ditemukan.",
      };
    }
    var schedule = getData("Schedules").find(function (s) {
      return String(s.id) === occurrenceRef;
    });
    if (!schedule) {
      return { status: "error", message: "Jadwal KBM tidak ditemukan." };
    }
    scheduledJtm = Number(schedule.jtm_val) || 0;
    originalUserId = String(payload.target_user_id || schedule.user_id || "");
    occurrenceDate = String(payload.date || todayDate);
    authResult = _jtmAuthorizeKbm(
      user,
      _isPicketOfficer(user),
      todayDate,
      occurrenceDate,
    );
  } else if (occurrenceType === "PIKET") {
    var picketTargetUid = String(payload.target_user_id || "");
    occurrenceDate = String(payload.date || todayDate);
    if (!picketTargetUid) {
      return {
        status: "error",
        message: "target_user_id tidak ditemukan untuk penyesuaian piket.",
      };
    }
    var picketLog = getData("Teaching_Logs").find(function (l) {
      return (
        String(l.user_id) === picketTargetUid &&
        String(l.schedule_id) === "PICKET-DUTY" &&
        String(l.date) === occurrenceDate
      );
    });
    if (!picketLog) {
      return {
        status: "error",
        message:
          "Kehadiran piket belum dikonfirmasi. Konfirmasi kehadiran terlebih dahulu sebelum menyesuaikan JTM.",
      };
    }
    scheduledJtm = Number(picketLog.jtm_val) || 4;
    originalUserId = picketTargetUid;
    occurrenceRef = picketTargetUid + "|" + occurrenceDate;
    authResult = _jtmAuthorizeKbm(
      user,
      _isPicketOfficer(user),
      todayDate,
      occurrenceDate,
    );
  } else if (occurrenceType === "EXAM") {
    occurrenceRef = String(payload.supervisor_id || "");
    if (!occurrenceRef) {
      return {
        status: "error",
        message: "ID pengawas (supervisor_id) tidak ditemukan.",
      };
    }
    var supervisor = getData(EXAM_SHEET.SUPERVISORS).find(function (s) {
      return String(s.id) === occurrenceRef;
    });
    if (!supervisor) {
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    }
    var room = _examGetRoom_(supervisor.room_id);
    if (!room) {
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    }
    var session = _examGetSession_(room.session_id);
    if (!session) {
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    }
    scheduledJtm =
      Number(supervisor.jtm_val) > 0
        ? Number(supervisor.jtm_val)
        : Number(session.jtm_val) || 0;
    originalUserId = String(supervisor.user_id || "");
    occurrenceDate = String(session.date || "");
    var examAuth = _examCheckAdminOrConfirmedPanitia_(user, occurrenceDate);
    authResult = _jtmAuthorizeExam(examAuth, todayDate, occurrenceDate);
  } else {
    return {
      status: "error",
      message: "Jenis kegiatan (occurrence_type) tidak valid.",
    };
  }
  if (!authResult.ok) {
    return {
      status: "error",
      code: authResult.code,
      message: authResult.message,
    };
  }
  var adjRes = _jtmValidateAdjusted(scheduledJtm, payload.adjusted_jtm);
  if (!adjRes.ok) {
    return { status: "error", code: adjRes.code, message: adjRes.message };
  }
  var adjustedJtm = adjRes.value;
  var reasonRes = _jtmValidateReason(adjustedJtm, scheduledJtm, payload.reason);
  if (!reasonRes.ok) {
    return {
      status: "error",
      code: reasonRes.code,
      message: reasonRes.message,
    };
  }
  var reason = reasonRes.reason;
  var existingAllocations = _jtmReadReallocations_(
    occurrenceType,
    occurrenceRef,
  ).map(function (r) {
    return Number(r.allocated_jtm) || 0;
  });
  var consRes = _jtmCheckConservation(
    scheduledJtm,
    adjustedJtm,
    existingAllocations,
    0,
  );
  if (!consRes.ok) {
    return { status: "error", code: consRes.code, message: consRes.message };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return {
      status: "error",
      code: JTM_ERROR_CODES.PERSISTENCE_FAILED,
      message:
        "Sistem sedang sibuk memproses penyesuaian lain. Silakan coba lagi.",
    };
  }
  try {
    var record = _jtmBuildAdjustmentRecord_({
      occurrenceType: occurrenceType,
      occurrenceRef: occurrenceRef,
      date: occurrenceDate,
      originalUserId: originalUserId,
      scheduledJtm: scheduledJtm,
      adjustedJtm: adjustedJtm,
      reason: reason,
      adjustedBy: user.id,
    });
    _jtmUpsertAdjustment_(record);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Penyesuaian JTM berhasil disimpan.",
      scheduled_jtm: scheduledJtm,
      adjusted_jtm: adjustedJtm,
      jtm_difference: record.jtm_difference,
    };
  } catch (writeErr) {
    return {
      status: "error",
      code: JTM_ERROR_CODES.PERSISTENCE_FAILED,
      message: "Penyesuaian JTM tidak dapat disimpan. Silakan coba lagi.",
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}
function getOccurrenceJtmState(token, occurrence_type, occurrence_ref) {
  try {
    return _getOccurrenceJtmState_(token, occurrence_type, occurrence_ref);
  } catch (e) {
    return {
      status: "error",
      message: "Gagal membaca status JTM: " + e.message,
    };
  }
}
function _getOccurrenceJtmState_(token, occurrence_type, occurrence_ref) {
  var user = verifySession(token);
  if (!user) {
    return {
      status: "error",
      message: "Sesi tidak valid, silakan login kembali.",
    };
  }
  var occurrenceType = String(occurrence_type || "").toUpperCase();
  var occurrenceRef = String(occurrence_ref || "");
  var scheduledJtm = 0;
  if (occurrenceType === "KBM") {
    if (!occurrenceRef) {
      return {
        status: "error",
        message: "ID jadwal (schedule_id) tidak ditemukan.",
      };
    }
    var schedule = getData("Schedules").find(function (s) {
      return String(s.id) === occurrenceRef;
    });
    if (!schedule) {
      return { status: "error", message: "Jadwal KBM tidak ditemukan." };
    }
    scheduledJtm = Number(schedule.jtm_val) || 0;
  } else if (occurrenceType === "EXAM") {
    if (!occurrenceRef) {
      return {
        status: "error",
        message: "ID pengawas (supervisor_id) tidak ditemukan.",
      };
    }
    var supervisor = getData(EXAM_SHEET.SUPERVISORS).find(function (s) {
      return String(s.id) === occurrenceRef;
    });
    if (!supervisor) {
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    }
    var room = _examGetRoom_(supervisor.room_id);
    if (!room) {
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    }
    var session = _examGetSession_(room.session_id);
    if (!session) {
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    }
    scheduledJtm =
      Number(supervisor.jtm_val) > 0
        ? Number(supervisor.jtm_val)
        : Number(session.jtm_val) || 0;
  } else if (occurrenceType === "PIKET") {
    if (!occurrenceRef) {
      return {
        status: "error",
        message: "occurrence_ref (target_user_id|date) tidak ditemukan.",
      };
    }
    var parts = occurrenceRef.split("|");
    var picketTargetUid = parts[0];
    var occurrenceDate = parts[1] || _examToday_();
    var picketLog = getData("Teaching_Logs").find(function (l) {
      return (
        String(l.user_id) === picketTargetUid &&
        String(l.schedule_id) === "PICKET-DUTY" &&
        String(l.date) === occurrenceDate
      );
    });
    scheduledJtm = picketLog ? Number(picketLog.jtm_val) || 4 : 4;
  } else {
    return {
      status: "error",
      message: "Jenis kegiatan (occurrence_type) tidak valid.",
    };
  }
  var adjustment = _jtmReadAdjustment_(occurrenceType, occurrenceRef);
  var adjustedJtm;
  var jtmDifference;
  var reason;
  if (adjustment) {
    adjustedJtm = Number(adjustment.adjusted_jtm) || 0;
    reason = adjustment.reason == null ? "" : String(adjustment.reason);
    jtmDifference = _jtmComputeDifference(scheduledJtm, adjustedJtm);
  } else {
    adjustedJtm = scheduledJtm;
    jtmDifference = 0;
    reason = "";
  }
  var allocations = _jtmReadReallocations_(occurrenceType, occurrenceRef) || [];
  return {
    status: "success",
    scheduled_jtm: scheduledJtm,
    adjusted_jtm: adjustedJtm,
    jtm_difference: jtmDifference,
    reason: reason,
    allocations: allocations,
  };
}
function reallocateJtmDifference(token, payload) {
  var user = verifySession(token);
  if (!user) {
    return {
      status: "error",
      message: "Sesi tidak valid, silakan login kembali.",
    };
  }
  payload = payload || {};
  var occurrenceType = String(payload.occurrence_type || "").toUpperCase();
  var occurrenceRef = String(payload.occurrence_ref || "");
  var todayDate = _examToday_(); 
  var occurrenceDate = "";
  var originalUserId = "";
  var scheduledJtm = 0;
  var partialSubScheduleId = "";
  var authResult;
  if (occurrenceType === "KBM") {
    if (!occurrenceRef) {
      return {
        status: "error",
        message: "ID jadwal (schedule_id) tidak ditemukan.",
      };
    }
    var schedule = getData("Schedules").find(function (s) {
      return String(s.id) === occurrenceRef;
    });
    if (!schedule) {
      return { status: "error", message: "Jadwal KBM tidak ditemukan." };
    }
    scheduledJtm = Number(schedule.jtm_val) || 0;
    originalUserId = String(schedule.user_id || "");
    occurrenceDate = String(payload.date || todayDate);
    partialSubScheduleId = "PARTIAL-SUB-KBM";
    authResult = _jtmAuthorizeKbm(
      user,
      _isPicketOfficer(user),
      todayDate,
      occurrenceDate,
    );
  } else if (occurrenceType === "PIKET") {
    var parts = occurrenceRef.split("|");
    var picketUid = parts[0] || "";
    occurrenceDate = parts[1] || todayDate;
    if (!picketUid) {
      return {
        status: "error",
        message: "occurrence_ref tidak valid untuk tipe PIKET.",
      };
    }
    var pLog = getData("Teaching_Logs").find(function (l) {
      return (
        String(l.user_id) === picketUid &&
        String(l.schedule_id) === "PICKET-DUTY" &&
        String(l.date) === occurrenceDate
      );
    });
    if (!pLog) {
      return {
        status: "error",
        message:
          "Kehadiran piket belum dikonfirmasi. Konfirmasi kehadiran terlebih dahulu.",
      };
    }
    scheduledJtm = Number(pLog.jtm_val) || 4;
    originalUserId = picketUid;
    partialSubScheduleId = "PARTIAL-SUB-PIKET";
    authResult = _jtmAuthorizeKbm(
      user,
      _isPicketOfficer(user),
      todayDate,
      occurrenceDate,
    );
  } else if (occurrenceType === "EXAM") {
    if (!occurrenceRef) {
      return {
        status: "error",
        message: "ID pengawas (supervisor_id) tidak ditemukan.",
      };
    }
    var supervisor = getData(EXAM_SHEET.SUPERVISORS).find(function (s) {
      return String(s.id) === occurrenceRef;
    });
    if (!supervisor) {
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    }
    var room = _examGetRoom_(supervisor.room_id);
    if (!room) {
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    }
    var session = _examGetSession_(room.session_id);
    if (!session) {
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    }
    scheduledJtm =
      Number(supervisor.jtm_val) > 0
        ? Number(supervisor.jtm_val)
        : Number(session.jtm_val) || 0;
    originalUserId = String(supervisor.user_id || "");
    occurrenceDate = String(session.date || "");
    partialSubScheduleId = "PARTIAL-SUB-EXAM";
    var examAuth = _examCheckAdminOrConfirmedPanitia_(user, occurrenceDate);
    authResult = _jtmAuthorizeExam(examAuth, todayDate, occurrenceDate);
  } else {
    return {
      status: "error",
      message: "Jenis kegiatan (occurrence_type) tidak valid.",
    };
  }
  if (!authResult.ok) {
    return {
      status: "error",
      code: authResult.code,
      message: authResult.message,
    };
  }
  var adjustment = _jtmReadAdjustment_(occurrenceType, occurrenceRef);
  var adjustedJtm = adjustment
    ? Number(adjustment.adjusted_jtm) || 0
    : scheduledJtm;
  var adjustmentId = adjustment ? String(adjustment.id || "") : "";
  var rawAllocations = payload.allocations;
  if (
    !rawAllocations ||
    typeof rawAllocations.length !== "number" ||
    rawAllocations.length === 0
  ) {
    return { status: "error", message: "Tidak ada alokasi JTM yang dikirim." };
  }
  var usersById = {};
  getData("Users").forEach(function (u) {
    usersById[String(u.id)] = u;
  });
  var existingRecords =
    _jtmReadReallocations_(occurrenceType, occurrenceRef) || [];
  var conservationBase = existingRecords.map(function (r) {
    return Number(r.allocated_jtm) || 0;
  });
  var validated = []; 
  var runningAllocations = conservationBase.slice();
  for (var i = 0; i < rawAllocations.length; i++) {
    var alloc = rawAllocations[i] || {};
    var substituteUserId = String(alloc.substitute_user_id || "");
    if (!substituteUserId) {
      return {
        status: "error",
        code: JTM_ERROR_CODES.ALLOCATION_INVALID,
        message: "ID guru pengganti (substitute_user_id) tidak ditemukan.",
      };
    }
    var subUser = usersById[substituteUserId];
    var isSubstituteAdmin = !!(
      subUser && String(subUser.role).toLowerCase() === "admin"
    );
    var allocRes = _jtmValidateAllocation(
      alloc.allocated_jtm,
      originalUserId,
      substituteUserId,
      isSubstituteAdmin,
    );
    if (!allocRes.ok) {
      return {
        status: "error",
        code: allocRes.code,
        message: allocRes.message,
      };
    }
    var consRes = _jtmCheckConservation(
      scheduledJtm,
      adjustedJtm,
      runningAllocations,
      allocRes.value,
    );
    if (!consRes.ok) {
      return { status: "error", code: consRes.code, message: consRes.message };
    }
    runningAllocations.push(allocRes.value);
    validated.push({
      substituteUserId: substituteUserId,
      value: allocRes.value,
    });
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return {
      status: "error",
      code: JTM_ERROR_CODES.PERSISTENCE_FAILED,
      message:
        "Sistem sedang sibuk memproses penyesuaian lain. Silakan coba lagi.",
    };
  }
  var written = []; 
  try {
    var savedAllocations = [];
    for (var j = 0; j < validated.length; j++) {
      var v = validated[j];
      var record = _jtmCreateReallocation_({
        adjustmentId: adjustmentId,
        occurrenceType: occurrenceType,
        occurrenceRef: occurrenceRef,
        date: occurrenceDate,
        substituteUserId: v.substituteUserId,
        allocatedJtm: v.value,
        createdBy: user.id,
      });
      var creditLogId;
      try {
        creditLogId = _jtmAddReallocationCreditLog_(
          partialSubScheduleId,
          v.substituteUserId,
          occurrenceDate,
          v.value,
          record.id,
        );
      } catch (creditErr) {
        try {
          _jtmDeleteReallocation_(record.id);
        } catch (_) {}
        throw creditErr;
      }
      try {
        _jtmSetReallocationLogId_(record.id, creditLogId);
        record.log_id = creditLogId;
      } catch (_) {}
      written.push({ reallocationId: record.id, logId: creditLogId });
      savedAllocations.push({
        id: record.id,
        substitute_user_id: record.substitute_user_id,
        allocated_jtm: record.allocated_jtm,
        log_id: creditLogId,
      });
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    var totalAllocated = runningAllocations.reduce(function (acc, n) {
      return acc + (Number(n) || 0);
    }, 0);
    var remainingDifference = scheduledJtm - adjustedJtm - totalAllocated;
    if (remainingDifference < 0) remainingDifference = 0;
    return {
      status: "success",
      message: "Alokasi JTM pengganti sebagian berhasil disimpan.",
      allocations: savedAllocations,
      remaining_difference: remainingDifference,
    };
  } catch (writeErr) {
    for (var k = written.length - 1; k >= 0; k--) {
      try {
        _jtmDeleteReallocation_(written[k].reallocationId);
      } catch (_) {}
      if (written[k].logId) {
        try {
          _examDeleteLogByRef_(written[k].reallocationId);
        } catch (_) {}
      }
    }
    return {
      status: "error",
      code: JTM_ERROR_CODES.PERSISTENCE_FAILED,
      message: "Alokasi JTM tidak dapat disimpan. Silakan coba lagi.",
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}
function _jtmAddReallocationCreditLog_(
  scheduleId,
  userId,
  date,
  jtmVal,
  reallocationId,
) {
  var logId = "LOG-PSUB-" + new Date().getTime();
  getSheet(SHEET_NAME.LOGS).appendRow([
    logId,
    String(scheduleId),
    String(userId),
    String(date),
    "Alokasi JTM pengganti sebagian",
    0,
    0,
    "REF:" + String(reallocationId),
    Number(jtmVal) || 0,
    _jtmNowIso_(),
  ]);
  return logId;
}
function _jtmSetReallocationLogId_(reallocationId, logId) {
  var sheet = _jtmGetReallocationsSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;
  var col = _jtmColMap_(sheet);
  var idIdx = col["id"];
  var logIdx = col["log_id"];
  if (idIdx === undefined || logIdx === undefined) return false;
  var id = String(reallocationId);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === id) {
      sheet.getRange(i + 1, logIdx + 1).setValue(String(logId));
      return true;
    }
  }
  return false;
}
function removeReallocation(token, reallocation_id) {
  var user = verifySession(token);
  if (!user) {
    return {
      status: "error",
      message: "Sesi tidak valid, silakan login kembali.",
    };
  }
  var reallocationId = String(reallocation_id || "");
  if (!reallocationId) {
    return {
      status: "error",
      message: "ID alokasi (reallocation_id) tidak ditemukan.",
    };
  }
  var record = _jtmReadReallocationById_(reallocationId);
  if (!record) {
    return {
      status: "error",
      message: "Alokasi JTM tidak ditemukan atau sudah dihapus.",
    };
  }
  var occurrenceType = String(record.occurrence_type || "").toUpperCase();
  var occurrenceDate = String(record.date || "");
  var todayDate = _examToday_(); 
  var authResult;
  if (occurrenceType === "KBM" || occurrenceType === "PIKET") {
    authResult = _jtmAuthorizeKbm(
      user,
      _isPicketOfficer(user),
      todayDate,
      occurrenceDate,
    );
  } else if (occurrenceType === "EXAM") {
    var examAuth = _examCheckAdminOrConfirmedPanitia_(user, occurrenceDate);
    authResult = _jtmAuthorizeExam(examAuth, todayDate, occurrenceDate);
  } else {
    return {
      status: "error",
      message: "Jenis kegiatan (occurrence_type) tidak valid.",
    };
  }
  if (!authResult.ok) {
    return {
      status: "error",
      code: authResult.code,
      message: authResult.message,
    };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return {
      status: "error",
      code: JTM_ERROR_CODES.PERSISTENCE_FAILED,
      message:
        "Sistem sedang sibuk memproses penyesuaian lain. Silakan coba lagi.",
    };
  }
  try {
    _examDeleteLogByRef_(reallocationId);
    _jtmDeleteReallocation_(reallocationId);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Alokasi JTM pengganti sebagian berhasil dihapus.",
    };
  } catch (writeErr) {
    return {
      status: "error",
      code: JTM_ERROR_CODES.PERSISTENCE_FAILED,
      message: "Alokasi JTM tidak dapat dihapus. Silakan coba lagi.",
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}
function resetOccurrenceJtm(token, payload) {
  try {
    var user = verifySession(token);
    if (!user) {
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    }
    payload = payload || {};
    var occurrenceType = String(payload.occurrence_type || "").toUpperCase();
    var occurrenceRef = String(payload.occurrence_ref || "");
    var todayDate = _examToday_();
    var occurrenceDate = "";
    var authResult;
    if (occurrenceType === "KBM") {
      if (!occurrenceRef) {
        return {
          status: "error",
          message: "ID jadwal (schedule_id) tidak ditemukan.",
        };
      }
      var schedule = getData("Schedules").find(function (s) {
        return String(s.id) === occurrenceRef;
      });
      if (!schedule) {
        return { status: "error", message: "Jadwal KBM tidak ditemukan." };
      }
      occurrenceDate = String(payload.date || todayDate);
      authResult = _jtmAuthorizeKbm(
        user,
        _isPicketOfficer(user),
        todayDate,
        occurrenceDate,
      );
    } else if (occurrenceType === "EXAM") {
      if (!occurrenceRef) {
        return {
          status: "error",
          message: "ID pengawas (supervisor_id) tidak ditemukan.",
        };
      }
      var supervisor = getData(EXAM_SHEET.SUPERVISORS).find(function (s) {
        return String(s.id) === occurrenceRef;
      });
      if (!supervisor) {
        return { status: "error", message: "Data pengawas tidak ditemukan." };
      }
      var room = _examGetRoom_(supervisor.room_id);
      if (!room)
        return { status: "error", message: "Ruang ujian tidak ditemukan." };
      var session = _examGetSession_(room.session_id);
      if (!session)
        return { status: "error", message: "Sesi ujian tidak ditemukan." };
      occurrenceDate = String(session.date || todayDate);
      var examAuth = _examCheckAdminOrConfirmedPanitia_(user, occurrenceDate);
      authResult = _jtmAuthorizeExam(examAuth, todayDate, occurrenceDate);
    } else if (occurrenceType === "PIKET") {
      if (!occurrenceRef) {
        return {
          status: "error",
          message: "occurrence_ref tidak ditemukan untuk tipe PIKET.",
        };
      }
      var piketParts = occurrenceRef.split("|");
      occurrenceDate = piketParts[1] || todayDate;
      authResult = _jtmAuthorizeKbm(
        user,
        _isPicketOfficer(user),
        todayDate,
        occurrenceDate,
      );
    } else {
      return {
        status: "error",
        message: "Jenis kegiatan (occurrence_type) tidak valid.",
      };
    }
    if (!authResult.ok) {
      return {
        status: "error",
        code: authResult.code,
        message: authResult.message,
      };
    }
    var existing = _jtmFindAdjustmentRow_(occurrenceType, occurrenceRef);
    if (!existing) {
      return {
        status: "error",
        code: "NOT_FOUND",
        message: "Tidak ada penyesuaian JTM aktif untuk kegiatan ini.",
      };
    }
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (lockErr) {
      return {
        status: "error",
        code: JTM_ERROR_CODES.PERSISTENCE_FAILED,
        message: "Sistem sedang sibuk. Silakan coba lagi.",
      };
    }
    try {
      var reallocations = _jtmReadReallocations_(occurrenceType, occurrenceRef);
      for (var i = 0; i < reallocations.length; i++) {
        var realloc = reallocations[i];
        if (realloc.id) {
          try {
            _examDeleteLogByRef_(realloc.id);
          } catch (_) {}
        }
      }
      _jtmDeleteAllReallocationsForOccurrence_(occurrenceType, occurrenceRef);
      _jtmDeleteAdjustment_(occurrenceType, occurrenceRef);
      try {
        _invalidateDataSnapshot();
      } catch (_) {}
      return {
        status: "success",
        message:
          "Penyesuaian JTM berhasil dibatalkan. JTM kembali ke nilai terjadwal.",
      };
    } catch (writeErr) {
      return {
        status: "error",
        code: JTM_ERROR_CODES.PERSISTENCE_FAILED,
        message: "Gagal membatalkan penyesuaian JTM: " + writeErr.message,
      };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  } catch (e) {
    return { status: "error", message: "Terjadi kesalahan: " + e.message };
  }
}
function _jtmDeleteAllReallocationsForOccurrence_(
  occurrenceType,
  occurrenceRef,
) {
  var sheet = _jtmGetReallocationsSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  var col = _jtmColMap_(sheet);
  var typeIdx = col["occurrence_type"];
  var refIdx = col["occurrence_ref"];
  var type = String(occurrenceType);
  var ref = String(occurrenceRef);
  var deleted = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][typeIdx]) === type && String(data[i][refIdx]) === ref) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted;
}
function _jtmDeleteKbmOriginalLog_(scheduleId, occurrenceDate, userId) {
  var schedId = String(scheduleId || "");
  if (!schedId) return 0;
  var sheet = getSheet(SHEET_NAME.LOGS);
  if (!sheet) return 0;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  var col = _jtmColMap_(sheet);
  var schedIdx = col["schedule_id"];
  var dateIdx = col["date"];
  var userIdx = col["user_id"];
  if (schedIdx === undefined || dateIdx === undefined) return 0;
  var tz = Session.getScriptTimeZone();
  var dateStr =
    occurrenceDate instanceof Date
      ? Utilities.formatDate(occurrenceDate, tz, "yyyy-MM-dd")
      : String(occurrenceDate || "");
  var filterUserId =
    userId !== undefined && userId !== null && String(userId).trim() !== ""
      ? String(userId).trim()
      : null;
  var deleted = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][schedIdx]) !== schedId) continue;
    if (dateStr) {
      var logDateVal = data[i][dateIdx];
      var logDateStr =
        logDateVal instanceof Date
          ? Utilities.formatDate(logDateVal, tz, "yyyy-MM-dd")
          : String(logDateVal);
      if (logDateStr !== dateStr) continue;
    }
    if (filterUserId !== null && userIdx !== undefined) {
      if (String(data[i][userIdx]) !== filterUserId) continue;
    }
    sheet.deleteRow(i + 1);
    deleted++;
  }
  return deleted;
}
function _jtmReverseOccurrence_(type, ref, date, originalUserId) {
  var occurrenceType = String(type || "").toUpperCase();
  var occurrenceRef = String(ref || "");
  if (!occurrenceRef) return;
  var reallocations = [];
  try {
    reallocations = _jtmReadReallocations_(occurrenceType, occurrenceRef) || [];
  } catch (e) {
    reallocations = [];
  }
  for (var i = 0; i < reallocations.length; i++) {
    var reallocationId = String(
      (reallocations[i] && reallocations[i].id) || "",
    );
    if (!reallocationId) continue;
    try {
      _examDeleteLogByRef_(reallocationId);
    } catch (e) {}
  }
  try {
    _jtmDeleteReallocationsForOccurrence_(occurrenceType, occurrenceRef);
  } catch (e) {}
  try {
    _jtmDeleteAdjustment_(occurrenceType, occurrenceRef);
  } catch (e) {}
  if (occurrenceType === "EXAM") {
    try {
      _examDeleteLogByRef_(occurrenceRef);
    } catch (e) {}
  } else if (occurrenceType === "KBM") {
    try {
      _jtmDeleteKbmOriginalLog_(occurrenceRef, date, originalUserId);
    } catch (e) {}
  }
  try {
    _invalidateDataSnapshot();
  } catch (e) {}
}
function getTransportConfig(token) {
  try {
    const user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali" };
    const cfg = _getConfigMap();
    const rawTarif = cfg["tarif_per_km"];
    const tarif_per_km =
      rawTarif !== undefined && rawTarif !== "" ? Number(rawTarif) : 0;
    const rawEnabled = cfg["transport_allowance_enabled"];
    const transport_allowance_enabled =
      String(rawEnabled).toLowerCase() === "true";
    return {
      status: "success",
      tarif_per_km: isNaN(tarif_per_km) ? 0 : tarif_per_km,
      transport_allowance_enabled: transport_allowance_enabled,
    };
  } catch (e) {
    console.error("getTransportConfig error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function saveTransportConfig(token, payload) {
  try {
    const user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali" };
    if (String(user.role).toLowerCase() !== "admin") {
      return { status: "error", message: "Akses hanya untuk Admin" };
    }
    const rawTarif = payload && payload.tarif_per_km;
    const tarif = Number(rawTarif);
    if (
      rawTarif === null ||
      rawTarif === undefined ||
      rawTarif === "" ||
      isNaN(tarif) ||
      !isFinite(tarif) ||
      tarif < 1 ||
      tarif > 999999999
    ) {
      return {
        status: "error",
        message:
          "Tarif per KM harus berupa angka positif antara 1 dan 999.999.999.",
      };
    }
    const enabledRaw = payload && payload.transport_allowance_enabled;
    const enabledStr =
      enabledRaw === true || String(enabledRaw).toLowerCase() === "true"
        ? "true"
        : "false";
    const sheet = getSheet("Config");
    const existingData = sheet.getDataRange().getValues();
    const updateConfigKey = (key, val) => {
      let rowIndex = -1;
      for (let i = 1; i < existingData.length; i++) {
        if (String(existingData[i][0]) === key) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex > 0) {
        sheet.getRange(rowIndex, 2).setValue(val);
      } else {
        sheet.appendRow([key, val, ""]);
      }
    };
    updateConfigKey("tarif_per_km", tarif);
    updateConfigKey("transport_allowance_enabled", enabledStr);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Konfigurasi tunjangan transportasi berhasil disimpan.",
    };
  } catch (e) {
    console.error("saveTransportConfig error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function _minutesBetween(timeStart, timeEnd) {
  if (
    typeof timeStart !== "string" ||
    typeof timeEnd !== "string" ||
    !timeStart.trim() ||
    !timeEnd.trim()
  ) {
    return 0;
  }
  var rTime = /^([01]\d|2[0-3]):([0-5]\d)$/;
  var startMatch = timeStart.trim().match(rTime);
  var endMatch = timeEnd.trim().match(rTime);
  if (!startMatch || !endMatch) return 0;
  var startMinutes =
    parseInt(startMatch[1], 10) * 60 + parseInt(startMatch[2], 10);
  var endMinutes = parseInt(endMatch[1], 10) * 60 + parseInt(endMatch[2], 10);
  var diff = endMinutes - startMinutes;
  return diff > 0 ? diff : 0;
}
function _calcAttendancePercent(durasiAktual, durasiJadwal) {
  var actual = Number(durasiAktual);
  var jadwal = Number(durasiJadwal);
  if (!isFinite(jadwal) || jadwal <= 0) return 0;
  if (!isFinite(actual) || actual < 0) actual = 0;
  var raw = (actual / jadwal) * 100;
  var capped = raw > 100 ? 100 : raw;
  return Math.round(capped * 1000) / 1000;
}
function _calcDailyTransport(
  kmDistance,
  tarifPerKm,
  durasiAktual,
  durasiJadwal,
  transportEnabled,
) {
  var enabled =
    transportEnabled === undefined || transportEnabled === null
      ? true
      : transportEnabled === true ||
        String(transportEnabled).toLowerCase() === "true";
  if (!enabled) return 0;
  var km = Number(kmDistance);
  var tarif = Number(tarifPerKm);
  var aktual = Number(durasiAktual);
  var jadwal = Number(durasiJadwal);
  if (
    !isFinite(km) ||
    !isFinite(tarif) ||
    !isFinite(aktual) ||
    !isFinite(jadwal)
  )
    return 0;
  if (km === 0 || tarif === 0) return 0;
  if (jadwal <= 0) return 0; 
  if (aktual <= 0) return 0; 
  var effectiveAktual = aktual > jadwal ? jadwal : aktual;
  return km * tarif * (effectiveAktual / jadwal);
}
function _applyThreePointRounding(totalDesimal) {
  var total = Number(totalDesimal);
  if (!isFinite(total) || isNaN(total) || total <= 0) return 0;
  var base = Math.floor(total / 1000) * 1000;
  var sisa = total - base;
  if (sisa < 333.34) {
    return base;
  } else if (sisa < 666.67) {
    return base + 500;
  } else {
    return base + 1000;
  }
}
function _getMonthlyTransportSum(userId, periodeStr) {
  _lastEventDaysCount = 0;
  var transportEnabled = String(
    getConfigValue("transport_allowance_enabled") || "false",
  )
    .toLowerCase()
    .trim();
  if (transportEnabled !== "true") return 0;
  var tarifPerKm = Number(getConfigValue("tarif_per_km") || 0);
  if (!isFinite(tarifPerKm) || tarifPerKm <= 0) return 0;
  var userRow = findData(SHEET_NAME.USERS, "id", userId);
  var kmDistance = userRow ? Number(userRow.km_distance || 0) : 0;
  if (!isFinite(kmDistance) || kmDistance <= 0) return 0;
  var monthNames = [
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
  var parts = String(periodeStr || "")
    .trim()
    .split(" ");
  var targetMonthIndex = monthNames.indexOf(parts[0]);
  var targetYear = parseInt(parts[1], 10);
  if (targetMonthIndex === -1 || isNaN(targetYear)) return 0;
  var mm =
    targetMonthIndex + 1 < 10
      ? "0" + (targetMonthIndex + 1)
      : String(targetMonthIndex + 1);
  var firstDay = targetYear + "-" + mm + "-01";
  var lastDayDate = new Date(targetYear, targetMonthIndex + 1, 0);
  var lastDd =
    lastDayDate.getDate() < 10
      ? "0" + lastDayDate.getDate()
      : String(lastDayDate.getDate());
  var lastDay = targetYear + "-" + mm + "-" + lastDd;
  var allAttendance = getData(SHEET_NAME.ATTENDANCE);
  var userEntries = allAttendance.filter(function (row) {
    if (String(row.user_id) !== String(userId)) return false;
    var d = new Date(row.date);
    if (isNaN(d.getTime())) return false;
    return d.getMonth() === targetMonthIndex && d.getFullYear() === targetYear;
  });
  var bestByDate = {};
  userEntries.forEach(function (entry) {
    var dateKey = String(entry.date || "");
    var isExtra = String(entry.extra_attendee || "").toLowerCase() === "true";
    var durasiJadwal;
    if (isExtra && !String(entry.sched_time_in || "").trim()) {
      durasiJadwal = _minutesBetween(
        String(entry.time_in || ""),
        String(entry.time_out || ""),
      );
    } else {
      durasiJadwal = _minutesBetween(
        String(entry.sched_time_in || ""),
        String(entry.sched_time_out || ""),
      );
    }
    if (!bestByDate[dateKey]) {
      bestByDate[dateKey] = { entry: entry, durasiJadwal: durasiJadwal };
    } else {
      if (durasiJadwal > bestByDate[dateKey].durasiJadwal) {
        bestByDate[dateKey] = { entry: entry, durasiJadwal: durasiJadwal };
      }
    }
  });
  var eventAttAll = getData(EVENT_SHEET.ATTENDANCE);
  var eventAttRecords = eventAttAll.filter(function (r) {
    return (
      String(r.user_id || "") === String(userId) &&
      String(r.date || "") >= firstDay &&
      String(r.date || "") <= lastDay
    );
  });
  var regularDates = Object.keys(bestByDate);
  var allUniqueDates = regularDates;
  var seenEventDates = {};
  var eventDaysCountVal = 0;
  eventAttRecords.forEach(function (r) {
    var submitted =
      r.journal_submitted === true ||
      String(r.journal_submitted).toUpperCase() === "TRUE";
    if (submitted && r.date) {
      var d = String(r.date);
      if (!seenEventDates[d]) {
        seenEventDates[d] = true;
        eventDaysCountVal++;
      }
    }
  });
  _lastEventDaysCount = eventDaysCountVal;
  if (allUniqueDates.length === 0) return 0;
  var total = 0;
  var allLeaves = getData(SHEET_NAME.ATTENDANCE_LEAVES);
  var userLeaves = allLeaves.filter(function (lv) {
    if (String(lv.user_id) !== String(userId)) return false;
    var d = new Date(lv.date);
    if (isNaN(d.getTime())) return false;
    return d.getMonth() === targetMonthIndex && d.getFullYear() === targetYear;
  });
  for (var i = 0; i < allUniqueDates.length; i++) {
    var dateKey = allUniqueDates[i];
    if (bestByDate[dateKey]) {
      var item = bestByDate[dateKey];
      var e = item.entry;
      var dJadwal = item.durasiJadwal;
      var dAktual = _minutesBetween(
        String(e.time_in || ""),
        String(e.time_out || ""),
      );
      var totalLeaveMinutes = 0;
      userLeaves.forEach(function (lv) {
        if (String(lv.date) === dateKey) {
          totalLeaveMinutes += _minutesBetween(
            String(lv.leave_time || ""),
            String(lv.return_time || ""),
          );
        }
      });
      var dAktualEfektif = dAktual - totalLeaveMinutes;
      if (dAktualEfektif < 0) dAktualEfektif = 0;
      total += _calcDailyTransport(
        kmDistance,
        tarifPerKm,
        dAktualEfektif,
        dJadwal,
        true,
      );
    }
  }
  return total;
}

function _getMonthlyTransportSumFast(userId, periodeStr, snapshot) {
  _lastEventDaysCount = 0;
  var transportEnabled = String(
    snapshot.config["transport_allowance_enabled"] || "false",
  )
    .toLowerCase()
    .trim();
  if (transportEnabled !== "true") return 0;
  var tarifPerKm = Number(snapshot.config["tarif_per_km"] || 0);
  if (!isFinite(tarifPerKm) || tarifPerKm <= 0) return 0;
  var userRow = snapshot.usersById[userId];
  var kmDistance = userRow ? Number(userRow.km_distance || 0) : 0;
  if (!isFinite(kmDistance) || kmDistance <= 0) return 0;
  var monthNames = [
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
  var parts = String(periodeStr || "")
    .trim()
    .split(" ");
  var targetMonthIndex = monthNames.indexOf(parts[0]);
  var targetYear = parseInt(parts[1], 10);
  if (targetMonthIndex === -1 || isNaN(targetYear)) return 0;
  var mm =
    targetMonthIndex + 1 < 10
      ? "0" + (targetMonthIndex + 1)
      : String(targetMonthIndex + 1);
  var firstDay = targetYear + "-" + mm + "-01";
  var lastDayDate = new Date(targetYear, targetMonthIndex + 1, 0);
  var lastDd =
    lastDayDate.getDate() < 10
      ? "0" + lastDayDate.getDate()
      : String(lastDayDate.getDate());
  var lastDay = targetYear + "-" + mm + "-" + lastDd;
  
  var userEntries = snapshot.attendanceByUserId[userId] || [];
  userEntries = userEntries.filter(function (row) {
    var d = new Date(row.date);
    if (isNaN(d.getTime())) return false;
    return d.getMonth() === targetMonthIndex && d.getFullYear() === targetYear;
  });
  
  var bestByDate = {};
  userEntries.forEach(function (entry) {
    var dateKey = String(entry.date || "");
    var isExtra = String(entry.extra_attendee || "").toLowerCase() === "true";
    var durasiJadwal;
    if (isExtra && !String(entry.sched_time_in || "").trim()) {
      durasiJadwal = _minutesBetween(
        String(entry.time_in || ""),
        String(entry.time_out || ""),
      );
    } else {
      durasiJadwal = _minutesBetween(
        String(entry.sched_time_in || ""),
        String(entry.sched_time_out || ""),
      );
    }
    if (!bestByDate[dateKey]) {
      bestByDate[dateKey] = { entry: entry, durasiJadwal: durasiJadwal };
    } else {
      if (durasiJadwal > bestByDate[dateKey].durasiJadwal) {
        bestByDate[dateKey] = { entry: entry, durasiJadwal: durasiJadwal };
      }
    }
  });
  
  var eventAttRecords = snapshot.eventAttendanceByUserId[userId] || [];
  eventAttRecords = eventAttRecords.filter(function (r) {
    return (
      String(r.date || "") >= firstDay &&
      String(r.date || "") <= lastDay
    );
  });
  
  var regularDates = Object.keys(bestByDate);
  var allUniqueDates = regularDates;
  var seenEventDates = {};
  var eventDaysCountVal = 0;
  eventAttRecords.forEach(function (r) {
    var submitted =
      r.journal_submitted === true ||
      String(r.journal_submitted).toUpperCase() === "TRUE";
    if (submitted && r.date) {
      var d = String(r.date);
      if (!seenEventDates[d]) {
        seenEventDates[d] = true;
        eventDaysCountVal++;
      }
    }
  });
  _lastEventDaysCount = eventDaysCountVal;
  if (allUniqueDates.length === 0) return 0;
  
  var total = 0;
  var userLeaves = snapshot.attendanceLeavesByUserId[userId] || [];
  userLeaves = userLeaves.filter(function (lv) {
    var d = new Date(lv.date);
    if (isNaN(d.getTime())) return false;
    return d.getMonth() === targetMonthIndex && d.getFullYear() === targetYear;
  });
  
  for (var i = 0; i < allUniqueDates.length; i++) {
    var dateKey = allUniqueDates[i];
    if (bestByDate[dateKey]) {
      var item = bestByDate[dateKey];
      var e = item.entry;
      var dJadwal = item.durasiJadwal;
      var dAktual = _minutesBetween(
        String(e.time_in || ""),
        String(e.time_out || ""),
      );
      var totalLeaveMinutes = 0;
      userLeaves.forEach(function (lv) {
        if (String(lv.date) === dateKey) {
          totalLeaveMinutes += _minutesBetween(
            String(lv.leave_time || ""),
            String(lv.return_time || ""),
          );
        }
      });
      var dAktualEfektif = dAktual - totalLeaveMinutes;
      if (dAktualEfektif < 0) dAktualEfektif = 0;
      total += _calcDailyTransport(
        kmDistance,
        tarifPerKm,
        dAktualEfektif,
        dJadwal,
        true,
      );
    }
  }
  return total;
}
function getDailyAttendanceList(token, dateStr) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var tz = Session.getScriptTimeZone();
    var targetDateStr;
    if (dateStr && /^\d{4}-\d{2}-\d{2}$/.test(String(dateStr))) {
      targetDateStr = String(dateStr);
    } else {
      targetDateStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    }
    var targetDateObj = new Date(targetDateStr + "T00:00:00");
    var targetDayIndex = (function () {
      var utc =
        targetDateObj.getTime() + targetDateObj.getTimezoneOffset() * 60000;
      var jkt = new Date(utc + 7 * 3600000);
      return jkt.getDay();
    })();
    var allUsers = getData(SHEET_NAME.USERS);
    var picketSchedules = getData(SHEET_NAME.PICKET_SCHEDULES);
    var allSubstitutes = getData(SHEET_NAME.SUBSTITUTES);
    var allSchedules = getData(SHEET_NAME.SCHEDULES);
    var allAttendance = getData(SHEET_NAME.ATTENDANCE);
    var allLeaves = getData(SHEET_NAME.ATTENDANCE_LEAVES);
    var allEventAttendance = getData(EVENT_SHEET.ATTENDANCE);
    var allAllowances = getData("Allowances");
    var allCeremonies = getData(SHEET_NAME.CEREMONY_SCHEDULES) || [];
    var cfg = _getConfigMap();
    var activeTP = cfg["tahun_pelajaran"] || "";
    var activeSem = cfg["semester"] || "";
    function findUser(uid) {
      return allUsers.find(function (u) {
        return String(u.id) === String(uid);
      });
    }
    function findAttendanceRecord(uid, dateS) {
      return (
        allAttendance.find(function (a) {
          return (
            String(a.user_id) === String(uid) &&
            String(a.date) === String(dateS)
          );
        }) || null
      );
    }
    if (!isAdmin) {
      var isPicketOriginal = picketSchedules.some(function (p) {
        return (
          String(p.user_id) === String(user.id) &&
          String(p.day_index).trim() === String(targetDayIndex)
        );
      });
      var isPicketSubstitute = allSubstitutes.some(function (s) {
        if (
          String(s.substitute_user_id) !== String(user.id) ||
          String(s.date) !== targetDateStr
        )
          return false;
        var sid = String(s.schedule_id);
        if (sid === "PICKET-DUTY" || sid === "PIKET") return true;
        return picketSchedules.some(function (p) {
          return String(p.id) === sid;
        });
      });
      var isPicketToday = isPicketOriginal || isPicketSubstitute;
      if (!isPicketToday) {
        return {
          status: "error",
          message: "Anda tidak memiliki akses ke menu ini.",
        };
      }
      var myAttendance = findAttendanceRecord(user.id, targetDateStr);
      var picketHasTimeIn =
        myAttendance &&
        myAttendance.time_in &&
        String(myAttendance.time_in).trim() !== "";
      if (!picketHasTimeIn) {
        return {
          status: "error",
          message:
            "Kehadiran piket Anda belum dikonfirmasi admin. Hubungi admin untuk konfirmasi terlebih dahulu.",
        };
      }
    }
    var teacherMap = {}; 
    function addTeacher(uid, roleType) {
      if (!uid || String(uid).trim() === "") return;
      var uidStr = String(uid);
      if (!teacherMap[uidStr]) {
        teacherMap[uidStr] = { role_type: roleType };
      } else {
        var priority = {
          piket: 5,
          pengganti: 4,
          ujian_pengawas: 3,
          ujian_panitia: 2,
          biasa: 1,
        };
        var curPrio = priority[teacherMap[uidStr].role_type] || 0;
        var newPrio = priority[roleType] || 0;
        if (newPrio > curPrio) teacherMap[uidStr].role_type = roleType;
      }
    }
    var substitutedSchedulesToday = {};
    allSubstitutes.forEach(function (s) {
      if (safeDate(s.date) === targetDateStr) {
        substitutedSchedulesToday[String(s.schedule_id)] = true;
      }
    });
    picketSchedules.forEach(function (p) {
      if (String(p.day_index).trim() === String(targetDayIndex)) {
        if (
          !substitutedSchedulesToday[String(p.id)] &&
          !substitutedSchedulesToday["PICKET-DUTY"] &&
          !substitutedSchedulesToday["PIKET"]
        ) {
          addTeacher(p.user_id, "piket");
        }
      }
    });
    allSubstitutes.forEach(function (s) {
      if (safeDate(s.date) === targetDateStr) {
        var sid = String(s.schedule_id);
        var isPicketSub = sid === "PICKET-DUTY" || sid === "PIKET";
        if (!isPicketSub) {
          isPicketSub = picketSchedules.some(function (p) {
            return String(p.id) === sid;
          });
        }
        if (isPicketSub) {
          addTeacher(s.substitute_user_id, "pengganti");
        }
      }
    });
    allSchedules.forEach(function (s) {
      var isToday = String(s.day_index).trim() === String(targetDayIndex);
      var sTP = s.tahun_pelajaran || activeTP;
      var sSem = s.semester || activeSem;
      if (isToday && sTP === activeTP && sSem === activeSem) {
        if (!substitutedSchedulesToday[String(s.id)]) {
          addTeacher(s.user_id, "biasa");
        }
      }
    });
    allSubstitutes.forEach(function (s) {
      var sid = String(s.schedule_id);
      var isPicketSub =
        sid === "PICKET-DUTY" ||
        sid === "PIKET" ||
        picketSchedules.some(function (p) {
          return String(p.id) === sid;
        });
      if (!isPicketSub && safeDate(s.date) === targetDateStr) {
        if (
          s.substitute_user_id &&
          String(s.substitute_user_id).trim() !== ""
        ) {
          addTeacher(s.substitute_user_id, "biasa");
        }
      }
    });
    var examPeriods = getData(EXAM_SHEET.PERIODS);
    var activePeriod = null;
    examPeriods.forEach(function (p) {
      var ps = String(p.date_start || "");
      var pe = String(p.date_end || "");
      if (ps && pe && targetDateStr >= ps && targetDateStr <= pe) {
        activePeriod = p;
      }
    });
    if (activePeriod) {
      var examSessions = getData(EXAM_SHEET.SESSIONS).filter(function (s) {
        return (
          String(s.period_id) === String(activePeriod.id) &&
          safeDate(s.date) === targetDateStr
        );
      });
      var todaySessionIds = {};
      examSessions.forEach(function (s) {
        todaySessionIds[String(s.id)] = true;
      });
      var examRooms = getData(EXAM_SHEET.ROOMS).filter(function (r) {
        return todaySessionIds[String(r.session_id)];
      });
      var todayRoomIds = {};
      examRooms.forEach(function (r) {
        todayRoomIds[String(r.id)] = true;
      });
      var examSupervisors = getData(EXAM_SHEET.SUPERVISORS).filter(
        function (sup) {
          return (
            todayRoomIds[String(sup.room_id)] && String(sup.status) === "active"
          );
        },
      );
      examSupervisors.forEach(function (sup) {
        addTeacher(sup.user_id, "ujian_pengawas");
      });
      var examCommittee = getData(EXAM_SHEET.COMMITTEE).filter(function (c) {
        return (
          String(c.period_id) === String(activePeriod.id) &&
          safeDate(c.date) === targetDateStr &&
          String(c.status) === "active"
        );
      });
      examCommittee.forEach(function (c) {
        addTeacher(c.user_id, "ujian_panitia");
      });
    }
    allCeremonies.forEach(function (c) {
      if (safeDate(c.date) === targetDateStr) {
        addTeacher(c.user_id, "biasa");
      }
    });
    var kepalaSekolahUserIds = {};
    allAllowances.forEach(function (a) {
      var dName = String(a.duty_name || "").toLowerCase();
      if (
        dName.indexOf("kepala madrasah") !== -1 ||
        dName.indexOf("kepala sekolah") !== -1
      ) {
        if (a.user_id && String(a.user_id).trim() !== "") {
          kepalaSekolahUserIds[String(a.user_id)] = true;
          addTeacher(a.user_id, "biasa");
        }
      }
    });
    allAttendance.forEach(function (a) {
      if (
        safeDate(a.date) === targetDateStr &&
        String(a.extra_attendee || "").toLowerCase() === "true" &&
        a.user_id &&
        String(a.user_id).trim() !== ""
      ) {
        addTeacher(a.user_id, "biasa");
      }
    });
    allEventAttendance.forEach(function (ea) {
      if (
        safeDate(ea.date) === targetDateStr &&
        ea.user_id &&
        String(ea.user_id).trim() !== ""
      ) {
        addTeacher(ea.user_id, "biasa");
      }
    });
    var picketUserIds = {}; 
    picketSchedules.forEach(function (p) {
      if (String(p.day_index).trim() === String(targetDayIndex)) {
        picketUserIds[String(p.user_id)] = true;
      }
    });
    allSubstitutes.forEach(function (s) {
      var sid = String(s.schedule_id);
      var isPicketSub =
        sid === "PICKET-DUTY" ||
        sid === "PIKET" ||
        picketSchedules.some(function (p) {
          return String(p.id) === sid;
        });
      if (isPicketSub && safeDate(s.date) === targetDateStr) {
        picketUserIds[String(s.substitute_user_id)] = true;
      }
    });
    (function () {
      try {
        var needsApply = Object.keys(teacherMap).some(function (uid) {
          var rec = allAttendance.find(function (a) {
            return (
              String(a.user_id) === uid && safeDate(a.date) === targetDateStr
            );
          });
          return !rec || !String(rec.sched_time_in || "").trim();
        });
        if (!needsApply) return; 
        var examPeriods2 = getData(EXAM_SHEET.PERIODS);
        var isExam2 = examPeriods2.some(function (p) {
          var ps = String(p.date_start || ""),
            pe = String(p.date_end || "");
          return ps && pe && targetDateStr >= ps && targetDateStr <= pe;
        });
        var activeType2 = isExam2 ? "UJIAN" : "KBM";
        var allTpl = getData(SHEET_NAME.ATTENDANCE_SCHED_TEMPLATES);
        var tpl = allTpl.find(function (t) {
          return (
            Number(t.day_index) === targetDayIndex &&
            String(t.sched_type || "")
              .toUpperCase()
              .trim() === activeType2
          );
        });
        if (!tpl) return; 
        var rTime = /^([01]\d|2[0-3]):([0-5]\d)$/;
        var schedIn = safeTime(tpl.sched_time_in).substring(0, 5);
        var schedOut = safeTime(tpl.sched_time_out).substring(0, 5);
        if (!rTime.test(schedIn) || !rTime.test(schedOut)) return; 
        var attSheet = getSheet(SHEET_NAME.ATTENDANCE);
        var attRows = attSheet.getDataRange().getValues();
        var attHdrs = attRows[0].map(function (h) {
          return String(h).toLowerCase().trim().replace(/\s+/g, "_");
        });
        var attCI = {};
        attHdrs.forEach(function (h, i) {
          attCI[h] = i;
        });
        var ensureCol2 = function (c) {
          if (attCI[c] !== undefined) return;
          var nc = attSheet.getLastColumn() + 1;
          attSheet.getRange(1, nc).setValue(c);
          attCI[c] = nc - 1;
        };
        ensureCol2("sched_time_in");
        ensureCol2("sched_time_out");
        var tz2 = Session.getScriptTimeZone();
        Object.keys(teacherMap).forEach(function (uid) {
          var foundRow = -1,
            hasSched = false,
            isExtra = false;
          for (var i = 1; i < attRows.length; i++) {
            var rowUid = String(
              attRows[i][
                attCI["user_id"] !== undefined ? attCI["user_id"] : 2
              ] || "",
            );
            var rawDate =
              attRows[i][attCI["date"] !== undefined ? attCI["date"] : 1];
            var rowDate =
              rawDate instanceof Date
                ? Utilities.formatDate(rawDate, tz2, "yyyy-MM-dd")
                : String(rawDate || "");
            if (rowUid === uid && rowDate === targetDateStr) {
              foundRow = i + 1;
              hasSched =
                String(attRows[i][attCI["sched_time_in"]] || "").trim() !== "";
              isExtra =
                String(
                  attRows[i][
                    attCI["extra_attendee"] !== undefined
                      ? attCI["extra_attendee"]
                      : -1
                  ] || "",
                ).toLowerCase() === "true";
              break;
            }
          }
          if (hasSched) return; 
          if (foundRow > 0) {
            attSheet
              .getRange(foundRow, attCI["sched_time_in"] + 1)
              .setValue(schedIn);
            attSheet
              .getRange(foundRow, attCI["sched_time_out"] + 1)
              .setValue(schedOut);
          } else {
            var nr = attHdrs.map(function () {
              return "";
            });
            if (attCI["id"] !== undefined) nr[attCI["id"]] = generateId("ATT");
            if (attCI["date"] !== undefined) nr[attCI["date"]] = targetDateStr;
            if (attCI["user_id"] !== undefined) nr[attCI["user_id"]] = uid;
            if (attCI["sched_time_in"] !== undefined)
              nr[attCI["sched_time_in"]] = schedIn;
            if (attCI["sched_time_out"] !== undefined)
              nr[attCI["sched_time_out"]] = schedOut;
            attSheet.appendRow(nr);
          }
        });
        allAttendance = getData(SHEET_NAME.ATTENDANCE);
        allLeaves = getData(SHEET_NAME.ATTENDANCE_LEAVES);
      } catch (autoApplyErr) {
        console.warn(
          "getDailyAttendanceList auto-apply warning: " + autoApplyErr,
        );
      }
    })();
    var now = new Date();
    var nowMinutes = (function () {
      var utcMs = now.getTime() + now.getTimezoneOffset() * 60000;
      var jkt = new Date(utcMs + 7 * 3600000);
      return jkt.getHours() * 60 + jkt.getMinutes();
    })();
    var list = [];
    Object.keys(teacherMap).forEach(function (uid) {
      var roleInfo = teacherMap[uid];
      var userObj = findUser(uid);
      if (!userObj) return; 
      var isKepalaSekolah = !!kepalaSekolahUserIds[uid];
      var roleType = roleInfo.role_type;
      var rec = findAttendanceRecord(uid, targetDateStr);
      var schedTimeIn = rec ? safeTime(rec.sched_time_in).substring(0, 5) : "";
      var schedTimeOut = rec
        ? safeTime(rec.sched_time_out).substring(0, 5)
        : "";
      var timeIn = rec ? safeTime(rec.time_in).substring(0, 5) : "";
      var timeOut = rec ? safeTime(rec.time_out).substring(0, 5) : "";
      var isExtraAttendee =
        rec && String(rec.extra_attendee || "").toLowerCase() === "true";
      var canInputSched = isAdmin;
      var canInputTimeIn = false;
      var timeInBlockReason = "";
      if (timeIn && timeIn !== "") {
        canInputTimeIn = false;
        timeInBlockReason = "Jam datang sudah dicatat.";
      } else if (!schedTimeIn || schedTimeIn === "") {
        if (isExtraAttendee && isAdmin) {
          canInputTimeIn = false;
          timeInBlockReason =
            "Gunakan tombol Detail untuk catat jam datang guru tambahan.";
        } else {
          canInputTimeIn = false;
          timeInBlockReason = "Jam jadwal belum diset oleh admin.";
        }
      } else {
        var isPicketEntry = !!picketUserIds[String(uid)]; 
        if (isAdmin) {
          canInputTimeIn = true;
        } else {
          if (isPicketEntry) {
            var isSelf = String(uid) === String(user.id);
            var hasTeachingScheduleToday =
              isSelf &&
              allSchedules.some(function (s) {
                var sTP = s.tahun_pelajaran || activeTP;
                var sSem = s.semester || activeSem;
                return (
                  String(s.user_id).trim() === String(user.id) &&
                  String(s.day_index).trim() === String(targetDayIndex) &&
                  sTP === activeTP &&
                  sSem === activeSem
                );
              });
            if (hasTeachingScheduleToday) {
              canInputTimeIn = true;
            } else {
              canInputTimeIn = false;
              timeInBlockReason = "Hanya admin yang dapat mencatat.";
            }
          } else {
            canInputTimeIn = true;
          }
        }
      }
      var canInputTimeOut = false;
      var timeOutBlockReason = "";
      if (!timeIn || timeIn === "") {
        canInputTimeOut = false;
        timeOutBlockReason = timeInBlockReason || "Jam datang belum dicatat.";
      } else if (timeOut && timeOut !== "") {
        canInputTimeOut = false;
        timeOutBlockReason = "Jam pulang sudah dicatat.";
      } else if (!schedTimeOut || schedTimeOut === "") {
        canInputTimeOut = false;
        timeOutBlockReason = "Jam jadwal pulang belum diset oleh admin.";
      } else {
        var schedOutMinutes = (function () {
          var m = String(schedTimeOut).match(/^([01]\d|2[0-3]):([0-5]\d)$/);
          if (!m) return -1;
          return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
        })();
        var availableFromMinutes =
          schedOutMinutes >= 0 ? schedOutMinutes + 5 : -1;
        if (availableFromMinutes < 0) {
          canInputTimeOut = false;
          timeOutBlockReason = "Format jam jadwal pulang tidak valid.";
        } else if (nowMinutes < availableFromMinutes) {
          var availHour = Math.floor(availableFromMinutes / 60);
          var availMin = availableFromMinutes % 60;
          var availStr =
            String(availHour).padStart(2, "0") +
            ":" +
            String(availMin).padStart(2, "0");
          canInputTimeOut = false;
          timeOutBlockReason = "Tersedia pada pukul " + availStr + ".";
        } else {
          var isPicketEntryOut = !!picketUserIds[String(uid)];
          if (isAdmin) {
            canInputTimeOut = true;
          } else {
            if (isPicketEntryOut) {
              var isSelfOut = String(uid) === String(user.id);
              var hasTeachingScheduleTodayOut =
                isSelfOut &&
                allSchedules.some(function (s) {
                  var sTP = s.tahun_pelajaran || activeTP;
                  var sSem = s.semester || activeSem;
                  return (
                    String(s.user_id).trim() === String(user.id) &&
                    String(s.day_index).trim() === String(targetDayIndex) &&
                    sTP === activeTP &&
                    sSem === activeSem
                  );
                });
              if (hasTeachingScheduleTodayOut) {
                canInputTimeOut = true;
              } else {
                canInputTimeOut = false;
                timeOutBlockReason = "Hanya admin yang dapat mencatat.";
              }
            } else {
              canInputTimeOut = true;
            }
          }
        }
      }
      var leavesForUser = allLeaves
        .filter(function (lv) {
          return (
            String(lv.user_id) === String(uid) &&
            String(lv.date) === targetDateStr
          );
        })
        .map(function (lv) {
          return {
            id: String(lv.id || ""),
            leave_time: String(lv.leave_time || "")
              .trim()
              .substring(0, 5),
            return_time: String(lv.return_time || "")
              .trim()
              .substring(0, 5),
            reason: String(lv.reason || ""),
          };
        });
      var canInputLeave = !!(
        timeIn &&
        timeIn !== "" &&
        (!timeOut || timeOut === "")
      );
      if (canInputLeave && !isAdmin && !!picketUserIds[String(uid)]) {
        canInputLeave = false;
      }
      list.push({
        user_id: String(uid),
        full_name: String(userObj.full_name || ""),
        role_type: roleType,
        isKepalaSekolah: isKepalaSekolah,
        is_extra_attendee: isExtraAttendee,
        extra_notes:
          isExtraAttendee && rec
            ? String(rec.notes || "")
                .replace(/^\[TAMBAH MANUAL\]\s*/, "")
                .trim()
            : "",
        sched_time_in: schedTimeIn,
        sched_time_out: schedTimeOut,
        time_in: timeIn,
        time_out: timeOut,
        canInputSched: canInputSched,
        canInputTimeIn: canInputTimeIn,
        canInputTimeOut: canInputTimeOut,
        timeInBlockReason: timeInBlockReason,
        timeOutBlockReason: timeOutBlockReason,
        canInputLeave: canInputLeave,
        leaves: leavesForUser,
      });
    });
    var rolePriority = {
      piket: 0,
      pengganti: 1,
      ujian_pengawas: 2,
      ujian_panitia: 3,
      biasa: 4,
    };
    list.sort(function (a, b) {
      var pa =
        rolePriority[a.role_type] !== undefined
          ? rolePriority[a.role_type]
          : 99;
      var pb =
        rolePriority[b.role_type] !== undefined
          ? rolePriority[b.role_type]
          : 99;
      if (pa !== pb) return pa - pb;
      return String(a.full_name).localeCompare(String(b.full_name), "id");
    });
    return {
      status: "success",
      date: targetDateStr,
      list: list,
    };
  } catch (e) {
    console.error("getDailyAttendanceList error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function saveAttendanceTimeIn(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var targetUserId = String((payload && payload.user_id) || "").trim();
    var dateStr = String((payload && payload.date) || "").trim();
    var timeIn = String((payload && payload.time_in) || "").trim();
    if (!targetUserId || !dateStr) {
      return { status: "error", message: "user_id dan date wajib diisi." };
    }
    var rTime = /^([01]\d|2[0-3]):([0-5]\d)$/;
    if (!rTime.test(timeIn)) {
      return {
        status: "error",
        message:
          "Format jam tidak valid. Gunakan format HH:mm (contoh: 07:30).",
      };
    }
    if (!isAdmin) {
      var targetDayIndex = (function () {
        try {
          var d = new Date(dateStr + "T00:00:00");
          var utc = d.getTime() + d.getTimezoneOffset() * 60000;
          return new Date(utc + 7 * 3600000).getDay();
        } catch (_) {
          return -1;
        }
      })();
      var picketSchedules = getData(SHEET_NAME.PICKET_SCHEDULES);
      var allSubstitutes = getData(SHEET_NAME.SUBSTITUTES);
      var callerIsPicketOriginal = picketSchedules.some(function (p) {
        return (
          String(p.user_id) === String(user.id) &&
          String(p.day_index).trim() === String(targetDayIndex)
        );
      });
      var callerIsPicketSubstitute = allSubstitutes.some(function (s) {
        if (
          String(s.substitute_user_id) !== String(user.id) ||
          String(s.date) !== dateStr
        )
          return false;
        var sid = String(s.schedule_id);
        if (sid === "PICKET-DUTY" || sid === "PIKET") return true;
        return picketSchedules.some(function (p) {
          return String(p.id) === sid;
        });
      });
      var callerIsPicket = callerIsPicketOriginal || callerIsPicketSubstitute;
      if (callerIsPicket) {
        var allAttendance = getData(SHEET_NAME.ATTENDANCE);
        var callerRecord = allAttendance.find(function (a) {
          return (
            String(a.user_id) === String(user.id) && String(a.date) === dateStr
          );
        });
        var callerHasTimeIn =
          callerRecord && String(callerRecord.time_in || "").trim() !== "";
        if (!callerHasTimeIn) {
          return {
            status: "error",
            message:
              "Kehadiran piket Anda belum dikonfirmasi admin. Hubungi admin untuk konfirmasi terlebih dahulu.",
          };
        }
        var targetIsPicketOriginal = picketSchedules.some(function (p) {
          return (
            String(p.user_id) === String(targetUserId) &&
            String(p.day_index).trim() === String(targetDayIndex)
          );
        });
        var targetIsPicketSubstitute = allSubstitutes.some(function (s) {
          if (
            String(s.substitute_user_id) !== String(targetUserId) ||
            String(s.date) !== dateStr
          )
            return false;
          var sid = String(s.schedule_id);
          if (sid === "PICKET-DUTY" || sid === "PIKET") return true;
          return picketSchedules.some(function (p) {
            return String(p.id) === sid;
          });
        });
        var targetIsPicket = targetIsPicketOriginal || targetIsPicketSubstitute;
        if (targetIsPicket) {
          var isSelf = String(targetUserId) === String(user.id);
          if (isSelf) {
            var cfg = _getConfigMap();
            var activeTP = cfg["tahun_pelajaran"] || "";
            var activeSem = cfg["semester"] || "";
            var allSchedules = getData(SHEET_NAME.SCHEDULES);
            var hasTeachingToday = allSchedules.some(function (s) {
              var sTP = s.tahun_pelajaran || activeTP;
              var sSem = s.semester || activeSem;
              return (
                String(s.user_id).trim() === String(user.id) &&
                String(s.day_index).trim() === String(targetDayIndex) &&
                sTP === activeTP &&
                sSem === activeSem
              );
            });
            if (!hasTeachingToday) {
              return {
                status: "error",
                message:
                  "Pencatatan jam kehadiran guru piket hanya dapat dilakukan oleh admin.",
              };
            }
          } else {
            return {
              status: "error",
              message:
                "Pencatatan jam kehadiran guru piket hanya dapat dilakukan oleh admin.",
            };
          }
        }
      } else {
        var isSelfRecord = String(targetUserId) === String(user.id);
        if (!isSelfRecord) {
          return {
            status: "error",
            message:
              "Anda tidak memiliki wewenang untuk mencatat kehadiran guru lain.",
          };
        }
      }
    }
    var sheet = getSheet(SHEET_NAME.ATTENDANCE);
    var rows = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var headers = rows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var colIdx = {};
    headers.forEach(function (h, i) {
      colIdx[h] = i;
    });
    var ensureCol = function (colName) {
      if (colIdx[colName] !== undefined) return;
      var newColNum = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColNum).setValue(colName);
      colIdx[colName] = newColNum - 1; 
    };
    ensureCol("sched_time_in");
    ensureCol("time_in");
    var foundRowNum = -1;
    var foundSchedTimeIn = "";
    for (var i = 1; i < rows.length; i++) {
      var rowUserId = String(
        rows[i][colIdx["user_id"] !== undefined ? colIdx["user_id"] : 2] || "",
      );
      var rowDateRaw =
        rows[i][colIdx["date"] !== undefined ? colIdx["date"] : 1];
      var rowDateStr =
        rowDateRaw instanceof Date
          ? Utilities.formatDate(rowDateRaw, tz, "yyyy-MM-dd")
          : String(rowDateRaw || "");
      if (rowUserId === targetUserId && rowDateStr === dateStr) {
        foundRowNum = i + 1; 
        var rawSched = rows[i][colIdx["sched_time_in"]];
        foundSchedTimeIn =
          rawSched instanceof Date
            ? Utilities.formatDate(rawSched, tz, "HH:mm")
            : String(rawSched || "")
                .trim()
                .substring(0, 5);
        break;
      }
    }
    if (!foundSchedTimeIn || !rTime.test(foundSchedTimeIn)) {
      return {
        status: "error",
        message: "Jam jadwal masuk belum diatur oleh admin.",
      };
    }
    var schedParts = foundSchedTimeIn.split(":");
    var inParts = timeIn.split(":");
    var schedMinutes =
      parseInt(schedParts[0], 10) * 60 + parseInt(schedParts[1], 10);
    var inMinutes = parseInt(inParts[0], 10) * 60 + parseInt(inParts[1], 10);
    if (inMinutes < schedMinutes) {
      return {
        status: "error",
        message:
          "Jam datang tidak boleh lebih awal dari jadwal masuk (" +
          foundSchedTimeIn +
          ").",
      };
    }
    if (foundRowNum > 0) {
      sheet.getRange(foundRowNum, colIdx["time_in"] + 1).setValue(timeIn);
    } else {
      return {
        status: "error",
        message:
          "Record kehadiran tidak ditemukan. Muat ulang halaman dan coba lagi.",
      };
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return { status: "success", message: "Jam datang berhasil dicatat." };
  } catch (e) {
    console.error("saveAttendanceTimeIn error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function saveAttendanceTimeOut(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var targetUserId = String((payload && payload.user_id) || "").trim();
    var dateStr = String((payload && payload.date) || "").trim();
    var timeOut = String((payload && payload.time_out) || "").trim();
    if (!targetUserId || !dateStr) {
      return { status: "error", message: "user_id dan date wajib diisi." };
    }
    var rTime = /^([01]\d|2[0-3]):([0-5]\d)$/;
    if (!rTime.test(timeOut)) {
      return {
        status: "error",
        message:
          "Format jam tidak valid. Gunakan format HH:mm (contoh: 14:30).",
      };
    }
    if (!isAdmin) {
      var targetDayIndex = (function () {
        try {
          var d = new Date(dateStr + "T00:00:00");
          var utc = d.getTime() + d.getTimezoneOffset() * 60000;
          return new Date(utc + 7 * 3600000).getDay();
        } catch (_) {
          return -1;
        }
      })();
      var picketSchedules = getData(SHEET_NAME.PICKET_SCHEDULES);
      var allSubstitutes = getData(SHEET_NAME.SUBSTITUTES);
      var callerIsPicketOriginal = picketSchedules.some(function (p) {
        return (
          String(p.user_id) === String(user.id) &&
          String(p.day_index).trim() === String(targetDayIndex)
        );
      });
      var callerIsPicketSubstitute = allSubstitutes.some(function (s) {
        if (
          String(s.substitute_user_id) !== String(user.id) ||
          String(s.date) !== dateStr
        )
          return false;
        var sid = String(s.schedule_id);
        if (sid === "PICKET-DUTY" || sid === "PIKET") return true;
        return picketSchedules.some(function (p) {
          return String(p.id) === sid;
        });
      });
      var callerIsPicket = callerIsPicketOriginal || callerIsPicketSubstitute;
      if (callerIsPicket) {
        var allAttendanceCheck = getData(SHEET_NAME.ATTENDANCE);
        var callerRecord = allAttendanceCheck.find(function (a) {
          return (
            String(a.user_id) === String(user.id) && String(a.date) === dateStr
          );
        });
        var callerHasTimeIn =
          callerRecord && String(callerRecord.time_in || "").trim() !== "";
        if (!callerHasTimeIn) {
          return {
            status: "error",
            message:
              "Kehadiran piket Anda belum dikonfirmasi admin. Hubungi admin untuk konfirmasi terlebih dahulu.",
          };
        }
        var targetIsPicketOriginal = picketSchedules.some(function (p) {
          return (
            String(p.user_id) === String(targetUserId) &&
            String(p.day_index).trim() === String(targetDayIndex)
          );
        });
        var targetIsPicketSubstitute = allSubstitutes.some(function (s) {
          if (
            String(s.substitute_user_id) !== String(targetUserId) ||
            String(s.date) !== dateStr
          )
            return false;
          var sid = String(s.schedule_id);
          if (sid === "PICKET-DUTY" || sid === "PIKET") return true;
          return picketSchedules.some(function (p) {
            return String(p.id) === sid;
          });
        });
        var targetIsPicket = targetIsPicketOriginal || targetIsPicketSubstitute;
        if (targetIsPicket) {
          var isSelf = String(targetUserId) === String(user.id);
          if (isSelf) {
            var cfg = _getConfigMap();
            var activeTP = cfg["tahun_pelajaran"] || "";
            var activeSem = cfg["semester"] || "";
            var allSchedulesCheck = getData(SHEET_NAME.SCHEDULES);
            var hasTeachingToday = allSchedulesCheck.some(function (s) {
              var sTP = s.tahun_pelajaran || activeTP;
              var sSem = s.semester || activeSem;
              return (
                String(s.user_id).trim() === String(user.id) &&
                String(s.day_index).trim() === String(targetDayIndex) &&
                sTP === activeTP &&
                sSem === activeSem
              );
            });
            if (!hasTeachingToday) {
              return {
                status: "error",
                message:
                  "Pencatatan jam kehadiran guru piket hanya dapat dilakukan oleh admin.",
              };
            }
          } else {
            return {
              status: "error",
              message:
                "Pencatatan jam kehadiran guru piket hanya dapat dilakukan oleh admin.",
            };
          }
        }
      } else {
        var isSelfRecord = String(targetUserId) === String(user.id);
        if (!isSelfRecord) {
          return {
            status: "error",
            message:
              "Anda tidak memiliki wewenang untuk mencatat kehadiran guru lain.",
          };
        }
      }
    }
    var sheet = getSheet(SHEET_NAME.ATTENDANCE);
    var rows = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var headers = rows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var colIdx = {};
    headers.forEach(function (h, i) {
      colIdx[h] = i;
    });
    var ensureCol = function (colName) {
      if (colIdx[colName] !== undefined) return;
      var newColNum = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColNum).setValue(colName);
      colIdx[colName] = newColNum - 1; 
    };
    ensureCol("sched_time_out");
    ensureCol("time_in");
    ensureCol("time_out");
    var foundRowNum = -1;
    var foundTimeIn = "";
    var foundSchedTimeOut = "";
    for (var i = 1; i < rows.length; i++) {
      var rowUserId = String(
        rows[i][colIdx["user_id"] !== undefined ? colIdx["user_id"] : 2] || "",
      );
      var rowDateRaw =
        rows[i][colIdx["date"] !== undefined ? colIdx["date"] : 1];
      var rowDateStr =
        rowDateRaw instanceof Date
          ? Utilities.formatDate(rowDateRaw, tz, "yyyy-MM-dd")
          : String(rowDateRaw || "");
      if (rowUserId === targetUserId && rowDateStr === dateStr) {
        foundRowNum = i + 1; 
        var rawTimeIn2 = rows[i][colIdx["time_in"]];
        var rawSchedOut = rows[i][colIdx["sched_time_out"]];
        foundTimeIn =
          rawTimeIn2 instanceof Date
            ? Utilities.formatDate(rawTimeIn2, tz, "HH:mm")
            : String(rawTimeIn2 || "")
                .trim()
                .substring(0, 5);
        foundSchedTimeOut =
          rawSchedOut instanceof Date
            ? Utilities.formatDate(rawSchedOut, tz, "HH:mm")
            : String(rawSchedOut || "")
                .trim()
                .substring(0, 5);
        break;
      }
    }
    if (!foundTimeIn || !rTime.test(foundTimeIn)) {
      return {
        status: "error",
        message: "Jam datang belum dicatat. Catat jam datang terlebih dahulu.",
      };
    }
    if (!foundSchedTimeOut || !rTime.test(foundSchedTimeOut)) {
      return {
        status: "error",
        message: "Jam jadwal pulang belum diatur oleh admin.",
      };
    }
    var schedOutParts = foundSchedTimeOut.split(":");
    var schedOutMinutes =
      parseInt(schedOutParts[0], 10) * 60 + parseInt(schedOutParts[1], 10);
    var availableFromMinutes = schedOutMinutes + 5;
    var nowJkt = (function () {
      var now = new Date();
      var utcMs = now.getTime() + now.getTimezoneOffset() * 60000;
      var jkt = new Date(utcMs + 7 * 3600000);
      return jkt.getHours() * 60 + jkt.getMinutes();
    })();
    if (!isAdmin && nowJkt < availableFromMinutes) {
      var availHour = Math.floor(availableFromMinutes / 60);
      var availMin = availableFromMinutes % 60;
      var availStr =
        String(availHour).padStart(2, "0") +
        ":" +
        String(availMin).padStart(2, "0");
      return {
        status: "error",
        message: "Input jam pulang tersedia pada pukul " + availStr + ".",
      };
    }
    var outParts = timeOut.split(":");
    var outMinutes = parseInt(outParts[0], 10) * 60 + parseInt(outParts[1], 10);
    if (outMinutes > schedOutMinutes) {
      return {
        status: "error",
        message:
          "Jam pulang tidak boleh melebihi jadwal pulang (" +
          foundSchedTimeOut +
          ").",
      };
    }
    if (foundRowNum > 0) {
      sheet.getRange(foundRowNum, colIdx["time_out"] + 1).setValue(timeOut);
    } else {
      var newRowLength = headers.length;
      var newRow = [];
      for (var j = 0; j < newRowLength; j++) newRow.push("");
      if (colIdx["id"] !== undefined) newRow[colIdx["id"]] = generateId("ATT");
      if (colIdx["date"] !== undefined) newRow[colIdx["date"]] = dateStr;
      if (colIdx["user_id"] !== undefined)
        newRow[colIdx["user_id"]] = targetUserId;
      if (colIdx["status"] !== undefined) newRow[colIdx["status"]] = "";
      if (colIdx["time_out"] !== undefined)
        newRow[colIdx["time_out"]] = timeOut;
      sheet.appendRow(newRow);
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return { status: "success", message: "Jam pulang berhasil dicatat." };
  } catch (e) {
    console.error("saveAttendanceTimeOut error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function getAttendanceRecord(token, targetUserId, dateStr) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    if (!isAdmin) {
      return {
        status: "error",
        message: "Hanya admin yang dapat memuat detail.",
      };
    }
    var sheet = getSheet(SHEET_NAME.ATTENDANCE);
    var rows = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var headers = rows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var colIdx = {};
    headers.forEach(function (h, i) {
      colIdx[h] = i;
    });
    var result = { time_in: "", time_out: "", status: "" };
    for (var i = 1; i < rows.length; i++) {
      var rowUserId = String(
        rows[i][colIdx["user_id"] !== undefined ? colIdx["user_id"] : 2] || "",
      );
      var rowDateRaw =
        rows[i][colIdx["date"] !== undefined ? colIdx["date"] : 1];
      var rowDateStr =
        rowDateRaw instanceof Date
          ? Utilities.formatDate(rowDateRaw, tz, "yyyy-MM-dd")
          : String(rowDateRaw || "");
      if (rowUserId === String(targetUserId) && rowDateStr === dateStr) {
        var rawTimeIn = rows[i][colIdx["time_in"]];
        var rawTimeOut = rows[i][colIdx["time_out"]];
        result.time_in =
          rawTimeIn instanceof Date
            ? Utilities.formatDate(rawTimeIn, tz, "HH:mm")
            : String(rawTimeIn || "")
                .trim()
                .substring(0, 5);
        result.time_out =
          rawTimeOut instanceof Date
            ? Utilities.formatDate(rawTimeOut, tz, "HH:mm")
            : String(rawTimeOut || "")
                .trim()
                .substring(0, 5);
        result.status = String(rows[i][colIdx["status"]] || "").trim();
        break;
      }
    }
    return { status: "success", data: result };
  } catch (e) {
    console.error("getAttendanceRecord error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function getMyAttendanceHistory(token, month, year) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var tz = Session.getScriptTimeZone();
    var m = parseInt(month, 10);
    var y = parseInt(year, 10);
    if (isNaN(m) || m < 1 || m > 12 || isNaN(y)) {
      return { status: "error", message: "Parameter bulan/tahun tidak valid." };
    }
    var firstDay = new Date(y, m - 1, 1);
    var lastDay = new Date(y, m, 0); 
    var today = new Date();
    today.setHours(23, 59, 59, 999);
    var upperBound = lastDay < today ? lastDay : today;
    var allAttendance = getData(SHEET_NAME.ATTENDANCE);
    var allLeaves = getData(SHEET_NAME.ATTENDANCE_LEAVES);
    var allSchedules = getData(SHEET_NAME.SCHEDULES);
    var picketSchedules = getData(SHEET_NAME.PICKET_SCHEDULES);
    var allSubstitutes = getData(SHEET_NAME.SUBSTITUTES);
    var allUsers = getData(SHEET_NAME.USERS);
    var allCalendar = getData(SHEET_NAME.CALENDAR);
    var allCeremonies = getData(SHEET_NAME.CEREMONY_SCHEDULES) || [];
    var examCommittee = getData(EXAM_SHEET.COMMITTEE) || [];
    var examSupervisors = getData(EXAM_SHEET.SUPERVISORS) || [];
    var examRooms = getData(EXAM_SHEET.ROOMS) || [];
    var examSessions = getData(EXAM_SHEET.SESSIONS) || [];
    var holidayMap = {};
    var calendarMap = {};
    allCalendar.forEach(function (h) {
      var hDateStr;
      try {
        hDateStr =
          h.date instanceof Date
            ? Utilities.formatDate(h.date, tz, "yyyy-MM-dd")
            : String(h.date || "").substring(0, 10);
      } catch (_) {
        return;
      }
      if (!hDateStr) return;
      var desc = String(h.description || "Event Kalender").trim();
      calendarMap[hDateStr] = desc; 
      if (String(h.is_holiday || "").toLowerCase() === "true") {
        holidayMap[hDateStr] = desc; 
      }
    });
    var cfg = _getConfigMap();
    var activeTP = cfg["tahun_pelajaran"] || "";
    var activeSem = cfg["semester"] || "";
    var userRow = allUsers.find(function (u) {
      return String(u.id) === String(user.id);
    });
    var isKepsek =
      userRow &&
      String(userRow.additional_role || "")
        .trim()
        .toUpperCase() === "KEPALA_SEKOLAH";
    var role = String(userRow ? userRow.role : user.role || "").toLowerCase();
    var transportEnabled = String(cfg["transport_allowance_enabled"] || "false").toLowerCase().trim() === "true";
    var tarifPerKm = Number(cfg["tarif_per_km"] || 0);
    var kmDistance = Number(userRow ? userRow.km_distance : user.km_distance) || 0;
    var myAttendance = allAttendance.filter(function (a) {
      return String(a.user_id) === String(user.id);
    });
    var myCeremonyDates = {};
    allCeremonies.forEach(function (c) {
      if (String(c.user_id) === String(user.id)) {
        myCeremonyDates[safeDate(c.date)] = true;
      }
    });
    var myExamCommDates = {};
    examCommittee.forEach(function (c) {
      if (String(c.user_id) === String(user.id) && String(c.status) !== "CANCELLED") {
        myExamCommDates[safeDate(c.date)] = true;
      }
    });
    var myExamSupvDates = {};
    examSupervisors.forEach(function (sup) {
      if (String(sup.user_id) === String(user.id) && String(sup.status) !== "CANCELLED") {
        var room = examRooms.find(function(r) { return String(r.id) === String(sup.room_id); });
        if (room) {
          var sess = examSessions.find(function(se) { return String(se.id) === String(room.session_id); });
          if (sess) {
            myExamSupvDates[safeDate(sess.date)] = true;
          }
        }
      }
    });
    var myLeaves = allLeaves.filter(function (lv) {
      return String(lv.user_id) === String(user.id);
    });
    function _hasScheduleOnDate(dateStr, dayIndex) {
      if (isKepsek) return true;
      var hasPiket = picketSchedules.some(function (p) {
        return (
          String(p.user_id) === String(user.id) &&
          String(p.day_index).trim() === String(dayIndex)
        );
      });
      if (hasPiket) return true;
      var hasPiketSub = allSubstitutes.some(function (s) {
        return (
          String(s.substitute_user_id) === String(user.id) &&
          String(s.schedule_id) === "PICKET-DUTY" &&
          safeDate(s.date) === dateStr
        );
      });
      if (hasPiketSub) return true;
      var hasMengajar = allSchedules.some(function (s) {
        if (String(s.user_id) !== String(user.id)) return false;
        if (String(s.day_index).trim() !== String(dayIndex)) return false;
        var sTP = s.tahun_pelajaran || activeTP;
        var sSem = s.semester || activeSem;
        return sTP === activeTP && sSem === activeSem;
      });
      if (hasMengajar) return true;
      var hasSub = allSubstitutes.some(function (s) {
        return (
          String(s.substitute_user_id) === String(user.id) &&
          String(s.schedule_id) !== "PICKET-DUTY" &&
          safeDate(s.date) === dateStr
        );
      });
      if (hasSub) return true;
      if (myCeremonyDates[dateStr]) return true;
      if (myExamCommDates[dateStr]) return true;
      if (myExamSupvDates[dateStr]) return true;
      return false;
    }
    var days = [];
    var totalHadir = 0,
      totalTidakHadir = 0,
      totalTerlambat = 0;
    var monthNames = [
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
    for (var day = 1; day <= lastDay.getDate(); day++) {
      var iterDate = new Date(y, m - 1, day);
      if (iterDate > upperBound) break;
      var dateStr = Utilities.formatDate(iterDate, tz, "yyyy-MM-dd");
      var dayIndex = (function () {
        var utc = iterDate.getTime() + iterDate.getTimezoneOffset() * 60000;
        return new Date(utc + 7 * 3600000).getDay();
      })();
      var dowNames = [
        "Minggu",
        "Senin",
        "Selasa",
        "Rabu",
        "Kamis",
        "Jumat",
        "Sabtu",
      ];
      var dow = dowNames[iterDate.getDay()];
      var dateFormatted = day + " " + monthNames[iterDate.getMonth()] + " " + y;
      var rec = myAttendance.find(function (a) {
        return safeDate(a.date) === dateStr;
      });
      var hasTimeIn = rec && String(rec.time_in || "").trim() !== "";
      var hasTimeOut = rec && String(rec.time_out || "").trim() !== "";
      if (!hasTimeIn && !hasTimeOut) {
        var holidayDesc = holidayMap[dateStr] || null; 
        var calendarDesc = calendarMap[dateStr] || ""; 
        var isHoliday = !!holidayDesc;
        var hasCalendar = !!calendarDesc;
        if (hasCalendar) {
          days.push({
            date: dateStr,
            date_formatted: dateFormatted,
            dow: dow,
            no_record: true,
            has_schedule: false,
            is_holiday: isHoliday,
            holiday_desc: holidayDesc || "",
            calendar_desc: calendarDesc,
            time_in: "",
            time_out: "",
            sched_time_in: "",
            sched_time_out: "",
            status: "",
            leaves: [],
          });
          continue;
        }
        var hasSchedule = _hasScheduleOnDate(dateStr, dayIndex);
        if (hasSchedule) totalTidakHadir++;
        days.push({
          date: dateStr,
          date_formatted: dateFormatted,
          dow: dow,
          no_record: true,
          has_schedule: hasSchedule,
          is_holiday: false,
          holiday_desc: "",
          calendar_desc: "",
          time_in: "",
          time_out: "",
          sched_time_in: "",
          sched_time_out: "",
          status: "",
          leaves: [],
        });
        continue;
      }
      var timeIn = safeTime(rec.time_in).substring(0, 5);
      var timeOut = safeTime(rec.time_out).substring(0, 5);
      var schedTimeIn = safeTime(rec.sched_time_in).substring(0, 5);
      var status = String(rec.status || "").trim();
      totalHadir++;
      if (timeIn && schedTimeIn && timeIn > schedTimeIn) totalTerlambat++;
      var dayLeaves = myLeaves
        .filter(function (lv) {
          return safeDate(lv.date) === dateStr;
        })
        .map(function (lv) {
          return {
            leave_time: safeTime(lv.leave_time).substring(0, 5),
            return_time: safeTime(lv.return_time).substring(0, 5),
            reason: String(lv.reason || "").trim(),
          };
        });

      var schedTimeOut = safeTime(rec.sched_time_out).substring(0, 5);
      var transportNominal = 0;
      var transportPercentage = 0;
      if (role === "guru" && transportEnabled && timeIn && timeOut) {
        var dAktual = _minutesBetween(timeIn, timeOut);
        var isExtra = String(rec.extra_attendee || "").toLowerCase() === "true";
        var dJadwal = (isExtra && !schedTimeIn) ? dAktual : _minutesBetween(schedTimeIn, schedTimeOut);
        
        var totalLeaveMinutes = 0;
        dayLeaves.forEach(function(lv) {
           totalLeaveMinutes += _minutesBetween(lv.leave_time, lv.return_time);
        });
        
        var dAktualEfektif = Math.max(0, dAktual - totalLeaveMinutes);
        transportPercentage = dJadwal > 0 ? (dAktualEfektif / dJadwal * 100) : 0;
        if (transportPercentage > 100) transportPercentage = 100;
        
        transportNominal = kmDistance * tarifPerKm * (transportPercentage / 100);
      }

      days.push({
        date: dateStr,
        date_formatted: dateFormatted,
        dow: dow,
        no_record: false,
        has_schedule: true,
        is_holiday: !!holidayMap[dateStr],
        holiday_desc: holidayMap[dateStr] || "",
        calendar_desc: calendarMap[dateStr] || "",
        time_in: timeIn,
        time_out: timeOut,
        sched_time_in: schedTimeIn,
        sched_time_out: schedTimeOut,
        status: status,
        leaves: dayLeaves,
        transport_nominal: transportNominal,
        transport_percentage: transportPercentage
      });
    }
    return {
      status: "success",
      days: days,
      summary: {
        total_days: days.length,
        hadir: totalHadir,
        tidak_hadir: totalTidakHadir,
        terlambat: totalTerlambat,
        libur: days.filter(function (d) {
          return !!d.calendar_desc && d.no_record;
        }).length,
      },
    };
  } catch (e) {
    console.error("getMyAttendanceHistory error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function addExtraAttendee(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var tz = Session.getScriptTimeZone();
    var targetUserId = String((payload && payload.target_user_id) || "").trim();
    var dateStr = String((payload && payload.date) || "").trim();
    var reason = String((payload && payload.reason) || "").trim();
    if (!targetUserId || !dateStr) {
      return {
        status: "error",
        message: "target_user_id dan date wajib diisi.",
      };
    }
    if (!reason) {
      return { status: "error", message: "Alasan penambahan wajib diisi." };
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      return { status: "error", message: "Format tanggal tidak valid." };
    }
    if (!isAdmin) {
      var picketSchedules = getData(SHEET_NAME.PICKET_SCHEDULES);
      var allSubstitutes = getData(SHEET_NAME.SUBSTITUTES);
      var allAttendance = getData(SHEET_NAME.ATTENDANCE);
      var targetDateObj = new Date(dateStr + "T00:00:00");
      var targetDayIndex = (function () {
        var utc =
          targetDateObj.getTime() + targetDateObj.getTimezoneOffset() * 60000;
        return new Date(utc + 7 * 3600000).getDay();
      })();
      var isPicketOriginal = picketSchedules.some(function (p) {
        return (
          String(p.user_id) === String(user.id) &&
          String(p.day_index).trim() === String(targetDayIndex)
        );
      });
      var isPicketSub = allSubstitutes.some(function (s) {
        return (
          String(s.substitute_user_id) === String(user.id) &&
          String(s.schedule_id) === "PICKET-DUTY" &&
          String(s.date) === dateStr
        );
      });
      if (!isPicketOriginal && !isPicketSub) {
        return {
          status: "error",
          message: "Hanya admin atau guru piket yang dapat menambahkan guru.",
        };
      }
      var myRec = allAttendance.find(function (a) {
        return (
          String(a.user_id) === String(user.id) && String(a.date) === dateStr
        );
      });
      if (!myRec || !String(myRec.time_in || "").trim()) {
        return {
          status: "error",
          message: "Kehadiran piket Anda belum dikonfirmasi admin.",
        };
      }
    }
    var targetUser = findData(SHEET_NAME.USERS, "id", targetUserId);
    if (!targetUser) {
      return { status: "error", message: "Guru yang dipilih tidak ditemukan." };
    }
    if (String(targetUser.role).toLowerCase() === "admin") {
      return {
        status: "error",
        message: "Akun admin tidak dapat ditambahkan ke daftar kehadiran guru.",
      };
    }
    var attSheet = getSheet(SHEET_NAME.ATTENDANCE);
    var attRows = attSheet.getDataRange().getValues();
    var attHdrs = attRows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var attCI = {};
    attHdrs.forEach(function (h, i) {
      attCI[h] = i;
    });
    var ensureCol = function (col) {
      if (attCI[col] !== undefined) return;
      var nc = attSheet.getLastColumn() + 1;
      attSheet.getRange(1, nc).setValue(col);
      attCI[col] = nc - 1;
      attHdrs.push(col);
    };
    ensureCol("notes");
    ensureCol("extra_attendee");
    var foundRowNum = -1;
    var existingRecord = null;
    for (var i = 1; i < attRows.length; i++) {
      var rowUid = String(
        attRows[i][attCI["user_id"] !== undefined ? attCI["user_id"] : 2] || "",
      );
      var rowDateRaw =
        attRows[i][attCI["date"] !== undefined ? attCI["date"] : 1];
      var rowDateStr =
        rowDateRaw instanceof Date
          ? Utilities.formatDate(rowDateRaw, tz, "yyyy-MM-dd")
          : String(rowDateRaw || "");
      if (rowUid === targetUserId && rowDateStr === dateStr) {
        foundRowNum = i + 1;
        existingRecord = attRows[i];
        break;
      }
    }
    if (existingRecord) {
      var existingSchedIn = String(
        existingRecord[attCI["sched_time_in"]] || "",
      ).trim();
      if (existingSchedIn) {
        return {
          status: "error",
          message:
            String(targetUser.full_name) +
            " sudah memiliki jadwal resmi hari ini dan sudah ada di Daftar Kehadiran.",
        };
      }
    }
    var targetDateObj2 = new Date(dateStr + "T00:00:00");
    var targetDayIndex2 = (function () {
      var utc =
        targetDateObj2.getTime() + targetDateObj2.getTimezoneOffset() * 60000;
      return new Date(utc + 7 * 3600000).getDay();
    })();
    var examPeriods3 = getData(EXAM_SHEET.PERIODS);
    var isExamDay = examPeriods3.some(function (p) {
      var ps = String(p.date_start || ""),
        pe = String(p.date_end || "");
      return ps && pe && dateStr >= ps && dateStr <= pe;
    });
    var activeType3 = isExamDay ? "UJIAN" : "KBM";
    var allTpl3 = getData(SHEET_NAME.ATTENDANCE_SCHED_TEMPLATES);
    var tpl3 = allTpl3.find(function (t) {
      return (
        Number(t.day_index) === targetDayIndex2 &&
        String(t.sched_type || "")
          .toUpperCase()
          .trim() === activeType3
      );
    });
    var rTime3 = /^([01]\d|2[0-3]):([0-5]\d)$/;
    var tplSchedIn = tpl3
      ? String(tpl3.sched_time_in || "")
          .trim()
          .substring(0, 5)
      : "";
    var tplSchedOut = tpl3
      ? String(tpl3.sched_time_out || "")
          .trim()
          .substring(0, 5)
      : "";
    if (!rTime3.test(tplSchedIn)) tplSchedIn = "";
    if (!rTime3.test(tplSchedOut)) tplSchedOut = "";
    ensureCol("sched_time_in");
    ensureCol("sched_time_out");
    var noteValue = "[TAMBAH MANUAL] " + reason.substring(0, 250);
    if (foundRowNum > 0) {
      if (attCI["notes"] !== undefined)
        attSheet.getRange(foundRowNum, attCI["notes"] + 1).setValue(noteValue);
      if (attCI["extra_attendee"] !== undefined)
        attSheet
          .getRange(foundRowNum, attCI["extra_attendee"] + 1)
          .setValue("true");
      if (tplSchedIn && attCI["sched_time_in"] !== undefined)
        attSheet
          .getRange(foundRowNum, attCI["sched_time_in"] + 1)
          .setValue(tplSchedIn);
      if (tplSchedOut && attCI["sched_time_out"] !== undefined)
        attSheet
          .getRange(foundRowNum, attCI["sched_time_out"] + 1)
          .setValue(tplSchedOut);
    } else {
      var newRow = attHdrs.map(function () {
        return "";
      });
      if (attCI["id"] !== undefined) newRow[attCI["id"]] = generateId("ATT");
      if (attCI["date"] !== undefined) newRow[attCI["date"]] = dateStr;
      if (attCI["user_id"] !== undefined)
        newRow[attCI["user_id"]] = targetUserId;
      if (attCI["notes"] !== undefined) newRow[attCI["notes"]] = noteValue;
      if (attCI["extra_attendee"] !== undefined)
        newRow[attCI["extra_attendee"]] = "true";
      if (tplSchedIn && attCI["sched_time_in"] !== undefined)
        newRow[attCI["sched_time_in"]] = tplSchedIn;
      if (tplSchedOut && attCI["sched_time_out"] !== undefined)
        newRow[attCI["sched_time_out"]] = tplSchedOut;
      attSheet.appendRow(newRow);
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message:
        String(targetUser.full_name) +
        " berhasil ditambahkan ke Daftar Kehadiran.",
    };
  } catch (e) {
    console.error("addExtraAttendee error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function saveAttendanceManual(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    if (!isAdmin) {
      return {
        status: "error",
        message: "Hanya admin yang dapat menyimpan detail.",
      };
    }
    var targetUserId = String((payload && payload.user_id) || "").trim();
    var dateStr = String((payload && payload.date) || "").trim();
    var timeIn = String((payload && payload.time_in) || "").trim();
    var timeOut = String((payload && payload.time_out) || "").trim();
    var attStatus = String((payload && payload.status) || "").trim();
    if (!targetUserId || !dateStr) {
      return { status: "error", message: "user_id dan date wajib diisi." };
    }
    var rTime = /^([01]\d|2[0-3]):([0-5]\d)$/;
    if (timeIn && !rTime.test(timeIn)) {
      return {
        status: "error",
        message:
          "Format jam datang tidak valid. Gunakan HH:mm (contoh: 07:30).",
      };
    }
    if (timeOut && !rTime.test(timeOut)) {
      return {
        status: "error",
        message:
          "Format jam pulang tidak valid. Gunakan HH:mm (contoh: 14:30).",
      };
    }
    if (timeIn && timeOut) {
      var inMins =
        parseInt(timeIn.split(":")[0], 10) * 60 +
        parseInt(timeIn.split(":")[1], 10);
      var outMins =
        parseInt(timeOut.split(":")[0], 10) * 60 +
        parseInt(timeOut.split(":")[1], 10);
      if (outMins < inMins) {
        return {
          status: "error",
          message: "Jam pulang tidak boleh lebih awal dari jam datang.",
        };
      }
    }
    var sheet = getSheet(SHEET_NAME.ATTENDANCE);
    var rows = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var headers = rows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var colIdx = {};
    headers.forEach(function (h, i) {
      colIdx[h] = i;
    });
    var ensureCol = function (colName) {
      if (colIdx[colName] !== undefined) return;
      var newColNum = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColNum).setValue(colName);
      colIdx[colName] = newColNum - 1;
    };
    ensureCol("time_in");
    ensureCol("time_out");
    ensureCol("status");
    var foundRowNum = -1;
    for (var i = 1; i < rows.length; i++) {
      var rowUserId = String(
        rows[i][colIdx["user_id"] !== undefined ? colIdx["user_id"] : 2] || "",
      );
      var rowDateRaw =
        rows[i][colIdx["date"] !== undefined ? colIdx["date"] : 1];
      var rowDateStr =
        rowDateRaw instanceof Date
          ? Utilities.formatDate(rowDateRaw, tz, "yyyy-MM-dd")
          : String(rowDateRaw || "");
      if (rowUserId === String(targetUserId) && rowDateStr === dateStr) {
        foundRowNum = i + 1;
        break;
      }
    }
    if (foundRowNum > 0) {
      sheet.getRange(foundRowNum, colIdx["time_in"] + 1).setValue(timeIn);
      sheet.getRange(foundRowNum, colIdx["time_out"] + 1).setValue(timeOut);
      sheet.getRange(foundRowNum, colIdx["status"] + 1).setValue(attStatus);
    } else {
      var newRowLength = headers.length;
      var newRow = [];
      for (var j = 0; j < newRowLength; j++) newRow.push("");
      if (colIdx["id"] !== undefined) newRow[colIdx["id"]] = generateId("ATT");
      if (colIdx["date"] !== undefined) newRow[colIdx["date"]] = dateStr;
      if (colIdx["user_id"] !== undefined)
        newRow[colIdx["user_id"]] = targetUserId;
      if (colIdx["status"] !== undefined) newRow[colIdx["status"]] = attStatus;
      if (colIdx["time_in"] !== undefined) newRow[colIdx["time_in"]] = timeIn;
      if (colIdx["time_out"] !== undefined)
        newRow[colIdx["time_out"]] = timeOut;
      sheet.appendRow(newRow);
    }
    return { status: "success", message: "Data kehadiran berhasil disimpan." };
  } catch (e) {
    console.error("saveAttendanceManual error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function getAttendanceSchedTemplates(token, dateStr) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var sheet = getSheet(SHEET_NAME.ATTENDANCE_SCHED_TEMPLATES);
    var rows = sheet.getDataRange().getValues();
    var result = [];
    if (rows.length >= 2) {
      var headers = rows[0].map(function (h) {
        return String(h).toLowerCase().trim().replace(/\s+/g, "_");
      });
      for (var i = 1; i < rows.length; i++) {
        var obj = {};
        headers.forEach(function (h, idx) {
          obj[h] = rows[i][idx];
        });
        result.push({
          id: String(obj.id || ""),
          day_index: Number(obj.day_index),
          sched_type: String(obj.sched_type || "KBM")
            .toUpperCase()
            .trim(),
          sched_time_in: String(obj.sched_time_in || ""),
          sched_time_out: String(obj.sched_time_out || ""),
        });
      }
    }
    result.sort(function (a, b) {
      if (a.sched_type !== b.sched_type)
        return a.sched_type < b.sched_type ? -1 : 1;
      return a.day_index - b.day_index;
    });
    var targetDate =
      dateStr && /^\d{4}-\d{2}-\d{2}$/.test(String(dateStr))
        ? String(dateStr)
        : Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
          );
    var isExamPeriod = getData(EXAM_SHEET.PERIODS).some(function (p) {
      var ps = String(p.date_start || ""),
        pe = String(p.date_end || "");
      return ps && pe && targetDate >= ps && targetDate <= pe;
    });
    return {
      status: "success",
      data: result,
      is_exam_period: isExamPeriod,
      active_type: isExamPeriod ? "UJIAN" : "KBM",
    };
  } catch (e) {
    console.error("getAttendanceSchedTemplates error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function saveAttendanceSchedTemplate(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    if (String(user.role).toLowerCase() !== "admin")
      return {
        status: "error",
        message: "Hanya Admin yang dapat mengatur jam jadwal.",
      };
    var dayIndex = parseInt(String((payload && payload.day_index) || ""), 10);
    var schedType = String((payload && payload.sched_type) || "")
      .toUpperCase()
      .trim();
    var schedTimeIn = String((payload && payload.sched_time_in) || "").trim();
    var schedTimeOut = String((payload && payload.sched_time_out) || "").trim();
    var applyDate = String((payload && payload.apply_date) || "").trim();
    if (isNaN(dayIndex) || dayIndex < 1 || dayIndex > 6)
      return { status: "error", message: "Hari tidak valid." };
    if (schedType !== "KBM" && schedType !== "UJIAN")
      return { status: "error", message: "Tipe jadwal tidak valid." };
    var rTime = /^([01]\d|2[0-3]):([0-5]\d)$/;
    if (!rTime.test(schedTimeIn) || !rTime.test(schedTimeOut))
      return {
        status: "error",
        message: "Format jam tidak valid. Gunakan HH:mm.",
      };
    var inMins =
      parseInt(schedTimeIn.split(":")[0], 10) * 60 +
      parseInt(schedTimeIn.split(":")[1], 10);
    var outMins =
      parseInt(schedTimeOut.split(":")[0], 10) * 60 +
      parseInt(schedTimeOut.split(":")[1], 10);
    if (outMins <= inMins)
      return {
        status: "error",
        message: "Jam pulang harus lebih akhir dari jam masuk.",
      };
    var sheet = getSheet(SHEET_NAME.ATTENDANCE_SCHED_TEMPLATES);
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var colIdx = {};
    headers.forEach(function (h, i) {
      colIdx[h] = i;
    });
    if (colIdx["sched_type"] === undefined) {
      var nc = sheet.getLastColumn() + 1;
      sheet.getRange(1, nc).setValue("sched_type");
      colIdx["sched_type"] = nc - 1;
    }
    var foundRow = -1;
    for (var i = 1; i < rows.length; i++) {
      var rDay = Number(
        rows[i][colIdx["day_index"] !== undefined ? colIdx["day_index"] : 1],
      );
      var rType = String(
        rows[i][
          colIdx["sched_type"] !== undefined ? colIdx["sched_type"] : 2
        ] || "",
      )
        .toUpperCase()
        .trim();
      if (rDay === dayIndex && rType === schedType) {
        foundRow = i + 1;
        break;
      }
    }
    if (foundRow > 0) {
      sheet
        .getRange(foundRow, colIdx["sched_time_in"] + 1)
        .setValue(schedTimeIn);
      sheet
        .getRange(foundRow, colIdx["sched_time_out"] + 1)
        .setValue(schedTimeOut);
    } else {
      var newRow = headers.map(function () {
        return "";
      });
      if (colIdx["id"] !== undefined)
        newRow[colIdx["id"]] = generateId("ASCHT");
      if (colIdx["day_index"] !== undefined)
        newRow[colIdx["day_index"]] = dayIndex;
      if (colIdx["sched_type"] !== undefined)
        newRow[colIdx["sched_type"]] = schedType;
      if (colIdx["sched_time_in"] !== undefined)
        newRow[colIdx["sched_time_in"]] = schedTimeIn;
      if (colIdx["sched_time_out"] !== undefined)
        newRow[colIdx["sched_time_out"]] = schedTimeOut;
      sheet.appendRow(newRow);
    }
    var applyResult = null;
    if (applyDate && /^\d{4}-\d{2}-\d{2}$/.test(applyDate)) {
      applyResult = _applyTemplateToDateInternal(
        schedTimeIn,
        schedTimeOut,
        applyDate,
      );
    }
    var msg =
      foundRow > 0
        ? "Jam jadwal berhasil diperbarui."
        : "Jam jadwal berhasil disimpan.";
    if (applyResult)
      msg +=
        " Diterapkan ke " +
        applyResult.applied +
        " guru" +
        (applyResult.skipped > 0
          ? ", " + applyResult.skipped + " dilewati."
          : ".");
    try { _invalidateDataSnapshot(); } catch(_) {}
    return { status: "success", message: msg };
  } catch (e) {
    console.error("saveAttendanceSchedTemplate error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function deleteAttendanceSchedTemplate(token, id) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    if (String(user.role).toLowerCase() !== "admin")
      return {
        status: "error",
        message: "Hanya Admin yang dapat menghapus jam jadwal.",
      };
    var sheet = getSheet(SHEET_NAME.ATTENDANCE_SCHED_TEMPLATES);
    var rows = sheet.getDataRange().getValues();
    var colId = rows[0]
      .map(function (h) {
        return String(h).toLowerCase().trim();
      })
      .indexOf("id");
    if (colId < 0) colId = 0;
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][colId]) === String(id)) {
        sheet.deleteRow(i + 1);
        try { _invalidateDataSnapshot(); } catch(_) {}
        return { status: "success", message: "Jam jadwal berhasil dihapus." };
      }
    }
    return { status: "error", message: "Data tidak ditemukan." };
  } catch (e) {
    console.error("deleteAttendanceSchedTemplate error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function _applyTemplateToDateInternal(schedTimeIn, schedTimeOut, dateStr) {
  var tz = Session.getScriptTimeZone();
  var allUsers = getData(SHEET_NAME.USERS);
  var pickets = getData(SHEET_NAME.PICKET_SCHEDULES);
  var subs = getData(SHEET_NAME.SUBSTITUTES);
  var schedules = getData(SHEET_NAME.SCHEDULES);
  var cfg = _getConfigMap();
  var activeTP = cfg["tahun_pelajaran"] || "",
    activeSem = cfg["semester"] || "";
  var dateObj = new Date(dateStr + "T00:00:00");
  var utcMs = dateObj.getTime() + dateObj.getTimezoneOffset() * 60000;
  var dayIndex = new Date(utcMs + 7 * 3600000).getDay();
  var userIds = {};
  
  var isExamPeriod = false;
  var activePeriodName = "";
  var examPeriods = getData(EXAM_SHEET.PERIODS);
  for (var i=0; i<examPeriods.length; i++) {
    var p = examPeriods[i];
    var ps = String(p.date_start||"");
    var pe = String(p.date_end||"");
    if (ps && pe && dateStr >= ps && dateStr <= pe) {
      isExamPeriod = true;
      activePeriodName = String(p.name||"").trim();
      break;
    }
  }

  if (isExamPeriod) {
    var proctors = getData(EXAM_SHEET.PROCTORS);
    proctors.forEach(function(p) {
      var pp = String(p.periode_ujian||"").trim();
      var pDate = p.date instanceof Date ? Utilities.formatDate(p.date, tz, "yyyy-MM-dd") : String(p.date||"");
      if ((!pp || pp === activePeriodName) && pDate === dateStr) {
        userIds[String(p.user_id)] = true;
      }
    });
    var committees = getData(EXAM_SHEET.COMMITTEES);
    committees.forEach(function(c) {
      var cp = String(c.periode_ujian||"").trim();
      if (!cp || cp === activePeriodName) {
        userIds[String(c.user_id)] = true;
      }
    });
  }
  pickets.forEach(function (p) {
    if (String(p.day_index).trim() === String(dayIndex))
      userIds[String(p.user_id)] = true;
  });
  subs.forEach(function (s) {
    if (
      String(s.schedule_id) === "PICKET-DUTY" &&
      String(s.date) === dateStr &&
      s.substitute_user_id
    )
      userIds[String(s.substitute_user_id)] = true;
  });
  schedules.forEach(function (s) {
    var sTP = s.tahun_pelajaran || activeTP,
      sSem = s.semester || activeSem;
    if (
      String(s.day_index).trim() === String(dayIndex) &&
      sTP === activeTP &&
      sSem === activeSem
    )
      userIds[String(s.user_id)] = true;
  });
  subs.forEach(function (s) {
    if (
      String(s.schedule_id) !== "PICKET-DUTY" &&
      String(s.date) === dateStr &&
      s.substitute_user_id
    )
      userIds[String(s.substitute_user_id)] = true;
  });
  allUsers.forEach(function (u) {
    if (
      String(u.additional_role || "")
        .trim()
        .toUpperCase() === "KEPALA_SEKOLAH"
    )
      userIds[String(u.id)] = true;
  });
  var attSheet = getSheet(SHEET_NAME.ATTENDANCE);
  var attRows = attSheet.getDataRange().getValues();
  var attHdrs = attRows[0].map(function (h) {
    return String(h).toLowerCase().trim().replace(/\s+/g, "_");
  });
  var attColIdx = {};
  attHdrs.forEach(function (h, i) {
    attColIdx[h] = i;
  });
  var ensureCol = function (c) {
    if (attColIdx[c] !== undefined) return;
    var nc = attSheet.getLastColumn() + 1;
    attSheet.getRange(1, nc).setValue(c);
    attColIdx[c] = nc - 1;
  };
  ensureCol("sched_time_in");
  ensureCol("sched_time_out");
  var applied = 0,
    skipped = 0;
  Object.keys(userIds).forEach(function (uid) {
    var foundRow = -1,
      hasSched = false;
    for (var i = 1; i < attRows.length; i++) {
      var rowUid = String(
        attRows[i][
          attColIdx["user_id"] !== undefined ? attColIdx["user_id"] : 2
        ] || "",
      );
      var rawDate =
        attRows[i][attColIdx["date"] !== undefined ? attColIdx["date"] : 1];
      var rowDate =
        rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd")
          : String(rawDate || "");
      if (rowUid === uid && rowDate === dateStr) {
        foundRow = i + 1;
        hasSched =
          String(attRows[i][attColIdx["sched_time_in"]] || "").trim() !== "";
        break;
      }
    }
    if (hasSched) {
      skipped++;
      return;
    }
    if (foundRow > 0) {
      attSheet
        .getRange(foundRow, attColIdx["sched_time_in"] + 1)
        .setValue(schedTimeIn);
      attSheet
        .getRange(foundRow, attColIdx["sched_time_out"] + 1)
        .setValue(schedTimeOut);
    } else {
      var nr = attHdrs.map(function () {
        return "";
      });
      if (attColIdx["id"] !== undefined)
        nr[attColIdx["id"]] = generateId("ATT");
      if (attColIdx["date"] !== undefined) nr[attColIdx["date"]] = dateStr;
      if (attColIdx["user_id"] !== undefined) nr[attColIdx["user_id"]] = uid;
      if (attColIdx["sched_time_in"] !== undefined)
        nr[attColIdx["sched_time_in"]] = schedTimeIn;
      if (attColIdx["sched_time_out"] !== undefined)
        nr[attColIdx["sched_time_out"]] = schedTimeOut;
      attSheet.appendRow(nr);
    }
    applied++;
  });
  return { applied: applied, skipped: skipped };
}
function applySchedTemplatesToDate(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    if (String(user.role).toLowerCase() !== "admin")
      return {
        status: "error",
        message: "Hanya Admin yang dapat menerapkan jam jadwal.",
      };
    var dateStr = String((payload && payload.date) || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr))
      return { status: "error", message: "Format tanggal tidak valid." };
    var dateObj = new Date(dateStr + "T00:00:00");
    var utcMs = dateObj.getTime() + dateObj.getTimezoneOffset() * 60000;
    var dayIndex = new Date(utcMs + 7 * 3600000).getDay();
    var isExamPeriod = getData(EXAM_SHEET.PERIODS).some(function (p) {
      var ps = String(p.date_start || ""),
        pe = String(p.date_end || "");
      return ps && pe && dateStr >= ps && dateStr <= pe;
    });
    var activeType = isExamPeriod ? "UJIAN" : "KBM";
    var templateId = String((payload && payload.template_id) || "").trim();
    var allTemplates = getData(SHEET_NAME.ATTENDANCE_SCHED_TEMPLATES);
    var tpl = templateId
      ? allTemplates.find(function (t) {
          return String(t.id) === templateId;
        })
      : allTemplates.find(function (t) {
          return (
            Number(t.day_index) === dayIndex &&
            String(t.sched_type || "KBM")
              .toUpperCase()
              .trim() === activeType
          );
        });
    if (!tpl) {
      var dayNames = [
        "Minggu",
        "Senin",
        "Selasa",
        "Rabu",
        "Kamis",
        "Jumat",
        "Sabtu",
      ];
      return {
        status: "error",
        message:
          "Belum ada jam jadwal " +
          activeType +
          " untuk hari " +
          (dayNames[dayIndex] || dayIndex) +
          ". Atur dulu di tab Atur Jam Jadwal.",
      };
    }
    var rTime = /^([01]\d|2[0-3]):([0-5]\d)$/;
    if (
      !rTime.test(String(tpl.sched_time_in || "")) ||
      !rTime.test(String(tpl.sched_time_out || ""))
    )
      return { status: "error", message: "Format jam di data tidak valid." };
    var res = _applyTemplateToDateInternal(
      tpl.sched_time_in,
      tpl.sched_time_out,
      dateStr,
    );
    try { _invalidateDataSnapshot(); } catch(_) {}
    return {
      status: "success",
      message:
        "Jam jadwal " +
        activeType +
        " diterapkan: " +
        res.applied +
        " guru diperbarui" +
        (res.skipped > 0
          ? ", " + res.skipped + " dilewati (sudah ada jadwal)."
          : "."),
      applied: res.applied,
      skipped: res.skipped,
    };
  } catch (e) {
    console.error("applySchedTemplatesToDate error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function saveAttendanceSched(token, payload) {
  try {
    var caller = verifySession(token);
    if (!caller) {
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    }
    if (String(caller.role || "").toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Hanya Admin yang dapat mengatur jam jadwal.",
      };
    }
    var userId = String(payload.user_id || "").trim();
    var dateStr = String(payload.date || "").trim();
    var schedTimeIn = String(payload.sched_time_in || "").trim();
    var schedTimeOut = String(payload.sched_time_out || "").trim();
    if (!userId || !dateStr) {
      return {
        status: "error",
        message: "Data tidak lengkap (user_id dan date wajib diisi).",
      };
    }
    var TIME_RE = /^([01]\d|2[0-3]):([0-5]\d)$/;
    if (!TIME_RE.test(schedTimeIn) || !TIME_RE.test(schedTimeOut)) {
      return {
        status: "error",
        message:
          "Format jam jadwal tidak valid. Gunakan format HH:mm (contoh: 07:20).",
      };
    }
    function toMins(t) {
      var m = t.match(/^([01]\d|2[0-3]):([0-5]\d)$/);
      return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : -1;
    }
    var inMins = toMins(schedTimeIn);
    var outMins = toMins(schedTimeOut);
    if (outMins <= inMins) {
      return {
        status: "error",
        message: "Jam pulang jadwal harus lebih akhir dari jam masuk jadwal.",
      };
    }
    var sheet = getSheet(SHEET_NAME.ATTENDANCE);
    var rawData = sheet.getDataRange().getValues();
    var headers = rawData[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var colIdx = {};
    headers.forEach(function (h, i) {
      colIdx[h] = i;
    });
    var requiredCols = ["user_id", "date", "sched_time_in", "sched_time_out"];
    var missingCol = requiredCols.find(function (c) {
      return colIdx[c] === undefined;
    });
    if (missingCol) {
      return {
        status: "error",
        message:
          'Struktur sheet Daily_Attendance tidak sesuai (kolom "' +
          missingCol +
          '" tidak ditemukan).',
      };
    }
    var targetRowIndex = -1;
    for (var r = 1; r < rawData.length; r++) {
      var rowUserId = String(rawData[r][colIdx["user_id"]] || "").trim();
      var rowDate = rawData[r][colIdx["date"]];
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(
          rowDate,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd",
        );
      } else {
        rowDate = String(rowDate || "").trim();
      }
      if (rowUserId === userId && rowDate === dateStr) {
        targetRowIndex = r;
        break;
      }
    }
    if (targetRowIndex >= 0) {
      var sheetRow = targetRowIndex + 1; 
      sheet
        .getRange(sheetRow, colIdx["sched_time_in"] + 1)
        .setValue(schedTimeIn);
      sheet
        .getRange(sheetRow, colIdx["sched_time_out"] + 1)
        .setValue(schedTimeOut);
    } else {
      var newRow = headers.map(function () {
        return "";
      });
      if (colIdx["id"] !== undefined) {
        newRow[colIdx["id"]] = generateId("ATT");
      }
      newRow[colIdx["user_id"]] = userId;
      newRow[colIdx["date"]] = dateStr;
      newRow[colIdx["sched_time_in"]] = schedTimeIn;
      newRow[colIdx["sched_time_out"]] = schedTimeOut;
      sheet.appendRow(newRow);
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {
    }
    return { status: "success", message: "Jam jadwal berhasil disimpan." };
  } catch (e) {
    console.error("saveAttendanceSched error: " + e.toString());
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function saveAttendanceLeave(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var targetUserId = String((payload && payload.user_id) || "").trim();
    var dateStr = String((payload && payload.date) || "").trim();
    var leaveTime = String((payload && payload.leave_time) || "")
      .trim()
      .substring(0, 5);
    var returnTime = String((payload && payload.return_time) || "")
      .trim()
      .substring(0, 5);
    var reason = String((payload && payload.reason) || "")
      .trim()
      .substring(0, 200);
    if (!targetUserId || !dateStr) {
      return { status: "error", message: "user_id dan date wajib diisi." };
    }
    if (!reason) {
      return { status: "error", message: "Keterangan wajib diisi." };
    }
    var rTime = /^([01]\d|2[0-3]):([0-5]\d)$/;
    if (!rTime.test(leaveTime)) {
      return {
        status: "error",
        message: "Format jam keluar tidak valid. Gunakan HH:mm.",
      };
    }
    if (!rTime.test(returnTime)) {
      return {
        status: "error",
        message: "Format jam kembali tidak valid. Gunakan HH:mm.",
      };
    }
    var leaveMins =
      parseInt(leaveTime.split(":")[0], 10) * 60 +
      parseInt(leaveTime.split(":")[1], 10);
    var returnMins =
      parseInt(returnTime.split(":")[0], 10) * 60 +
      parseInt(returnTime.split(":")[1], 10);
    if (returnMins <= leaveMins) {
      return {
        status: "error",
        message: "Jam kembali harus lebih besar dari jam keluar.",
      };
    }
    if (!isAdmin) {
      var targetDayIndex = (function () {
        try {
          var d = new Date(dateStr + "T00:00:00");
          var utc = d.getTime() + d.getTimezoneOffset() * 60000;
          return new Date(utc + 7 * 3600000).getDay();
        } catch (_) {
          return -1;
        }
      })();
      var picketSchedules = getData(SHEET_NAME.PICKET_SCHEDULES);
      var allSubstitutes = getData(SHEET_NAME.SUBSTITUTES);
      var callerIsPicket =
        picketSchedules.some(function (p) {
          return (
            String(p.user_id) === String(user.id) &&
            String(p.day_index).trim() === String(targetDayIndex)
          );
        }) ||
        allSubstitutes.some(function (s) {
          return (
            String(s.substitute_user_id) === String(user.id) &&
            String(s.schedule_id) === "PICKET-DUTY" &&
            String(s.date) === dateStr
          );
        });
      if (callerIsPicket) {
        var callerRec = getData(SHEET_NAME.ATTENDANCE).find(function (a) {
          return (
            String(a.user_id) === String(user.id) && String(a.date) === dateStr
          );
        });
        if (!callerRec || !String(callerRec.time_in || "").trim()) {
          return {
            status: "error",
            message: "Kehadiran piket Anda belum dikonfirmasi admin.",
          };
        }
        var targetIsPicket =
          picketSchedules.some(function (p) {
            return (
              String(p.user_id) === String(targetUserId) &&
              String(p.day_index).trim() === String(targetDayIndex)
            );
          }) ||
          allSubstitutes.some(function (s) {
            return (
              String(s.substitute_user_id) === String(targetUserId) &&
              String(s.schedule_id) === "PICKET-DUTY" &&
              String(s.date) === dateStr
            );
          });
        if (targetIsPicket) {
          return {
            status: "error",
            message:
              "Pencatatan izin keluar guru piket hanya dapat dilakukan oleh admin.",
          };
        }
      } else {
        if (String(targetUserId) !== String(user.id)) {
          return {
            status: "error",
            message:
              "Anda tidak memiliki wewenang untuk mencatat kehadiran guru lain.",
          };
        }
      }
    }
    var attRec = getData(SHEET_NAME.ATTENDANCE).find(function (a) {
      return (
        String(a.user_id) === String(targetUserId) && String(a.date) === dateStr
      );
    });
    if (!attRec || !String(attRec.time_in || "").trim()) {
      return {
        status: "error",
        message:
          "Jam datang guru belum dicatat. Catat jam datang terlebih dahulu.",
      };
    }
    var existingLeaves = getData(SHEET_NAME.ATTENDANCE_LEAVES).filter(
      function (lv) {
        return (
          String(lv.user_id) === String(targetUserId) &&
          String(lv.date) === dateStr
        );
      },
    );
    for (var i = 0; i < existingLeaves.length; i++) {
      var exLv = existingLeaves[i];
      var exLeave = String(exLv.leave_time || "")
        .trim()
        .substring(0, 5);
      var exReturn = String(exLv.return_time || "")
        .trim()
        .substring(0, 5);
      if (!rTime.test(exLeave) || !rTime.test(exReturn)) continue;
      var exLeaveMins =
        parseInt(exLeave.split(":")[0], 10) * 60 +
        parseInt(exLeave.split(":")[1], 10);
      var exReturnMins =
        parseInt(exReturn.split(":")[0], 10) * 60 +
        parseInt(exReturn.split(":")[1], 10);
      if (leaveMins < exReturnMins && returnMins > exLeaveMins) {
        return {
          status: "error",
          message:
            "Rentang waktu tumpang tindih dengan izin keluar yang sudah ada (" +
            exLeave +
            "–" +
            exReturn +
            ").",
        };
      }
    }
    var newId = generateId("LV");
    getSheet(SHEET_NAME.ATTENDANCE_LEAVES).appendRow([
      newId,
      dateStr,
      targetUserId,
      leaveTime,
      returnTime,
      reason,
    ]);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Izin keluar sementara berhasil dicatat.",
      id: newId,
    };
  } catch (e) {
    console.error("saveAttendanceLeave error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function deleteAttendanceLeave(token, leaveId) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var sheet = getSheet(SHEET_NAME.ATTENDANCE_LEAVES);
    var rows = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var headers = rows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var colIdx = {};
    headers.forEach(function (h, i) {
      colIdx[h] = i;
    });
    var foundRow = -1;
    var entryUserId = "";
    var entryDate = "";
    for (var i = 1; i < rows.length; i++) {
      var rowId = String(
        rows[i][colIdx["id"] !== undefined ? colIdx["id"] : 0] || "",
      );
      if (rowId === String(leaveId)) {
        foundRow = i + 1; 
        entryUserId = String(
          rows[i][colIdx["user_id"] !== undefined ? colIdx["user_id"] : 2] ||
            "",
        );
        var rawDate =
          rows[i][colIdx["date"] !== undefined ? colIdx["date"] : 1];
        entryDate =
          rawDate instanceof Date
            ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd")
            : String(rawDate || "");
        break;
      }
    }
    if (foundRow < 0) {
      return { status: "error", message: "Entri izin keluar tidak ditemukan." };
    }
    if (!isAdmin) {
      var targetDayIndex = (function () {
        try {
          var d = new Date(entryDate + "T00:00:00");
          var utc = d.getTime() + d.getTimezoneOffset() * 60000;
          return new Date(utc + 7 * 3600000).getDay();
        } catch (_) {
          return -1;
        }
      })();
      var picketSchedules = getData(SHEET_NAME.PICKET_SCHEDULES);
      var allSubstitutes = getData(SHEET_NAME.SUBSTITUTES);
      var callerIsPicket =
        picketSchedules.some(function (p) {
          return (
            String(p.user_id) === String(user.id) &&
            String(p.day_index).trim() === String(targetDayIndex)
          );
        }) ||
        allSubstitutes.some(function (s) {
          return (
            String(s.substitute_user_id) === String(user.id) &&
            String(s.schedule_id) === "PICKET-DUTY" &&
            String(s.date) === entryDate
          );
        });
      if (!callerIsPicket && String(entryUserId) !== String(user.id)) {
        return {
          status: "error",
          message: "Anda tidak berwenang menghapus entri ini.",
        };
      }
    }
    sheet.deleteRow(foundRow);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Entri izin keluar berhasil dihapus.",
    };
  } catch (e) {
    console.error("deleteAttendanceLeave error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}
function getEventDefinitions(token) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    if (!_authorizeEventWriter(user)) {
      return {
        status: "error",
        message: "Akses Ditolak: Halaman ini hanya untuk Admin.",
      };
    }
    var rows = getData(EVENT_SHEET.DEFINITIONS);
    return { status: "success", events: rows };
  } catch (e) {
    console.error("getEventDefinitions error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function saveEventDefinition(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    if (!_authorizeEventWriter(user)) {
      return {
        status: "error",
        message: "Akses Ditolak: Halaman ini hanya untuk Admin.",
      };
    }
    var validation = _validateEventDefinition(payload);
    if (!validation.ok) {
      return { status: "error", message: validation.message };
    }
    var payloadId = payload && payload.id ? String(payload.id).trim() : "";
    var allDefs = getData(EVENT_SHEET.DEFINITIONS);
    var nameNorm = String(payload.name).trim().toLowerCase();
    var duplicate = allDefs.find(function (r) {
      return (
        String(r.name || "")
          .trim()
          .toLowerCase() === nameNorm && String(r.id || "") !== payloadId
      );
    });
    if (duplicate) {
      return {
        status: "error",
        message: "Nama acara sudah digunakan oleh acara lain.",
      };
    }
    var sheet = getSheet(EVENT_SHEET.DEFINITIONS);
    var savedId;
    if (payloadId) {
      var rows = sheet.getDataRange().getValues();
      var headers = rows[0].map(function (h) {
        return String(h).toLowerCase().trim().replace(/\s+/g, "_");
      });
      var colIdx = {};
      headers.forEach(function (h, i) {
        colIdx[h] = i;
      });
      var foundRowNum = -1;
      for (var i = 1; i < rows.length; i++) {
        if (
          String(rows[i][colIdx["id"] !== undefined ? colIdx["id"] : 0]) ===
          payloadId
        ) {
          foundRowNum = i + 1; 
          break;
        }
      }
      if (foundRowNum === -1) {
        return { status: "error", message: "Definisi acara tidak ditemukan." };
      }
      sheet
        .getRange(
          foundRowNum,
          (colIdx["name"] !== undefined ? colIdx["name"] : 1) + 1,
        )
        .setValue(String(payload.name || "").trim());
      sheet
        .getRange(
          foundRowNum,
          (colIdx["type"] !== undefined ? colIdx["type"] : 2) + 1,
        )
        .setValue(String(payload.type || "").trim());
      sheet
        .getRange(
          foundRowNum,
          (colIdx["recurrence_day_index"] !== undefined
            ? colIdx["recurrence_day_index"]
            : 3) + 1,
        )
        .setValue(
          payload.recurrence_day_index !== null &&
            payload.recurrence_day_index !== undefined
            ? payload.recurrence_day_index
            : "",
        );
      sheet
        .getRange(
          foundRowNum,
          (colIdx["dates_json"] !== undefined ? colIdx["dates_json"] : 4) + 1,
        )
        .setValue(
          payload.dates_json !== null && payload.dates_json !== undefined
            ? payload.dates_json
            : "",
        );
      sheet
        .getRange(
          foundRowNum,
          (colIdx["time_start"] !== undefined ? colIdx["time_start"] : 5) + 1,
        )
        .setValue(
          String(payload.time_start || "")
            .trim()
            .slice(0, 5),
        );
      sheet
        .getRange(
          foundRowNum,
          (colIdx["time_end"] !== undefined ? colIdx["time_end"] : 6) + 1,
        )
        .setValue(
          String(payload.time_end || "")
            .trim()
            .slice(0, 5),
        );
      sheet
        .getRange(
          foundRowNum,
          (colIdx["jtm_val"] !== undefined ? colIdx["jtm_val"] : 7) + 1,
        )
        .setValue(Number(payload.jtm_val));
      sheet
        .getRange(
          foundRowNum,
          (colIdx["description"] !== undefined ? colIdx["description"] : 8) + 1,
        )
        .setValue(String(payload.description || ""));
      savedId = payloadId;
    } else {
      savedId = generateId("EVD");
      var now = new Date().toISOString();
      sheet.appendRow([
        savedId,
        String(payload.name || "").trim(),
        String(payload.type || "").trim(),
        payload.recurrence_day_index !== null &&
        payload.recurrence_day_index !== undefined
          ? payload.recurrence_day_index
          : "",
        payload.dates_json !== null && payload.dates_json !== undefined
          ? payload.dates_json
          : "",
        String(payload.time_start || "").trim(),
        String(payload.time_end || "").trim(),
        Number(payload.jtm_val),
        String(payload.description || ""),
        String(user.id),
        now,
      ]);
    }
    return { status: "success", id: savedId };
  } catch (e) {
    console.error("saveEventDefinition error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function deleteEventDefinition(token, eventId, force) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    if (!_authorizeEventWriter(user)) {
      return {
        status: "error",
        message: "Akses Ditolak: Halaman ini hanya untuk Admin.",
      };
    }
    var eId = String(eventId || "").trim();
    if (!eId) return { status: "error", message: "ID acara tidak valid." };
    var allAttendance = getData(EVENT_SHEET.ATTENDANCE);
    var allJournals = getData(EVENT_SHEET.JOURNALS);
    var relatedAttendance = allAttendance.filter(function (r) {
      return String(r.event_id || "") === eId;
    });
    var relatedJournals = allJournals.filter(function (r) {
      return String(r.event_id || "") === eId;
    });
    var relatedCount = {
      attendance: relatedAttendance.length,
      journals: relatedJournals.length,
    };
    if (force !== true) {
      return { status: "success", deleted: false, relatedCount: relatedCount };
    }
    var journalSheet = getSheet(EVENT_SHEET.JOURNALS);
    var journalRows = journalSheet.getDataRange().getValues();
    var journalHeaders = journalRows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var jEventIdCol = journalHeaders.indexOf("event_id");
    if (jEventIdCol === -1) jEventIdCol = 1; 
    for (var j = journalRows.length - 1; j >= 1; j--) {
      if (String(journalRows[j][jEventIdCol]) === eId) {
        journalSheet.deleteRow(j + 1);
      }
    }
    var attendanceSheet = getSheet(EVENT_SHEET.ATTENDANCE);
    var attRows = attendanceSheet.getDataRange().getValues();
    var attHeaders = attRows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var aEventIdCol = attHeaders.indexOf("event_id");
    if (aEventIdCol === -1) aEventIdCol = 1;
    for (var a = attRows.length - 1; a >= 1; a--) {
      if (String(attRows[a][aEventIdCol]) === eId) {
        attendanceSheet.deleteRow(a + 1);
      }
    }
    var defSheet = getSheet(EVENT_SHEET.DEFINITIONS);
    var defRows = defSheet.getDataRange().getValues();
    var defHeaders = defRows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var dIdCol = defHeaders.indexOf("id");
    if (dIdCol === -1) dIdCol = 0;
    for (var d = defRows.length - 1; d >= 1; d--) {
      if (String(defRows[d][dIdCol]) === eId) {
        defSheet.deleteRow(d + 1);
        break; 
      }
    }
    return { status: "success", deleted: true, relatedCount: relatedCount };
  } catch (e) {
    console.error("deleteEventDefinition error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function getTodayEvents(token) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var isPicket = _isPicketOfficer(user);
    if (!isAdmin && !isPicket) {
      return {
        status: "error",
        message:
          "Akses Ditolak: Halaman ini hanya untuk Admin atau Guru Piket.",
      };
    }
    var isPicketConfirmed = isAdmin
      ? true
      : isPicket && _hasTimeInToday(user.id);
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    var dayIndexStr = Utilities.formatDate(now, tz, "u");
    var dayOfWeek = Number(dayIndexStr) % 7;
    var allDefs = getData(EVENT_SHEET.DEFINITIONS);
    var filteredEvents = allDefs.filter(function (ev) {
      try {
        var type = String(ev.type || "")
          .trim()
          .toLowerCase();
        if (type === "rutin") {
          return Number(ev.recurrence_day_index) === dayOfWeek;
        } else if (type === "insidental") {
          if (!ev.dates_json) return false;
          var dates;
          try {
            dates = JSON.parse(String(ev.dates_json));
          } catch (_) {
            return false;
          }
          if (!Array.isArray(dates)) return false;
          return dates.indexOf(todayStr) !== -1;
        }
        return false;
      } catch (_) {
        return false;
      }
    });
    return {
      status: "success",
      events: filteredEvents,
      today: todayStr,
      isPicketConfirmed: isPicketConfirmed,
    };
  } catch (e) {
    console.error("getTodayEvents error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function getEventAttendance(token, eventId, date) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var isPicketConfirmed = _isPicketOfficer(user);
    if (!isAdmin && !isPicketConfirmed) {
      return {
        status: "error",
        message: "Akses Ditolak: Konfirmasi piket Anda belum disetujui Admin.",
      };
    }
    var eId = String(eventId || "").trim();
    var dateStr = String(date || "").trim();
    var allUsers = getData(SHEET_NAME.USERS).filter(function (u) {
      return (
        String(u.role || "").toLowerCase() === "guru" &&
        String(u.username || "") !== "PenjagaMadrasah"
      );
    });
    var allAttendance = getData(EVENT_SHEET.ATTENDANCE).filter(function (r) {
      return (
        String(r.event_id || "") === eId && String(r.date || "") === dateStr
      );
    });
    var rows = allUsers.map(function (u) {
      var userId = String(u.id || "");
      var name = String(u.full_name || u.name || "");
      var record = allAttendance.find(function (r) {
        return String(r.user_id || "") === userId;
      });
      if (record) {
        return {
          user_id: userId,
          full_name: name,
          time_in: String(record.time_in || ""),
          journal_submitted:
            record.journal_submitted === true ||
            String(record.journal_submitted).toLowerCase() === "true",
          jtm_val:
            record.jtm_val !== undefined && record.jtm_val !== ""
              ? record.jtm_val
              : "",
          attendance_id: String(record.id || ""),
        };
      } else {
        return {
          user_id: userId,
          full_name: name,
          time_in: "",
          journal_submitted: false,
          jtm_val: "",
          attendance_id: "",
        };
      }
    });
    return { status: "success", rows: rows };
  } catch (e) {
    console.error("getEventAttendance error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function _upsertAttendanceRecord(
  existing,
  newTimeIn,
  confirmedBy,
  confirmedAt,
) {
  if (existing) {
    var updated = {};
    var keys = Object.keys(existing);
    for (var k = 0; k < keys.length; k++) {
      updated[keys[k]] = existing[keys[k]];
    }
    updated.time_in = newTimeIn;
    updated.confirmed_by = confirmedBy;
    updated.confirmed_at = confirmedAt;
    return updated;
  }
  return {
    time_in: newTimeIn,
    confirmed_by: confirmedBy,
    confirmed_at: confirmedAt,
    journal_submitted: false,
    jtm_val: "",
  };
}
function saveEventAttendance(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var isAdmin = String(user.role).toLowerCase() === "admin";
    var isPicketConfirmed = isAdmin
      ? true
      : _isPicketOfficer(user) && _hasTimeInToday(user.id);
    if (!_authorizeAttendanceWriter(user, isPicketConfirmed)) {
      return {
        status: "error",
        message:
          "Akses Ditolak: Anda tidak berwenang mencatat kehadiran acara (piket belum dikonfirmasi).",
      };
    }
    var HH_MM = /^([01]\d|2[0-3]):[0-5]\d$/;
    var YYYY_MM_DD = /^\d{4}-\d{2}-\d{2}$/;
    var timeIn =
      payload && typeof payload.time_in === "string"
        ? payload.time_in.trim()
        : "";
    var dateStr =
      payload && typeof payload.date === "string" ? payload.date.trim() : "";
    if (timeIn !== "" && !HH_MM.test(timeIn)) {
      return {
        status: "error",
        message: "Format jam masuk tidak valid (gunakan HH:mm).",
      };
    }
    if (!YYYY_MM_DD.test(dateStr)) {
      return {
        status: "error",
        message: "Format tanggal tidak valid (gunakan YYYY-MM-DD).",
      };
    }
    var eventId =
      payload && payload.event_id ? String(payload.event_id).trim() : "";
    var userId =
      payload && payload.user_id ? String(payload.user_id).trim() : "";
    if (!eventId)
      return { status: "error", message: "event_id tidak boleh kosong." };
    if (!userId)
      return { status: "error", message: "user_id tidak boleh kosong." };
    var eventDef = findData(EVENT_SHEET.DEFINITIONS, "id", eventId);
    if (!eventDef) {
      return {
        status: "error",
        message: "Acara dengan event_id tersebut tidak ditemukan.",
      };
    }
    // Validasi: jam masuk tidak boleh kurang dari jam mulai acara
    var eventTimeStart = String(eventDef.time_start || "").trim();
    if (timeIn !== "" && eventTimeStart && HH_MM.test(eventTimeStart) && timeIn < eventTimeStart) {
      return {
        status: "error",
        message: "Jam masuk (" + timeIn + ") tidak boleh kurang dari jam mulai acara (" + eventTimeStart + ").",
      };
    }
    // Validasi: jam masuk tidak boleh melebihi jam selesai acara
    var eventTimeEnd = String(eventDef.time_end || "").trim();
    if (timeIn !== "" && eventTimeEnd && HH_MM.test(eventTimeEnd) && timeIn >= eventTimeEnd) {
      return {
        status: "error",
        message: "Jam masuk (" + timeIn + ") tidak boleh melebihi jam selesai acara (" + eventTimeEnd + ").",
      };
    }
    var targetUser = findData(SHEET_NAME.USERS, "id", userId);
    if (!targetUser || String(targetUser.role || "").toLowerCase() !== "guru") {
      return {
        status: "error",
        message: "Pengguna dengan user_id tersebut bukan guru yang terdaftar.",
      };
    }
    var confirmedAt = new Date().toISOString();
    var confirmedBy = String(user.id);
    var allAttendance = getData(EVENT_SHEET.ATTENDANCE);
    var existing = allAttendance.find(function (r) {
      return (
        String(r.event_id || "") === eventId &&
        String(r.user_id || "") === userId &&
        String(r.date || "") === dateStr
      );
    });
    if (existing) {
      var attSheet = getSheet(EVENT_SHEET.ATTENDANCE);
      var attRows = attSheet.getDataRange().getValues();
      var attHeaders = attRows[0].map(function (h) {
        return String(h).toLowerCase().trim().replace(/\s+/g, "_");
      });
      var colOf = {};
      attHeaders.forEach(function (h, i) {
        colOf[h] = i;
      });
      var existingId = String(existing.id || "");
      var idCol = colOf["id"] !== undefined ? colOf["id"] : 0;
      for (var i = 1; i < attRows.length; i++) {
        if (String(attRows[i][idCol]) === existingId) {
          var rowNum = i + 1; 
          if (timeIn === "") {
            var js = existing.journal_submitted;
            if (js === true || String(js).toUpperCase() === "TRUE") {
              return {
                status: "error",
                message:
                  "Tidak dapat membatalkan kehadiran karena jurnal acara sudah diisi.",
              };
            }
            attSheet.deleteRow(rowNum);
          } else {
            attSheet
              .getRange(
                rowNum,
                (colOf["time_in"] !== undefined ? colOf["time_in"] : 4) + 1,
              )
              .setValue(timeIn);
            attSheet
              .getRange(
                rowNum,
                (colOf["confirmed_by"] !== undefined
                  ? colOf["confirmed_by"]
                  : 5) + 1,
              )
              .setValue(confirmedBy);
            attSheet
              .getRange(
                rowNum,
                (colOf["confirmed_at"] !== undefined
                  ? colOf["confirmed_at"]
                  : 6) + 1,
              )
              .setValue(confirmedAt);
          }
          break;
        }
      }
    } else {
      if (timeIn === "") {
        return { status: "success" };
      }
      var newRecord = _upsertAttendanceRecord(
        null,
        timeIn,
        confirmedBy,
        confirmedAt,
      );
      var newId = generateId("EVA");
      getSheet(EVENT_SHEET.ATTENDANCE).appendRow([
        newId,
        eventId,
        dateStr,
        userId,
        newRecord.time_in,
        newRecord.confirmed_by,
        newRecord.confirmed_at,
        newRecord.journal_submitted,
        newRecord.jtm_val,
      ]);
    }
    try {
      var email = _notifGetEmail_(payload.user_id);
      if (email) {
        var isCancelled = timeIn === "";
        var eventName = _escHtml_(eventDef.name || "Acara");
        var teacherName = _escHtml_(
          targetUser.full_name || targetUser.name || "Guru",
        );
        var subject = isCancelled
          ? "Pembatalan Kehadiran Acara"
          : "Kehadiran Acara Tercatat";
        var gradient = isCancelled
          ? "linear-gradient(135deg,#DC2626,#EF4444)"
          : "linear-gradient(135deg,#059669,#10B981)";
        var actionText = isCancelled
          ? "Kehadiran Anda pada acara <strong>" +
            eventName +
            "</strong> telah dibatalkan."
          : "Kehadiran Anda pada acara <strong>" +
            eventName +
            "</strong> telah berhasil dicatat pada jam <strong>" +
            _escHtml_(timeIn) +
            "</strong>.";
        var htmlBody =
          '<div style="font-family:Inter,Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px;color:#0f172a;">' +
          '<div style="background:' +
          gradient +
          ';color:white;padding:20px 24px;border-radius:14px 14px 0 0;">' +
          '<div style="font-size:13px;opacity:0.85;letter-spacing:0.05em;">SiM-GURU</div>' +
          '<div style="font-size:20px;font-weight:800;margin-top:4px;">' +
          subject +
          "</div>" +
          "</div>" +
          '<div style="border:1px solid #e2e8f0;border-top:0;border-radius:0 0 14px 14px;padding:24px;background:#fff;">' +
          '<p style="font-size:14px;color:#334155;margin:0 0 16px;">Halo <strong>' +
          teacherName +
          "</strong>,</p>" +
          '<p style="font-size:14px;color:#334155;margin:0 0 16px;">' +
          actionText +
          "</p>" +
          '<div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:12px;padding:16px;margin:16px 0;">' +
          '<p style="margin:0 0 8px;font-size:13px;"><span style="color:#64748B;display:inline-block;width:100px;">Acara</span>: <strong>' +
          eventName +
          "</strong></p>" +
          '<p style="margin:0 0 8px;font-size:13px;"><span style="color:#64748B;display:inline-block;width:100px;">Tanggal</span>: <strong>' +
          _escHtml_(dateStr) +
          "</strong></p>" +
          (!isCancelled
            ? '<p style="margin:0;font-size:13px;"><span style="color:#64748B;display:inline-block;width:100px;">Waktu Hadir</span>: <strong>' +
              _escHtml_(timeIn) +
              "</strong></p>"
            : "") +
          "</div>" +
          '<p style="font-size:13px;color:#64748b;margin:16px 0 16px;">' +
          (!isCancelled
            ? "Jangan lupa untuk mengisi jurnal acara terkait pada sistem SiM-Guru setelah kegiatan selesai."
            : "Jika Anda merasa ini adalah kesalahan, silakan hubungi admin atau guru piket.") +
          "</p>" +
          '<hr style="border:0;border-top:1px solid #e2e8f0;margin:20px 0;">' +
          '<p style="font-size:11px;color:#94a3b8;margin:0;">Email otomatis dari SiM-Guru. Jangan balas email ini.</p>' +
          "</div>" +
          "</div>";
        var textBody =
          "Halo " +
          (targetUser.full_name || "Guru") +
          ",\n\n" +
          (isCancelled
            ? "Kehadiran Anda pada acara " +
              (eventDef.name || "") +
              " telah dibatalkan."
            : "Kehadiran Anda pada acara " +
              (eventDef.name || "") +
              " telah berhasil dicatat pada jam " +
              timeIn +
              ".") +
          "\n\nTanggal: " +
          dateStr;
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: textBody,
          htmlBody: htmlBody,
          name: "SiM-Guru",
        });
      }
    } catch (emailErr) {
    }
    return { status: "success" };
  } catch (e) {
    console.error("saveEventAttendance error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function getMyPendingJournals(token) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    if (String(user.role || "").toLowerCase() !== "guru") {
      return { status: "error", message: "Akses Ditolak." };
    }
    var todayStr = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    var userId = String(user.id);
    var allAttendance = getData(EVENT_SHEET.ATTENDANCE);
    var pending = allAttendance.filter(function (r) {
      if (String(r.user_id || "") !== userId) return false;
      if (String(r.date || "") !== todayStr) return false;
      var js = r.journal_submitted;
      if (js === true || String(js).toUpperCase() === "TRUE") return false;
      return true;
    });
    var allDefs = getData(EVENT_SHEET.DEFINITIONS);
    var defsMap = {};
    allDefs.forEach(function (d) {
      defsMap[String(d.id || "")] = String(d.name || "");
    });
    var items = pending.map(function (r) {
      return {
        attendance_id: String(r.id || ""),
        event_id: String(r.event_id || ""),
        event_name: defsMap[String(r.event_id || "")] || "",
        date: String(r.date || ""),
        time_in: String(r.time_in || ""),
      };
    });
    return { status: "success", items: items };
  } catch (e) {
    console.error("getMyPendingJournals error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function saveEventJournal(token, payload) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    if (String(user.role || "").toLowerCase() !== "guru") {
      return { status: "error", message: "Akses Ditolak." };
    }
    var attendanceId =
      payload && payload.attendance_id
        ? String(payload.attendance_id).trim()
        : "";
    var attendanceRecord = findData(EVENT_SHEET.ATTENDANCE, "id", attendanceId);
    if (!attendanceRecord) {
      return { status: "error", message: "Catatan kehadiran tidak ditemukan." };
    }
    if (String(attendanceRecord.user_id) !== String(user.id)) {
      return {
        status: "error",
        message: "Akses Ditolak: Jurnal ini bukan milik Anda.",
      };
    }
    var js = attendanceRecord.journal_submitted;
    if (js === true || String(js).toUpperCase() === "TRUE") {
      return {
        status: "error",
        message: "Jurnal untuk acara ini sudah pernah diisi.",
      };
    }
    var todayStrJrn = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    if (String(attendanceRecord.date) !== todayStrJrn) {
      return {
        status: "error",
        message:
          "Jurnal hanya dapat diisi pada hari yang sama dengan kehadiran.",
      };
    }
    var description =
      payload && typeof payload.description === "string"
        ? payload.description.trim()
        : "";
    if (!description) {
      return {
        status: "error",
        message: "Deskripsi jurnal tidak boleh kosong.",
      };
    }
    if (description.length > 1000) {
      return {
        status: "error",
        message: "Deskripsi jurnal tidak boleh melebihi 1000 karakter.",
      };
    }
    var eventDef = findData(
      EVENT_SHEET.DEFINITIONS,
      "id",
      attendanceRecord.event_id,
    );
    var jtmSnapshot = eventDef ? Number(eventDef.jtm_val) : 0;
    try {
      getSheet(EVENT_SHEET.JOURNALS).appendRow([
        generateId("EVJ"),
        attendanceRecord.event_id,
        attendanceRecord.id,
        String(user.id),
        attendanceRecord.date,
        description,
        new Date().toISOString(),
      ]);
    } catch (insertErr) {
      console.error("saveEventJournal INSERT error: " + insertErr);
      return { status: "error", message: "Terjadi kesalahan server." };
    }
    var attSheet = getSheet(EVENT_SHEET.ATTENDANCE);
    var attRows = attSheet.getDataRange().getValues();
    var attHeaders = attRows[0].map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
    var colOf = {};
    attHeaders.forEach(function (h, i) {
      colOf[h] = i;
    });
    var idCol = colOf["id"] !== undefined ? colOf["id"] : 0;
    var jsCol =
      colOf["journal_submitted"] !== undefined ? colOf["journal_submitted"] : 7;
    var jtmCol = colOf["jtm_val"] !== undefined ? colOf["jtm_val"] : 8;
    for (var i = 1; i < attRows.length; i++) {
      if (String(attRows[i][idCol]) === String(attendanceRecord.id)) {
        var rowNum = i + 1; 
        attSheet.getRange(rowNum, jsCol + 1).setValue(true);
        attSheet.getRange(rowNum, jtmCol + 1).setValue(jtmSnapshot);
        break;
      }
    }
    try {
      _invalidateDataSnapshot();
    } catch (e) {}
    return { status: "success" };
  } catch (e) {
    console.error("saveEventJournal error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function getEventAttendanceHistory(token, month, year) {
  try {
    var user = verifySession(token);
    if (!user) {
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    }
    var userId = String(user.id);
    var m = Number(month);
    var y = Number(year);
    function pad(n) {
      return String(n).padStart(2, "0");
    }
    var firstDay = y + "-" + pad(m) + "-01";
    var lastDayNum = new Date(y, m, 0).getDate();
    var lastDay = y + "-" + pad(m) + "-" + pad(lastDayNum);
    var allAttendances = getData(EVENT_SHEET.ATTENDANCE);
    var filtered = allAttendances.filter(function (r) {
      return (
        String(r.user_id) === userId &&
        String(r.date) >= firstDay &&
        String(r.date) <= lastDay
      );
    });
    var allDefinitions = getData(EVENT_SHEET.DEFINITIONS);
    var eventMap = {};
    allDefinitions.forEach(function (ev) {
      eventMap[String(ev.id)] = String(ev.name || "");
    });
    var history = filtered.map(function (r) {
      var dateParts = String(r.date).split("-");
      var dateDisplay =
        dateParts.length === 3
          ? dateParts[2] + "/" + dateParts[1] + "/" + dateParts[0]
          : String(r.date);
      var submitted = r.journal_submitted;
      var isSubmitted =
        submitted === true || String(submitted).toUpperCase() === "TRUE";
      var journalStatus = isSubmitted ? "Sudah Diisi" : "Belum Diisi";
      var jtmVal = isSubmitted ? Number(r.jtm_val) || 0 : 0;
      return {
        attendance_id: String(r.id),
        event_id: String(r.event_id),
        event_name: eventMap[String(r.event_id)] || "",
        date: dateDisplay,
        date_raw: String(r.date),
        time_in: String(r.time_in || ""),
        journal_status: journalStatus,
        jtm_val: jtmVal,
      };
    });
    return { status: "success", history: history };
  } catch (e) {
    console.error("getEventAttendanceHistory error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
function getEventJournalHistory(token, month, year) {
  try {
    var user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali." };
    var role = String(user.role || "").toLowerCase();
    var userId = String(user.id);
    if (role !== "guru" && role !== "admin") {
      return { status: "error", message: "Akses ditolak." };
    }
    var m = parseInt(month, 10);
    var y = parseInt(year, 10);
    if (isNaN(m) || isNaN(y) || m < 1 || m > 12) {
      return { status: "error", message: "Parameter bulan/tahun tidak valid." };
    }
    var pad = function (n) {
      return n < 10 ? "0" + n : String(n);
    };
    var firstDay = y + "-" + pad(m) + "-01";
    var lastDayNum = new Date(y, m, 0).getDate();
    var lastDay = y + "-" + pad(m) + "-" + pad(lastDayNum);
    var allJournals = getData(EVENT_SHEET.JOURNALS);
    var filtered = allJournals.filter(function (r) {
      var dateOk = String(r.date) >= firstDay && String(r.date) <= lastDay;
      var roleOk = role === "admin" ? true : String(r.user_id) === userId;
      return dateOk && roleOk;
    });
    var allDefinitions = getData(EVENT_SHEET.DEFINITIONS);
    var eventMap = {};
    allDefinitions.forEach(function (ev) {
      eventMap[String(ev.id)] = String(ev.name || "");
    });
    var allUsers = getData(SHEET_NAME.USERS);
    var userMap = {};
    allUsers.forEach(function (u) {
      userMap[String(u.id)] = String(u.full_name || "");
    });
    var history = filtered.map(function (r) {
      var dateParts = String(r.date).split("-");
      var dateDisplay =
        dateParts.length === 3
          ? dateParts[2] + "/" + dateParts[1] + "/" + dateParts[0]
          : String(r.date);
      var desc = String(r.description || "");
      return {
        id: String(r.id),
        event_name: eventMap[String(r.event_id)] || "Unknown Event",
        user_name: userMap[String(r.user_id)] || "Unknown User",
        date: dateDisplay,
        date_raw: String(r.date),
        description: desc,
        timestamp:
          r.submitted_at instanceof Date
            ? r.submitted_at.toISOString()
            : String(r.submitted_at),
      };
    });
    history.sort(function (a, b) {
      if (a.date_raw > b.date_raw) return -1;
      if (a.date_raw < b.date_raw) return 1;
      if (a.timestamp > b.timestamp) return -1;
      if (a.timestamp < b.timestamp) return 1;
      return 0;
    });
    return { status: "success", history: history };
  } catch (e) {
    console.error("getEventJournalHistory error: " + e);
    return { status: "error", message: "Terjadi kesalahan server." };
  }
}
