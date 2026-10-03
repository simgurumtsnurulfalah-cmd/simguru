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
