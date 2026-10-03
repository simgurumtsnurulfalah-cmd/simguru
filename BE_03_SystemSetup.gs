// SiM-Guru — BE_03_SystemSetup.gs
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
