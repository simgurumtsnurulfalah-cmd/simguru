// SiM-Guru — BE_24_Events.gs
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
