// SiM-Guru — BE_17_CeremonySubstitutes.gs
function saveCeremonySchedule(token, date, userId) {
  const user = verifySession(token);
  if (!user || user.role !== "admin") return { status: "error" };
  const d = new Date(date + "T00:00:00");
  if (d.getDay() !== 1)
    return {
      status: "error",
      message: "Jadwal upacara hanya boleh pada hari Senin!",
    };
  const existing = getData(SHEET_NAME.CEREMONY_SCHEDULES);
  const duplicate = existing.find((c) => safeDate(c.date) === date);
  if (duplicate)
    return {
      status: "error",
      message: "Tanggal ini sudah memiliki jadwal pembina upacara!",
    };

  const config = {};
  (getData("Config") || []).forEach((c) => { config[c.key] = c.value; });
  const activeTP = String(config.tahun_pelajaran || "").toLowerCase();
  const activeSem = String(config.semester || "");

  const sheet = getSheet("Ceremony_Schedules");
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0] || [];
  if (!headers.some((h) => String(h).toLowerCase().trim() === "tahun_pelajaran")) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("tahun_pelajaran");
  }
  if (!headers.some((h) => String(h).toLowerCase().trim() === "semester")) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("semester");
  }

  const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newId = generateId("CER");
  const newRow = new Array(updatedHeaders.length).fill("");
  updatedHeaders.forEach((h, i) => {
    const key = String(h).toLowerCase().trim();
    if (key === "id") newRow[i] = newId;
    else if (key === "date") newRow[i] = date;
    else if (key === "user_id") newRow[i] = userId;
    else if (key === "tahun_pelajaran") newRow[i] = activeTP;
    else if (key === "semester") newRow[i] = activeSem;
  });
  
  sheet.appendRow(newRow);
  return { status: "success" };
}

function deleteCeremonySchedule(token, id) {
  return deleteRowById(token, "Ceremony_Schedules", id);
}

function confirmCeremonyLeader(token, userId) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const isAdmin = String(user.role).toLowerCase() === "admin";
  const isSelf = String(user.id) === String(userId);
  if (!isAdmin && !isSelf) {
    return {
      status: "error",
      message: "Hanya admin atau pembina upacara yang berwenang.",
    };
  }
  const attRes = confirmTeacherPresence(token, userId, "");
  if (attRes.status === "error" && attRes.message !== "Sudah dikonfirmasi" && attRes.message !== "Anda tidak berwenang mengonfirmasi kehadiran guru lain.") {
    return attRes;
  }
  const sheet = getSheet("Teaching_Logs");
  const newId = generateId("LOG-CER");
  const todayStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const existing = getData("Teaching_Logs").find(
    (l) =>
      l.date === todayStr &&
      String(l.user_id) === String(userId) &&
      l.schedule_id === "CEREMONY-DUTY",
  );
  if (existing) return { status: "success", message: "Sudah ada" };
  sheet.appendRow([
    newId,
    "CEREMONY-DUTY",
    userId,
    todayStr,
    "Melaksanakan Tugas Pembina Upacara",
    0,
    0,
    "Dikonfirmasi",
    5,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm"),
  ]);
  return { status: "success" };
}

function revokeCeremonyLeader(token, userId) {
  const user = verifySession(token);
  if (!user) return { status: "error" };
  const todayStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const sheet = getSheet("Teaching_Logs");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rowDate = Utilities.formatDate(
      new Date(rows[i][3]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    if (
      String(rows[i][2]) === String(userId) &&
      rowDate === todayStr &&
      rows[i][1] === "CEREMONY-DUTY"
    ) {
      sheet.deleteRow(i + 1);
      return { status: "success" };
    }
  }
  return { status: "error" };
}

function substituteCeremonyLeader(token, subId) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const isAdmin = String(user.role).toLowerCase() === "admin";
  const isPicketToday = _isPicketOfficer(user);
  if (!isAdmin && !isPicketToday) {
    return {
      status: "error",
      message: "Hanya admin atau guru piket hari ini yang berwenang.",
    };
  }
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const sheet = getSheet("Ceremony_Schedules");
  const rows = sheet.getDataRange().getValues();
  let targetRow = -1;
  let originalUserId = null;
  for (let i = 1; i < rows.length; i++) {
    const rowDate = Utilities.formatDate(
      new Date(rows[i][1]),
      tz,
      "yyyy-MM-dd",
    );
    if (rowDate === todayStr) {
      targetRow = i + 1;
      originalUserId = rows[i][2];
      break;
    }
  }
  if (targetRow === -1) {
    return {
      status: "error",
      message: "Jadwal upacara hari ini tidak ditemukan.",
    };
  }
  if (String(subId) === String(originalUserId)) {
    return {
      status: "error",
      message: "Guru tidak dapat menggantikan dirinya sendiri.",
    };
  }
  const subUser = findData("Users", "id", subId);
  if (!subUser)
    return { status: "error", message: "Guru pengganti tidak ditemukan." };
  if (String(subUser.role).toLowerCase() === "admin") {
    return {
      status: "error",
      message: "Admin tidak dapat ditugaskan sebagai pembina upacara.",
    };
  }
  const logSheet = getSheet("Teaching_Logs");
  const logRows = logSheet.getDataRange().getValues();
  for (let i = logRows.length - 1; i >= 1; i--) {
    try {
      const logDate = Utilities.formatDate(
        new Date(logRows[i][3]),
        tz,
        "yyyy-MM-dd",
      );
      if (
        logDate === todayStr &&
        String(logRows[i][1]) === "CEREMONY-DUTY" &&
        String(logRows[i][2]) === String(originalUserId)
      ) {
        logSheet.deleteRow(i + 1);
        break;
      }
    } catch (e) {}
  }
  sheet.getRange(targetRow, 3).setValue(subId);
  return {
    status: "success",
    message:
      "Pembina upacara hari ini diganti menjadi " +
      subUser.full_name +
      ". Pembina baru perlu konfirmasi kehadiran.",
  };
}

function assignSubstitute(token, scheduleId, subId) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const todayStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const sheet = getSheet("Substitutes");
  const existingSubs = getData("Substitutes").filter((s) => {
    const subDate = Utilities.formatDate(
      new Date(s.date),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    return String(s.schedule_id) === String(scheduleId) && subDate === todayStr;
  });
  if (existingSubs.length > 0) {
    const existingSub = existingSubs[0];
    const existingUser = findData(
      "Users",
      "id",
      existingSub.substitute_user_id,
    );
    const existingName = existingUser
      ? existingUser.full_name
      : "Guru Pengganti";
    return {
      status: "error",
      code: "COLLISION",
      collision: true,
      existing_sub_id: String(existingSub.substitute_user_id || ""),
      existing_sub_name: existingName,
      message: `Jadwal ini sudah memiliki guru pengganti: ${existingName}. Silakan batalkan terlebih dahulu jika ingin mengganti.`,
    };
  }
  const sched = findData("Schedules", "id", scheduleId);
  if (!sched) return { status: "error", message: "Jadwal tidak ditemukan." };
  const originalUser = sched.user_id;
  if (String(subId) === String(originalUser)) {
    return {
      status: "error",
      message:
        "Guru tidak dapat ditugaskan sebagai pengganti untuk jadwalnya sendiri.",
    };
  }
  const subUser = findData("Users", "id", subId);
  if (!subUser)
    return { status: "error", message: "Guru pengganti tidak ditemukan." };
  if (String(subUser.role).toLowerCase() === "admin") {
    return {
      status: "error",
      message: "Admin tidak dapat ditugaskan sebagai pengganti mengajar.",
    };
  }
  const newId = generateId("SUB");
  sheet.appendRow([newId, todayStr, originalUser, subId, scheduleId]);
  try {
    var sched2 = sched || {};
    var scheduleInfo =
      (sched2.subject ? String(sched2.subject) : "") +
      (sched2.class_name ? " · Kelas " + sched2.class_name : "");
    var byName = String(
      user.full_name ||
        user.name ||
        user.username ||
        (String(user.role).toLowerCase() === "admin" ? "Admin" : "Guru Piket"),
    );
    _notifKbmSubstituted_(originalUser, subId, scheduleInfo, byName);
  } catch (_) {}
  return {
    status: "success",
    message: `Berhasil menetapkan ${subUser.full_name} sebagai guru pengganti.`,
    data: {
      substitute_id: newId,
      original_user: originalUser,
      substitute_user: subId,
      schedule_id: scheduleId,
    },
  };
}

function replaceSubstitute(token, scheduleId, newSubId) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const isAdmin = String(user.role).toLowerCase() === "admin";
    const isPicketToday = _isPicketOfficer(user);
    if (!isAdmin && !isPicketToday) {
      return {
        status: "error",
        message:
          "Hanya admin atau guru piket hari ini yang dapat mengganti pengganti.",
      };
    }
    const revokeRes = revokeSubstitute(token, scheduleId);
    if (revokeRes.status !== "success") {
      if (
        !revokeRes.message ||
        revokeRes.message.indexOf("tidak ditemukan") === -1
      ) {
        return revokeRes;
      }
    }
    const assignRes = assignSubstitute(token, scheduleId, newSubId);
    if (assignRes.status !== "success") return assignRes;
    return {
      status: "success",
      message: "Pengganti berhasil diganti.",
      data: assignRes.data,
    };
  } catch (e) {
    return { status: "error", message: "Server error: " + e.toString() };
  }
}

function revokeSubstitute(token, scheduleId) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const todayStr = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    const isAdmin = String(user.role).toLowerCase() === "admin";
    const isPicketToday = _isPicketOfficer(user);
    if (!isAdmin && !isPicketToday) {
      return {
        status: "error",
        message:
          "Hanya admin atau guru piket hari ini yang dapat membatalkan guru pengganti.",
      };
    }
    const sheet = getSheet("Substitutes");
    const rows = sheet.getDataRange().getValues();
    let deletedRow = -1;
    let deletedData = null;
    for (let i = 1; i < rows.length; i++) {
      const rowDate = Utilities.formatDate(
        new Date(rows[i][1]),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
      if (String(rows[i][4]) === String(scheduleId) && rowDate === todayStr) {
        deletedData = {
          id: rows[i][0],
          date: rows[i][1],
          original_user_id: rows[i][2],
          substitute_user_id: rows[i][3],
          schedule_id: rows[i][4],
        };
        sheet.deleteRow(i + 1);
        deletedRow = i;
        break;
      }
    }
    if (deletedRow === -1) {
      return {
        status: "error",
        message: "Data guru pengganti tidak ditemukan atau sudah dibatalkan.",
      };
    }
    const logsSheet = getSheet("Teaching_Logs");
    const logRows = logsSheet.getDataRange().getValues();
    const picketSchedules = getData("Picket_Schedules");
    const isPicketSub = picketSchedules.some(
      (p) => String(p.id) === String(scheduleId),
    );
    const targetLogScheduleId = isPicketSub
      ? "PICKET-DUTY"
      : String(scheduleId);
    for (let i = 1; i < logRows.length; i++) {
      const logDate = Utilities.formatDate(
        new Date(logRows[i][3]),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
      const logScheduleId = String(logRows[i][1]);
      const logUserId = String(logRows[i][2]);
      if (
        logDate === todayStr &&
        logScheduleId === targetLogScheduleId &&
        logUserId === String(deletedData.substitute_user_id)
      ) {
        logsSheet.deleteRow(i + 1);
        break;
      }
    }
    const attSheet = getSheet("Daily_Attendance");
    const attRows = attSheet.getDataRange().getValues();
    const attSchedColIdx = _attendanceScheduleColIndex_(attSheet);
    for (let i = attRows.length - 1; i >= 1; i--) {
      try {
        const attDate = Utilities.formatDate(
          new Date(attRows[i][1]),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd",
        );
        const attSched = String(
          attRows[i][attSchedColIdx] === undefined
            ? ""
            : attRows[i][attSchedColIdx],
        );
        const isTargetSched = isPicketSub
          ? attSched.trim() === ""
          : attSched === String(scheduleId);
        if (
          String(attRows[i][2]) === String(deletedData.substitute_user_id) &&
          attDate === todayStr &&
          isTargetSched
        ) {
          attSheet.deleteRow(i + 1);
          break;
        }
      } catch (e) {}
    }
    const allUsers = getData("Users");
    const subUser = allUsers.find(
      (u) => String(u.id) === String(deletedData.substitute_user_id),
    );
    const origUser = allUsers.find(
      (u) => String(u.id) === String(deletedData.original_user_id),
    );
    // Jika ini pengganti piket, bersihkan JTM adjustment & reallocations-nya agar tidak orphan
    if (isPicketSub) {
      try {
        const piketRef = String(deletedData.substitute_user_id) + "|" + todayStr;
        const reallocations = _jtmReadReallocations_("PIKET", piketRef) || [];
        for (var _ri = 0; _ri < reallocations.length; _ri++) {
          try { _examDeleteLogByRef_(reallocations[_ri].id); } catch (_) {}
        }
        _jtmDeleteAllReallocationsForOccurrence_("PIKET", piketRef);
        _jtmDeleteAdjustment_("PIKET", piketRef);
      } catch (_) {}
    }
    try { _invalidateDataSnapshot(); } catch (_) {}
    try {
      var sched3 = findData("Schedules", "id", scheduleId) || {};
      var scheduleInfo =
        (sched3.subject ? String(sched3.subject) : "") +
        (sched3.class_name ? " · Kelas " + sched3.class_name : "");
      var byName = String(
        user.full_name ||
          user.name ||
          user.username ||
          (isAdmin ? "Admin" : "Guru Piket"),
      );
      _notifKbmSubCancelled_(
        deletedData.substitute_user_id,
        scheduleInfo,
        byName,
      );
    } catch (_) {}
    return {
      status: "success",
      message: `Berhasil membatalkan guru pengganti ${subUser ? subUser.full_name : ""}. ${origUser ? origUser.full_name : "Guru asli"} kembali bertugas.`,
      data: deletedData,
      schedule_id: scheduleId,
    };
  } catch (error) {
    return { status: "error", message: "Server error: " + error.toString() };
  }
}

function getTodaySubstitutes(token) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const todayStr = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    const substitutes = getData("Substitutes");
    const users = getData("Users");
    const schedules = getData("Schedules");
    const attendanceData = getData("Daily_Attendance");
    const todaySubs = substitutes.filter((s) => {
      const subDate = safeDate(s.date);
      const isNotPicket = String(s.schedule_id).indexOf("PKT") === -1;
      return subDate === todayStr && isNotPicket;
    });
    const formattedSubs = todaySubs.map((s) => {
      const subUser = users.find(
        (u) => String(u.id) === String(s.substitute_user_id),
      );
      const origUser = users.find(
        (u) => String(u.id) === String(s.original_user_id),
      );
      const sched = schedules.find(
        (sch) => String(sch.id) === String(s.schedule_id),
      );
      const attRecord = attendanceData.find(
        (a) =>
          String(a.user_id) === String(s.substitute_user_id) &&
          safeDate(a.date) === todayStr &&
          String(a.schedule_id || "") === String(s.schedule_id),
      );
      let timeStr = "-";
      if (sched) {
        const tStart = safeTime(sched.time_start);
        const tEnd = safeTime(sched.time_end);
        timeStr = `${tStart.substring(0, 5)} - ${tEnd.substring(0, 5)}`;
      }
      return {
        id: s.id,
        substitute_name: subUser ? subUser.full_name : "Unknown",
        substitute_user_id: s.substitute_user_id,
        original_name: origUser ? origUser.full_name : "Unknown",
        schedule_info: sched
          ? `${sched.subject} - Kelas ${sched.class_name}`
          : "Unknown",
        time: timeStr,
        schedule_id: s.schedule_id,
        class_name: sched ? sched.class_name : "-",
        subject: sched ? sched.subject : "-",
        is_confirmed: !!attRecord,
      };
    });
    formattedSubs.sort((a, b) => a.time.localeCompare(b.time));
    return {
      status: "success",
      data: formattedSubs,
      count: formattedSubs.length,
    };
  } catch (error) {
    return {
      status: "error",
      message: "Server error: " + error.toString(),
    };
  }
}

function safeDate(val) {
  if (!val) return "";
  if (val instanceof Date)
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(val);
}

function safeTime(val) {
  if (!val) return "";
  if (val instanceof Date)
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  return String(val);
}
