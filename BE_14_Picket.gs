// SiM-Guru — BE_14_Picket.gs
function getPicketPageData(token) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const picketSheet = getSheet("Picket_Schedules");
    const pickets = picketSheet.getDataRange().getValues();
    const users = getData("Users");
    const logs = getData("Teaching_Logs");
    const schedules = getData("Schedules");
    const ceremonies = getData("Ceremony_Schedules");
    const attendance = getData("Daily_Attendance");
    const picketSchedules = getData("Picket_Schedules");
    const substitutes = getData("Substitutes");
    const scheduleData = [];
    const days = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    for (let i = 1; i < pickets.length; i++) {
      if (pickets[i].length < 3) continue;
      const u = users.find((usr) => String(usr.id) === String(pickets[i][2]));
      scheduleData.push({
        id: String(pickets[i][0]),
        day_index: parseInt(pickets[i][1]),
        day_name: days[parseInt(pickets[i][1])],
        user_id: String(pickets[i][2]),
        name: u ? u.full_name : "Unknown",
      });
    }
    scheduleData.sort((a, b) => a.day_index - b.day_index);
    const today = new Date();
    const dayIdx = _getIndoDayIndex();
    const todayDateStr = Utilities.formatDate(
      today,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    const todayPicketList = scheduleData.filter(
      (s) => String(s.day_index) === String(dayIdx),
    );
    const todayPickets = todayPicketList.map((p) => {
      const isPresent = logs.some(
        (l) =>
          String(l.user_id) === String(p.user_id) &&
          String(l.schedule_id) === "PICKET-DUTY" &&
          safeDate(l.date) === todayDateStr,
      );
      return { ...p, is_present: isPresent };
    });
    const isPicketToday = todayPickets.some(
      (p) => String(p.user_id) === String(user.id),
    );
    const cfg = _getConfigMap();
    const activeTP = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    const allSched = schedules.filter((s) => {
      const isToday = String(s.day_index) === String(dayIdx);
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return isToday && sTP === activeTP && sSem === activeSem;
    });
    const todayAtt = attendance.filter(
      (a) => safeDate(a.date) === todayDateStr,
    );
    const todaySubstitutes = substitutes.filter((s) => {
      const subDate = safeDate(s.date);
      return subDate === todayDateStr;
    });
    allSched.sort((a, b) => {
      return safeTime(a.time_start).localeCompare(safeTime(b.time_start));
    });
    var dailySchedules = allSched.map(function (s) {
      var guru = users.find(function (u) {
        return String(u.id) === String(s.user_id);
      });
      var isPres = todayAtt.some(function (a) {
        return (
          String(a.user_id) === String(s.user_id) &&
          String(a.schedule_id || "") === String(s.id) &&
          a.status === "Present"
        );
      });
      var subRecord = todaySubstitutes.find(function (sub) {
        return String(sub.schedule_id) === String(s.id);
      });
      var substituteInfo = null;
      if (subRecord) {
        var subUser = users.find(function (u) {
          return String(u.id) === String(subRecord.substitute_user_id);
        });
        substituteInfo = {
          id: subRecord.id,
          user_id: subRecord.substitute_user_id,
          name: subUser ? subUser.full_name : "Unknown",
        };
      }
      var tStart = safeTime(s.time_start);
      if (tStart.length > 5) tStart = tStart.substring(0, 5);
      var tEnd = safeTime(s.time_end);
      if (tEnd.length > 5) tEnd = tEnd.substring(0, 5);
      return {
        schedule_id: String(s.id),
        time: tStart + " - " + tEnd,
        class_name: String(s.class_name || "-"),
        subject: String(s.subject || "-"),
        name: guru ? guru.full_name : "Unknown",
        user_id: String(s.user_id),
        is_present: isPres,
        substitute_info: substituteInfo,
      };
    });
    const todayCeremony = ceremonies.find(
      (c) => safeDate(c.date) === todayDateStr,
    );
    let ceremonyInfo = {
      is_ceremony_day: dayIdx === 1,
      data: null,
      upcoming: [],
    };
    if (todayCeremony) {
      const cLeader = users.find(
        (u) => String(u.id) === String(todayCeremony.user_id),
      );
      const isConfirmed = logs.some(
        (l) =>
          String(l.schedule_id) === "CEREMONY-DUTY" &&
          safeDate(l.date) === todayDateStr &&
          String(l.user_id) === String(todayCeremony.user_id),
      );
      ceremonyInfo.data = {
        id: String(todayCeremony.id),
        user_id: String(todayCeremony.user_id),
        name: cLeader ? cLeader.full_name : "Unknown",
        is_confirmed: isConfirmed,
      };
    }
    const upcoming = ceremonies
      .filter((c) => new Date(c.date) >= today)
      .sort((a, b) => new Date(a.date) - new Date(b.date))
      .slice(0, 5)
      .map((c) => {
        const u = users.find((usr) => String(usr.id) === String(c.user_id));
        return {
          id: String(c.id),
          date: safeDate(c.date),
          name: u ? u.full_name : "Unknown",
        };
      });
    ceremonyInfo.upcoming = upcoming;
    const myCeremonies = ceremonies
      .filter((c) => String(c.user_id).trim() === String(user.id).trim())
      .map((c) => {
        const dateStr = safeDate(c.date);
        const cDate = new Date(dateStr + "T00:00:00");
        const tDate = new Date(todayDateStr + "T00:00:00");
        const isPast = cDate < tDate;
        const isToday = dateStr === todayDateStr;
        const isConfirmed = logs.some(
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
    const originalPicketIds = new Set(
      todayPicketList.map((p) => String(p.user_id)),
    );
    const todayPicketSubRecords = substitutes.filter((s) => {
      const subDate = safeDate(s.date);
      return (
        subDate === todayDateStr &&
        picketSchedules.some(
          (p) =>
            String(p.id) === String(s.schedule_id) &&
            String(p.day_index) === String(dayIdx),
        )
      );
    });
    const todayPicketSubstitutes = todayPicketSubRecords.map((s) => {
      const subUser = users.find(
        (u) => String(u.id) === String(s.substitute_user_id),
      );
      const origUser = users.find(
        (u) => String(u.id) === String(s.original_user_id),
      );
      const noteStr = "";
      const logRec = logs.find(
        (l) =>
          String(l.user_id) === String(s.substitute_user_id) &&
          String(l.schedule_id) === "PICKET-DUTY" &&
          safeDate(l.date) === todayDateStr,
      );
      return {
        log_id: logRec ? String(logRec.log_id) : "",
        schedule_id: String(s.schedule_id),
        sub_id: String(s.id),
        user_id: String(s.substitute_user_id),
        name: subUser ? subUser.full_name : "Unknown",
        original_name: origUser ? origUser.full_name : "Petugas Asli",
        notes: noteStr,
        is_confirmed: !!logRec,
      };
    });
    return {
      status: "success",
      is_admin: user.role === "admin",
      is_picket_today:
        isPicketToday ||
        todayPicketSubstitutes.some(
          (s) => String(s.user_id) === String(user.id),
        ) ||
        logs.some(
          (l) =>
            String(l.user_id) === String(user.id) &&
            String(l.schedule_id) === "PARTIAL-SUB-PIKET" &&
            safeDate(l.date) === todayDateStr,
        ),
      current_user_id: String(user.id),
      picket_schedule: scheduleData,
      today_pickets: todayPickets,
      daily_schedules: dailySchedules,
      ceremony_info: ceremonyInfo,
      my_ceremonies: myCeremonies,
      today_picket_subs: todayPicketSubstitutes,
    };
  } catch (error) {
    return {
      status: "error",
      message: "Server Error (Picket): " + error.toString(),
    };
  }
}

function _isPicketOfficer(user) {
  const dayIndex = _getIndoDayIndex();
  const picketSchedules = getData("Picket_Schedules");
  const isRegularPicket = picketSchedules.some(
    (p) =>
      String(p.user_id) === String(user.id) &&
      String(p.day_index).trim() === String(dayIndex),
  );
  if (isRegularPicket) return true;
  const todayStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const substitutes = getData("Substitutes");
  const isAssignedSubstitute = substitutes.some((s) => {
    if (String(s.substitute_user_id) !== String(user.id)) return false;
    if (safeDate(s.date) !== todayStr) return false;
    return picketSchedules.some((p) => String(p.id) === String(s.schedule_id));
  });
  if (isAssignedSubstitute) {
    const logs = getData("Teaching_Logs");
    return logs.some(
      (l) =>
        String(l.user_id) === String(user.id) &&
        String(l.schedule_id) === "PICKET-DUTY" &&
        safeDate(l.date) === todayStr,
    );
  }
  const logsAll = getData("Teaching_Logs");
  // Guru piket biasa yang sudah dikonfirmasi hari ini
  if (logsAll.some(
    (l) =>
      String(l.user_id) === String(user.id) &&
      String(l.schedule_id) === "PICKET-DUTY" &&
      safeDate(l.date) === todayStr,
  )) return true;
  // Guru penerima relokasi JTM piket (PARTIAL-SUB-PIKET) hari ini
  // juga mendapat wewenang piket seperti guru piket pengganti
  return logsAll.some(
    (l) =>
      String(l.user_id) === String(user.id) &&
      String(l.schedule_id) === "PARTIAL-SUB-PIKET" &&
      safeDate(l.date) === todayStr,
  );
}

function _hasTimeInToday(userId) {
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const allAttendance = getData("Daily_Attendance");
  const callerRecord = allAttendance.find(
    (a) => String(a.user_id) === String(userId) && String(a.date) === todayStr,
  );
  return callerRecord && String(callerRecord.time_in || "").trim() !== "";
}

function savePicketSchedule(token, data) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const dayIdx = parseInt(data && data.day_index, 10);
  const userId = String((data && data.user_id) || "").trim();
  if (isNaN(dayIdx) || dayIdx < 0 || dayIdx > 6) {
    return { status: "error", message: "Hari tidak valid." };
  }
  if (!userId) {
    return { status: "error", message: "Pilih guru terlebih dahulu." };
  }
  const targetUser = findData("Users", "id", userId);
  if (!targetUser) return { status: "error", message: "Guru tidak ditemukan." };
  if (String(targetUser.role).toLowerCase() === "admin") {
    return {
      status: "error",
      message: "Admin tidak dapat dijadwalkan sebagai piket.",
    };
  }
  const existing = getData("Picket_Schedules");
  const dup = existing.find(
    (p) =>
      String(p.user_id) === userId &&
      String(p.day_index).trim() === String(dayIdx),
  );
  if (dup) {
    return {
      status: "error",
      code: "DUPLICATE",
      message: "Guru ini sudah terdaftar piket pada hari yang sama.",
    };
  }
  const sheet = getSheet("Picket_Schedules");
  const newId = generateId("PKT");
  sheet.appendRow([newId, dayIdx, userId]);
  return { status: "success", message: "Jadwal piket berhasil disimpan." };
}

function revokePicketDaily(token, userId) {
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
      rows[i][1] === "PICKET-DUTY"
    ) {
      sheet.deleteRow(i + 1);
      // Bersihkan JTM adjustment & reallocations yang terkait agar tidak orphan
      try {
        const piketRef = String(userId) + "|" + todayStr;
        const reallocations = _jtmReadReallocations_("PIKET", piketRef) || [];
        for (var _ri = 0; _ri < reallocations.length; _ri++) {
          try { _examDeleteLogByRef_(reallocations[_ri].id); } catch (_) {}
        }
        _jtmDeleteAllReallocationsForOccurrence_("PIKET", piketRef);
        _jtmDeleteAdjustment_("PIKET", piketRef);
      } catch (_) {}
      // Hapus juga baris Daily_Attendance piket agar jam masuk/pulang
      // tidak muncul lagi setelah konfirmasi dibatalkan.
      try {
        const attSheet = getSheet("Daily_Attendance");
        const attSchedColIdx = _attendanceScheduleColIndex_(attSheet);
        const attRows = attSheet.getDataRange().getValues();
        for (var _ai = attRows.length - 1; _ai >= 1; _ai--) {
          var attUserId = String(attRows[_ai][2] || "");
          var attDateRaw = attRows[_ai][1];
          var attDate = "";
          try {
            attDate = Utilities.formatDate(new Date(attDateRaw), Session.getScriptTimeZone(), "yyyy-MM-dd");
          } catch (_e) { attDate = String(attDateRaw); }
          var attSched = String(attRows[_ai][attSchedColIdx] === undefined ? "" : attRows[_ai][attSchedColIdx]);
          if (attUserId === String(userId) && attDate === todayStr &&
              (attSched === "" || attSched === "PICKET-DUTY" || attSched === "PIKET")) {
            attSheet.deleteRow(_ai + 1);
            break; // hanya satu baris piket per user per hari
          }
        }
      } catch (_) {}
      try { _invalidateDataSnapshot(); } catch (_) {}
      try {
        _notifPicketRevoked_(userId);
      } catch (_) {}
      return { status: "success" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}

function substitutePicketDaily(token, originalUserId, substituteUserId) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const isAdmin = String(user.role).toLowerCase() === "admin";
  const isPicketToday = _isPicketOfficer(user);
  if (!isAdmin && !isPicketToday) {
    return {
      status: "error",
      message:
        "Hanya admin atau guru piket hari ini yang berwenang menetapkan pengganti piket.",
    };
  }
  if (String(substituteUserId) === String(originalUserId)) {
    return {
      status: "error",
      message: "Guru tidak dapat menggantikan dirinya sendiri.",
    };
  }
  const users = getData("Users");
  const origUser = users.find((u) => String(u.id) === String(originalUserId));
  const subUser = users.find((u) => String(u.id) === String(substituteUserId));
  if (!origUser)
    return { status: "error", message: "Petugas asli tidak ditemukan." };
  if (!subUser)
    return { status: "error", message: "Guru pengganti tidak ditemukan." };
  if (String(subUser.role).toLowerCase() === "admin") {
    return {
      status: "error",
      message: "Admin tidak dapat ditugaskan sebagai piket pengganti.",
    };
  }
  const dayIndex = _getIndoDayIndex();
  const picketSchedules = getData("Picket_Schedules");
  const originalPicketSchedule = picketSchedules.find(
    (p) =>
      String(p.user_id) === String(originalUserId) &&
      String(p.day_index).trim() === String(dayIndex),
  );
  if (!originalPicketSchedule) {
    return {
      status: "error",
      message: "Guru asli bukan petugas piket hari ini.",
    };
  }
  const origName = origUser.full_name;
  const sheet = getSheet("Substitutes");
  const todayStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const existingSubs = getData("Substitutes").filter((s) => {
    const subDate = Utilities.formatDate(
      new Date(s.date),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    return (
      String(s.schedule_id) === String(originalPicketSchedule.id) &&
      subDate === todayStr
    );
  });
  if (existingSubs.length > 0) {
    const existingSub = existingSubs[0];
    const existingSubUser = users.find(
      (u) => String(u.id) === String(existingSub.substitute_user_id),
    );
    const existingName = existingSubUser
      ? existingSubUser.full_name
      : "Guru Pengganti";
    return {
      status: "error",
      code: "COLLISION",
      collision: true,
      existing_sub_id: String(existingSub.substitute_user_id || ""),
      existing_sub_name: existingName,
      message:
        "Jadwal ini sudah memiliki guru pengganti: " +
        existingName +
        ". Silakan batalkan terlebih dahulu.",
    };
  }
  const newId = generateId("SUB");
  sheet.appendRow([
    newId,
    todayStr,
    originalUserId,
    substituteUserId,
    originalPicketSchedule.id,
  ]);
  try {
    _notifPicketSubstituted_(originalUserId, substituteUserId);
  } catch (_) {}
  return {
    status: "success",
    message: subUser.full_name + " ditugaskan sebagai piket pengganti.",
  };
}

function cancelPicketSubstituteAttendance(token, logId) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const sheet = getSheet("Teaching_Logs");
  const rows = sheet.getDataRange().getValues();
  let deletedUserId = null;
  let deletedDate = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(logId)) {
      deletedUserId = String(rows[i][2]);
      try {
        deletedDate = Utilities.formatDate(
          new Date(rows[i][3]),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd",
        );
      } catch (e) {
        deletedDate = String(rows[i][3]);
      }
      sheet.deleteRow(i + 1);
      break;
    }
  }
  if (!deletedUserId) {
    return { status: "error", message: "Data tidak ditemukan" };
  }
  if (deletedDate) {
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
        if (
          String(attRows[i][2]) === deletedUserId &&
          attDate === deletedDate &&
          !attSched
        ) {
          attSheet.deleteRow(i + 1);
          break;
        }
      } catch (e) {}
    }
  }
  try {
    _notifPicketSubCancelled_(deletedUserId);
  } catch (_) {}
  return {
    status: "success",
    message: "Data piket pengganti berhasil dihapus",
  };
}

function deletePicketSchedule(token, id) {
  return deleteRowById(token, "Picket_Schedules", id);
}

function confirmPicketDaily(token, userId) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const isAdmin = String(user.role).toLowerCase() === "admin";
  const isSelf = String(user.id) === String(userId);
  const isPicketToday = _isPicketOfficer(user);
  if (!isAdmin && !isSelf && !isPicketToday) {
    return {
      status: "error",
      message: "Anda tidak berwenang mengonfirmasi kehadiran piket guru lain.",
    };
  }
  const safeDate = (d) =>
    Utilities.formatDate(
      new Date(d),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  const dayIndex = _getIndoDayIndex();
  const picketSchedules = getData("Picket_Schedules");
  const isPicket = picketSchedules.some(
    (p) =>
      String(p.user_id) === String(userId) &&
      String(p.day_index) === String(dayIndex),
  );
  const todayStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const substitutes = getData("Substitutes");
  const picketSubRecord = substitutes.find((s) => {
    return (
      String(s.substitute_user_id) === String(userId) &&
      safeDate(s.date) === todayStr &&
      picketSchedules.some(
        (p) =>
          String(p.id) === String(s.schedule_id) &&
          String(p.day_index) === String(dayIndex),
      )
    );
  });
  if (!isPicket && !picketSubRecord)
    return { status: "error", message: "Anda bukan petugas piket hari ini" };
  let roleLabel;
  if (isAdmin) roleLabel = "Admin";
  else if (isPicketToday && !isSelf) roleLabel = "Guru Piket";
  else if (isSelf) roleLabel = "Diri Sendiri";
  else roleLabel = "Pengguna";
  const recorderName = String(
    user.full_name || user.name || user.username || "",
  ).trim();
  const catatan =
    "Dikonfirmasi oleh " +
    roleLabel +
    (recorderName ? " (" + recorderName + ")" : "");
  const sheet = getSheet("Teaching_Logs");
  const newId = generateId("LOG-PCK");
  const existing = getData("Teaching_Logs").find(
    (l) =>
      safeDate(l.date) === todayStr &&
      String(l.user_id) === String(userId) &&
      l.schedule_id === "PICKET-DUTY",
  );
  if (existing) return { status: "success", message: "Sudah ada" };
  sheet.appendRow([
    newId,
    "PICKET-DUTY",
    userId,
    todayStr,
    "Melaksanakan Tugas Piket",
    picketSubRecord ? 1 : 0,
    picketSubRecord ? picketSubRecord.original_user_id : 0,
    catatan,
    4,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm"),
  ]);
  try {
    _notifPicketConfirmed_(userId, { byRole: roleLabel, byName: recorderName });
  } catch (_) {}
  return { status: "success" };
}

function confirmPicketSubPresenceBackend(token, targetUserId) {
  // Langkah 1: Konfirmasi kehadiran di Daily_Attendance dengan schedule_id = ""
  // (menandai piket). Namun cek dulu apakah baris dengan schedule_id apapun
  // sudah ada hari ini — jika guru juga punya jadwal KBM yang sudah dikonfirmasi
  // (baris dengan schedule_id != ""), kita tidak perlu membuat baris baru lagi
  // untuk menghindari duplikasi di Daily_Attendance.
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const attSheet = getSheet("Daily_Attendance");
  const attSchedColIdx = _attendanceScheduleColIndex_(attSheet);
  const attRows = attSheet.getDataRange().getValues();
  let piketAttExists = false;
  for (let i = 1; i < attRows.length; i++) {
    const rowUserId = String(attRows[i][2] || "");
    let rowDate = "";
    try { rowDate = Utilities.formatDate(new Date(attRows[i][1]), tz, "yyyy-MM-dd"); } catch (_) { rowDate = String(attRows[i][1]); }
    const rowSched = String(attRows[i][attSchedColIdx] === undefined ? "" : attRows[i][attSchedColIdx]);
    if (rowUserId === String(targetUserId) && rowDate === todayStr &&
        (rowSched === "" || rowSched === "PICKET-DUTY" || rowSched === "PIKET")) {
      piketAttExists = true;
      break;
    }
  }
  if (!piketAttExists) {
    const res1 = confirmTeacherPresence(token, targetUserId, "");
    if (res1.status !== "success" && res1.message !== "Sudah dikonfirmasi") {
      return res1;
    }
  }
  // Langkah 2: Catat log piket di Teaching_Logs
  const res2 = confirmPicketDaily(token, targetUserId);
  if (res2.status !== "success" && res2.message !== "Sudah ada") {
    return res2;
  }
  return { status: "success" };
}

function getPicketStatus(token) {
  const user = verifySession(token);
  if (!user) return { status: "error" };
  const todayIndex = _getIndoDayIndex();
  const pickets = getData("Picket_Schedules").filter(
    (p) =>
      Number(p.day_index) === todayIndex &&
      String(p.user_id) === String(user.id),
  );
  return { status: "success", is_picket: pickets.length > 0 };
}
