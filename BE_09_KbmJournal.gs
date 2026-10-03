// SiM-Guru — BE_09_KbmJournal.gs
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
