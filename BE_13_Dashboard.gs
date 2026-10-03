// SiM-Guru — BE_13_Dashboard.gs
function getDashboardStats(token) {
  const user = verifySession(token);
  if (!user) return { status: "error" };
  try {
    const now = new Date();
    const currentMonth = parseInt(Utilities.formatDate(now, "Asia/Jakarta", "M"), 10) - 1;
    const currentYear = parseInt(Utilities.formatDate(now, "Asia/Jakarta", "yyyy"), 10);
    const snap = _getDataSnapshot();
    const userLogs = snap.logsByUserId[user.id] || [];
    const logs = userLogs.filter((l) => {
      const d = new Date(l.date + (String(l.date).length === 10 ? "T00:00:00" : ""));
      return d.getMonth() === currentMonth && d.getFullYear() === currentYear;
    });
    const jtmAdjMap = _jtmBuildKbmAdjustmentMap_();
    const piketAdjMap = _jtmBuildPiketAdjustmentMap_();
    let totalJtmRecorded = 0;
    logs.forEach((l) => {
      const _sid = String(l.schedule_id);
      const isAutoLibur = String(l.notes).includes(
        "Auto-generated: Libur Bonus",
      );
      if (isAutoLibur) {
      } else if (_sid === "PICKET-DUTY" || _sid === "PIKET") {
        totalJtmRecorded += _jtmResolvePiketLogJtm_(piketAdjMap, l.user_id, l.date, 4);
      } else if (_sid === "CEREMONY-DUTY" || _sid === "UPACARA") {
        totalJtmRecorded += 5;
      } else if (
        _sid === "EXAM-SUPERVISOR" ||
        _sid === "EXAM-COMMITTEE" ||
        _sid === "PARTIAL-SUB-KBM" ||
        _sid === "PARTIAL-SUB-EXAM"
      ) {
        totalJtmRecorded += Number(l.jtm_val || 0);
      } else {
        const s = snap.schedulesById[_sid];
        if (s) {
          totalJtmRecorded += _jtmResolveKbmLogJtm_(
            jtmAdjMap,
            _sid,
            l.date,
            Number(s.jtm_val || 0),
          );
        } else {
          totalJtmRecorded += Number(l.jtm_val || 0);
        }
      }
    });
    const userEventLogs = snap.eventAttendanceByUserId[user.id] || [];
    const eventLogs = userEventLogs.filter((l) => {
      const d = new Date(l.date + (String(l.date).length === 10 ? "T00:00:00" : ""));
      return d.getMonth() === currentMonth && d.getFullYear() === currentYear;
    });
    const totalEventJtm = _computeEventJtm(eventLogs);
    const bonusDetail = _calculateMonthlyBonusFast_Detailed(
      user.id,
      currentMonth,
      currentYear,
      snap,
    );
    const autoRecordedBonus = bonusDetail.auto;
    const totalPendingBonus = bonusDetail.estimated;
    const displayTotalBonus = autoRecordedBonus + totalPendingBonus;
    const grandTotalJtm = totalJtmRecorded + totalEventJtm + displayTotalBonus;
    const allowances = snap.allowancesByUserId[user.id] || [];
    const totalAllowance = allowances.reduce(
      (sum, a) => sum + Number(a.amount || 0),
      0,
    );
    const baseSalary = Number(snap.config["base_salary"] || 0);
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
    const currentPeriodId = `${months[currentMonth]} ${currentYear}`;
    const transportEnabled =
      String(snap.config["transport_allowance_enabled"] || "false")
        .toLowerCase()
        .trim() === "true";
    const transportDesimal = transportEnabled
      ? _getMonthlyTransportSumFast(String(user.id), currentPeriodId, snap)
      : 0;
    const transportTotal = _applyThreePointRounding(transportDesimal);
    const grandTotalHonor =
      grandTotalJtm * baseSalary + totalAllowance + transportTotal;
    const todayStr = Utilities.formatDate(
      now,
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    let jtmToday = 0;
    for (let i = 0; i < logs.length; i++) {
      const l = logs[i];
      const dStr = Utilities.formatDate(
        new Date(l.date + (String(l.date).length === 10 ? "T00:00:00" : "")),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (dStr !== todayStr) continue;
      const _sid = String(l.schedule_id);
      const isAutoLibur = String(l.notes).includes("Auto-generated: Libur Bonus");
      if (isAutoLibur) continue;
      if (_sid === "PICKET-DUTY" || _sid === "PIKET") {
        jtmToday += _jtmResolvePiketLogJtm_(piketAdjMap, l.user_id, l.date, 4);
      } else if (_sid === "CEREMONY-DUTY" || _sid === "UPACARA") {
        jtmToday += 5;
      } else if (
        _sid === "EXAM-SUPERVISOR" || _sid === "EXAM-COMMITTEE" ||
        _sid === "PARTIAL-SUB-KBM" || _sid === "PARTIAL-SUB-EXAM" ||
        _sid === "PARTIAL-SUB-PIKET"
      ) {
        jtmToday += Number(l.jtm_val || 0);
      } else {
        const s = snap.schedulesById[_sid];
        if (s) {
          jtmToday += _jtmResolveKbmLogJtm_(jtmAdjMap, _sid, l.date, Number(s.jtm_val || 0));
        } else {
          jtmToday += Number(l.jtm_val || 0);
        }
      }
    }
    for (let i = 0; i < eventLogs.length; i++) {
      const eLog = eventLogs[i];
      if (
        eLog.journal_submitted !== true &&
        String(eLog.journal_submitted).toLowerCase() !== "true"
      )
        continue;
      const dStr = Utilities.formatDate(
        new Date(eLog.date + (String(eLog.date).length === 10 ? "T00:00:00" : "")),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (dStr === todayStr) {
        const val = Number(eLog.jtm_val);
        if (!isNaN(val) && val > 0) jtmToday += val;
      }
    }
    return {
      status: "success",
      total_jtm: grandTotalJtm,
      bonus_jtm: displayTotalBonus,
      est_honor_formatted: formatRupiah(grandTotalHonor),
      jtm_hari_ini: jtmToday,
      transport_monthly: transportDesimal,
      transport_monthly_formatted: formatRupiah(transportDesimal),
      transport_enabled: transportEnabled,
      breakdown: {
        auto_libur_jtm: autoRecordedBonus,
        estimated_remaining: totalPendingBonus,
        event_jtm: totalEventJtm,
      },
    };
  } catch (e) {
    return {
      status: "error",
      message: "Gagal memuat statistik: " + (e && e.message ? e.message : e),
    };
  }
}

function _calculateMonthlyBonusFast_Detailed(userId, month, year, snapshot) {
  const cutoffDate = new Date();
  cutoffDate.setHours(23, 59, 59, 999);
  const holidays = snapshot.holidays.filter((h) => {
    const d = new Date(h.date);
    return (
      d.getMonth() === month &&
      d.getFullYear() === year &&
      String(h.is_holiday).toLowerCase() === "true"
    );
  });
  if (holidays.length === 0) return { auto: 0, estimated: 0 };
  const userLogs = snapshot.logsByUserId[userId] || [];
  const autoLogJtmByDate = {};
  for (let i = 0; i < userLogs.length; i++) {
    const l = userLogs[i];
    if (String(l.notes).indexOf("Auto-generated: Libur Bonus") < 0) continue;
    const ld = new Date(l.date);
    if (ld.getMonth() !== month || ld.getFullYear() !== year) continue;
    const dStr = Utilities.formatDate(ld, "Asia/Jakarta", "yyyy-MM-dd");
    autoLogJtmByDate[dStr] =
      (autoLogJtmByDate[dStr] || 0) + Number(l.jtm_val || 0);
  }
  // Filter jadwal: hanya user yang bersangkutan DAN semester/tahun pelajaran aktif
  const activeTP = (snapshot.config && snapshot.config["tahun_pelajaran"]) || "";
  const activeSem = (snapshot.config && snapshot.config["semester"]) || "";
  const userSchedules = snapshot.schedules.filter((s) => {
    if (String(s.user_id) !== String(userId)) return false;
    const sTP = s.tahun_pelajaran || activeTP;
    const sSem = s.semester || activeSem;
    return sTP === activeTP && sSem === activeSem;
  });
  const schedulesByDay = {};
  userSchedules.forEach((s) => {
    const k = Number(s.day_index);
    if (!schedulesByDay[k]) schedulesByDay[k] = [];
    schedulesByDay[k].push(s);
  });
  let estimated = 0;
  holidays.forEach((h) => {
    const d = new Date(h.date);
    const dStr = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
    if (autoLogJtmByDate[dStr]) return;
    const dayIdx = d.getDay();
    const list = schedulesByDay[dayIdx] || [];
    list.forEach((s) => {
      estimated += Number(s.jtm_val || 0);
    });
  });
  let auto = 0;
  Object.keys(autoLogJtmByDate).forEach((k) => {
    auto += Number(autoLogJtmByDate[k] || 0);
  });
  return { auto: auto, estimated: estimated };
}

function _getIndoDayIndex() {
  const now = new Date();
  const utc = now.getTime() + now.getTimezoneOffset() * 60000;
  const jakartaTime = new Date(utc + 7 * 3600000);
  return jakartaTime.getDay();
}

function _getConfigMap() {
  const data = getData("Config");
  let conf = {};
  data.forEach((d) => (conf[d.key] = d.value));
  return conf;
}

function getTodaySchedule(token) {
  try {
    const user = verifySession(token);
    const dayIndex = _getIndoDayIndex();
    const snap = _getDataSnapshot();
    const allSchedules = snap.schedules;
    const allLogs = snap.logs;
    const attendanceData = snap.attendance;
    const substitutes = snap.substitutes;
    const todayStr = Utilities.formatDate(
      new Date(),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    let config = snap.config;
    const activeTP = config.tahun_pelajaran || "";
    const activeSem = config.semester || "";
    let mySchedules = allSchedules.filter((s) => {
      const isMe = String(s.user_id).trim() === String(user.id).trim();
      const isToday = String(s.day_index).trim() === String(dayIndex);
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return isMe && isToday && sTP === activeTP && sSem === activeSem;
    });
    const mySubJobs = substitutes.filter(
      (s) =>
        String(s.substitute_user_id).trim() === String(user.id).trim() &&
        Utilities.formatDate(new Date(s.date + (String(s.date).length === 10 ? "T00:00:00" : "")), "Asia/Jakarta", "yyyy-MM-dd") ===
          todayStr,
    );
    mySubJobs.forEach((sub) => {
      const sched = allSchedules.find(
        (s) => String(s.id) === String(sub.schedule_id),
      );
      if (sched) {
        const subSched = { ...sched, is_substitute_job: true };
        mySchedules.push(subSched);
      }
    });
    const finalResult = mySchedules.map((s) => {
      let start = s.time_start;
      let end = s.time_end;
      if (start instanceof Date)
        start = Utilities.formatDate(start, "Asia/Jakarta", "HH:mm");
      if (end instanceof Date)
        end = Utilities.formatDate(end, "Asia/Jakarta", "HH:mm");
      start = String(start).substring(0, 5);
      end = String(end).substring(0, 5);
      const isDone = allLogs.some((log) => {
        const logDate = Utilities.formatDate(
          new Date(log.date + (String(log.date).length === 10 ? "T00:00:00" : "")),
          "Asia/Jakarta",
          "yyyy-MM-dd",
        );
        return String(log.schedule_id) === String(s.id) && logDate === todayStr;
      });
      const attendanceRecord = attendanceData.find(
        (a) =>
          String(a.user_id).trim() === String(user.id).trim() &&
          Utilities.formatDate(
            new Date(a.date + (String(a.date).length === 10 ? "T00:00:00" : "")),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          ) === todayStr &&
          String(a.schedule_id || "") === String(s.id),
      );
      let subjectLabel = s.subject;
      if (s.is_substitute_job) subjectLabel += " (GURU PENGGANTI)";
      return {
        id: s.id,
        time_start: start,
        time_end: end,
        class_name: String(s.class_name || "-"),
        subject: subjectLabel,
        jtm_val: String(s.jtm_val || "0"),
        has_log: isDone,
        attendance_confirmed: !!attendanceRecord,
      };
    });
    finalResult.sort((a, b) => a.time_start.localeCompare(b.time_start));
    return finalResult;
  } catch (e) {
    return [];
  }
}

function deleteTeachingLog(token, id) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    const isAdmin = String(user.role).toLowerCase() === "admin";
    // Guru hanya boleh hapus jurnal miliknya sendiri; admin boleh hapus milik siapapun
    if (!isAdmin) {
      const logs = getData("Teaching_Logs");
      const log = logs.find((l) => String(l.log_id) === String(id));
      if (!log) return { status: "error", message: "Jurnal tidak ditemukan." };
      if (String(log.user_id).trim() !== String(user.id).trim()) {
        return { status: "error", message: "Anda tidak memiliki akses untuk menghapus jurnal ini." };
      }
      // Guru tidak boleh hapus jurnal auto-generated libur bonus
      if (String(log.notes || "").includes("Auto-generated: Libur Bonus")) {
        return { status: "error", message: "Jurnal yang dibuat otomatis oleh sistem tidak dapat dihapus." };
      }
    }
    const sheet = getSheet("Teaching_Logs");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        try { _invalidateDataSnapshot(); } catch (_) {}
        return { status: "success" };
      }
    }
    return { status: "error", message: "Data tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: "Server Error: " + e.toString() };
  }
}

function getAdminDashboardData_V2(token) {
  try {
    const user = verifySession(token);
    if (String(user.role).toLowerCase() !== "admin")
      return { status: "error", message: "Akses Ditolak." };
    const snap = _getDataSnapshot();
    const cfg = snap.config;
    const TARIF = Number(cfg["base_salary"] || 0);
    const usersData = snap.users;
    const allSchedules = snap.schedules;
    const allLogs = snap.logs;
    const allAllowances = snap.allowances;
    const allSubstitutes = snap.substitutes;
    const teachers = usersData.filter(
      (u) => String(u.role).toLowerCase() === "guru",
    );
    const userMap = {};
    usersData.forEach((u) => (userMap[u.id] = u.full_name));
    let totalJTM = 0;
    const now = new Date();
    const m = now.getMonth();
    const y = now.getFullYear();
    const logsThisMonth = allLogs.filter((l) => {
      const d = new Date(l.date);
      return d.getMonth() === m && d.getFullYear() === y;
    });
    const jtmAdjMap = _jtmBuildKbmAdjustmentMap_();
    const piketAdjMap = _jtmBuildPiketAdjustmentMap_();
    logsThisMonth.forEach((l) => {
      const _sid = String(l.schedule_id);
      if (_sid === "PICKET-DUTY" || _sid === "PIKET") totalJTM += _jtmResolvePiketLogJtm_(piketAdjMap, l.user_id, l.date, 4);
      else if (_sid === "CEREMONY-DUTY" || _sid === "UPACARA") totalJTM += 5;
      else if (_sid === "EXAM-SUPERVISOR" || _sid === "EXAM-COMMITTEE")
        totalJTM += Number(l.jtm_val || 0);
      else if (_sid === "PARTIAL-SUB-KBM" || _sid === "PARTIAL-SUB-EXAM")
        totalJTM += Number(l.jtm_val || 0);
      else {
        const sched = allSchedules.find((s) => String(s.id) === _sid);
        if (sched)
          totalJTM += _jtmResolveKbmLogJtm_(
            jtmAdjMap,
            _sid,
            l.date,
            Number(sched.jtm_val || 0),
          );
      }
    });
    const allEvents = snap.eventAttendance;
    const eventsThisMonth = allEvents.filter((a) => {
      if (
        a.journal_submitted !== true &&
        String(a.journal_submitted).toLowerCase() !== "true"
      )
        return false;
      const d = new Date(a.date);
      return d.getMonth() === m && d.getFullYear() === y;
    });
    totalJTM += _computeEventJtm(eventsThisMonth);
    teachers.forEach((t) => {
      const bonusDetail = _calculateMonthlyBonusFast_Detailed(t.id, m, y, snap);
      totalJTM += bonusDetail.estimated;
    });
    const totalAllowancesSum = allAllowances.reduce(
      (sum, a) => sum + Number(a.amount || 0),
      0,
    );
    const totalHonor = totalJTM * TARIF + totalAllowancesSum;
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
    const periodId = `${months[m]} ${y}`;
    const transportEnabled =
      String(cfg["transport_allowance_enabled"] || "false")
        .toLowerCase()
        .trim() === "true";
    let totalTransportSchool = 0;
    if (transportEnabled) {
      teachers.forEach((u) => {
        totalTransportSchool += _getMonthlyTransportSumFast(String(u.id), periodId, snap);
      });
    }
    const dayIndex = _getIndoDayIndex();
    const todayStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    const activeTP = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    const todaysSchedules = allSchedules
      .filter((s) => {
        const isToday = String(s.day_index).trim() === String(dayIndex);
        const sTP = s.tahun_pelajaran || activeTP;
        const sSem = s.semester || activeSem;
        return isToday && sTP === activeTP && sSem === activeSem;
      })
      .map((s) => {
        const isDone = allLogs.some(
          (l) =>
            String(l.schedule_id) === String(s.id) &&
            Utilities.formatDate(
              new Date(l.date),
              "Asia/Jakarta",
              "yyyy-MM-dd",
            ) === todayStr,
        );
        const subRecord = allSubstitutes.find(
          (sub) =>
            String(sub.schedule_id) === String(s.id) &&
            Utilities.formatDate(
              new Date(sub.date),
              "Asia/Jakarta",
              "yyyy-MM-dd",
            ) === todayStr,
        );
        let displayGuruName = userMap[s.user_id] || "Unknown";
        if (subRecord) {
          const subName = userMap[subRecord.substitute_user_id] || "Unknown";
          displayGuruName = `${subName} (Gnt)`;
        }
        let tStart = safeTime(s.time_start);
        if (tStart.length > 5) tStart = tStart.substring(0, 5);
        if (!tStart) tStart = "??:??";
        
        let tEnd = safeTime(s.time_end);
        if (tEnd.length > 5) tEnd = tEnd.substring(0, 5);
        if (!tEnd) tEnd = "??:??";
        
        return {
          id: s.id,
          time: tStart + " - " + tEnd,
          guru_name: displayGuruName,
          class_name: s.class_name,
          subject: s.subject,
          status: isDone ? "✅ Sudah Mengajar" : "⏳ Belum",
        };
      })
      .sort((a, b) => a.time.localeCompare(b.time));
    const eventsTodayForAdmin = allEvents.filter((a) => {
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
    eventsTodayForAdmin.forEach((e) => {
      const isDone = String(e.journal_submitted).toLowerCase() === "true";
      todaysSchedules.push({
        id: "EVENT-" + e.id,
        time: "Acara",
        guru_name: userMap[e.user_id] || "Unknown",
        class_name: "Kegiatan/Acara",
        subject: e.event_name || "Acara Sekolah",
        status: isDone ? "✅ Sudah Mengisi" : "⏳ Belum",
      });
    });
    return {
      status: "success",
      total_guru: teachers.length,
      total_jtm_school: totalJTM,
      est_honor_school: formatRupiah(totalHonor),
      transport_total_school: totalTransportSchool,
      transport_total_school_formatted: formatRupiah(totalTransportSchool),
      transport_enabled: transportEnabled,
      schedules: todaysSchedules,
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function getAdminDailyAttendance(token) {
  try {
    const user = verifySession(token);
    if (String(user.role).toLowerCase() !== "admin") return { status: "error" };
    const dayIndex = _getIndoDayIndex();
    const todayStr = Utilities.formatDate(
      new Date(),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const snap = _getDataSnapshot();
    const schedules = snap.schedules;
    const users = snap.users;
    const attendanceData = snap.attendance;
    const substitutes = snap.substitutes;
    const cfg = snap.config;
    const activeTP = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    const todaySchedules = schedules.filter((s) => {
      const isToday = String(s.day_index).trim() === String(dayIndex);
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return isToday && sTP === activeTP && sSem === activeSem;
    });
    const todaySubstitutes = substitutes.filter(
      (s) =>
        Utilities.formatDate(new Date(s.date), "Asia/Jakarta", "yyyy-MM-dd") ===
        todayStr,
    );
    todaySchedules.sort((a, b) =>
      safeTime(a.time_start).localeCompare(safeTime(b.time_start)),
    );
    const result = todaySchedules.map((s) => {
      const sub = todaySubstitutes.find(
        (x) => String(x.schedule_id) === String(s.id),
      );
      const effectiveUserId = sub
        ? String(sub.substitute_user_id)
        : String(s.user_id);
      const teacher = users.find((u) => String(u.id) === effectiveUserId);
      let displayName = teacher ? teacher.full_name : "Unknown";
      if (sub) displayName += " (Guru Pengganti)";
      const isConfirmed = attendanceData.some(
        (a) =>
          String(a.user_id) === effectiveUserId &&
          Utilities.formatDate(
            new Date(a.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          ) === todayStr &&
          String(a.schedule_id || "") === String(s.id),
      );
      let tStart = safeTime(s.time_start);
      if (tStart.length > 5) tStart = tStart.substring(0, 5);
      let tEnd = safeTime(s.time_end);
      if (tEnd.length > 5) tEnd = tEnd.substring(0, 5);
      return {
        user_id: effectiveUserId,
        schedule_id: String(s.id),
        full_name: displayName,
        class_name: String(s.class_name || "-"),
        subject: String(s.subject || "-"),
        time: tStart + " - " + tEnd,
        is_confirmed: isConfirmed,
      };
    });
    return { status: "success", data: result };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function getAdminDailyAttendanceHistory(token, dateStr) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin")
      return { status: "error", message: "Unauthorized" };
    if (!dateStr)
      dateStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    const users = getData("Users");
    const attendanceData = getData("Daily_Attendance");
    const targetAttendances = attendanceData.filter(
      (a) => safeDate(a.date) === dateStr,
    );
    const transportEnabledStr = String(getConfigValue("transport_allowance_enabled") || "false").toLowerCase().trim();
    const transportEnabled = (transportEnabledStr === "true");
    const tarifPerKm = Number(getConfigValue("tarif_per_km") || 0);
    const result = [];
    const processedUserIds = new Set();
    targetAttendances.forEach((a) => {
      const uId = String(a.user_id);
      if (processedUserIds.has(uId)) return; 
      processedUserIds.add(uId);
      const t = users.find((u) => String(u.id) === uId);
      const fullName = t ? t.full_name : "Unknown";
      const role = t ? t.role : "Guru";
      const rawSchedIn =
        a.sched_time_in ||
        a.jam_jadwal_masuk ||
        a.jam_jadwal_masuk_guru ||
        null;
      const rawSchedOut =
        a.sched_time_out ||
        a.jam_jadwal_pulang ||
        a.jam_jadwal_pulang_guru ||
        null;
      const rawTimeIn =
        a.time_in || a.jam_masuk_aktual || a.jam_masuk_aktual_guru || null;
      const rawTimeOut =
        a.time_out || a.jam_pulang_aktual || a.jam_pulang_aktual_guru || null;
      const schedIn = safeTime(rawSchedIn);
      const schedOut = safeTime(rawSchedOut);
      const timeIn = safeTime(rawTimeIn);
      const timeOut = safeTime(rawTimeOut);
      let durationActualMin = 0;
      if (timeIn && timeOut) {
        const inParts = timeIn.split(":");
        const outParts = timeOut.split(":");
        if (inParts.length >= 2 && outParts.length >= 2) {
          const inMins = parseInt(inParts[0]) * 60 + parseInt(inParts[1]);
          const outMins = parseInt(outParts[0]) * 60 + parseInt(outParts[1]);
          durationActualMin = Math.max(0, outMins - inMins);
        }
      }
      let status = "Belum Datang";
      if (timeIn && timeOut) {
        status = "Hadir";
      } else if (timeIn) {
        status = "Sedang Bekerja";
      }
      let percentage = 0;
      let schedMins = 0;
      if (schedIn && schedOut && durationActualMin > 0) {
        const sInParts = schedIn.split(":");
        const sOutParts = schedOut.split(":");
        if (sInParts.length >= 2 && sOutParts.length >= 2) {
          const sInMins = parseInt(sInParts[0]) * 60 + parseInt(sInParts[1]);
          const sOutMins = parseInt(sOutParts[0]) * 60 + parseInt(sOutParts[1]);
          schedMins = Math.max(0, sOutMins - sInMins);
          if (schedMins > 0) {
            percentage = Math.round((durationActualMin / schedMins) * 100);
            if (percentage > 100) percentage = 100;
          }
        }
      }

      let transportNominal = 0;
      if (transportEnabled && String(role).toLowerCase() === "guru") {
        const km = t ? Number(t.km_distance || 0) : 0;
        if (timeIn && timeOut) {
           transportNominal = km * tarifPerKm * (percentage / 100);
        }
      }

      result.push({
        user_id: uId,
        full_name: fullName,
        role: role,
        sched_time_in: schedIn ? schedIn.substring(0, 5) : null,
        sched_time_out: schedOut ? schedOut.substring(0, 5) : null,
        time_in: timeIn ? timeIn.substring(0, 5) : null,
        time_out: timeOut ? timeOut.substring(0, 5) : null,
        duration_actual_min: durationActualMin,
        percentage: percentage,
        status: status,
        transport_nominal_formatted: formatRupiah(Math.round(transportNominal)),
      });
    });
    users
      .filter(
        (u) =>
          String(u.role).toLowerCase() === "guru" &&
          String(u.is_active) !== "FALSE" &&
          String(u.is_active) !== "false" &&
          u.is_active !== false,
      )
      .forEach((u) => {
        const uId = String(u.id);
        if (!processedUserIds.has(uId)) {
          result.push({
            user_id: uId,
            full_name: u.full_name,
            role: u.role,
            sched_time_in: null,
            sched_time_out: null,
            time_in: null,
            time_out: null,
            duration_actual_min: 0,
            percentage: 0,
            status: "Belum Jadwal",
            transport_nominal_formatted: "Rp 0",
          });
        }
      });
    result.sort((a, b) => a.full_name.localeCompare(b.full_name));
    return { status: "success", list: result };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function _attendanceScheduleColIndex_(sheet) {
  const lastCol = sheet.getLastColumn() || 1;
  const headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h).toLowerCase().trim().replace(/\s+/g, "_");
    });
  let idx = headers.indexOf("schedule_id");
  if (idx === -1) {
    sheet.getRange(1, lastCol + 1).setValue("schedule_id");
    idx = lastCol; 
  }
  return idx;
}

function confirmTeacherPresence(token, targetUserId, scheduleId) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const isAdmin = String(user.role).toLowerCase() === "admin";
  const isSelf = String(user.id) === String(targetUserId);
  const isPicketToday = _isPicketOfficer(user);
  if (!isAdmin && !isSelf && !isPicketToday) {
    return {
      status: "error",
      message: "Anda tidak berwenang mengonfirmasi kehadiran guru lain.",
    };
  }
  const schedId =
    scheduleId === undefined || scheduleId === null ? "" : String(scheduleId);
  const todayStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const sheet = getSheet("Daily_Attendance");
  const schedColIdx = _attendanceScheduleColIndex_(sheet);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rowSched = String(
      rows[i][schedColIdx] === undefined ? "" : rows[i][schedColIdx],
    );
    if (
      String(rows[i][2]) === String(targetUserId) &&
      Utilities.formatDate(
        new Date(rows[i][1]),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      ) === todayStr &&
      rowSched === schedId
    ) {
      return { status: "success", message: "Sudah dikonfirmasi" };
    }
  }
  const newId = generateId("ATT");
  const rowData = [newId, todayStr, targetUserId, "Present", new Date()];
  while (rowData.length < schedColIdx) rowData.push("");
  rowData[schedColIdx] = schedId;
  sheet.appendRow(rowData);
  try {
    var byName = String(
      user.full_name ||
        user.name ||
        user.username ||
        (isAdmin ? "Admin" : "Guru Piket"),
    );
    _notifKbmConfirmed_(targetUserId, { byName: byName });
  } catch (_) {}
  return { status: "success" };
}

function revokeTeacherPresence(token, targetUserId, scheduleId) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const isAdmin = String(user.role).toLowerCase() === "admin";
  const isSelf = String(user.id) === String(targetUserId);
  const isPicketToday = _isPicketOfficer(user);
  if (!isAdmin && !isSelf && !isPicketToday) {
    return {
      status: "error",
      message: "Anda tidak berwenang membatalkan kehadiran guru lain.",
    };
  }
  const schedId =
    scheduleId === undefined || scheduleId === null ? "" : String(scheduleId);
  const todayStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const sheet = getSheet("Daily_Attendance");
  const schedColIdx = _attendanceScheduleColIndex_(sheet);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rowDate = Utilities.formatDate(
      new Date(rows[i][1]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    const rowSched = String(
      rows[i][schedColIdx] === undefined ? "" : rows[i][schedColIdx],
    );
    if (
      String(rows[i][2]) === String(targetUserId) &&
      rowDate === todayStr &&
      rowSched === schedId
    ) {
      sheet.deleteRow(i + 1);
      try {
        _jtmReverseOccurrence_("KBM", schedId, rowDate, targetUserId);
      } catch (_) {}
      try {
        var byName = String(
          user.full_name ||
            user.name ||
            user.username ||
            (isAdmin ? "Admin" : "Guru Piket"),
        );
        _notifKbmRevoked_(targetUserId, { byName: byName });
      } catch (_) {}
      return { status: "success" };
    }
  }
  return {
    status: "error",
    code: JTM_ERROR_CODES.NO_ACTIVE_CONFIRMATION,
    message: "Data tidak ditemukan",
  };
}
