// SiM-Guru — BE_23_AttendanceTransport.gs
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
