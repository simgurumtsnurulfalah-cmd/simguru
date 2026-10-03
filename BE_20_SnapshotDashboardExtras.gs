// SiM-Guru — BE_20_SnapshotDashboardExtras.gs
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
