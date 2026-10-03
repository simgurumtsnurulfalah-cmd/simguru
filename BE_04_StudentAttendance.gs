// SiM-Guru — student attendance backend
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
