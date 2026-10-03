// SiM-Guru — BE_16_UsersAdmin.gs
function getAllUsers(token) {
  const user = verifySession(token);
  if (!user) return [];
  return getData("Users").map((u) => ({
    id: u.id,
    full_name: u.full_name,
    username: u.username,
    role: u.role,
    nip: u.nip,
    km_distance:
      u.km_distance !== undefined && u.km_distance !== null
        ? u.km_distance
        : "",
  }));
}

function _validateKmDistance_(raw) {
  if (raw === null || raw === undefined || raw === "") {
    return { ok: true, value: 0 };
  }
  var s = String(raw).trim();
  if (s === "") return { ok: true, value: 0 };
  var num = Number(s);
  if (isNaN(num) || !isFinite(num)) {
    return {
      ok: false,
      message:
        "Jarak harus berupa angka antara 0,01 dan 999,99 dengan maksimal 2 angka di belakang koma.",
    };
  }
  if (num === 0) return { ok: true, value: 0 };
  if (num < 0.01 || num > 999.99) {
    return {
      ok: false,
      message:
        "Jarak harus berupa angka antara 0,01 dan 999,99 dengan maksimal 2 angka di belakang koma.",
    };
  }
  var decimalMatch = s.match(/\.(\d+)$/);
  if (decimalMatch && decimalMatch[1].length > 2) {
    return {
      ok: false,
      message:
        "Jarak harus berupa angka antara 0,01 dan 999,99 dengan maksimal 2 angka di belakang koma.",
    };
  }
  return { ok: true, value: num };
}

function _validateAdditionalRole_(raw) {
  if (raw === null || raw === undefined || raw === "") {
    return { ok: true, value: "" };
  }
  var s = String(raw).trim().toUpperCase();
  if (s === "") return { ok: true, value: "" };
  if (s === "KEPALA_SEKOLAH") return { ok: true, value: "KEPALA_SEKOLAH" };
  return {
    ok: false,
    message:
      "Nilai additional_role tidak valid. Gunakan KEPALA_SEKOLAH atau kosongkan.",
  };
}

function saveUser(token, data) {
  const admin = verifySession(token);
  if (!admin || admin.role !== "admin")
    return { status: "error", message: "Akses ditolak" };
  var kmValidation = _validateKmDistance_(
    Object.prototype.hasOwnProperty.call(data, "km_distance")
      ? data.km_distance
      : null,
  );
  if (!kmValidation.ok)
    return { status: "error", message: kmValidation.message };
  var roleValidation = _validateAdditionalRole_(
    Object.prototype.hasOwnProperty.call(data, "additional_role")
      ? data.additional_role
      : null,
  );
  if (!roleValidation.ok)
    return { status: "error", message: roleValidation.message };
  const sheet = getSheet("Users");
  if (data.id) {
    const rows = sheet.getDataRange().getValues();
    for (let j = 1; j < rows.length; j++) {
      if (
        String(rows[j][0]) !== String(data.id) &&
        String(rows[j][1]).toLowerCase() === String(data.username).toLowerCase()
      ) {
        return {
          status: "error",
          message: "Username sudah digunakan oleh pengguna lain",
        };
      }
    }
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        const oldUsername = String(rows[i][1]);
        sheet.getRange(i + 1, 2).setValue(data.username);
        sheet.getRange(i + 1, 4).setValue(data.role);
        sheet.getRange(i + 1, 5).setValue(data.full_name);
        if (Object.prototype.hasOwnProperty.call(data, "nip")) {
          sheet
            .getRange(i + 1, 6)
            .setValue(String(data.nip || "").trim() || "-");
        }
        if (oldUsername !== String(data.username)) {
          _extUpdateUsername_(oldUsername, data.username);
        }
        if (Object.prototype.hasOwnProperty.call(data, "km_distance")) {
          const kmCol = _ensureUserColumn_(sheet, "km_distance", 13);
          sheet
            .getRange(i + 1, kmCol)
            .setValue(kmValidation.value === 0 ? "" : kmValidation.value);
        }
        if (Object.prototype.hasOwnProperty.call(data, "additional_role")) {
          const arCol = _ensureUserColumn_(sheet, "additional_role", 14);
          sheet.getRange(i + 1, arCol).setValue(roleValidation.value);
        }
        try {
          _invalidateDataSnapshot();
        } catch (_) {}
        return { status: "success" };
      }
    }
    return { status: "error", message: "User tidak ditemukan" };
  } else {
    const rows = sheet.getDataRange().getValues();
    for (let j = 1; j < rows.length; j++) {
      if (
        String(rows[j][1]).toLowerCase() === String(data.username).toLowerCase()
      ) {
        return { status: "error", message: "Username sudah digunakan" };
      }
    }
    const newId = generateId("USR");
    const defaultPass = data.password || "123456";
    const kmValue = Object.prototype.hasOwnProperty.call(data, "km_distance")
      ? kmValidation.value === 0
        ? ""
        : kmValidation.value
      : "";
    const arValue = Object.prototype.hasOwnProperty.call(
      data,
      "additional_role",
    )
      ? roleValidation.value
      : "";
    const nipValue = String(data.nip || "").trim() || "-";
    sheet.appendRow([
      newId,
      data.username,
      defaultPass,
      data.role,
      data.full_name,
      nipValue,
      0,
      "",
      "",
      "",
      "",
      "",
      kmValue,
      arValue,
    ]);
    const extOk = _extAddUser_(data.username, defaultPass);
    if (!extOk) {
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][0]) === newId) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return {
        status: "error",
        message: "Gagal menambah user ke sumber credential. Coba lagi.",
      };
    }
    return { status: "success" };
  }
}

function deleteUser(token, id) {
  const admin = verifySession(token);
  if (!admin || admin.role !== "admin")
    return { status: "error", message: "Akses ditolak" };
  const users = getData("Users");
  const target = users.find((u) => String(u.id) === String(id));
  const result = deleteRowById(token, "Users", id);
  if (target && target.username) {
    _extDeleteUser_(target.username);
  }
  return result;
}

function resetUserPassword(token, id) {
  const admin = verifySession(token);
  if (!admin || admin.role !== "admin")
    return { status: "error", message: "Akses ditolak" };
  const sheet = getSheet("Users");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      const username = String(rows[i][1]);
      const extOk = _extSetPassword_(username, "123456");
      if (!extOk) {
        return {
          status: "error",
          message:
            "Gagal reset password di sumber credential. Hubungi Administrator.",
        };
      }
      sheet.getRange(i + 1, 3).setValue("123456");
      return {
        status: "success",
        message: "Password berhasil direset ke 123456",
      };
    }
  }
  return { status: "error", message: "User tidak ditemukan" };
}

function getConfigValue(key) {
  const conf = getData("Config");
  const item = conf.find((c) => c.key === key);
  return item ? item.value : "";
}

function savePdfToDrive(base64, filename) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      "application/pdf",
      filename,
    );
    const folder = DriveApp.getRootFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function getAdminJournalHistory(
  token,
  page,
  limit,
  filterDateFrom,
  filterDateTo,
  filterGuru,
  filterClass,
) {
  try {
    const user = verifySession(token);
    if (!user || user.role !== "admin")
      return { status: "error", message: "Unauthorized" };
    page = parseInt(page) || 1;
    limit = parseInt(limit) || 10;
    const logsRaw = getData("Teaching_Logs");
    const users = getData("Users");
    const schedules = getData("Schedules");
    const tz = Session.getScriptTimeZone();
    const getTimestamp = (entry) => {
      let dStr = entry.date;
      if (dStr instanceof Date) {
        dStr = Utilities.formatDate(dStr, tz, "yyyy-MM-dd");
      } else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
        let dObj = new Date(dStr);
        if (!isNaN(dObj.getTime()))
          dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
      }
      let tStr = entry.waktu_submit;
      if (tStr instanceof Date) {
        tStr = Utilities.formatDate(tStr, tz, "HH:mm:ss");
      } else if (typeof tStr === "string") {
        if (tStr.indexOf("T") !== -1) {
          let tObj = new Date(tStr);
          if (!isNaN(tObj.getTime()))
            tStr = Utilities.formatDate(tObj, tz, "HH:mm:ss");
        } else if (tStr.length === 5) {
          tStr = tStr + ":00";
        }
      }
      let finalDateStr = dStr + "T" + (tStr || "00:00:00");
      let ts = new Date(finalDateStr).getTime();
      return isNaN(ts) ? 0 : ts;
    };
    const sortedAll = logsRaw.sort((a, b) => getTimestamp(b) - getTimestamp(a));
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    let logsFiltered = sortedAll;
    if (filterDateFrom) {
      logsFiltered = logsFiltered.filter((l) => {
        let dStr = l.date;
        if (l.date instanceof Date)
          dStr = Utilities.formatDate(l.date, tz, "yyyy-MM-dd");
        else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
          let dObj = new Date(dStr);
          if (!isNaN(dObj.getTime()))
            dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
        }
        return String(dStr) >= String(filterDateFrom);
      });
    }
    if (filterDateTo) {
      logsFiltered = logsFiltered.filter((l) => {
        let dStr = l.date;
        if (l.date instanceof Date)
          dStr = Utilities.formatDate(l.date, tz, "yyyy-MM-dd");
        else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
          let dObj = new Date(dStr);
          if (!isNaN(dObj.getTime()))
            dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
        }
        return String(dStr) <= String(filterDateTo);
      });
    }
    if (filterGuru) {
      const fGuruLower = String(filterGuru).toLowerCase();
      logsFiltered = logsFiltered.filter((l) => {
        const teacher = users.find((u) => String(u.id) === String(l.user_id));
        return (
          teacher &&
          (teacher.full_name || "").toLowerCase().includes(fGuruLower)
        );
      });
    }
    if (filterClass) {
      const fClassLower = String(filterClass).toLowerCase();
      logsFiltered = logsFiltered.filter((l) => {
        const sid = String(l.schedule_id);
        // Log special-duty tidak memiliki kelas — jangan ikut difilter (dibuang) saat filter kelas aktif
        if (
          sid === "PICKET-DUTY" ||
          sid === "CEREMONY-DUTY" ||
          sid === "EXAM-SUPERVISOR" ||
          sid === "EXAM-COMMITTEE" ||
          sid.indexOf("PARTIAL-SUB-") === 0
        ) {
          return false;
        }
        const sched = schedules.find(
          (s) => String(s.id) === sid,
        );
        return (
          sched && (sched.class_name || "").toLowerCase().includes(fClassLower)
        );
      });
    }
    // BUG FIX: Stats dihitung dari logsFiltered (setelah filter aktif), bukan dari semua data
    let totalJtmAll = 0;
    let totalThisMonth = 0;
    let sumHadirAll = 0;
    let sumAbsenAll = 0;
    logsFiltered.forEach((l) => {
      totalJtmAll += Number(l.jtm_val || 0);
      sumHadirAll += Number(l.siswa_hadir || 0);
      sumAbsenAll += Number(l.siswa_absen || 0);
      let dStr = l.date;
      if (l.date instanceof Date) {
        dStr = Utilities.formatDate(l.date, tz, "yyyy-MM-dd");
      } else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
        let dObj = new Date(dStr);
        if (!isNaN(dObj.getTime()))
          dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
      }
      const d = new Date(dStr);
      if (d.getMonth() === currentMonth && d.getFullYear() === currentYear) {
        totalThisMonth++;
      }
    });
    const totalAllSiswa = sumHadirAll + sumAbsenAll;
    const attRateAll =
      totalAllSiswa > 0
        ? Math.round((sumHadirAll * 100) / totalAllSiswa)
        : null;
    const totalItems = logsFiltered.length;
    const totalPages = Math.ceil(totalItems / limit) || 1;
    const startIndex = (page - 1) * limit;
    const pagedLogs = logsFiltered.slice(startIndex, startIndex + limit);
    const history = pagedLogs.map((log) => {
      const teacher = users.find((u) => String(u.id) === String(log.user_id));
      let sched = null;
      if (
        log.schedule_id !== "PICKET-DUTY" &&
        log.schedule_id !== "CEREMONY-DUTY"
      ) {
        sched = schedules.find((s) => String(s.id) === String(log.schedule_id));
      }
      let subj = "N/A",
        cls = "N/A";
      if (log.schedule_id === "PICKET-DUTY") {
        subj = "Tugas Piket";
        cls = "-";
      } else if (log.schedule_id === "CEREMONY-DUTY") {
        subj = "Pembina Upacara";
        cls = "-";
      } else if (log.schedule_id === "EXAM-SUPERVISOR") {
        subj = "Pengawas Ujian";
        cls = "-";
      } else if (log.schedule_id === "EXAM-COMMITTEE") {
        subj = "Panitia Ujian";
        cls = "-";
      } else if (String(log.schedule_id).indexOf("PARTIAL-SUB-") === 0) {
        subj = "Substitusi Parsial";
        cls = "-";
      } else if (sched) {
        subj = sched.subject;
        cls = sched.class_name;
      }
      let safeDate = log.date;
      if (safeDate instanceof Date) {
        safeDate = Utilities.formatDate(safeDate, tz, "yyyy-MM-dd");
      } else if (typeof safeDate === "string" && safeDate.indexOf("T") !== -1) {
        let dObj = new Date(safeDate);
        if (!isNaN(dObj.getTime()))
          safeDate = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
      }
      let safeTime = log.waktu_submit;
      if (safeTime instanceof Date) {
        safeTime = Utilities.formatDate(safeTime, tz, "HH:mm");
      } else if (typeof safeTime === "string") {
        if (safeTime.indexOf("T") !== -1) {
          let tObj = new Date(safeTime);
          if (!isNaN(tObj.getTime()))
            safeTime = Utilities.formatDate(tObj, tz, "HH:mm");
        } else if (safeTime.length >= 5) {
          safeTime = safeTime.substring(0, 5);
        }
      }
      if (!safeTime) safeTime = "-";
      return {
        log_id: log.log_id,
        date: String(safeDate),
        time: String(safeTime),
        guru_name: teacher ? teacher.full_name : "Unknown",
        guru_id: log.user_id,
        class_name: cls,
        subject: subj,
        materi: log.materi,
        students_present: log.siswa_hadir,
        students_absent: log.siswa_absen,
        notes: log.notes || "",
        jtm_val: Number(log.jtm_val || 0),
      };
    });
    return {
      status: "success",
      data: history,
      stats: {
        total_this_month: totalThisMonth,
        total_jtm_all: totalJtmAll,
        att_rate_all: attRateAll,
      },
      pagination: {
        current_page: page,
        total_pages: totalPages,
        total_items: totalItems,
        items_per_page: limit,
      },
    };
  } catch (error) {
    return { status: "error", message: "Server Error: " + error.toString() };
  }
}
