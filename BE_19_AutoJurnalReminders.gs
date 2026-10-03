// SiM-Guru — BE_19_AutoJurnalReminders.gs
function validateDocument(trxId) {
  if (!trxId || String(trxId).trim() === "") {
    return { status: "invalid", message: "ID Transaksi tidak boleh kosong." };
  }
  const cleanId = String(trxId).trim().toUpperCase();
  const allHistory = getData("Honor_History");
  const history = allHistory.find(
    (h) =>
      String(h.trx_id || "")
        .trim()
        .toUpperCase() === cleanId,
  );
  if (history) {
    const user = findData("Users", "id", history.user_id);
    const config = getData("Config");
    const schoolNameCfg = config.find((c) => String(c.key) === "school_name");
    const kepalaSekolahCfg = config.find(
      (c) => String(c.key) === "kepala_sekolah",
    );
    const totalJtm = Number(history.total_jtm) || 0;
    const totalHonor = Number(history.total_honor) || 0;
    const baseSalary =
      Number(
        (config.find((c) => String(c.key) === "base_salary") || {}).value,
      ) || 0;
    let periodeStr = String(history.periode || "").trim();
    if (history.periode instanceof Date) {
      periodeStr = Utilities.formatDate(
        history.periode,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
    }
    let tglSimpan = history.tanggal_simpan;
    if (tglSimpan instanceof Date) {
      tglSimpan = formatDateIndo(tglSimpan);
    } else if (tglSimpan) {
      tglSimpan = formatDateIndo(new Date(tglSimpan));
    } else {
      tglSimpan = "-";
    }
    return {
      status: "valid",
      data: {
        trx_id: String(history.trx_id),
        nama_guru: user ? String(user.full_name) : "Unknown",
        nip: user ? String(user.nip || "-") : "-",
        periode: formatPeriode(periodeStr),
        total_jtm: totalJtm,
        total_honor: formatRupiah(totalHonor),
        tgl_simpan: tglSimpan,
        tgl_verifikasi: formatDateIndo(new Date()),
        school_name: schoolNameCfg
          ? String(schoolNameCfg.value)
          : "MTs Nurul Falah",
        kepala_sekolah: kepalaSekolahCfg ? String(kepalaSekolahCfg.value) : "-",
      },
    };
  }
  return {
    status: "invalid",
    message:
      "Dokumen tidak ditemukan. Pastikan ID Transaksi yang Anda masukkan benar dan sesuai dengan slip yang diterima.",
  };
}

function getGuruJournalHistoryPaginated(
  token,
  page,
  limit,
  filterDateFrom,
  filterDateTo,
  searchQuery,
) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    page = parseInt(page) || 1;
    limit = parseInt(limit) || 10;
    // BUG FIX: tz harus dideklarasikan SEBELUM getTimestamp agar tidak undefined saat closure dibuat
    const tz = Session.getScriptTimeZone();
    const allUserLogs = getData("Teaching_Logs").filter(
      (l) => String(l.user_id).trim() === String(user.id).trim(),
    );
    const schedules = getData("Schedules");
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
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    let logsFiltered = allUserLogs;
    if (filterDateFrom) {
      logsFiltered = logsFiltered.filter((l) => {
        let dStr = l.date;
        if (l.date instanceof Date) {
          dStr = Utilities.formatDate(l.date, tz, "yyyy-MM-dd");
        } else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
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
        if (l.date instanceof Date) {
          dStr = Utilities.formatDate(l.date, tz, "yyyy-MM-dd");
        } else if (typeof dStr === "string" && dStr.indexOf("T") !== -1) {
          let dObj = new Date(dStr);
          if (!isNaN(dObj.getTime()))
            dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
        }
        return String(dStr) <= String(filterDateTo);
      });
    }
    if (searchQuery) {
      const q = String(searchQuery).toLowerCase();
      logsFiltered = logsFiltered.filter((l) => {
        let subj = "",
          cls = "";
        const sid = String(l.schedule_id).toUpperCase();
        if (sid === "PICKET-DUTY") {
          subj = "tugas piket";
        } else if (sid === "CEREMONY-DUTY") {
          subj = "pembina upacara";
        } else if (sid === "EXAM-SUPERVISOR") {
          subj = "pengawas ujian";
        } else if (sid === "EXAM-COMMITTEE") {
          subj = "panitia ujian";
        } else if (sid.indexOf("PARTIAL-SUB-") === 0) {
          subj = "substitusi parsial";
        } else {
          const s = schedules.find(
            (sc) => String(sc.id) === String(l.schedule_id),
          );
          if (s) {
            subj = String(s.subject || "").toLowerCase();
            cls = String(s.class_name || "").toLowerCase();
          }
        }
        const mat = String(l.materi || "").toLowerCase();
        return (
          subj.indexOf(q) !== -1 ||
          cls.indexOf(q) !== -1 ||
          mat.indexOf(q) !== -1
        );
      });
    }
    // BUG FIX: Stats dihitung dari logsFiltered (setelah filter aktif), bukan dari semua data
    let totalJtmAll = 0;
    let totalThisMonth = 0;
    let sumHadirAll = 0;
    let sumAbsenAll = 0;
    logsFiltered.forEach((l) => {
      const jtm = Number(l.jtm_val || 0);
      totalJtmAll += jtm;
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
    const sortedLogs = logsFiltered.sort(
      (a, b) => getTimestamp(b) - getTimestamp(a),
    );
    const totalItems = sortedLogs.length;
    const totalPages = Math.ceil(totalItems / limit) || 1;
    const startIndex = (page - 1) * limit;
    const pagedLogs = sortedLogs.slice(startIndex, startIndex + limit);
    const history = pagedLogs.map((log) => {
      let sched = null;
      const sid = String(log.schedule_id).toUpperCase();
      if (sid !== "PICKET-DUTY" && sid !== "CEREMONY-DUTY") {
        sched = schedules.find((s) => String(s.id) === String(log.schedule_id));
      }
      let subj = "N/A",
        cls = "N/A";
      if (sid === "PICKET-DUTY") {
        subj = "Tugas Piket";
        cls = "-";
      } else if (sid === "CEREMONY-DUTY") {
        subj = "Pembina Upacara";
        cls = "-";
      } else if (sid === "EXAM-SUPERVISOR") {
        subj = "Pengawas Ujian";
        cls = "-";
      } else if (sid === "EXAM-COMMITTEE") {
        subj = "Panitia Ujian";
        cls = "-";
      } else if (sid.indexOf("PARTIAL-SUB-") === 0) {
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
        class_name: cls,
        subject: subj,
        materi: log.materi,
        students: log.siswa_hadir,
        students_absent: log.siswa_absen || 0,
        notes: log.notes || "",
        clean_notes: log.notes || "-",
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

function generateAutoJurnalForHoliday(token, targetDate) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized" };
    return _runAutoJurnalForHoliday_(user.id, user.full_name, targetDate);
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function _runAutoJurnalForHoliday_(actorUserId, actorName, targetDate) {
  try {
    const checkDate = targetDate ? new Date(targetDate) : new Date();
    const dateStr = Utilities.formatDate(
      checkDate,
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const dayIndex = checkDate.getDay();
    const lockId = `AUTO_JURNAL_${dateStr}`;
    const lockSheet = getSheet("System_Locks");
    const lockData = lockSheet.getDataRange().getValues();
    const existingLock = lockData.find(
      (row) => row[0] === lockId && row[1] === "COMPLETED",
    );
    if (existingLock) {
      return {
        status: "info",
        message: "Auto-jurnal untuk hari ini sudah pernah berhasil dibuat",
        date: dateStr,
        locked: true,
        locked_at: existingLock[2],
      };
    }
    const holidays = getData("Academic_Calendar");
    const holiday = holidays.find((h) => {
      const hDate = new Date(h.date);
      const hDateStr = Utilities.formatDate(
        hDate,
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        hDateStr === dateStr && String(h.is_holiday).toLowerCase() === "true"
      );
    });
    if (!holiday) {
      return {
        status: "error",
        message: "Hari ini bukan hari libur",
        date: dateStr,
      };
    }
    const holidayDesc = holiday.description || "Hari Libur";
    const lockTimestamp = Utilities.formatDate(
      new Date(),
      "Asia/Jakarta",
      "yyyy-MM-dd HH:mm:ss",
    );
    lockSheet.appendRow([
      lockId,
      "RUNNING",
      lockTimestamp,
      actorUserId || "SYSTEM",
      "",
    ]);
    const sheet = getSheet("Teaching_Logs");
    const sheetCols = Math.max(sheet.getLastColumn(), 10);
    // Ambil konfigurasi aktif (tahun pelajaran & semester)
    const configRaw = getData("Config");
    const cfg = {};
    (configRaw || []).forEach((c) => { cfg[c.key] = c.value; });
    const activeTP = cfg["tahun_pelajaran"] || "";
    const activeSem = cfg["semester"] || "";
    // Filter jadwal: hanya hari yang sama DAN semester/tahun pelajaran yang aktif
    const allSchedules = getData("Schedules").filter((s) => {
      if (String(s.day_index).trim() !== String(dayIndex)) return false;
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return sTP === activeTP && sSem === activeSem;
    });
    const allLogsToday = getData("Teaching_Logs").filter((l) => {
      try {
        return (
          Utilities.formatDate(
            new Date(l.date),
            "Asia/Jakarta",
            "yyyy-MM-dd",
          ) === dateStr
        );
      } catch (_) {
        return false;
      }
    });
    const autoLogKeys = new Set();
    const manualLogKeys = new Set();
    allLogsToday.forEach((l) => {
      const k = String(l.user_id) + "|" + String(l.schedule_id);
      const notes = String(l.notes || "");
      if (notes.includes("Auto-generated: Libur Bonus")) {
        autoLogKeys.add(k);
      } else if (!notes.includes("Auto-generated")) {
        manualLogKeys.add(k);
      }
    });
    const generatedLogs = [];
    const skippedEntries = [];
    const newRowsToWrite = [];
    function _bufferLog(row, summary) {
      newRowsToWrite.push(row);
      autoLogKeys.add(String(row[2]) + "|" + String(row[1]));
      generatedLogs.push(summary);
    }
    allSchedules.forEach((schedule) => {
      try {
        if (
          schedule.status &&
          String(schedule.status).toLowerCase() !== "active"
        ) {
          skippedEntries.push({
            user_id: schedule.user_id,
            schedule_id: schedule.id,
            reason: "SCHEDULE_NOT_ACTIVE",
          });
          return;
        }
        const key = String(schedule.user_id) + "|" + String(schedule.id);
        if (autoLogKeys.has(key)) {
          skippedEntries.push({
            user_id: schedule.user_id,
            schedule_id: schedule.id,
            reason: "EXACT_DUPLICATE",
          });
          return;
        }
        if (manualLogKeys.has(key)) {
          skippedEntries.push({
            user_id: schedule.user_id,
            schedule_id: schedule.id,
            reason: "HAS_MANUAL_LOG",
          });
          return;
        }
        const newLogId = generateId("LOG-AUTO");
        const timeStr = Utilities.formatDate(
          new Date(),
          "Asia/Jakarta",
          "HH:mm:ss",
        );
        const materiOtomatis =
          `[LIBUR: ${holidayDesc}] - Bonus JTM Otomatis | ` +
          `Jadwal: ${schedule.subject} Kelas ${schedule.class_name} | ` +
          `JTM: ${schedule.jtm_val}`;
        _bufferLog(
          [
            newLogId,
            schedule.id,
            schedule.user_id,
            dateStr,
            materiOtomatis,
            0,
            0,
            "Auto-generated: Libur Bonus | Idempotent: " + lockId,
            Number(schedule.jtm_val) || 0,
            timeStr,
          ],
          {
            log_id: newLogId,
            user_id: schedule.user_id,
            schedule_id: schedule.id,
            subject: schedule.subject,
            class_name: schedule.class_name,
            jtm_val: Number(schedule.jtm_val) || 0,
          },
        );
      } catch (errSchedule) {
        skippedEntries.push({
          user_id: schedule && schedule.user_id,
          schedule_id: schedule && schedule.id,
          reason: "ERROR: " + errSchedule.toString(),
        });
      }
    });
    const todayPickets = getData("Picket_Schedules").filter(
      (p) => Number(p.day_index) === dayIndex,
    );
    todayPickets.forEach((picket) => {
      try {
        const picketUserId = String(picket.user_id || "").trim();
        if (!picketUserId) return;
        const key = picketUserId + "|PICKET-DUTY";
        if (autoLogKeys.has(key) || manualLogKeys.has(key)) {
          skippedEntries.push({
            user_id: picketUserId,
            schedule_id: "PICKET-DUTY",
            reason: "ALREADY_CONFIRMED",
          });
          return;
        }
        const newLogId = generateId("LOG-PCK-AUTO");
        const timeStr = Utilities.formatDate(
          new Date(),
          "Asia/Jakarta",
          "HH:mm:ss",
        );
        _bufferLog(
          [
            newLogId,
            "PICKET-DUTY",
            picketUserId,
            dateStr,
            "Melaksanakan Tugas Piket — Auto-confirmed (" + holidayDesc + ")",
            0,
            0,
            "Auto-generated: Libur Bonus | Idempotent: " + lockId,
            4,
            timeStr,
          ],
          {
            log_id: newLogId,
            user_id: picketUserId,
            schedule_id: "PICKET-DUTY",
            subject: "Tugas Piket",
            class_name: "-",
            jtm_val: 4,
          },
        );
      } catch (errPicket) {
        skippedEntries.push({
          user_id: picket && picket.user_id,
          schedule_id: "PICKET-DUTY",
          reason: "ERROR: " + errPicket.toString(),
        });
      }
    });
    if (dayIndex === 1) {
      const todayCeremonies = getData("Ceremony_Schedules").filter((c) => {
        try {
          return (
            Utilities.formatDate(
              new Date(c.date),
              "Asia/Jakarta",
              "yyyy-MM-dd",
            ) === dateStr
          );
        } catch (_) {
          return false;
        }
      });
      todayCeremonies.forEach((cer) => {
        try {
          const cerUserId = String(cer.user_id || "").trim();
          if (!cerUserId) return;
          const key = cerUserId + "|CEREMONY-DUTY";
          if (autoLogKeys.has(key) || manualLogKeys.has(key)) {
            skippedEntries.push({
              user_id: cerUserId,
              schedule_id: "CEREMONY-DUTY",
              reason: "ALREADY_CONFIRMED",
            });
            return;
          }
          const newLogId = generateId("LOG-CER-AUTO");
          const timeStr = Utilities.formatDate(
            new Date(),
            "Asia/Jakarta",
            "HH:mm:ss",
          );
          _bufferLog(
            [
              newLogId,
              "CEREMONY-DUTY",
              cerUserId,
              dateStr,
              "Melaksanakan Tugas Pembina Upacara — Auto-confirmed (" +
                holidayDesc +
                ")",
              0,
              0,
              "Auto-generated: Libur Bonus | Idempotent: " + lockId,
              5,
              timeStr,
            ],
            {
              log_id: newLogId,
              user_id: cerUserId,
              schedule_id: "CEREMONY-DUTY",
              subject: "Pembina Upacara",
              class_name: "-",
              jtm_val: 5,
            },
          );
        } catch (errCer) {
          skippedEntries.push({
            user_id: cer && cer.user_id,
            schedule_id: "CEREMONY-DUTY",
            reason: "ERROR: " + errCer.toString(),
          });
        }
      });
    }
    if (newRowsToWrite.length > 0) {
      const normalized = newRowsToWrite.map((r) => {
        const out = r.slice(0, sheetCols);
        while (out.length < sheetCols) out.push("");
        return out;
      });
      const startRow = sheet.getLastRow() + 1;
      sheet
        .getRange(startRow, 1, normalized.length, sheetCols)
        .setValues(normalized);
      SpreadsheetApp.flush();
    }
    const lockMetadata = {
      generated_by: actorName || "SYSTEM",
      generated_count: generatedLogs.length,
      total_jtm: generatedLogs.reduce((sum, l) => sum + l.jtm_val, 0),
      holiday_desc: holidayDesc,
    };
    const lockRowIndex = lockData.findIndex((row) => row[0] === lockId);
    if (lockRowIndex > 0) {
      lockSheet.getRange(lockRowIndex + 1, 2).setValue("COMPLETED");
      lockSheet.getRange(lockRowIndex + 1, 4).setValue(actorUserId || "SYSTEM");
      lockSheet
        .getRange(lockRowIndex + 1, 5)
        .setValue(JSON.stringify(lockMetadata));
    } else {
      const newLockData = lockSheet.getDataRange().getValues();
      const runningRowIndex = newLockData.findIndex(
        (row) => row[0] === lockId && row[1] === "RUNNING",
      );
      if (runningRowIndex > 0) {
        lockSheet.getRange(runningRowIndex + 1, 2).setValue("COMPLETED");
        lockSheet
          .getRange(runningRowIndex + 1, 5)
          .setValue(JSON.stringify(lockMetadata));
      }
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    var notifSummary = { sent: 0, failed: 0, recipients: 0 };
    try {
      notifSummary =
        _notifAutoJurnalLiburGenerated_(
          generatedLogs,
          dateStr,
          holidayDesc,
          actorName || "Sistem (Otomatis)",
        ) || notifSummary;
    } catch (errNotif) {
      try {
        console.warn("[autoJurnal] gagal kirim notifikasi: " + errNotif);
      } catch (_) {}
    }
    return {
      status: "success",
      message: `Berhasil membuat ${generatedLogs.length} jurnal otomatis`,
      date: dateStr,
      lock_id: lockId,
      generated: generatedLogs.length,
      skipped: skippedEntries.length,
      details: generatedLogs,
      holiday_description: holidayDesc,
      notification: notifSummary,
    };
  } catch (error) {
    try {
      const lockSheet = getSheet("System_Locks");
      const lockId = `AUTO_JURNAL_${Utilities.formatDate(targetDate ? new Date(targetDate) : new Date(), "Asia/Jakarta", "yyyy-MM-dd")}`;
      const lockData = lockSheet.getDataRange().getValues();
      const lockRowIndex = lockData.findIndex((row) => row[0] === lockId);
      if (lockRowIndex > 0) {
        lockSheet.getRange(lockRowIndex + 1, 2).setValue("FAILED");
        lockSheet.getRange(lockRowIndex + 1, 5).setValue(error.toString());
      }
    } catch (e) {}
    return { status: "error", message: error.toString() };
  }
}

function checkAutoJurnalStatus(token, targetDate) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error" };
    const checkDate = targetDate ? new Date(targetDate) : new Date();
    const dateStr = Utilities.formatDate(
      checkDate,
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const dayIndex = checkDate.getDay();
    const lockId = `AUTO_JURNAL_${dateStr}`;
    const holidays = getData("Academic_Calendar");
    const holiday = holidays.find((h) => {
      const hDate = new Date(h.date);
      const hDateStr = Utilities.formatDate(
        hDate,
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        hDateStr === dateStr && String(h.is_holiday).toLowerCase() === "true"
      );
    });
    if (!holiday) {
      return { status: "info", is_holiday: false, date: dateStr };
    }
    const lockSheet = getSheet("System_Locks");
    const lockData = lockSheet.getDataRange().getValues();
    const lockRecord = lockData.find((row) => String(row[0]) === lockId);
    let hasAutoJurnal = false;
    let generatedBy = null;
    let generatedAt = null;
    let existingCount = 0;
    const autoLogs = getData("Teaching_Logs").filter((l) => {
      const logDate = Utilities.formatDate(
        new Date(l.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        logDate === dateStr &&
        String(l.notes).includes("Auto-generated: Libur Bonus")
      );
    });
    if (autoLogs.length > 0) {
      hasAutoJurnal = true;
      existingCount = autoLogs.length;
      if (lockRecord) {
        const metadata = lockRecord[4] ? JSON.parse(String(lockRecord[4])) : {};
        generatedBy = metadata.generated_by || "Unknown";
        generatedAt = lockRecord[2]
          ? Utilities.formatDate(
              new Date(lockRecord[2]),
              "Asia/Jakarta",
              "yyyy-MM-dd HH:mm",
            )
          : "Unknown";
      }
    }
    if (String(user.role).toLowerCase() !== "admin") {
      const configRaw2 = getData("Config");
      const cfg2 = {};
      (configRaw2 || []).forEach((c) => { cfg2[c.key] = c.value; });
      const activeTP2 = cfg2["tahun_pelajaran"] || "";
      const activeSem2 = cfg2["semester"] || "";
      const mySchedules = getData("Schedules").filter((s) => {
        if (String(s.user_id) !== String(user.id)) return false;
        if (String(s.day_index) !== String(dayIndex)) return false;
        const sTP = s.tahun_pelajaran || activeTP2;
        const sSem = s.semester || activeSem2;
        return sTP === activeTP2 && sSem === activeSem2;
      });
      const myAutoLogs = autoLogs.filter(
        (l) => String(l.user_id) === String(user.id),
      );
      return {
        status: "success",
        is_holiday: true,
        date: dateStr,
        holiday_description: holiday.description,
        has_auto_jurnal: hasAutoJurnal,
        my_schedule_count: mySchedules.length,
        my_bonus_jtm: mySchedules.reduce(
          (sum, s) => sum + Number(s.jtm_val),
          0,
        ),
        my_auto_logs: myAutoLogs.map((l) => ({
          log_id: l.log_id,
          materi: l.materi,
          jtm_val: l.jtm_val,
        })),
      };
    }
    return {
      status: "success",
      is_holiday: true,
      date: dateStr,
      holiday_description: holiday.description,
      has_auto_jurnal: hasAutoJurnal,
      can_generate: !hasAutoJurnal,
      can_cancel: hasAutoJurnal,
      existing_count: existingCount,
      lock_id: lockId,
      generated_by: generatedBy,
      generated_at: generatedAt,
    };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function cancelAutoJurnalForHoliday(token, targetDate) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Unauthorized: Hanya admin yang dapat membatalkan",
      };
    }
    const checkDate = targetDate ? new Date(targetDate) : new Date();
    const dateStr = Utilities.formatDate(
      checkDate,
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const lockId = `AUTO_JURNAL_${dateStr}`;
    const allLogs = getData("Teaching_Logs");
    const autoLogs = allLogs.filter((l) => {
      const logDate = Utilities.formatDate(
        new Date(l.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        logDate === dateStr &&
        String(l.notes).includes("Auto-generated: Libur Bonus")
      );
    });
    if (autoLogs.length === 0) {
      return {
        status: "error",
        message:
          "Tidak ada jurnal otomatis yang dapat dibatalkan untuk hari ini",
        date: dateStr,
      };
    }
    const logsSheet = getSheet("Teaching_Logs");
    const logsData = logsSheet.getDataRange().getValues();
    const deletedLogs = [];
    let totalJtmRemoved = 0;
    const affectedTeacherIds = new Set();
    const users = getData("Users");
    const schedules = getData("Schedules");
    for (let i = logsData.length - 1; i >= 1; i--) {
      const row = logsData[i];
      const rowDate = Utilities.formatDate(
        new Date(row[3]),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      const notes = String(row[7] || "");
      if (
        rowDate === dateStr &&
        notes.includes("Auto-generated: Libur Bonus")
      ) {
        const userId = String(row[2]);
        const scheduleId = String(row[1]);
        const jtmVal = Number(row[8] || 0);
        const teacher = users.find((u) => String(u.id) === userId);
        const schedule = schedules.find((s) => String(s.id) === scheduleId);
        deletedLogs.push({
          log_id: row[0],
          user_id: userId,
          teacher_name: teacher ? teacher.full_name : "Unknown",
          schedule_id: scheduleId,
          subject: schedule ? schedule.subject : "Unknown",
          class_name: schedule ? schedule.class_name : "-",
          jtm_val: jtmVal,
          materi: String(row[4] || ""),
        });
        totalJtmRemoved += jtmVal;
        affectedTeacherIds.add(userId);
        logsSheet.deleteRow(i + 1);
      }
    }
    const lockSheet = getSheet("System_Locks");
    const lockData = lockSheet.getDataRange().getValues();
    for (let i = 1; i < lockData.length; i++) {
      if (String(lockData[i][0]) === lockId) {
        lockSheet.getRange(i + 1, 2).setValue("CANCELLED");
        lockSheet
          .getRange(i + 1, 5)
          .setValue(
            `Cancelled by ${user.full_name} at ${Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss")}`,
          );
        break;
      }
    }
    const auditSheet = getSheet("System_Locks");
    const cancelId = `CANCEL_${dateStr}_${new Date().getTime()}`;
    auditSheet.appendRow([
      cancelId,
      "CANCELLED",
      Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss"),
      user.id,
      JSON.stringify({
        deleted_count: deletedLogs.length,
        total_jtm_removed: totalJtmRemoved,
        affected_teachers: affectedTeacherIds.size,
        lock_id: lockId,
      }),
    ]);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: `Berhasil membatalkan ${deletedLogs.length} jurnal otomatis`,
      date: dateStr,
      deleted_count: deletedLogs.length,
      total_jtm_removed: totalJtmRemoved,
      affected_teachers: affectedTeacherIds.size,
      details: deletedLogs,
      cancelled_by: user.full_name,
      cancelled_at: Utilities.formatDate(
        new Date(),
        "Asia/Jakarta",
        "yyyy-MM-dd HH:mm:ss",
      ),
    };
  } catch (error) {
    return {
      status: "error",
      message: "Server error saat pembatalan: " + error.toString(),
    };
  }
}

function runAutoJurnalScheduler() {
  try {
    const now = new Date();
    const dateStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    const holidays = getData("Academic_Calendar");
    const holiday = holidays.find((h) => {
      const hDate = new Date(h.date);
      const hDateStr = Utilities.formatDate(
        hDate,
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return (
        hDateStr === dateStr && String(h.is_holiday).toLowerCase() === "true"
      );
    });
    if (!holiday) {
      console.log(
        "[runAutoJurnalScheduler] " + dateStr + " bukan hari libur. Skip.",
      );
      return { status: "skip", reason: "NOT_HOLIDAY", date: dateStr };
    }
    console.log(
      "[runAutoJurnalScheduler] " +
        dateStr +
        ' adalah hari libur ("' +
        holiday.description +
        '"). Menjalankan auto-jurnal...',
    );
    const result = _runAutoJurnalForHoliday_(
      "SYSTEM",
      "Sistem (Otomatis)",
      dateStr,
    );
    console.log(
      "[runAutoJurnalScheduler] Hasil: " +
        JSON.stringify({
          status: result.status,
          generated: result.generated,
          skipped: result.skipped,
          message: result.message,
          notification: result.notification || null,
        }),
    );
    return result;
  } catch (e) {
    console.error("[runAutoJurnalScheduler] Error: " + e.toString());
    return { status: "error", message: e.toString() };
  }
}

function installAutoJurnalScheduler() {
  try {
    uninstallAutoJurnalScheduler();
    const trigger = ScriptApp.newTrigger("runAutoJurnalScheduler")
      .timeBased()
      .atHour(16)
      .nearMinute(0)
      .everyDays(1)
      .inTimezone("Asia/Jakarta")
      .create();
    console.log(
      "[installAutoJurnalScheduler] Trigger terpasang. ID: " +
        trigger.getUniqueId(),
    );
    return {
      status: "success",
      message: "Trigger otomatis terpasang. Akan jalan setiap hari ~16.00 WIB.",
      trigger_id: trigger.getUniqueId(),
    };
  } catch (e) {
    console.error("[installAutoJurnalScheduler] Error: " + e.toString());
    return { status: "error", message: e.toString() };
  }
}

function uninstallAutoJurnalScheduler() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = 0;
    triggers.forEach((t) => {
      if (t.getHandlerFunction() === "runAutoJurnalScheduler") {
        ScriptApp.deleteTrigger(t);
        removed++;
      }
    });
    console.log(
      "[uninstallAutoJurnalScheduler] " + removed + " trigger dihapus.",
    );
    return { status: "success", removed: removed };
  } catch (e) {
    console.error("[uninstallAutoJurnalScheduler] Error: " + e.toString());
    return { status: "error", message: e.toString() };
  }
}

function getAutoJurnalSchedulerStatus(token) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Hanya admin yang dapat melihat status trigger.",
      };
    }
    const triggers = ScriptApp.getProjectTriggers().filter(
      (t) => t.getHandlerFunction() === "runAutoJurnalScheduler",
    );
    if (triggers.length === 0) {
      return { status: "success", installed: false, count: 0 };
    }
    return {
      status: "success",
      installed: true,
      count: triggers.length,
      trigger_ids: triggers.map((t) => t.getUniqueId()),
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function installAutoJurnalSchedulerByAdmin(token) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return {
      status: "error",
      message: "Hanya admin yang dapat memasang trigger.",
    };
  }
  return installAutoJurnalScheduler();
}

function uninstallAutoJurnalSchedulerByAdmin(token) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return {
      status: "error",
      message: "Hanya admin yang dapat menghapus trigger.",
    };
  }
  return uninstallAutoJurnalScheduler();
}

function runReminderScheduler() {
  try {
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    var hour = Number(Utilities.formatDate(now, "Asia/Jakarta", "HH"));
    console.log("[runReminderScheduler] Berjalan: " + dateStr + " jam " + hour);
    var stats = {
      missing_journal: 0,
      missing_bap: 0,
      pending_kbm: 0,
      pending_supervisor: 0,
    };
    try {
      stats.missing_journal = _reminderMissingJournal_(dateStr);
    } catch (e) {
      console.warn("reminder journal: " + e);
    }
    try {
      stats.missing_bap = _reminderMissingBap_(dateStr);
    } catch (e) {
      console.warn("reminder bap: " + e);
    }
    try {
      stats.pending_kbm = _reminderPendingKbm_(dateStr);
    } catch (e) {
      console.warn("reminder pending kbm: " + e);
    }
    try {
      stats.pending_supervisor = _reminderPendingSupervisor_(dateStr);
    } catch (e) {
      console.warn("reminder pending supervisor: " + e);
    }
    console.log(
      "[runReminderScheduler] Selesai. Stats: " + JSON.stringify(stats),
    );
    return { status: "success", date: dateStr, stats: stats };
  } catch (e) {
    console.error("[runReminderScheduler] Error: " + e);
    return { status: "error", message: e && e.message ? e.message : String(e) };
  }
}

function _reminderMissingJournal_(dateStr) {
  // isTeachingDay menentukan apakah ada jadwal KBM reguler hari ini
  var isTeachingDay = _notifIsTeachingDay_(dateStr);
  // Jurnal acara tetap diperiksa meski bukan hari KBM biasa (misal: hari ujian)
  // kecuali hari libur (hari libur berarti tidak ada kegiatan apapun)
  var isHoliday = _notifIsHoliday_(dateStr);
  if (!isTeachingDay && isHoliday) return 0;

  var d = new Date(dateStr + "T00:00:00");
  var dayIdx = d.getDay();
  var cfg = _getConfigMap();
  var activeTP = cfg["tahun_pelajaran"] || "";
  var activeSem = cfg["semester"] || "";

  // --- Ambil jadwal KBM aktif hari ini (hanya jika hari KBM reguler) ---
  var schedules = [];
  if (isTeachingDay) {
    schedules = getData("Schedules").filter(function (s) {
      if (String(s.day_index).trim() !== String(dayIdx)) return false;
      var sTP = s.tahun_pelajaran || activeTP;
      var sSem = s.semester || activeSem;
      return sTP === activeTP && sSem === activeSem;
    });
  }

  // --- Ambil substitusi hari ini ---
  var subs = isTeachingDay ? getData("Substitutes").filter(function (s) {
    try {
      var sd = Utilities.formatDate(
        new Date(s.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return sd === dateStr;
    } catch (_) {
      return false;
    }
  }) : [];

  // Kumpulkan schedule_id yang sudah punya pengganti hari ini
  // agar guru asli tidak diingatkan untuk jadwal yang sudah digantikan
  var substitutedScheduleIds = new Set();
  subs.forEach(function (sub) {
    substitutedScheduleIds.add(String(sub.schedule_id));
  });

  // --- Bangun set log yang sudah masuk hari ini ---
  var logs = (isTeachingDay && schedules.length > 0) ? getData("Teaching_Logs") : [];
  var loggedSet = {};
  logs.forEach(function (l) {
    try {
      var ld = Utilities.formatDate(
        new Date(l.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (ld !== dateStr) return;
      loggedSet[String(l.user_id) + "|" + String(l.schedule_id)] = true;
    } catch (_) {}
  });

  // --- byUser: map userId → daftar item jurnal KBM yang belum diisi ---
  var byUser = {};

  // 1) Jadwal reguler (guru asli, tidak ada pengganti)
  schedules.forEach(function (s) {
    var schedId = String(s.id);
    if (substitutedScheduleIds.has(schedId)) {
      // Jadwal ini sudah digantikan — tangani via blok guru pengganti di bawah
      return;
    }
    var originalUid = String(s.user_id);
    var key = originalUid + "|" + schedId;
    if (loggedSet[key]) return;
    if (!byUser[originalUid]) byUser[originalUid] = { kbm: [], events: [] };
    byUser[originalUid].kbm.push({
      subject: String(s.subject || "-"),
      class_name: String(s.class_name || "-"),
      time: _notifFmtTime_(s.time_start) + "–" + _notifFmtTime_(s.time_end),
    });
  });

  // 2) Guru pengganti — untuk setiap substitusi hari ini,
  //    cek apakah guru pengganti sudah mengisi jurnal
  subs.forEach(function (sub) {
    var schedId = String(sub.schedule_id);
    var substituteUid = String(sub.substitute_user_id);
    // Cari data jadwal yang bersangkutan
    var s = schedules.find(function (sc) {
      return String(sc.id) === schedId;
    });
    if (!s) return; // Jadwal tidak ada di semester aktif — abaikan
    var key = substituteUid + "|" + schedId;
    if (loggedSet[key]) return;
    if (!byUser[substituteUid]) byUser[substituteUid] = { kbm: [], events: [] };
    byUser[substituteUid].kbm.push({
      subject: String(s.subject || "-"),
      class_name: String(s.class_name || "-"),
      time: _notifFmtTime_(s.time_start) + "–" + _notifFmtTime_(s.time_end),
      is_substitute: true,
    });
  });

  // --- Jurnal Acara/Kegiatan yang belum diisi ---
  // Guru perlu mengisi jurnal acara jika:
  //   - sudah dikonfirmasi hadir (confirmed_by terisi)
  //   - belum mengisi jurnal (journal_submitted != true)
  //   - tanggal kehadiran = hari ini
  var eventAttendances = getData(EVENT_SHEET.ATTENDANCE).filter(function (a) {
    try {
      var ad = Utilities.formatDate(
        new Date(a.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (ad !== dateStr) return false;
      // Hanya yang sudah dikonfirmasi hadir
      if (!a.confirmed_by || String(a.confirmed_by).trim() === "") return false;
      // Belum isi jurnal
      var js = a.journal_submitted;
      if (js === true || String(js).toLowerCase() === "true") return false;
      return true;
    } catch (_) {
      return false;
    }
  });

  // Ambil definisi event untuk mendapat nama acara
  var eventDefs = {};
  if (eventAttendances.length > 0) {
    getData(EVENT_SHEET.DEFINITIONS).forEach(function (ev) {
      eventDefs[String(ev.id)] = ev;
    });
  }

  eventAttendances.forEach(function (a) {
    var uid = String(a.user_id || "").trim();
    if (!uid) return;
    var evDef = eventDefs[String(a.event_id)] || {};
    var evName = String(evDef.name || "Acara/Kegiatan");
    var evTime = evDef.time_start
      ? _notifFmtTime_(evDef.time_start) + "–" + _notifFmtTime_(evDef.time_end)
      : "-";
    if (!byUser[uid]) byUser[uid] = { kbm: [], events: [] };
    byUser[uid].events.push({
      event_name: evName,
      time: evTime,
    });
  });

  // --- Kirim notifikasi per guru ---
  var sent = 0;
  Object.keys(byUser).forEach(function (uid) {
    var data = byUser[uid];
    var kbmItems = data.kbm || [];
    var eventItems = data.events || [];
    if (!kbmItems.length && !eventItems.length) return;

    var paragraphs = [];

    // Bagian jurnal KBM
    if (kbmItems.length > 0) {
      paragraphs.push(
        "📚 <strong>Jurnal Mengajar (" + kbmItems.length + " jadwal):</strong>",
      );
      kbmItems.forEach(function (it) {
        var label = "• <strong>" + _escHtml_(it.subject) + "</strong>" +
          " — Kelas " + _escHtml_(it.class_name) +
          " (" + _escHtml_(it.time) + ")";
        if (it.is_substitute) {
          label += ' <span style="color:#8B5CF6;font-size:11px;">(Guru Pengganti)</span>';
        }
        paragraphs.push(label);
      });
    }

    // Bagian jurnal acara/kegiatan
    if (eventItems.length > 0) {
      paragraphs.push(
        "🎌 <strong>Jurnal Acara/Kegiatan (" + eventItems.length + " kegiatan):</strong>",
      );
      eventItems.forEach(function (it) {
        paragraphs.push(
          "• <strong>" + _escHtml_(it.event_name) + "</strong>" +
          " (" + _escHtml_(it.time) + ")",
        );
      });
    }

    paragraphs.push(
      '<span style="color:#94a3b8;font-size:12px;">Mohon segera lengkapi jurnal melalui menu ' +
      '<strong>Jurnal Mengajar</strong>' +
      (eventItems.length > 0 ? " dan <strong>Jurnal Acara</strong>" : "") +
      ".</span>",
    );

    var name = _notifGetName_(uid);
    var totalMissing = kbmItems.length + eventItems.length;
    var html = _notifBuildHtml_({
      title: "⏰ Pengingat: Jurnal Belum Diisi",
      accent: "#F59E0B",
      accent2: "#B45309",
      name: name,
      intro:
        "Kami mendeteksi <strong>" + totalMissing + " jurnal</strong> yang belum Anda lengkapi untuk hari <strong>" +
        _escHtml_(_notifFmtDateLong_(dateStr)) +
        "</strong>:",
      badges: [
        "Tanggal: " + _notifFmtDateLong_(dateStr),
        "Total Belum Diisi: " + totalMissing,
      ],
      paragraphs: paragraphs,
    });
    if (
      _notifToUser_(
        uid,
        "[SiM-Guru] Pengingat Jurnal · " + _notifFmtDate_(dateStr),
        html,
      )
    )
      sent++;
  });
  return sent;
}

function _reminderMissingBap_(dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return 0;
  var period = getData(EXAM_SHEET.PERIODS).find(function (p) {
    return dateStr >= String(p.date_start) && dateStr <= String(p.date_end);
  });
  if (!period) return 0;
  var sessions = getData(EXAM_SHEET.SESSIONS).filter(function (s) {
    return (
      String(s.period_id) === String(period.id) && String(s.date) === dateStr
    );
  });
  if (sessions.length === 0) return 0;
  var rooms = getData(EXAM_SHEET.ROOMS);
  var supervisors = getData(EXAM_SHEET.SUPERVISORS);
  var baps = getData(EXAM_SHEET.BAP);
  var nowTime = Utilities.formatDate(new Date(), "Asia/Jakarta", "HH:mm");
  var byUser = {};
  sessions.forEach(function (session) {
    var te = String(session.time_end || "");
    if (!te) return;
    var teH = Number(te.substring(0, 2));
    var teM = Number(te.substring(3, 5));
    var bapMin = teH * 60 + teM + 5;
    var nowMin =
      Number(nowTime.substring(0, 2)) * 60 + Number(nowTime.substring(3, 5));
    if (nowMin < bapMin) return;
    var roomsInSession = rooms.filter(function (r) {
      return String(r.session_id) === String(session.id);
    });
    roomsInSession.forEach(function (room) {
      var sups = supervisors.filter(function (sup) {
        return (
          String(sup.room_id) === String(room.id) &&
          String(sup.status) === "active"
        );
      });
      sups.forEach(function (sup) {
        var isConf = !!(
          sup.confirmed_by && String(sup.confirmed_by).trim() !== ""
        );
        if (!isConf) return;
        var bap = baps.find(function (b) {
          return String(b.supervisor_id) === String(sup.id);
        });
        var hasBap = !!(
          bap &&
          bap.submitted_at &&
          String(bap.submitted_at).trim() !== ""
        );
        if (hasBap) return;
        var uid = String(sup.user_id);
        if (!byUser[uid]) byUser[uid] = [];
        byUser[uid].push({
          session: String(session.session_name || ""),
          room: String(room.room_name || ""),
          subject: String(room.subject || ""),
          time:
            _notifFmtTime_(session.time_start) +
            "–" +
            _notifFmtTime_(session.time_end),
        });
      });
    });
  });
  var sent = 0;
  Object.keys(byUser).forEach(function (uid) {
    var items = byUser[uid];
    if (!items.length) return;
    var paragraphs = items.map(function (it) {
      return (
        "• <strong>" +
        _escHtml_(it.session) +
        "</strong> di Ruang " +
        _escHtml_(it.room) +
        " — " +
        _escHtml_(it.subject) +
        " (" +
        _escHtml_(it.time) +
        ")"
      );
    });
    var name = _notifGetName_(uid);
    var html = _notifBuildHtml_({
      title: "⏰ Pengingat: BAP Pengawas Ujian Belum Diisi",
      accent: "#F59E0B",
      accent2: "#B45309",
      name: name,
      intro:
        "Sesi ujian yang Anda awasi pada <strong>" +
        _escHtml_(_notifFmtDateLong_(dateStr)) +
        "</strong> sudah berakhir, namun BAP-nya belum disubmit:",
      badges: ["Tanggal: " + _notifFmtDateLong_(dateStr)],
      paragraphs: paragraphs.concat([
        '<span style="color:#94a3b8;font-size:12px;">Mohon segera lengkapi BAP melalui halaman <strong>Jadwal Ujian</strong> atau panel <strong>Panitia Ujian</strong>.</span>',
      ]),
    });
    if (
      _notifToUser_(
        uid,
        "[SiM-Guru] Pengingat BAP Pengawas · " + _notifFmtDate_(dateStr),
        html,
      )
    )
      sent++;
  });
  return sent;
}

function _reminderPendingKbm_(dateStr) {
  if (!_notifIsTeachingDay_(dateStr)) return 0;
  var d = new Date(dateStr + "T00:00:00");
  var dayIdx = d.getDay();
  var cfg = _getConfigMap();
  var activeTP = cfg["tahun_pelajaran"] || "";
  var activeSem = cfg["semester"] || "";
  var picketSchedules = getData("Picket_Schedules").filter(function (p) {
    return String(p.day_index).trim() === String(dayIdx);
  });
  if (picketSchedules.length === 0) return 0;
  var logs = getData("Teaching_Logs");
  var picketIds = picketSchedules.map(function (p) {
    return String(p.user_id);
  });
  var picketLogsToday = logs.filter(function (l) {
    try {
      var ld = Utilities.formatDate(
        new Date(l.date),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      return ld === dateStr && String(l.schedule_id) === "PICKET-DUTY";
    } catch (_) {
      return false;
    }
  });
  picketLogsToday.forEach(function (l) {
    var uid = String(l.user_id);
    if (picketIds.indexOf(uid) < 0) picketIds.push(uid);
  });
  var schedules = getData("Schedules").filter(function (s) {
    if (String(s.day_index).trim() !== String(dayIdx)) return false;
    var sTP = s.tahun_pelajaran || activeTP;
    var sSem = s.semester || activeSem;
    return sTP === activeTP && sSem === activeSem;
  });
  var subs = getData("Substitutes").filter(function (s) {
    var sd = Utilities.formatDate(
      new Date(s.date),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    return sd === dateStr;
  });
  var attendance = getData("Daily_Attendance").filter(function (a) {
    var ad = Utilities.formatDate(
      new Date(a.date),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    return ad === dateStr;
  });
  var pendingUserIds = {};
  schedules.forEach(function (s) {
    var sub = subs.find(function (x) {
      return String(x.schedule_id) === String(s.id);
    });
    var effectiveUid = sub ? String(sub.substitute_user_id) : String(s.user_id);
    var confirmed = attendance.some(function (a) {
      return (
        String(a.user_id) === effectiveUid &&
        String(a.schedule_id || "") === String(s.id)
      );
    });
    if (!confirmed) pendingUserIds[effectiveUid] = true;
  });
  var pendingCount = Object.keys(pendingUserIds).length;
  if (pendingCount === 0) return 0;
  var users = getData("Users");
  var pendingNames = Object.keys(pendingUserIds)
    .slice(0, 8)
    .map(function (uid) {
      var u = users.find(function (x) {
        return String(x.id) === uid;
      });
      return u ? String(u.full_name || u.username) : "Guru";
    });
  var sent = 0;
  picketIds.forEach(function (picketUid) {
    var name = _notifGetName_(picketUid);
    var paragraphs = [
      "<strong>" +
        pendingCount +
        "</strong> guru mata pelajaran masih menunggu konfirmasi kehadiran Anda hari ini.",
      "Daftar singkat: " +
        pendingNames
          .map(function (n) {
            return _escHtml_(n);
          })
          .join(", ") +
        (pendingCount > pendingNames.length
          ? ", dan " + (pendingCount - pendingNames.length) + " lainnya"
          : ""),
      '<span style="color:#94a3b8;font-size:12px;">Buka halaman <strong>Piket &amp; Upacara</strong> untuk segera memprosesnya.</span>',
    ];
    var html = _notifBuildHtml_({
      title: "⏰ Pengingat: Konfirmasi Kehadiran Guru",
      accent: "#3B82F6",
      accent2: "#1D4ED8",
      name: name,
      intro:
        "Sebagai Guru Piket pada <strong>" +
        _escHtml_(_notifFmtDateLong_(dateStr)) +
        "</strong>, Anda perlu mengonfirmasi kehadiran rekan-rekan guru.",
      badges: ["Tanggal: " + _notifFmtDateLong_(dateStr)],
      paragraphs: paragraphs,
    });
    if (
      _notifToUser_(
        picketUid,
        "[SiM-Guru] Pengingat Konfirmasi Kehadiran Guru · " +
          _notifFmtDate_(dateStr),
        html,
      )
    )
      sent++;
  });
  return sent;
}

function _reminderPendingSupervisor_(dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return 0;
  var period = getData(EXAM_SHEET.PERIODS).find(function (p) {
    return dateStr >= String(p.date_start) && dateStr <= String(p.date_end);
  });
  if (!period) return 0;
  var committee = getData(EXAM_SHEET.COMMITTEE).filter(function (c) {
    return (
      String(c.date) === dateStr &&
      String(c.status) === "active" &&
      c.confirmed_by &&
      String(c.confirmed_by).trim() !== ""
    );
  });
  if (committee.length === 0) return 0;
  var sessions = getData(EXAM_SHEET.SESSIONS).filter(function (s) {
    return (
      String(s.period_id) === String(period.id) && String(s.date) === dateStr
    );
  });
  if (sessions.length === 0) return 0;
  var rooms = getData(EXAM_SHEET.ROOMS);
  var supervisors = getData(EXAM_SHEET.SUPERVISORS);
  var users = getData("Users");
  var pendingNames = [];
  sessions.forEach(function (session) {
    var roomsInSession = rooms.filter(function (r) {
      return String(r.session_id) === String(session.id);
    });
    roomsInSession.forEach(function (room) {
      supervisors
        .filter(function (sup) {
          return (
            String(sup.room_id) === String(room.id) &&
            String(sup.status) === "active"
          );
        })
        .forEach(function (sup) {
          var isConf = !!(
            sup.confirmed_by && String(sup.confirmed_by).trim() !== ""
          );
          if (isConf) return;
          var u = users.find(function (x) {
            return String(x.id) === String(sup.user_id);
          });
          var name = u ? String(u.full_name || u.username) : "Pengawas";
          pendingNames.push(
            name +
              " — " +
              (session.session_name || "") +
              " / " +
              (room.room_name || ""),
          );
        });
    });
  });
  if (pendingNames.length === 0) return 0;
  var sent = 0;
  committee.forEach(function (c) {
    var picketUid = String(c.user_id);
    var name = _notifGetName_(picketUid);
    var paragraphs = [
      "<strong>" +
        pendingNames.length +
        "</strong> pengawas ruang ujian masih menunggu konfirmasi kehadiran.",
      pendingNames
        .slice(0, 8)
        .map(function (n) {
          return "• " + _escHtml_(n);
        })
        .join("<br>") +
        (pendingNames.length > 8
          ? '<br><span style="color:#94a3b8">… dan ' +
            (pendingNames.length - 8) +
            " lainnya</span>"
          : ""),
      '<span style="color:#94a3b8;font-size:12px;">Buka halaman <strong>Panitia Ujian</strong> untuk segera memprosesnya.</span>',
    ];
    var html = _notifBuildHtml_({
      title: "⏰ Pengingat: Konfirmasi Pengawas Ruang Ujian",
      accent: "#0891B2",
      accent2: "#0369A1",
      name: name,
      intro:
        "Sebagai Panitia Ujian pada <strong>" +
        _escHtml_(_notifFmtDateLong_(dateStr)) +
        "</strong>, Anda perlu mengonfirmasi kehadiran para pengawas ruang ujian.",
      badges: ["Tanggal: " + _notifFmtDateLong_(dateStr)],
      paragraphs: paragraphs,
    });
    if (
      _notifToUser_(
        picketUid,
        "[SiM-Guru] Pengingat Konfirmasi Pengawas Ujian · " +
          _notifFmtDate_(dateStr),
        html,
      )
    )
      sent++;
  });
  return sent;
}

function formatPeriode(periodeValue) {
  if (!periodeValue) return "Unknown";
  const monthNames = [
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
  const parts = String(periodeValue).split(" ");
  if (
    parts.length === 2 &&
    monthNames.includes(parts[0]) &&
    !isNaN(parseInt(parts[1]))
  ) {
    return periodeValue;
  }
  let date;
  if (periodeValue instanceof Date) {
    date = periodeValue;
  } else {
    date = new Date(periodeValue);
  }
  if (!isNaN(date.getTime())) {
    try {
      const monthIndex =
        parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), "M")) -
        1;
      const year = Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "yyyy",
      );
      if (monthIndex >= 0 && monthIndex < 12) {
        return `${monthNames[monthIndex]} ${year}`;
      }
    } catch (e) {
      const month = monthNames[date.getMonth()];
      const year = date.getFullYear();
      return `${month} ${year}`;
    }
  }
  return String(periodeValue);
}

function getPeriodeSortKey(periodeValue) {
  const monthNames = [
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
  let date;
  if (periodeValue instanceof Date) {
    date = periodeValue;
  } else {
    date = new Date(periodeValue);
  }
  if (!isNaN(date.getTime())) {
    return date.getFullYear() * 12 + date.getMonth();
  }
  const parts = String(periodeValue).split(" ");
  if (parts.length === 2) {
    const monthIndex = monthNames.indexOf(parts[0]);
    const year = parseInt(parts[1]);
    if (monthIndex !== -1 && !isNaN(year)) {
      return year * 12 + monthIndex;
    }
  }
  return 0;
}
let _cachedData = null;
let _cacheTime = 0;
const CACHE_TTL = 300000;
