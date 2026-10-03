// SiM-Guru — BE_21_Exam.gs
function _examNow_() {
  return Utilities.formatDate(
    new Date(),
    "Asia/Jakarta",
    "yyyy-MM-dd'T'HH:mm:ss",
  );
}

function _examToday_() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
}

function _examGetSession_(session_id) {
  return (
    getData(EXAM_SHEET.SESSIONS).find(
      (s) => String(s.id) === String(session_id),
    ) || null
  );
}

function _examGetRoom_(room_id) {
  return (
    getData(EXAM_SHEET.ROOMS).find((r) => String(r.id) === String(room_id)) ||
    null
  );
}

function _examGetRoomsBySession_(session_id) {
  return getData(EXAM_SHEET.ROOMS).filter(
    (r) => String(r.session_id) === String(session_id),
  );
}

function _examGetSessionsByPeriod_(period_id) {
  return getData(EXAM_SHEET.SESSIONS).filter(
    (s) => String(s.period_id) === String(period_id),
  );
}

function _examDeleteBySet_(sheetName, colIndex, valueSet) {
  if (!valueSet || valueSet.size === 0) return;
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (valueSet.has(String(data[i][colIndex]))) {
      sheet.deleteRow(i + 1);
    }
  }
}

function _examDeleteById_(sheetName, id) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function _examFormatTime_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Jakarta", "HH:mm");
  }
  return String(val || "").substring(0, 5);
}

function _examCheckSupervisorConflict_(user_id, session_id, exclude_sup_id) {
  const session = _examGetSession_(session_id);
  if (!session) return "Sesi ujian tidak ditemukan.";
  const targetDate = String(session.date);
  const sessionRoomIds = new Set(
    _examGetRoomsBySession_(session_id).map((r) => String(r.id)),
  );
  const conflictSameSesi = getData(EXAM_SHEET.SUPERVISORS).some(
    (sup) =>
      String(sup.user_id) === String(user_id) &&
      String(sup.status) === "active" &&
      sessionRoomIds.has(String(sup.room_id)) &&
      (!exclude_sup_id || String(sup.id) !== String(exclude_sup_id)),
  );
  if (conflictSameSesi) {
    return "Guru ini sudah menjadi pengawas di ruang lain pada sesi yang sama.";
  }
  return null;
}

function _examCheckCommitteeConflict_(user_id, date, exclude_com_id) {
  const targetDate = String(date);
  const alreadyCommittee = getData(EXAM_SHEET.COMMITTEE).some(
    (c) =>
      String(c.user_id) === String(user_id) &&
      String(c.date) === targetDate &&
      String(c.status) === "active" &&
      (!exclude_com_id || String(c.id) !== String(exclude_com_id)),
  );
  if (alreadyCommittee) {
    return "Guru ini sudah terdaftar sebagai Panitia pada tanggal tersebut.";
  }
  return null;
}

function getExamPeriods(token) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    const today = _examToday_();
    const data = getData(EXAM_SHEET.PERIODS)
      .map((p) => {
        const ds = String(p.date_start);
        const de = String(p.date_end);
        let periodStatus;
        if (today < ds) periodStatus = "Akan Datang";
        else if (today > de) periodStatus = "Selesai";
        else periodStatus = "Berlangsung";
        return {
          id: String(p.id),
          name: String(p.name || ""),
          date_start: ds,
          date_end: de,
          jtm_committee_per_day: Number(p.jtm_committee_per_day) || 0,
          description: String(p.description || ""),
          created_by: String(p.created_by || ""),
          created_at: String(p.created_at || ""),
          status: periodStatus,
        };
      })
      .sort((a, b) => b.date_start.localeCompare(a.date_start));
    return { status: "success", data };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function addExamPeriod(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat membuat periode ujian.",
      };
    }
    if (!data.name || !String(data.name).trim()) {
      return { status: "error", message: "Nama periode ujian wajib diisi." };
    }
    if (!data.date_start) {
      return { status: "error", message: "Tanggal mulai wajib diisi." };
    }
    if (!data.date_end) {
      return { status: "error", message: "Tanggal selesai wajib diisi." };
    }
    if (String(data.date_start) > String(data.date_end)) {
      return {
        status: "error",
        message: "Tanggal mulai tidak boleh setelah tanggal selesai.",
      };
    }
    const jtm = Number(data.jtm_committee_per_day);
    if (isNaN(jtm) || jtm < 0) {
      return {
        status: "error",
        message: "Nilai JTM Panitia per hari harus berupa angka (minimal 0).",
      };
    }
    const newId = "EXP-" + new Date().getTime();
    getSheet(EXAM_SHEET.PERIODS).appendRow([
      newId,
      String(data.name).trim(),
      String(data.date_start),
      String(data.date_end),
      jtm,
      String(data.description || "").trim(),
      String(user.id),
      _examNow_(),
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Periode ujian berhasil dibuat.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function editExamPeriod(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat mengedit periode ujian.",
      };
    }
    if (!data.id)
      return { status: "error", message: "ID periode tidak ditemukan." };
    if (!data.name || !String(data.name).trim()) {
      return { status: "error", message: "Nama periode wajib diisi." };
    }
    if (!data.date_start || !data.date_end) {
      return {
        status: "error",
        message: "Tanggal mulai dan selesai wajib diisi.",
      };
    }
    if (String(data.date_start) > String(data.date_end)) {
      return {
        status: "error",
        message: "Tanggal mulai tidak boleh setelah tanggal selesai.",
      };
    }
    const jtm = Number(data.jtm_committee_per_day);
    if (isNaN(jtm) || jtm < 0) {
      return {
        status: "error",
        message: "Nilai JTM Panitia per hari harus berupa angka (minimal 0).",
      };
    }
    const existing = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(data.id),
    );
    if (!existing)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    if (_examToday_() >= String(existing.date_start)) {
      return {
        status: "error",
        message: "Periode yang sudah dimulai atau selesai tidak dapat diedit.",
      };
    }
    const sheet = getSheet(EXAM_SHEET.PERIODS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 2).setValue(String(data.name).trim());
        sheet.getRange(i + 1, 3).setValue(String(data.date_start));
        sheet.getRange(i + 1, 4).setValue(String(data.date_end));
        sheet.getRange(i + 1, 5).setValue(jtm);
        sheet
          .getRange(i + 1, 6)
          .setValue(String(data.description || "").trim());
        return {
          status: "success",
          message: "Periode ujian berhasil diperbarui.",
        };
      }
    }
    return { status: "error", message: "Periode ujian tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function deleteExamPeriod(token, id, force) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat menghapus periode ujian.",
      };
    }
    if (!id) return { status: "error", message: "ID periode tidak ditemukan." };
    const existing = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(id),
    );
    if (!existing)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const sessions = _examGetSessionsByPeriod_(id);
    const sessionIds = new Set(sessions.map((s) => String(s.id)));
    const allRooms = getData(EXAM_SHEET.ROOMS);
    const rooms = allRooms.filter((r) => sessionIds.has(String(r.session_id)));
    const roomIds = new Set(rooms.map((r) => String(r.id)));
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const supervisors = allSups.filter((sup) =>
      roomIds.has(String(sup.room_id)),
    );
    const supIds = new Set(supervisors.map((s) => String(s.id)));
    const allCom = getData(EXAM_SHEET.COMMITTEE);
    const committee = allCom.filter((c) => String(c.period_id) === String(id));
    if (!force) {
      const hasConfirmedSup = supervisors.some(
        (sup) => sup.confirmed_by && String(sup.confirmed_by).trim() !== "",
      );
      const hasConfirmedCom = committee.some(
        (c) => c.confirmed_by && String(c.confirmed_by).trim() !== "",
      );
      if (hasConfirmedSup || hasConfirmedCom) {
        return {
          status: "confirm_required",
          message:
            "Periode ini sudah memiliki data konfirmasi dan JTM terhitung. " +
            "Penghapusan akan menghapus semua data terkait secara permanen. Lanjutkan?",
        };
      }
    } else {
      supIds.forEach(function (supId) {
        _examDeleteLogByRef_(supId);
      });
      committee.forEach(function (c) {
        _examDeleteLogByRef_(String(c.id));
      });
    }
    _examDeleteBySet_(EXAM_SHEET.BAP, 1, supIds);
    _examDeleteBySet_(EXAM_SHEET.SUPERVISORS, 1, roomIds);
    _examDeleteBySet_(EXAM_SHEET.ROOMS, 1, sessionIds);
    const sesSheet = getSheet(EXAM_SHEET.SESSIONS);
    const sesData = sesSheet.getDataRange().getValues();
    for (let i = sesData.length - 1; i >= 1; i--) {
      if (String(sesData[i][1]) === String(id)) sesSheet.deleteRow(i + 1);
    }
    const comSheet = getSheet(EXAM_SHEET.COMMITTEE);
    const comData = comSheet.getDataRange().getValues();
    for (let i = comData.length - 1; i >= 1; i--) {
      if (String(comData[i][1]) === String(id)) comSheet.deleteRow(i + 1);
    }
    _examDeleteById_(EXAM_SHEET.PERIODS, id);
    return {
      status: "success",
      message: "Periode ujian dan semua data terkait berhasil dihapus.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function addExamSession(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menambah sesi ujian.",
      };
    }
    if (!data.period_id)
      return { status: "error", message: "ID periode wajib diisi." };
    if (!data.date)
      return { status: "error", message: "Tanggal sesi wajib diisi." };
    if (!data.session_name || !String(data.session_name).trim()) {
      return { status: "error", message: "Nama sesi wajib diisi." };
    }
    if (!data.time_start)
      return { status: "error", message: "Jam mulai sesi wajib diisi." };
    if (!data.time_end)
      return { status: "error", message: "Jam selesai sesi wajib diisi." };
    const ts = String(data.time_start).substring(0, 5);
    const te = String(data.time_end).substring(0, 5);
    if (ts >= te) {
      return {
        status: "error",
        message: "Jam mulai harus sebelum jam selesai.",
      };
    }
    const jtm = Number(data.jtm_val);
    if (isNaN(jtm) || jtm < 0) {
      return {
        status: "error",
        message: "Nilai JTM per sesi harus berupa angka (minimal 0).",
      };
    }
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(data.period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const sesDate = String(data.date);
    if (
      sesDate < String(period.date_start) ||
      sesDate > String(period.date_end)
    ) {
      return {
        status: "error",
        message:
          "Tanggal sesi harus berada dalam rentang periode ujian (" +
          period.date_start +
          " s.d. " +
          period.date_end +
          ").",
      };
    }
    const newId = "EXS-" + new Date().getTime();
    getSheet(EXAM_SHEET.SESSIONS).appendRow([
      newId,
      String(data.period_id),
      sesDate,
      String(data.session_name).trim(),
      ts,
      te,
      jtm,
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Sesi ujian berhasil ditambahkan.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function editExamSession(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat mengedit sesi ujian.",
      };
    }
    if (!data.id)
      return { status: "error", message: "ID sesi tidak ditemukan." };
    if (!data.session_name || !String(data.session_name).trim()) {
      return { status: "error", message: "Nama sesi wajib diisi." };
    }
    if (!data.time_start || !data.time_end) {
      return {
        status: "error",
        message: "Jam mulai dan jam selesai wajib diisi.",
      };
    }
    const ts = String(data.time_start).substring(0, 5);
    const te = String(data.time_end).substring(0, 5);
    if (ts >= te) {
      return {
        status: "error",
        message: "Jam mulai harus sebelum jam selesai.",
      };
    }
    const jtm = Number(data.jtm_val);
    if (isNaN(jtm) || jtm < 0) {
      return {
        status: "error",
        message: "Nilai JTM per sesi harus berupa angka (minimal 0).",
      };
    }
    const roomIds = new Set(
      _examGetRoomsBySession_(data.id).map((r) => String(r.id)),
    );
    const hasConfirmed = getData(EXAM_SHEET.SUPERVISORS).some(
      (sup) =>
        roomIds.has(String(sup.room_id)) &&
        sup.confirmed_by &&
        String(sup.confirmed_by).trim() !== "",
    );
    if (hasConfirmed) {
      return {
        status: "error",
        message:
          "Sesi ini sudah memiliki pengawas yang dikonfirmasi kehadirannya dan tidak dapat diedit.",
      };
    }
    const sheet = getSheet(EXAM_SHEET.SESSIONS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 4).setValue(String(data.session_name).trim());
        sheet.getRange(i + 1, 5).setValue(ts);
        sheet.getRange(i + 1, 6).setValue(te);
        sheet.getRange(i + 1, 7).setValue(jtm);
        return {
          status: "success",
          message: "Sesi ujian berhasil diperbarui.",
        };
      }
    }
    return { status: "error", message: "Sesi ujian tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function deleteExamSession(token, id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menghapus sesi ujian.",
      };
    }
    if (!id) return { status: "error", message: "ID sesi tidak ditemukan." };
    const rooms = _examGetRoomsBySession_(id);
    const roomIds = new Set(rooms.map((r) => String(r.id)));
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const supervisors = allSups.filter((sup) =>
      roomIds.has(String(sup.room_id)),
    );
    const supIds = new Set(supervisors.map((s) => String(s.id)));
    const hasConfirmed = supervisors.some(
      (sup) => sup.confirmed_by && String(sup.confirmed_by).trim() !== "",
    );
    if (hasConfirmed) {
      return {
        status: "error",
        message:
          "Sesi ini sudah memiliki pengawas yang dikonfirmasi kehadirannya dan tidak dapat dihapus.",
      };
    }
    _examDeleteBySet_(EXAM_SHEET.BAP, 1, supIds);
    _examDeleteBySet_(EXAM_SHEET.SUPERVISORS, 1, roomIds);
    _examDeleteBySet_(EXAM_SHEET.ROOMS, 1, new Set([String(id)]));
    _examDeleteById_(EXAM_SHEET.SESSIONS, id);
    return { status: "success", message: "Sesi ujian berhasil dihapus." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function addExamRoom(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menambah ruang ujian.",
      };
    }
    if (!data.session_id)
      return { status: "error", message: "ID sesi wajib diisi." };
    if (!data.room_name || !String(data.room_name).trim())
      return { status: "error", message: "Nama/nomor ruang wajib diisi." };
    if (!data.subject || !String(data.subject).trim())
      return { status: "error", message: "Mata ujian wajib diisi." };
    if (!data.class_name || !String(data.class_name).trim())
      return { status: "error", message: "Nama kelas wajib diisi." };
    const totalSiswa = Number(data.total_siswa) || 0;
    if (totalSiswa < 0) {
      return { status: "error", message: "Total siswa tidak boleh negatif." };
    }
    if (!_examGetSession_(data.session_id)) {
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    }
    const newId = "EXR-" + new Date().getTime();
    getSheet(EXAM_SHEET.ROOMS).appendRow([
      newId,
      String(data.session_id),
      String(data.room_name).trim(),
      String(data.subject).trim(),
      String(data.class_name).trim(),
      totalSiswa,
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Ruang ujian berhasil ditambahkan.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function editExamRoom(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat mengedit ruang ujian.",
      };
    }
    if (!data.id)
      return { status: "error", message: "ID ruang tidak ditemukan." };
    if (!data.room_name || !String(data.room_name).trim())
      return { status: "error", message: "Nama/nomor ruang wajib diisi." };
    if (!data.subject || !String(data.subject).trim())
      return { status: "error", message: "Mata ujian wajib diisi." };
    if (!data.class_name || !String(data.class_name).trim())
      return { status: "error", message: "Nama kelas wajib diisi." };
    const totalSiswa = Number(data.total_siswa) || 0;
    if (totalSiswa < 0) {
      return { status: "error", message: "Total siswa tidak boleh negatif." };
    }
    const hasConfirmed = getData(EXAM_SHEET.SUPERVISORS).some(
      (sup) =>
        String(sup.room_id) === String(data.id) &&
        sup.confirmed_by &&
        String(sup.confirmed_by).trim() !== "",
    );
    if (hasConfirmed) {
      return {
        status: "error",
        message:
          "Ruang ini sudah memiliki pengawas yang dikonfirmasi dan tidak dapat diedit.",
      };
    }
    const sheet = getSheet(EXAM_SHEET.ROOMS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 3).setValue(String(data.room_name).trim());
        sheet.getRange(i + 1, 4).setValue(String(data.subject).trim());
        sheet.getRange(i + 1, 5).setValue(String(data.class_name).trim());
        sheet.getRange(i + 1, 6).setValue(totalSiswa);
        return {
          status: "success",
          message: "Ruang ujian berhasil diperbarui.",
        };
      }
    }
    return { status: "error", message: "Ruang ujian tidak ditemukan." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function deleteExamRoom(token, id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menghapus ruang ujian.",
      };
    }
    if (!id) return { status: "error", message: "ID ruang tidak ditemukan." };
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const supsInRoom = allSups.filter(
      (sup) => String(sup.room_id) === String(id),
    );
    const supIds = new Set(supsInRoom.map((s) => String(s.id)));
    const hasConfirmed = supsInRoom.some(
      (sup) => sup.confirmed_by && String(sup.confirmed_by).trim() !== "",
    );
    if (hasConfirmed) {
      return {
        status: "error",
        message:
          "Ruang ini sudah memiliki pengawas yang dikonfirmasi dan tidak dapat dihapus.",
      };
    }
    _examDeleteBySet_(EXAM_SHEET.BAP, 1, supIds);
    _examDeleteBySet_(EXAM_SHEET.SUPERVISORS, 1, new Set([String(id)]));
    _examDeleteById_(EXAM_SHEET.ROOMS, id);
    return { status: "success", message: "Ruang ujian berhasil dihapus." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function assignExamSupervisor(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menetapkan pengawas.",
      };
    }
    if (!data.room_id)
      return { status: "error", message: "ID ruang wajib diisi." };
    if (!data.user_id)
      return { status: "error", message: "ID guru wajib diisi." };
    const room = _examGetRoom_(data.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const targetGuru = getData("Users").find(
      (u) => String(u.id) === String(data.user_id),
    );
    if (!targetGuru)
      return { status: "error", message: "Guru tidak ditemukan." };
    const conflict = _examCheckSupervisorConflict_(
      data.user_id,
      room.session_id,
    );
    if (conflict) return { status: "error", message: conflict };
    const session = _examGetSession_(room.session_id);
    let jtmVal = Number(data.jtm_val);
    if (isNaN(jtmVal) || jtmVal < 0) {
      jtmVal = session ? Number(session.jtm_val) || 0 : 0;
    }
    const newId = "EXSUP-" + new Date().getTime();
    getSheet(EXAM_SHEET.SUPERVISORS).appendRow([
      newId,
      String(data.room_id),
      String(data.user_id),
      "active",
      false,
      "",
      "",
      "",
      "",
      "",
      jtmVal,
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Pengawas berhasil ditetapkan.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function editExamSupervisor(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat mengedit pengawas.",
      };
    }
    if (!data || !data.id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    if (!data.user_id)
      return { status: "error", message: "ID guru wajib diisi." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(data.id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message:
          "Pengawas yang sudah dikonfirmasi kehadirannya tidak dapat diedit. " +
          "Cabut konfirmasi kehadiran terlebih dahulu.",
      };
    }
    var isSub =
      target.is_substitute === true ||
      String(target.is_substitute).toLowerCase() === "true";
    if (isSub) {
      return {
        status: "error",
        message:
          "Pengawas pengganti tidak dapat diedit langsung. Gunakan tombol Batal Ganti terlebih dahulu.",
      };
    }
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return {
        status: "error",
        message: "Ruang ujian tidak ditemukan untuk pengawas ini.",
      };
    const targetGuru = getData("Users").find(
      (u) => String(u.id) === String(data.user_id),
    );
    if (!targetGuru)
      return { status: "error", message: "Guru tidak ditemukan." };
    var userChanged = String(target.user_id) !== String(data.user_id);
    if (userChanged) {
      const conflict = _examCheckSupervisorConflict_(
        data.user_id,
        room.session_id,
        data.id,
      );
      if (conflict) return { status: "error", message: conflict };
    }
    const session = _examGetSession_(room.session_id);
    let jtmVal = Number(data.jtm_val);
    if (isNaN(jtmVal) || jtmVal < 0) {
      jtmVal = session
        ? Number(session.jtm_val) || 0
        : Number(target.jtm_val) || 0;
    }
    const sheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const rows = sheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 3).setValue(String(data.user_id));
        sheet.getRange(i + 1, 11).setValue(jtmVal);
        updated = true;
        break;
      }
    }
    if (!updated)
      return {
        status: "error",
        message: "Gagal memperbarui data pengawas (baris tidak ditemukan).",
      };
    SpreadsheetApp.flush();
    return { status: "success", message: "Data pengawas berhasil diperbarui." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function removeExamSupervisor(token, supervisor_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menghapus pengawas.",
      };
    }
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message:
          "Pengawas yang sudah dikonfirmasi kehadirannya tidak dapat dihapus langsung. " +
          "Cabut konfirmasi kehadiran terlebih dahulu.",
      };
    }
    _examDeleteBySet_(EXAM_SHEET.BAP, 1, new Set([String(supervisor_id)]));
    _examDeleteById_(EXAM_SHEET.SUPERVISORS, supervisor_id);
    return {
      status: "success",
      message: "Pengawas berhasil dihapus dari jadwal.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function assignExamCommittee(token, data) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menetapkan panitia.",
      };
    }
    if (!data.period_id)
      return { status: "error", message: "ID periode wajib diisi." };
    if (!data.date) return { status: "error", message: "Tanggal wajib diisi." };
    if (!data.user_id)
      return { status: "error", message: "ID guru wajib diisi." };
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(data.period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const comDate = String(data.date);
    if (
      comDate < String(period.date_start) ||
      comDate > String(period.date_end)
    ) {
      return {
        status: "error",
        message:
          "Tanggal panitia harus berada dalam rentang periode ujian (" +
          period.date_start +
          " s.d. " +
          period.date_end +
          ").",
      };
    }
    const targetGuru = getData("Users").find(
      (u) => String(u.id) === String(data.user_id),
    );
    if (!targetGuru)
      return { status: "error", message: "Guru tidak ditemukan." };
    const conflict = _examCheckCommitteeConflict_(data.user_id, comDate);
    if (conflict) return { status: "error", message: conflict };
    const jtmVal = Number(period.jtm_committee_per_day) || 0;
    const newId = "EXCOM-" + new Date().getTime();
    getSheet(EXAM_SHEET.COMMITTEE).appendRow([
      newId,
      String(data.period_id),
      comDate,
      String(data.user_id),
      jtmVal,
      "active",
      false,
      "",
      "",
      "",
      "",
      "",
    ]);
    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Panitia berhasil ditetapkan.",
      id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function removeExamCommittee(token, committee_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat menghapus panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message:
          "Panitia yang sudah dikonfirmasi kehadirannya tidak dapat dihapus langsung. " +
          "Cabut konfirmasi kehadiran terlebih dahulu.",
      };
    }
    _examDeleteById_(EXAM_SHEET.COMMITTEE, committee_id);
    return {
      status: "success",
      message: "Panitia berhasil dihapus dari jadwal.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function getExamAdminMasterData(token) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return { status: "error", message: "Akses ditolak." };
    }
    const periodsResult = getExamPeriods(token);
    const periods =
      periodsResult.status === "success" ? periodsResult.data : [];
    const teachers = getData("Users")
      .filter((u) => String(u.role).toLowerCase() !== "admin")
      .map((u) => ({
        id: String(u.id),
        name: String(u.full_name || ""),
        nip: String(u.nip || ""),
        username: String(u.username || ""),
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
    return { status: "success", periods, teachers };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function getExamPeriodDetail(token, period_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return { status: "error", message: "Akses ditolak." };
    }
    if (!period_id)
      return { status: "error", message: "ID periode wajib diisi." };
    const periods = getData(EXAM_SHEET.PERIODS);
    const period = periods.find((p) => String(p.id) === String(period_id));
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const today = _examToday_();
    const ds = String(period.date_start);
    const de = String(period.date_end);
    const periodStatus =
      today < ds ? "Akan Datang" : today > de ? "Selesai" : "Berlangsung";
    const periodObj = {
      id: String(period.id),
      name: String(period.name || ""),
      date_start: ds,
      date_end: de,
      jtm_committee_per_day: Number(period.jtm_committee_per_day) || 0,
      description: String(period.description || ""),
      status: periodStatus,
    };
    const allUsers = getData("Users");
    const allSessions = getData(EXAM_SHEET.SESSIONS).filter(
      (s) => String(s.period_id) === String(period_id),
    );
    const allRooms = getData(EXAM_SHEET.ROOMS);
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const allBaps = getData(EXAM_SHEET.BAP);
    const sessions = allSessions
      .sort((a, b) => {
        const da = String(a.date),
          db = String(b.date);
        if (da !== db) return da.localeCompare(db);
        return _examFormatTime_(a.time_start).localeCompare(
          _examFormatTime_(b.time_start),
        );
      })
      .map((s) => {
        const rooms = allRooms
          .filter((r) => String(r.session_id) === String(s.id))
          .map((r) => {
            const supervisors = allSups
              .filter((sup) => String(sup.room_id) === String(r.id))
              .map((sup) => {
                const g = allUsers.find(
                  (u) => String(u.id) === String(sup.user_id),
                );
                const bapEntry = allBaps.find(
                  (b) =>
                    String(b.supervisor_id) === String(sup.id) &&
                    b.submitted_at &&
                    String(b.submitted_at).trim() !== "",
                );
                return {
                  id: String(sup.id),
                  user_id: String(sup.user_id),
                  guru_name: g ? String(g.full_name) : "Tidak Dikenal",
                  status: String(sup.status || "active"),
                  is_substitute:
                    sup.is_substitute === true ||
                    String(sup.is_substitute) === "true",
                  original_user_id: String(sup.original_user_id || ""),
                  confirmed_by: String(sup.confirmed_by || ""),
                  confirmed_at: String(sup.confirmed_at || ""),
                  substitution_cancelled_by: String(
                    sup.substitution_cancelled_by || "",
                  ),
                  substitution_cancelled_at: String(
                    sup.substitution_cancelled_at || "",
                  ),
                  jtm_val: Number(sup.jtm_val) || 0,
                  has_bap: !!bapEntry,
                };
              });
            return {
              id: String(r.id),
              room_name: String(r.room_name || ""),
              subject: String(r.subject || ""),
              class_name: String(r.class_name || ""),
              total_siswa: Number(r.total_siswa) || 0,
              supervisors,
            };
          })
          .sort((a, b) => a.room_name.localeCompare(b.room_name));
        return {
          id: String(s.id),
          date: String(s.date),
          session_name: String(s.session_name || ""),
          time_start: _examFormatTime_(s.time_start),
          time_end: _examFormatTime_(s.time_end),
          jtm_val: Number(s.jtm_val) || 0,
          rooms,
        };
      });
    const committee = getData(EXAM_SHEET.COMMITTEE)
      .filter((c) => String(c.period_id) === String(period_id))
      .map((c) => {
        const g = allUsers.find((u) => String(u.id) === String(c.user_id));
        return {
          id: String(c.id),
          date: String(c.date),
          user_id: String(c.user_id),
          guru_name: g ? String(g.full_name) : "Tidak Dikenal",
          jtm_val: Number(c.jtm_val) || 0,
          status: String(c.status || "active"),
          is_substitute:
            c.is_substitute === true || String(c.is_substitute) === "true",
          original_user_id: String(c.original_user_id || ""),
          confirmed_by: String(c.confirmed_by || ""),
          confirmed_at: String(c.confirmed_at || ""),
          substitution_cancelled_by: String(c.substitution_cancelled_by || ""),
          substitution_cancelled_at: String(c.substitution_cancelled_at || ""),
        };
      })
      .sort((a, b) => a.date.localeCompare(b.date));
    return { status: "success", period: periodObj, sessions, committee };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function _examAddJtmLog_(
  scheduleId,
  userId,
  date,
  materi,
  hadir,
  absen,
  jtmVal,
  refId,
) {
  const logId = "LOG-EXAM-" + new Date().getTime();
  getSheet(SHEET_NAME.LOGS).appendRow([
    logId,
    String(scheduleId),
    String(userId),
    String(date),
    String(materi),
    Number(hadir) || 0,
    Number(absen) || 0,
    "REF:" + String(refId),
    Number(jtmVal) || 0,
    _examNow_(),
  ]);
  return logId;
}

function _examDeleteLogByRef_(refId) {
  const sheet = getSheet(SHEET_NAME.LOGS);
  const rows = sheet.getDataRange().getValues();
  const REF = "REF:" + String(refId);
  let deleted = false;
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][7]).indexOf(REF) !== -1) {
      sheet.deleteRow(i + 1);
      deleted = true;
    }
  }
  return deleted;
}

function _examUpdateLogByRef_(refId, hadir, absen) {
  const sheet = getSheet(SHEET_NAME.LOGS);
  const rows = sheet.getDataRange().getValues();
  const REF = "REF:" + String(refId);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][7]).indexOf(REF) !== -1) {
      sheet.getRange(i + 1, 6).setValue(Number(hadir) || 0);
      sheet.getRange(i + 1, 7).setValue(Number(absen) || 0);
      return true;
    }
  }
  return false;
}

function _examGetBap_(supervisor_id) {
  return (
    getData(EXAM_SHEET.BAP).find(
      (b) => String(b.supervisor_id) === String(supervisor_id),
    ) || null
  );
}

function _examIsConfirmedPanitia_(user_id, date) {
  return getData(EXAM_SHEET.COMMITTEE).some(
    (c) =>
      String(c.user_id) === String(user_id) &&
      String(c.date) === String(date) &&
      String(c.status) === "active" &&
      c.confirmed_by &&
      String(c.confirmed_by).trim() !== "",
  );
}

function _examCheckAdminOrConfirmedPanitia_(user, date) {
  if (String(user.role).toLowerCase() === "admin") {
    return { ok: true, isAdmin: true };
  }
  if (_examIsConfirmedPanitia_(user.id, date)) {
    return { ok: true, isAdmin: false };
  }
  return {
    ok: false,
    isAdmin: false,
    error:
      "Akses ditolak. Hanya Admin atau Panitia yang sudah dikonfirmasi yang dapat melakukan aksi ini pada hari H.",
  };
}

function _examCheckBapTiming_(sessionDate, timeEnd) {
  const today = _examToday_();
  if (today < sessionDate) {
    return "Sesi ujian belum berlangsung.";
  }
  if (today > sessionDate) {
    return null;
  }
  const [teH, teM] = String(timeEnd).split(":").map(Number);
  const activeTotal = teH * 60 + teM + 5;
  const activeTime =
    String(Math.floor(activeTotal / 60)).padStart(2, "0") +
    ":" +
    String(activeTotal % 60).padStart(2, "0");
  const nowTime = Utilities.formatDate(new Date(), "Asia/Jakarta", "HH:mm");
  if (nowTime < activeTime) {
    return (
      "BAP belum bisa diisi. Sesi berakhir pukul " +
      timeEnd +
      ", tombol aktif mulai pukul " +
      activeTime +
      "."
    );
  }
  return null;
}

function confirmCommitteeAttendance(token, committee_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat mengonfirmasi kehadiran Panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    if (_examToday_() !== String(target.date)) {
      return {
        status: "error",
        message:
          "Konfirmasi hanya dapat dilakukan pada hari pelaksanaan (hari H).",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord panitia ini tidak aktif." };
    }
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message: "Panitia ini sudah dikonfirmasi kehadirannya.",
      };
    }
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(target.period_id),
    );
    const periodName = period ? String(period.name) : "Ujian";
    const now = _examNow_();
    const sheet = getSheet(EXAM_SHEET.COMMITTEE);
    const rows = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(committee_id)) {
        sheet.getRange(i + 1, 9).setValue(String(user.id));
        sheet.getRange(i + 1, 10).setValue(now);
        found = true;
        break;
      }
    }
    if (!found)
      return { status: "error", message: "Gagal memperbarui data konfirmasi." };
    _examAddJtmLog_(
      "EXAM-COMMITTEE",
      target.user_id,
      String(target.date),
      "Panitia Ujian: " + periodName,
      0,
      0,
      Number(target.jtm_val) || 0,
      committee_id,
    );
    try {
      _notifCommitteeConfirmed_(target.user_id, target.date);
    } catch (_) {}
    return {
      status: "success",
      message: "Kehadiran Panitia berhasil dikonfirmasi. JTM telah dicatat.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function revokeCommitteeAttendance(token, committee_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat mencabut konfirmasi Panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    if (_examToday_() !== String(target.date)) {
      return {
        status: "error",
        message:
          "Konfirmasi hanya dapat dicabut pada hari pelaksanaan (hari H).",
      };
    }
    if (!target.confirmed_by || String(target.confirmed_by).trim() === "") {
      return {
        status: "error",
        message: "Panitia ini belum dikonfirmasi kehadirannya.",
      };
    }
    _examDeleteLogByRef_(committee_id);
    const sheet = getSheet(EXAM_SHEET.COMMITTEE);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(committee_id)) {
        sheet.getRange(i + 1, 9).setValue("");
        sheet.getRange(i + 1, 10).setValue("");
        break;
      }
    }
    try {
      _notifCommitteeRevoked_(target.user_id, target.date);
    } catch (_) {}
    return {
      status: "success",
      message:
        "Konfirmasi kehadiran Panitia berhasil dicabut. JTM telah dihapus dari honorarium.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function confirmSupervisorAttendance(token, supervisor_id) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const sessionDate = String(session.date);
    if (_examToday_() !== sessionDate) {
      return {
        status: "error",
        message:
          "Konfirmasi hanya dapat dilakukan pada hari pelaksanaan (hari H).",
      };
    }
    const auth = _examCheckAdminOrConfirmedPanitia_(user, sessionDate);
    if (!auth.ok) return { status: "error", message: auth.error };
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord pengawas ini tidak aktif." };
    }
    if (target.confirmed_by && String(target.confirmed_by).trim() !== "") {
      return {
        status: "error",
        message: "Pengawas ini sudah dikonfirmasi kehadirannya.",
      };
    }
    const sheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(supervisor_id)) {
        sheet.getRange(i + 1, 7).setValue(String(user.id));
        sheet.getRange(i + 1, 8).setValue(_examNow_());
        try {
          var byName = String(
            user.full_name ||
              user.name ||
              user.username ||
              (auth.isAdmin ? "Admin" : "Panitia Ujian"),
          );
          var sessionInfo =
            String(session.session_name || "") +
            " · " +
            String(room.room_name || "") +
            " (" +
            String(room.subject || "") +
            ")";
          _notifSupervisorConfirmed_(
            target.user_id,
            sessionDate,
            byName,
            sessionInfo,
          );
        } catch (_) {}
        return {
          status: "success",
          message:
            "Kehadiran Pengawas berhasil dikonfirmasi. Pengawas dapat mengisi BAP setelah sesi berakhir + 5 menit.",
        };
      }
    }
    return { status: "error", message: "Gagal memperbarui data konfirmasi." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function revokeSupervisorAttendance(token, supervisor_id, force) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const sessionDate = String(session.date);
    if (_examToday_() !== sessionDate) {
      return {
        status: "error",
        message:
          "Konfirmasi hanya dapat dicabut pada hari pelaksanaan (hari H).",
      };
    }
    const auth = _examCheckAdminOrConfirmedPanitia_(user, sessionDate);
    if (!auth.ok) return { status: "error", message: auth.error };
    if (!target.confirmed_by || String(target.confirmed_by).trim() === "") {
      return {
        status: "error",
        code: JTM_ERROR_CODES.NO_ACTIVE_CONFIRMATION,
        message: "Pengawas ini belum dikonfirmasi kehadirannya.",
      };
    }
    const bap = _examGetBap_(supervisor_id);
    const hasBap = !!(
      bap &&
      bap.submitted_at &&
      String(bap.submitted_at).trim() !== ""
    );
    if (hasBap && !force) {
      return {
        status: "confirm_required",
        message:
          "Pencabutan konfirmasi akan mengunci BAP Pengawas dan menghapus JTM-nya dari honorarium. Lanjutkan?",
      };
    }
    if (hasBap) {
      _examDeleteLogByRef_(supervisor_id);
    }
    const sheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(supervisor_id)) {
        sheet.getRange(i + 1, 7).setValue("");
        sheet.getRange(i + 1, 8).setValue("");
        break;
      }
    }
    try {
      _jtmReverseOccurrence_("EXAM", supervisor_id, sessionDate);
    } catch (_) {}
    try {
      var byName = String(
        user.full_name ||
          user.name ||
          user.username ||
          (auth.isAdmin ? "Admin" : "Panitia Ujian"),
      );
      var sessionInfo =
        String(session.session_name || "") +
        " · " +
        String(room.room_name || "") +
        " (" +
        String(room.subject || "") +
        ")";
      _notifSupervisorRevoked_(
        target.user_id,
        sessionDate,
        byName,
        sessionInfo,
      );
    } catch (_) {}
    return {
      status: "success",
      message: hasBap
        ? "Konfirmasi dicabut. BAP terkunci dan JTM telah dihapus dari honorarium."
        : "Konfirmasi kehadiran Pengawas berhasil dicabut.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function submitExamBAP(token, supervisor_id, bapData) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (String(target.user_id) !== String(user.id)) {
      return {
        status: "error",
        message: "Anda tidak terdaftar sebagai pengawas pada jadwal ini.",
      };
    }
    if (String(target.status) === "cancelled") {
      return {
        status: "error",
        message: "Substitusi Anda pada jadwal ini telah dibatalkan.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord pengawas ini tidak aktif." };
    }
    if (!target.confirmed_by || String(target.confirmed_by).trim() === "") {
      return {
        status: "error",
        message:
          "Kehadiran Anda belum dikonfirmasi Panitia. BAP belum dapat diisi.",
      };
    }
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(session.period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    if (_examToday_() > String(period.date_end)) {
      return {
        status: "error",
        message: "Periode ujian sudah selesai. BAP tidak dapat diubah.",
      };
    }
    const timingError = _examCheckBapTiming_(
      String(session.date),
      _examFormatTime_(session.time_end),
    );
    if (timingError) return { status: "error", message: timingError };
    const existing = _examGetBap_(supervisor_id);
    if (
      existing &&
      existing.submitted_at &&
      String(existing.submitted_at).trim() !== ""
    ) {
      return {
        status: "error",
        message:
          "BAP sudah pernah disubmit. Gunakan fitur Edit BAP untuk mengubah isian.",
      };
    }
    const hadir = Number(bapData.peserta_hadir);
    const absen = Number(bapData.peserta_absen);
    if (isNaN(hadir) || hadir < 0 || isNaN(absen) || absen < 0) {
      return {
        status: "error",
        message:
          "Jumlah peserta hadir dan tidak hadir wajib diisi dan tidak boleh negatif.",
      };
    }
    const totalSiswaAdmin = Number(room.total_siswa) || 0;
    const totalSiswaBAP = hadir + absen;
    if (totalSiswaAdmin > 0 && totalSiswaBAP !== totalSiswaAdmin) {
      return {
        status: "error",
        message: `Jumlah total siswa tidak sesuai. Total siswa di ruang ini: ${totalSiswaAdmin} (Hadir: ${hadir} + Alpa: ${absen} = ${totalSiswaBAP}). Silakan periksa kembali data kehadiran.`,
      };
    }
    const now = _examNow_();
    const catatan = String(bapData.catatan || "").trim();
    const bapId = "BAP-" + new Date().getTime();
    getSheet(EXAM_SHEET.BAP).appendRow([
      bapId,
      String(supervisor_id),
      hadir,
      absen,
      catatan,
      now,
      now,
    ]);
    let jtmToLog =
      Number(target.jtm_val) > 0
        ? Number(target.jtm_val)
        : Number(session.jtm_val) || 0;
    const jtmAdjustment = _jtmReadAdjustment_("EXAM", supervisor_id);
    if (jtmAdjustment) {
      const tz = Session.getScriptTimeZone();
      const sessionDateStr =
        session.date instanceof Date
          ? Utilities.formatDate(session.date, tz, "yyyy-MM-dd")
          : String(session.date);
      const adjDateStr =
        jtmAdjustment.date instanceof Date
          ? Utilities.formatDate(jtmAdjustment.date, tz, "yyyy-MM-dd")
          : String(jtmAdjustment.date);
      const storedAdjusted = jtmAdjustment.adjusted_jtm;
      const dateMatches = !adjDateStr || adjDateStr === sessionDateStr;
      if (
        dateMatches &&
        storedAdjusted !== null &&
        storedAdjusted !== undefined &&
        storedAdjusted !== "" &&
        !isNaN(Number(storedAdjusted))
      ) {
        jtmToLog = Number(storedAdjusted);
      }
    }
    _examAddJtmLog_(
      "EXAM-SUPERVISOR",
      target.user_id,
      String(session.date),
      "Pengawas Ujian: " +
        String(session.session_name || "") +
        " — " +
        String(room.room_name || ""),
      hadir,
      absen,
      jtmToLog,
      supervisor_id,
    );
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message:
        "BAP berhasil dikirim. JTM Pengawas telah dicatat ke sistem honorarium.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function editExamBAP(token, supervisor_id, bapData) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (String(target.user_id) !== String(user.id)) {
      return {
        status: "error",
        message: "Anda tidak terdaftar sebagai pengawas pada jadwal ini.",
      };
    }
    if (String(target.status) === "cancelled") {
      return {
        status: "error",
        message:
          "Substitusi Anda pada jadwal ini telah dibatalkan. BAP bersifat read-only.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord pengawas ini tidak aktif." };
    }
    if (!target.confirmed_by || String(target.confirmed_by).trim() === "") {
      return {
        status: "error",
        message: "Kehadiran Anda belum dikonfirmasi. BAP tidak dapat diedit.",
      };
    }
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(session.period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    if (_examToday_() > String(period.date_end)) {
      return {
        status: "error",
        message: "Periode ujian sudah selesai. BAP tidak dapat diubah.",
      };
    }
    const timingError = _examCheckBapTiming_(
      String(session.date),
      _examFormatTime_(session.time_end),
    );
    if (timingError) return { status: "error", message: timingError };
    const existing = _examGetBap_(supervisor_id);
    if (
      !existing ||
      !existing.submitted_at ||
      String(existing.submitted_at).trim() === ""
    ) {
      return {
        status: "error",
        message:
          "BAP belum pernah disubmit. Gunakan Submit BAP terlebih dahulu.",
      };
    }
    const hadir = Number(bapData.peserta_hadir);
    const absen = Number(bapData.peserta_absen);
    if (isNaN(hadir) || hadir < 0 || isNaN(absen) || absen < 0) {
      return {
        status: "error",
        message:
          "Jumlah peserta hadir dan tidak hadir wajib diisi dan tidak boleh negatif.",
      };
    }
    const totalSiswaAdmin = Number(room.total_siswa) || 0;
    const totalSiswaBAP = hadir + absen;
    if (totalSiswaAdmin > 0 && totalSiswaBAP !== totalSiswaAdmin) {
      return {
        status: "error",
        message: `Jumlah total siswa tidak sesuai. Total siswa di ruang ini: ${totalSiswaAdmin} (Hadir: ${hadir} + Alpa: ${absen} = ${totalSiswaBAP}). Silakan periksa kembali data kehadiran.`,
      };
    }
    const now = _examNow_();
    const catatan = String(bapData.catatan || "").trim();
    const bapSheet = getSheet(EXAM_SHEET.BAP);
    const bapRows = bapSheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < bapRows.length; i++) {
      if (String(bapRows[i][1]) === String(supervisor_id)) {
        bapSheet.getRange(i + 1, 3).setValue(hadir);
        bapSheet.getRange(i + 1, 4).setValue(absen);
        bapSheet.getRange(i + 1, 5).setValue(catatan);
        bapSheet.getRange(i + 1, 7).setValue(now);
        updated = true;
        break;
      }
    }
    if (!updated)
      return { status: "error", message: "Gagal memperbarui data BAP." };
    _examUpdateLogByRef_(supervisor_id, hadir, absen);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return { status: "success", message: "BAP berhasil diperbarui." };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function substituteExamSupervisor(token, supervisor_id, new_user_id) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    if (!new_user_id)
      return { status: "error", message: "ID guru pengganti tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const sessionDate = String(session.date);
    if (_examToday_() !== sessionDate) {
      return {
        status: "error",
        message:
          "Batas waktu substitusi sudah terlewat. Substitusi hanya dapat dilakukan pada hari H.",
      };
    }
    const auth = _examCheckAdminOrConfirmedPanitia_(user, sessionDate);
    if (!auth.ok) return { status: "error", message: auth.error };
    if (String(target.status) !== "active") {
      return {
        status: "error",
        message: "Rekord pengawas ini tidak aktif dan tidak dapat digantikan.",
      };
    }
    const newGuru = getData("Users").find(
      (u) => String(u.id) === String(new_user_id),
    );
    if (!newGuru)
      return { status: "error", message: "Guru pengganti tidak ditemukan." };
    if (String(new_user_id) === String(target.user_id)) {
      return {
        status: "error",
        message: "Guru pengganti tidak boleh sama dengan guru yang diganti.",
      };
    }
    const conflict = _examCheckSupervisorConflict_(
      new_user_id,
      room.session_id,
      supervisor_id,
    );
    if (conflict)
      return { status: "error", message: "Guru pengganti: " + conflict };
    const supSheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const supRows = supSheet.getDataRange().getValues();
    for (let i = 1; i < supRows.length; i++) {
      if (String(supRows[i][0]) === String(supervisor_id)) {
        supSheet.getRange(i + 1, 4).setValue("substituted");
        break;
      }
    }
    const newId = "EXSUP-" + new Date().getTime();
    supSheet.appendRow([
      newId,
      String(target.room_id),
      String(new_user_id),
      "active",
      true,
      String(target.user_id),
      "",
      "",
      "",
      "",
      Number(target.jtm_val) || 0,
    ]);
    try {
      var byName = String(
        user.full_name ||
          user.name ||
          user.username ||
          (auth.isAdmin ? "Admin" : "Panitia Ujian"),
      );
      var sessionInfo =
        String(session.session_name || "") +
        " · " +
        String(room.room_name || "") +
        " (" +
        String(room.subject || "") +
        ")";
      _notifSupervisorSubstituted_(
        target.user_id,
        new_user_id,
        sessionDate,
        byName,
        sessionInfo,
      );
    } catch (_) {}
    return {
      status: "success",
      message:
        String(newGuru.full_name) + " kini menggantikan posisi pengawas.",
      new_supervisor_id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function cancelExamSupervisorSubstitution(token, supervisor_id, force) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisor_id)
      return { status: "error", message: "ID pengawas tidak ditemukan." };
    const target = getData(EXAM_SHEET.SUPERVISORS).find(
      (s) => String(s.id) === String(supervisor_id),
    );
    if (!target)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    if (!(
      target.is_substitute === true || String(target.is_substitute) === "true"
    )) {
      return {
        status: "error",
        message: "Rekord ini bukan pengganti. Pembatalan tidak berlaku.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Substitusi ini sudah tidak aktif." };
    }
    const room = _examGetRoom_(target.room_id);
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const session = _examGetSession_(room.session_id);
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const sessionDate = String(session.date);
    if (_examToday_() !== sessionDate) {
      return {
        status: "error",
        message: "Pembatalan substitusi hanya dapat dilakukan pada hari H.",
      };
    }
    const auth = _examCheckAdminOrConfirmedPanitia_(user, sessionDate);
    if (!auth.ok) return { status: "error", message: auth.error };
    const isConfirmed = !!(
      target.confirmed_by && String(target.confirmed_by).trim() !== ""
    );
    const bap = _examGetBap_(supervisor_id);
    const hasBap = !!(
      bap &&
      bap.submitted_at &&
      String(bap.submitted_at).trim() !== ""
    );
    if (isConfirmed && hasBap && !force) {
      return {
        status: "confirm_required",
        message:
          "Guru pengganti sudah mengisi dan menyerahkan BAP. " +
          "Pembatalan akan mengunci BAP pengganti dan menghapus JTM-nya dari honorarium. " +
          "Guru asli akan dikembalikan ke jadwal. Lanjutkan?",
      };
    }
    const now = _examNow_();
    const supSheet = getSheet(EXAM_SHEET.SUPERVISORS);
    const supRows = supSheet.getDataRange().getValues();
    if (isConfirmed && hasBap) {
      for (let i = 1; i < supRows.length; i++) {
        if (String(supRows[i][0]) === String(supervisor_id)) {
          supSheet.getRange(i + 1, 4).setValue("cancelled");
          supSheet.getRange(i + 1, 9).setValue(String(user.id));
          supSheet.getRange(i + 1, 10).setValue(now);
          break;
        }
      }
      _examDeleteLogByRef_(supervisor_id);
    } else {
      for (let i = supRows.length - 1; i >= 1; i--) {
        if (String(supRows[i][0]) === String(supervisor_id)) {
          supSheet.deleteRow(i + 1);
          break;
        }
      }
    }
    const originalUserId = String(target.original_user_id);
    const freshRows = getSheet(EXAM_SHEET.SUPERVISORS)
      .getDataRange()
      .getValues();
    for (let i = 1; i < freshRows.length; i++) {
      if (
        String(freshRows[i][1]) === String(target.room_id) &&
        String(freshRows[i][2]) === originalUserId &&
        String(freshRows[i][3]) === "substituted"
      ) {
        const s = getSheet(EXAM_SHEET.SUPERVISORS);
        s.getRange(i + 1, 4).setValue("active");
        s.getRange(i + 1, 7).setValue("");
        s.getRange(i + 1, 8).setValue("");
        break;
      }
    }
    try {
      var byName = String(
        user.full_name ||
          user.name ||
          user.username ||
          (auth.isAdmin ? "Admin" : "Panitia Ujian"),
      );
      var sessionInfo =
        String(session.session_name || "") +
        " · " +
        String(room.room_name || "") +
        " (" +
        String(room.subject || "") +
        ")";
      _notifSupervisorSubCancelled_(
        target.user_id,
        sessionDate,
        byName,
        sessionInfo,
      );
    } catch (_) {}
    return {
      status: "success",
      message:
        "Substitusi pengawas berhasil dibatalkan. Guru asli telah dikembalikan ke jadwal.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function substituteExamCommittee(token, committee_id, new_user_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat mengganti Panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    if (!new_user_id)
      return { status: "error", message: "ID guru pengganti tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    const comDate = String(target.date);
    if (_examToday_() !== comDate) {
      return {
        status: "error",
        message: "Substitusi hanya dapat dilakukan pada hari H.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Rekord panitia ini tidak aktif." };
    }
    const newGuru = getData("Users").find(
      (u) => String(u.id) === String(new_user_id),
    );
    if (!newGuru)
      return { status: "error", message: "Guru pengganti tidak ditemukan." };
    if (String(new_user_id) === String(target.user_id)) {
      return {
        status: "error",
        message: "Guru pengganti tidak boleh sama dengan guru yang diganti.",
      };
    }
    const conflict = _examCheckCommitteeConflict_(
      new_user_id,
      comDate,
      committee_id,
    );
    if (conflict)
      return { status: "error", message: "Guru pengganti: " + conflict };
    const comSheet = getSheet(EXAM_SHEET.COMMITTEE);
    const comRows = comSheet.getDataRange().getValues();
    for (let i = 1; i < comRows.length; i++) {
      if (String(comRows[i][0]) === String(committee_id)) {
        comSheet.getRange(i + 1, 6).setValue("substituted");
        break;
      }
    }
    const newId = "EXCOM-" + new Date().getTime();
    const jtmVal = Number(target.jtm_val) || 0;
    comSheet.appendRow([
      newId,
      String(target.period_id),
      comDate,
      String(new_user_id),
      jtmVal,
      "active",
      true,
      String(target.user_id),
      "",
      "",
      "",
      "",
    ]);
    try {
      _notifCommitteeSubstituted_(target.user_id, new_user_id, comDate);
    } catch (_) {}
    return {
      status: "success",
      message: String(newGuru.full_name) + " kini menggantikan posisi panitia.",
      new_committee_id: newId,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function cancelExamCommitteeSubstitution(token, committee_id, force) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message:
          "Akses ditolak. Hanya Admin yang dapat membatalkan substitusi Panitia.",
      };
    }
    if (!committee_id)
      return { status: "error", message: "ID panitia tidak ditemukan." };
    const target = getData(EXAM_SHEET.COMMITTEE).find(
      (c) => String(c.id) === String(committee_id),
    );
    if (!target)
      return { status: "error", message: "Data panitia tidak ditemukan." };
    if (!(
      target.is_substitute === true || String(target.is_substitute) === "true"
    )) {
      return {
        status: "error",
        message: "Rekord ini bukan pengganti. Pembatalan tidak berlaku.",
      };
    }
    if (String(target.status) !== "active") {
      return { status: "error", message: "Substitusi ini sudah tidak aktif." };
    }
    const comDate = String(target.date);
    if (_examToday_() !== comDate) {
      return {
        status: "error",
        message: "Pembatalan substitusi hanya dapat dilakukan pada hari H.",
      };
    }
    const isConfirmed = !!(
      target.confirmed_by && String(target.confirmed_by).trim() !== ""
    );
    if (isConfirmed && !force) {
      return {
        status: "confirm_required",
        message:
          "Guru pengganti sudah dikonfirmasi kehadirannya dan JTM-nya sudah terhitung. " +
          "Pembatalan akan menghapus JTM pengganti dari honorarium. " +
          "Guru asli akan dikembalikan ke jadwal. Lanjutkan?",
      };
    }
    const now = _examNow_();
    const comSheet = getSheet(EXAM_SHEET.COMMITTEE);
    const comRows = comSheet.getDataRange().getValues();
    if (isConfirmed) {
      for (let i = 1; i < comRows.length; i++) {
        if (String(comRows[i][0]) === String(committee_id)) {
          comSheet.getRange(i + 1, 6).setValue("cancelled");
          comSheet.getRange(i + 1, 11).setValue(String(user.id));
          comSheet.getRange(i + 1, 12).setValue(now);
          break;
        }
      }
      _examDeleteLogByRef_(committee_id);
    } else {
      for (let i = comRows.length - 1; i >= 1; i--) {
        if (String(comRows[i][0]) === String(committee_id)) {
          comSheet.deleteRow(i + 1);
          break;
        }
      }
    }
    const originalUserId = String(target.original_user_id);
    const freshRows = getSheet(EXAM_SHEET.COMMITTEE).getDataRange().getValues();
    for (let i = 1; i < freshRows.length; i++) {
      if (
        String(freshRows[i][1]) === String(target.period_id) &&
        String(freshRows[i][2]) === comDate &&
        String(freshRows[i][3]) === originalUserId &&
        String(freshRows[i][5]) === "substituted"
      ) {
        const s = getSheet(EXAM_SHEET.COMMITTEE);
        s.getRange(i + 1, 6).setValue("active");
        s.getRange(i + 1, 9).setValue("");
        s.getRange(i + 1, 10).setValue("");
        break;
      }
    }
    try {
      _notifCommitteeSubCancelled_(target.user_id, comDate);
    } catch (_) {}
    return {
      status: "success",
      message:
        "Substitusi panitia berhasil dibatalkan. Guru asli telah dikembalikan ke jadwal.",
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function getExamPageData(token, dateStr) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    const targetDate = dateStr ? String(dateStr) : _examToday_();
    const activePeriod = getData(EXAM_SHEET.PERIODS).find(
      (p) =>
        targetDate >= String(p.date_start) && targetDate <= String(p.date_end),
    );
    if (!activePeriod) {
      return {
        status: "success",
        is_exam_period: false,
        supervisor_duties: [],
        committee_duties: [],
        role: "none",
        is_confirmed_panitia: false,
        supervisors_to_confirm: [],
      };
    }
    const periodObj = {
      id: String(activePeriod.id),
      name: String(activePeriod.name || ""),
      date_start: String(activePeriod.date_start),
      date_end: String(activePeriod.date_end),
      jtm_committee_per_day: Number(activePeriod.jtm_committee_per_day) || 0,
      description: String(activePeriod.description || ""),
    };
    const allUsers = getData("Users");
    const allSessions = getData(EXAM_SHEET.SESSIONS).filter(
      (s) =>
        String(s.period_id) === String(activePeriod.id) &&
        String(s.date) === targetDate,
    );
    const allRooms = getData(EXAM_SHEET.ROOMS);
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const allBaps = getData(EXAM_SHEET.BAP);
    const allCom = getData(EXAM_SHEET.COMMITTEE);
    const supervisorDuties = [];
    allSessions.forEach((s) => {
      const te = _examFormatTime_(s.time_end);
      const [teH, teM] = te.split(":").map(Number);
      const activeTotal = teH * 60 + teM + 5;
      const bapActiveTime =
        String(Math.floor(activeTotal / 60)).padStart(2, "0") +
        ":" +
        String(activeTotal % 60).padStart(2, "0");
      allRooms
        .filter((r) => String(r.session_id) === String(s.id))
        .forEach((r) => {
          const mySup = allSups.find(
            (sup) =>
              String(sup.room_id) === String(r.id) &&
              String(sup.user_id) === String(user.id) &&
              String(sup.status) === "active",
          );
          if (!mySup) return;
          const bap = allBaps.find(
            (b) => String(b.supervisor_id) === String(mySup.id),
          );
          const hasBap = !!(
            bap &&
            bap.submitted_at &&
            String(bap.submitted_at).trim() !== ""
          );
          supervisorDuties.push({
            supervisor_id: String(mySup.id),
            session_id: String(s.id),
            session_name: String(s.session_name || ""),
            date: String(s.date),
            time_start: _examFormatTime_(s.time_start),
            time_end: te,
            bap_active_time: bapActiveTime,
            jtm_val: Number(mySup.jtm_val) || 0,
            room_id: String(r.id),
            room_name: String(r.room_name || ""),
            subject: String(r.subject || ""),
            class_name: String(r.class_name || ""),
            is_substitute:
              mySup.is_substitute === true ||
              String(mySup.is_substitute) === "true",
            confirmed_by: String(mySup.confirmed_by || ""),
            confirmed_at: String(mySup.confirmed_at || ""),
            has_bap: hasBap,
            bap: hasBap
              ? {
                  id: String(bap.id),
                  peserta_hadir: Number(bap.peserta_hadir) || 0,
                  peserta_absen: Number(bap.peserta_absen) || 0,
                  catatan: String(bap.catatan || ""),
                  submitted_at: String(bap.submitted_at || ""),
                  updated_at: String(bap.updated_at || ""),
                }
              : null,
          });
        });
    });
    const committeeDuties = allCom
      .filter(
        (c) =>
          String(c.period_id) === String(activePeriod.id) &&
          String(c.date) === targetDate &&
          String(c.user_id) === String(user.id) &&
          String(c.status) === "active",
      )
      .map((c) => ({
        committee_id: String(c.id),
        date: String(c.date),
        jtm_val: Number(c.jtm_val) || 0,
        is_substitute:
          c.is_substitute === true || String(c.is_substitute) === "true",
        confirmed_by: String(c.confirmed_by || ""),
        confirmed_at: String(c.confirmed_at || ""),
      }));
    const isConfirmedPanitia = committeeDuties.some(
      (c) => c.confirmed_by && String(c.confirmed_by).trim() !== "",
    );
    let role = "none";
    if (supervisorDuties.length > 0) role = "supervisor";
    else if (committeeDuties.length > 0) role = "committee";
    const supervisorsToConfirm = [];
    if (isConfirmedPanitia) {
      allSessions.forEach((s) => {
        const roomsInSession = allRooms.filter(
          (r) => String(r.session_id) === String(s.id),
        );
        roomsInSession.forEach((r) => {
          const supsInRoom = allSups
            .filter((sup) => String(sup.room_id) === String(r.id))
            .map((sup) => {
              const g = allUsers.find(
                (u) => String(u.id) === String(sup.user_id),
              );
              const bap = allBaps.find(
                (b) => String(b.supervisor_id) === String(sup.id),
              );
              return {
                supervisor_id: String(sup.id),
                user_id: String(sup.user_id),
                guru_name: g ? String(g.full_name) : "Tidak Dikenal",
                status: String(sup.status || "active"),
                is_substitute:
                  sup.is_substitute === true ||
                  String(sup.is_substitute) === "true",
                original_user_id: String(sup.original_user_id || ""),
                confirmed_by: String(sup.confirmed_by || ""),
                confirmed_at: String(sup.confirmed_at || ""),
                has_bap: !!(
                  bap &&
                  bap.submitted_at &&
                  String(bap.submitted_at).trim() !== ""
                ),
              };
            });
          if (supsInRoom.length > 0) {
            supervisorsToConfirm.push({
              session_id: String(s.id),
              session_name: String(s.session_name || ""),
              time_start: _examFormatTime_(s.time_start),
              time_end: _examFormatTime_(s.time_end),
              room_id: String(r.id),
              room_name: String(r.room_name || ""),
              subject: String(r.subject || ""),
              supervisors: supsInRoom,
            });
          }
        });
      });
    }
    return {
      status: "success",
      is_exam_period: true,
      period: periodObj,
      role,
      supervisor_duties: supervisorDuties,
      committee_duties: committeeDuties,
      is_confirmed_panitia: isConfirmedPanitia,
      supervisors_to_confirm: supervisorsToConfirm,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function getSupervisorBapDetail(token, supervisorId) {
  try {
    const user = verifySession(token);
    if (!user)
      return {
        status: "error",
        message: "Sesi tidak valid, silakan login kembali.",
      };
    if (!supervisorId)
      return { status: "error", message: "Supervisor ID kosong." };
    const sups = getData(EXAM_SHEET.SUPERVISORS);
    const sup = sups.find((s) => String(s.id) === String(supervisorId));
    if (!sup)
      return { status: "error", message: "Data pengawas tidak ditemukan." };
    const isOwner = String(sup.user_id) === String(user.id);
    const isAdmin = String(user.role).toLowerCase() === "admin";
    if (!isOwner && !isAdmin) {
      return {
        status: "error",
        message: "Anda tidak memiliki akses ke BAP ini.",
      };
    }
    const rooms = getData(EXAM_SHEET.ROOMS);
    const room = rooms.find((r) => String(r.id) === String(sup.room_id));
    if (!room)
      return { status: "error", message: "Ruang ujian tidak ditemukan." };
    const sessions = getData(EXAM_SHEET.SESSIONS);
    const session = sessions.find(
      (s) => String(s.id) === String(room.session_id),
    );
    if (!session)
      return { status: "error", message: "Sesi ujian tidak ditemukan." };
    const periods = getData(EXAM_SHEET.PERIODS);
    const period = periods.find(
      (p) => String(p.id) === String(session.period_id),
    );
    const baps = getData(EXAM_SHEET.BAP);
    const bap = baps.find(
      (b) => String(b.supervisor_id) === String(supervisorId),
    );
    const hasBap = !!(
      bap &&
      bap.submitted_at &&
      String(bap.submitted_at).trim() !== ""
    );
    const today = _examToday_();
    const periodOver = !!(period && today > String(period.date_end));
    const isCancelled = String(sup.status) === "cancelled";
    const te = _examFormatTime_(session.time_end);
    return {
      status: "success",
      supervisor_id: String(sup.id),
      session_name: String(session.session_name || ""),
      date: String(session.date),
      time_start: _examFormatTime_(session.time_start),
      time_end: te,
      jtm_val: Number(session.jtm_val) || 0,
      room_name: String(room.room_name || ""),
      subject: String(room.subject || ""),
      class_name: String(room.class_name || ""),
      is_substitute:
        sup.is_substitute === true || String(sup.is_substitute) === "true",
      status: String(sup.status || ""),
      is_cancelled: isCancelled,
      period_over: periodOver,
      period: period
        ? {
            id: String(period.id),
            name: String(period.name || ""),
            date_start: String(period.date_start),
            date_end: String(period.date_end),
          }
        : null,
      has_bap: hasBap,
      bap: hasBap
        ? {
            id: String(bap.id),
            peserta_hadir: Number(bap.peserta_hadir) || 0,
            peserta_absen: Number(bap.peserta_absen) || 0,
            catatan: String(bap.catatan || ""),
            submitted_at: String(bap.submitted_at || ""),
            updated_at: String(bap.updated_at || ""),
          }
        : null,
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}

function getExamRecapData(token, period_id) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Akses ditolak. Hanya Admin yang dapat melihat rekap.",
      };
    }
    if (!period_id)
      return { status: "error", message: "ID periode wajib diisi." };
    const period = getData(EXAM_SHEET.PERIODS).find(
      (p) => String(p.id) === String(period_id),
    );
    if (!period)
      return { status: "error", message: "Periode ujian tidak ditemukan." };
    const today = _examToday_();
    const ds = String(period.date_start);
    const de = String(period.date_end);
    const pStatus =
      today < ds ? "Akan Datang" : today > de ? "Selesai" : "Berlangsung";
    const allUsers = getData("Users");
    const allSessions = getData(EXAM_SHEET.SESSIONS).filter(
      (s) => String(s.period_id) === String(period_id),
    );
    const allRooms = getData(EXAM_SHEET.ROOMS);
    const allSups = getData(EXAM_SHEET.SUPERVISORS);
    const allCom = getData(EXAM_SHEET.COMMITTEE).filter(
      (c) => String(c.period_id) === String(period_id),
    );
    const allBaps = getData(EXAM_SHEET.BAP);
    const sessionIds = new Set(allSessions.map((s) => String(s.id)));
    const sessionMap = {};
    allSessions.forEach((s) => {
      sessionMap[String(s.id)] = s;
    });
    const roomsInPeriod = allRooms.filter((r) =>
      sessionIds.has(String(r.session_id)),
    );
    const roomMap = {};
    roomsInPeriod.forEach((r) => {
      roomMap[String(r.id)] = r;
    });
    const roomIds = new Set(roomsInPeriod.map((r) => String(r.id)));
    const supsInPeriod = allSups.filter((sup) =>
      roomIds.has(String(sup.room_id)),
    );
    const jtmMap = {};
    const initGuru = (userId) => {
      if (!jtmMap[userId]) {
        const g = allUsers.find((u) => String(u.id) === userId);
        jtmMap[userId] = {
          user_id: userId,
          guru_name: g ? String(g.full_name) : "Tidak Dikenal",
          jtm_supervisor: 0,
          jtm_committee: 0,
          sessions_done: 0,
          sessions_confirmed: 0,
          bap_submitted: 0,
          duty_detail: [],
        };
      }
    };
    supsInPeriod.forEach((sup) => {
      if (String(sup.status) === "cancelled") return;
      const room = roomMap[String(sup.room_id)];
      const session = room ? sessionMap[String(room.session_id)] : null;
      if (!session) return;
      const userId = String(sup.user_id);
      initGuru(userId);
      const isConfirmed = !!(
        sup.confirmed_by && String(sup.confirmed_by).trim() !== ""
      );
      const bap = allBaps.find(
        (b) => String(b.supervisor_id) === String(sup.id),
      );
      const hasBap = !!(
        bap &&
        bap.submitted_at &&
        String(bap.submitted_at).trim() !== ""
      );
      let supJtm =
        Number(sup.jtm_val) > 0
          ? Number(sup.jtm_val)
          : Number(session.jtm_val) || 0;
      const jtmAdjustment = _jtmReadAdjustment_("EXAM", String(sup.id));
      if (jtmAdjustment) {
        const tz = Session.getScriptTimeZone();
        const sessionDateStr =
          session.date instanceof Date
            ? Utilities.formatDate(session.date, tz, "yyyy-MM-dd")
            : String(session.date);
        const adjDateStr =
          jtmAdjustment.date instanceof Date
            ? Utilities.formatDate(jtmAdjustment.date, tz, "yyyy-MM-dd")
            : String(jtmAdjustment.date);
        const storedAdjusted = jtmAdjustment.adjusted_jtm;
        const dateMatches = !adjDateStr || adjDateStr === sessionDateStr;
        if (
          dateMatches &&
          storedAdjusted !== null &&
          storedAdjusted !== undefined &&
          storedAdjusted !== "" &&
          !isNaN(Number(storedAdjusted))
        ) {
          supJtm = Number(storedAdjusted);
        }
      }
      if (isConfirmed) jtmMap[userId].sessions_confirmed++;
      if (hasBap) {
        jtmMap[userId].bap_submitted++;
        jtmMap[userId].sessions_done++;
        jtmMap[userId].jtm_supervisor += supJtm;
      }
      jtmMap[userId].duty_detail.push({
        type: "supervisor",
        date: String(session.date),
        session_name: String(session.session_name || ""),
        room_name: String(room.room_name || ""),
        is_substitute:
          sup.is_substitute === true || String(sup.is_substitute) === "true",
        is_confirmed: isConfirmed,
        has_bap: hasBap,
        jtm_val: hasBap ? supJtm : 0,
      });
    });
    allCom.forEach((c) => {
      if (String(c.status) === "cancelled") return;
      const userId = String(c.user_id);
      initGuru(userId);
      const isConfirmed = !!(
        c.confirmed_by && String(c.confirmed_by).trim() !== ""
      );
      if (isConfirmed) {
        jtmMap[userId].jtm_committee += Number(c.jtm_val) || 0;
      }
      jtmMap[userId].duty_detail.push({
        type: "committee",
        date: String(c.date),
        is_substitute:
          c.is_substitute === true || String(c.is_substitute) === "true",
        is_confirmed: isConfirmed,
        jtm_val: isConfirmed ? Number(c.jtm_val) || 0 : 0,
      });
    });
    const recapJtm = Object.values(jtmMap)
      .map((g) => ({
        ...g,
        total_jtm: g.jtm_supervisor + g.jtm_committee,
      }))
      .sort((a, b) => a.guru_name.localeCompare(b.guru_name));
    const recapBap = [];
    supsInPeriod.forEach((sup) => {
      const bap = allBaps.find(
        (b) => String(b.supervisor_id) === String(sup.id),
      );
      if (!bap || !bap.submitted_at || String(bap.submitted_at).trim() === "")
        return;
      const room = roomMap[String(sup.room_id)];
      const session = room ? sessionMap[String(room.session_id)] : null;
      const g = allUsers.find((u) => String(u.id) === String(sup.user_id));
      recapBap.push({
        supervisor_id: String(sup.id),
        user_id: String(sup.user_id),
        guru_name: g ? String(g.full_name) : "Tidak Dikenal",
        is_substitute:
          sup.is_substitute === true || String(sup.is_substitute) === "true",
        status: String(sup.status || ""),
        date: session ? String(session.date) : "",
        session_name: session ? String(session.session_name || "") : "",
        time_start: session ? _examFormatTime_(session.time_start) : "",
        time_end: session ? _examFormatTime_(session.time_end) : "",
        jtm_val:
          Number(sup.jtm_val) > 0
            ? Number(sup.jtm_val)
            : session
              ? Number(session.jtm_val) || 0
              : 0,
        room_name: room ? String(room.room_name || "") : "",
        subject: room ? String(room.subject || "") : "",
        class_name: room ? String(room.class_name || "") : "",
        peserta_hadir: Number(bap.peserta_hadir) || 0,
        peserta_absen: Number(bap.peserta_absen) || 0,
        catatan: String(bap.catatan || ""),
        submitted_at: String(bap.submitted_at || ""),
        updated_at: String(bap.updated_at || ""),
      });
    });
    recapBap.sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      return a.session_name.localeCompare(b.session_name);
    });
    return {
      status: "success",
      period: {
        id: String(period.id),
        name: String(period.name || ""),
        date_start: ds,
        date_end: de,
        jtm_committee_per_day: Number(period.jtm_committee_per_day) || 0,
        status: pStatus,
      },
      recap_jtm: recapJtm,
      recap_bap: recapBap,
    };
  } catch (e) {
    return { status: "error", message: e.message };
  }
}

function _cleanup(token, periodId) {
  if (!periodId) return;
  try {
    deleteExamPeriod(token, periodId, true);
    Logger.log("🧹 Cleanup: Data uji [" + periodId + "] berhasil dihapus.");
  } catch (e) {
    Logger.log(
      "⚠️  Cleanup gagal: " +
        e.message +
        " — Hapus manual via Admin > Jadwal Ujian.",
    );
  }
}
var JTM_SHEET = {
  ADJUSTMENTS: "JTM_Adjustments",
  REALLOCATIONS: "JTM_Reallocations",
};
var JTM_ADJUSTMENTS_HEADERS = [
  "id",
  "occurrence_type",
  "occurrence_ref",
  "date",
  "original_user_id",
  "scheduled_jtm",
  "adjusted_jtm",
  "jtm_difference",
  "reason",
  "adjusted_by",
  "adjusted_at",
];
var JTM_REALLOCATIONS_HEADERS = [
  "id",
  "adjustment_id",
  "occurrence_type",
  "occurrence_ref",
  "date",
  "substitute_user_id",
  "allocated_jtm",
  "created_by",
  "created_at",
  "log_id",
];
