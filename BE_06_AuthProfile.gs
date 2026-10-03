// SiM-Guru — BE_06_AuthProfile.gs
function loginUser(username, password) {
  const extPassword = _extGetPassword_(username);
  if (extPassword === null || String(extPassword) !== String(password)) {
    return { status: "error", message: "Username atau Password salah!" };
  }
  const users = getData("Users");
  const user = users.find(
    (u) => String(u.username).trim() === String(username).trim(),
  );
  if (!user) {
    return {
      status: "error",
      message: "Profil pengguna tidak ditemukan. Hubungi Administrator.",
    };
  }
  // --- Maintenance Mode Check ---
  const configData = getData("Config");
  const _cfgVal = (key) => { const item = configData.find((c) => c.key === key); return item ? item.value : ""; };
  let isMaintenance = String(_cfgVal("maintenance_mode")).toLowerCase() === "true";
  
  if (!isMaintenance && String(_cfgVal("maintenance_scheduled")).toLowerCase() === "true") {
    const startStr = _cfgVal("maintenance_start");
    const endStr = _cfgVal("maintenance_end");
    if (startStr && endStr) {
      const start = new Date(startStr);
      const end = new Date(endStr);
      const now = new Date();
      if (!isNaN(start.getTime()) && !isNaN(end.getTime()) && now >= start && now <= end) {
        isMaintenance = true;
      }
    }
  }

  if (isMaintenance) {
    if (String(user.role).toLowerCase() !== "admin") {
      const maintenanceMsg = _cfgVal("maintenance_message") || "";
      return {
        status: "maintenance",
        message: maintenanceMsg,
      };
    }
  }
  // --- End Maintenance Mode Check ---
  const token = "TKN-" + Utilities.getUuid();
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + 24);
  try {
    const sessionsSheet = getSheet("Sessions");
    const lastRow = sessionsSheet.getLastRow();
    if (lastRow > 1) {
      const range = sessionsSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      const now = new Date();
      for (let i = range.length - 1; i >= 0; i--) {
        const row = range[i];
        const expiryRaw = row[2];
        let expDate = null;
        if (expiryRaw instanceof Date) expDate = expiryRaw;
        else if (expiryRaw) {
          const tryD = new Date(expiryRaw);
          if (!isNaN(tryD.getTime())) expDate = tryD;
        }
        const isExpired = !expDate || expDate <= now;
        if (isExpired) {
          sessionsSheet.deleteRow(i + 2);
        }
      }
    }
  } catch (e) {}
  getSheet("Sessions").appendRow([token, user.id, expiry, new Date()]);
  const isDefaultPass = String(password) === "123456";
  const hasSecurityQuestion = !!(
    user.security_question && String(user.security_question).trim()
  );
  const hasEmailVerified = !!(
    user.email &&
    String(user.email).trim() &&
    _isValidEmail_(String(user.email))
  );
  return {
    status: "success",
    token,
    user: {
      id: user.id,
      name: user.full_name,
      role: user.role,
      is_default_pass: isDefaultPass,
      has_security_question: hasSecurityQuestion,
      has_email_verified: hasEmailVerified,
      email: String(user.email || ""),
    },
  };
}

function logoutUser(token) {
  const sheet = getSheet("Sessions");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { status: "success" };
}

function verifySession(token) {
  const sheet = getSheet("Sessions");
  const data = sheet.getDataRange().getValues();
  const session = data.find((row) => row[0] === token);
  if (!session) return null;
  const expiry = new Date(session[2]);
  if (new Date() > expiry) return null;
  const userId = session[1];
  const user = findData("Users", "id", userId);
  return user;
}

function checkSessionStatus(token) {
  try {
    if (!token) return { status: "expired" };
    const user = verifySession(token);
    return user ? { status: "ok" } : { status: "expired" };
  } catch (e) {
    console.error("checkSessionStatus error: " + e);
    return { status: "expired" };
  }
}

function updateUserProfile(token, payload) {
  const user = verifySession(token);
  if (!user)
    return { status: "error", message: "Sesi habis, silakan login kembali" };
  const sheetData = getData("Users");
  const userRowIndex = sheetData.findIndex(
    (u) => String(u.id) === String(user.id),
  );
  if (userRowIndex === -1)
    return { status: "error", message: "User tidak ditemukan" };
  const sheetObj = getSheet("Users");
  const sheetRowNumber = userRowIndex + 2;
  if (payload.full_name && !payload.phone && payload.phone !== "") {
    sheetObj.getRange(sheetRowNumber, 5).setValue(payload.full_name);
    return { status: "success", message: "Nama berhasil diubah" };
  }
  if (typeof payload.phone === "string") {
    const cleaned = _normalizePhoneNumber_(payload.phone);
    if (payload.phone.trim() !== "" && !cleaned.valid) {
      return {
        status: "error",
        message: cleaned.message || "Format nomor HP tidak valid.",
      };
    }
    const phoneCol = _ensureUserColumn_(sheetObj, "phone", 11);
    sheetObj.getRange(sheetRowNumber, phoneCol).setValue(cleaned.normalized);
    if (payload.full_name)
      sheetObj.getRange(sheetRowNumber, 5).setValue(payload.full_name);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Data berhasil disimpan",
      phone: cleaned.normalized,
    };
  }
  if (payload.new_password && payload.old_password) {
    const storedUsername = sheetData[userRowIndex].username;
    const currentExtPass = _extGetPassword_(storedUsername);
    if (
      currentExtPass === null ||
      String(currentExtPass) !== String(payload.old_password)
    ) {
      return {
        status: "error",
        message: "Password lama yang Anda masukkan salah.",
      };
    }
    if (String(payload.new_password) === String(payload.old_password)) {
      return {
        status: "error",
        message: "Password baru tidak boleh sama dengan password lama.",
      };
    }
    const complexityCheck = _checkPasswordComplexity(
      payload.new_password,
      user.full_name,
      storedUsername,
    );
    if (!complexityCheck.valid) {
      return { status: "error", message: complexityCheck.message };
    }
    const extOk = _extSetPassword_(storedUsername, payload.new_password);
    if (!extOk) {
      return {
        status: "error",
        message: "Gagal memperbarui password. Hubungi Administrator.",
      };
    }
    sheetObj.getRange(sheetRowNumber, 3).setValue(payload.new_password);
    return { status: "success", message: "Password berhasil diubah" };
  }
  return { status: "error", message: "Data tidak lengkap" };
}

function updateGuruProfileDetails(token, payload) {
  const user = verifySession(token);
  if (!user)
    return { status: "error", message: "Sesi habis, silakan login kembali" };
  const sheetData = getData("Users");
  const userRowIndex = sheetData.findIndex(
    (u) => String(u.id) === String(user.id),
  );
  if (userRowIndex === -1)
    return { status: "error", message: "User tidak ditemukan" };
  const sheetObj = getSheet("Users");
  const sheetRowNumber = userRowIndex + 2;
  sheetObj
    .getRange(sheetRowNumber, 15)
    .setValue(payload.status_kepegawaian || "");
  sheetObj.getRange(sheetRowNumber, 16).setValue(payload.golongan || "");
  sheetObj.getRange(sheetRowNumber, 17).setValue(payload.alamat || "");
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  return { status: "success", message: "Profil guru berhasil disimpan" };
}

function getMyProfile(token) {
  const user = verifySession(token);
  if (!user)
    return { status: "error", message: "Sesi habis, silakan login kembali" };
  const sheetData = getData("Users");
  const me = sheetData.find((u) => String(u.id) === String(user.id));
  if (!me) return { status: "error", message: "User tidak ditemukan" };
  const configRaw = getData("Config");
  let config = {};
  (configRaw || []).forEach((c) => {
    config[c.key] = c.value;
  });
  const activeTP = config.tahun_pelajaran || "";
  const activeSem = config.semester || "";
  const allowancesData = getData("Allowances") || [];
  const dutyNames = allowancesData
    .filter((a) => String(a.user_id).trim() === String(user.id).trim())
    .map((a) => a.duty_name)
    .filter(Boolean);
  const jabatan = [...new Set(dutyNames)];
  const schedulesData = getData("Schedules") || [];
  const mySubjects = schedulesData
    .filter((s) => {
      const sTP = s.tahun_pelajaran || activeTP;
      const sSem = s.semester || activeSem;
      return (
        String(s.user_id).trim() === String(user.id).trim() &&
        sTP === activeTP &&
        sSem === activeSem
      );
    })
    .map((s) => `${s.subject || ""} (${s.class_name || ""})`.trim());
  const mapel_diampu = [...new Set(mySubjects)].filter(Boolean);
  const userRowIndex = sheetData.findIndex(
    (u) => String(u.id) === String(user.id),
  );
  const sheetObj = getSheet("Users");
  const userRowNumber = userRowIndex + 2;
  let profileDetails = ["", "", ""];
  if (userRowIndex !== -1) {
    try {
      profileDetails = sheetObj
        .getRange(userRowNumber, 15, 1, 3)
        .getValues()[0];
    } catch (e) {}
  }
  return {
    status: "success",
    profile: {
      id: String(me.id || ""),
      username: String(me.username || ""),
      full_name: String(me.full_name || ""),
      role: String(me.role || ""),
      nip: String(me.nip || ""),
      phone: String(me.phone || ""),
      email: String(me.email || ""),
      km_distance: String(me.km_distance || ""),
      jabatan: jabatan,
      mapel_diampu: mapel_diampu,
      status_kepegawaian: String(profileDetails[0] || ""),
      golongan: String(profileDetails[1] || ""),
      alamat: String(profileDetails[2] || ""),
    },
  };
}

function _ensureUserColumn_(sheetObj, headerName, defaultPos) {
  const lastCol = sheetObj.getLastColumn();
  const headers = sheetObj.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (
      String(headers[i]).trim().toLowerCase() ===
      String(headerName).toLowerCase()
    ) {
      return i + 1;
    }
  }
  const newCol = lastCol + 1;
  sheetObj.getRange(1, newCol).setValue(headerName);
  return newCol;
}

function _normalizePhoneNumber_(input) {
  let s = String(input || "").trim();
  if (!s) return { valid: true, normalized: "" };
  s = s.replace(/[\s\-\(\)\.]/g, "");
  if (s.indexOf("+") === 0) s = s.substring(1);
  if (!/^\d+$/.test(s)) {
    return {
      valid: false,
      normalized: "",
      message: "Nomor HP hanya boleh mengandung angka.",
    };
  }
  if (s.charAt(0) === "0") {
    s = "62" + s.substring(1);
  } else if (s.charAt(0) !== "6" || s.charAt(1) !== "2") {
    if (s.length >= 8 && s.length <= 13) {
      s = "62" + s;
    }
  }
  if (s.length < 10 || s.length > 15) {
    return {
      valid: false,
      normalized: "",
      message: "Nomor HP harus 10–15 digit (termasuk kode negara).",
    };
  }
  if (s.indexOf("62") !== 0) {
    return {
      valid: false,
      normalized: "",
      message: "Nomor HP harus menggunakan kode negara Indonesia (62).",
    };
  }
  return { valid: true, normalized: s };
}
