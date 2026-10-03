// SiM-Guru — BE_07_Notifications.gs
function requestEmailOTP(token, newEmail) {
  try {
    const user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali" };
    const email = String(newEmail || "")
      .trim()
      .toLowerCase();
    if (!email) return { status: "error", message: "Email wajib diisi." };
    if (!_isValidEmail_(email)) {
      return { status: "error", message: "Format email tidak valid." };
    }
    const sheetData = getData("Users");
    const otherWithEmail = sheetData.find(
      (u) =>
        String(u.email || "").toLowerCase() === email &&
        String(u.id) !== String(user.id),
    );
    if (otherWithEmail) {
      return {
        status: "error",
        message: "Email ini sudah digunakan oleh akun lain.",
      };
    }
    const props = PropertiesService.getScriptProperties();
    const rateKey = "otp_rate_" + user.id;
    const lastReq = Number(props.getProperty(rateKey) || 0);
    const now = Date.now();
    const SECS = 60;
    if (lastReq && now - lastReq < SECS * 1000) {
      const wait = Math.ceil((SECS * 1000 - (now - lastReq)) / 1000);
      return {
        status: "error",
        message: "Mohon tunggu " + wait + " detik sebelum meminta kode lagi.",
      };
    }
    const code = String(Math.floor(100000 + Math.random() * 900000));
    const dataKey = "otp_data_" + user.id;
    const payload = {
      code: code,
      email: email,
      expires: now + 10 * 60 * 1000,
      attempts: 0,
    };
    props.setProperty(dataKey, JSON.stringify(payload));
    props.setProperty(rateKey, String(now));
    try {
      const subject = "Kode Verifikasi Email — SiM-Guru";
      const htmlBody =
        '<div style="font-family:Inter,Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px;color:#0f172a;">' +
        '<div style="background:linear-gradient(135deg,#4338CA,#4F46E5);color:white;padding:20px 24px;border-radius:14px 14px 0 0;">' +
        '<div style="font-size:13px;opacity:0.85;letter-spacing:0.05em;">SiM-GURU</div>' +
        '<div style="font-size:20px;font-weight:800;margin-top:4px;">Verifikasi Alamat Email</div>' +
        "</div>" +
        '<div style="border:1px solid #e2e8f0;border-top:0;border-radius:0 0 14px 14px;padding:24px;background:#fff;">' +
        '<p style="font-size:14px;color:#334155;margin:0 0 16px;">Halo <strong>' +
        _escHtml_(user.full_name || "Pengguna") +
        "</strong>,</p>" +
        '<p style="font-size:14px;color:#334155;margin:0 0 16px;">Anda telah meminta untuk mengubah alamat email akun SiM-Guru. Gunakan kode verifikasi berikut:</p>' +
        '<div style="background:#EEF2FF;border:1px dashed #6366F1;border-radius:12px;padding:18px;text-align:center;margin:16px 0;">' +
        "<div style=\"font-family:'JetBrains Mono',monospace;font-size:32px;font-weight:800;letter-spacing:0.4em;color:#4338CA;\">" +
        code +
        "</div>" +
        "</div>" +
        '<p style="font-size:13px;color:#64748b;margin:16px 0 8px;">⏱️ Kode berlaku <strong>10 menit</strong>.</p>' +
        '<p style="font-size:13px;color:#64748b;margin:0 0 16px;">Jika Anda tidak meminta perubahan ini, abaikan email ini.</p>' +
        '<hr style="border:0;border-top:1px solid #e2e8f0;margin:20px 0;">' +
        '<p style="font-size:11px;color:#94a3b8;margin:0;">Email otomatis dari SiM-Guru. Jangan balas email ini.</p>' +
        "</div>" +
        "</div>";
      const textBody =
        "Kode verifikasi email SiM-Guru Anda: " +
        code +
        "\n\nKode berlaku 10 menit. Jika Anda tidak meminta perubahan ini, abaikan email ini.";
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: textBody,
        htmlBody: htmlBody,
        name: "SiM-Guru",
      });
    } catch (mailErr) {
      props.deleteProperty(dataKey);
      props.deleteProperty(rateKey);
      return {
        status: "error",
        message:
          "Gagal mengirim email verifikasi: " +
          (mailErr && mailErr.message ? mailErr.message : mailErr),
      };
    }
    return {
      status: "success",
      message:
        "Kode OTP telah dikirim ke " + email + ". Kode berlaku 10 menit.",
      expires_in: 600,
      cooldown: SECS,
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}

function verifyEmailOTP(token, newEmail, otpCode) {
  try {
    const user = verifySession(token);
    if (!user)
      return { status: "error", message: "Sesi habis, silakan login kembali" };
    const email = String(newEmail || "")
      .trim()
      .toLowerCase();
    const code = String(otpCode || "").trim();
    if (!_isValidEmail_(email))
      return { status: "error", message: "Format email tidak valid." };
    if (!/^\d{6}$/.test(code))
      return { status: "error", message: "Kode OTP harus 6 digit angka." };
    const props = PropertiesService.getScriptProperties();
    const dataKey = "otp_data_" + user.id;
    const raw = props.getProperty(dataKey);
    if (!raw)
      return {
        status: "error",
        message: "Tidak ada permintaan OTP aktif. Minta kode baru.",
      };
    let data;
    try {
      data = JSON.parse(raw);
    } catch (e) {
      props.deleteProperty(dataKey);
      return { status: "error", message: "Data OTP rusak. Minta kode baru." };
    }
    if (Date.now() > Number(data.expires || 0)) {
      props.deleteProperty(dataKey);
      return {
        status: "error",
        message: "Kode OTP sudah kedaluwarsa. Minta kode baru.",
      };
    }
    if (String(data.email).toLowerCase() !== email) {
      return {
        status: "error",
        message:
          "Email yang Anda masukkan berbeda dengan permintaan OTP terakhir. Minta kode baru.",
      };
    }
    if (Number(data.attempts || 0) >= 5) {
      props.deleteProperty(dataKey);
      return {
        status: "error",
        message: "Terlalu banyak percobaan salah. Minta kode baru.",
      };
    }
    if (String(data.code) !== code) {
      data.attempts = Number(data.attempts || 0) + 1;
      props.setProperty(dataKey, JSON.stringify(data));
      const remain = 5 - data.attempts;
      return {
        status: "error",
        message:
          "Kode OTP salah. " +
          (remain > 0 ? "Sisa percobaan: " + remain : "Percobaan habis."),
      };
    }
    const sheetData = getData("Users");
    const userRowIndex = sheetData.findIndex(
      (u) => String(u.id) === String(user.id),
    );
    if (userRowIndex === -1)
      return { status: "error", message: "User tidak ditemukan" };
    const otherWithEmail = sheetData.find(
      (u) =>
        String(u.email || "").toLowerCase() === email &&
        String(u.id) !== String(user.id),
    );
    if (otherWithEmail) {
      props.deleteProperty(dataKey);
      return {
        status: "error",
        message: "Email ini sudah digunakan oleh akun lain.",
      };
    }
    const sheetObj = getSheet("Users");
    const emailCol = _ensureUserColumn_(sheetObj, "email", 12);
    sheetObj.getRange(userRowIndex + 2, emailCol).setValue(email);
    props.deleteProperty(dataKey);
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return {
      status: "success",
      message: "Email berhasil diverifikasi & disimpan.",
      email: email,
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}

function _isValidEmail_(s) {
  if (!s) return false;
  return /^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$/.test(String(s));
}

function _escHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function _notifGetEmail_(userId) {
  if (!userId) return null;
  try {
    var users = getData("Users");
    var u = users.find(function (x) {
      return String(x.id) === String(userId);
    });
    if (!u) return null;
    var email = String(u.email || "")
      .trim()
      .toLowerCase();
    if (!email || !_isValidEmail_(email)) return null;
    return email;
  } catch (_) {
    return null;
  }
}

function _notifGetName_(userId) {
  try {
    var u = getData("Users").find(function (x) {
      return String(x.id) === String(userId);
    });
    return u ? String(u.full_name || u.username || "Pengguna") : "Pengguna";
  } catch (_) {
    return "Pengguna";
  }
}

function _notifBuildHtml_(opts) {
  opts = opts || {};
  var color = opts.accent || "#4F46E5";
  var color2 = opts.accent2 || "#4338CA";
  var schoolName = "";
  try {
    var cfg = _getConfigMap();
    schoolName = String(cfg["school_name"] || cfg["app_name"] || "");
  } catch (_) {}
  var badgesHtml = "";
  if (opts.badges && opts.badges.length) {
    badgesHtml =
      '<div style="display:flex;flex-wrap:wrap;gap:6px;margin:12px 0 4px;">' +
      opts.badges
        .map(function (b) {
          return (
            '<span style="display:inline-block;background:#EEF2FF;color:#4338CA;border:1px solid #C7D2FE;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:600;">' +
            _escHtml_(b) +
            "</span>"
          );
        })
        .join("") +
      "</div>";
  }
  var parasHtml = "";
  if (opts.paragraphs && opts.paragraphs.length) {
    parasHtml = opts.paragraphs
      .map(function (p) {
        return (
          '<p style="font-size:14px;color:#334155;line-height:1.55;margin:10px 0;">' +
          p +
          "</p>"
        );
      })
      .join("");
  }
  var ctaHtml = "";
  if (opts.ctaText && opts.ctaUrl) {
    ctaHtml =
      '<div style="text-align:center;margin:18px 0 4px;">' +
      '<a href="' +
      _escHtml_(opts.ctaUrl) +
      '" style="display:inline-block;background:linear-gradient(135deg,' +
      color2 +
      "," +
      color +
      ');color:white;text-decoration:none;padding:10px 20px;border-radius:10px;font-weight:700;font-size:13px;">' +
      _escHtml_(opts.ctaText) +
      "</a></div>";
  }
  return (
    "" +
    '<div style="font-family:Inter,Arial,sans-serif;max-width:560px;margin:0 auto;padding:0;color:#0f172a;background:#F8FAFC;">' +
    '<div style="padding:24px 16px;">' +
    '<div style="background:linear-gradient(135deg,' +
    color2 +
    "," +
    color +
    ');color:white;padding:20px 24px;border-radius:14px 14px 0 0;">' +
    '<div style="font-size:11px;opacity:0.85;letter-spacing:0.08em;font-weight:600;text-transform:uppercase;">SiM-GURU' +
    (schoolName ? " · " + _escHtml_(schoolName) : "") +
    "</div>" +
    '<div style="font-size:18px;font-weight:800;margin-top:6px;line-height:1.25;">' +
    _escHtml_(opts.title || "Notifikasi") +
    "</div>" +
    "</div>" +
    '<div style="border:1px solid #E2E8F0;border-top:0;border-radius:0 0 14px 14px;padding:22px 24px;background:#fff;">' +
    (opts.name
      ? '<p style="font-size:14px;color:#334155;margin:0 0 14px;">Halo <strong>' +
        _escHtml_(opts.name) +
        "</strong>,</p>"
      : "") +
    (opts.intro
      ? '<p style="font-size:14px;color:#334155;line-height:1.55;margin:0 0 12px;">' +
        opts.intro +
        "</p>"
      : "") +
    badgesHtml +
    parasHtml +
    ctaHtml +
    '<hr style="border:0;border-top:1px solid #E2E8F0;margin:18px 0 12px;">' +
    '<p style="font-size:11px;color:#94A3B8;margin:0;">Email otomatis dari SiM-Guru. Mohon tidak membalas email ini. Jika ada pertanyaan, hubungi admin sekolah.</p>' +
    "</div>" +
    "</div>" +
    "</div>"
  );
}

function _notifSend_(toEmail, subject, htmlBody, textBody) {
  if (!toEmail || !_isValidEmail_(toEmail)) return false;
  try {
    MailApp.sendEmail({
      to: toEmail,
      subject: subject,
      body:
        textBody || subject + "\n\n" + "Buka aplikasi SiM-Guru untuk detail.",
      htmlBody: htmlBody,
      name: "SiM-Guru",
    });
    return true;
  } catch (e) {
    try {
      console.warn("[notif] gagal kirim ke " + toEmail + ": " + e);
    } catch (_) {}
    return false;
  }
}

function _notifToUser_(userId, subject, htmlBody, textBody) {
  var email = _notifGetEmail_(userId);
  if (!email) return false;
  return _notifSend_(email, subject, htmlBody, textBody);
}

function _notifIsHoliday_(dateStr) {
  try {
    var ds =
      dateStr || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    var holidays = getData("Academic_Calendar") || [];
    return holidays.some(function (h) {
      try {
        var hStr = Utilities.formatDate(
          new Date(h.date),
          "Asia/Jakarta",
          "yyyy-MM-dd",
        );
        return hStr === ds;
      } catch (_) {
        return false;
      }
    });
  } catch (_) {
    return false;
  }
}

function _notifIsKbmEnabled_(dateStr) {
  try {
    var ds =
      dateStr || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    if (_notifIsHoliday_(ds)) return false;
    return true;
  } catch (e) {
    return true;
  }
}

function _notifIsTeachingDay_(dateStr) {
  try {
    var ds =
      dateStr || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    if (_notifIsHoliday_(ds)) return false;
    if (_notifIsExamEnabled_(ds)) return false;
    return true;
  } catch (e) {
    return true;
  }
}

function _notifIsExamEnabled_(dateStr) {
  try {
    var ds =
      dateStr || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    if (_notifIsHoliday_(ds)) return false;
    var periods = getData(EXAM_SHEET.PERIODS) || [];
    return periods.some(function (p) {
      var ps = String(p.date_start || "");
      var pe = String(p.date_end || "");
      return ps && pe && ds >= ps && ds <= pe;
    });
  } catch (_) {
    return false;
  }
}

function _notifPicketConfirmed_(userId, opts) {
  if (!_notifIsKbmEnabled_()) return;
  opts = opts || {};
  var byRole = String(opts.byRole || "Admin");
  var byName = String(opts.byName || "").trim();
  var byLabel = byRole + (byName ? " (" + byName + ")" : "");
  var name = _notifGetName_(userId);
  var dateStr = _notifTodayID_();
  var html = _notifBuildHtml_({
    title: "✅ Kehadiran Piket Anda Dikonfirmasi",
    accent: "#3B82F6",
    accent2: "#1D4ED8",
    name: name,
    intro:
      "Kehadiran Anda sebagai <strong>Guru Piket</strong> hari ini telah dikonfirmasi oleh <strong>" +
      _escHtml_(byLabel) +
      "</strong>.",
    badges: ["Tugas: Guru Piket", "Tanggal: " + dateStr],
    paragraphs: [
      "JTM piket sudah masuk ke sistem honorarium Anda. Tetap pantau dashboard untuk update lainnya.",
    ],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Piket Dikonfirmasi · " + dateStr,
    html,
  );
}

function _notifPicketRevoked_(userId) {
  if (!_notifIsKbmEnabled_()) return;
  var name = _notifGetName_(userId);
  var dateStr = _notifTodayID_();
  var html = _notifBuildHtml_({
    title: "⚠️ Konfirmasi Piket Dibatalkan",
    accent: "#F59E0B",
    accent2: "#B45309",
    name: name,
    intro:
      "Konfirmasi kehadiran Anda sebagai <strong>Guru Piket</strong> hari ini telah <strong>dibatalkan</strong> oleh admin.",
    badges: ["Tugas: Guru Piket", "Tanggal: " + dateStr],
    paragraphs: [
      "JTM piket Anda telah dihapus dari sistem honorarium. Hubungi admin jika ada pertanyaan.",
    ],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Konfirmasi Piket Dibatalkan · " + dateStr,
    html,
  );
}

function _notifPicketSubstituted_(originalUserId, substituteUserId) {
  if (!_notifIsKbmEnabled_()) return;
  var dateStr = _notifTodayID_();
  var subName = _notifGetName_(substituteUserId);
  var origName = _notifGetName_(originalUserId);
  var htmlSub = _notifBuildHtml_({
    title: "📌 Anda Ditugaskan sebagai Guru Piket Pengganti",
    accent: "#8B5CF6",
    accent2: "#6D28D9",
    name: subName,
    intro:
      "Admin menugaskan Anda untuk menggantikan <strong>" +
      _escHtml_(origName) +
      "</strong> sebagai Guru Piket hari ini.",
    badges: ["Pengganti dari: " + origName, "Tanggal: " + dateStr],
    paragraphs: [
      "Mohon segera hadir di pos piket. JTM akan dihitung setelah kehadiran Anda dikonfirmasi.",
    ],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Piket Pengganti · " + dateStr,
    htmlSub,
  );
  var htmlOrig = _notifBuildHtml_({
    title: "ℹ️ Tugas Piket Digantikan",
    accent: "#64748B",
    accent2: "#334155",
    name: origName,
    intro:
      "Tugas piket Anda hari ini telah dialihkan ke <strong>" +
      _escHtml_(subName) +
      "</strong> oleh admin.",
    badges: ["Diganti oleh: " + subName, "Tanggal: " + dateStr],
  });
  _notifToUser_(
    originalUserId,
    "[SiM-Guru] Tugas Piket Digantikan · " + dateStr,
    htmlOrig,
  );
}

function _notifPicketSubCancelled_(substituteUserId) {
  if (!_notifIsKbmEnabled_()) return;
  var name = _notifGetName_(substituteUserId);
  var dateStr = _notifTodayID_();
  var html = _notifBuildHtml_({
    title: "🚫 Tugas Piket Pengganti Dibatalkan",
    accent: "#F97316",
    accent2: "#C2410C",
    name: name,
    intro:
      "Tugas Anda sebagai <strong>Guru Piket Pengganti</strong> hari ini telah <strong>dibatalkan</strong> oleh admin.",
    badges: ["Tanggal: " + dateStr],
    paragraphs: [
      "JTM yang sebelumnya tercatat sudah dihapus. Tidak ada tindakan yang perlu Anda lakukan.",
    ],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengganti Piket Dibatalkan · " + dateStr,
    html,
  );
}

function _notifKbmConfirmed_(userId, opts) {
  if (!_notifIsTeachingDay_()) return;
  opts = opts || {};
  var name = _notifGetName_(userId);
  var dateStr = _notifTodayID_();
  var byName = opts.byName || "Guru Piket / Admin";
  var html = _notifBuildHtml_({
    title: "✅ Kehadiran Mengajar Dikonfirmasi",
    accent: "#10B981",
    accent2: "#047857",
    name: name,
    intro:
      "Kehadiran Anda sebagai Guru Mata Pelajaran hari ini telah dikonfirmasi oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Mengajar Dikonfirmasi · " + dateStr,
    html,
  );
}

function _notifKbmRevoked_(userId, opts) {
  if (!_notifIsTeachingDay_()) return;
  opts = opts || {};
  var name = _notifGetName_(userId);
  var dateStr = _notifTodayID_();
  var byName = opts.byName || "Guru Piket / Admin";
  var html = _notifBuildHtml_({
    title: "⚠️ Konfirmasi Kehadiran Mengajar Dibatalkan",
    accent: "#F59E0B",
    accent2: "#B45309",
    name: name,
    intro:
      "Konfirmasi kehadiran mengajar Anda hari ini telah <strong>dibatalkan</strong> oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr],
    paragraphs: ["Mohon segera tindak lanjuti."],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Mengajar Dibatalkan · " + dateStr,
    html,
  );
}

function _notifKbmSubstituted_(
  originalUserId,
  substituteUserId,
  scheduleInfo,
  byName,
) {
  if (!_notifIsTeachingDay_()) return;
  var dateStr = _notifTodayID_();
  var origName = _notifGetName_(originalUserId);
  var subName = _notifGetName_(substituteUserId);
  byName = byName || "Guru Piket / Admin";
  var info = scheduleInfo ? " · " + scheduleInfo : "";
  var htmlSub = _notifBuildHtml_({
    title: "📌 Anda Ditugaskan sebagai Guru Pengganti KBM",
    accent: "#8B5CF6",
    accent2: "#6D28D9",
    name: subName,
    intro:
      "<strong>" +
      _escHtml_(byName) +
      "</strong> menugaskan Anda untuk menggantikan <strong>" +
      _escHtml_(origName) +
      "</strong> dalam jadwal mengajar.",
    badges: ["Pengganti dari: " + origName, "Tanggal: " + dateStr + info],
    paragraphs: ["JTM akan dihitung setelah Anda mengisi jurnal mengajar."],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengganti KBM · " + dateStr,
    htmlSub,
  );
  var htmlOrig = _notifBuildHtml_({
    title: "ℹ️ Jadwal Mengajar Digantikan",
    accent: "#64748B",
    accent2: "#334155",
    name: origName,
    intro:
      "Salah satu jadwal mengajar Anda hari ini dialihkan ke <strong>" +
      _escHtml_(subName) +
      "</strong> oleh " +
      _escHtml_(byName) +
      ".",
    badges: ["Diganti oleh: " + subName, "Tanggal: " + dateStr + info],
  });
  _notifToUser_(
    originalUserId,
    "[SiM-Guru] Jadwal Mengajar Digantikan · " + dateStr,
    htmlOrig,
  );
}

function _notifKbmSubCancelled_(substituteUserId, scheduleInfo, byName) {
  if (!_notifIsTeachingDay_()) return;
  var name = _notifGetName_(substituteUserId);
  var dateStr = _notifTodayID_();
  byName = byName || "Guru Piket / Admin";
  var info = scheduleInfo ? " · " + scheduleInfo : "";
  var html = _notifBuildHtml_({
    title: "🚫 Tugas Pengganti KBM Dibatalkan",
    accent: "#F97316",
    accent2: "#C2410C",
    name: name,
    intro:
      "Tugas Anda sebagai pengganti KBM telah <strong>dibatalkan</strong> oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr + info],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengganti KBM Dibatalkan · " + dateStr,
    html,
  );
}

function _notifAutoJurnalLiburGenerated_(
  generatedLogs,
  dateStr,
  holidayDesc,
  actorName,
) {
  var summary = { sent: 0, failed: 0, recipients: 0 };
  if (!Array.isArray(generatedLogs) || generatedLogs.length === 0)
    return summary;
  var byUser = {};
  generatedLogs.forEach(function (l) {
    var uid = String(l.user_id || "").trim();
    if (!uid) return;
    if (!byUser[uid]) byUser[uid] = [];
    byUser[uid].push(l);
  });
  var displayDate = _notifFmtDateLong_(dateStr);
  var holidayLabel = String(holidayDesc || "Hari Libur").trim();
  var actor = String(actorName || "Sistem").trim();
  Object.keys(byUser).forEach(function (uid) {
    var items = byUser[uid];
    if (!items || !items.length) return;
    var name = _notifGetName_(uid) || "Bapak/Ibu Guru";
    var teachingItems = [];
    var picketItems = [];
    var ceremonyItems = [];
    items.forEach(function (it) {
      var sid = String(it.schedule_id || "");
      if (sid === "PICKET-DUTY") picketItems.push(it);
      else if (sid === "CEREMONY-DUTY") ceremonyItems.push(it);
      else teachingItems.push(it);
    });
    var totalJtm = items.reduce(function (s, it) {
      return s + (Number(it.jtm_val) || 0);
    }, 0);
    var paragraphs = [];
    if (teachingItems.length) {
      var rowsHtml = teachingItems
        .map(function (it) {
          return (
            '<li style="margin:4px 0;color:#334155;font-size:13px;">' +
            "<strong>" +
            _escHtml_(it.subject || "-") +
            "</strong>" +
            " · Kelas <strong>" +
            _escHtml_(it.class_name || "-") +
            "</strong>" +
            ' · <span style="color:#6366F1;font-weight:700;">' +
            (Number(it.jtm_val) || 0) +
            " JTM</span>" +
            "</li>"
          );
        })
        .join("");
      paragraphs.push(
        "📚 <strong>Jurnal Mengajar Otomatis (" +
          teachingItems.length +
          " jadwal):</strong>" +
          '<ul style="margin:8px 0 4px 18px;padding:0;">' +
          rowsHtml +
          "</ul>",
      );
    }
    if (picketItems.length) {
      paragraphs.push(
        "🛡️ <strong>Tugas Piket:</strong> Kehadiran Anda sebagai <strong>Guru Piket</strong> hari ini " +
          "telah dikonfirmasi otomatis oleh sistem " +
          '(<span style="color:#10B981;font-weight:700;">+' +
          4 * picketItems.length +
          " JTM</span>).",
      );
    }
    if (ceremonyItems.length) {
      paragraphs.push(
        "🎌 <strong>Tugas Pembina Upacara:</strong> Anda terdaftar sebagai Pembina Upacara dan " +
          "kehadirannya telah dikonfirmasi otomatis " +
          '(<span style="color:#F59E0B;font-weight:700;">+' +
          5 * ceremonyItems.length +
          " JTM</span>).",
      );
    }
    paragraphs.push(
      "✨ Semua entri di atas sudah masuk ke <strong>jurnal mengajar</strong> dan akan terhitung pada " +
        "<strong>honorarium bulan ini</strong>. Anda tidak perlu mengisi jurnal manual untuk tanggal ini.",
    );
    paragraphs.push(
      '<span style="color:#64748B;font-size:12px;">Dibuat oleh: ' +
        _escHtml_(actor) +
        "</span>",
    );
    var html = _notifBuildHtml_({
      title: "🎉 Bonus JTM Libur — Jurnal Otomatis",
      accent: "#7C3AED",
      accent2: "#5B21B6",
      name: name,
      intro:
        'Hari ini adalah <strong>hari libur bonus</strong> ("' +
        _escHtml_(holidayLabel) +
        '"). ' +
        "Sistem telah <strong>otomatis membuat jurnal</strong> untuk Anda sehingga tidak perlu input manual.",
      badges: ["Tanggal: " + displayDate, "Total JTM Bonus: " + totalJtm],
      paragraphs: paragraphs,
    });
    var textBody =
      "Halo " +
      name +
      ",\n\n" +
      'Hari ini hari libur bonus ("' +
      holidayLabel +
      '"). Sistem otomatis membuat jurnal untuk Anda.\n\n' +
      "Tanggal: " +
      displayDate +
      "\n" +
      "Total JTM Bonus: " +
      totalJtm +
      "\n" +
      (teachingItems.length
        ? "Jadwal mengajar otomatis: " + teachingItems.length + "\n"
        : "") +
      (picketItems.length
        ? "Tugas Piket: dikonfirmasi otomatis (+" +
          4 * picketItems.length +
          " JTM)\n"
        : "") +
      (ceremonyItems.length
        ? "Pembina Upacara: dikonfirmasi otomatis (+" +
          5 * ceremonyItems.length +
          " JTM)\n"
        : "") +
      "\nDibuat oleh: " +
      actor +
      "\n\nSemua entri ini sudah masuk jurnal dan akan terhitung pada honorarium bulan ini.";
    summary.recipients++;
    var ok = _notifToUser_(
      uid,
      "[SiM-Guru] 🎉 Bonus JTM Libur · " + displayDate,
      html,
      textBody,
    );
    if (ok) summary.sent++;
    else summary.failed++;
  });
  return summary;
}

function _notifCommitteeConfirmed_(userId, dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  var name = _notifGetName_(userId);
  dateStr = _notifFmtDateLong_(dateStr);
  var html = _notifBuildHtml_({
    title: "✅ Kehadiran Panitia Ujian Dikonfirmasi",
    accent: "#0891B2",
    accent2: "#0369A1",
    name: name,
    intro:
      "Kehadiran Anda sebagai <strong>Panitia Ujian</strong> telah dikonfirmasi oleh admin.",
    badges: ["Tanggal: " + dateStr],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Panitia Ujian Dikonfirmasi · " + dateStr,
    html,
  );
}

function _notifCommitteeRevoked_(userId, dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  var name = _notifGetName_(userId);
  dateStr = _notifFmtDateLong_(dateStr);
  var html = _notifBuildHtml_({
    title: "⚠️ Konfirmasi Panitia Ujian Dibatalkan",
    accent: "#F59E0B",
    accent2: "#B45309",
    name: name,
    intro:
      "Konfirmasi kehadiran Anda sebagai Panitia Ujian telah <strong>dibatalkan</strong> oleh admin.",
    badges: ["Tanggal: " + dateStr],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Konfirmasi Panitia Ujian Dibatalkan · " + dateStr,
    html,
  );
}

function _notifCommitteeSubstituted_(
  originalUserId,
  substituteUserId,
  dateStr,
) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  dateStr = _notifFmtDateLong_(dateStr);
  var origName = _notifGetName_(originalUserId);
  var subName = _notifGetName_(substituteUserId);
  var htmlSub = _notifBuildHtml_({
    title: "📌 Anda Ditugaskan sebagai Panitia Ujian Pengganti",
    accent: "#8B5CF6",
    accent2: "#6D28D9",
    name: subName,
    intro:
      "Admin menugaskan Anda menggantikan <strong>" +
      _escHtml_(origName) +
      "</strong> sebagai Panitia Ujian.",
    badges: ["Pengganti dari: " + origName, "Tanggal: " + dateStr],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Panitia Ujian Pengganti · " + dateStr,
    htmlSub,
  );
  var htmlOrig = _notifBuildHtml_({
    title: "ℹ️ Tugas Panitia Ujian Digantikan",
    accent: "#64748B",
    accent2: "#334155",
    name: origName,
    intro:
      "Tugas Panitia Ujian Anda telah dialihkan ke <strong>" +
      _escHtml_(subName) +
      "</strong> oleh admin.",
    badges: ["Diganti oleh: " + subName, "Tanggal: " + dateStr],
  });
  _notifToUser_(
    originalUserId,
    "[SiM-Guru] Tugas Panitia Ujian Digantikan · " + dateStr,
    htmlOrig,
  );
}

function _notifCommitteeSubCancelled_(substituteUserId, dateStr) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  dateStr = _notifFmtDateLong_(dateStr);
  var name = _notifGetName_(substituteUserId);
  var html = _notifBuildHtml_({
    title: "🚫 Tugas Panitia Ujian Pengganti Dibatalkan",
    accent: "#F97316",
    accent2: "#C2410C",
    name: name,
    intro:
      "Tugas Anda sebagai Panitia Ujian Pengganti telah <strong>dibatalkan</strong> oleh admin.",
    badges: ["Tanggal: " + dateStr],
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Panitia Pengganti Dibatalkan · " + dateStr,
    html,
  );
}

function _notifSupervisorConfirmed_(userId, dateStr, byName, sessionInfo) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  var name = _notifGetName_(userId);
  dateStr = _notifFmtDateLong_(dateStr);
  byName = byName || "Panitia Ujian / Admin";
  var html = _notifBuildHtml_({
    title: "✅ Kehadiran Pengawas Ruang Ujian Dikonfirmasi",
    accent: "#10B981",
    accent2: "#047857",
    name: name,
    intro:
      "Kehadiran Anda sebagai <strong>Pengawas Ruang Ujian</strong> telah dikonfirmasi oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr].concat(sessionInfo ? [sessionInfo] : []),
    paragraphs: ["Mohon segera mengisi BAP setelah sesi berakhir."],
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Kehadiran Pengawas Dikonfirmasi · " + dateStr,
    html,
  );
}

function _notifSupervisorRevoked_(userId, dateStr, byName, sessionInfo) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  var name = _notifGetName_(userId);
  dateStr = _notifFmtDateLong_(dateStr);
  byName = byName || "Panitia Ujian / Admin";
  var html = _notifBuildHtml_({
    title: "⚠️ Konfirmasi Pengawas Ruang Ujian Dibatalkan",
    accent: "#F59E0B",
    accent2: "#B45309",
    name: name,
    intro:
      "Konfirmasi kehadiran Anda sebagai Pengawas Ruang Ujian telah <strong>dibatalkan</strong> oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr].concat(sessionInfo ? [sessionInfo] : []),
  });
  _notifToUser_(
    userId,
    "[SiM-Guru] Konfirmasi Pengawas Dibatalkan · " + dateStr,
    html,
  );
}

function _notifSupervisorSubstituted_(
  originalUserId,
  substituteUserId,
  dateStr,
  byName,
  sessionInfo,
) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  dateStr = _notifFmtDateLong_(dateStr);
  byName = byName || "Panitia Ujian / Admin";
  var origName = _notifGetName_(originalUserId);
  var subName = _notifGetName_(substituteUserId);
  var htmlSub = _notifBuildHtml_({
    title: "📌 Anda Ditugaskan sebagai Pengawas Ruang Ujian Pengganti",
    accent: "#8B5CF6",
    accent2: "#6D28D9",
    name: subName,
    intro:
      "<strong>" +
      _escHtml_(byName) +
      "</strong> menugaskan Anda menggantikan <strong>" +
      _escHtml_(origName) +
      "</strong> sebagai Pengawas Ruang Ujian.",
    badges: ["Pengganti dari: " + origName, "Tanggal: " + dateStr].concat(
      sessionInfo ? [sessionInfo] : [],
    ),
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengawas Pengganti · " + dateStr,
    htmlSub,
  );
  var htmlOrig = _notifBuildHtml_({
    title: "ℹ️ Tugas Pengawas Ruang Ujian Digantikan",
    accent: "#64748B",
    accent2: "#334155",
    name: origName,
    intro:
      "Tugas Pengawas Ruang Ujian Anda dialihkan ke <strong>" +
      _escHtml_(subName) +
      "</strong> oleh " +
      _escHtml_(byName) +
      ".",
    badges: ["Diganti oleh: " + subName, "Tanggal: " + dateStr].concat(
      sessionInfo ? [sessionInfo] : [],
    ),
  });
  _notifToUser_(
    originalUserId,
    "[SiM-Guru] Tugas Pengawas Digantikan · " + dateStr,
    htmlOrig,
  );
}

function _notifSupervisorSubCancelled_(
  substituteUserId,
  dateStr,
  byName,
  sessionInfo,
) {
  if (!_notifIsExamEnabled_(dateStr)) return;
  dateStr = _notifFmtDateLong_(dateStr);
  byName = byName || "Panitia Ujian / Admin";
  var name = _notifGetName_(substituteUserId);
  var html = _notifBuildHtml_({
    title: "🚫 Tugas Pengawas Pengganti Dibatalkan",
    accent: "#F97316",
    accent2: "#C2410C",
    name: name,
    intro:
      "Tugas Anda sebagai Pengawas Ruang Ujian Pengganti telah <strong>dibatalkan</strong> oleh <strong>" +
      _escHtml_(byName) +
      "</strong>.",
    badges: ["Tanggal: " + dateStr].concat(sessionInfo ? [sessionInfo] : []),
  });
  _notifToUser_(
    substituteUserId,
    "[SiM-Guru] Tugas Pengawas Pengganti Dibatalkan · " + dateStr,
    html,
  );
}

function _notifTodayID_() {
  return _notifFmtDateLong_(new Date());
}

function _notifFmtTime_(val) {
  if (val == null || val === "") return "";
  var tz = "Asia/Jakarta";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, tz, "HH:mm");
  }
  var s = String(val).trim();
  if (!s) return "";
  var m = s.match(/(\d{1,2}):(\d{1,2})/);
  if (m) {
    return ("0" + m[1]).slice(-2) + ":" + ("0" + m[2]).slice(-2);
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, "HH:mm");
  return s;
}

function _notifResolveDate_(value) {
  var tz = "Asia/Jakarta";
  if (value == null || value === "") return null;
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : { d: value, hasTime: true };
  }
  var s = String(value).trim();
  if (!s) return null;
  var dateOnly = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dateOnly) {
    var d = new Date(
      parseInt(dateOnly[1], 10),
      parseInt(dateOnly[2], 10) - 1,
      parseInt(dateOnly[3], 10),
    );
    return { d: d, hasTime: false };
  }
  var dateTime = s.match(
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/,
  );
  if (dateTime) {
    var d2 = new Date(
      parseInt(dateTime[1], 10),
      parseInt(dateTime[2], 10) - 1,
      parseInt(dateTime[3], 10),
      parseInt(dateTime[4], 10),
      parseInt(dateTime[5], 10),
      dateTime[6] ? parseInt(dateTime[6], 10) : 0,
    );
    return { d: d2, hasTime: true };
  }
  var fallback = new Date(s);
  if (!isNaN(fallback.getTime())) {
    return { d: fallback, hasTime: true };
  }
  return null;
}

function _notifFmtDate_(value) {
  var tz = "Asia/Jakarta";
  var info = _notifResolveDate_(value);
  if (!info) return String(value || "");
  var months = [
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
  var dd = Utilities.formatDate(info.d, tz, "d");
  var monthIdx = parseInt(Utilities.formatDate(info.d, tz, "M"), 10) - 1;
  var yyyy = Utilities.formatDate(info.d, tz, "yyyy");
  return ("0" + dd).slice(-2) + " " + (months[monthIdx] || "") + " " + yyyy;
}

function _notifFmtDateLong_(value) {
  var tz = "Asia/Jakarta";
  var info = _notifResolveDate_(value);
  if (!info) return String(value || "");
  var days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  var months = [
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
  var dayIdx = parseInt(Utilities.formatDate(info.d, tz, "u"), 10) % 7;
  var dd = Utilities.formatDate(info.d, tz, "d");
  var monthIdx = parseInt(Utilities.formatDate(info.d, tz, "M"), 10) - 1;
  var yyyy = Utilities.formatDate(info.d, tz, "yyyy");
  return (
    (days[dayIdx] || "") +
    ", " +
    ("0" + dd).slice(-2) +
    " " +
    (months[monthIdx] || "") +
    " " +
    yyyy
  );
}

function _notifFmtDateTime_(value) {
  var tz = "Asia/Jakarta";
  var info = _notifResolveDate_(value);
  if (!info) return String(value || "");
  var datePart = _notifFmtDate_(value);
  if (!info.hasTime) return datePart;
  var hhmm = Utilities.formatDate(info.d, tz, "HH:mm");
  return datePart + " " + hhmm;
}

function _notifFmtDateLongTime_(value) {
  var info = _notifResolveDate_(value);
  if (!info) return String(value || "");
  var datePart = _notifFmtDateLong_(value);
  if (!info.hasTime) return datePart;
  var hhmm = Utilities.formatDate(info.d, "Asia/Jakarta", "HH:mm");
  return datePart + " " + hhmm;
}
