// SiM-Guru — BE_18_SecurityRecovery.gs
function verifyForgotSecurityAnswer(username, securityAnswer) {
  try {
    const uname = String(username || "").trim();
    if (!uname)
      return { status: "error", message: "Username tidak boleh kosong." };
    const user = findData("Users", "username", uname);
    if (!user) return { status: "error", message: "Username tidak ditemukan." };
    const storedQuestion = String(user.security_question || "").trim();
    if (!storedQuestion) {
      return {
        status: "error",
        message: "Akun ini belum memiliki pertanyaan keamanan. Hubungi Admin.",
      };
    }
    const storedAnswer = String(user.security_answer || "")
      .toLowerCase()
      .trim();
    const givenAnswer = String(securityAnswer || "")
      .toLowerCase()
      .trim();
    if (!givenAnswer)
      return {
        status: "error",
        message: "Jawaban keamanan tidak boleh kosong.",
      };
    if (storedAnswer !== givenAnswer) {
      return {
        status: "error",
        message: "Jawaban keamanan tidak sesuai. Periksa kembali jawaban Anda.",
      };
    }
    const resetToken = _issueForgotResetToken_(uname, "secq");
    return {
      status: "success",
      message: "Verifikasi berhasil. Silakan buat password baru.",
      reset_token: resetToken,
      expires_in: 5 * 60,
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}

function _issueForgotResetToken_(username, method) {
  const props = PropertiesService.getScriptProperties();
  const token =
    Utilities.getUuid().replace(/-/g, "") +
    String(Math.floor(Math.random() * 1e8));
  const key = "fp_reset_token_" + token;
  const payload = {
    username: String(username),
    method: String(method || "unknown"),
    expires: Date.now() + 5 * 60 * 1000,
  };
  props.setProperty(key, JSON.stringify(payload));
  return token;
}

function _consumeForgotResetToken_(resetToken) {
  try {
    const t = String(resetToken || "").trim();
    if (!t) return { ok: false, message: "Token tidak valid." };
    const props = PropertiesService.getScriptProperties();
    const key = "fp_reset_token_" + t;
    const raw = props.getProperty(key);
    if (!raw)
      return {
        ok: false,
        message: "Token tidak ditemukan atau sudah dipakai.",
      };
    let data;
    try {
      data = JSON.parse(raw);
    } catch (e) {
      props.deleteProperty(key);
      return { ok: false, message: "Token rusak." };
    }
    if (Date.now() > Number(data.expires || 0)) {
      props.deleteProperty(key);
      return {
        ok: false,
        message: "Token sudah kedaluwarsa. Ulangi proses verifikasi.",
      };
    }
    props.deleteProperty(key);
    return {
      ok: true,
      username: String(data.username || ""),
      method: String(data.method || ""),
    };
  } catch (e) {
    return {
      ok: false,
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}

function getForgotMethods(username) {
  const user = findData("Users", "username", String(username || "").trim());
  if (!user) return { status: "error", message: "Username tidak ditemukan." };
  const hasSecQ = !!(
    user.security_question && String(user.security_question).trim()
  );
  const email = String(user.email || "").trim();
  const hasEmail = !!(email && _isValidEmail_(email));
  return {
    status: "success",
    has_security_question: hasSecQ,
    has_email_verified: hasEmail,
    masked_email: hasEmail ? _maskEmail_(email) : "",
  };
}

function _maskEmail_(email) {
  try {
    const e = String(email || "").trim();
    const at = e.indexOf("@");
    if (at < 1) return e;
    const local = e.substring(0, at);
    const domain = e.substring(at);
    if (local.length <= 2) return local.charAt(0) + "***" + domain;
    return local.substring(0, 2) + "***" + domain;
  } catch (_) {
    return String(email || "");
  }
}

function requestForgotPasswordOTP(username) {
  try {
    const uname = String(username || "").trim();
    if (!uname)
      return { status: "error", message: "Username tidak boleh kosong." };
    const user = findData("Users", "username", uname);
    if (!user) return { status: "error", message: "Username tidak ditemukan." };
    const email = String(user.email || "")
      .trim()
      .toLowerCase();
    if (!email || !_isValidEmail_(email)) {
      return {
        status: "error",
        message:
          "Akun ini belum memiliki email terverifikasi. Hubungi Admin atau gunakan pertanyaan keamanan.",
      };
    }
    const props = PropertiesService.getScriptProperties();
    const rateKey = "fp_otp_rate_" + uname;
    const dataKey = "fp_otp_data_" + uname;
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
    const payload = {
      code: code,
      email: email,
      username: uname,
      expires: now + 10 * 60 * 1000,
      attempts: 0,
    };
    props.setProperty(dataKey, JSON.stringify(payload));
    props.setProperty(rateKey, String(now));
    try {
      const subject = "Kode Reset Password — SiM-Guru";
      const htmlBody =
        '<div style="font-family:Inter,Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px;color:#0f172a;">' +
        '<div style="background:linear-gradient(135deg,#B45309,#D97706);color:white;padding:20px 24px;border-radius:14px 14px 0 0;">' +
        '<div style="font-size:13px;opacity:0.85;letter-spacing:0.05em;">SiM-GURU</div>' +
        '<div style="font-size:20px;font-weight:800;margin-top:4px;">Kode Reset Password</div>' +
        "</div>" +
        '<div style="border:1px solid #e2e8f0;border-top:0;border-radius:0 0 14px 14px;padding:24px;background:#fff;">' +
        '<p style="font-size:14px;color:#334155;margin:0 0 16px;">Halo <strong>' +
        _escHtml_(user.full_name || "Pengguna") +
        "</strong>,</p>" +
        '<p style="font-size:14px;color:#334155;margin:0 0 16px;">Anda meminta reset password untuk akun <strong>' +
        _escHtml_(uname) +
        "</strong>. Gunakan kode verifikasi berikut untuk mengirim permintaan reset ke admin:</p>" +
        '<div style="background:#FEF3C7;border:1px dashed #D97706;border-radius:12px;padding:18px;text-align:center;margin:16px 0;">' +
        "<div style=\"font-family:'JetBrains Mono',monospace;font-size:32px;font-weight:800;letter-spacing:0.4em;color:#92400E;\">" +
        code +
        "</div>" +
        "</div>" +
        '<p style="font-size:13px;color:#64748b;margin:16px 0 8px;">⏱️ Kode berlaku <strong>10 menit</strong>.</p>' +
        '<p style="font-size:13px;color:#dc2626;margin:0 0 16px;font-weight:600;">⚠️ Jika Anda tidak meminta reset password, abaikan email ini dan amankan akun Anda.</p>' +
        '<hr style="border:0;border-top:1px solid #e2e8f0;margin:20px 0;">' +
        '<p style="font-size:11px;color:#94a3b8;margin:0;">Email otomatis dari SiM-Guru. Jangan balas email ini.</p>' +
        "</div>" +
        "</div>";
      const textBody =
        "Kode reset password SiM-Guru untuk akun " +
        uname +
        ": " +
        code +
        "\n\nKode berlaku 10 menit. Jika Anda tidak meminta reset password, abaikan email ini.";
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
          "Gagal mengirim email: " +
          (mailErr && mailErr.message ? mailErr.message : mailErr),
      };
    }
    return {
      status: "success",
      message: "Kode OTP telah dikirim ke " + _maskEmail_(email) + ".",
      masked_email: _maskEmail_(email),
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

function requestPasswordResetByOTP(username, otpCode) {
  try {
    const uname = String(username || "").trim();
    const code = String(otpCode || "").trim();
    if (!uname)
      return { status: "error", message: "Username tidak boleh kosong." };
    if (!/^\d{6}$/.test(code))
      return { status: "error", message: "Kode OTP harus 6 digit angka." };
    const user = findData("Users", "username", uname);
    if (!user) return { status: "error", message: "Username tidak ditemukan." };
    const props = PropertiesService.getScriptProperties();
    const dataKey = "fp_otp_data_" + uname;
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
    props.deleteProperty(dataKey);
    const resetToken = _issueForgotResetToken_(uname, "otp");
    return {
      status: "success",
      message: "Verifikasi OTP berhasil. Silakan buat password baru.",
      reset_token: resetToken,
      expires_in: 5 * 60,
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}

function resetPasswordWithToken(resetToken, newPassword) {
  try {
    const consumed = _consumeForgotResetToken_(resetToken);
    if (!consumed.ok)
      return {
        status: "error",
        message: consumed.message || "Token tidak valid.",
      };
    const uname = String(consumed.username || "").trim();
    if (!uname)
      return {
        status: "error",
        message: "Token tidak terkait dengan akun manapun.",
      };
    const user = findData("Users", "username", uname);
    if (!user) return { status: "error", message: "Akun tidak ditemukan." };
    const newPass = String(newPassword || "");
    if (!newPass)
      return { status: "error", message: "Password baru wajib diisi." };
    if (newPass === "123456") {
      return {
        status: "error",
        message: "Password baru tidak boleh password default.",
      };
    }
    const currentExtPass = _extGetPassword_(uname);
    if (currentExtPass !== null && String(currentExtPass) === newPass) {
      return {
        status: "error",
        message: "Password baru tidak boleh sama dengan password lama.",
      };
    }
    const complexityCheck = _checkPasswordComplexity(
      newPass,
      user.full_name,
      uname,
    );
    if (!complexityCheck.valid) {
      return { status: "error", message: complexityCheck.message };
    }
    const extOk = _extSetPassword_(uname, newPass);
    if (!extOk)
      return {
        status: "error",
        message: "Gagal memperbarui password. Hubungi Administrator.",
      };
    const sheetData = getData("Users");
    const userRowIndex = sheetData.findIndex(
      (u) => String(u.username) === uname,
    );
    if (userRowIndex !== -1) {
      const sheetObj = getSheet("Users");
      sheetObj.getRange(userRowIndex + 2, 3).setValue(newPass);
    }
    try {
      const sheet = getSheet("Reset_Requests");
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (
          String(rows[i][1]) === uname &&
          String(rows[i][3]).toLowerCase() === "pending"
        ) {
          sheet.deleteRow(i + 1);
        }
      }
    } catch (_) {}
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    try {
      const email = String(user.email || "")
        .trim()
        .toLowerCase();
      if (email && _isValidEmail_(email)) {
        const subject = "Password SiM-Guru Berhasil Diubah";
        const htmlBody =
          '<div style="font-family:Inter,Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px;color:#0f172a;">' +
          '<div style="background:linear-gradient(135deg,#16A34A,#15803D);color:white;padding:18px 22px;border-radius:14px 14px 0 0;">' +
          '<div style="font-size:13px;opacity:0.85;letter-spacing:0.05em;">SiM-GURU</div>' +
          '<div style="font-size:18px;font-weight:800;margin-top:4px;">Password Berhasil Diubah</div>' +
          "</div>" +
          '<div style="border:1px solid #e2e8f0;border-top:0;border-radius:0 0 14px 14px;padding:22px;background:#fff;">' +
          '<p style="font-size:14px;color:#334155;margin:0 0 12px;">Halo <strong>' +
          _escHtml_(user.full_name || uname) +
          "</strong>,</p>" +
          '<p style="font-size:14px;color:#334155;margin:0 0 12px;">Password akun <strong>' +
          _escHtml_(uname) +
          "</strong> baru saja diubah melalui flow Lupa Password (" +
          (consumed.method === "otp" ? "OTP email" : "pertanyaan keamanan") +
          ").</p>" +
          '<p style="font-size:13px;color:#dc2626;margin:0 0 12px;font-weight:600;">⚠️ Jika Anda merasa tidak melakukan perubahan ini, segera hubungi Admin.</p>' +
          '<p style="font-size:11px;color:#94a3b8;margin:12px 0 0;">Email otomatis dari SiM-Guru. Jangan balas email ini.</p>' +
          "</div>" +
          "</div>";
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: htmlBody,
          body:
            "Password akun " +
            uname +
            " berhasil diubah. Jika bukan Anda, segera hubungi Admin.",
          name: "SiM-Guru",
        });
      }
    } catch (_) {}
    return {
      status: "success",
      message:
        "Password berhasil diubah. Silakan login dengan password baru Anda.",
    };
  } catch (e) {
    return {
      status: "error",
      message: "Server error: " + (e && e.message ? e.message : e),
    };
  }
}

function getSecurityQuestion(username) {
  const user = findData("Users", "username", username);
  if (!user) return { status: "error", message: "Username tidak ditemukan." };
  const question = String(user.security_question || "").trim();
  if (!question) {
    return {
      status: "error",
      message:
        "Akun ini belum memiliki pertanyaan keamanan. Hubungi Admin secara langsung.",
    };
  }
  return { status: "success", question: question };
}

function getMySecurityQuestion(token) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Sesi habis" };
  const question = String(user.security_question || "").trim();
  return { status: "success", question: question };
}

function saveSecurityQuestion(token, question, answer) {
  const user = verifySession(token);
  if (!user)
    return { status: "error", message: "Sesi habis, silakan login kembali" };
  if (!question || !answer)
    return { status: "error", message: "Pertanyaan dan jawaban harus diisi" };
  if (String(answer).trim().length < 2)
    return { status: "error", message: "Jawaban terlalu pendek" };
  const sheetObj = getSheet("Users");
  const headers = sheetObj
    .getRange(1, 1, 1, sheetObj.getLastColumn())
    .getValues()[0];
  let sqIdx = headers.findIndex(
    (h) => String(h).toLowerCase() === "security_question",
  );
  let saIdx = headers.findIndex(
    (h) => String(h).toLowerCase() === "security_answer",
  );
  if (sqIdx === -1) {
    const lastCol = sheetObj.getLastColumn();
    sheetObj.getRange(1, lastCol + 1).setValue("security_question");
    sheetObj.getRange(1, lastCol + 2).setValue("security_answer");
    sqIdx = lastCol;
    saIdx = lastCol + 1;
  }
  const sheetData = getData("Users");
  const userRowIndex = sheetData.findIndex(
    (u) => String(u.id) === String(user.id),
  );
  if (userRowIndex === -1)
    return { status: "error", message: "User tidak ditemukan" };
  const sheetRowNumber = userRowIndex + 2;
  sheetObj.getRange(sheetRowNumber, sqIdx + 1).setValue(question.trim());
  sheetObj
    .getRange(sheetRowNumber, saIdx + 1)
    .setValue(String(answer).toLowerCase().trim());
  return {
    status: "success",
    message: "Pertanyaan keamanan berhasil disimpan",
  };
}

function getResetRequests(token) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return { status: "error", message: "Unauthorized" };
    }
    const allData = getData("Reset_Requests");
    if (!allData || !Array.isArray(allData)) {
      return { status: "success", data: [], message: "No data available" };
    }
    const reqs = allData
      .filter((r) => r && String(r.status).toLowerCase() === "pending")
      .map((r) => ({
        request_id: String(r.request_id || ""),
        username: String(r.username || ""),
        full_name: String(r.full_name || ""),
        status: String(r.status || ""),
        created_at:
          r.created_at instanceof Date
            ? Utilities.formatDate(
                r.created_at,
                Session.getScriptTimeZone(),
                "yyyy-MM-dd HH:mm",
              )
            : String(r.created_at || ""),
      }));
    return { status: "success", data: reqs };
  } catch (error) {
    return { status: "error", message: error.toString(), data: [] };
  }
}

function approveResetRequest(token, requestId, username) {
  const user = verifySession(token);
  if (!user || user.role !== "admin") return { status: "error" };
  const targetUser = findData("Users", "username", username);
  if (targetUser) {
    resetUserPassword(token, targetUser.id);
  }
  const sheet = getSheet("Reset_Requests");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(requestId)) {
      sheet.getRange(i + 1, 4).setValue("Completed");
      return { status: "success" };
    }
  }
  return { status: "error" };
}

function getPendingResetCount(token) {
  const user = verifySession(token);
  if (!user || user.role !== "admin") return { status: "error" };
  const count = getData("Reset_Requests").filter(
    (r) => r.status === "Pending",
  ).length;
  return { status: "success", count: count };
}
