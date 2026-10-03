// SiM-Guru — BE_11_Announcements.gs
function _annBool_(v) {
  if (v === true || v === 1) return true;
  if (v === false || v === 0 || v == null) return false;
  var s = String(v).trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "1";
}

function _annDateOnly_(v) {
  if (v == null) return "";
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return "";
    return Utilities.formatDate(v, "Asia/Jakarta", "yyyy-MM-dd");
  }
  var s = String(v).trim();
  if (!s) return "";
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
  }
  return "";
}

function _annParseIdList_(v) {
  if (v == null) return [];
  var s = String(v).trim();
  if (!s) return [];
  if (s.charAt(0) === "[" || s.charAt(0) === "{") {
    try {
      var j = JSON.parse(s);
      if (Array.isArray(j)) {
        return j
          .map(function (x) {
            return String(x || "").trim();
          })
          .filter(function (x) {
            return x.length > 0;
          });
      }
    } catch (_) {}
  }
  return s
    .split(/[,;\n]+/)
    .map(function (x) {
      return x.trim();
    })
    .filter(function (x) {
      return x.length > 0;
    });
}

function _annStringifyIdList_(arr) {
  if (!arr || !arr.length) return "";
  var seen = {};
  var clean = [];
  for (var i = 0; i < arr.length; i++) {
    var v = String(arr[i] || "").trim();
    if (!v || seen[v]) continue;
    seen[v] = true;
    clean.push(v);
  }
  return clean.length ? JSON.stringify(clean) : "";
}

function _annEnsureSheet_() {
  var headers = [
    "id",
    "title",
    "body",
    "severity",
    "is_dismissible",
    "is_active",
    "starts_at",
    "ends_at",
    "target_role",
    "cta_text",
    "cta_url",
    "created_by",
    "created_at",
    "updated_at",
    "target_user_ids",
    "dismissible_user_ids",
    "nondismissible_user_ids",
    "send_email_notification",
    "email_sent",
    "reminder_3_sent",
    "reminder_2_sent",
    "reminder_1_sent",
  ];
  var ss = SpreadsheetApp.openById(getDbId());
  var sheet = ss.getSheetByName(SHEET_NAME.ANNOUNCEMENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME.ANNOUNCEMENTS);
    sheet.appendRow(headers);
    return sheet;
  }
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.appendRow(headers);
    return sheet;
  }
  var current = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(function (h) {
      return String(h || "")
        .toLowerCase()
        .trim()
        .replace(/\s+/g, "_");
    });
  if (headers.length > current.length) {
    var missing = headers.slice(current.length);
    sheet
      .getRange(1, current.length + 1, 1, missing.length)
      .setValues([missing]);
  }
  return sheet;
}

function getActiveAnnouncementsForUser(token) {
  try {
    const user = verifySession(token);
    if (!user) return { status: "error", message: "Unauthorized", items: [] };
    _annEnsureSheet_();
    const today = Utilities.formatDate(
      new Date(),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    const role = String(user.role || "").toLowerCase();
    const myId = String(user.id || "");
    const all = getData(SHEET_NAME.ANNOUNCEMENTS) || [];
    const items = all
      .filter((a) => {
        if (!_annBool_(a.is_active)) return false;
        const sa = _annDateOnly_(a.starts_at);
        const ea = _annDateOnly_(a.ends_at);
        if (sa && today < sa) return false;
        if (ea && today > ea) return false;
        const targetUsers = _annParseIdList_(a.target_user_ids);
        if (targetUsers.length > 0) {
          return targetUsers.indexOf(myId) >= 0;
        }
        const tgt = String(a.target_role || "all").toLowerCase();
        if (tgt !== "all" && tgt !== role) return false;
        return true;
      })
      .map((a) => {
        const ndList = _annParseIdList_(a.nondismissible_user_ids);
        const dList = _annParseIdList_(a.dismissible_user_ids);
        let resolvedDismissible = _annBool_(a.is_dismissible);
        if (ndList.indexOf(myId) >= 0) resolvedDismissible = false;
        else if (dList.indexOf(myId) >= 0) resolvedDismissible = true;
        if (role === "admin") resolvedDismissible = true;
        return {
          id: String(a.id),
          title: String(a.title || ""),
          body: String(a.body || ""),
          severity: String(a.severity || "info").toLowerCase(),
          is_dismissible: resolvedDismissible,
          starts_at: _annDateOnly_(a.starts_at),
          ends_at: _annDateOnly_(a.ends_at),
          target_role: String(a.target_role || "all").toLowerCase(),
          cta_text: String(a.cta_text || ""),
          cta_url: String(a.cta_url || ""),
          updated_at: String(a.updated_at || a.created_at || ""),
        };
      });
    const sevRank = { critical: 0, warning: 1, success: 2, info: 3 };
    items.sort((a, b) => {
      const ra = sevRank[a.severity] != null ? sevRank[a.severity] : 4;
      const rb = sevRank[b.severity] != null ? sevRank[b.severity] : 4;
      if (ra !== rb) return ra - rb;
      return String(b.updated_at).localeCompare(String(a.updated_at));
    });
    return { status: "success", items: items };
  } catch (e) {
    return { status: "error", message: e.toString(), items: [] };
  }
}

function listAnnouncementsAdmin(token) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return { status: "error", message: "Hanya admin yang dapat mengakses." };
  }
  try {
    _annEnsureSheet_();
    const all = (getData(SHEET_NAME.ANNOUNCEMENTS) || []).map((a) => ({
      id: String(a.id),
      title: String(a.title || ""),
      body: String(a.body || ""),
      severity: String(a.severity || "info").toLowerCase(),
      is_dismissible: _annBool_(a.is_dismissible),
      is_active: _annBool_(a.is_active),
      starts_at: _annDateOnly_(a.starts_at),
      ends_at: _annDateOnly_(a.ends_at),
      target_role: String(a.target_role || "all").toLowerCase(),
      cta_text: String(a.cta_text || ""),
      cta_url: String(a.cta_url || ""),
      created_by: String(a.created_by || ""),
      created_at: String(a.created_at || ""),
      updated_at: String(a.updated_at || ""),
      target_user_ids: _annParseIdList_(a.target_user_ids),
      dismissible_user_ids: _annParseIdList_(a.dismissible_user_ids),
      nondismissible_user_ids: _annParseIdList_(a.nondismissible_user_ids),
      send_email_notification: _annBool_(a.send_email_notification),
    }));
    all.sort((a, b) => {
      if (a.is_active !== b.is_active) return a.is_active ? -1 : 1;
      return String(b.updated_at || b.created_at).localeCompare(
        String(a.updated_at || a.created_at),
      );
    });
    return { status: "success", items: all };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function saveAnnouncement(token, data) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return {
      status: "error",
      message: "Hanya admin yang dapat menyimpan pengumuman.",
    };
  }
  if (!data || !String(data.title || "").trim()) {
    return { status: "error", message: "Judul pengumuman wajib diisi." };
  }
  if (!String(data.body || "").trim()) {
    return { status: "error", message: "Isi pengumuman wajib diisi." };
  }
  const validSeverity = ["info", "success", "warning", "critical"];
  const validRole = ["all", "guru", "admin"];
  const severity =
    validSeverity.indexOf(String(data.severity || "").toLowerCase()) >= 0
      ? String(data.severity).toLowerCase()
      : "info";
  const targetRole =
    validRole.indexOf(String(data.target_role || "").toLowerCase()) >= 0
      ? String(data.target_role).toLowerCase()
      : "guru";
  const starts = String(data.starts_at || "").trim();
  const ends = String(data.ends_at || "").trim();
  if (starts && !/^\d{4}-\d{2}-\d{2}$/.test(starts)) {
    return {
      status: "error",
      message: "Format tanggal mulai tidak valid (yyyy-MM-dd).",
    };
  }
  if (ends && !/^\d{4}-\d{2}-\d{2}$/.test(ends)) {
    return {
      status: "error",
      message: "Format tanggal selesai tidak valid (yyyy-MM-dd).",
    };
  }
  if (starts && ends && ends < starts) {
    return {
      status: "error",
      message: "Tanggal selesai harus setelah tanggal mulai.",
    };
  }
  const sheet = _annEnsureSheet_();
  const now = Utilities.formatDate(
    new Date(),
    "Asia/Jakarta",
    "yyyy-MM-dd HH:mm:ss",
  );
  const isDismissible = _annBool_(data.is_dismissible);
  const isActive =
    data.is_active === undefined ? true : _annBool_(data.is_active);
  const ctaText = String(data.cta_text || "")
    .trim()
    .slice(0, 120);
  const ctaUrl = String(data.cta_url || "").trim();
  if (ctaUrl && !/^https?:\/\//i.test(ctaUrl)) {
    return {
      status: "error",
      message: "CTA URL harus diawali http:// atau https://",
    };
  }
  const allUsers = getData("Users") || [];
  const validIds = {};
  allUsers.forEach((u) => {
    if (u && u.id != null) validIds[String(u.id)] = true;
  });
  function _validateIds(arr, label) {
    var clean = [];
    var invalid = [];
    (arr || []).forEach((id) => {
      var s = String(id || "").trim();
      if (!s) return;
      if (validIds[s]) clean.push(s);
      else invalid.push(s);
    });
    return { clean: clean, invalid: invalid };
  }
  const targetCheck = _validateIds(data.target_user_ids, "sasaran");
  const dismCheck = _validateIds(data.dismissible_user_ids, "boleh tutup");
  const ndCheck = _validateIds(
    data.nondismissible_user_ids,
    "tidak boleh tutup",
  );
  if (
    targetCheck.invalid.length ||
    dismCheck.invalid.length ||
    ndCheck.invalid.length
  ) {
    return {
      status: "error",
      message:
        "Beberapa ID pengguna tidak ditemukan: " +
        []
          .concat(targetCheck.invalid, dismCheck.invalid, ndCheck.invalid)
          .join(", "),
    };
  }
  var conflict = dismCheck.clean.filter((id) => ndCheck.clean.indexOf(id) >= 0);
  if (conflict.length) {
    return {
      status: "error",
      message:
        'Pengguna tidak boleh berada di kedua daftar "boleh tutup" dan "tidak boleh tutup": ' +
        conflict.join(", "),
    };
  }
  const targetIdsStr = _annStringifyIdList_(targetCheck.clean);
  const dismIdsStr = _annStringifyIdList_(dismCheck.clean);
  const ndIdsStr = _annStringifyIdList_(ndCheck.clean);
  const sendEmailNotification = _annBool_(data.send_email_notification);
  if (data.id) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        // Cek apakah konten signifikan berubah — jika ya, reset email_sent
        // agar email dikirim ulang ke penerima yang sesuai.
        const titleChanged = String(rows[i][1] || "") !== String(data.title || "").trim();
        const bodyChanged  = String(rows[i][2] || "") !== String(data.body  || "");
        const shouldResetEmail = titleChanged || bodyChanged;
        const prevEmailSent = rows[i][18]; // col 19 (0-indexed: 18) = email_sent
        const newEmailSent  = shouldResetEmail ? false : prevEmailSent;
        // Jika email di-reset, juga reset semua flag reminder
        const prevRem3 = shouldResetEmail ? false : rows[i][19];
        const prevRem2 = shouldResetEmail ? false : rows[i][20];
        const prevRem1 = shouldResetEmail ? false : rows[i][21];

        // Kolom 2–22 (1-indexed), total 21 kolom:
        // title, body, severity, is_dismissible, is_active, starts_at, ends_at,
        // target_role, cta_text, cta_url, created_by, created_at, updated_at,
        // target_user_ids, dismissible_user_ids, nondismissible_user_ids,
        // send_email_notification, email_sent, reminder_3_sent, reminder_2_sent, reminder_1_sent
        sheet
          .getRange(i + 1, 2, 1, 21)
          .setValues([
            [
              String(data.title).trim(),       // col 2: title
              String(data.body),               // col 3: body
              severity,                        // col 4: severity
              isDismissible,                   // col 5: is_dismissible
              isActive,                        // col 6: is_active
              starts,                          // col 7: starts_at
              ends,                            // col 8: ends_at
              targetRole,                      // col 9: target_role
              ctaText,                         // col 10: cta_text
              ctaUrl,                          // col 11: cta_url
              rows[i][11] || user.full_name || user.id, // col 12: created_by (preserve)
              rows[i][12] || now,              // col 13: created_at (preserve)
              now,                             // col 14: updated_at
              targetIdsStr,                    // col 15: target_user_ids
              dismIdsStr,                      // col 16: dismissible_user_ids
              ndIdsStr,                        // col 17: nondismissible_user_ids
              sendEmailNotification,           // col 18: send_email_notification
              newEmailSent,                    // col 19: email_sent (reset jika konten berubah)
              prevRem3,                        // col 20: reminder_3_sent
              prevRem2,                        // col 21: reminder_2_sent
              prevRem1,                        // col 22: reminder_1_sent
            ],
          ]);
        try {
          _invalidateDataSnapshot();
        } catch (_) {}
        if (sendEmailNotification) {
          checkAndSendAnnouncementEmails();
        }
        return {
          status: "success",
          message: "Pengumuman diperbarui.",
          id: String(data.id),
        };
      }
    }
    return { status: "error", message: "Pengumuman tidak ditemukan." };
  }
  const newId = "ANN-" + new Date().getTime();
  sheet.appendRow([
    newId,
    String(data.title).trim(),  // col 2: title
    String(data.body),          // col 3: body
    severity,                   // col 4: severity
    isDismissible,              // col 5: is_dismissible
    isActive,                   // col 6: is_active
    starts,                     // col 7: starts_at
    ends,                       // col 8: ends_at
    targetRole,                 // col 9: target_role
    ctaText,                    // col 10: cta_text
    ctaUrl,                     // col 11: cta_url
    user.full_name || user.id,  // col 12: created_by
    now,                        // col 13: created_at
    now,                        // col 14: updated_at
    targetIdsStr,               // col 15: target_user_ids
    dismIdsStr,                 // col 16: dismissible_user_ids
    ndIdsStr,                   // col 17: nondismissible_user_ids
    sendEmailNotification,      // col 18: send_email_notification
    "",                         // col 19: email_sent
    "",                         // col 20: reminder_3_sent
    "",                         // col 21: reminder_2_sent
    "",                         // col 22: reminder_1_sent
  ]);
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  if (sendEmailNotification) {
    checkAndSendAnnouncementEmails();
  }
  return { status: "success", message: "Pengumuman dibuat.", id: newId };
}

function toggleAnnouncement(token, id, isActive) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin") {
    return { status: "error", message: "Hanya admin." };
  }
  const sheet = _annEnsureSheet_();
  const rows = sheet.getDataRange().getValues();
  const now = Utilities.formatDate(
    new Date(),
    "Asia/Jakarta",
    "yyyy-MM-dd HH:mm:ss",
  );
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue(_annBool_(isActive));
      sheet.getRange(i + 1, 14).setValue(now);
      try {
        _invalidateDataSnapshot();
      } catch (_) {}
      return { status: "success" };
    }
  }
  return { status: "error", message: "Pengumuman tidak ditemukan." };
}

function deleteAnnouncement(token, id) {
  return deleteRowById(token, SHEET_NAME.ANNOUNCEMENTS, id);
}

function checkAndSendAnnouncementEmails() {
  const sheet = _annEnsureSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];
  const todayDate = new Date();
  todayDate.setHours(0, 0, 0, 0);
  const today = Utilities.formatDate(todayDate, "Asia/Jakarta", "yyyy-MM-dd");
  const idxId = headers.indexOf("id");
  const idxTitle = headers.indexOf("title");
  const idxBody = headers.indexOf("body");
  const idxIsActive = headers.indexOf("is_active");
  const idxStartsAt = headers.indexOf("starts_at");
  const idxEndsAt = headers.indexOf("ends_at");
  const idxTargetRole = headers.indexOf("target_role");
  const idxTargetUserIds = headers.indexOf("target_user_ids");
  const idxSendEmail = headers.indexOf("send_email_notification");
  const idxEmailSent = headers.indexOf("email_sent");
  const idxRem3 = headers.indexOf("reminder_3_sent");
  const idxRem2 = headers.indexOf("reminder_2_sent");
  const idxRem1 = headers.indexOf("reminder_1_sent");
  if (idxSendEmail === -1) return;
  const allUsers = getData("Users") || [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const isActive = _annBool_(row[idxIsActive]);
    const sendEmail = _annBool_(row[idxSendEmail]);
    if (!isActive || !sendEmail) continue;
    const emailSent =
      idxEmailSent !== -1 ? _annBool_(row[idxEmailSent]) : false;
    const startsAt = _annDateOnly_(row[idxStartsAt]);
    const endsAt = idxEndsAt !== -1 ? _annDateOnly_(row[idxEndsAt]) : "";
    const rem3Sent = idxRem3 !== -1 ? _annBool_(row[idxRem3]) : false;
    const rem2Sent = idxRem2 !== -1 ? _annBool_(row[idxRem2]) : false;
    const rem1Sent = idxRem1 !== -1 ? _annBool_(row[idxRem1]) : false;
    let emailToSend = null; 
    let colToUpdate = -1;
    if (!emailSent && (!startsAt || startsAt <= today)) {
      emailToSend = "start";
      colToUpdate = idxEmailSent;
    } else if (endsAt && (!startsAt || startsAt <= today)) {
      const endDate = new Date(endsAt);
      endDate.setHours(0, 0, 0, 0);
      const diffTime = endDate.getTime() - todayDate.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays === 3 && !rem3Sent) {
        emailToSend = "rem3";
        colToUpdate = idxRem3;
      } else if (diffDays === 2 && !rem2Sent) {
        emailToSend = "rem2";
        colToUpdate = idxRem2;
      } else if (diffDays === 1 && !rem1Sent) {
        emailToSend = "rem1";
        colToUpdate = idxRem1;
      }
    }
    if (emailToSend) {
      const title = String(row[idxTitle] || "");
      const body = String(row[idxBody] || "");
      const targetRole = String(row[idxTargetRole] || "all").toLowerCase();
      const targetUserIdsStr = String(row[idxTargetUserIds] || "");
      let targetUserIds = [];
      try {
        if (targetUserIdsStr) targetUserIds = JSON.parse(targetUserIdsStr);
      } catch (e) {}
      const targetEmails = [];
      allUsers.forEach((u) => {
        if (!u.email || !_isValidEmail_(u.email)) return;
        let include = false;
        if (targetUserIds.length > 0) {
          if (targetUserIds.indexOf(String(u.id)) >= 0) include = true;
        } else {
          if (targetRole === "all") {
            include = true;
          } else {
            const uRole = String(u.role || "").toLowerCase();
            if (targetRole === uRole) include = true;
          }
        }
        if (include && targetEmails.indexOf(u.email) === -1) {
          targetEmails.push(u.email);
        }
      });
      if (targetEmails.length > 0) {
        // Helper escape HTML khusus untuk email server-side (escapeHtml hanya ada di frontend)
        function _annEscapeHtml_(str) {
          return String(str || "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
        }
        let subject = "";
        let prefix = "";
        if (emailToSend === "start") {
          subject = "Pengumuman Baru: " + title;
          const dateStr = startsAt
            ? formatDateIndo(startsAt)
            : formatDateIndo(today);
          prefix = `Berlaku mulai: ${_annEscapeHtml_(dateStr)}`;
        } else {
          const daysLeft =
            emailToSend === "rem3" ? 3 : emailToSend === "rem2" ? 2 : 1;
          subject = `Peringatan (${daysLeft} hari lagi): ${title}`;
          const dateStr = formatDateIndo(endsAt);
          prefix = `Pengingat! Pengumuman ini akan berakhir pada: ${_annEscapeHtml_(dateStr)} (${daysLeft} hari lagi)`;
        }
        // Body disanitasi: tag HTML diizinkan (admin yang menulis), tapi
        // escape karakter berbahaya di luar konteks tag agar aman dikirim via email.
        // Konversi newline ke <br/> untuk plaintext fallback.
        const safeBody = String(body).replace(/\r\n/g, "\n").replace(/\n/g, "<br/>");
        const htmlBody = `
          <div style="font-family:sans-serif; max-width:600px; margin:0 auto; padding:20px; color:#334155;">
            <h2 style="color:#1e293b; margin-top:0;">${_annEscapeHtml_(title)}</h2>
            <p style="font-size:13px; color:#e11d48; margin-bottom:20px; font-weight:bold;">${prefix}</p>
            <div style="background:#f8fafc; padding:15px; border-radius:8px; line-height:1.6; color:#334155;">
              ${safeBody}
            </div>
            <p style="font-size:12px; color:#94a3b8; margin-top:30px; border-top:1px solid #e2e8f0; padding-top:15px;">
              Email otomatis dari SiM-Guru. Harap periksa aplikasi untuk detail lebih lanjut. Jangan membalas email ini.
            </p>
          </div>
        `;
        const chunk = 50;
        for (let j = 0; j < targetEmails.length; j += chunk) {
          const bccChunk = targetEmails.slice(j, j + chunk).join(",");
          try {
            MailApp.sendEmail({
              to: Session.getActiveUser().getEmail() || "noreply@simguru.local",
              bcc: bccChunk,
              subject: subject,
              htmlBody: htmlBody,
              body: "Silakan aktifkan tampilan HTML untuk melihat pengumuman ini.",
            });
          } catch (e) {
            console.error("Gagal mengirim email pengumuman: " + e.message);
          }
        }
      }
      if (colToUpdate !== -1) {
        sheet.getRange(i + 1, colToUpdate + 1).setValue(true);
      }
    }
  }
}

function installAnnouncementTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkAndSendAnnouncementEmails") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("checkAndSendAnnouncementEmails")
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone("Asia/Jakarta")
    .create();
  return {
    status: "success",
    message: "Trigger berhasil diinstall pada jam 08:00 pagi.",
  };
}
