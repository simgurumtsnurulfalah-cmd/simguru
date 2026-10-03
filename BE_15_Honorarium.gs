// SiM-Guru — BE_15_Honorarium.gs
function calculateMonthlyBonus(userId, month, year) {
  const snapshot = _getDataSnapshot();
  return _calculateMonthlyBonusFast(userId, month, year, snapshot);
}

function getAvailableHonorPeriods(token) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Unauthorized" };
  const history = getData("Honor_History");
  const months = [
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
  const parseP = (p) => {
    const parts = String(p || "").split(" ");
    return (parseInt(parts[1]) || 0) * 12 + months.indexOf(parts[0]);
  };
  let filtered = history;
  if (String(user.role).toLowerCase() !== "admin") {
    filtered = history.filter((h) => String(h.user_id) === String(user.id));
  }
  const periods = [...new Set(filtered.map((h) => h.periode))].sort(
    (a, b) => parseP(b) - parseP(a),
  );
  return { status: "success", periods: periods };
}

function _sendHonorariumEmail(
  user,
  periodId,
  finalJtm,
  totalAllow,
  grandTotal,
  allowancesInfo,
  transportTotal = 0,
  kbmJtm = 0,
  bonusJtm = 0,
  eventJtm = 0,
) {
  if (!user.email || String(user.email).indexOf("@") === -1) return;
  const subject = "[SiM-Guru] Slip Honorarium · " + periodId;
  const baseSalary = Number(getConfigValue("base_salary") || 0);
  const jtmNominal = finalJtm * baseSalary;
  const fmtRp = (num) => "Rp " + Number(num).toLocaleString("id-ID");
  const recipientName = _escHtml_(user.full_name || user.username || "Guru");

  // --- Bangun baris tabel rincian ---
  const rowStyle =
    "border-bottom:1px solid #E2E8F0;";
  const tdLabelStyle =
    "padding:10px 12px;" + rowStyle + "color:#475569;font-size:13px;";
  const tdValueStyle =
    "padding:10px 12px;" + rowStyle + "color:#0f172a;font-weight:600;font-size:13px;text-align:right;";
  const tdSubLabelStyle =
    "padding:6px 12px 6px 24px;" + rowStyle + "color:#64748b;font-size:12px;";
  const tdSubValueStyle =
    "padding:6px 12px;" + rowStyle + "color:#475569;font-size:12px;text-align:right;";

  // Tentukan apakah perlu tampilkan rincian komponen JTM (jika ada lebih dari satu komponen)
  const hasBreakdown = (bonusJtm > 0 || eventJtm > 0) && kbmJtm > 0;

  let bodyRows =
    "<tr>" +
    '<td style="' + tdLabelStyle + '">' +
    "Honor Mengajar (" + finalJtm + " JTM &times; " + fmtRp(baseSalary) + ")" +
    "</td>" +
    '<td style="' + tdValueStyle + '">' + fmtRp(jtmNominal) + "</td>" +
    "</tr>";

  // Sub-baris rincian komponen JTM jika ada bonus atau event
  if (hasBreakdown) {
    bodyRows +=
      "<tr>" +
      '<td style="' + tdSubLabelStyle + '">↳ KBM / Piket / Ujian</td>' +
      '<td style="' + tdSubValueStyle + '">' + kbmJtm + " JTM</td>" +
      "</tr>";
    if (bonusJtm > 0) {
      bodyRows +=
        "<tr>" +
        '<td style="' + tdSubLabelStyle + '">↳ Bonus JTM Libur</td>' +
        '<td style="' + tdSubValueStyle + '">' + bonusJtm + " JTM</td>" +
        "</tr>";
    }
    if (eventJtm > 0) {
      bodyRows +=
        "<tr>" +
        '<td style="' + tdSubLabelStyle + '">↳ Kehadiran Acara/Kegiatan</td>' +
        '<td style="' + tdSubValueStyle + '">' + eventJtm + " JTM</td>" +
        "</tr>";
    }
  } else if (eventJtm > 0 && kbmJtm === 0) {
    // Edge case: hanya ada event JTM tanpa KBM
    bodyRows +=
      "<tr>" +
      '<td style="' + tdSubLabelStyle + '">↳ Kehadiran Acara/Kegiatan</td>' +
      '<td style="' + tdSubValueStyle + '">' + eventJtm + " JTM</td>" +
      "</tr>";
  }

  if (allowancesInfo && allowancesInfo.length > 0) {
    allowancesInfo.forEach((a) => {
      bodyRows +=
        "<tr>" +
        '<td style="' + tdLabelStyle + '">' + _escHtml_(String(a.duty_name || "")) + "</td>" +
        '<td style="' + tdValueStyle + '">' + fmtRp(a.amount) + "</td>" +
        "</tr>";
    });
  }
  if (transportTotal > 0) {
    bodyRows +=
      "<tr>" +
      '<td style="' + tdLabelStyle + '">Tunjangan Transportasi</td>' +
      '<td style="' + tdValueStyle + '">' + fmtRp(transportTotal) + "</td>" +
      "</tr>";
  }
  if (!(allowancesInfo && allowancesInfo.length > 0) && transportTotal <= 0) {
    bodyRows +=
      "<tr>" +
      '<td colspan="2" style="padding:10px 12px;' + rowStyle + 'color:#94a3b8;font-size:13px;font-style:italic;text-align:center;">Tidak ada tunjangan tambahan</td>' +
      "</tr>";
  }

  // --- Blok tabel slip (disisipkan ke dalam _notifBuildHtml_ via paragraphs) ---
  const tableBlock =
    '<div style="margin:16px 0;">' +
    '<table style="width:100%;border-collapse:collapse;font-family:Inter,Arial,sans-serif;">' +
    "<thead>" +
    "<tr>" +
    '<th style="padding:10px 12px;background:#EEF2FF;color:#4338CA;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;text-align:left;border-radius:8px 0 0 8px;">Keterangan</th>' +
    '<th style="padding:10px 12px;background:#EEF2FF;color:#4338CA;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;text-align:right;border-radius:0 8px 8px 0;">Nominal</th>' +
    "</tr>" +
    "</thead>" +
    "<tbody>" + bodyRows + "</tbody>" +
    "<tfoot>" +
    "<tr>" +
    '<td style="padding:12px;border-top:2px solid #E2E8F0;color:#0f172a;font-size:14px;font-weight:800;">TOTAL DITERIMA</td>' +
    '<td style="padding:12px;border-top:2px solid #E2E8F0;color:#10B981;font-size:16px;font-weight:800;text-align:right;">' + fmtRp(grandTotal) + "</td>" +
    "</tr>" +
    "</tfoot>" +
    "</table>" +
    "</div>";

  // Badge tambahan: tampilkan rincian JTM jika ada komponen acara/bonus
  const badges = ["Periode: " + periodId, "Total JTM: " + finalJtm + " JTM"];
  if (eventJtm > 0) badges.push("Acara/Kegiatan: " + eventJtm + " JTM");
  if (bonusJtm > 0) badges.push("Bonus Libur: " + bonusJtm + " JTM");

  const htmlBody = _notifBuildHtml_({
    title: "💰 Slip Honorarium Telah Diterbitkan",
    accent: "#10B981",
    accent2: "#059669",
    name: recipientName,
    intro:
      "Honorarium Anda untuk periode <strong>" +
      _escHtml_(periodId) +
      "</strong> telah difinalisasi. Berikut rincian slip Anda:",
    badges: badges,
    paragraphs: [
      tableBlock,
      '<span style="color:#94a3b8;font-size:12px;">Harap simpan email ini sebagai bukti penerimaan honorarium Anda. Detail lengkap dapat dilihat melalui menu <strong>Honorarium</strong> di aplikasi SiM-Guru.</span>',
    ],
  });

  try {
    MailApp.sendEmail({
      to: user.email,
      subject: subject,
      body:
        "Slip Honorarium Periode " +
        periodId +
        "\n\nTotal JTM: " + finalJtm + " JTM" +
        (eventJtm > 0 ? "\n  - Kehadiran Acara/Kegiatan: " + eventJtm + " JTM" : "") +
        (bonusJtm > 0 ? "\n  - Bonus JTM Libur: " + bonusJtm + " JTM" : "") +
        "\nTotal Diterima: " + fmtRp(grandTotal) +
        "\n\nBuka aplikasi SiM-Guru untuk detail lengkap.",
      htmlBody: htmlBody,
      name: "SiM-Guru",
    });
  } catch (e) {
    console.error("Gagal mengirim email ke " + user.email + ": " + e);
  }
}

function _processHonorariumFinalization(useMonth, useYear) {
  const months = [
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
  const periodId = `${months[useMonth]} ${useYear}`;
  const historySheet = getSheet("Honor_History");
  const allHistory = getData("Honor_History");
  const existing = allHistory.find(
    (h) => formatPeriode(h.periode) === periodId,
  );
  if (existing)
    return {
      status: "error",
      message: "Periode " + periodId + " sudah difinalisasi sebelumnya.",
    };
  const allUsers = getData("Users").filter((u) => u.role === "guru");
  const allLogs = getData("Teaching_Logs");
  const allAllowances = getData("Allowances");
  const allSchedules = getData("Schedules");
  const baseSalary = Number(getConfigValue("base_salary"));
  const newRows = [];
  const emailsToSend = [];
  const jtmAdjMap = _jtmBuildKbmAdjustmentMap_();
  const piketAdjMap = _jtmBuildPiketAdjustmentMap_();
  const firstDayEvent =
    useYear + "-" + String(useMonth + 1).padStart(2, "0") + "-01";
  const lastDayEventDate = new Date(useYear, useMonth + 1, 0);
  const lastDayEvent =
    useYear +
    "-" +
    String(useMonth + 1).padStart(2, "0") +
    "-" +
    String(lastDayEventDate.getDate()).padStart(2, "0");
  const eventAttAll = getData(EVENT_SHEET.ATTENDANCE);
  allUsers.forEach((u) => {
    const logs = allLogs.filter((l) => {
      const d = new Date(l.date);
      return (
        String(l.user_id) === String(u.id) &&
        d.getMonth() === useMonth &&
        d.getFullYear() === useYear
      );
    });
    let totalJtm = 0;
    logs.forEach((l) => {
      const _sid = String(l.schedule_id);
      if (_sid === "PICKET-DUTY" || _sid === "PIKET") totalJtm += _jtmResolvePiketLogJtm_(piketAdjMap, l.user_id, l.date, 4);
      else if (_sid === "CEREMONY-DUTY" || _sid === "UPACARA") totalJtm += 5;
      else if (_sid === "EXAM-SUPERVISOR" || _sid === "EXAM-COMMITTEE")
        totalJtm += Number(l.jtm_val || 0);
      else if (_sid === "PARTIAL-SUB-KBM" || _sid === "PARTIAL-SUB-EXAM")
        totalJtm += Number(l.jtm_val || 0);
      else {
        const s = allSchedules.find((sch) => String(sch.id) === _sid);
        if (s)
          totalJtm += _jtmResolveKbmLogJtm_(
            jtmAdjMap,
            _sid,
            l.date,
            Number(s.jtm_val || 0),
          );
        else totalJtm += Number(l.jtm_val || 0);
      }
    });
    const totalBonus = calculateMonthlyBonus(u.id, useMonth, useYear);
    const eventJtm = _computeEventJtm(
      eventAttAll.filter(
        (r) =>
          String(r.user_id) === String(u.id) &&
          r.date >= firstDayEvent &&
          r.date <= lastDayEvent,
      ),
    );
    const finalJtm = totalJtm + totalBonus + eventJtm;
    const userAllowances = allAllowances.filter(
      (a) => String(a.user_id) === String(u.id),
    );
    const totalAllow = userAllowances.reduce(
      (sum, a) => sum + Number(a.amount),
      0,
    );
    const transportDesimal = _getMonthlyTransportSum(String(u.id), periodId);
    const transportTotal = _applyThreePointRounding(transportDesimal);
    const grandTotal = finalJtm * baseSalary + totalAllow + transportTotal;
    if (grandTotal > 0) {
      const trxId = "TRX-" + Utilities.getUuid();
      const allowanceSnapshot = userAllowances.map((a) => ({
        duty_name: String(a.duty_name || ""),
        amount: Number(a.amount || 0),
      }));
      const detailsJson = JSON.stringify({
        allowances: allowanceSnapshot,
        allow: totalAllow,
        bonus: totalBonus,
        transport_total: transportTotal,
        event_jtm: eventJtm,
      });
      newRows.push([
        generateId("HIS"),
        u.id,
        periodId,
        finalJtm,
        grandTotal,
        new Date(),
        detailsJson,
        trxId,
      ]);
      emailsToSend.push({
        user: u,
        periodId: periodId,
        finalJtm: finalJtm,
        totalAllow: totalAllow,
        grandTotal: grandTotal,
        allowancesInfo: allowanceSnapshot,
        transportTotal: transportTotal,
        kbmJtm: totalJtm,
        bonusJtm: totalBonus,
        eventJtm: eventJtm,
      });
    }
  });
  if (newRows.length > 0) {
    const lastRow = historySheet.getLastRow();
    const startRow = lastRow + 1;
    historySheet.getRange(startRow, 3, newRows.length, 1).setNumberFormat("@");
    historySheet.getRange(startRow, 1, newRows.length, 8).setValues(newRows);
    emailsToSend.forEach((e) => {
      _sendHonorariumEmail(
        e.user,
        e.periodId,
        e.finalJtm,
        e.totalAllow,
        e.grandTotal,
        e.allowancesInfo,
        e.transportTotal,
        e.kbmJtm,
        e.bonusJtm,
        e.eventJtm,
      );
    });
  }
  return {
    status: "success",
    message:
      "Berhasil memfinalisasi " +
      newRows.length +
      " data honorarium untuk periode " +
      periodId,
  };
}

function saveMonthlyHonorarium(token, targetMonth, targetYear) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Akses ditolak" };
  const now = new Date();
  const useMonth =
    targetMonth !== undefined && targetMonth !== null && targetMonth !== ""
      ? parseInt(targetMonth)
      : now.getMonth();
  const useYear =
    targetYear !== undefined && targetYear !== null && targetYear !== ""
      ? parseInt(targetYear)
      : now.getFullYear();
  return _processHonorariumFinalization(useMonth, useYear);
}

function autoFinalizeHonorarium() {
  const now = new Date();
  let targetMonth = now.getMonth() - 1;
  let targetYear = now.getFullYear();
  if (targetMonth < 0) {
    targetMonth = 11;
    targetYear--;
  }
  _processHonorariumFinalization(targetMonth, targetYear);
}

function setupAutoFinalizeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "autoFinalizeHonorarium") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("autoFinalizeHonorarium")
    .timeBased()
    .onMonthDay(2)
    .atHour(10)
    .create();
  return "Trigger finalisasi otomatis berhasil dibuat (Setiap tanggal 2 jam 10.00).";
}

function getMyHonorariumData(token, period) {
  const user = verifySession(token);
  if (!user) return { status: "error", message: "Session expired" };
  const snapshot = _getDataSnapshot();
  const now = new Date();
  let targetMonth = now.getMonth();
  let targetYear = now.getFullYear();
  if (period && period !== "current") {
    const parts = period.split(" ");
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
    targetMonth = monthNames.indexOf(parts[0]);
    targetYear = parseInt(parts[1]);
    if (isNaN(targetMonth) || targetMonth === -1) {
      const d = new Date(period);
      if (!isNaN(d.getTime())) {
        targetMonth = d.getMonth();
        targetYear = d.getFullYear();
      }
    }
  }
  const baseSalary = Number(snapshot.config["base_salary"] || 0);
  const userLogs = (snapshot.logsByUserId[user.id] || []).filter((l) => {
    const d = new Date(l.date);
    return d.getMonth() === targetMonth && d.getFullYear() === targetYear;
  });
  const isExamLog_ = (l) => {
    const s = String(l.schedule_id);
    return s === "EXAM-SUPERVISOR" || s === "EXAM-COMMITTEE";
  };
  const isAutoLibur_ = (l) =>
    String(l.notes).includes("Auto-generated: Libur Bonus");
  const examLogs = userLogs.filter((l) => isExamLog_(l));
  const manualLogs = userLogs.filter((l) => !isAutoLibur_(l) && !isExamLog_(l));
  const autoLiburLogs = userLogs.filter(
    (l) => isAutoLibur_(l) && !isExamLog_(l),
  );
  let manualJTM = 0;
  const jtmAdjMap = _jtmBuildKbmAdjustmentMap_();
  const piketAdjMap = _jtmBuildPiketAdjustmentMap_();
  manualLogs.forEach((l) => {
    const _sid = String(l.schedule_id);
    if (_sid === "PICKET-DUTY" || _sid === "PIKET") manualJTM += _jtmResolvePiketLogJtm_(piketAdjMap, l.user_id, l.date, 4);
    else if (_sid === "CEREMONY-DUTY" || _sid === "UPACARA") manualJTM += 5;
    else if (_sid === "PARTIAL-SUB-KBM" || _sid === "PARTIAL-SUB-EXAM")
      manualJTM += Number(l.jtm_val || 0);
    else {
      const sched = snapshot.schedulesById[_sid];
      if (sched)
        manualJTM += _jtmResolveKbmLogJtm_(
          jtmAdjMap,
          _sid,
          l.date,
          Number(sched.jtm_val || 0),
        );
      else manualJTM += Number(l.jtm_val || 0);
    }
  });
  let examSupJTM = examLogs
    .filter((l) => l.schedule_id === "EXAM-SUPERVISOR")
    .reduce((sum, l) => sum + Number(l.jtm_val || 0), 0);
  let examComJTM = examLogs
    .filter((l) => l.schedule_id === "EXAM-COMMITTEE")
    .reduce((sum, l) => sum + Number(l.jtm_val || 0), 0);
  let autoLiburJTM = autoLiburLogs.reduce(
    (sum, l) => sum + Number(l.jtm_val || 0),
    0,
  );
  let estimatedRemainingJTM = 0;
  const isCurrentMonth =
    targetMonth === now.getMonth() && targetYear === now.getFullYear();
  if (isCurrentMonth) {
    estimatedRemainingJTM = _calculateMonthlyBonusFast(
      user.id,
      targetMonth,
      targetYear,
      snapshot,
    );
  }
  const finalJTM = manualJTM + autoLiburJTM + examSupJTM + examComJTM;
  const displayJTM = isCurrentMonth
    ? finalJTM + estimatedRemainingJTM
    : finalJTM;
  const userAllowances = snapshot.allowancesByUserId[user.id] || [];
  const totalAllowanceAmount = userAllowances.reduce(
    (sum, a) => sum + Number(a.amount || 0),
    0,
  );
  const allowanceDetails = userAllowances.map((a) => ({
    desc: a.duty_name,
    qty: 1,
    rate: formatRupiah(a.amount),
    amount: formatRupiah(a.amount),
  }));
  const honorFromJTM = displayJTM * baseSalary;
  const schoolName = snapshot.config["app_school_name"] || "MTs Nurul Falah";
  const kepalaSekolah =
    snapshot.config["kepala_sekolah"] || "[Nama Kepala Sekolah]";
  const months = [
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
  const currentPeriodId = `${months[targetMonth]} ${targetYear}`;
  const historyRecord = snapshot.honorHistory.find(
    (h) =>
      String(h.user_id) === String(user.id) &&
      formatPeriode(h.periode) === currentPeriodId,
  );
  if (historyRecord && !isCurrentMonth) {
    let histDetails = [];
    let histTransportTotal = 0;
    try {
      const parsed = JSON.parse(historyRecord.details_json || "{}");
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
          histDetails = parsed.allowances.map((item) => ({
            desc: String(item.duty_name),
            qty: 1,
            rate: formatRupiah(Number(item.amount || 0)),
            amount: formatRupiah(Number(item.amount || 0)),
          }));
        } else {
          const storedNames = parsed.allowances.map((n) =>
            String(n).trim().toLowerCase(),
          );
          histDetails = userAllowances
            .filter((a) =>
              storedNames.includes(
                String(a.duty_name || "")
                  .trim()
                  .toLowerCase(),
              ),
            )
            .map((a) => ({
              desc: a.duty_name,
              qty: 1,
              rate: formatRupiah(a.amount),
              amount: formatRupiah(a.amount),
            }));
        }
      } else if (parsed.allow && Number(parsed.allow) > 0) {
        histDetails.push({
          desc: "Tunjangan Tambahan",
          qty: 1,
          rate: "-",
          amount: formatRupiah(parsed.allow),
        });
      }
      if (parsed.transport_total && Number(parsed.transport_total) > 0) {
        histTransportTotal = Math.round(Number(parsed.transport_total));
      }
    } catch (e) {}
    const histJtm = Number(historyRecord.total_jtm);
    const histTotal = Number(historyRecord.total_honor);
    let histEventJtm = 0;
    let histKbmJtm = histJtm;
    try {
      const parsedForEvent = JSON.parse(historyRecord.details_json || "{}");
      if (Number(parsedForEvent.event_jtm || 0) > 0) {
        histEventJtm = Number(parsedForEvent.event_jtm);
        histKbmJtm = histJtm - histEventJtm;
      }
    } catch (e) {}
    const honorJtm = histKbmJtm * baseSalary;
    if (histDetails.length === 0) {
      histDetails = userAllowances.map((a) => ({
        desc: a.duty_name,
        qty: 1,
        rate: formatRupiah(a.amount),
        amount: formatRupiah(a.amount),
      }));
    }
    if (histTransportTotal > 0) {
      histDetails.push({
        desc: "Tunjangan Transportasi",
        qty: 1,
        rate: "-",
        amount: formatRupiah(histTransportTotal),
      });
    }
    const eventJtmDetail = histEventJtm > 0 ? [{
      desc: "JTM Kehadiran Acara",
      qty: histEventJtm,
      rate: formatRupiah(baseSalary),
      amount: formatRupiah(histEventJtm * baseSalary),
      note: "Akumulasi JTM acara/kegiatan",
    }] : [];
    return {
      status: "success",
      month: targetMonth,
      year: targetYear,
      periode: currentPeriodId,
      school_name: schoolName,
      kepala_sekolah: kepalaSekolah,
      nip: user.nip || "-",
      breakdown: {
        manual_jtm: histKbmJtm,
        auto_libur_jtm: 0,
        estimated_remaining: 0,
        event_jtm: histEventJtm,
      },
      details: [
        {
          desc: "JTM Mengajar & Tugas",
          qty: histKbmJtm,
          rate: formatRupiah(baseSalary),
          amount: formatRupiah(honorJtm),
        },
        ...eventJtmDetail,
        ...histDetails,
      ],
      total: histTotal,
      formatted_total: formatRupiah(histTotal),
      trx_id: historyRecord.trx_id || null,
      is_estimate: false,
    };
  }
  var firstDayEvent =
    targetYear + "-" + String(targetMonth + 1).padStart(2, "0") + "-01";
  var lastDayEventDate = new Date(targetYear, targetMonth + 1, 0);
  var lastDayEvent =
    targetYear +
    "-" +
    String(targetMonth + 1).padStart(2, "0") +
    "-" +
    String(lastDayEventDate.getDate()).padStart(2, "0");
  var eventAttAll = getData(EVENT_SHEET.ATTENDANCE);
  var liveEventJtm = _computeEventJtm(
    eventAttAll.filter(function (r) {
      return (
        String(r.user_id || "") === String(user.id) &&
        String(r.date || "") >= firstDayEvent &&
        String(r.date || "") <= lastDayEvent
      );
    }),
  );
  const transportEnabled =
    String(snapshot.config["transport_allowance_enabled"] || "false")
      .toLowerCase()
      .trim() === "true";
  let liveTransportTotal = 0;
  if (transportEnabled) {
    liveTransportTotal = _getMonthlyTransportSumFast(
      String(user.id),
      currentPeriodId,
      snapshot
    );
  }
  const grandTotal =
    honorFromJTM +
    liveEventJtm * baseSalary +
    totalAllowanceAmount +
    liveTransportTotal;
  return {
    status: "success",
    month: targetMonth,
    year: targetYear,
    periode: currentPeriodId,
    school_name: schoolName,
    kepala_sekolah: kepalaSekolah,
    nip: user.nip || "-",
    breakdown: {
      manual_jtm: manualJTM,
      auto_libur_jtm: autoLiburJTM,
      exam_sup_jtm: examSupJTM,
      exam_com_jtm: examComJTM,
      estimated_remaining: isCurrentMonth ? estimatedRemainingJTM : 0,
      event_jtm: liveEventJtm,
    },
    details: [
      {
        desc: "JTM Mengajar & Tugas (Manual)",
        qty: manualJTM,
        rate: formatRupiah(baseSalary),
        amount: formatRupiah(manualJTM * baseSalary),
      },
      ...(autoLiburJTM > 0
        ? [
            {
              desc: "Bonus JTM Hari Libur (Auto)",
              qty: autoLiburJTM,
              rate: formatRupiah(baseSalary),
              amount: formatRupiah(autoLiburJTM * baseSalary),
              note: "Otomatis dari sistem",
            },
          ]
        : []),
      ...(examSupJTM > 0
        ? [
            {
              desc: "JTM Pengawas Ujian",
              qty: examSupJTM,
              rate: formatRupiah(baseSalary),
              amount: formatRupiah(examSupJTM * baseSalary),
            },
          ]
        : []),
      ...(examComJTM > 0
        ? [
            {
              desc: "JTM Panitia Ujian",
              qty: examComJTM,
              rate: formatRupiah(baseSalary),
              amount: formatRupiah(examComJTM * baseSalary),
            },
          ]
        : []),
      ...(isCurrentMonth && estimatedRemainingJTM > 0
        ? [
            {
              desc: "Estimasi Bonus Libur Mendatang",
              qty: estimatedRemainingJTM,
              rate: formatRupiah(baseSalary),
              amount: formatRupiah(estimatedRemainingJTM * baseSalary),
              note: "Prediksi berdasarkan kalender",
            },
          ]
        : []),
      {
        desc: "JTM Kehadiran Acara",
        qty: liveEventJtm,
        rate: formatRupiah(baseSalary),
        amount: formatRupiah(liveEventJtm * baseSalary),
        note: "Akumulasi JTM acara/kegiatan",
      },
      ...allowanceDetails,
      ...(liveTransportTotal > 0
        ? [
            {
              desc: "Tunjangan Transportasi",
              qty: 1,
              rate: formatRupiah(liveTransportTotal),
              amount: formatRupiah(liveTransportTotal),
              note: "Akumulasi tgl 1 s/d hari ini",
            },
          ]
        : []),
    ],
    total: grandTotal,
    formatted_total: formatRupiah(grandTotal),
    trx_id: null,
    is_estimate: isCurrentMonth && estimatedRemainingJTM > 0,
  };
}

function _calculateMonthlyBonusFast(userId, month, year, snapshot) {
  const now = new Date();
  const cutoffDate = new Date();
  cutoffDate.setHours(23, 59, 59, 999);
  const holidays = snapshot.holidays.filter((h) => {
    const d = new Date(h.date);
    return d.getMonth() === month && d.getFullYear() === year &&
      String(h.is_holiday).toLowerCase() === "true";
  });
  const userLogs = snapshot.logsByUserId[userId] || [];
  const autoLogDates = new Set();
  userLogs.forEach((l) => {
    if (String(l.notes).includes("Auto-generated: Libur Bonus")) {
      const d = new Date(l.date);
      autoLogDates.add(Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd"));
    }
  });
  // Filter jadwal: hanya user yang bersangkutan DAN semester/tahun pelajaran aktif
  const activeTP = (snapshot.config && snapshot.config["tahun_pelajaran"]) || "";
  const activeSem = (snapshot.config && snapshot.config["semester"]) || "";
  const userSchedules = snapshot.schedules.filter((s) => {
    if (String(s.user_id) !== String(userId)) return false;
    const sTP = s.tahun_pelajaran || activeTP;
    const sSem = s.semester || activeSem;
    return sTP === activeTP && sSem === activeSem;
  });
  const schedulesByDay = {};
  userSchedules.forEach((s) => {
    schedulesByDay[s.day_index] = s;
  });
  let totalBonus = 0;
  holidays.forEach((h) => {
    const dateStr = Utilities.formatDate(
      new Date(h.date),
      "Asia/Jakarta",
      "yyyy-MM-dd",
    );
    if (autoLogDates.has(dateStr)) return;
    const dayIdx = new Date(h.date).getDay();
    const sched = schedulesByDay[dayIdx];
    if (sched) totalBonus += Number(sched.jtm_val);
  });
  return totalBonus;
}

function getAdminHonorRecap(token, targetMonth, targetYear) {
  const user = verifySession(token);
  if (!user || String(user.role).toLowerCase() !== "admin")
    return { status: "error" };
  const now = new Date();
  const useMonth =
    targetMonth !== undefined && targetMonth !== null && targetMonth !== ""
      ? parseInt(targetMonth)
      : now.getMonth();
  const useYear =
    targetYear !== undefined && targetYear !== null && targetYear !== ""
      ? parseInt(targetYear)
      : now.getFullYear();
  const months = [
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
  const periodId = `${months[useMonth]} ${useYear}`;
  const snapshot = _getDataSnapshot();
  const finalizedPeriod = snapshot.honorHistory.filter(
    (h) => h.periode === periodId,
  );
  const isFinalized = finalizedPeriod.length > 0;
  const teachers = snapshot.users.filter(
    (u) => String(u.role).toLowerCase() === "guru",
  );
  const baseSalary = Number(snapshot.config["base_salary"] || 0);
  if (isFinalized) {
    let grandTotal = 0;
    const data = finalizedPeriod.map((h) => {
      const u = snapshot.usersById[h.user_id];
      grandTotal += Number(h.total_honor);
      let details = {};
      try {
        if (h.details_json) details = JSON.parse(h.details_json);
      } catch (e) {}
      return {
        id: h.user_id,
        nama: u ? u.full_name : "Unknown",
        nip: u ? u.nip || "-" : "-",
        total_jtm: h.total_jtm,
        nominal_tugas_tambahan: formatRupiah(details.allow || 0),
        nominal_transport: formatRupiah(details.transport_total || 0),
        nominal: formatRupiah(h.total_honor),
        trx_id: h.trx_id,
      };
    });
    return {
      status: "success",
      periode: periodId,
      grand_total: formatRupiah(grandTotal),
      is_finalized: true,
      data: data,
    };
  }
  let grandTotalAll = 0;
  const jtmAdjMap = _jtmBuildKbmAdjustmentMap_();
  const piketAdjMap = _jtmBuildPiketAdjustmentMap_();
  const firstDayEvent =
    useYear + "-" + String(useMonth + 1).padStart(2, "0") + "-01";
  const lastDayEventDate = new Date(useYear, useMonth + 1, 0);
  const lastDayEvent =
    useYear +
    "-" +
    String(useMonth + 1).padStart(2, "0") +
    "-" +
    String(lastDayEventDate.getDate()).padStart(2, "0");
  const eventAttAll = snapshot.eventAttendance;
  const transportEnabled =
    String(snapshot.config["transport_allowance_enabled"] || "false")
      .toLowerCase()
      .trim() === "true";
  const data = teachers.map((u) => {
    const logs = (snapshot.logsByUserId[u.id] || []).filter((l) => {
      const d = new Date(l.date);
      return d.getMonth() === useMonth && d.getFullYear() === useYear;
    });
    let totalJtm = 0;
    logs.forEach((l) => {
      const isAutoLibur = String(l.notes).includes(
        "Auto-generated: Libur Bonus",
      );
      if (isAutoLibur) return;
      const _sid = String(l.schedule_id);
      if (_sid === "PICKET-DUTY" || _sid === "PIKET") totalJtm += _jtmResolvePiketLogJtm_(piketAdjMap, l.user_id, l.date, 4);
      else if (_sid === "CEREMONY-DUTY" || _sid === "UPACARA") totalJtm += 5;
      else if (_sid === "EXAM-SUPERVISOR" || _sid === "EXAM-COMMITTEE")
        totalJtm += Number(l.jtm_val || 0);
      else if (_sid === "PARTIAL-SUB-KBM" || _sid === "PARTIAL-SUB-EXAM")
        totalJtm += Number(l.jtm_val || 0);
      else {
        const s = snapshot.schedulesById[_sid];
        if (s)
          totalJtm += _jtmResolveKbmLogJtm_(
            jtmAdjMap,
            _sid,
            l.date,
            Number(s.jtm_val || 0),
          );
        else totalJtm += Number(l.jtm_val || 0);
      }
    });
    const bonusDetail = _calculateMonthlyBonusFast_Detailed(
      u.id,
      useMonth,
      useYear,
      snapshot,
    );
    const totalBonus = bonusDetail.auto + bonusDetail.estimated;
    const eventJtm = _computeEventJtm(
      eventAttAll.filter(
        (r) =>
          String(r.user_id) === String(u.id) &&
          r.date >= firstDayEvent &&
          r.date <= lastDayEvent,
      ),
    );
    const finalJtm = totalJtm + totalBonus + eventJtm;
    const userAllowances = snapshot.allowancesByUserId[u.id] || [];
    const totalAllow = userAllowances.reduce(
      (sum, a) => sum + Number(a.amount),
      0,
    );
    let transportTotal = 0;
    if (transportEnabled) {
      const transportDesimal = _getMonthlyTransportSumFast(String(u.id), periodId, snapshot);
      transportTotal = _applyThreePointRounding(transportDesimal);
    }
    const nominal = finalJtm * baseSalary + totalAllow + transportTotal;
    grandTotalAll += nominal;
    return {
      id: u.id,
      nama: u.full_name,
      nip: u.nip || "-",
      total_jtm: finalJtm,
      nominal_tugas_tambahan: formatRupiah(totalAllow),
      nominal_transport: formatRupiah(transportTotal),
      nominal: formatRupiah(nominal),
    };
  });
  return {
    status: "success",
    periode: periodId,
    grand_total: formatRupiah(grandTotalAll),
    is_finalized: false,
    data: data,
  };
}

function getHonorariumHistoryGuru(token) {
  try {
    const user = verifySession(token);
    if (!user) {
      return {
        status: "error",
        message: "Sesi tidak valid. Silakan login kembali.",
      };
    }
    const snap = _getDataSnapshot();
    const allHistory = snap.honorHistory;
    if (!allHistory || !Array.isArray(allHistory)) {
      return { status: "success", data: [] };
    }
    const filtered = allHistory.filter(
      (h) => h && String(h.user_id) === String(user.id),
    );
    filtered.sort((a, b) => {
      const keyA = getPeriodeSortKey(a && a.periode);
      const keyB = getPeriodeSortKey(b && b.periode);
      return keyB - keyA;
    });
    const data = filtered
      .map((h) => {
        if (!h) return null;
        const periodeFormatted = formatPeriode(h.periode);
        let tanggalStr = "";
        if (h.tanggal_simpan) {
          if (h.tanggal_simpan instanceof Date) {
            try {
              tanggalStr = Utilities.formatDate(
                h.tanggal_simpan,
                Session.getScriptTimeZone(),
                "dd/MM/yyyy",
              );
            } catch (e) {
              tanggalStr = String(h.tanggal_simpan);
            }
          } else {
            tanggalStr = String(h.tanggal_simpan);
          }
        }
        let details = {};
        try {
          details = JSON.parse(h.details_json || "{}");
        } catch (e) {}
        return {
          id: String(h.id || ""),
          periode: String(periodeFormatted),
          total_jtm: Number(h.total_jtm) || 0,
          jtm_acara: Number(details.event_jtm || 0),
          nominal_tugas_tambahan: formatRupiah(details.allow || 0),
          nominal_transport: formatRupiah(details.transport_total || 0),
          nominal: formatRupiah(h.total_honor),
          trx_id: String(h.trx_id || "-"),
          tanggal_simpan: tanggalStr,
        };
      })
      .filter((item) => item !== null);
    return { status: "success", data: data };
  } catch (e) {
    console.error("Error in getHonorariumHistoryGuru: " + e.toString());
    return {
      status: "error",
      message: "Terjadi kesalahan server: " + e.toString(),
    };
  }
}

function getHonorariumHistoryAdmin(token) {
  try {
    const user = verifySession(token);
    if (!user || String(user.role).toLowerCase() !== "admin") {
      return {
        status: "error",
        message: "Unauthorized: Hanya admin yang dapat mengakses.",
      };
    }
    const snap = _getDataSnapshot();
    const allHistory = snap.honorHistory;
    const allUsers = snap.users;
    if (!allHistory || !Array.isArray(allHistory)) {
      return { status: "success", data: [], periods: [] };
    }
    const data = [];
    for (let i = 0; i < allHistory.length; i++) {
      const h = allHistory[i];
      if (!h) continue;
      const u = (allUsers || []).find(
        (usr) => usr && String(usr.id) === String(h.user_id),
      );
      const periodeFormatted = formatPeriode(h.periode);
      let tanggalStr = "";
      if (h.tanggal_simpan) {
        if (h.tanggal_simpan instanceof Date) {
          try {
            tanggalStr = Utilities.formatDate(
              h.tanggal_simpan,
              Session.getScriptTimeZone(),
              "dd/MM/yyyy",
            );
          } catch (e) {
            tanggalStr = String(h.tanggal_simpan);
          }
        } else {
          tanggalStr = String(h.tanggal_simpan);
        }
      }
      let details = {};
      try {
        details = JSON.parse(h.details_json || "{}");
      } catch (e) {}
      data.push({
        id: String(h.id || ""),
        periode: String(periodeFormatted),
        guru_name: u ? String(u.full_name || "Unknown") : "Unknown",
        user_id: String(h.user_id || ""),
        total_jtm: Number(h.total_jtm) || 0,
        nominal_tugas_tambahan: formatRupiah(details.allow || 0),
        nominal_transport: formatRupiah(details.transport_total || 0),
        nominal: formatRupiah(h.total_honor),
        trx_id: String(h.trx_id || "-"),
        tanggal_simpan: tanggalStr,
      });
    }
    data.sort((a, b) => {
      const keyA = getPeriodeSortKey(a.periode);
      const keyB = getPeriodeSortKey(b.periode);
      if (keyA !== keyB) return keyB - keyA;
      return (a.guru_name || "").localeCompare(b.guru_name || "");
    });
    const periodSet = new Set();
    for (let i = 0; i < data.length; i++) {
      const p = data[i].periode;
      if (p && p !== "Unknown") periodSet.add(p);
    }
    const periods = Array.from(periodSet);
    periods.sort((a, b) => getPeriodeSortKey(b) - getPeriodeSortKey(a));
    return { status: "success", data: data, periods: periods };
  } catch (e) {
    console.error("Error in getHonorariumHistoryAdmin: " + e.toString());
    return {
      status: "error",
      message: "Terjadi kesalahan server: " + e.toString(),
    };
  }
}
