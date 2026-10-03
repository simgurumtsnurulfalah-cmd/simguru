// SiM-Guru — JTM validation/authorization helpers
function _jtmValidateAdjusted(scheduledJtm, rawAdjusted) {
  var invalid = {
    ok: false,
    code: JTM_ERROR_CODES.ADJUSTED_INVALID,
    message:
      "Adjusted JTM harus berupa bilangan bulat antara 0 dan " +
      scheduledJtm +
      ".",
  };
  var value;
  if (typeof rawAdjusted === "number") {
    value = rawAdjusted;
  } else if (typeof rawAdjusted === "string") {
    var trimmed = rawAdjusted.trim();
    if (trimmed === "") return invalid;
    value = Number(trimmed);
  } else {
    return invalid;
  }
  if (typeof value !== "number" || isNaN(value) || !isFinite(value))
    return invalid;
  if (!Number.isInteger(value) || value < 0) return invalid;
  if (value > scheduledJtm) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.ADJUSTED_EXCEEDS_SCHEDULED,
      message:
        "Adjusted JTM tidak boleh melebihi Scheduled JTM (" +
        scheduledJtm +
        ").",
    };
  }
  return { ok: true, value: value };
}

function _jtmComputeDifference(scheduledJtm, adjustedJtm) {
  var difference = Math.round(scheduledJtm - adjustedJtm);
  return difference < 0 ? 0 : difference;
}

function _jtmValidateReason(adjustedJtm, scheduledJtm, rawReason) {
  var trimmed = "";
  if (typeof rawReason === "string") {
    trimmed = rawReason.trim();
  } else if (
    rawReason !== null &&
    rawReason !== undefined &&
    typeof rawReason !== "object"
  ) {
    trimmed = String(rawReason).trim();
  }
  if (!(adjustedJtm < scheduledJtm)) return { ok: true, reason: trimmed };
  if (trimmed.length === 0) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.REASON_REQUIRED,
      message: "Alasan penyesuaian wajib diisi ketika JTM dikurangi.",
    };
  }
  if (trimmed.length > JTM_REASON_MAX_LENGTH) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.REASON_TOO_LONG,
      message:
        "Alasan penyesuaian tidak boleh melebihi " +
        JTM_REASON_MAX_LENGTH +
        " karakter.",
    };
  }
  return { ok: true, reason: trimmed };
}

function _jtmValidateAllocation(
  rawAllocation,
  originalUserId,
  substituteUserId,
  isSubstituteAdmin,
) {
  var invalid = {
    ok: false,
    code: JTM_ERROR_CODES.ALLOCATION_INVALID,
    message: "Alokasi JTM harus berupa angka lebih besar dari 0.",
  };
  var value;
  if (typeof rawAllocation === "number") {
    value = rawAllocation;
  } else if (typeof rawAllocation === "string") {
    var trimmed = rawAllocation.trim();
    if (trimmed === "") return invalid;
    value = Number(trimmed);
  } else {
    return invalid;
  }
  if (typeof value !== "number" || isNaN(value) || !isFinite(value))
    return invalid;
  if (value <= 0) return invalid;
  if (substituteUserId === originalUserId) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.SUBSTITUTE_IS_ORIGINAL,
      message:
        "Guru pengganti tidak boleh sama dengan guru/pengawas asli untuk kegiatan ini.",
    };
  }
  if (isSubstituteAdmin) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.SUBSTITUTE_IS_ADMIN,
      message: "Guru pengganti tidak boleh seorang Admin.",
    };
  }
  return { ok: true, value: value };
}

function _jtmCheckConservation(
  scheduledJtm,
  adjustedJtm,
  existingAllocations,
  newAllocation,
) {
  var sumExisting = 0;
  if (existingAllocations && typeof existingAllocations.length === "number") {
    for (var i = 0; i < existingAllocations.length; i++) {
      var entry = Number(existingAllocations[i]);
      if (!isNaN(entry) && isFinite(entry)) sumExisting += entry;
    }
  }
  var addition = Number(newAllocation);
  if (isNaN(addition) || !isFinite(addition)) addition = 0;
  var base = Number(adjustedJtm);
  if (isNaN(base) || !isFinite(base)) base = 0;
  var total = base + sumExisting + addition;
  if (total <= scheduledJtm) return { ok: true, total: total };
  return {
    ok: false,
    code: JTM_ERROR_CODES.CONSERVATION_VIOLATION,
    total: total,
    message:
      "Total JTM (" +
      total +
      ") melebihi Scheduled JTM (" +
      scheduledJtm +
      "). Alokasi ditolak agar tidak melampaui batas.",
  };
}

function _jtmAuthorizeKbm(user, isPicketToday, todayDate, occurrenceDate) {
  var role =
    user && typeof user.role === "string" ? user.role.toLowerCase() : "";
  var isAdmin = role === "admin";
  var isAuthorizedRole = isAdmin || !!isPicketToday;
  if (!isAuthorizedRole) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.NOT_AUTHORIZED,
      message:
        "Anda tidak berwenang melakukan penyesuaian JTM untuk kegiatan KBM ini.",
    };
  }
  if (todayDate !== occurrenceDate) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.NOT_DAY_OF_OCCURRENCE,
      message:
        "Penyesuaian JTM hanya dapat dilakukan pada hari pelaksanaan (hari H) kegiatan.",
    };
  }
  return { ok: true };
}

function _jtmAuthorizeExam(authResult, todayDate, sessionDate) {
  var isAuthorizedRole = !!(authResult && authResult.ok);
  if (!isAuthorizedRole) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.NOT_AUTHORIZED,
      message:
        "Anda tidak berwenang melakukan penyesuaian JTM untuk pengawas ujian ini.",
    };
  }
  if (todayDate !== sessionDate) {
    return {
      ok: false,
      code: JTM_ERROR_CODES.NOT_DAY_OF_OCCURRENCE,
      message:
        "Penyesuaian JTM hanya dapat dilakukan pada hari pelaksanaan (hari H) sesi ujian.",
    };
  }
  return { ok: true };
}
