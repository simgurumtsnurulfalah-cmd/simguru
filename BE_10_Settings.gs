// SiM-Guru — BE_10_Settings.gs
function saveSystemConfig(token, data) {
  try {
    const user = verifySession(token);
    if (!user || user.role !== "admin")
      return { status: "error", message: "Akses hanya untuk Admin" };
    const sheet = getSheet("Config");
    const existingData = sheet.getDataRange().getValues();
    const updateKey = (key, val) => {
      let rowIndex = -1;
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][0] === key) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex > 0) {
        const cell = sheet.getRange(rowIndex, 2);
        if (key === "app_version" || key === "maintenance_start" || key === "maintenance_end") cell.setNumberFormat("@");
        cell.setValue(val);
      } else {
        sheet.appendRow([key, ""]);
        const cell = sheet.getRange(sheet.getLastRow(), 2);
        if (key === "app_version" || key === "maintenance_start" || key === "maintenance_end") cell.setNumberFormat("@");
        cell.setValue(val);
      }
    };
    for (const [key, value] of Object.entries(data)) {
      updateKey(key, value);
    }
    try {
      _invalidateDataSnapshot();
    } catch (_) {}
    return { status: "success" };
  } catch (e) {
    console.error("saveSystemConfig error: " + e);
    return { status: "error", message: "Terjadi kesalahan server. Coba lagi." };
  }
}

function saveSubject(token, data) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const nameTrimmed = (data.name || "").trim();
  if (!nameTrimmed)
    return { status: "error", message: "Nama mapel tidak boleh kosong." };
  const sheet = getSheet("Subjects");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toLowerCase() === nameTrimmed.toLowerCase()) {
      if (!data.id || String(rows[i][0]) !== String(data.id)) {
        return {
          status: "error",
          message: "Mata pelajaran dengan nama tersebut sudah ada.",
        };
      }
    }
  }
  if (data.id) {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2, 1, 1).setValue(nameTrimmed);
        found = true;
        break;
      }
    }
    if (!found)
      return { status: "error", message: "Data Mapel tidak ditemukan." };
  } else {
    const newId = generateId("SBJ");
    sheet.appendRow([newId, nameTrimmed]);
  }
  return { status: "success" };
}

function deleteSubject(token, id) {
  return deleteRowById(token, "Subjects", id);
}

function saveAllowance(token, data) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const sheet = getSheet("Allowances");
  if (data.id) {
    const rows = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2).setValue(data.duty_name);
        sheet.getRange(i + 1, 3).setValue(data.user_id);
        sheet.getRange(i + 1, 4).setValue(data.amount);
        found = true;
        break;
      }
    }
    if (!found)
      return { status: "error", message: "Tunjangan tidak ditemukan." };
  } else {
    const newId = generateId("ALW");
    sheet.appendRow([newId, data.duty_name, data.user_id, data.amount]);
  }
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  return { status: "success" };
}

function deleteAllowance(token, id) {
  return deleteRowById(token, "Allowances", id);
}

function addHoliday(token, date, desc, isHoliday, id) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const sheet = getSheet("Academic_Calendar");
  const isHolStr = isHoliday ? "True" : "False";
  if (id) {
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        let existingDate = String(data[i][1]).trim();
        let todayStr = Utilities.formatDate(
          new Date(),
          "Asia/Jakarta",
          "yyyy-MM-dd",
        );
        if (existingDate < todayStr) {
          return {
            status: "error",
            message: "Hari libur yang sudah lewat tidak dapat diubah",
          };
        }
        sheet.getRange(i + 1, 2).setValue(date);
        sheet.getRange(i + 1, 3).setValue(desc);
        sheet.getRange(i + 1, 4).setValue(isHolStr);
        found = true;
        break;
      }
    }
    if (!found) return { status: "error", message: "Data tidak ditemukan" };
  } else {
    const newId = generateId("HOL");
    sheet.appendRow([newId, date, desc, isHolStr]);
  }
  try {
    _invalidateDataSnapshot();
  } catch (_) {}
  return { status: "success" };
}

function deleteHoliday(token, id) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const sheet = getSheet("Academic_Calendar");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      let existingDate = String(data[i][1]).trim();
      let todayStr = Utilities.formatDate(
        new Date(),
        "Asia/Jakarta",
        "yyyy-MM-dd",
      );
      if (existingDate < todayStr) {
        return {
          status: "error",
          message: "Hari libur yang sudah lewat tidak dapat dihapus",
        };
      }
      sheet.deleteRow(i + 1);
      try {
        _invalidateDataSnapshot();
      } catch (_) {}
      return { status: "success" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}

function deleteRowById(token, sheetName, id) {
  const user = verifySession(token);
  if (!user || user.role !== "admin")
    return { status: "error", message: "Unauthorized" };
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      try {
        _invalidateDataSnapshot();
      } catch (_) {}
      return { status: "success" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}
