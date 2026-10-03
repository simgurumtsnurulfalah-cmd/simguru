// SiM-Guru — BE_05_CoreData.gs
function getDbId() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getId();
  } catch (e) {
    throw new Error(
      "Script tidak terhubung ke Spreadsheet. Pastikan script ini berada di dalam file Google Sheets (Container-bound) atau masukkan ID secara manual di Server_Database.gs",
    );
  }
}
let _cachedSS = null;

function getSheet(name) {
  if (!_cachedSS) {
    const id = getDbId();
    _cachedSS = SpreadsheetApp.openById(id);
  }
  let sheet = _cachedSS.getSheetByName(name);
  if (!sheet) {
    sheet = _cachedSS.insertSheet(name);
  }
  return sheet;
}

function getData(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0].map((h) =>
      String(h).toLowerCase().trim().replace(/\s+/g, "_"),
    );
    const rows = data.slice(1);
    var tz = Session.getScriptTimeZone();
    return rows.map((row) => {
      let obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (h.includes("date") && val instanceof Date) {
          obj[h] = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        } else if (h === "periode" && val instanceof Date) {
          obj[h] = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        } else if (
          (h === "time_in" ||
            h === "time_out" ||
            h === "sched_time_in" ||
            h === "sched_time_out" ||
            h === "leave_time" ||
            h === "return_time" ||
            h === "time_start" ||
            h === "time_end" ||
            h === "submitted_at") &&
          val instanceof Date
        ) {
          obj[h] = Utilities.formatDate(val, tz, "HH:mm");
        } else if (val instanceof Date) {
          obj[h] = val.toISOString();
        } else {
          obj[h] = val;
        }
      });
      return obj;
    });
  } catch (e) {
    console.error("Error in getData: " + e.toString());
    return [];
  }
}

function findData(sheetName, key, value) {
  const all = getData(sheetName);
  return all.find((item) => String(item[key]) === String(value));
}

function generateId(prefix) {
  return prefix + "-" + new Date().getTime();
}

function formatRupiah(num) {
  let nominal = Number(num);
  if (isNaN(nominal)) nominal = 0;
  return new Intl.NumberFormat("id-ID", {
    style: "currency",
    currency: "IDR",
    minimumFractionDigits: 0,
  }).format(nominal);
}

function formatDateIndo(dateInput) {
  if (!dateInput) return "-";
  const d = new Date(dateInput);
  if (isNaN(d.getTime())) return String(dateInput);
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
  const day = String(d.getDate()).padStart(2, "0");
  const month = months[d.getMonth()];
  const year = d.getFullYear();
  return `${day} ${month} ${year}`;
}
