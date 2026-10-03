// SiM-Guru — BE_08_IdentityExternalUsers.gs
function _checkPasswordComplexity(password, fullName, username) {
  if (password.length < 8) {
    return { valid: false, message: "Password baru minimal harus 8 karakter." };
  }
  if (!/[A-Z]/.test(password)) {
    return { valid: false, message: "Password harus mengandung huruf besar." };
  }
  if (!/[a-z]/.test(password)) {
    return { valid: false, message: "Password harus mengandung huruf kecil." };
  }
  if (!/[0-9]/.test(password)) {
    return { valid: false, message: "Password harus mengandung angka." };
  }
  if (!/[!@#$%^&*(),.?":{}|<>]/.test(password)) {
    return {
      valid: false,
      message: "Password harus mengandung karakter khusus (simbol).",
    };
  }
  const lowerPass = password.toLowerCase();
  if (username && lowerPass.includes(username.toLowerCase())) {
    return {
      valid: false,
      message: "Password tidak boleh mengandung username Anda.",
    };
  }
  if (fullName) {
    const nameParts = fullName.split(" ").filter((part) => part.length > 2);
    for (let part of nameParts) {
      if (lowerPass.includes(part.toLowerCase())) {
        return {
          valid: false,
          message:
            "Password tidak boleh mengandung unsur nama Anda (" + part + ").",
        };
      }
    }
  }
  return { valid: true };
}

function _extUsersSheet_() {
  return SpreadsheetApp.openById(EXT_USERS_SS_ID).getSheetByName(
    EXT_USERS_SHEET,
  );
}

function _extFindUserRow_(username) {
  const sheet = _extUsersSheet_();
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(username).trim()) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function _extGetPassword_(username) {
  const found = _extFindUserRow_(username);
  return found ? String(found.row[2]) : null;
}

function _extSetPassword_(username, newPassword) {
  const sheet = _extUsersSheet_();
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(username).trim()) {
      sheet.getRange(i + 1, 3).setValue(newPassword);
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}

function _extAddUser_(username, password) {
  const sheet = _extUsersSheet_();
  if (!sheet) return false;
  sheet.appendRow(["", String(username), String(password)]);
  SpreadsheetApp.flush();
  return true;
}

function _extUpdateUsername_(oldUsername, newUsername) {
  const sheet = _extUsersSheet_();
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(oldUsername).trim()) {
      sheet.getRange(i + 1, 2).setValue(String(newUsername));
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}

function _extDeleteUser_(username) {
  const sheet = _extUsersSheet_();
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim() === String(username).trim()) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}
