function doGet() {
  try { setupDatabase(); } catch (e) {}
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('JLN MANAGEMEN SYSTEM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * SETUP DATABASE LENGKAP
 * Update: foto_ktp, foto_rumah
 * Update: lat & lng untuk koordinat pelanggan
 * Update: catatan pelanggan
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'users': ['id', 'username', 'password', 'nama', 'email', 'role', 'telp', 'alamat', 'created_at'],
    'customers': ['id', 'nama', 'username', 'password', 'alamat', 'lat', 'lng', 'lokasi_id', 'paket_id', 'sales_id', 'status', 'telp', 'foto_ktp', 'foto_rumah', 'catatan', 'created_at'],
    'locations': ['id', 'nama_blok', 'keterangan', 'created_at'],
    'packages': ['id', 'nama_paket', 'harga', 'bandwidth', 'status_aktif', 'ketersediaan', 'created_at'],
    'tickets': ['id', 'pelanggan_id', 'teknisi_id', 'judul_laporan', 'status', 'keterangan_teknisi', 'created_at'],
    'announcements': ['id', 'target', 'judul', 'isi', 'created_at'],
    'tutorials': ['id', 'target', 'kategori', 'judul', 'link_isi', 'created_at'],
    'reports': ['id', 'user_id', 'jenis_laporan', 'jumlah', 'keterangan', 'created_at']
  };

  for (let name in sheets) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
      sheet.getRange(1, 1, 1, sheets[name].length).setFontWeight('bold').setBackground('#f3f3f3');
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(sheets[name]);
      sheet.getRange(1, 1, 1, sheets[name].length).setFontWeight('bold').setBackground('#f3f3f3');
    }
  }

  // Default admin
  const userSheet = ss.getSheetByName('users');
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow([
      Utilities.getUuid(), 'admin', 'Admin2024!', 'Super Admin', 'admin@example.com', 'Admin', '08123456789', 'Server Room', new Date()
    ]);
  }

  // Ensure columns aman untuk DB lama
  ensureColumns('customers', ['lat', 'lng', 'foto_ktp', 'foto_rumah', 'catatan', 'created_at']);
  ensureColumns('users', ['created_at']);
  ensureColumns('locations', ['created_at']);
  ensureColumns('packages', ['created_at']);
  ensureColumns('tickets', ['created_at']);
  ensureColumns('announcements', ['created_at']);
  ensureColumns('tutorials', ['created_at']);
  ensureColumns('reports', ['created_at']);

  // PAKSA kolom lat/lng jadi TEXT supaya titik tidak jadi pemisah ribuan
  ensureTextColumns('customers', ['lat', 'lng']);
}

function ensureColumns(sheetName, requiredHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return false;

  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());

  let changed = false;
  requiredHeaders.forEach(h => {
    if (!headers.includes(h)) {
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn()).setValue(h).setFontWeight('bold').setBackground('#f3f3f3');
      changed = true;
    }
  });
  return changed;
}

/**
 * Paksa kolom tertentu menjadi Plain text (@)
 */
function ensureTextColumns(sheetName, headersToText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());

  headersToText.forEach(h => {
    const idx = headers.indexOf(h);
    if (idx !== -1) {
      // set format text untuk seluruh kolom (mulai baris 2)
      const col = idx + 1;
      const maxRows = Math.max(2, sh.getMaxRows());
      sh.getRange(2, col, maxRows - 1, 1).setNumberFormat('@');
    }
  });
}

/**
 * Normalisasi koordinat:
 * - terima "106,815546" atau "106.815546" atau "106815546"
 * - keluaran selalu string desimal pakai titik, 6 digit
 */
function normalizeCoord(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).trim();
  if (!s) return '';

  // hilangkan spasi
  s = s.replace(/\s+/g, '');

  // jika ada koma, anggap itu desimal -> ubah ke titik
  // contoh: 106,815546 -> 106.815546
  if (s.includes(',') && !s.includes('.')) {
    s = s.replace(',', '.');
  }

  // jika tidak ada titik/koma sama sekali dan angka besar -> kemungkinan "106815546" (x 1e6)
  if (!s.includes('.') && !s.includes(',')) {
    const n = Number(s);
    if (!isNaN(n)) {
      // heuristik: lat harus <= 90, lng <= 180
      const abs = Math.abs(n);
      if (abs > 180 && abs < 1000000000) {
        // asumsikan * 1e6
        return (n / 1e6).toFixed(6);
      }
    }
  }

  const num = Number(s);
  if (isNaN(num)) return '';

  return num.toFixed(6);
}

/**
 * AUTH SYSTEM
 */
function loginUser(formData) {
  const users = getSheetData('users');
  const found = users.find(u => u.username == formData.username && u.password == formData.password);
  if (found) return { success: true, user: found };

  const custs = getSheetData('customers');
  const foundCust = custs.find(c => c.username == formData.username && c.password == formData.password);
  if (foundCust) {
    foundCust.role = 'Pelanggan';
    return { success: true, user: foundCust };
  }

  return { success: false, message: 'Username atau Password salah!' };
}

/**
 * FILE UPLOAD HANDLER (GOOGLE DRIVE)
 */
function uploadFileToDrive(base64Data, fileName, mimeType) {
  try {
    const folderName = "JLN_DATA_UPLOAD";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Error Upload: " + e.toString();
  }
}

/**
 * CORE CRUD FUNCTIONS
 */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      }
      obj[h] = val === undefined || val === null ? "" : val;
    });
    return obj;
  });
}

function saveData(sheetName, data) {
  try {
    // Process file uploads
    for (let key in data) {
      if (data[key] && typeof data[key] === 'object' && data[key].isFile) {
        data[key] = uploadFileToDrive(data[key].data, data[key].name, data[key].mimeType);
      }
    }

    // Khusus customers: paksa lat/lng sebagai TEXT + normalisasi
    if (sheetName === 'customers') {
      if (data.lat) data.lat = "'" + normalizeCoord(data.lat);
      if (data.lng) data.lng = "'" + normalizeCoord(data.lng);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    data.id = Utilities.getUuid();
    data.created_at = new Date();

    if (!data.status && sheetName === 'customers') data.status = 'Aktif';
    if (!data.status_aktif && sheetName === 'packages') data.status_aktif = 'Aktif';
    if (!data.status && sheetName === 'tickets') data.status = 'Open';

    const row = headers.map(h => data[h] || "");

    // Gunakan setValues (lebih aman untuk text format)
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, headers.length).setValues([row]);

    // Pastikan kolom lat/lng tetap text
    if (sheetName === 'customers') ensureTextColumns('customers', ['lat', 'lng']);

    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function updateData(sheetName, id, updatedData) {
  try {
    // Process file uploads
    for (let key in updatedData) {
      if (updatedData[key] && typeof updatedData[key] === 'object' && updatedData[key].isFile) {
        updatedData[key] = uploadFileToDrive(updatedData[key].data, updatedData[key].name, updatedData[key].mimeType);
      }
    }

    // Khusus customers: paksa lat/lng sebagai TEXT + normalisasi
    if (sheetName === 'customers') {
      if (updatedData.lat) updatedData.lat = "'" + normalizeCoord(updatedData.lat);
      if (updatedData.lng) updatedData.lng = "'" + normalizeCoord(updatedData.lng);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('id');
    if (idIdx === -1) return { success: false, message: 'ID column not found' };

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] == id) {
        headers.forEach((h, colIdx) => {
          if (updatedData.hasOwnProperty(h) && h !== 'id' && h !== 'created_at') {
            // jika kosong jangan timpa
            if (updatedData[h] !== "") {
              sheet.getRange(i + 1, colIdx + 1).setValue(updatedData[h]);
            }
          }
        });

        // Pastikan kolom lat/lng tetap text
        if (sheetName === 'customers') ensureTextColumns('customers', ['lat', 'lng']);

        return { success: true };
      }
    }

    return { success: false, message: 'Data ID not found' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteData(sheetName, id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    const idIdx = data[0].indexOf('id');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] == id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Data ID not found' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getReferenceData() {
  return {
    packages: getSheetData('packages'),
    sales: getSheetData('users').filter(u => u.role === 'Sales'),
    teknisi: getSheetData('users').filter(u => u.role === 'Teknisi'),
    locations: getSheetData('locations'),
    customers: getSheetData('customers')
  };
}

function getDashboardStats() {
  return {
    success: true,
    customers: getSheetData('customers').length,
    packages: getSheetData('packages').length,
    tickets: getSheetData('tickets').filter(t => t.status !== 'Selesai').length,
    sales: getSheetData('users').filter(u => u.role === 'Sales').length,
    income: getSheetData('reports').reduce((acc, curr) => acc + (parseInt(curr.jumlah) || 0), 0)
  };
}

/**
 * ✅ Jalankan SEKALI untuk memperbaiki data customers yang sudah keburu "hilang titik"
 * Cara pakai:
 * - Buka Apps Script -> pilih fungsi fixCoordinatesInCustomers -> Run
 */
function fixCoordinatesInCustomers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('customers');
  if (!sh || sh.getLastRow() <= 1) return;

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const latIdx = headers.indexOf('lat');
  const lngIdx = headers.indexOf('lng');
  if (latIdx === -1 || lngIdx === -1) return;

  const range = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol);
  const values = range.getValues();

  let changed = false;

  for (let r = 0; r < values.length; r++) {
    let lat = values[r][latIdx];
    let lng = values[r][lngIdx];

    const latNorm = normalizeCoord(lat);
    const lngNorm = normalizeCoord(lng);

    // kalau normalization menghasilkan valid, paksa text
    if (latNorm) values[r][latIdx] = latNorm;
    if (lngNorm) values[r][lngIdx] = lngNorm;

    if (latNorm || lngNorm) changed = true;
  }

  if (changed) {
    range.setValues(values);
    // paksa format text
    ensureTextColumns('customers', ['lat', 'lng']);
  }
}
