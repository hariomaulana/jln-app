function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('JLN MANAGEMEN SYSTEM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * SETUP DATABASE LENGKAP
 * Update: Menambahkan kolom foto_ktp dan foto_rumah
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'users': ['id', 'username', 'password', 'nama', 'email', 'role', 'telp', 'alamat', 'created_at'],
    'customers': ['id', 'nama', 'username', 'password', 'alamat', 'lokasi_id', 'paket_id', 'sales_id', 'status', 'telp', 'foto_ktp', 'foto_rumah', 'created_at'],
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
    }
  }
  
  // Create default admin if users empty
  const userSheet = ss.getSheetByName('users');
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow([
      Utilities.getUuid(), 'admin', 'Admin2024!', 'Super Admin', 'admin@example.com', 'Admin', '08123456789', 'Server Room', new Date()
    ]);
  }
}

/**
 * AUTH SYSTEM
 */
function loginUser(formData) {
  const users = getSheetData('users');
  const found = users.find(u => u.username == formData.username && u.password == formData.password);
  
  if (found) {
    return { success: true, user: found };
  }
  
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
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

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
    // Process File Uploads inside data object
    for (let key in data) {
      if (data[key] && typeof data[key] === 'object' && data[key].isFile) {
        data[key] = uploadFileToDrive(data[key].data, data[key].name, data[key].mimeType);
      }
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
    sheet.appendRow(row);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function updateData(sheetName, id, updatedData) {
  try {
    // Process File Uploads
    for (let key in updatedData) {
      if (updatedData[key] && typeof updatedData[key] === 'object' && updatedData[key].isFile) {
        updatedData[key] = uploadFileToDrive(updatedData[key].data, updatedData[key].name, updatedData[key].mimeType);
      }
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
          // Hanya update jika ada data baru, jika string kosong (karena tidak upload file baru), jangan timpa
          if (updatedData.hasOwnProperty(h) && h !== 'id' && h !== 'created_at') {
             // Logic khusus file: jika user tidak upload file baru, jangan update kolom foto
             if (updatedData[h] !== "") {
                sheet.getRange(i + 1, colIdx + 1).setValue(updatedData[h]);
             }
          }
        });
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
