/*******************************************************
 * JLN MANAGEMEN SYSTEM - Code.gs (FULL, PATCHED & FIXED)
 *
 * PERBAIKAN PENTING:
 * 1. DATE FIX: Google Script sering gagal mengirim data Tanggal ke HTML.
 * Solusi: Semua tanggal dikonversi jadi Teks (String) sebelum dikirim.
 * 2. READ FIX: Mencegah data hilang jika kolom 'id' kosong/header rusak.
 * 3. SCAN FIX: Pembacaan kolom lebih teliti agar tidak ada data terpotong.
 *******************************************************/

const TZ = 'Asia/Jakarta';

// Jika project standalone, set ID Spreadsheet DB di Script Properties via setDatabaseSpreadsheetId()
const DB_PROP_KEYS = ['JLN_DB_SPREADSHEET_ID', 'DB_SPREADSHEET_ID', 'SPREADSHEET_ID', 'DB_ID'];

const SESSION_PREFIX = 'JLN_SESS_';
const SESSION_TTL_MS = 12 * 60 * 60 * 1000; // 12 jam
const ROOT_UPLOAD_FOLDER_NAME = 'JLN_UPLOADS';

// Scan header per blok agar tidak membaca ribuan kolom kosong/formatting
const HEADER_SCAN_BLOCK = 50;

// Skema kolom (canonical)
const SCHEMAS = {
  users: [
    'id','username','password','nama','email','role','telp','alamat',
    'must_change_password','created_at','last_login_at','updated_at'
  ],
  packages: [
    'id','nama_paket','harga','bandwidth','status_aktif','ketersediaan','created_at','updated_at'
  ],
  locations: [
    'id','nama_blok','keterangan','created_at','updated_at'
  ],
  customers: [
    'id','nama','username','password','alamat','lokasi_id','paket_id',
    'sales_id','status','telp','foto_ktp','foto_rumah','lng','lat',
    'catatan','created_at','updated_at'
  ],
  tickets: [
    'id','judul_laporan','pelanggan_id','teknisi_id','status','keterangan_teknisi',
    'created_by','created_role','created_at','updated_at'
  ],
  ticket_logs: [
    'id','ticket_id','actor_id','actor_role','actor_name','message','attachment','created_at'
  ],
  announcements: [
    'id','judul','target','isi','created_by','created_at','updated_at'
  ],
  tutorials: [
    'id','judul','kategori','link_isi','created_by','created_at','updated_at'
  ],
  reports: [
    'id','jenis_laporan','jumlah','keterangan','created_by','created_at','updated_at'
  ]
};

/* =========================
 * WEB APP
 * ========================= */
function doGet(e) {
  ensureDatabase_();
  const t = HtmlService.createTemplateFromFile('Index');
  return t.evaluate()
    .setTitle('JLN MANAGEMEN SYSTEM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================
 * RUN ONCE HELPER
 * ========================= */
function RUN_ONCE__ENSURE_DB(){
  ensureDatabase_();
  return 'OK: ensureDatabase_ selesai.';
}

/* =========================
 * DB ID SETTER
 * ========================= */
function setDatabaseSpreadsheetId(id){
  PropertiesService.getScriptProperties().setProperty('JLN_DB_SPREADSHEET_ID', String(id || '').trim());
  return 'OK';
}

/* =========================
 * ENSURE DATABASE
 * ========================= */
function ensureDatabase_(){
  const ss = getDb_();

  Object.keys(SCHEMAS).forEach(name => {
    let sh = getSheet_(name, ss);

    const schema = SCHEMAS[name].map(canonKey_);

    let headerCol = getLastHeaderCol_(sh);
    let rawHeader = sh.getRange(1, 1, 1, headerCol).getValues()[0];
    let header = rawHeader.map(canonKey_);

    const isAllEmpty = header.every(h => !h);
    if (isAllEmpty){
      sh.getRange(1, 1, 1, schema.length).setValues([schema]);
      sh.setFrozenRows(1);
      autoResize_(sh, schema.length);
      return;
    }

    // Standardkan header
    for (let i = 0; i < header.length; i++){
      const can = header[i];
      const raw = String(rawHeader[i] === undefined || rawHeader[i] === null ? '' : rawHeader[i]).trim();
      if (can && raw !== can){
        sh.getRange(1, i+1).setValue(can);
      }
    }

    // reload header
    headerCol = getLastHeaderCol_(sh);
    rawHeader = sh.getRange(1, 1, 1, headerCol).getValues()[0];
    header = rawHeader.map(canonKey_);

    // Tambah kolom schema yang hilang
    const map = firstHeaderMap_(header);
    schema.forEach(col => {
      if (!(col in map)){
        sh.insertColumnAfter(headerCol);
        headerCol += 1;
        sh.getRange(1, headerCol).setValue(col);
        map[col] = headerCol;
      }
    });

    dedupeAndMergeColumns_(sh);
    sh.setFrozenRows(1);
    autoResize_(sh, getLastHeaderCol_(sh));
  });

  seedAdminIfEmpty_();
}

function dedupeAndMergeColumns_(sh){
  const headerCol = getLastHeaderCol_(sh);
  const headerRaw = sh.getRange(1, 1, 1, headerCol).getValues()[0];
  const header = headerRaw.map(canonKey_);

  const groups = {};
  header.forEach((k, idx) => {
    if (!k) return;
    if (!groups[k]) groups[k] = [];
    groups[k].push(idx + 1);
  });

  const lastRow = sh.getLastRow();
  const numRows = Math.max(lastRow - 1, 0);

  Object.keys(groups).forEach(key => {
    const cols = groups[key];
    if (cols.length <= 1) return;

    const primary = cols[0];

    for (let i = 1; i < cols.length; i++){
      const dupCol = cols[i];

      if (numRows > 0){
        const rangeA = sh.getRange(2, primary, numRows, 1);
        const rangeB = sh.getRange(2, dupCol, numRows, 1);
        const a = rangeA.getValues();
        const b = rangeB.getValues();

        let changed = false;
        for (let r = 0; r < numRows; r++){
          if (isBlank_(a[r][0]) && !isBlank_(b[r][0])){
            a[r][0] = b[r][0];
            changed = true;
          }
        }
        if (changed) rangeA.setValues(a);
      }
      sh.getRange(1, dupCol).setValue(`old_${key}_${i+1}`);
    }
  });
}

function seedAdminIfEmpty_(){
  const users = readAll_('users');
  if (users.length > 0) return;

  appendRow_('users', {
    id: uid_(),
    username: 'admin',
    password: 'Admin2024!',
    nama: 'Super Admin',
    email: 'admin@example.com',
    role: 'Admin',
    telp: '08xxxxxxxxxx',
    alamat: '-',
    must_change_password: 'FALSE',
    created_at: now_(),
    updated_at: now_()
  });
}

/* =========================
 * AUTH / SESSION
 * ========================= */
function loginUser(payload){
  ensureDatabase_();

  const username = String((payload && payload.username) || '').trim();
  const password = String((payload && payload.password) || '').trim();
  if (!username || !password) return {success:false, message:'Username & Password wajib diisi.'};

  // Users
  const users = readAll_('users');
  const u = users.find(x => String(x.username||'') === username && String(x.password||'') === password);
  if (u){
    try{
      const sh = getSheet_('users');
      const found = findRowById_(sh, u.id);
      if (found){
        const headers = getHeaders_(sh);
        const arr = sh.getRange(found.row, 1, 1, headers.length).getValues()[0];
        const obj = rowToObj_(headers, arr);
        obj.last_login_at = now_();
        writeRow_(sh, headers, found.row, obj);
      }
    } catch(_) {}

    const token = newToken_();
    const must = truthy_(u.must_change_password);

    const userObj = sanitizeUser_(u);
    saveSession_(token, { user: userObj, exp: Date.now() + SESSION_TTL_MS });

    return {success:true, token, user:userObj, mustChangePassword: must};
  }

  // Customers
  const customers = readAll_('customers');
  const c = customers.find(x => String(x.username||'') === username && String(x.password||'') === password);
  if (c){
    const token = newToken_();
    const userObj = {
      id: c.id,
      username: c.username,
      nama: c.nama,
      role: 'Pelanggan',
      telp: c.telp || '',
      alamat: c.alamat || ''
    };
    saveSession_(token, { user: userObj, exp: Date.now() + SESSION_TTL_MS });
    return {success:true, token, user:userObj, mustChangePassword:false};
  }

  return {success:false, message:'Username atau Password salah.'};
}

function getSession(token){
  const sess = getSession_(token);
  if (!sess) return {success:false, message:'SESSION_EXPIRED'};
  return {success:true, session: sess};
}

function logoutSession(token){
  clearSession_(token);
  return {success:true};
}

function changeOwnPassword(token, oldPassword, newPassword){
  const sess = requireSession_(token);
  const user = sess.user;

  if (!['Admin','Sales','Teknisi'].includes(String(user.role||''))){
    return {success:false, message:'Hanya user internal yang bisa ganti password.'};
  }

  const oldP = String(oldPassword||'').trim();
  const newP = String(newPassword||'').trim();
  if (!oldP || !newP) return {success:false, message:'Password tidak valid.'};
  if (newP.length < 6) return {success:false, message:'Password minimal 6 karakter.'};

  const sh = getSheet_('users');
  const found = findRowById_(sh, user.id);
  if (!found) return {success:false, message:'User tidak ditemukan.'};

  const headers = getHeaders_(sh);
  const rowArr = sh.getRange(found.row, 1, 1, headers.length).getValues()[0];
  const obj = rowToObj_(headers, rowArr);

  if (String(obj.password||'') !== oldP) return {success:false, message:'Password lama salah.'};

  obj.password = newP;
  obj.must_change_password = 'FALSE';
  obj.updated_at = now_();
  writeRow_(sh, headers, found.row, obj);

  sess.user.must_change_password = 'FALSE';
  saveSession_(token, sess);

  return {success:true};
}

/* =========================
 * REFERENCE / DASHBOARD
 * ========================= */
function getReferenceData(token){
  const sess = requireSession_(token);
  const user = sess.user;

  const packages = readAll_('packages');
  const locations = readAll_('locations');
  const users = readAll_('users');
  const allCustomers = readAll_('customers');

  const sales = users.filter(u => String(u.role||'') === 'Sales').map(sanitizeUser_);
  const teknisi = users.filter(u => String(u.role||'') === 'Teknisi').map(sanitizeUser_);

  let customers = allCustomers;
  if (user.role === 'Sales'){
    customers = allCustomers.filter(c => String(c.sales_id||'') === String(user.id||''));
  } else if (user.role === 'Pelanggan'){
    customers = allCustomers.filter(c => String(c.id||'') === String(user.id||''));
  }

  return {packages, locations, sales, teknisi, customers};
}

function getDashboardStats(token){
  const sess = requireSession_(token);
  const user = sess.user;

  const usersAll = readAll_('users');
  const packages = readAll_('packages');
  const tickets = readAll_('tickets');
  const customersAll = readAll_('customers');
  const reports = readAll_('reports');

  let customers = customersAll;
  let ticketsFiltered = tickets;

  if (user.role === 'Sales'){
    customers = customersAll.filter(c => String(c.sales_id||'') === String(user.id||''));
    ticketsFiltered = tickets.filter(t => String(t.created_by||'') === String(user.id||''));
  } else if (user.role === 'Teknisi'){
    ticketsFiltered = tickets.filter(t => String(t.teknisi_id||'') === String(user.id||''));
  } else if (user.role === 'Pelanggan'){
    ticketsFiltered = tickets.filter(t => String(t.pelanggan_id||'') === String(user.id||''));
    customers = customersAll.filter(c => String(c.id||'') === String(user.id||''));
  }

  const salesCount = usersAll.filter(u => String(u.role||'') === 'Sales').length;
  const income = computeIncome_(reports);

  return {
    success:true,
    customers: customers.length,
    packages: packages.length,
    tickets: ticketsFiltered.length,
    sales: salesCount,
    income: income
  };
}

function computeIncome_(reports){
  let total = 0;
  (reports||[]).forEach(r => {
    const jenis = String(r.jenis_laporan||'');
    const n = parseInt(r.jumlah||0, 10) || 0;
    if (jenis === 'Pemasukan') total += n;
    else if (jenis === 'Pengeluaran') total -= n;
  });
  return total;
}

/* =========================
 * CRUD
 * ========================= */
function getSheetData(token, tableName){
  const sess = requireSession_(token);
  const user = sess.user;
  const t = canonKey_(tableName);

  if (!SCHEMAS[t]) return [];

  // guard by role
  if (user.role === 'Sales' && t !== 'customers') return [];
  if (user.role === 'Teknisi' && t !== 'tickets') return [];
  if (user.role === 'Pelanggan' && !['tickets','announcements','tutorials'].includes(t)) return [];

  let data = readAll_(t);

  if (t === 'customers' && user.role === 'Sales'){
    data = data.filter(c => String(c.sales_id||'') === String(user.id||''));
  }
  if (t === 'tickets' && user.role === 'Teknisi'){
    data = data.filter(x => String(x.teknisi_id||'') === String(user.id||''));
  }
  if (t === 'tickets' && user.role === 'Pelanggan'){
    data = data.filter(x => String(x.pelanggan_id||'') === String(user.id||''));
  }
  if (t === 'announcements' && user.role !== 'Admin'){
    data = data.filter(a => {
      const target = String(a.target||'Semua');
      return target === 'Semua' || target === user.role;
    });
  }
  if (t === 'tutorials' && user.role !== 'Admin'){
    data = data.filter(a => String(a.kategori||'') === user.role);
  }

  return data;
}

function saveData(token, tableName, data){
  const sess = requireSession_(token);
  const user = sess.user;
  const t = canonKey_(tableName);

  if (!SCHEMAS[t]) return {success:false, message:'Tabel tidak dikenal.'};

  if (user.role === 'Sales' && t !== 'customers') return {success:false, message:'Akses ditolak.'};
  if (user.role === 'Teknisi') return {success:false, message:'Teknisi tidak bisa tambah data.'};
  if (user.role === 'Pelanggan' && t !== 'tickets') return {success:false, message:'Akses ditolak.'};

  const obj = normalizeObjKeys_(data || {});
  obj.id = obj.id || uid_();
  obj.created_at = obj.created_at || now_();
  obj.updated_at = obj.updated_at || obj.created_at;

  if (t === 'tickets'){
    return saveTicket_(user, obj);
  }

  if (t === 'customers'){
    if (user.role === 'Sales'){
      obj.sales_id = String(user.id);
      if (!obj.status) obj.status = 'Antrian';
      if (obj.username === undefined) obj.username = '';
      if (obj.password === undefined) obj.password = '';
    }
    if (user.role === 'Admin'){
      if (!String(obj.sales_id||'').trim()){
        return {success:false, message:'Sales wajib dipilih untuk input oleh Admin.'};
      }
      if (!obj.status) obj.status = 'Antrian';
    }
  }

  if (t === 'users'){
    if (obj.must_change_password === undefined || obj.must_change_password === ''){
      obj.must_change_password = 'TRUE';
    }
  }

  const uploaded = handleFileFields_(t, obj);
  if (!uploaded.success) return uploaded;

  appendRow_(t, obj);
  return {success:true};
}

function updateData(token, tableName, id, data){
  const sess = requireSession_(token);
  const user = sess.user;
  const t = canonKey_(tableName);

  if (!SCHEMAS[t]) return {success:false, message:'Tabel tidak dikenal.'};

  if (user.role === 'Sales' && t !== 'customers') return {success:false, message:'Akses ditolak.'};
  if (user.role === 'Teknisi' && t !== 'tickets') return {success:false, message:'Akses ditolak.'};
  if (user.role === 'Pelanggan' && t !== 'tickets') return {success:false, message:'Akses ditolak.'};

  const sh = getSheet_(t);
  const found = findRowById_(sh, String(id||''));
  if (!found) return {success:false, message:'Data tidak ditemukan.'};

  const headers = getHeaders_(sh);
  const currentArr = sh.getRange(found.row, 1, 1, headers.length).getValues()[0];
  const currentObj = rowToObj_(headers, currentArr);

  // ownership checks
  if (t === 'customers' && user.role === 'Sales'){
    if (String(currentObj.sales_id||'') !== String(user.id||'')){
      return {success:false, message:'Akses ditolak (bukan customer Anda).'};
    }
  }
  if (t === 'tickets' && user.role === 'Teknisi'){
    if (String(currentObj.teknisi_id||'') !== String(user.id||'')){
      return {success:false, message:'Akses ditolak (bukan tiket Anda).'};
    }
  }
  if (t === 'tickets' && user.role === 'Pelanggan'){
    if (String(currentObj.pelanggan_id||'') !== String(user.id||'')){
      return {success:false, message:'Akses ditolak (bukan tiket Anda).'};
    }
  }

  const patch = normalizeObjKeys_(data || {});
  patch.id = String(currentObj.id || id);

  if (t === 'customers' && user.role === 'Sales'){
    patch.sales_id = String(user.id);
    delete patch.status;
    delete patch.username;
    delete patch.password;
  }

  const uploaded = handleFileFields_(t, patch);
  if (!uploaded.success) return uploaded;

  const merged = Object.assign({}, currentObj, patch);
  merged.updated_at = now_();
  if (t === 'tickets') merged.updated_at = now_();

  writeRow_(sh, headers, found.row, merged);
  return {success:true};
}

function deleteData(token, tableName, id){
  const sess = requireSession_(token);
  const user = sess.user;
  const t = canonKey_(tableName);

  if (!SCHEMAS[t]) return {success:false, message:'Tabel tidak dikenal.'};
  if (user.role === 'Sales' && t !== 'customers') return {success:false, message:'Akses ditolak.'};
  if (user.role !== 'Admin' && user.role !== 'Sales') return {success:false, message:'Akses ditolak.'};

  const sh = getSheet_(t);
  const found = findRowById_(sh, String(id||''));
  if (!found) return {success:false, message:'Data tidak ditemukan.'};

  if (t === 'customers' && user.role === 'Sales'){
    const headers = getHeaders_(sh);
    const arr = sh.getRange(found.row, 1, 1, headers.length).getValues()[0];
    const obj = rowToObj_(headers, arr);
    if (String(obj.sales_id||'') !== String(user.id||'')){
      return {success:false, message:'Akses ditolak (bukan customer Anda).'};
    }
  }

  sh.deleteRow(found.row);
  return {success:true};
}

/* =========================
 * TICKETS
 * ========================= */
function saveTicket_(user, obj){
  if (user.role === 'Pelanggan'){
    obj.pelanggan_id = String(user.id);
    obj.status = obj.status || 'Open';
    obj.created_by = String(user.id);
    obj.created_role = 'Pelanggan';
    obj.created_at = obj.created_at || now_();
    obj.updated_at = obj.updated_at || obj.created_at;

    appendRow_('tickets', obj);

    appendRow_('ticket_logs', {
      id: uid_(),
      ticket_id: obj.id,
      actor_id: user.id,
      actor_role: 'Pelanggan',
      actor_name: user.nama || user.username,
      message: 'Tiket dibuat: ' + (obj.judul_laporan || ''),
      attachment: '',
      created_at: now_()
    });
    return {success:true};
  }

  if (user.role === 'Admin'){
    obj.status = obj.status || 'Open';
    obj.created_by = String(user.id);
    obj.created_role = 'Admin';
    obj.created_at = obj.created_at || now_();
    obj.updated_at = obj.updated_at || obj.created_at;

    appendRow_('tickets', obj);

    appendRow_('ticket_logs', {
      id: uid_(),
      ticket_id: obj.id,
      actor_id: user.id,
      actor_role: 'Admin',
      actor_name: user.nama || user.username,
      message: 'Tiket dibuat oleh Admin.',
      attachment: '',
      created_at: now_()
    });
    return {success:true};
  }
  return {success:false, message:'Akses ditolak.'};
}

function getTicketDetail(token, ticketId){
  const sess = requireSession_(token);
  const user = sess.user;
  const id = String(ticketId||'').trim();

  const tickets = readAll_('tickets');
  const ticket = tickets.find(t => String(t.id||'') === id);
  if (!ticket) return {success:false, message:'Tiket tidak ditemukan.'};

  if (user.role === 'Pelanggan' && String(ticket.pelanggan_id||'') !== String(user.id||'')) return {success:false, message:'Akses ditolak.'};
  if (user.role === 'Teknisi' && String(ticket.teknisi_id||'') !== String(user.id||'')) return {success:false, message:'Akses ditolak.'};
  if (user.role === 'Sales') return {success:false, message:'Akses ditolak.'};

  let logs = readAll_('ticket_logs').filter(l => String(l.ticket_id||'') === id);
  logs.sort((a,b) => String(a.created_at||'').localeCompare(String(b.created_at||'')));

  return {success:true, ticket, logs};
}

function addTicketMessage(token, ticketId, payload){
  const sess = requireSession_(token);
  const user = sess.user;

  const id = String(ticketId||'').trim();
  const tickets = readAll_('tickets');
  const ticket = tickets.find(t => String(t.id||'') === id);
  if (!ticket) return {success:false, message:'Tiket tidak ditemukan.'};

  const isAdmin = user.role === 'Admin';
  const isTek = user.role === 'Teknisi' && String(ticket.teknisi_id||'') === String(user.id||'');
  const isCust = user.role === 'Pelanggan' && String(ticket.pelanggan_id||'') === String(user.id||'');
  if (!isAdmin && !isTek && !isCust) return {success:false, message:'Akses ditolak.'};

  const p = payload || {};
  let attachmentUrl = '';

  if (p.attachment && p.attachment.isFile){
    const up = uploadBase64File_(ROOT_UPLOAD_FOLDER_NAME, 'ticket_logs', p.attachment);
    if (!up.success) return up;
    attachmentUrl = up.url;
  }

  appendRow_('ticket_logs', {
    id: uid_(),
    ticket_id: id,
    actor_id: user.id,
    actor_role: user.role,
    actor_name: user.nama || user.username,
    message: String(p.message||''),
    attachment: attachmentUrl,
    created_at: now_()
  });

  updateTicketTimestamp_(id);
  return {success:true};
}

function updateTicketStatus(token, ticketId, status, note){
  const sess = requireSession_(token);
  const user = sess.user;

  if (user.role !== 'Admin' && user.role !== 'Teknisi'){
    return {success:false, message:'Akses ditolak.'};
  }

  const id = String(ticketId||'').trim();
  const st = String(status||'').trim();

  const sh = getSheet_('tickets');
  const found = findRowById_(sh, id);
  if (!found) return {success:false, message:'Tiket tidak ditemukan.'};

  const headers = getHeaders_(sh);
  const arr = sh.getRange(found.row, 1, 1, headers.length).getValues()[0];
  const obj = rowToObj_(headers, arr);

  if (user.role === 'Teknisi' && String(obj.teknisi_id||'') !== String(user.id||'')){
    return {success:false, message:'Akses ditolak (bukan tiket Anda).'};
  }

  obj.status = st || obj.status;
  obj.updated_at = now_();
  writeRow_(sh, headers, found.row, obj);

  appendRow_('ticket_logs', {
    id: uid_(),
    ticket_id: id,
    actor_id: user.id,
    actor_role: user.role,
    actor_name: user.nama || user.username,
    message: 'Update status: ' + obj.status + (note ? (' | ' + String(note)) : ''),
    attachment: '',
    created_at: now_()
  });

  return {success:true};
}

function assignTicketTechnician(token, ticketId, teknisiId){
  const sess = requireSession_(token);
  const user = sess.user;

  if (user.role !== 'Admin') return {success:false, message:'Akses ditolak.'};

  const id = String(ticketId||'').trim();
  const tekId = String(teknisiId||'').trim();

  const sh = getSheet_('tickets');
  const found = findRowById_(sh, id);
  if (!found) return {success:false, message:'Tiket tidak ditemukan.'};

  const headers = getHeaders_(sh);
  const arr = sh.getRange(found.row, 1, 1, headers.length).getValues()[0];
  const obj = rowToObj_(headers, arr);

  obj.teknisi_id = tekId;
  obj.updated_at = now_();
  writeRow_(sh, headers, found.row, obj);

  appendRow_('ticket_logs', {
    id: uid_(),
    ticket_id: id,
    actor_id: user.id,
    actor_role: 'Admin',
    actor_name: user.nama || user.username,
    message: 'Assign teknisi: ' + (tekId || '(kosong)'),
    attachment: '',
    created_at: now_()
  });

  return {success:true};
}

function updateTicketTimestamp_(ticketId){
  const sh = getSheet_('tickets');
  const found = findRowById_(sh, String(ticketId||''));
  if (!found) return;

  const headers = getHeaders_(sh);
  const arr = sh.getRange(found.row, 1, 1, headers.length).getValues()[0];
  const obj = rowToObj_(headers, arr);
  obj.updated_at = now_();
  writeRow_(sh, headers, found.row, obj);
}

/* =========================
 * FILE HANDLING
 * ========================= */
function handleFileFields_(tableName, obj){
  try {
    Object.keys(obj || {}).forEach(k => {
      const v = obj[k];
      if (v && typeof v === 'object' && v.isFile && v.data){
        const up = uploadBase64File_(ROOT_UPLOAD_FOLDER_NAME, tableName, v);
        if (!up.success) throw new Error(up.message || 'Upload gagal');
        obj[k] = up.url;
      }
    });
    return {success:true};
  } catch (e){
    return {success:false, message:'Upload file gagal: ' + e};
  }
}

function uploadBase64File_(rootFolderName, subFolderName, fileObj){
  try{
    const root = getOrCreateFolder_(rootFolderName);
    const sub = getOrCreateChildFolder_(root, String(subFolderName||'files'));

    const bytes = Utilities.base64Decode(String(fileObj.data||''));
    const blob = Utilities.newBlob(bytes, String(fileObj.mimeType||'application/octet-stream'), String(fileObj.name||('file_'+Date.now())));
    const file = sub.createFile(blob);

    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}

    return {success:true, url:file.getUrl(), fileId:file.getId()};
  } catch(e){
    return {success:false, message:String(e)};
  }
}

function getOrCreateFolder_(name){
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function getOrCreateChildFolder_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* =========================
 * DATABASE HELPERS
 * ========================= */
function getDb_(){
  const sp = PropertiesService.getScriptProperties();
  for (const k of DB_PROP_KEYS){
    const id = sp.getProperty(k);
    if (id && String(id).trim()){
      return SpreadsheetApp.openById(String(id).trim());
    }
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  throw new Error('DB Spreadsheet tidak ditemukan. Jika project ini standalone, jalankan setDatabaseSpreadsheetId("SPREADSHEET_ID") dulu.');
}

function getSheet_(name, ss){
  if (!ss) ss = getDb_();
  let sh = ss.getSheetByName(name);
  
  if (!sh) {
    // FIX: Case-Insensitive Search
    const all = ss.getSheets();
    for (const s of all){
      if (s.getName().toLowerCase() === name.toLowerCase()){
        sh = s;
        break;
      }
    }
  }

  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/**
 * FIXED: Scan header lebih pintar agar tidak berhenti di tengah jalan
 */
function getLastHeaderCol_(sheet){
  const max = Math.max(sheet.getLastColumn(), 1);
  if (max === 1){
    const v = sheet.getRange(1,1).getValue();
    return canonKey_(v) ? 1 : 1;
  }

  let end = max;
  while (end >= 1){
    const start = Math.max(1, end - HEADER_SCAN_BLOCK + 1);
    const vals = sheet.getRange(1, start, 1, end - start + 1).getValues()[0];

    for (let i = vals.length - 1; i >= 0; i--){
      if (canonKey_(vals[i])) return start + i;
    }
    end = start - 1;
  }
  return 1;
}

function getHeaders_(sheet){
  const lastHeaderCol = getLastHeaderCol_(sheet);
  const raw = sheet.getRange(1, 1, 1, lastHeaderCol).getValues()[0];
  return raw.map(canonKey_);
}

/**
 * FIXED readAll_ (Permissive & Safe)
 */
function readAll_(sheetName){
  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), getLastHeaderCol_(sh));
  if (lastRow < 2) return [];

  const headers = getHeaders_(sh);
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const out = [];
  for (let r = 0; r < values.length; r++){
    const obj = rowToObj_(headers, values[r]);
    
    // Check if row is empty
    if (Object.keys(obj).length > 0){
       if (!obj.id || String(obj.id).trim() === '') {
         obj.id = 'row_' + (r + 2); 
       }
       out.push(obj);
    }
  }
  return out;
}

function appendRow_(sheetName, obj){
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const sh = getSheet_(sheetName);
    const headers = getHeaders_(sh);

    const row = headers.map(h => (h ? (obj[h] !== undefined ? obj[h] : '') : ''));
    sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

function writeRow_(sheet, headers, rowIndex, obj){
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const current = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const merged = headers.map((h, i) => {
      if (!h) return current[i];
      if (obj[h] === undefined) return current[i];
      return obj[h];
    });
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([merged]);
  } finally {
    lock.releaseLock();
  }
}

function findRowById_(sheet, id){
  const headers = getHeaders_(sheet);
  const idCol = headers.findIndex(h => h === 'id');
  if (idCol === -1) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const range = sheet.getRange(2, idCol+1, lastRow-1, 1).getValues();
  for (let i = 0; i < range.length; i++){
    if (String(range[i][0]||'') === String(id||'')){
      return {row: i+2, col: idCol+1};
    }
  }
  return null;
}

/**
 * FIXED: Auto convert DATE to STRING
 */
function rowToObj_(headers, row){
  const obj = {};
  for (let i = 0; i < headers.length; i++){
    const key = headers[i];
    if (!key) continue;

    let val = row[i];

    // FIX: Convert Date object to String to avoid GAS serialization error
    if (val instanceof Date) {
      try {
        val = Utilities.formatDate(val, TZ, 'yyyy-MM-dd HH:mm:ss');
      } catch(e) {
        val = String(val);
      }
    }

    // kalau sudah ada dan value baru kosong -> jangan timpa
    if (obj[key] !== undefined && isBlank_(val)) continue;

    obj[key] = val;
  }
  return obj;
}

function firstHeaderMap_(headers){
  const map = {};
  headers.forEach((h, idx) => {
    if (!h) return;
    if (!(h in map)) map[h] = idx+1;
  });
  return map;
}

function autoResize_(sh, maxCols){
  try {
    const n = Math.min(maxCols || sh.getLastColumn(), 30);
    sh.autoResizeColumns(1, n);
  } catch(_) {}
}

/* =========================
 * SESSION STORAGE
 * ========================= */
function newToken_(){
  return Utilities.getUuid().replace(/-/g,'');
}

function saveSession_(token, sess){
  const sp = PropertiesService.getScriptProperties();
  sp.setProperty(SESSION_PREFIX + token, JSON.stringify(sess));
  cleanupSessionsMaybe_();
}

function getSession_(token){
  const sp = PropertiesService.getScriptProperties();
  const raw = sp.getProperty(SESSION_PREFIX + String(token||''));
  if (!raw) return null;

  try{
    const sess = JSON.parse(raw);
    if (!sess || !sess.exp || Date.now() > sess.exp){
      sp.deleteProperty(SESSION_PREFIX + token);
      return null;
    }
    return sess;
  } catch(e){
    sp.deleteProperty(SESSION_PREFIX + token);
    return null;
  }
}

function clearSession_(token){
  PropertiesService.getScriptProperties().deleteProperty(SESSION_PREFIX + String(token||''));
}

function requireSession_(token){
  const sess = getSession_(token);
  if (!sess) throw new Error('SESSION_EXPIRED');
  return sess;
}

function cleanupSessionsMaybe_(){
  if (Math.random() > 0.01) return;

  const sp = PropertiesService.getScriptProperties();
  const props = sp.getProperties();
  const now = Date.now();

  Object.keys(props).forEach(k => {
    if (k.indexOf(SESSION_PREFIX) !== 0) return;
    try{
      const sess = JSON.parse(props[k]);
      if (!sess || !sess.exp || now > sess.exp) sp.deleteProperty(k);
    } catch(e){
      sp.deleteProperty(k);
    }
  });
}

/* =========================
 * UTIL
 * ========================= */
function now_(){
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function uid_(){
  return 'id_' + Utilities.getUuid().replace(/-/g,'').slice(0, 12);
}

function isBlank_(v){
  return v === null || v === undefined || String(v).trim() === '';
}

function canonKey_(v){
  let s = String(v === undefined || v === null ? '' : v).trim().toLowerCase();
  if (!s) return '';
  s = s.replace(/[^a-z0-9]+/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '');
  return s;
}

function normalizeObjKeys_(obj){
  const out = {};
  Object.keys(obj || {}).forEach(k => {
    out[canonKey_(k)] = obj[k];
  });
  return out;
}

function truthy_(v){
  const s = String(v === undefined || v === null ? '' : v).trim().toLowerCase();
  return ['1','true','ya','yes','y'].indexOf(s) !== -1;
}

function sanitizeUser_(u){
  return {
    id: u.id,
    username: u.username,
    nama: u.nama,
    email: u.email,
    role: u.role,
    telp: u.telp,
    alamat: u.alamat,
    must_change_password: u.must_change_password
  };
}
