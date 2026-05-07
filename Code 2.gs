/**
 * ═══════════════════════════════════════════════════════════════
 * LAS Fleet Monitoring System — Google Apps Script Backend
 * PT. Pelayaran Lestari Abadi Serasi
 * ═══════════════════════════════════════════════════════════════
 *
 * CARA DEPLOY:
 * ─────────────────────────────────────────────────────────────
 * 1. Buka https://script.google.com → New Project
 * 2. Rename project: "LAS Fleet Monitoring System"
 * 3. Paste seluruh kode ini ke Code.gs
 * 4. Klik (+) Files → HTML → beri nama "index" (tanpa .html)
 *    → paste seluruh isi index.html ke file tersebut
 * 5. Ganti SPREADSHEET_ID di bawah dengan ID Google Sheet Anda
 * 6. Jalankan fungsi setupSheets() sekali dari editor
 * 7. Deploy → New deployment:
 *      Type        : Web app
 *      Execute as  : Me (akun Google Anda)
 *      Who access  : Anyone with Google Account
 * 8. Authorize → Copy URL → buka di browser
 *
 * CARA DAPAT SPREADSHEET_ID:
 * ─────────────────────────────────────────────────────────────
 *   Buka Google Sheet → lihat URL:
 *   https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
 */

// ── KONFIGURASI ──────────────────────────────────────────────
const SPREADSHEET_ID  = 'YOUR_SPREADSHEET_ID_HERE'; // ← WAJIB GANTI
const DB_PROPERTY_KEY = 'las_fms_db_v2';

// ═════════════════════════════════════════════════════════════
// HTTP ENTRY POINT
// ═════════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('LAS Fleet Monitoring System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ═════════════════════════════════════════════════════════════
// CLIENT-CALLABLE FUNCTIONS (dipanggil via google.script.run)
// ═════════════════════════════════════════════════════════════

/** Muat DB dari PropertiesService */
function serverLoadDB() {
  try {
    const props = PropertiesService.getScriptProperties();
    const flag  = props.getProperty(DB_PROPERTY_KEY);
    if (!flag) return null;
    if (flag === '__SPLIT__') {
      return (props.getProperty(DB_PROPERTY_KEY + '_1') || '') +
             (props.getProperty(DB_PROPERTY_KEY + '_2') || '');
    }
    return flag;
  } catch(e) { console.error('serverLoadDB:', e); return null; }
}

/** Simpan DB ke PropertiesService + tulis ke Sheets */
function serverSaveDB(jsonStr) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (jsonStr.length > 480000) {
      props.setProperty(DB_PROPERTY_KEY, '__SPLIT__');
      props.setProperty(DB_PROPERTY_KEY + '_1', jsonStr.slice(0, 240000));
      props.setProperty(DB_PROPERTY_KEY + '_2', jsonStr.slice(240000));
    } else {
      props.setProperty(DB_PROPERTY_KEY, jsonStr);
      props.deleteProperty(DB_PROPERTY_KEY + '_1');
      props.deleteProperty(DB_PROPERTY_KEY + '_2');
    }
    // Tulis ke Sheets (non-critical, best-effort)
    try { writeAllSheets_(JSON.parse(jsonStr)); } catch(e) { console.warn('writeSheets:', e); }
    return true;
  } catch(e) { console.error('serverSaveDB:', e); return false; }
}

/** Autentikasi login dari sheet USERS */
function serverLogin(email, password) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName('USERS');
    if (!sheet) {
      sheet = ss.insertSheet('USERS');
      sheet.getRange(1,1,1,5).setValues([['email','password','name','role','active']]);
      const defaults = [
        ['ahmadmabrur.lasshipping@gmail.com','Admin77','Ahmad Mabrur','Admin','true'],
        ['operation@lasshipping.co.id','Las2024','Operator LAS','Operator','true'],
        ['crewing@lasshipping.co.id','Crew2024','Crewing Manager','Crewing Manager','true'],
        ['viewer@lasshipping.co.id','View2024','Viewer LAS','Viewer','true'],
      ];
      sheet.getRange(2,1,defaults.length,5).setValues(defaults);
      sheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a1f3a').setFontColor('#00e5a0');
      sheet.setFrozenRows(1);
    }
    if (email === '__setup__') return { ok: false };
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const [rEmail, rPass, rName, rRole, rActive] = rows[i];
      if (rEmail === email && rPass === password && String(rActive).toLowerCase() === 'true') {
        return { ok: true, email: rEmail, name: rName, role: rRole };
      }
    }
    return { ok: false, error: 'Email atau password salah, atau akun tidak aktif.' };
  } catch(e) {
    console.error('serverLogin:', e);
    if (email === 'ahmadmabrur.lasshipping@gmail.com' && password === 'Admin77')
      return { ok: true, email, name: 'Ahmad Mabrur', role: 'Admin' };
    return { ok: false, error: 'Server error: ' + e.message };
  }
}

/** Tambah / update user di sheet USERS */
function serverSaveUser(email, password, name, role) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('USERS');
    if (!sheet) return { ok: false, error: 'Sheet USERS tidak ditemukan' };
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === email) {
        sheet.getRange(i+1, 1, 1, 5).setValues([[email, password, name, role, 'true']]);
        return { ok: true, action: 'updated' };
      }
    }
    sheet.appendRow([email, password, name, role, 'true']);
    return { ok: true, action: 'created' };
  } catch(e) { return { ok: false, error: e.message }; }
}

/** Nonaktifkan user */
function serverDeactivateUser(email) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('USERS');
    if (!sheet) return { ok: false };
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === email) {
        sheet.getRange(i+1, 5).setValue('false');
        return { ok: true };
      }
    }
    return { ok: false, error: 'User tidak ditemukan' };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ═════════════════════════════════════════════════════════════
// INTERNAL — tulis ke Sheets untuk keterbacaan
// ═════════════════════════════════════════════════════════════
function writeAllSheets_(db) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  writeSheet_(ss, 'VESSELS',
    ['ID','Nama Kapal','Tipe','GT','DWT','Tahun','Kelas','Status','Lokasi','Catatan'],
    (db.vessels||[]).map(v=>[v.id,v.name,v.type,v.gt,v.dwt,v.year,v.cls,v.status,v.location,v.notes||''])
  );
  writeSheet_(ss, 'CREW_DATA',
    ['ID','Nama','Jabatan','Kapal','Status','Seaman Book','Kode Pelaut','COC','COC Exp','Medical Exp','Emergency'],
    (db.crew||[]).map(c=>{
      const v=(db.vessels||[]).find(x=>x.id===c.vessel)||{name:''};
      return [c.id,c.name,c.rank,v.name,c.status||'onboard',c.seamanbook||'',c.seamanCode||'',c.coc||'',c.cocExp||'',c.medExp||'',c.emergency||''];
    })
  );
  writeSheet_(ss, 'PAYROLL',
    ['ID','Nama Kru','Jabatan','Kapal','Bulan','Hari Aktif','Gaji Pokok','Gaji Aktual','Tunjangan Ops','Tambahan Lain','Total Gaji','BPJS Kes','BPJS TK','PPh 21','Gaji Bersih','Rekening','Bank','Keterangan'],
    (db.payroll||[]).map(p=>{
      const c=(db.crew||[]).find(x=>x.id===p.crewId)||{name:'',rank:''};
      const v=(db.vessels||[]).find(x=>x.id===p.vesselId)||{name:''};
      return [p.id,c.name,c.rank,v.name,p.month,p.hariAktif||30,p.gajiPokok||0,p.gajiAktual||p.gajiPokok||0,p.tunjanganOps||0,p.tambahanLain||0,p.totalGaji||0,p.bpjsKes||0,p.bpjsTK||0,p.potonganPPh21||0,p.gajiBersih||0,p.rekening||'',p.bank||'',p.keterangan||''];
    }),
    {currencyCols:[7,8,9,10,11,12,13,14,15]}
  );
  writeSheet_(ss, 'DOCUMENTS',
    ['ID','Kapal','Jenis','Nama Dokumen','Penerbit','Terbit','Berlaku S/D','Catatan'],
    (db.documents||[]).map(d=>{
      const v=(db.vessels||[]).find(x=>x.id===d.vessel)||{name:''};
      return [d.id,v.name,d.type,d.name,d.issuer||'',d.issue||'',d.expiry==='9999-12-31'?'PERMANEN':(d.expiry||''),d.notes||''];
    })
  );
  writeSheet_(ss, 'CREW_DOCS',
    ['ID','Nama Kru','Jabatan','Kapal','Jenis','No. Dokumen','Penerbit','Terbit','Berlaku S/D','Status','Catatan'],
    (db.crewDocs||[]).map(d=>{
      const c=(db.crew||[]).find(x=>x.id===d.crewId)||{name:'',rank:''};
      const v=(db.vessels||[]).find(x=>x.id===c.vessel)||{name:''};
      return [d.id,c.name,c.rank,v.name,d.type,d.docNumber||'',d.issuer||'',d.issue||'',d.expiry==='9999-12-31'?'PERMANEN':(d.expiry||''),d.status||'valid',d.notes||''];
    })
  );
}

function writeSheet_(ss, name, headers, rows, opts={}) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  else sheet.clearContents();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
    if (opts.currencyCols) {
      opts.currencyCols.forEach(col=>{
        sheet.getRange(2,col,rows.length,1).setNumberFormat('"Rp "#,##0');
      });
    }
  }
  sheet.getRange(1,1,1,headers.length)
    .setFontWeight('bold').setBackground('#1a1f3a')
    .setFontColor('#00e5a0').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  try { sheet.autoResizeColumns(1, headers.length); } catch(e){}
}

// ═════════════════════════════════════════════════════════════
// UTILITY — jalankan manual dari editor GAS
// ═════════════════════════════════════════════════════════════
/** Jalankan sekali sebelum deploy untuk inisialisasi sheet */
function setupSheets() {
  serverLogin('__setup__', '__setup__');
  console.log('✅ Setup OK. Buka Google Sheet Anda untuk melihat sheet USERS.');
  try {
    SpreadsheetApp.getUi().alert('✅ Setup berhasil!\nSheet USERS telah dibuat.\nGanti password default sebelum deploy!');
  } catch(e) { console.log('(Jalankan dari web - alert tidak ditampilkan)'); }
}

/** Reset seluruh DB (HATI-HATI: tidak bisa dibatalkan) */
function resetDB() {
  const p = PropertiesService.getScriptProperties();
  [DB_PROPERTY_KEY, DB_PROPERTY_KEY+'_1', DB_PROPERTY_KEY+'_2'].forEach(k=>p.deleteProperty(k));
  console.log('✅ DB dihapus. Data default akan muncul saat app dibuka kembali.');
}

/** Cek ukuran DB */
function getDBInfo() {
  const p = PropertiesService.getScriptProperties();
  const d = p.getProperty(DB_PROPERTY_KEY);
  if (!d) { console.log('DB kosong'); return; }
  if (d === '__SPLIT__') {
    const s1 = (p.getProperty(DB_PROPERTY_KEY+'_1')||'').length;
    const s2 = (p.getProperty(DB_PROPERTY_KEY+'_2')||'').length;
    console.log(`DB size: ${(s1+s2).toLocaleString()} chars (SPLIT: part1=${s1}, part2=${s2})`);
  } else {
    console.log(`DB size: ${d.length.toLocaleString()} chars (~${Math.round(d.length/1024)}KB / 500KB limit)`);
  }
}

/** Test koneksi dan login */
function testAll() {
  console.log('=== TEST LAS FMS ===');
  console.log('SPREADSHEET_ID:', SPREADSHEET_ID);
  const loginOk = serverLogin('ahmadmabrur.lasshipping@gmail.com','Admin77');
  console.log('Login test:', JSON.stringify(loginOk));
  const dbData = serverLoadDB();
  console.log('DB loaded:', dbData ? `${dbData.length} chars` : 'null (kosong)');
  console.log('=== END TEST ===');
}
