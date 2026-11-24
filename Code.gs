/**
 * ============================================================================
 * APLIKASI BUKU INDUK SISWA SMP NEGERI 104 JAKARTA
 * Backend Google Apps Script & Otomatisasi Setup Google Spreadsheet
 * Created by Supry
 * ============================================================================
 */

// Nama-nama Sheet (Tab) di dalam Google Spreadsheet
const SHEET_BIODATA = 'Biodata_Siswa';
const SHEET_NILAI = 'Nilai_Semester';
const SHEET_CONFIG = 'Pengaturan_Admin';

// Folder Google Drive tujuan penyimpanan pas foto siswa (3x4)
const PHOTO_FOLDER_ID = '1a5ByuM2dr1_zoFfgxlSwItZmXjWffN1J';

// Daftar Header Horizontal Lengkap untuk Sheet Biodata_Siswa (68 Kolom termasuk Pas Foto)
const BIODATA_HEADERS = [
  // A. KETERANGAN SISWA (27 Kolom)
  'NIS',
  'Nama Lengkap',
  'Jenis Kelamin',
  'Kelas',
  'Tahun Pelajaran',
  'NIK',
  'Nomer KK',
  'Anak ke',
  'Tempat Lahir Siswa',
  'Tanggal Lahir Siswa',
  'No. Akte Lahir',
  'Asal Sekolah SD/MI',
  'NISN',
  'Nomer Seri Ijazah SD/MI',
  'Nomer Peserta Ujian SD/MI',
  'Agama',
  'Alamat',
  'RT',
  'RW',
  'Kelurahan',
  'Kecamatan',
  'Kode Pos',
  'Tempat Tinggal',
  'Transportasi',
  'Nomor HP/WA',
  'Nomer Telepon',
  'Email',

  // B. DATA PERIODIK (5 Kolom)
  'Tinggi Badan (cm)',
  'Berat Badan (kg)',
  'Jarak Rumah Ke Sekolah',
  'Waktu tempuh ke sekolah (Jam/Menit)',
  'Jumlah Saudara Kandung',

  // C. KETERANGAN AYAH KANDUNG (5 Kolom)
  'Nama Ayah',
  'Tanggal Lahir Ayah',
  'Pendidikan Ayah',
  'Pekerjaan Ayah',
  'Penghasilan Ayah',

  // D. KETERANGAN IBU KANDUNG (5 Kolom)
  'Nama Ibu',
  'Tanggal Lahir Ibu',
  'Pendidikan Ibu',
  'Pekerjaan Ibu',
  'Penghasilan Ibu',

  // E. KETERANGAN WALI (5 Kolom)
  'Nama Wali',
  'Tanggal Lahir Wali',
  'Pendidikan Wali',
  'Pekerjaan Wali',
  'Penghasilan Wali',

  // F. PENDIDIKAN SEBELUMNYA - Mutasi Masuk (5 Kolom)
  'Mutasi Masuk - Nama Sekolah Asal',
  'Mutasi Masuk - Alamat Sekolah Asal',
  'Mutasi Masuk - Nomor Surat',
  'Mutasi Masuk - Diterima Tanggal',
  'Mutasi Masuk - Di Kelas',

  // F. PENDIDIKAN SEBELUMNYA - Mutasi Keluar (6 Kolom)
  'Mutasi Keluar - Nomor Surat',
  'Mutasi Keluar - Tanggal',
  'Mutasi Keluar - Kelas yang ditinggalkan',
  'Mutasi Keluar - Alasan Pindah',
  'Mutasi Keluar - Sekolah Tujuan',
  'Mutasi Keluar - Alamat Sekolah',

  // F. PENDIDIKAN SEBELUMNYA - Tamat Belajar / Lulus (4 Kolom)
  'Tamat Belajar - Tanggal Tamat Belajar',
  'Tamat Belajar - Nomer Seri Ijazah',
  'Tamat Belajar - Nomer Peserta',
  'Tamat Belajar - Melanjutkan ke Sekolah',

  // F. PENDIDIKAN SEBELUMNYA - Putus Sekolah (3 Kolom)
  'Putus Sekolah - Hari/Tanggal',
  'Putus Sekolah - Kelas',
  'Putus Sekolah - Alasan',

  // G. LAIN-LAIN (1 Kolom)
  'Catatan Penting',

  // H. PAS FOTO SISWA (Google Drive)
  'foto_drive_id',
  'foto_url'
];

// Daftar 10 Mata Pelajaran Resmi Kurikulum Merdeka SMP
const MAPEL_LIST = [
  'PABP',
  'Pendidikan Pancasila',
  'Bahasa Indonesia',
  'Matematika',
  'IPA',
  'IPS',
  'Bahasa Inggris',
  'PJOK',
  'Informatika',
  'Seni Budaya'
];

// Daftar 14 Ekstrakurikuler
const EKSKUL_LIST = [
  'Pramuka',
  'Basket',
  'KIR',
  'Paduan Suara',
  'Paskibra',
  'Hadroh',
  'Kaligrafi',
  'Karate',
  'PMR',
  'Tari',
  'Pencak Silat',
  'Futsal',
  'English Community',
  'Voly'
];

/**
 * Membangun daftar header horizontal untuk Sheet Nilai_Semester
 */
function getNilaiHeaders() {
  const headers = ['NIS', 'Nama Lengkap', 'Kelas'];
  for (let sem = 1; sem <= 6; sem++) {
    MAPEL_LIST.forEach(mapel => {
      headers.push(`S${sem} - ${mapel}`);
    });
    headers.push(`S${sem} - Sakit`);
    headers.push(`S${sem} - Izin`);
    headers.push(`S${sem} - Alpa`);
    EKSKUL_LIST.forEach(ekskul => {
      headers.push(`S${sem} - ${ekskul}`);
    });
    headers.push(`Rata-Rata Semester ${sem}`);
  }
  headers.push('Keputusan Kelas 7');
  headers.push('Keputusan Kelas 8');
  headers.push('Keputusan Kelas 9');
  headers.push('Rata-Rata Keseluruhan');
  return headers;
}

/**
 * Mendapatkan instance Spreadsheet aktif atau membuat baru jika belum ada
 */
function getSpreadsheet() {
  const scriptProps = PropertiesService.getScriptProperties();
  let ssId = scriptProps.getProperty('SPREADSHEET_ID');
  let ss = null;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ss = null;
    }
  }

  if (!ss) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        scriptProps.setProperty('SPREADSHEET_ID', ss.getId());
      }
    } catch (e) {
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('Database Buku Induk Siswa SMP Negeri 104 Jakarta');
    scriptProps.setProperty('SPREADSHEET_ID', ss.getId());
  }

  return ss;
}

/**
 * ============================================================================
 * FUNGSI SETUP DATABASE (JALANKAN SEKALI DI EDITOR APPS SCRIPT)
 * ============================================================================
 */
function setupDatabase() {
  const ss = getSpreadsheet();
  const ssUrl = ss.getUrl();

  // 1. Setup Sheet Biodata_Siswa
  let sheetBiodata = ss.getSheetByName(SHEET_BIODATA);
  if (!sheetBiodata) {
    sheetBiodata = ss.insertSheet(SHEET_BIODATA);
  }
  if (sheetBiodata.getLastRow() === 0) {
    sheetBiodata.getRange(1, 1, 1, BIODATA_HEADERS.length).setValues([BIODATA_HEADERS]);
    const headerRange = sheetBiodata.getRange(1, 1, 1, BIODATA_HEADERS.length);
    headerRange
      .setBackground('#1e3a8a')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheetBiodata.setFrozenRows(1);

    const sampleStudent = new Array(BIODATA_HEADERS.length).fill('');
    sampleStudent[0] = '242507001';
    sampleStudent[1] = 'Ahmad Fauzan Pratama';
    sampleStudent[2] = 'Laki-laki';
    sampleStudent[3] = 'VII-A';
    sampleStudent[4] = '2024/2025';
    sampleStudent[5] = '3174011505120001';
    sampleStudent[6] = '3174010203100005';
    sampleStudent[7] = '1';
    sampleStudent[8] = 'Jakarta';
    sampleStudent[9] = '2012-05-15';
    sampleStudent[10] = '3174-LU-20052012-0012';
    sampleStudent[11] = 'SD Negeri 01 Mampang Prapatan';
    sampleStudent[12] = '0124589632';
    sampleStudent[13] = 'DN-01/D-SD/24/0012451';
    sampleStudent[14] = '24-01-001-012-4';
    sampleStudent[15] = 'Islam';
    sampleStudent[16] = 'Jl. Mampang Prapatan VIII No. 14';
    sampleStudent[17] = '004';
    sampleStudent[18] = '002';
    sampleStudent[19] = 'Tegal Parang';
    sampleStudent[20] = 'Mampang Prapatan';
    sampleStudent[21] = '12790';
    sampleStudent[22] = 'Bersama Orang Tua';
    sampleStudent[23] = 'Jalan Kaki';
    sampleStudent[24] = '081234567890';
    sampleStudent[25] = '0217981234';
    sampleStudent[26] = 'ahmad.fauzan@siswa.smpn104jkt.sch.id';
    sampleStudent[27] = '156';
    sampleStudent[28] = '45';
    sampleStudent[29] = '1.2 km';
    sampleStudent[30] = '0 Jam 15 Menit';
    sampleStudent[31] = '1';
    sampleStudent[32] = 'Budi Santoso, S.E.';
    sampleStudent[33] = '1982-08-17';
    sampleStudent[34] = 'S1';
    sampleStudent[35] = 'Karyawan Swasta';
    sampleStudent[36] = 'Rp 5.000.000 - Rp 10.000.000';
    sampleStudent[37] = 'Siti Aminah, S.Pd.';
    sampleStudent[38] = '1985-04-21';
    sampleStudent[39] = 'S1';
    sampleStudent[40] = 'Guru';
    sampleStudent[41] = 'Rp 5.000.000 - Rp 10.000.000';
    sampleStudent[65] = 'Siswa aktif berprestasi di bidang Olimpiade Sains Nasional (OSN) IPA.';
    sampleStudent[66] = '';
    sampleStudent[67] = '';

    sheetBiodata.getRange(2, 1, 1, sampleStudent.length).setValues([sampleStudent]);
  }

  // 2. Setup Sheet Nilai_Semester
  let sheetNilai = ss.getSheetByName(SHEET_NILAI);
  const nilaiHeaders = getNilaiHeaders();
  if (!sheetNilai) {
    sheetNilai = ss.insertSheet(SHEET_NILAI);
  }
  if (sheetNilai.getLastRow() === 0) {
    sheetNilai.getRange(1, 1, 1, nilaiHeaders.length).setValues([nilaiHeaders]);
    const headerNilaiRange = sheetNilai.getRange(1, 1, 1, nilaiHeaders.length);
    headerNilaiRange
      .setBackground('#0f766e')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheetNilai.setFrozenRows(1);

    const sampleNilai = ['242507001', 'Ahmad Fauzan Pratama', 'VII-A'];
    const baseScores = [88, 86, 85, 84, 89, 87, 86, 90, 88, 87];
    let totalAll = 0;
    for (let sem = 1; sem <= 6; sem++) {
      let sumSem = 0;
      baseScores.forEach((sc, idx) => {
        const val = Math.min(98, sc + ((sem + idx) % 4));
        sampleNilai.push(val);
        sumSem += val;
      });
      sampleNilai.push(0); // Sakit
      sampleNilai.push(0); // Izin
      sampleNilai.push(0); // Alpa
      EKSKUL_LIST.forEach(ekskul => {
        sampleNilai.push(ekskul === 'Pramuka' ? 'A' : '');
      });
      const avgSem = Number((sumSem / 10).toFixed(2));
      sampleNilai.push(avgSem);
      totalAll += avgSem;
    }
    sampleNilai.push('Naik ke kelas 8');
    sampleNilai.push('Naik ke kelas 9');
    sampleNilai.push('Lulus');
    sampleNilai.push(Number((totalAll / 6).toFixed(2)));
    sheetNilai.getRange(2, 1, 1, sampleNilai.length).setValues([sampleNilai]);
  }

  // 3. Setup Sheet Pengaturan_Admin
  let sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet(SHEET_CONFIG);
  }
  if (sheetConfig.getLastRow() === 0) {
    const defaultConfigs = [
      ['KUNCI_PENGATURAN', 'NILAI_PENGATURAN', 'KETERANGAN'],
      ['ADMIN_PASSWORD', 'admin104', 'Password Login Admin Buku Induk Siswa'],
      ['NAMA_SEKOLAH', 'SMP NEGERI 104 JAKARTA', 'Nama Resmi Satuan Pendidikan'],
      ['NPSN', '20102534', 'Nomor Pokok Sekolah Nasional'],
      ['ALAMAT_SEKOLAH', 'Jl. Mampang Prapatan XIII, Kel. Tegal Parang, Kec. Mampang Prapatan, Jakarta Selatan 12790', 'Alamat Lengkap Sekolah'],
      ['KEPALA_SEKOLAH', 'Drs. H. Hendra Wijaya, M.Pd.', 'Nama Kepala Sekolah untuk Tanda Tangan PDF'],
      ['NIP_KEPALA_SEKOLAH', '196805121994121002', 'NIP Kepala Sekolah']
    ];
    sheetConfig.getRange(1, 1, defaultConfigs.length, 3).setValues(defaultConfigs);
    sheetConfig.getRange(1, 1, 1, 3)
      .setBackground('#334155')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheetConfig.setFrozenRows(1);
    sheetConfig.autoResizeColumns(1, 3);
  }

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('SETUP DATABASE BERHASIL! URL: ' + ssUrl);
  return {
    success: true,
    message: 'Database Google Spreadsheet berhasil dibuat dan dikonfigurasi.',
    spreadsheetUrl: ssUrl
  };
}

/**
 * Entry Point Web App Google Apps Script
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Buku Induk Siswa SMP Negeri 104 Jakarta')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Mengambil seluruh konfigurasi admin dari sheet Pengaturan_Admin
 */
function getConfigMap() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName(SHEET_CONFIG);
  }
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      config[String(data[i][0]).trim()] = String(data[i][1] || '');
    }
  }
  return config;
}

/**
 * Verifikasi Login Admin
 */
function verifyAdminLogin(password) {
  try {
    const config = getConfigMap();
    const validPassword = config['ADMIN_PASSWORD'] || 'admin104';
    if (String(password) === String(validPassword)) {
      return { success: true, message: 'Login berhasil! Selamat datang di Buku Induk Siswa SMPN 104 Jakarta.' };
    }
    return { success: false, message: 'Password salah. Silakan periksa kembali password Anda.' };
  } catch (err) {
    return { success: false, message: 'Terjadi kesalahan server: ' + err.message };
  }
}

/**
 * Menormalisasi berbagai format tanggal menjadi YYYY-MM-DD
 */
function normalizeDateString(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(val).trim();
  let m = str.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) {
    return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
  }
  m = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) {
    return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return str;
}

/**
 * Verifikasi Login Siswa (Mode Hanya Lihat / Read-Only) menggunakan NIS & Tanggal Lahir Siswa
 */
function verifyStudentLogin(nisInput, tanggalLahirInput) {
  try {
    const targetNis = String(nisInput || '').trim();
    const targetDob = normalizeDateString(tanggalLahirInput);

    if (!targetNis || !targetDob) {
      return { success: false, message: 'NIS dan Tanggal Lahir wajib diisi.' };
    }

    const ss = getSpreadsheet();
    let sheetBio = ss.getSheetByName(SHEET_BIODATA);
    if (!sheetBio) {
      setupDatabase();
      sheetBio = ss.getSheetByName(SHEET_BIODATA);
    }

    const displayValues = sheetBio.getDataRange().getDisplayValues();
    const rawValues = sheetBio.getDataRange().getValues();
    if (displayValues.length <= 1) {
      return { success: false, message: 'Belum ada data siswa terdaftar di database.' };
    }

    const headers = displayValues[0];
    const dobIdx = headers.indexOf('Tanggal Lahir Siswa');
    let matchedStudent = null;

    for (let r = 1; r < displayValues.length; r++) {
      if (String(displayValues[r][0]).trim() === targetNis) {
        const storedDobNorm = dobIdx !== -1
          ? (normalizeDateString(rawValues[r][dobIdx]) || normalizeDateString(displayValues[r][dobIdx]))
          : '';
        if (storedDobNorm === targetDob) {
          matchedStudent = {};
          headers.forEach((h, idx) => {
            matchedStudent[h] = displayValues[r][idx] !== undefined ? String(displayValues[r][idx]) : '';
          });
        }
        break;
      }
    }

    if (!matchedStudent) {
      return { success: false, message: 'NIS atau Tanggal Lahir tidak cocok. Periksa kembali data Anda.' };
    }

    let matchedNilai = {};
    const sheetNilai = ss.getSheetByName(SHEET_NILAI);
    if (sheetNilai) {
      const nilaiValues = sheetNilai.getDataRange().getDisplayValues();
      if (nilaiValues.length > 1) {
        const headersN = nilaiValues[0];
        for (let r = 1; r < nilaiValues.length; r++) {
          if (String(nilaiValues[r][0]).trim() === targetNis) {
            headersN.forEach((h, idx) => {
              matchedNilai[h] = nilaiValues[r][idx] !== undefined ? String(nilaiValues[r][idx]) : '';
            });
            break;
          }
        }
      }
    }

    return {
      success: true,
      student: matchedStudent,
      nilai: matchedNilai,
      message: 'Login siswa berhasil!'
    };
  } catch (err) {
    return { success: false, message: 'Terjadi kesalahan saat verifikasi siswa: ' + err.message };
  }
}

/**
 * Mengubah Password Admin dan menyimpannya ke Google Spreadsheet
 */
function changeAdminPassword(oldPassword, newPassword) {
  try {
    if (!newPassword || String(newPassword).trim().length < 4) {
      return { success: false, message: 'Password baru minimal terdiri dari 4 karakter.' };
    }
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CONFIG);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'ADMIN_PASSWORD') {
        const currentPass = String(data[i][1]);
        if (String(oldPassword) !== currentPass) {
          return { success: false, message: 'Password lama tidak sesuai!' };
        }
        sheet.getRange(i + 1, 2).setValue(String(newPassword).trim());
        return { success: true, message: 'Password admin berhasil diperbarui di Google Spreadsheet!' };
      }
    }
    return { success: false, message: 'Konfigurasi password tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal mengubah password: ' + err.message };
  }
}

/**
 * Mengambil seluruh data awal (Biodata Siswa, Nilai 6 Semester, dan Info Sekolah)
 */
function getAllInitialData() {
  try {
    const ss = getSpreadsheet();
    if (!ss.getSheetByName(SHEET_BIODATA)) {
      setupDatabase();
    }

    const config = getConfigMap();
    delete config['ADMIN_PASSWORD'];

    const sheetBio = ss.getSheetByName(SHEET_BIODATA);
    ensurePhotoColumnsExist(sheetBio);
    const bioValues = sheetBio.getDataRange().getDisplayValues();
    const biodataList = [];

    if (bioValues.length > 1) {
      const headers = bioValues[0];
      for (let r = 1; r < bioValues.length; r++) {
        const row = bioValues[r];
        if (!row[0] && !row[1]) continue;
        const obj = {};
        headers.forEach((h, idx) => {
          obj[h] = row[idx] !== undefined ? String(row[idx]) : '';
        });
        biodataList.push(obj);
      }
    }

    const sheetNilai = ss.getSheetByName(SHEET_NILAI);
    const nilaiValues = sheetNilai ? sheetNilai.getDataRange().getDisplayValues() : [];
    const nilaiList = [];

    if (nilaiValues.length > 1) {
      const headersN = nilaiValues[0];
      for (let r = 1; r < nilaiValues.length; r++) {
        const row = nilaiValues[r];
        if (!row[0]) continue;
        const obj = {};
        headersN.forEach((h, idx) => {
          obj[h] = row[idx] !== undefined ? String(row[idx]) : '';
        });
        nilaiList.push(obj);
      }
    }

    return {
      success: true,
      biodata: biodataList,
      nilai: nilaiList,
      config: config,
      spreadsheetUrl: ss.getUrl()
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat data: ' + err.message };
  }
}

/**
 * Memastikan kolom foto_drive_id dan foto_url tersedia pada sheet Biodata_Siswa
 */
function ensurePhotoColumnsExist(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  let nextCol = lastCol;

  if (headers.indexOf('foto_drive_id') === -1) {
    nextCol++;
    sheet.getRange(1, nextCol).setValue('foto_drive_id')
      .setBackground('#1e3a8a')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  if (headers.indexOf('foto_url') === -1) {
    nextCol++;
    sheet.getRange(1, nextCol).setValue('foto_url')
      .setBackground('#1e3a8a')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
}

/**
 * Mengunggah atau menimpa Pas Foto Siswa ([NIS].jpg) ke Folder Google Drive tujuan
 */
function uploadStudentPhoto(nisInput, base64Data) {
  try {
    const nis = String(nisInput || '').trim();
    if (!nis) {
      return { success: false, message: 'NIS siswa wajib diisi.' };
    }
    if (!base64Data) {
      return { success: false, message: 'Data gambar tidak ditemukan.' };
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_BIODATA);
    if (!sheet) {
      return { success: false, message: 'Sheet Biodata_Siswa tidak ditemukan.' };
    }
    ensurePhotoColumnsExist(sheet);

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const fotoIdCol = headers.indexOf('foto_drive_id') + 1;
    const fotoUrlCol = headers.indexOf('foto_url') + 1;

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === nis) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'NIS ' + nis + ' tidak ditemukan di database Biodata_Siswa.' };
    }

    const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    const fileName = nis + '.jpg';

    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const cleanBase64 = String(base64Data).replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, '');
    const bytes = Utilities.base64Decode(cleanBase64);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', fileName);
    const newFile = folder.createFile(blob);

    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = newFile.getId();
    const photoUrl = 'https://lh3.googleusercontent.com/d/' + fileId;

    if (fotoIdCol > 0) sheet.getRange(targetRow, fotoIdCol).setValue(fileId);
    if (fotoUrlCol > 0) sheet.getRange(targetRow, fotoUrlCol).setValue(photoUrl);

    return {
      success: true,
      nis: nis,
      fileId: fileId,
      photoUrl: photoUrl,
      message: 'Pas foto siswa (NIS: ' + nis + ') berhasil disimpan ke Google Drive!'
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengunggah foto ke Google Drive: ' + err.message };
  }
}

/**
 * Menghapus Pas Foto Siswa dari Google Drive
 */
function deleteStudentPhoto(nisInput) {
  try {
    const nis = String(nisInput || '').trim();
    if (!nis) {
      return { success: false, message: 'NIS siswa tidak valid.' };
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_BIODATA);
    if (!sheet) {
      return { success: false, message: 'Sheet Biodata_Siswa tidak ditemukan.' };
    }
    ensurePhotoColumnsExist(sheet);

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const fotoIdCol = headers.indexOf('foto_drive_id') + 1;
    const fotoUrlCol = headers.indexOf('foto_url') + 1;

    let targetRow = -1;
    let existingFileId = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === nis) {
        targetRow = i + 1;
        if (fotoIdCol > 0) existingFileId = String(data[i][fotoIdCol - 1] || '').trim();
        break;
      }
    }

    if (existingFileId) {
      try {
        DriveApp.getFileById(existingFileId).setTrashed(true);
      } catch (e) {}
    }

    try {
      const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
      const files = folder.getFilesByName(nis + '.jpg');
      while (files.hasNext()) {
        files.next().setTrashed(true);
      }
    } catch (e) {}

    if (targetRow > 0) {
      if (fotoIdCol > 0) sheet.getRange(targetRow, fotoIdCol).setValue('');
      if (fotoUrlCol > 0) sheet.getRange(targetRow, fotoUrlCol).setValue('');
    }

    return { success: true, message: 'Pas foto siswa (NIS: ' + nis + ') berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus pas foto: ' + err.message };
  }
}

/**
 * Mengambil gambar dari Google Drive sebagai Data URI Base64 untuk PDF bebas CORS
 */
function getStudentPhotoBase64(fileId) {
  try {
    if (!fileId) {
      return { success: false, message: 'File ID foto kosong.' };
    }
    const file = DriveApp.getFileById(String(fileId).trim());
    const blob = file.getBlob();
    const mimeType = blob.getContentType() || 'image/jpeg';
    const base64 = Utilities.base64Encode(blob.getBytes());
    return {
      success: true,
      dataUrl: 'data:' + mimeType + ';base64,' + base64
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat gambar Base64: ' + err.message };
  }
}

/**
 * Menyimpan atau Memperbarui 1 Data Biodata Siswa
 */
function saveBiodataSiswa(studentObj, originalNis) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_BIODATA);
    ensurePhotoColumnsExist(sheet);
    const nis = String(studentObj['NIS'] || '').trim();
    if (!nis) {
      return { success: false, message: 'NIS wajib diisi sebelum menyimpan data!' };
    }

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const searchNis = originalNis ? String(originalNis).trim() : nis;

    let targetRow = -1;
    let existingRow = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchNis) {
        targetRow = i + 1;
        existingRow = data[i];
        break;
      }
    }

    if (existingRow) {
      const idxId = headers.indexOf('foto_drive_id');
      const idxUrl = headers.indexOf('foto_url');
      if (idxId !== -1 && (studentObj['foto_drive_id'] === undefined || studentObj['foto_drive_id'] === '')) {
        studentObj['foto_drive_id'] = existingRow[idxId] || '';
      }
      if (idxUrl !== -1 && (studentObj['foto_url'] === undefined || studentObj['foto_url'] === '')) {
        studentObj['foto_url'] = existingRow[idxUrl] || '';
      }
    }

    const rowData = BIODATA_HEADERS.map(header => studentObj[header] !== undefined ? String(studentObj[header]) : '');

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
      syncStudentMetaToNilai(nis, studentObj['Nama Lengkap'] || '', studentObj['Kelas'] || '', searchNis);
      return { success: true, message: 'Data biodata siswa (' + studentObj['Nama Lengkap'] + ') berhasil diperbarui!' };
    } else {
      sheet.appendRow(rowData);
      return { success: true, message: 'Data biodata siswa baru (' + studentObj['Nama Lengkap'] + ') berhasil disimpan!' };
    }
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan biodata: ' + err.message };
  }
}

/**
 * Sinkronisasi perubahan NIS, Nama, dan Kelas ke Sheet Nilai_Semester
 */
function syncStudentMetaToNilai(newNis, nama, kelas, oldNis) {
  const ss = getSpreadsheet();
  const sheetNilai = ss.getSheetByName(SHEET_NILAI);
  if (!sheetNilai) return;
  const data = sheetNilai.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(oldNis).trim()) {
      sheetNilai.getRange(i + 1, 1, 1, 3).setValues([[newNis, nama, kelas]]);
      break;
    }
  }
}

/**
 * Menghapus Data Siswa dari Sheet Biodata dan Sheet Nilai
 */
function deleteBiodataSiswa(nis) {
  try {
    const ss = getSpreadsheet();
    const targetNis = String(nis).trim();

    const sheetBio = ss.getSheetByName(SHEET_BIODATA);
    const bioData = sheetBio.getDataRange().getDisplayValues();
    for (let i = bioData.length - 1; i >= 1; i--) {
      if (String(bioData[i][0]).trim() === targetNis) {
        sheetBio.deleteRow(i + 1);
      }
    }

    const sheetNilai = ss.getSheetByName(SHEET_NILAI);
    if (sheetNilai) {
      const nilaiData = sheetNilai.getDataRange().getDisplayValues();
      for (let i = nilaiData.length - 1; i >= 1; i--) {
        if (String(nilaiData[i][0]).trim() === targetNis) {
          sheetNilai.deleteRow(i + 1);
        }
      }
    }

    return { success: true, message: 'Data siswa dengan NIS ' + targetNis + ' berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus data siswa: ' + err.message };
  }
}

/**
 * Impor Massal Biodata Siswa dari File Excel (Upsert berdasarkan NIS)
 */
function importBiodataBatch(studentsArray) {
  try {
    if (!Array.isArray(studentsArray) || studentsArray.length === 0) {
      return { success: false, message: 'Tidak ada baris data valid di dalam file Excel.' };
    }
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_BIODATA);
    ensurePhotoColumnsExist(sheet);
    const existingData = sheet.getDataRange().getDisplayValues();
    const headers = existingData[0];
    const idxFotoId = headers.indexOf('foto_drive_id');
    const idxFotoUrl = headers.indexOf('foto_url');
    const nisRowMap = {};

    for (let i = 1; i < existingData.length; i++) {
      const existingNis = String(existingData[i][0]).trim();
      if (existingNis) nisRowMap[existingNis] = i + 1;
    }

    let inserted = 0;
    let updated = 0;
    const rowsToAppend = [];

    studentsArray.forEach(item => {
      const nis = String(item['NIS'] || '').trim();
      if (!nis) return;
      if (nisRowMap[nis]) {
        const oldRow = existingData[nisRowMap[nis] - 1];
        if (idxFotoId !== -1 && !item['foto_drive_id']) item['foto_drive_id'] = oldRow[idxFotoId] || '';
        if (idxFotoUrl !== -1 && !item['foto_url']) item['foto_url'] = oldRow[idxFotoUrl] || '';
      }
      const rowValues = BIODATA_HEADERS.map(h => item[h] !== undefined ? String(item[h]) : '');
      if (nisRowMap[nis]) {
        sheet.getRange(nisRowMap[nis], 1, 1, rowValues.length).setValues([rowValues]);
        updated++;
      } else {
        rowsToAppend.push(rowValues);
        inserted++;
      }
    });

    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, BIODATA_HEADERS.length).setValues(rowsToAppend);
    }

    return {
      success: true,
      message: `Impor Excel Biodata selesai: ${inserted} siswa ditambahkan, ${updated} siswa diperbarui.`
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengimpor Excel biodata: ' + err.message };
  }
}

/**
 * Menyimpan atau Memperbarui Nilai 6 Semester Siswa
 */
function saveNilaiSiswa(nilaiObj) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NILAI);
    const headers = getNilaiHeaders();
    const nis = String(nilaiObj['NIS'] || '').trim();

    if (!nis) {
      return { success: false, message: 'Pilih siswa terlebih dahulu sebelum menyimpan nilai.' };
    }

    const currentLastCol = sheet.getLastColumn();
    if (currentLastCol < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const rowValues = headers.map(h => nilaiObj[h] !== undefined ? String(nilaiObj[h]) : '');
    const data = sheet.getDataRange().getDisplayValues();
    let targetRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === nis) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return { success: true, message: 'Nilai 6 Semester untuk ' + (nilaiObj['Nama Lengkap'] || nis) + ' berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan nilai: ' + err.message };
  }
}

/**
 * Impor Massal Nilai 6 Semester dari File Excel (Partial/Merge Update)
 */
function importNilaiBatch(nilaiArray) {
  try {
    if (!Array.isArray(nilaiArray) || nilaiArray.length === 0) {
      return { success: false, message: 'Tidak ada data nilai valid di dalam file Excel.' };
    }
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NILAI);
    const headers = getNilaiHeaders();

    if (sheet.getLastColumn() < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const existingData = sheet.getDataRange().getDisplayValues();
    const nisRowMap = {};

    for (let i = 1; i < existingData.length; i++) {
      const existingNis = String(existingData[i][0]).trim();
      if (existingNis) nisRowMap[existingNis] = i + 1;
    }

    let inserted = 0;
    let updated = 0;
    const rowsToAppend = [];

    nilaiArray.forEach(item => {
      const nis = String(item['NIS'] || '').trim();
      if (!nis) return;

      if (nisRowMap[nis]) {
        const existingRow = existingData[nisRowMap[nis] - 1];
        const mergedRowValues = headers.map((h, idx) => item[h] !== undefined ? String(item[h]) : String(existingRow[idx] || ''));
        sheet.getRange(nisRowMap[nis], 1, 1, mergedRowValues.length).setValues([mergedRowValues]);
        existingData[nisRowMap[nis] - 1] = mergedRowValues;
        updated++;
      } else {
        const rowValues = headers.map(h => item[h] !== undefined ? String(item[h]) : '');
        rowsToAppend.push(rowValues);
        inserted++;
      }
    });

    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    }

    return {
      success: true,
      message: `Impor Excel Nilai selesai: ${inserted} data ditambahkan, ${updated} data diperbarui.`
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengimpor Excel nilai: ' + err.message };
  }
}

/**
 * Menyimpan Pengaturan Identitas Sekolah
 */
function saveSchoolConfig(configUpdates) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CONFIG);
    const data = sheet.getDataRange().getValues();

    Object.keys(configUpdates).forEach(key => {
      if (key === 'ADMIN_PASSWORD') return;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === key) {
          sheet.getRange(i + 1, 2).setValue(String(configUpdates[key]));
        }
      }
    });

    return { success: true, message: 'Profil sekolah untuk dokumen cetak berhasil diperbarui!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan pengaturan sekolah: ' + err.message };
  }
}
