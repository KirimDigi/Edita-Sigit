/***************************************************************
 * GOOGLE APPS SCRIPT - Buku Tamu & RSVP Edita & Sigit
 * Spreadsheet: https://docs.google.com/spreadsheets/d/1FCIDO7H27j6KAqD9NTVewLJa1q6l78rDZWWBINsw1BQ/edit
 * Sheet: Sheet1
 *
 * Kolom yang dipakai (baris 1 = header, urutan HARUS ini):
 *   A: timestamp | B: nama tamu | C: ucapan | D: konfirmasi kehadiran | E: jumlah tamu
 *
 * CARA PAKAI (sekali saja):
 * 1. Buka spreadsheet di atas > menu Ekstensi > Apps Script.
 * 2. Hapus isi Code.gs, paste SEMUA file ini, lalu Save (Ctrl+S).
 * 3. Jalankan sekali fungsi setupSheet() > klik Run > izinkan akses.
 *    (Fungsi ini hanya memastikan baris header benar.)
 * 4. Deploy > New deployment > tipe Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone  (PENTING agar web undangan bisa akses)
 * 5. Copy URL Web app yang muncul, misal:
 *    https://script.google.com/macros/s/AKfyc.../exec
 * 6. Buka index.html, cari GS_URL lalu ganti dengan URL tersebut.
 ***************************************************************/

const SPREADSHEET_ID = '1FCIDO7H27j6KAqD9NTVewLJa1q6l78rDZWWBINsw1BQ';
const SHEET_NAME = 'Sheet1';
const HEADERS = ['timestamp', 'nama tamu', 'ucapan', 'konfirmasi kehadiran', 'jumlah tamu'];

// Dipanggil manual sekali untuk memastikan header benar
function setupSheet() {
  const sheet = getSheet_();
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needHeader = firstRow.join('').trim() === '';
  if (needHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return 'Header OK: ' + HEADERS.join(' | ');
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

// Normalisasi label kehadiran dari web ke label rapi
function normAttendance_(v) {
  v = String(v || '').trim().toLowerCase();
  if (v === 'present' || v === 'hadir') return 'Hadir';
  if (v === 'notpresent' || v === 'tidak hadir' || v === 'tidak_hadir') return 'Tidak Hadir';
  if (v === 'notsure' || v === 'masih ragu' || v === 'ragu') return 'Masih Ragu';
  return v || '-';
}

/**
 * GET  -> ?action=list  (atau tanpa param) mengembalikan JSON daftar ucapan
 * Contoh respon: { "status":"ok", "count": 3, "data":[ {...}, ... ] }
 */
function doGet(e) {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    const data = [];
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      for (let i = 0; i < values.length; i++) {
        const r = values[i];
        // Lewati baris kosong
        if (!r[1] && !r[2]) continue;
        data.push({
          timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(r[0] || ''),
          nama: String(r[1] || ''),
          ucapan: String(r[2] || ''),
          kehadiran: String(r[3] || ''),
          jumlah: String(r[4] || '')
        });
      }
      data.reverse(); // terbaru di atas
    }
    return jsonOut_({ status: 'ok', count: data.length, data: data });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

/**
 * POST -> menyimpan 1 baris ucapan baru.
 * Body JSON yang diharapkan dari web:
 * { nama, ucapan, kehadiran, jumlah }
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err2) {
        payload = (e.parameter) || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const nama = String(payload.nama || payload.name || payload.author || '').trim();
    const ucapan = String(payload.ucapan || payload.comment || payload.message || '').trim();
    let kehadiran = normAttendance_(payload.kehadiran || payload.attendance || '');
    let jumlah = String(payload.jumlah || payload.guest || payload.jumlahTamu || '').trim();

    if (nama.length < 1) return jsonOut_({ status: 'error', message: 'Nama wajib diisi.' });
    if (ucapan.length < 2) return jsonOut_({ status: 'error', message: 'Ucapan minimal 2 karakter.' });
    if (!kehadiran || kehadiran === '-') return jsonOut_({ status: 'error', message: 'Pilih konfirmasi kehadiran.' });
    // Jumlah tamu hanya wajib saat Hadir
    if (kehadiran === 'Hadir' && !jumlah) return jsonOut_({ status: 'error', message: 'Pilih jumlah tamu.' });
    if (!jumlah) jumlah = '-';

    const sheet = getSheet_();
    // Pastikan header ada
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }

    const now = new Date();
    sheet.appendRow([now, nama, ucapan, kehadiran, jumlah]);

    return jsonOut_({
      status: 'ok',
      message: 'Terima kasih atas ucapan Anda!',
      data: {
        timestamp: Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
        nama: nama,
        ucapan: ucapan,
        kehadiran: kehadiran,
        jumlah: jumlah
      }
    });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
