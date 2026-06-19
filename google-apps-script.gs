/**
 * Kanvas Inovasi Probis - jembatan ke Google Sheet
 *
 * Skrip ini membuat halaman kanvas bisa menyimpan isian peserta ke Google
 * Sheet milik Anda, dan halaman Cek Hasil membacanya kembali. Jadi peserta
 * boleh memakai laptop masing-masing, lalu Anda memantau semuanya terpusat.
 *
 * CARA PASANG (sekali saja, sekitar lima menit):
 * 1. Buka Google Sheet baru di akun Anda. Beri nama bebas, misalnya
 *    "Kanvas Inovasi Probis".
 * 2. Di menu, klik Extensions lalu Apps Script.
 * 3. Hapus kode contoh yang ada, lalu tempel SELURUH isi file ini.
 * 4. Klik Save (ikon disket).
 * 5. Klik Deploy lalu New deployment.
 * 6. Pada "Select type" pilih Web app.
 * 7. Isi Description bebas. Pada "Execute as" pilih Me. Pada "Who has access"
 *    pilih Anyone. Klik Deploy.
 * 8. Saat diminta, klik Authorize access dan izinkan dengan akun Anda.
 *    (Bila muncul peringatan "Google hasn't verified this app", klik Advanced
 *    lalu Go to ... untuk melanjutkan. Ini wajar karena skrip milik Anda
 *    sendiri.)
 * 9. Salin "Web app URL" yang muncul. Bentuknya seperti
 *    https://script.google.com/macros/s/XXXXXXXX/exec
 * 10. Tempel URL itu ke dalam index.html dan dashboard.html pada baris
 *     var SHEET_URL = ""; menjadi var SHEET_URL = "URL_ANDA";
 *
 * Bila nanti Anda mengubah skrip ini, jalankan Deploy lalu Manage deployments,
 * pilih deployment yang ada, klik edit (pensil), ubah Version ke New version,
 * lalu Deploy. URL tetap sama.
 */

var NAMA_SHEET = 'Data';

var KOLOM = [
  'Waktu', 'Nama', 'Proses yang dipilih', 'Pengguna dan kebutuhannya',
  'Titik nyeri', 'Akar masalah', 'Ide desain ulang', 'Nilai tambah',
  'Sumber daya dan mitra', 'Uji coba cepat', 'Indikator keberhasilan'
];

var URUTAN_BLOK = [
  'proses', 'pengguna', 'nyeri', 'akar', 'ide',
  'nilai', 'sumberdaya', 'uji', 'indikator'
];

function ambilSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NAMA_SHEET);
  if (!sheet) sheet = ss.insertSheet(NAMA_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, KOLOM.length).setValues([KOLOM]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Tulis atau perbarui satu setoran, dikunci agar aman saat banyak peserta
function handleWrite(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = ambilSheet();
    var nama = String(data.nama || '').trim();
    if (nama === '') return { ok: false, error: 'nama kosong' };

    var blok = data.blok || {};
    var baris = [new Date()];
    baris.push(nama);
    for (var i = 0; i < URUTAN_BLOK.length; i++) {
      baris.push(String(blok[URUTAN_BLOK[i]] || ''));
    }

    // Cari nama yang sama agar setoran terbaru menimpa yang lama
    var nilai = sheet.getDataRange().getValues();
    var temu = -1;
    for (var r = 1; r < nilai.length; r++) {
      if (String(nilai[r][1]).trim().toLowerCase() === nama.toLowerCase()) {
        temu = r + 1;
        break;
      }
    }
    if (temu > 0) sheet.getRange(temu, 1, 1, baris.length).setValues([baris]);
    else sheet.appendRow(baris);

    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

// Baca semua setoran menjadi array objek
function handleRead() {
  var sheet = ambilSheet();
  var nilai = sheet.getDataRange().getValues();
  var keluar = [];
  for (var r = 1; r < nilai.length; r++) {
    var row = nilai[r];
    if (!row[1]) continue;
    var blok = {};
    for (var i = 0; i < URUTAN_BLOK.length; i++) {
      blok[URUTAN_BLOK[i]] = String(row[i + 2] || '');
    }
    var waktu = row[0] ? new Date(row[0]).getTime() : 0;
    keluar.push({ nama: String(row[1]), ts: waktu, blok: blok });
  }
  return keluar;
}

function doPost(e) {
  var hasil;
  try {
    hasil = handleWrite(JSON.parse(e.postData.contents));
  } catch (err) {
    hasil = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(hasil))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var data = handleRead();
  var json = JSON.stringify(data);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
