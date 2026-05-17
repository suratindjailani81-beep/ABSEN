// ============================================================
//   SISTEM ABSENSI SISWA ONLINE - Google Apps Script
//   Koordinat Sekolah: 0.5084878, 121.4663607
// ============================================================

// ===== KONFIGURASI =====
var LATITUDE_SEKOLAH  = 0.5084878150913366;
var LONGITUDE_SEKOLAH = 121.4663606701032;
var RADIUS_METER      = 100;

var JAM_MASUK_MULAI   = 6;   // 06.00
var JAM_MASUK_BATAS   = 7;   // 07.30 → terlambat setelah ini
var MENIT_MASUK_BATAS = 30;
var JAM_PULANG_MULAI  = 15;  // 15.00

var SHEET_SISWA   = "DATA_SISWA";
var SHEET_ABSENSI = "ABSENSI";
var SHEET_LOG     = "LOG";

// ============================================================
//   INISIALISASI SHEET (jalankan sekali manual)
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- DATA_SISWA ---
  var shSiswa = ss.getSheetByName(SHEET_SISWA);
  if (!shSiswa) {
    shSiswa = ss.insertSheet(SHEET_SISWA);
    shSiswa.appendRow(["NIS", "Nama Lengkap", "Password", "Tanggal Daftar"]);
    shSiswa.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#4a90d9").setFontColor("#ffffff");
  }

  // --- ABSENSI ---
  var shAbsen = ss.getSheetByName(SHEET_ABSENSI);
  if (!shAbsen) {
    shAbsen = ss.insertSheet(SHEET_ABSENSI);
    shAbsen.appendRow([
      "Hari/Tanggal", "NIS", "Nama Siswa",
      "Waktu Absen Datang", "Waktu Absen Pulang", "Keterangan"
    ]);
    shAbsen.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#4a90d9").setFontColor("#ffffff");
  }

  // --- LOG ---
  var shLog = ss.getSheetByName(SHEET_LOG);
  if (!shLog) {
    shLog = ss.insertSheet(SHEET_LOG);
    shLog.appendRow(["Waktu", "Aksi", "NIS", "Detail"]);
    shLog.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#888888").setFontColor("#ffffff");
  }

  SpreadsheetApp.getUi().alert("✅ Setup selesai! Sheet DATA_SISWA, ABSENSI, dan LOG sudah dibuat.");
}

// ============================================================
//   ENTRY POINT HTTP
// ============================================================
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var aksi   = params.aksi || "";

    switch (aksi) {
      case "register":     return respond(registerSiswa(params.nis, params.nama, params.password));
      case "login":        return respond(loginSiswa(params.nis, params.password));
      case "absenDatang":  return respond(absenDatang(params.nis, params.latitude, params.longitude));
      case "absenPulang":  return respond(absenPulang(params.nis, params.latitude, params.longitude));
      default:             return respond({status:"error", message:"Aksi tidak dikenal: " + aksi});
    }
  } catch (err) {
    tulisLog("ERROR", "-", err.toString());
    return respond({status:"error", message:"Terjadi kesalahan server: " + err.toString()});
  }
}

function doGet(e) {
  try {
    var aksi = e.parameter.aksi || "";
    switch (aksi) {
      case "riwayat": return respond(getRiwayatAbsen(e.parameter.nis));
      case "ping":    return respond({status:"success", message:"Server aktif ✅"});
      default:        return respond({status:"error", message:"Aksi GET tidak dikenal"});
    }
  } catch (err) {
    return respond({status:"error", message:err.toString()});
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//   1. REGISTER SISWA
// ============================================================
function registerSiswa(nis, nama, password) {
  if (!nis || !nama || !password)
    return {status:"error", message:"NIS, nama, dan password wajib diisi"};

  nis = nis.toString().trim();
  nama = nama.toString().trim();

  var sh   = getSheet(SHEET_SISWA);
  var data = sh.getDataRange().getValues();

  // Cek duplikat NIS (mulai baris ke-2, index 1)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nis) {
      return {status:"error", message:"NIS " + nis + " sudah terdaftar. Silakan login."};
    }
  }

  var hash    = hashSederhana(password);
  var tanggal = formatTanggal(new Date());
  sh.appendRow([nis, nama, hash, tanggal]);

  tulisLog("REGISTER", nis, "Siswa baru: " + nama);
  return {status:"success", message:"Pendaftaran berhasil! Selamat datang, " + nama};
}

// ============================================================
//   2. LOGIN SISWA
// ============================================================
function loginSiswa(nis, password) {
  if (!nis || !password)
    return {status:"error", message:"NIS dan password wajib diisi"};

  nis = nis.toString().trim();
  var hash = hashSederhana(password);
  var sh   = getSheet(SHEET_SISWA);
  var data = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nis) {
      if (data[i][2].toString() === hash) {
        tulisLog("LOGIN", nis, "Berhasil login");
        return {
          status:"success",
          message:"Login berhasil",
          data:{nis:data[i][0], nama:data[i][1]}
        };
      } else {
        tulisLog("LOGIN_GAGAL", nis, "Password salah");
        return {status:"error", message:"Password salah"};
      }
    }
  }
  return {status:"error", message:"NIS tidak ditemukan"};
}

// ============================================================
//   3. ABSEN DATANG
// ============================================================
function absenDatang(nis, lat, lng) {
  var cekInput = validasiInputAbsen(nis, lat, lng);
  if (cekInput) return cekInput;

  // Cek lokasi
  var cekLokasi = validasiLokasi(parseFloat(lat), parseFloat(lng));
  if (cekLokasi) return cekLokasi;

  // Cek siswa terdaftar
  var siswa = cariSiswa(nis);
  if (!siswa) return {status:"error", message:"NIS tidak ditemukan"};

  var now      = new Date();
  var jam      = now.getHours();
  var menit    = now.getMinutes();
  var tanggal  = formatTanggal(now);
  var waktuStr = formatWaktu(now);

  // Cek apakah sudah absen datang hari ini
  var baris = cariBasisAbsen(nis, tanggal);
  if (baris && baris[3] !== "") {
    return {status:"error", message:"Anda sudah absen datang hari ini pukul " + baris[3]};
  }

  // Tentukan keterangan
  var keterangan;
  if (jam < JAM_MASUK_MULAI) {
    return {status:"error", message:"Absen datang belum dibuka. Mulai pukul 06.00 WIB"};
  } else if (jam < JAM_MASUK_BATAS || (jam === JAM_MASUK_BATAS && menit === 0)) {
    keterangan = "Datang Tepat Waktu";
  } else {
    keterangan = "Datang Terlambat";
  }

  var sh = getSheet(SHEET_ABSENSI);

  if (baris) {
    // Update baris yang sudah ada (misal sudah dibuat trigger "Tidak Hadir")
    var nomBaris = cariBarisBaris(nis, tanggal);
    sh.getRange(nomBaris, 4).setValue(waktuStr); // Waktu Datang
    sh.getRange(nomBaris, 6).setValue(keterangan);
  } else {
    sh.appendRow([tanggal, nis, siswa.nama, waktuStr, "", keterangan]);
  }

  tulisLog("ABSEN_DATANG", nis, keterangan + " pukul " + waktuStr);
  return {
    status:"success",
    message:"Absen datang berhasil! " + keterangan,
    data:{waktu:waktuStr, keterangan:keterangan}
  };
}

// ============================================================
//   4. ABSEN PULANG
// ============================================================
function absenPulang(nis, lat, lng) {
  var cekInput = validasiInputAbsen(nis, lat, lng);
  if (cekInput) return cekInput;

  var cekLokasi = validasiLokasi(parseFloat(lat), parseFloat(lng));
  if (cekLokasi) return cekLokasi;

  var siswa = cariSiswa(nis);
  if (!siswa) return {status:"error", message:"NIS tidak ditemukan"};

  var now      = new Date();
  var jam      = now.getHours();
  var tanggal  = formatTanggal(now);
  var waktuStr = formatWaktu(now);

  if (jam < JAM_PULANG_MULAI) {
    return {status:"error", message:"Absen pulang baru bisa dilakukan mulai pukul 15.00 WIB"};
  }

  var nomBaris = cariBarisBaris(nis, tanggal);
  if (!nomBaris) {
    return {status:"error", message:"Anda belum absen datang hari ini"};
  }

  var sh   = getSheet(SHEET_ABSENSI);
  var baris = sh.getRange(nomBaris, 1, 1, 6).getValues()[0];

  if (baris[4] !== "") {
    return {status:"error", message:"Anda sudah absen pulang hari ini pukul " + baris[4]};
  }

  sh.getRange(nomBaris, 5).setValue(waktuStr);
  // Jika keterangan masih "Tidak Hadir" atau kosong, update
  var ketLama = baris[5];
  var ketBaru = (ketLama === "Datang Tepat Waktu" || ketLama === "Datang Terlambat")
                ? ketLama : "Datang Tepat Waktu";
  sh.getRange(nomBaris, 6).setValue(ketBaru);

  tulisLog("ABSEN_PULANG", nis, "Pulang pukul " + waktuStr);
  return {
    status:"success",
    message:"Absen pulang berhasil! Pukul " + waktuStr,
    data:{waktu:waktuStr}
  };
}

// ============================================================
//   5. RIWAYAT ABSEN
// ============================================================
function getRiwayatAbsen(nis) {
  if (!nis) return {status:"error", message:"NIS wajib diisi"};
  nis = nis.toString().trim();

  var siswa = cariSiswa(nis);
  if (!siswa) return {status:"error", message:"NIS tidak ditemukan"};

  var sh   = getSheet(SHEET_ABSENSI);
  var data = sh.getDataRange().getValues();
  var hasil = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString() === nis) {
      hasil.push({
        tanggal       : data[i][0].toString(),
        nis           : data[i][1].toString(),
        nama          : data[i][2].toString(),
        waktuDatang   : data[i][3].toString(),
        waktuPulang   : data[i][4].toString(),
        keterangan    : data[i][5].toString()
      });
    }
  }

  return {
    status  : "success",
    message : "Data riwayat absen " + siswa.nama,
    data    : hasil
  };
}

// ============================================================
//   6. TRIGGER HARIAN 23.59 — ISI "TIDAK HADIR"
// ============================================================
function triggerMalam() {
  var shSiswa = getSheet(SHEET_SISWA);
  var shAbsen = getSheet(SHEET_ABSENSI);
  var siswaData = shSiswa.getDataRange().getValues();
  var tanggal   = formatTanggal(new Date());
  var jumlah    = 0;

  for (var i = 1; i < siswaData.length; i++) {
    var nis  = siswaData[i][0].toString();
    var nama = siswaData[i][1].toString();
    if (!nis) continue;

    var sudahAda = cariBasisAbsen(nis, tanggal);
    if (!sudahAda) {
      shAbsen.appendRow([tanggal, nis, nama, "", "", "Tidak Hadir"]);
      jumlah++;
    }
  }

  tulisLog("TRIGGER_MALAM", "-", "Dijalankan. " + jumlah + " siswa ditandai Tidak Hadir");
  Logger.log("Trigger malam selesai. " + jumlah + " siswa Tidak Hadir.");
}

// ============================================================
//   HELPER: VALIDASI & UTILITAS
// ============================================================

function validasiInputAbsen(nis, lat, lng) {
  if (!nis)        return {status:"error", message:"NIS wajib diisi"};
  if (lat === undefined || lat === null || lat === "")
                   return {status:"error", message:"Latitude wajib diisi"};
  if (lng === undefined || lng === null || lng === "")
                   return {status:"error", message:"Longitude wajib diisi"};
  return null;
}

function validasiLokasi(lat, lng) {
  var jarak = hitungHaversine(lat, lng, LATITUDE_SEKOLAH, LONGITUDE_SEKOLAH);
  if (jarak > RADIUS_METER) {
    return {
      status  : "error",
      message : "Anda berada di luar area sekolah. Jarak: " + Math.round(jarak) + " meter (maks 100 m)"
    };
  }
  return null;
}

// Rumus Haversine — menghitung jarak dua koordinat GPS dalam meter
function hitungHaversine(lat1, lon1, lat2, lon2) {
  var R    = 6371000; // radius bumi dalam meter
  var dLat = toRad(lat2 - lat1);
  var dLon = toRad(lon2 - lon1);
  var a    = Math.sin(dLat/2) * Math.sin(dLat/2) +
             Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
             Math.sin(dLon/2) * Math.sin(dLon/2);
  var c    = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

function toRad(deg) { return deg * Math.PI / 180; }

// Hash sederhana — gabungkan karakter dan kalikan nilai ASCII
function hashSederhana(str) {
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return "H" + Math.abs(hash).toString(16).toUpperCase();
}

function cariSiswa(nis) {
  nis = nis.toString().trim();
  var data = getSheet(SHEET_SISWA).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === nis)
      return {nis:data[i][0], nama:data[i][1]};
  }
  return null;
}

// Kembalikan array baris absen [tanggal,nis,nama,datang,pulang,ket] atau null
function cariBasisAbsen(nis, tanggal) {
  nis = nis.toString().trim();
  var data = getSheet(SHEET_ABSENSI).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString() === nis && data[i][0].toString() === tanggal)
      return data[i];
  }
  return null;
}

// Kembalikan nomor baris (1-based) di sheet ABSENSI
function cariBarisBaris(nis, tanggal) {
  nis = nis.toString().trim();
  var data = getSheet(SHEET_ABSENSI).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString() === nis && data[i][0].toString() === tanggal)
      return i + 1;
  }
  return null;
}

function getSheet(nama) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nama);
  if (!sh) sh = ss.insertSheet(nama);
  return sh;
}

function formatTanggal(d) {
  var hari  = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  var h     = hari[d.getDay()];
  var tgl   = pad(d.getDate()) + "/" + pad(d.getMonth()+1) + "/" + d.getFullYear();
  return h + ", " + tgl;
}

function formatWaktu(d) {
  return pad(d.getHours()) + ":" + pad(d.getMinutes()) + ":" + pad(d.getSeconds()) + " WIB";
}

function pad(n) { return n < 10 ? "0" + n : n.toString(); }

function tulisLog(aksi, nis, detail) {
  try {
    var sh = getSheet(SHEET_LOG);
    sh.appendRow([formatWaktu(new Date()), aksi, nis, detail]);
  } catch(e) {}
}
