/**
 * ============================================================================
 * SISTEM LAPORAN HARIAN & PRESENSI MULTI-GERAI (Code.gs)
 * Backend REST API Google Apps Script (GAS)
 * Basis Data: Google Sheets | Penyimpanan Media: Google Drive
 * ============================================================================
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const FOLDER_PRESENSI_NAME = "DAILY_REPORT_PRESENSI_PHOTOS";
const FOLDER_IZIN_NAME = "DAILY_REPORT_IZIN_DOCS";
const FOLDER_BUKTI_NAME = "DAILY_REPORT_BUKTI_TUGAS";
const TZ = "Asia/Jakarta";

function doGet(e) { return handleRequest(e, "GET"); }
function doPost(e) { return handleRequest(e, "POST"); }

function handleRequest(e, method) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000); // Mencegah konflik penulisan bersamaan

  try {
    let p = {};
    if (method === "POST" && e.postData && e.postData.contents) {
      p = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      p = e.parameter;
    }

    const action = p.action || (e.parameter ? e.parameter.action : "");
    let res = { status: "error", message: "Aksi tidak dikenali" };

    switch (action) {
      case "init_database": res = initDatabase(); break;
      case "register": res = handleRegister(p); break;
      case "login": res = handleLogin(p); break;
      case "get_users": res = handleGetUsers(); break;
      case "approve_user": res = handleApproveUser(p); break;
      case "get_settings": res = handleGetSettings(); break;
      case "update_settings": res = handleUpdateSettings(p); break;
      case "get_outlets": res = handleGetOutlets(); break;
      case "add_outlet": res = handleAddOutlet(p); break;
      case "delete_outlet": res = handleDeleteOutlet(p); break;
      case "checkin_presensi": res = handleCheckinPresensi(p); break;
      case "get_presensi_history": res = handleGetPresensiHistory(p); break;
      case "get_dept_report": res = handleGetDeptReport(p); break;
      case "get_tasks": res = handleGetTasks(p); break;
      case "save_task": res = handleSaveTask(p); break;
      case "update_task_status": res = handleUpdateTaskStatus(p); break;
      case "get_daily_reports": res = handleGetDailyReports(p); break;
      case "save_daily_report": res = handleSaveDailyReport(p); break;
      case "update_report_status": res = handleUpdateReportStatus(p); break;
      case "submit_izin": res = handleSubmitIzin(p); break;
      case "get_izin_list": res = handleGetIzinList(p); break;
      case "respond_izin": res = handleRespondIzin(p); break;
      case "get_kpi_summary": res = handleGetKpiSummary(p); break;
      case "get_dashboard_summary": res = handleGetDashboardSummary(p); break;
      default: res = { status: "error", message: "Aksi tidak valid: " + action };
    }

    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Inisialisasi Otomatis 7 Tab Basis Data Spreadsheet
function initDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const schemas = {
    "Users": ["ID", "Nama", "Email", "Password", "Role", "Divisi", "Status", "Created_At"],
    "Daily_Tasks": ["ID", "User_ID", "Nama_Karyawan", "Divisi", "Tanggal", "Judul", "Deskripsi", "Status", "Bukti_URL", "Catatan", "Created_At", "Updated_At"],
    "Daily_Reports": ["ID", "User_ID", "Nama_Karyawan", "Divisi", "Tanggal", "Jam", "Task_Open", "Task_Proses", "Task_Selesai", "Status", "Updated_At"],
    "Izin_Cuti": ["ID", "User_ID", "Nama_Karyawan", "Tipe", "Alasan", "Tanggal_Mulai", "Tanggal_Selesai", "Lampiran_URL", "Status_Approval", "Approved_By_HRD", "Catatan_HRD", "Created_At"],
    "Presensi": ["ID", "User_ID", "Nama_Karyawan", "Divisi", "Tanggal", "Jam_Masuk", "Lokasi_Gerai", "Foto_Selfie_URL", "Latitude", "Longitude", "Distance_Meters", "Status_Radius", "Status_Kehadiran"],
    "Settings": ["Nama_Perusahaan", "Work_Start_Time", "Default_Max_Radius"],
    "Outlets": ["ID", "Nama_Outlet", "Alamat", "Latitude", "Longitude", "Max_Radius_Meters", "Status_Aktif"],
    "KPI_Karyawan": ["User_ID", "Nama", "Divisi", "Periode_Bulan", "Hari_Kerja", "Total_Hadir", "Total_Late", "Task_Target", "Task_Selesai", "Nilai_Target_Divisi", "Alpha_Pelanggaran", "Skor_KPI_Total", "Grade_KPI", "Status_Evaluasi"]
  };

  for (let s in schemas) {
    let sheet = ss.getSheetByName(s);
    if (!sheet) {
      sheet = ss.insertSheet(s);
      sheet.appendRow(schemas[s]);
      sheet.getRange(1, 1, 1, schemas[s].length).setFontWeight("bold").setBackground("#0a0f2b").setFontColor("#38bdf8");
      schemas[s].forEach((h, i) => { if (/^(Tanggal|Jam|Work_Start)/.test(h)) sheet.getRange(1, i + 1, sheet.getMaxRows(), 1).setNumberFormat("@"); });

      if (s === "Settings") sheet.appendRow(["Chicken Crush", "08:30", 100]);
      if (s === "Users") {
        // Akun Super Admin awal: sandi acak, dicatat sekali di Log eksekusi. Segera login & simpan, lalu ganti email.
        const tmpPass = Utilities.getUuid().replace(/-/g, "").slice(0, 12);
        sheet.appendRow(["USR-ADMIN-01", "Super Administrator", "admin@chickencrush.local", hashPassword(tmpPass), "Super Admin", "HRD", "Active", new Date().toISOString()]);
        Logger.log("SUPER ADMIN -> email: admin@chickencrush.local | sandi sementara: " + tmpPass);
      }
    }
  }
  return { status: "success", message: "Tab spreadsheet siap. Lihat Log eksekusi (View > Logs) untuk sandi Super Admin awal." };
}

// Formula Haversine Server-Side (Mencegah Pemalsuan Lokasi)
function calculateHaversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000; // Radius Bumi dalam meter
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2)*Math.sin(dLat/2) + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)*Math.sin(dLon/2);
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function hashPassword(str) {
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return hash.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function saveBase64ToDrive(base64Data, filename, folderName) {
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const clean = base64Data.replace(/^data:image\/(png|jpeg|jpg);base64,/, "");
  const blob = Utilities.newBlob(Utilities.base64Decode(clean), "image/jpeg", filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, TZ, /^Jam/.test(h) ? "HH:mm:ss" : "yyyy-MM-dd");
      obj[h] = v;
    });
    return obj;
  });
}

// Verifikasi Presensi Multi-Gerai
function handleCheckinPresensi(p) {
  const outlets = getSheetData("Outlets").filter(o => o.Status_Aktif === "Active");
  const targetOutlet = outlets.find(o => o.ID === p.outletId) || outlets[0];
  
  if (!p.fotoSelfieBase64) return { status: "error", message: "Foto swafoto wajib diambil langsung dari kamera." };
  if (!targetOutlet) return { status: "error", message: "Gerai penugasan tidak ditemukan!" };

  const distance = calculateHaversineDistance(Number(p.latitude), Number(p.longitude), Number(targetOutlet.Latitude), Number(targetOutlet.Longitude));
  const maxRadius = Number(targetOutlet.Max_Radius_Meters) || 100;

  if (distance > maxRadius) {
    return { status: "error", message: `Presensi Ditolak! Anda berada di luar radius ${targetOutlet.Nama_Outlet} (${distance}m dari maks ${maxRadius}m).` };
  }

  const todayStr = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(new Date(), TZ, "HH:mm:ss");
  
  const existing = getSheetData("Presensi");
  if (existing.some(r => r.User_ID === p.userId && r.Tanggal === todayStr)) {
    return { status: "error", message: "Anda sudah melakukan presensi masuk hari ini!" };
  }

  const settings = handleGetSettings().data;
  const fotoUrl = saveBase64ToDrive(p.fotoSelfieBase64, `PRESENSI_${p.userId}_${todayStr}.jpg`, FOLDER_PRESENSI_NAME);
  const statusKehadiran = timeStr > settings.workStartTime ? "Late" : "Present";
  const newId = "PRS-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Presensi");
  sheet.appendRow([newId, p.userId, p.namaKaryawan, p.divisi, todayStr, timeStr, targetOutlet.Nama_Outlet, fotoUrl, Number(p.latitude), Number(p.longitude), distance, "Dalam Radius", statusKehadiran]);

  return {
    status: "success",
    message: `Presensi Berhasil (${statusKehadiran}) di ${targetOutlet.Nama_Outlet}! Jarak: ${distance}m.`,
    data: { id: newId, jamMasuk: timeStr, distance: distance, statusKehadiran: statusKehadiran, fotoUrl: fotoUrl }
  };
}

function handleGetOutlets() { return { status: "success", data: getSheetData("Outlets") }; }
function handleAddOutlet(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Outlets");
  const newId = "GERAI-0" + (sheet.getLastRow());
  sheet.appendRow([newId, p.namaOutlet, p.alamat || "-", Number(p.latitude), Number(p.longitude), Number(p.maxRadius) || 100, "Active"]);
  return { status: "success", message: `Gerai ${p.namaOutlet} berhasil didaftarkan!`, id: newId };
}

function handleDeleteOutlet(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Outlets");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.outletId) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "Gerai berhasil dihapus" };
    }
  }
  return { status: "error", message: "Gerai tidak ditemukan" };
}

function handleGetSettings() {
  const rows = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Settings").getDataRange().getDisplayValues();
  return {
    status: "success",
    data: {
      namaPerusahaan: rows[1][0],
      workStartTime: rows[1][1] || "08:30",
      defaultMaxRadius: Number(rows[1][2]) || 100
    }
  };
}

function handleUpdateSettings(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Settings");
  sheet.getRange(2, 2).setNumberFormat("@");
  sheet.getRange(2, 1, 1, 3).setValues([[p.namaPerusahaan, p.workStartTime || "08:30", Number(p.defaultMaxRadius) || 100]]);
  return { status: "success", message: "Pengaturan berhasil diperbarui" };
}

function handleLogin(p) {
  const users = getSheetData("Users");
  const user = users.find(u => u.Email.toLowerCase() === p.email.toLowerCase() && u.Password === hashPassword(p.password));
  if (!user) return { status: "error", message: "Email atau Password salah!" };
  if (user.Status !== "Active") return { status: "error", message: "Akun Pending! Menunggu persetujuan HRD/Super Admin." };
  return { status: "success", user: { id: user.ID, nama: user.Nama, email: user.Email, role: user.Role, divisi: user.Divisi } };
}

function handleRegister(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Users");
  const users = getSheetData("Users");
  if (users.some(u => u.Email.toLowerCase() === p.email.toLowerCase())) return { status: "error", message: "Email sudah terdaftar!" };
  const newId = "USR-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
  sheet.appendRow([newId, p.nama, p.email, hashPassword(p.password), p.role || "KARYAWAN", p.divisi, "Pending", new Date().toISOString()]);
  return { status: "success", message: "Pendaftaran berhasil! Menunggu persetujuan aktivasi oleh HRD/Super Admin." };
}

function handleGetUsers() { return { status: "success", data: getSheetData("Users") }; }
function handleApproveUser(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.userId) {
      sheet.getRange(i + 1, 7).setValue(p.status || "Active");
      return { status: "success", message: "Status User Berhasil Diperbarui" };
    }
  }
  return { status: "error", message: "User tidak ditemukan" };
}

function handleSaveDailyReport(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Daily_Reports");
  const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss");
  const newId = "RPT-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
  sheet.appendRow([newId, p.userId, p.namaKaryawan, p.divisi, todayStr, timeStr, p.taskOpen, p.taskProses, p.taskSelesai, p.status || "OPEN", new Date().toISOString()]);
  return { status: "success", message: "Laporan harian berhasil disimpan" };
}
function handleGetDailyReports(p) { return { status: "success", data: getSheetData("Daily_Reports").reverse() }; }
function handleUpdateReportStatus(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Daily_Reports");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.reportId) {
      if (p.status) sheet.getRange(i + 1, 10).setValue(p.status);
      return { status: "success", message: "Status Report Berhasil Diubah" };
    }
  }
  return { status: "error", message: "Report tidak ditemukan" };
}

function handleSubmitIzin(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Izin_Cuti");
  let url = p.lampiranBase64 ? saveBase64ToDrive(p.lampiranBase64, `IZIN_${p.userId}_${Date.now()}.jpg`, FOLDER_IZIN_NAME) : "-";
  const newId = "IZN-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
  sheet.appendRow([newId, p.userId, p.namaKaryawan, p.tipe, p.alasan, p.tanggalMulai, p.tanggalSelesai, url, "Pending", "-", "-", new Date().toISOString()]);
  return { status: "success", message: "Pengajuan izin berhasil diajukan" };
}
function handleGetIzinList() { return { status: "success", data: getSheetData("Izin_Cuti").reverse() }; }
function handleRespondIzin(p) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Izin_Cuti");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.izinId) {
      sheet.getRange(i + 1, 9).setValue(p.statusApproval);
      sheet.getRange(i + 1, 10).setValue(p.hrdName || "HRD");
      return { status: "success", message: "Status pengajuan izin berhasil diupdate" };
    }
  }
  return { status: "error", message: "ID Izin tidak ditemukan" };
}

function handleGetDeptReport(p) {
  const presensi = getSheetData("Presensi");
  let filtered = presensi;
  if (p.divisi && p.divisi !== "Semua") filtered = filtered.filter(pr => pr.Divisi === p.divisi);
  if (p.tanggal) filtered = filtered.filter(pr => pr.Tanggal === p.tanggal);
  return { status: "success", data: filtered.reverse() };
}

// KPI bulanan otomatis. Asumsi: hari libur mingguan = Minggu (ubah HARI_LIBUR bila perlu).
const HARI_LIBUR = [0];
function handleGetKpiSummary(p) {
  const now = new Date();
  const periode = /^\d{4}-\d{2}$/.test(p.periode || "") ? p.periode : Utilities.formatDate(now, TZ, "yyyy-MM");
  const y = Number(periode.slice(0, 4)), m = Number(periode.slice(5, 7));
  const today = Utilities.formatDate(now, TZ, "yyyy-MM-dd");
  const lastDay = new Date(y, m, 0).getDate();
  const pad = n => ("0" + n).slice(-2);
  const days = [];  // hari kerja yang sudah berjalan
  for (let d = 1; d <= lastDay; d++) {
    const ds = periode + "-" + pad(d);
    if (ds > today) break;
    if (HARI_LIBUR.indexOf(new Date(y, m - 1, d).getDay()) < 0) days.push(ds);
  }
  const users = getSheetData("Users").filter(u => u.Status === "Active" && u.Role !== "Super Admin");
  const presensi = getSheetData("Presensi").filter(r => String(r.Tanggal).slice(0, 7) === periode);
  const tasks = getSheetData("Daily_Tasks").filter(t => String(t.Tanggal).slice(0, 7) === periode);
  const izin = getSheetData("Izin_Cuti").filter(i => i.Status_Approval === "Approved");

  // Target divisi = rasio tugas selesai seluruh anggota divisi pada periode (0-100)
  const divStat = {};
  tasks.forEach(t => { const d = divStat[t.Divisi] = divStat[t.Divisi] || { total: 0, done: 0 }; d.total++; if (t.Status === "SELESAI") d.done++; });

  const data = users.map(u => {
    const mine = presensi.filter(r => r.User_ID === u.ID);
    const hadirDays = {}; mine.forEach(r => hadirDays[r.Tanggal] = true);
    const izinDays = days.filter(ds => izin.some(i => i.User_ID === u.ID && ds >= i.Tanggal_Mulai && ds <= i.Tanggal_Selesai));
    const hariKerja = Math.max(0, days.length - izinDays.length);
    const hadir = Object.keys(hadirDays).filter(ds => izinDays.indexOf(ds) < 0).length;
    const myTasks = tasks.filter(t => t.User_ID === u.ID);
    const ds = divStat[u.Divisi];
    return {
      userId: u.ID, nama: u.Nama, divisi: u.Divisi, hariKerja: hariKerja, hadir: Math.min(hadir, hariKerja),
      late: mine.filter(r => r.Status_Kehadiran === "Late").length,
      taskTotal: myTasks.length, taskSelesai: myTasks.filter(t => t.Status === "SELESAI").length,
      targetDivisi: ds && ds.total ? Math.round(ds.done / ds.total * 100) : 0,
      alpha: Math.max(0, hariKerja - hadir)
    };
  });
  return { status: "success", periode: periode, data: data };
}

function handleGetDashboardSummary(p) {
  const todayStr = p.tanggal || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  const users = getSheetData("Users").filter(u => u.Status === "Active");
  const presensi = getSheetData("Presensi").filter(pr => pr.Tanggal === todayStr);
  const izinList = getSheetData("Izin_Cuti").filter(iz => iz.Status_Approval === "Approved" && todayStr >= iz.Tanggal_Mulai && todayStr <= iz.Tanggal_Selesai);

  const totalEmployees = users.length || 1;
  const presentCount = presensi.filter(pr => pr.Status_Kehadiran === "Present").length;
  const lateCount = presensi.filter(pr => pr.Status_Kehadiran === "Late").length;
  const leaveCount = izinList.length;
  const absentCount = Math.max(0, totalEmployees - (presentCount + lateCount + leaveCount));

  return {
    status: "success",
    data: {
      metrics: {
        totalEmployees,
        present: { count: presentCount, percent: ((presentCount / totalEmployees) * 100).toFixed(2) },
        absent: { count: absentCount, percent: ((absentCount / totalEmployees) * 100).toFixed(2) },
        late: { count: lateCount, percent: ((lateCount / totalEmployees) * 100).toFixed(2) },
        leave: { count: leaveCount, percent: ((leaveCount / totalEmployees) * 100).toFixed(2) }
      },
      recentAttendance: presensi.slice(-6).reverse()
    }
  };
}


function handleGetPresensiHistory(p) {
  const rows = getSheetData("Presensi").filter(r => r.User_ID === p.userId).reverse().slice(0, 30);
  return { status: "success", data: rows };
}

// ===== TUGAS HARIAN: PROSES -> HOLD -> SELESAI (Hold & Selesai wajib bukti) =====
function handleSaveTask(p) {
  if (!p.userId || !p.judul) return { status: "error", message: "Judul tugas wajib diisi." };
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Daily_Tasks");
  const now = new Date();
  const id = "TSK-" + Utilities.formatDate(now, TZ, "yyyyMMdd") + "-" + Math.floor(10000 + Math.random() * 90000);
  sheet.appendRow([id, p.userId, p.namaKaryawan, p.divisi, Utilities.formatDate(now, TZ, "yyyy-MM-dd"), p.judul, p.deskripsi || "-", "PROSES", "-", "-", now.toISOString(), now.toISOString()]);
  return { status: "success", message: "Laporan tersimpan", id: id };
}

function handleGetTasks(p) {
  const cutoff = Utilities.formatDate(new Date(Date.now() - 7 * 86400000), TZ, "yyyy-MM-dd");
  let rows = getSheetData("Daily_Tasks").filter(t => t.Status !== "SELESAI" || t.Tanggal >= cutoff);
  if (String(p.all) !== "1") rows = rows.filter(t => t.User_ID === p.userId);
  return { status: "success", data: rows.reverse() };
}

function handleUpdateTaskStatus(p) {
  const allowed = ["PROSES", "HOLD", "SELESAI"];
  if (allowed.indexOf(p.status) < 0) return { status: "error", message: "Status tidak valid." };
  const link = String(p.buktiLink || "").trim();
  const hasLink = /^https?:\/\/\S+$/i.test(link);
  if ((p.status === "HOLD" || p.status === "SELESAI") && !hasLink && !p.buktiBase64) {
    return { status: "error", message: "Status " + p.status + " wajib disertai bukti (tautan atau foto)." };
  }
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Daily_Tasks");
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== p.taskId) continue;
    if (rows[i][1] !== p.userId) return { status: "error", message: "Hanya pembuat laporan yang dapat mengubah statusnya." };
    const bukti = [];
    if (hasLink) bukti.push(link);
    if (p.buktiBase64) bukti.push(saveBase64ToDrive(p.buktiBase64, "BUKTI_" + p.taskId + "_" + Date.now() + ".jpg", FOLDER_BUKTI_NAME));
    const r = i + 1;
    sheet.getRange(r, 8).setValue(p.status);
    if (bukti.length) {
      const old = String(rows[i][8] || "-");
      sheet.getRange(r, 9).setValue((old !== "-" ? old + " | " : "") + bukti.join(" | "));
    }
    if (p.catatan) sheet.getRange(r, 10).setValue(p.catatan);
    sheet.getRange(r, 12).setValue(new Date().toISOString());
    return { status: "success", message: "Status diperbarui" };
  }
  return { status: "error", message: "Tugas tidak ditemukan" };
}
