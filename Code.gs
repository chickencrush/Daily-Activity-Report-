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
      
      if (s === "Settings") sheet.appendRow(["PT Global Distribusi Sejahtera", "08:30", 100]);
      if (s === "Users") {
        sheet.appendRow(["USR-ADMIN-01", "Super Administrator", "admin@company.com", hashPassword("Admin123!"), "Super Admin", "HRD", "Active", new Date().toISOString()]);
      }
      if (s === "Outlets") {
        sheet.appendRow(["GERAI-01", "Gerai Pusat (Kantor Pusat)", "Jakarta Pusat", -6.2088, 106.8456, 100, "Active"]);
        sheet.appendRow(["GERAI-02", "Gerai Cabang Kemang", "Jakarta Selatan", -6.2738, 106.8156, 150, "Active"]);
      }
    }
  }
  return { status: "success", message: "Seluruh 7 Tab Spreadsheet Berhasil Diinisialisasi!" };
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
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// Verifikasi Presensi Multi-Gerai
function handleCheckinPresensi(p) {
  const outlets = getSheetData("Outlets").filter(o => o.Status_Aktif === "Active");
  const targetOutlet = outlets.find(o => o.ID === p.outletId) || outlets[0];
  
  if (!targetOutlet) return { status: "error", message: "Gerai penugasan tidak ditemukan!" };

  const distance = calculateHaversineDistance(Number(p.latitude), Number(p.longitude), Number(targetOutlet.Latitude), Number(targetOutlet.Longitude));
  const maxRadius = Number(targetOutlet.Max_Radius_Meters) || 100;

  if (distance > maxRadius) {
    return { status: "error", message: `Presensi Ditolak! Anda berada di luar radius ${targetOutlet.Nama_Outlet} (${distance}m dari maks ${maxRadius}m).` };
  }

  const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss");
  
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
  const rows = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Settings").getDataRange().getValues();
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

function handleGetKpiSummary() {
  return { status: "success", data: getSheetData("KPI_Karyawan") };
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
