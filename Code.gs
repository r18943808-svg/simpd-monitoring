/**
 * SIMPD Monitoring KPU Backend - REST API
 * Data dari spreadsheet "Form Responses 1" dan Monitoring Sheets
 */

// ========== KONFIGURASI ==========
const SPREADSHEET_ID = '1ZAR3OKwG28yI-knUtFU6ujZs9e-_ZCkl6w8wYIAfZP4';
const SHEET_NAME = 'Form Responses 1';
const DRIVE_FOLDER_ID = '1nBZmPI-NzNuy6PEIhuCtkdy8mcAjb1vo'; // Folder untuk monitoring sheets

// Nama bulan dalam Bahasa Indonesia
const BULAN_INDONESIA = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

// Header untuk sheet monitoring
const MONITORING_HEADERS = [
  'NO', 'NAMA', 'NIP', 'PANGKAT / GOLONGAN', 'JABATAN',
  'MAKSUD PERJALANAN DINAS', 'ALAT ANGKUTAN YANG DIGUNAKAN',
  'TEMPAT BERANGKAT', 'TEMPAT TUJUAN', 'LAMA SPPD (HARI)',
  'TANGGAL BERANGKAT', 'TANGGAL HARUS KEMBALI',
  'NOMOR SPT', 'TGL SPT & SPD', 'DASAR SPT', 'STATUS', 'LINK FOLDER'
];

// Mapping kolom Form Responses (0-indexed)
const COL = {
  TIMESTAMP: 0, NIP: 1, NAMA: 2, JABATAN: 3, UNIT_KERJA: 4
};

// ========== API HANDLERS ==========

function doGet(e) {
  // Handle direct execution in Apps Script editor (testing)
  if (!e || !e.parameter) {
    return ContentService.createTextOutput(JSON.stringify({
      error: 'No parameters provided. This API must be called via HTTP request.'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  const action = e.parameter.action;
  let result;
  
  try {
    switch (action) {
      case 'getEmployeeList':
        result = getEmployeeList();
        break;
      case 'getEmployeeByNIP':
        result = getEmployeeByNIP(e.parameter.nip);
        break;
      case 'getEmployeeData':
        result = getEmployeeData(e.parameter.name);
        break;
      case 'getReportStats':
        result = getReportStats(e.parameter.name);
        break;
      case 'getRecentReports':
        result = getRecentReports(e.parameter.name, parseInt(e.parameter.limit) || 5);
        break;
      case 'getMonitoringData':
        result = getMonitoringData(e.parameter.bulan, e.parameter.tahun);
        break;
      case 'getMonitoringSheets':
        result = getMonitoringSheets();
        break;
      case 'getAllData':
        result = getAllData();
        break;
      case 'getReportFileDetails':
        result = getReportFileDetails(e.parameter.spreadsheetName, e.parameter.employeeName, e.parameter.tglBerangkat);
        break;
      default:
        result = { 
          error: 'Unknown action', 
          availableActions: ['getEmployeeList', 'getEmployeeByNIP', 'getEmployeeData', 'getReportStats', 'getRecentReports', 'getMonitoringData', 'getMonitoringSheets', 'getAllData'] 
        };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch (action) {
      case 'submitMonitoring':
        result = submitMonitoring(data.payload);
        break;
      case 'submitMonitoringWithFiles': // Legacy - kept for backward compatibility if needed
        result = submitMonitoringWithFiles(data.payload);
        break;
      case 'uploadSingleFile':
        result = handleFileUpload(data.payload);
        break;
      case 'submitData':
        result = submitData(data.payload);
        break;
      case 'completeReport':
        result = completeReport(data.payload);
        break;
      case 'checkReportFolder':
        result = checkReportFolder(data.payload);
        break;
      case 'updateReportStatus':
        result = updateReportStatus(data.payload);
        break;
      default:
        result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== HELPER FUNCTIONS ==========

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet() {
  return getSpreadsheet().getSheetByName(SHEET_NAME);
}

function formatDate(dateVal) {
  if (!dateVal) return '-';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'dd MMM yyyy');
  }
  return dateVal.toString();
}

function formatDateShort(dateVal) {
  if (!dateVal) return '-';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return dateVal.toString();
}

// ========== EMPLOYEE FUNCTIONS ==========

function getEmployeeList() {
  try {
    const sheet = getSheet();
    if (!sheet) return { error: 'Sheet not found: ' + SHEET_NAME };
    
    const data = sheet.getDataRange().getValues();
    const names = [];
    const seenNames = {};
    
    for (let i = 1; i < data.length; i++) {
      const nama = data[i][COL.NAMA];
      if (nama && !seenNames[nama]) {
        seenNames[nama] = true;
        names.push(nama);
      }
    }
    
    return names.sort();
  } catch (e) {
    Logger.log('getEmployeeList error: ' + e);
    return { error: e.toString() };
  }
}

function getEmployeeByNIP(nip) {
  try {
    const sheet = getSheet();
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL.NIP] == nip) {
        return {
          nip: data[i][COL.NIP],
          nama: data[i][COL.NAMA],
          jabatan: data[i][COL.JABATAN] || '-',
          unitKerja: data[i][COL.UNIT_KERJA] || '-'
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('getEmployeeByNIP error: ' + e);
    return null;
  }
}

function getEmployeeData(name) {
  try {
    const sheet = getSheet();
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL.NAMA] === name) {
        return {
          nama: data[i][COL.NAMA],
          nip: data[i][COL.NIP] || '-',
          jabatan: data[i][COL.JABATAN] || '-',
          unitKerja: data[i][COL.UNIT_KERJA] || '-'
        };
      }
    }
    return { nama: name, nip: '-', jabatan: '-', unitKerja: '-' };
  } catch (e) {
    Logger.log('getEmployeeData error: ' + e);
    return { nama: name, nip: '-', jabatan: '-', unitKerja: '-' };
  }
}

// ========== MONITORING FUNCTIONS ==========

/**
 * Get or create the monitoring spreadsheet for a specific month/year
 */
function getOrCreateMonitoringSpreadsheet(bulan, tahun) {
  const sheetName = bulan + ' - ' + tahun;
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  
  // Search for existing spreadsheet
  const files = folder.getFilesByName(sheetName);
  
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.openById(file.getId());
  }
  
  // Create new spreadsheet
  const newSs = SpreadsheetApp.create(sheetName);
  const fileId = newSs.getId();
  
  // Move to target folder
  const file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
  
  // Setup the "Monitoring" sheet (main sheet)
  const monitoringSheet = newSs.getActiveSheet();
  monitoringSheet.setName('Monitoring');
  monitoringSheet.appendRow(MONITORING_HEADERS);
  
  // Format header
  const headerRange = monitoringSheet.getRange(1, 1, 1, MONITORING_HEADERS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#2ebcb3');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
  
  // Set column widths
  monitoringSheet.setColumnWidth(1, 40);   // NO
  monitoringSheet.setColumnWidth(2, 180);  // NAMA
  monitoringSheet.setColumnWidth(3, 180);  // NIP
  monitoringSheet.setColumnWidth(4, 120);  // PANGKAT
  monitoringSheet.setColumnWidth(5, 200);  // JABATAN
  monitoringSheet.setColumnWidth(6, 300);  // MAKSUD
  monitoringSheet.setColumnWidth(7, 180);  // ANGKUTAN
  monitoringSheet.setColumnWidth(8, 120);  // BERANGKAT
  monitoringSheet.setColumnWidth(9, 120);  // TUJUAN
  monitoringSheet.setColumnWidth(10, 80);  // LAMA
  monitoringSheet.setColumnWidth(11, 120); // TGL BERANGKAT
  monitoringSheet.setColumnWidth(12, 120); // TGL KEMBALI
  monitoringSheet.setColumnWidth(13, 120); // NOMOR SPT
  monitoringSheet.setColumnWidth(14, 120); // TGL SPT
  monitoringSheet.setColumnWidth(15, 200); // DASAR SPT
  
  Logger.log('Created new monitoring spreadsheet: ' + sheetName);
  return newSs;
}

/**
 * Get or create employee-specific sheet within a monitoring spreadsheet
 */
function getOrCreateEmployeeSheet(ss, employeeName) {
  let sheet = ss.getSheetByName(employeeName);
  
  if (!sheet) {
    sheet = ss.insertSheet(employeeName);
    sheet.appendRow(MONITORING_HEADERS);
    
    // Format header
    const headerRange = sheet.getRange(1, 1, 1, MONITORING_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#3498db');
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    
    Logger.log('Created new employee sheet: ' + employeeName);
  }
  
  return sheet;
}

/**
 * Submit monitoring data - main function
 */
function submitMonitoring(payload) {
  try {
    // Parse tanggal berangkat to get month/year
    const tglBerangkat = new Date(payload.tglBerangkat);
    const bulan = BULAN_INDONESIA[tglBerangkat.getMonth()];
    const tahun = tglBerangkat.getFullYear();
    
    // Get or create the monitoring spreadsheet
    const ss = getOrCreateMonitoringSpreadsheet(bulan, tahun);
    
    // Get the monitoring sheet (all employees)
    const monitoringSheet = ss.getSheetByName('Monitoring');
    
    // Get or create employee-specific sheet
    const employeeSheet = getOrCreateEmployeeSheet(ss, payload.nama);
    
    // Get or create folder structure (Critical for file uploads)
    const monthYearFolder = getOrCreateUploadFolder(bulan, tahun);
    const employeeFolder = getOrCreateEmployeeFolder(monthYearFolder, payload.tglBerangkat, payload.nama);
    
    // Get next row number for monitoring sheet
    const lastRowMonitoring = monitoringSheet.getLastRow();
    const noUrut = lastRowMonitoring; // Since row 1 is header
    
    const rowData = [
      noUrut,
      payload.nama,
      payload.nip,
      payload.pangkat,
      payload.jabatan,
      payload.maksud,
      payload.angkutan,
      payload.tempatBerangkat,
      payload.tempatTujuan,
      payload.lamaHari,
      payload.tglBerangkat,
      payload.tglKembali,
      payload.nomorSpt,
      payload.tglSpt,
      payload.dasarSpt,
      'Belum Lengkap',           // STATUS - Column 16
      employeeFolder.getUrl()    // LINK FOLDER - Column 17
    ];
    
    // Append to monitoring sheet (all employees)
    monitoringSheet.appendRow(rowData);
    
    // Get next row number for employee sheet
    const lastRowEmployee = employeeSheet.getLastRow();
    rowData[0] = lastRowEmployee; // Update NO for employee sheet
    
    // Append to employee-specific sheet
    employeeSheet.appendRow(rowData);
    
    return { 
      success: true, 
      message: 'Data monitoring berhasil disimpan!',
      spreadsheetName: bulan + ' - ' + tahun,
      spreadsheetUrl: ss.getUrl(),
      folderId: employeeFolder.getId(),
      folderUrl: employeeFolder.getUrl()
    };
    
  } catch (e) {
    Logger.log('submitMonitoring error: ' + e);
    return { success: false, message: e.toString() };
  }
}

/**
 * Get list of available monitoring spreadsheets
 */
function getMonitoringSheets() {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    const sheets = [];
    while (files.hasNext()) {
      const file = files.next();
      sheets.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        lastUpdated: file.getLastUpdated()
      });
    }
    
    // Sort by name (most recent first)
    sheets.sort((a, b) => b.name.localeCompare(a.name));
    
    return sheets;
  } catch (e) {
    Logger.log('getMonitoringSheets error: ' + e);
    return { error: e.toString() };
  }
}

/**
 * Get monitoring data from a specific month/year
 */
function getMonitoringData(bulan, tahun) {
  try {
    const sheetName = bulan + ' - ' + tahun;
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFilesByName(sheetName);
    
    if (!files.hasNext()) {
      return { error: 'Spreadsheet not found: ' + sheetName, data: [] };
    }
    
    const file = files.next();
    const ss = SpreadsheetApp.openById(file.getId());
    const sheet = ss.getSheetByName('Monitoring');
    
    if (!sheet) {
      return { error: 'Monitoring sheet not found', data: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      rows.push(row);
    }
    
    return { 
      spreadsheetName: sheetName,
      spreadsheetUrl: ss.getUrl(),
      headers: headers, 
      data: rows, 
      count: rows.length 
    };
  } catch (e) {
    Logger.log('getMonitoringData error: ' + e);
    return { error: e.toString() };
  }
}

/**
 * Get report statistics for an employee from monitoring data
 */
function getReportStats(name) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    let lengkap = 0;
    let belumLengkap = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      try {
        const ss = SpreadsheetApp.openById(file.getId());
        const sheet = ss.getSheetByName(name); // Employee-specific sheet
        
        if (sheet) {
          const lastRow = sheet.getLastRow();
          if (lastRow > 1) {
            const data = sheet.getDataRange().getValues();
            // Iterate from row 1 (skipping header row 0)
            for (let i = 1; i < data.length; i++) {
               // Column 16 (index 15) is STATUS
               const status = data[i][15];
               if (status === 'Lengkap') {
                 lengkap++;
               } else {
                 belumLengkap++;
               }
            }
          }
        }
      } catch (err) {
        // Skip files that can't be opened
        continue;
      }
    }
    
    return { 
      lengkap: lengkap, 
      belumLengkap: belumLengkap, 
      total: lengkap + belumLengkap 
    };
  } catch (e) {
    Logger.log('getReportStats error: ' + e);
    return { lengkap: 0, belumLengkap: 0, total: 0 };
  }
}

/**
 * Get recent reports for an employee
 */
function getRecentReports(name, limit) {
  try {
    limit = limit || 5;
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    const allReports = [];
    
    while (files.hasNext()) {
      const file = files.next();
      try {
        const ss = SpreadsheetApp.openById(file.getId());
        const sheet = ss.getSheetByName(name);
        
        if (sheet && sheet.getLastRow() > 1) {
          const data = sheet.getDataRange().getValues();
          
          for (let i = 1; i < data.length; i++) {
            allReports.push({
              id: i,
              nama: data[i][1],
              keperluan: data[i][5] || '-', // MAKSUD PERJALANAN
              tglBerangkat: formatDateShort(data[i][10]),
              tglKembali: formatDateShort(data[i][11]),
              status: data[i][15] || 'Belum Lengkap',
              folderUrl: data[i][16] || '',
              source: file.getName()
            });
          }
        }
      } catch (err) {
        continue;
      }
    }
    
    // Sort by date descending
    allReports.sort((a, b) => new Date(b.tglBerangkat) - new Date(a.tglBerangkat));
    
    return allReports.slice(0, limit);
  } catch (e) {
    Logger.log('getRecentReports error: ' + e);
    return [];
  }
}

/**
 * Get all data from form responses (for debugging)
 */
function getAllData() {
  try {
    const sheet = getSheet();
    if (!sheet) return { error: 'Sheet not found: ' + SHEET_NAME };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      rows.push(row);
    }
    
    return { headers: headers, data: rows, count: rows.length };
  } catch (e) {
    Logger.log('getAllData error: ' + e);
    return { error: e.toString() };
  }
}

/**
 * Submit data to form responses (legacy)
 */
function submitData(payload) {
  try {
    const sheet = getSheet();
    if (!sheet) return { success: false, message: 'Sheet not found' };
    
    const timestamp = new Date();
    const rowData = [
      timestamp,
      payload.nip,
      payload.nama,
      payload.jabatan,
      payload.unitKerja
    ];
    
    sheet.appendRow(rowData);
    return { success: true, message: 'Data berhasil disimpan' };
  } catch (e) {
    Logger.log('submitData error: ' + e);
    return { success: false, message: e.toString() };
  }
}

/**
 * Test function - run this to verify connection
 */
function testConnection() {
  const result = getEmployeeList();
  Logger.log('Employee List: ' + JSON.stringify(result));
  return result;
}

/**
 * Test monitoring submission
 */
function testSubmitMonitoring() {
  const testData = {
    nama: 'RINDY',
    nip: '200101012020011001',
    pangkat: 'III/a',
    jabatan: 'Penata Kelola Sistem dan Teknologi Informasi',
    maksud: 'Monitoring dan Evaluasi Sistem Informasi',
    angkutan: 'Pesawat Udara',
    tempatBerangkat: 'Kendari',
    tempatTujuan: 'Jakarta',
    lamaHari: 3,
    tglBerangkat: '2026-01-29',
    tglKembali: '2026-01-31',
    nomorSpt: '001/SPT/KPU/2026',
    tglSpt: '2026-01-28',
    dasarSpt: 'Surat Undangan Rapat Koordinasi'
  };
  
  const result = submitMonitoring(testData);
  Logger.log('Submit result: ' + JSON.stringify(result));
  return result;
}

// ========== FILE UPLOAD FUNCTIONS ==========

/**
 * Get or create folder for file uploads based on month-year
 */
function getOrCreateUploadFolder(bulan, tahun) {
  const parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const folderName = bulan + ' - ' + tahun;
  
  // Search for existing folder
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  
  // Create new folder
  return parentFolder.createFolder(folderName);
}

/**
 * Get or create subfolder for specific employee and date
 */
function getOrCreateEmployeeFolder(monthYearFolder, tglBerangkat, nama) {
  // Format: DD-MM-YYYY - NAMA
  const date = parseDateRobust(tglBerangkat);
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  const folderName = day + '-' + month + '-' + year + ' - ' + nama;
  
  // Search for existing folder
  const folders = monthYearFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  
  // Create new folder
  return monthYearFolder.createFolder(folderName);
}

/**
 * Upload a single file to specified folder
 */
function uploadFile(folder, fileData, fileName) {
  const maxRetries = 3;
  let lastError;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // fileData is base64 encoded
      const decoded = Utilities.base64Decode(fileData.split(',')[1] || fileData);
      const blob = Utilities.newBlob(decoded, 'application/octet-stream', fileName);
      const file = folder.createFile(blob);
      
      return {
        success: true,
        fileId: file.getId(),
        fileName: file.getName(),
        fileUrl: file.getUrl()
      };
    } catch (e) {
      lastError = e;
      Logger.log('uploadFile error (attempt ' + attempt + '): ' + e);
      // Wait before retrying (exponential backoff: 1s, 2s, 4s)
      if (attempt < maxRetries) {
        Utilities.sleep(1000 * Math.pow(2, attempt - 1));
      }
    }
  }
  
  return { success: false, error: lastError.toString() };
}


/**
 * Robust Date Parser for DD/MM/YYYY, YYYY-MM-DD, or Date object
 */
function parseDateRobust(dateInput) {
  if (!dateInput) return new Date();
  if (dateInput instanceof Date) return dateInput;
  
  // Handle DD/MM/YYYY (e.g. 04/02/2026)
  if (typeof dateInput === 'string' && dateInput.includes('/')) {
    const parts = dateInput.split('/');
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // Months are 0-indexed
      const year = parseInt(parts[2], 10);
      return new Date(year, month, day);
    }
  }
  
  // Handle YYYY-MM-DD or other standard formats
  return new Date(dateInput);
}

/**
 * Get details of files currently in the report folder
 */
function getReportFileDetails(spreadsheetName, employeeName, tglBerangkat) {
  try {
    // 1. Find Folder
    // Use the logic to find the specific employee folder
    const tgl = parseDateRobust(tglBerangkat);

    const bulan = BULAN_INDONESIA[tgl.getMonth()];
    const tahun = tgl.getFullYear();
    
    // Check if month folder exists
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const monthFolders = folder.getFoldersByName(bulan + ' - ' + tahun);
    
    if (!monthFolders.hasNext()) {
       return { 
         exists: false, 
         files: [], 
         fileCount: 0, 
         status: 'Folder Bulan tidak ditemukan' 
       };
    }
    
    const monthFolder = monthFolders.next();
    
    // Check employee folder
    const day = String(tgl.getDate()).padStart(2, '0');
    const month = String(tgl.getMonth() + 1).padStart(2, '0');
    const year = tgl.getFullYear();
    const folderName = day + '-' + month + '-' + year + ' - ' + employeeName;
    
    const empFolders = monthFolder.getFoldersByName(folderName);
    
    if (!empFolders.hasNext()) {
      return { 
         exists: false, 
         files: [], 
         fileCount: 0, 
         status: 'Folder Laporan tidak ditemukan'
       };
    }
    
    const empFolder = empFolders.next();
    const files = [];
    const fileIter = empFolder.getFiles();
    
    while (fileIter.hasNext()) {
      const f = fileIter.next();
      files.push({
        name: f.getName(),
        id: f.getId(),
        url: f.getUrl(),
        type: f.getMimeType()
      });
    }
    
    return {
      exists: true,
      files: files,
      fileCount: files.length,
      folderUrl: empFolder.getUrl()
    };
    
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Upload multiple files for a monitoring submission
 * Payload should include: tglBerangkat, nama, files (array of {data, name, type})
 */
function uploadMonitoringFiles(payload) {
  try {
    const tglBerangkat = new Date(payload.tglBerangkat);
    const bulan = BULAN_INDONESIA[tglBerangkat.getMonth()];
    const tahun = tglBerangkat.getFullYear();
    
    // Get or create folder structure
    const monthYearFolder = getOrCreateUploadFolder(bulan, tahun);
    const employeeFolder = getOrCreateEmployeeFolder(monthYearFolder, payload.tglBerangkat, payload.nama);
    
    const uploadedFiles = [];
    const failedFiles = [];
    const files = payload.files || [];
    
    // Define expected file types
    const expectedFileTypes = ['SPT', 'SPPD', 'KUITANSI', 'BOARDING_PASS', 'LAPORAN', 'UNDANGAN'];
    let uploadedCount = 0;
    
    // Upload each file
    for (const file of files) {
      if (file && file.data && file.name) {
        const result = uploadFile(employeeFolder, file.data, file.name);
        if (result.success) {
          uploadedFiles.push(result);
          uploadedCount++;
        } else {
          failedFiles.push({
            name: file.name,
            error: result.error
          });
        }
        // Add safeguard delay between uploads
        Utilities.sleep(1000);
      }
    }
    
    // Determine status
    // If there are failures, it's definitely not complete (or partial)
    const status = uploadedCount >= expectedFileTypes.length ? 'Lengkap' : 'Belum Lengkap';
    
    return {
      success: true,
      message: failedFiles.length > 0 ? 'Beberapa file gagal diupload.' : 'File upload berhasil!',
      folderUrl: employeeFolder.getUrl(),
      folderId: employeeFolder.getId(),
      uploadedFiles: uploadedFiles,
      failedFiles: failedFiles,
      uploadedCount: uploadedCount,
      status: status
    };
    
  } catch (e) {
    Logger.log('uploadMonitoringFiles error: ' + e);
    return { success: false, message: e.toString() };
  }
}

/**
 * Handle single file upload request
 */
function handleFileUpload(payload) {
  try {
    // payload: { folderId, fileData, fileName }
    if (!payload.folderId) {
      return { success: false, message: 'Folder ID missing' };
    }
    
    const folder = DriveApp.getFolderById(payload.folderId);
    return uploadFile(folder, payload.fileData, payload.fileName);
    
  } catch (e) {
    Logger.log('handleFileUpload error: ' + e);
    return { success: false, message: e.toString() };
  }
}

/**
 * Submit monitoring data WITH file uploads (Legacy/Batch)
 */
function submitMonitoringWithFiles(payload) {
  try {
    const tglBerangkat = new Date(payload.tglBerangkat);
    const bulan = BULAN_INDONESIA[tglBerangkat.getMonth()];
    const tahun = tglBerangkat.getFullYear();
    
    // Get or create folder structure for files
    const monthYearFolder = getOrCreateUploadFolder(bulan, tahun);
    const employeeFolder = getOrCreateEmployeeFolder(monthYearFolder, payload.tglBerangkat, payload.nama);
    
    // Upload files if provided
    const files = payload.files || [];
    let uploadedCount = 0;
    const failedFiles = [];
    const expectedFileCount = 6; // SPT, SPPD, KUITANSI, BOARDING_PASS, LAPORAN, UNDANGAN
    
    for (const file of files) {
      if (file && file.data && file.name) {
        const result = uploadFile(employeeFolder, file.data, file.name);
        if (result.success) {
          uploadedCount++;
        } else {
          failedFiles.push({
             name: file.name,
             error: result.error
          });
        }
        // Add safeguard delay between uploads
        Utilities.sleep(1000);
      }
    }
    
    // Determine status based on uploaded files
    const status = uploadedCount >= expectedFileCount ? 'Lengkap' : 'Belum Lengkap';
    const folderUrl = employeeFolder.getUrl();
    
    // Get or create the monitoring spreadsheet
    const ss = getOrCreateMonitoringSpreadsheet(bulan, tahun);
    const monitoringSheet = ss.getSheetByName('Monitoring');
    const employeeSheet = getOrCreateEmployeeSheet(ss, payload.nama);
    
    // Get next row number
    const lastRowMonitoring = monitoringSheet.getLastRow();
    const noUrut = lastRowMonitoring;
    
    // Prepare row data with STATUS and LINK FOLDER
    const rowData = [
      noUrut,
      payload.nama,
      payload.nip,
      payload.pangkat,
      payload.jabatan,
      payload.maksud,
      payload.angkutan,
      payload.tempatBerangkat,
      payload.tempatTujuan,
      payload.lamaHari,
      payload.tglBerangkat,
      payload.tglKembali,
      payload.nomorSpt,
      payload.tglSpt,
      payload.dasarSpt,
      status,
      folderUrl
    ];
    
    // Append to monitoring sheet
    monitoringSheet.appendRow(rowData);
    
    // Update NO and append to employee sheet
    const lastRowEmployee = employeeSheet.getLastRow();
    rowData[0] = lastRowEmployee;
    employeeSheet.appendRow(rowData);
    
    return { 
      success: true, 
      message: failedFiles.length > 0 ? 'Data disimpan, namun beberapa file gagal diupload.' : 'Data monitoring dan file berhasil disimpan!',
      spreadsheetName: bulan + ' - ' + tahun,
      spreadsheetUrl: ss.getUrl(),
      folderUrl: folderUrl,
      status: status,
      uploadedCount: uploadedCount,
      failedFiles: failedFiles
    };
    
  } catch (e) {
    Logger.log('submitMonitoringWithFiles error: ' + e);
    return { success: false, message: e.toString() };
  }
}
/**
 * Complete a report by uploading additional files
 */
function completeReport(payload) {
  try {
    // Payload: { spreadsheetName, employeeName, tglBerangkat, files }
    
    // 1. Open Spreadsheet
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const filesIter = folder.getFilesByName(payload.spreadsheetName);
    
    if (!filesIter.hasNext()) {
      return { success: false, message: 'Spreadsheet tidak ditemukan: ' + payload.spreadsheetName };
    }
    
    const ss = SpreadsheetApp.openById(filesIter.next().getId());
    
    // 2. Find Folder for Uploads
    // We try to find the folder based on naming convention
    const tglBerangkat = new Date(payload.tglBerangkat);
    const bulan = BULAN_INDONESIA[tglBerangkat.getMonth()];
    const tahun = tglBerangkat.getFullYear();
    
    const monthYearFolder = getOrCreateUploadFolder(bulan, tahun);
    const employeeFolder = getOrCreateEmployeeFolder(monthYearFolder, payload.tglBerangkat, payload.employeeName);
    
    // 3. Upload New Files
    const files = payload.files || [];
    let newUploadCount = 0;
    const failedFiles = [];
    
    for (const file of files) {
      if (file && file.data && file.name) {
        // If file object has a 'type' property (e.g. 'SPT', 'SPPD'), prepend it to filename
        // This makes it easier to track which file is which
        let fileName = file.name;
        if (file.docType && !fileName.startsWith(file.docType)) {
             fileName = file.docType + '_' + fileName;
        }
        
        const result = uploadFile(employeeFolder, file.data, fileName);
        if (result.success) {
          newUploadCount++;
        } else {
          failedFiles.push({
             name: file.name,
             error: result.error
          });
        }
        // Add safeguard delay between uploads
        Utilities.sleep(1000);
      }
    }
    
    // 4. Check Total File Count in Folder
    const fileIter = employeeFolder.getFiles();
    let totalFiles = 0;
    while (fileIter.hasNext()) {
      fileIter.next();
      totalFiles++;
    }
    
    // 5. Update Status
    // Rule: At least 6 files required
    const expectedFileCount = 6;
    let newStatus = totalFiles >= expectedFileCount ? 'Lengkap' : 'Belum Lengkap';
    
    // Update in Monitoring Sheet (Main)
    const monitoringSheet = ss.getSheetByName('Monitoring');
    if (monitoringSheet) {
      const data = monitoringSheet.getDataRange().getValues();
      // Logger.log('Searching for row to update. Name: ' + payload.employeeName + ', Date: ' + formatDateShort(tglBerangkat));
      
      for (let i = 1; i < data.length; i++) {
        const rowName = data[i][1];
        const rowDate = data[i][10]; // Tgl Berangkat
        
        // Robust Date Matching
        let dateMatch = false;
        const targetDateStr = formatDateShort(tglBerangkat);
        
        if (rowDate instanceof Date) {
          dateMatch = formatDateShort(rowDate) === targetDateStr;
        } else {
          // If string, try to match directly or parse
          dateMatch = rowDate.toString().includes(targetDateStr) || rowDate == payload.tglBerangkat;
        }
        
        if (rowName === payload.employeeName && dateMatch) {
          // Update Status (Col 16)
          monitoringSheet.getRange(i + 1, 16).setValue(newStatus);
          // Logger.log('Updated Main Sheet Row ' + (i+1));
          break;
        }
      }
    }
    
    // Update in Employee Sheet
    const employeeSheet = ss.getSheetByName(payload.employeeName);
    if (employeeSheet) {
      const data = employeeSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const rowDate = data[i][10];
        
        let dateMatch = false;
        const targetDateStr = formatDateShort(tglBerangkat);
        
        if (rowDate instanceof Date) {
          dateMatch = formatDateShort(rowDate) === targetDateStr;
        } else {
           dateMatch = rowDate.toString().includes(targetDateStr) || rowDate == payload.tglBerangkat;
        }
        
        if (dateMatch) {
           // Update Status (Col 16)
           employeeSheet.getRange(i + 1, 16).setValue(newStatus);
           // Logger.log('Updated Employee Sheet Row ' + (i+1));
           break;
        }
      }
    }
    
    return {
      success: true,
      message: failedFiles.length > 0 ? 'Berkas ditambahkan, namun beberapa gagal.' : 'Berkas berhasil ditambahkan',
      newStatus: newStatus,
      totalFiles: totalFiles,
      uploadedCount: newUploadCount,
      failedFiles: failedFiles
    };
    
  } catch (e) {
    Logger.log('completeReport error: ' + e);
    return { success: false, message: e.toString() };
  }
}

/**
 * Check/Get report folder ID for sequential upload
 */
function checkReportFolder(payload) {
  try {
     const tglBerangkat = parseDateRobust(payload.tglBerangkat);
     const bulan = BULAN_INDONESIA[tglBerangkat.getMonth()];
     const tahun = tglBerangkat.getFullYear();
    
     const monthYearFolder = getOrCreateUploadFolder(bulan, tahun);
     const employeeFolder = getOrCreateEmployeeFolder(monthYearFolder, payload.tglBerangkat, payload.employeeName);
     
     return {
       success: true,
       folderId: employeeFolder.getId(),
       folderUrl: employeeFolder.getUrl()
     };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Update report status after sequential uploads
 */
function updateReportStatus(payload) {
  try {
     // 1. Get Spreadsheet
     const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
     const filesIter = folder.getFilesByName(payload.spreadsheetName);
    
     if (!filesIter.hasNext()) {
       return { success: false, message: 'Spreadsheet not found' };
     }
     
     const ss = SpreadsheetApp.openById(filesIter.next().getId());
     
     // Wait for Drive propagation
     Utilities.sleep(2000);
     
     // 2. Count files in drive folder
     let employeeFolder;
     
     if (payload.folderId) {
       // Use provided ID directly - robust against date parsing issues
       employeeFolder = DriveApp.getFolderById(payload.folderId);
     } else {
       // Fallback logic
       const tglBerangkat = parseDateRobust(payload.tglBerangkat);
       const bulan = BULAN_INDONESIA[tglBerangkat.getMonth()];
       const tahun = tglBerangkat.getFullYear();
       
       const monthYearFolder = getOrCreateUploadFolder(bulan, tahun);
       employeeFolder = getOrCreateEmployeeFolder(monthYearFolder, payload.tglBerangkat, payload.employeeName);
     }
     
     
     
     
     let totalFiles = 0;
     const foundTypes = new Set();
     // Flexible matching types
     const requiredTypes = ['SPT', 'SPPD', 'KUITANSI', 'TIKET', 'HOTEL', 'UNDANGAN', 'DOKUMENTASI', 'LAPORAN']; 
     // Note: TIKET usually matches BOARDING_PASS logic, adjusting to match typical filenames
     
     const debugFiles = [];
     
     const fileIter = employeeFolder.getFiles();
     while (fileIter.hasNext()) {
       const file = fileIter.next();
       const fileName = file.getName().toUpperCase();
       totalFiles++;
       debugFiles.push(fileName);
       
       for (const type of requiredTypes) {
         // Special handling for synonyms
         if (fileName.includes(type) || 
            (type === 'TIKET' && fileName.includes('BOARDING')) ||
            (type === 'LAPORAN' && fileName.includes('HASIL')) ) {
           foundTypes.add(type);
         }
       }
     }
     
     // Determine missing based on Core Requirements (adjust as needed)
     // If you only strictly need 6 specific types:
     const strictRequirements = ['SPT', 'SPPD', 'KUITANSI'];
     // But for now let's rely on what was found vs what is strictly missing from a standard set
     // Determine missing based on Core Requirements
     // ALIGNMENT FIX: Must match frontend REQUIRED_DOCS
     // Frontend: SPT, SPPD, KUITANSI, BOARDING_PASS (TIKET), LAPORAN, UNDANGAN
     const standardSet = ['SPT', 'SPPD', 'KUITANSI', 'TIKET', 'LAPORAN', 'UNDANGAN'];
     
     const missingTypes = standardSet.filter(type => {
        // Check if type was found OR if it has a synonym that was found
        if (foundTypes.has(type)) return false;
        
        // Synonyms
        if (type === 'TIKET' && (foundTypes.has('BOARDING') || foundTypes.has('BOARDING_PASS'))) return false; 
        if (type === 'LAPORAN' && foundTypes.has('HASIL')) return false;
        
        return true;
     });
     
     // Status is Complete ONLY if all required types are found
     let newStatus = missingTypes.length === 0 ? 'Lengkap' : 'Belum Lengkap';
     
     // SAFETY OVERRIDE: If user uploaded 6+ files, assume they are correct to avoid blocking
     if (totalFiles >= 6) {
        newStatus = 'Lengkap';
     }
     
     
     // Update in Monitoring Sheet (Main)
     let sheetUpdateStatus = 'None';
     let debugRowInfo = [];
     
     const monitoringSheet = ss.getSheetByName('Monitoring');
     if (monitoringSheet) {
       const data = monitoringSheet.getDataRange().getValues();
       // Target Date details for robust comparison
       const targetDate = parseDateRobust(payload.tglBerangkat);
       const targetDay = targetDate.getDate();
       const targetMonth = targetDate.getMonth();
       const targetYear = targetDate.getFullYear();
       
       let foundRow = -1;
       
       for (let i = 1; i < data.length; i++) {
         const rowName = (data[i][1] || '').toString().trim().toUpperCase();
         const targetName = (payload.employeeName || '').toString().trim().toUpperCase();
         const rowDateRaw = data[i][10];
         
         // Parse row date robustly
         const rowDate = parseDateRobust(rowDateRaw);
         
         // Compare Day/Month/Year
         const dateMatch = (rowDate.getDate() === targetDay) && 
                           (rowDate.getMonth() === targetMonth) && 
                           (rowDate.getFullYear() === targetYear);
         
         // Flexible name matching (includes or exact)
         const nameMatch = (rowName === targetName) || 
                           rowName.includes(targetName) || 
                           targetName.includes(rowName);
         
         if (nameMatch && dateMatch) {
           monitoringSheet.getRange(i + 1, 16).setValue(newStatus);
           sheetUpdateStatus = 'Updated Row ' + (i + 1);
           foundRow = i + 1;
           break;
         }
         
         // Debug: collect first 5 rows info
         if (i <= 5) {
           debugRowInfo.push({
             row: i + 1,
             name: rowName.substring(0, 20),
             date: rowDateRaw ? rowDateRaw.toString().substring(0, 15) : 'null'
           });
         }
       }
       
       if (foundRow === -1) {
         sheetUpdateStatus = 'Row Not Found. Target: ' + payload.employeeName + ' / ' + payload.tglBerangkat;
       }
     }
     
     // Update in Employee Sheet - CRITICAL: This is what getRecentReports reads!
     // IMPORTANT: Only update the SPECIFIC row for this report, not all rows!
     const employeeSheet = ss.getSheetByName(payload.employeeName);
     if (employeeSheet) {
       const empData = employeeSheet.getDataRange().getValues();
       const targetDate = parseDateRobust(payload.tglBerangkat);
       const targetDay = targetDate.getDate();
       const targetMonth = targetDate.getMonth();
       const targetYear = targetDate.getFullYear();

       let employeeRowUpdated = false;
       
       // Find the row with matching date (column K = TANGGAL BERANGKAT, 0-indexed = 10)
       for (let i = 1; i < empData.length; i++) {
         const rowDateRaw = empData[i][10];
         const rowDate = parseDateRobust(rowDateRaw);
         
         const dateMatch = (rowDate.getDate() === targetDay) && 
                           (rowDate.getMonth() === targetMonth) && 
                           (rowDate.getFullYear() === targetYear);
         
         if (dateMatch) {
           // Only update THIS specific row
           employeeSheet.getRange(i + 1, 16).setValue(newStatus);
           sheetUpdateStatus += ' & Employee Row ' + (i + 1) + ' Updated';
           employeeRowUpdated = true;
           break; // Stop after finding the first match
         }
       }
       
       // Fallback message if date matching failed
       if (!employeeRowUpdated) {
         sheetUpdateStatus += ' | Employee date match failed, tried ' + empData.length + ' rows';
       }
     } else {
       sheetUpdateStatus += ' | Employee Sheet "' + payload.employeeName + '" NOT FOUND';
     }
     
     // RETURN to caller
     return {
       success: true,
       message: 'Status laporan diperbarui',
       status: newStatus,
       totalFiles: totalFiles,
       missingTypes: missingTypes,
       foundTypes: Array.from(foundTypes),
       debugFiles: debugFiles,
       debugFolderId: employeeFolder.getId(),
       sheetUpdateStatus: sheetUpdateStatus,
       debugRowInfo: debugRowInfo
     };
         

     
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
