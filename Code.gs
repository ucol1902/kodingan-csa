/**
 * SelesaiSurvei - Google Apps Script Backend (Code.gs)
 * ID Spreadsheet Terpaut: 1TL94WviGR7-SyIGP10p7KeVLTmtpe7wpsb0c-0W6XIA
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  var userParam = '';
  if (e && e.parameter) {
    userParam = e.parameter.user || e.parameter.email || e.parameter.username || '';
  }
  template.csaUser = userParam.trim();
  
  return template.evaluate()
    .setTitle('SelesaiSurvei - Cybersecurity Awareness Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const SPREADSHEET_ID = '1TL94WviGR7-SyIGP10p7KeVLTmtpe7wpsb0c-0W6XIA';

/**
 * Helper dynamically adapter to use the bound active sheet if present, or fallback to SPREADSHEET_ID.
 */
function getActiveSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      return ss;
    }
  } catch (e) {
    Logger.log('getActiveSpreadsheet failed, falling back to ID: ' + e.toString());
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * 1. Mengambil data pegawai aktif yang belum survey/kuis dari spreadsheet
 */
function cleanNipBackend(v) {
  return String(v || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function isNipValidBackend(nipClean) {
  return nipClean && nipClean !== '0' && nipClean !== '-' && nipClean !== 'TIDAKADA' && nipClean !== 'BELUMADA' && nipClean.length >= 3;
}

function getEmployeeUniqueKeyBackend(emp, index) {
  const nipClean = cleanNipBackend(emp.nip);
  if (isNipValidBackend(nipClean)) {
    return 'NIP_' + nipClean;
  } else {
    const nameClean = String(emp.nama || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    const unitClean = String(emp.unit || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!nameClean) {
      return 'RANDOM_' + (index || Math.random());
    }
    return 'NAME_' + nameClean + '_UNIT_' + unitClean;
  }
}

function deduplicateEmployeesBackend(empList) {
  if (!empList || !Array.isArray(empList)) return [];

  // Fetch reference employees from Referensi_Pegawai sheet to enforce strict NIP matching
  let refEmployees = [];
  try {
    if (typeof getReferensiPegawai === 'function') {
      refEmployees = getReferensiPegawai();
    }
  } catch (e) {
    Logger.log('Error loading reference employees in backend: ' + e.toString());
  }

  const refNipMap = {};
  if (refEmployees && Array.isArray(refEmployees) && refEmployees.length > 0) {
    refEmployees.forEach(ref => {
      const cNip = cleanNipBackend(ref.nip);
      if (isNipValidBackend(cNip)) {
        refNipMap[cNip] = ref;
      }
    });
  }

  const hasReference = Object.keys(refNipMap).length > 0;

  const keyMap = {};
  empList.forEach((emp, index) => {
    if (!emp) return;
    const nipClean = cleanNipBackend(emp.nip);

    // CRITICAL: If reference database exists in spreadsheet, strictly check that emp.nip is in Referensi_Pegawai!
    if (hasReference) {
      if (!isNipValidBackend(nipClean) || !refNipMap[nipClean]) {
        // Skip employee because NIP is invalid, missing, or does NOT exist in Referensi_Pegawai!
        return;
      }
      // Override/normalize with reference employee's official name and NIP
      const refItem = refNipMap[nipClean];
      if (refItem && refItem.nama) {
        emp.nama = String(refItem.nama).trim();
        emp.nip = String(refItem.nip).trim();
      }
    }

    const key = isNipValidBackend(nipClean) ? ('NIP_' + nipClean) : getEmployeeUniqueKeyBackend(emp, index);

    if (!keyMap[key]) {
      keyMap[key] = { ...emp };
    } else {
      const existing = keyMap[key];
      const p1 = parseFloat(existing.persentase) || 0;
      const p2 = parseFloat(emp.persentase) || 0;
      if (p2 > p1) {
        existing.persentase = p2;
        existing.jawabanBenar = emp.jawabanBenar;
        existing.mengikutiAwareness = emp.mengikutiAwareness;
      }
      const existingNipClean = cleanNipBackend(existing.nip);
      if (!isNipValidBackend(existingNipClean) && isNipValidBackend(nipClean)) {
        existing.nip = emp.nip;
      }
    }
  });

  const results = [];
  for (const k in keyMap) {
    if (Object.prototype.hasOwnProperty.call(keyMap, k)) {
      results.push(keyMap[k]);
    }
  }

  results.forEach((e, idx) => {
    e.no = idx + 1;
  });
  return results;
}

function getUncompletedEmployees() {
  try {
    const ss = getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Daftar_Belum_Survei');
    
    // Jika sheet belum ada, buat baru dengan baris header default
    if (!sheet) {
      sheet = ss.insertSheet('Daftar_Belum_Survei');
      sheet.appendRow(['No', 'Nama', 'NIP', 'Unit', 'Mengikuti_Awareness', 'Jawaban_Benar', 'Persentase']);
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const employees = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][1]) continue; // skip kosong
      employees.push({
        no: Number(data[i][0]) || i,
        nama: String(data[i][1]),
        nip: String(data[i][2] || '-'),
        unit: String(data[i][3] || 'UP3 MUARA BUNGO'),
        mengikutiAwareness: String(data[i][4] || 'Ya'),
        jawabanBenar: String(data[i][5] || 'Tidak Quiz'),
        persentase: Number(data[i][6]) || 0
      });
    }
    return deduplicateEmployeesBackend(employees);
  } catch (error) {
    Logger.log('Gagal mengambil data: ' + error.toString());
    return [];
  }
}

/**
 * 2. Memperbarui seluruh daftar pegawai belum survey di spreadsheet
 */
function updateAllEmployeesOnSheet(employeesList) {
  try {
    const ss = getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Daftar_Belum_Survei');
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    sheet = ss.insertSheet('Daftar_Belum_Survei');
    sheet.appendRow(['No', 'Nama', 'NIP', 'Unit', 'Mengikuti_Awareness', 'Jawaban_Benar', 'Persentase']);
    
    const deduplicated = deduplicateEmployeesBackend(employeesList || []);

    if (deduplicated && deduplicated.length > 0) {
      const rows = deduplicated.map((e, idx) => [
        idx + 1, 
        e.nama, 
        e.nip, 
        e.unit, 
        e.mengikutiAwareness, 
        e.jawabanBenar, 
        e.persentase
      ]);
      sheet.getRange(2, 1, rows.length, 7).setValues(rows);
    }
    return { success: true, count: deduplicated ? deduplicated.length : 0 };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 3. Mengambil konfigurasi target sasaran pegawai per unit dari Spreadsheet
 */
function getUnitTargets() {
  try {
    const ss = getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Target_Unit');
    
    // Jika sheet belum ada, buat baru dan isi dengan default benchmark terbaru (Gambar 2)
    if (!sheet) {
      sheet = ss.insertSheet('Target_Unit');
      sheet.appendRow(['ID', 'Nama_Unit', 'Kelompok_Induk', 'Target', 'Bulan', 'Realisasi (Target - Belum Survey)', 'Tanggal Update']);
      
      const defaults = [
        ['1_Mei 2026', 'UP3 MUARA BUNGO', 'UP3 MUARA BUNGO', 0, 'Mei 2026', 0, ''],
        ['2_Mei 2026', 'ULP SAROLANGUN', 'UP3 MUARA BUNGO', 0, 'Mei 2026', 0, ''],
        ['3_Mei 2026', 'ULP SNGAI PNUH', 'UP3 MUARA BUNGO', 0, 'Mei 2026', 0, ''],
        ['4_Mei 2026', 'ULP MUARA TEBO', 'UP3 MUARA BUNGO', 0, 'Mei 2026', 0, ''],
        ['5_Mei 2026', 'ULP KERSIK TUO', 'UP3 MUARA BUNGO', 0, 'Mei 2026', 0, ''],
        ['6_Mei 2026', 'ULP BANGKO', 'UP3 MUARA BUNGO', 0, 'Mei 2026', 0, ''],
        ['7_Mei 2026', 'ULP RMB BUJANG', 'UP3 MUARA BUNGO', 0, 'Mei 2026', 0, ''],
        ['8_Mei 2026', 'ULP KOTA BUNGO', 'UP3 MUARA BUNGO', 0, 'Mei 2026', 0, ''],
        ['9_Mei 2026', 'UP3 JAMBI', 'UP3 JAMBI', 0, 'Mei 2026', 0, ''],
        ['10_Mei 2026', 'ULP KOTA BARU', 'UP3 JAMBI', 0, 'Mei 2026', 0, ''],
        ['11_Mei 2026', 'ULP KUALA TNGKL', 'UP3 JAMBI', 0, 'Mei 2026', 0, ''],
        ['12_Mei 2026', 'ULP MRA BULIAN', 'UP3 JAMBI', 0, 'Mei 2026', 0, ''],
        ['13_Mei 2026', 'ULP MUARA SABAK', 'UP3 JAMBI', 0, 'Mei 2026', 0, ''],
        ['14_Mei 2026', 'ULP SBRNG KOTA', 'UP3 JAMBI', 0, 'Mei 2026', 0, ''],
        ['15_Mei 2026', 'ULP TLANAI PURA', 'UP3 JAMBI', 0, 'Mei 2026', 0, ''],
        ['16_Mei 2026', 'UPT JAMBI', 'UPT JAMBI', 0, 'Mei 2026', 0, ''],
        ['17_Mei 2026', 'ULTG AUR DURI', 'UPT JAMBI', 0, 'Mei 2026', 0, ''],
        ['18_Mei 2026', 'ULTG MUARABUNGO', 'UPT JAMBI', 0, 'Mei 2026', 0, ''], // Default Mei (Gambar 2). Maret/April adalah 23.
        ['19_Mei 2026', 'UP2K JAMBI', 'UP2K JAMBI', 0, 'Mei 2026', 0, ''],
        ['20_Mei 2026', 'UPP SUMBAGTENG3', 'UPP SUMBAGTENG3', 0, 'Mei 2026', 0, '']
      ];
      sheet.getRange(2, 1, defaults.length, 7).setValues(defaults);
      
      return defaults.map(d => ({
        id: d[0],
        name: d[1],
        parent: d[2],
        target: Number(d[3]),
        bulan: String(d[4]),
        realisasi: Number(d[5]),
        tanggalUpdate: String(d[6])
      }));
    }
    
    const data = sheet.getDataRange().getValues();
    const targets = [];
    const seenKeys = {};

    for (var i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      const rowName = String(data[i][1]);
      
      const rowBulan = getNormalizedBulanLabel(data[i].length > 4 ? data[i][4] : 'Mei 2026');

      const key = (rowName + '_' + rowBulan).toUpperCase().trim();
      
      if (seenKeys[key]) continue; // Skip duplicates on load
      seenKeys[key] = true;

      let rowId = String(data[i][0] || i);
      if (rowId.includes('_')) {
        const baseId = rowId.split('_')[0];
        rowId = baseId + '_' + rowBulan;
      } else {
        rowId = rowId + '_' + rowBulan;
      }

      targets.push({
        id: rowId,
        name: rowName,
        parent: String(data[i][2] || 'UP3 JAMBI'),
        target: Number(data[i][3]) || 0,
        bulan: rowBulan,
        realisasi: data[i].length > 5 ? Number(data[i][5] || 0) : 0,
        tanggalUpdate: data[i].length > 6 ? String(data[i][6] || '') : ''
      });
    }
    return targets;
  } catch (error) {
    Logger.log('Gagal mengambil target unit: ' + error.toString());
    // Fallback default jika terjadi masalah koneksi sheet
    return [
      { id: '1_Mei 2026', name: 'UP3 MUARA BUNGO', parent: 'UP3 MUARA BUNGO', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '2_Mei 2026', name: 'ULP SAROLANGUN', parent: 'UP3 MUARA BUNGO', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '3_Mei 2026', name: 'ULP SNGAI PNUH', parent: 'UP3 MUARA BUNGO', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '4_Mei 2026', name: 'ULP MUARA TEBO', parent: 'UP3 MUARA BUNGO', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '5_Mei 2026', name: 'ULP KERSIK TUO', parent: 'UP3 MUARA BUNGO', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '6_Mei 2026', name: 'ULP BANGKO', parent: 'UP3 MUARA BUNGO', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '7_Mei 2026', name: 'ULP RMB BUJANG', parent: 'UP3 MUARA BUNGO', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '8_Mei 2026', name: 'ULP KOTA BUNGO', parent: 'UP3 MUARA BUNGO', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '9_Mei 2026', name: 'UP3 JAMBI', parent: 'UP3 JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '10_Mei 2026', name: 'ULP KOTA BARU', parent: 'UP3 JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '11_Mei 2026', name: 'ULP KUALA TNGKL', parent: 'UP3 JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '12_Mei 2026', name: 'ULP MRA BULIAN', parent: 'UP3 JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '13_Mei 2026', name: 'ULP MUARA SABAK', parent: 'UP3 JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '14_Mei 2026', name: 'ULP SBRNG KOTA', parent: 'UP3 JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '15_Mei 2026', name: 'ULP TLANAI PURA', parent: 'UP3 JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '16_Mei 2026', name: 'UPT JAMBI', parent: 'UPT JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '17_Mei 2026', name: 'ULTG AUR DURI', parent: 'UPT JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '18_Mei 2026', name: 'ULTG MUARABUNGO', parent: 'UPT JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '19_Mei 2026', name: 'UP2K JAMBI', parent: 'UP2K JAMBI', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' },
      { id: '20_Mei 2026', name: 'UPP SUMBAGTENG3', parent: 'UPP SUMBAGTENG3', target: 0, bulan: 'Mei 2026', realisasi: 0, tanggalUpdate: '' }
    ];
  }
}

/**
 * 6. Memperbarui target sasaran di Spreadsheet
 */
function updateUnitTargets(targetsList) {
  try {
    // Ambil data target yang ada saat ini dari spreadsheet sebagai referensi (untuk mencegah pengeditan bulan lalu)
    const existingTargets = getUnitTargets();
    const currentMonthLabel = getCurrentMonthLabelBackend();

    const ss = getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Target_Unit');
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet('Target_Unit');
    sheet.appendRow(['ID', 'Nama_Unit', 'Kelompok_Induk', 'Target', 'Bulan', 'Realisasi (Target - Belum Survey)', 'Tanggal Update']);
    
    if (targetsList && targetsList.length > 0) {
      const rows = [];
      const seenKeys = {};

      targetsList.forEach(t => {
        const uBulan = getNormalizedBulanLabel(t.bulan || 'Mei 2026');
        const uName = t.name;
        const key = (uName + '_' + uBulan).toUpperCase().trim();

        if (seenKeys[key]) return; // Skip duplicates on write
        seenKeys[key] = true;

        let targetVal = Number(t.target) || 0;
        let realisasiVal = t.realisasi || 0;
        let tanggalUpdateVal = t.tanggalUpdate || '';

        // Ensure compound ID is clean with the appropriate month suffix
        let compoundId = String(t.id);
        if (compoundId.includes('_')) {
          const baseId = compoundId.split('_')[0];
          compoundId = baseId + '_' + uBulan;
        } else {
          compoundId = compoundId + '_' + uBulan;
        }

        // Cari data yang sudah ada sebelumnya di spreadsheet untuk unit dan bulan ini
        const existing = existingTargets.find(ext => ext.id === compoundId && ext.bulan === uBulan);

        // BATASAN HISTORI DIHAPUS (Bisa memperbarui data kapan saja untuk backfile lancar)

        rows.push([
          compoundId, 
          uName, 
          t.parent, 
          targetVal,
          uBulan,
          realisasiVal,
          tanggalUpdateVal
        ]);
      });

      if (rows.length > 0) {
        sheet.getRange(2, 5, rows.length, 1).setNumberFormat("@"); // Set format BEFORE setValues
        sheet.getRange(2, 1, rows.length, 7).setValues(rows);
      }
    }
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Normalizes any Sheet date cell value (whether string or Date object) 
 * into a clean Indonesian month name like "Juni 2026".
 */
function getNormalizedBulanLabel(val) {
  if (!val) return "Mei 2026";
  
  const indonesianMonths = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  
  // 1. If it's a native Date, format it safely with a 12-hour buffer to avoid midnight timezone shift!
  if (val instanceof Date || Object.prototype.toString.call(val) === '[object Date]') {
    const adjustedDate = new Date(val.getTime() + 12 * 60 * 60 * 1000);
    return indonesianMonths[adjustedDate.getMonth()] + " " + adjustedDate.getFullYear();
  }
  
  const valStr = String(val).trim();
  if (!valStr) return "Mei 2026";
  
  const lower = valStr.toLowerCase();
  
  // 2. Year Extraction (default to 2026 if not found)
  const yearMatch = valStr.match(/\b(20\d{2})\b/);
  const year = yearMatch ? parseInt(yearMatch[1]) : 2026;
  
  // 3. Month string keyword matcher (performs text-based mapping to bypass timezone-unsafe new Date() parsing)
  const monthMap = {
    'januari': 'Januari', 'februari': 'Februari', 'maret': 'Maret', 'april': 'April', 'mei': 'Mei', 'juni': 'Juni',
    'juli': 'Juli', 'agustus': 'Agustus', 'september': 'September', 'oktober': 'Oktober', 'november': 'November', 'desember': 'Desember',
    'january': 'Januari', 'february': 'Februari', 'march': 'Maret', 'june': 'Juni',
    'july': 'Juli', 'august': 'Agustus', 'october': 'Oktober', 'december': 'Desember',
    'jan': 'Januari', 'feb': 'Februari', 'mar': 'Maret', 'apr': 'April', 'may': 'Mei', 'jun': 'Juni',
    'jul': 'Juli', 'aug': 'Agustus', 'sep': 'September', 'oct': 'Oktober', 'nov': 'November', 'dec': 'Desember'
  };
  
  for (const [key, indMonth] of Object.entries(monthMap)) {
    if (lower.indexOf(key) !== -1) {
      return indMonth + " " + year;
    }
  }
  
  // 4. Numerical date parsing fallback (e.g., 2026-05-01, 05/2026)
  if (valStr.includes('/') || valStr.includes('-')) {
    const numbers = valStr.match(/\b\d+\b/g);
    if (numbers && numbers.length >= 2) {
      let monthIndex = -1;
      if (numbers[0].length === 4) {
        // Format: YYYY-MM-DD
        monthIndex = parseInt(numbers[1]) - 1;
      } else if (numbers[1].length === 4) {
        // Format: MM/YYYY or DD/MM/YYYY
        monthIndex = parseInt(numbers[0]) - 1;
      } else {
        // Format: DD-MM-YYYY
        monthIndex = parseInt(numbers[1]) - 1;
      }
      if (monthIndex >= 0 && monthIndex < 12) {
        return indonesianMonths[monthIndex] + " " + year;
      }
    }
  }
  
  return valStr;
}

function getCurrentMonthLabelBackend() {
  const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const now = new Date();
  return months[now.getMonth()] + " " + now.getFullYear();
}

function parseMonthLabelBackend(label) {
  if (!label) return null;
  const parts = label.trim().split(/\s+/);
  if (parts.length < 2) return null;
  const indonesianMonths = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const mIndex = indonesianMonths.findIndex(m => m.toLowerCase() === parts[0].toLowerCase());
  const year = parseInt(parts[1]);
  if (mIndex === -1 || isNaN(year)) return null;
  return { month: mIndex, year: year };
}

function isPastMonthBackend(label, currentLabel) {
  const current = parseMonthLabelBackend(currentLabel);
  const target = parseMonthLabelBackend(label);
  if (!current || !target) return false;
  if (target.year < current.year) return true;
  if (target.year === current.year && target.month < current.month) return true;
  return false;
}

/**
 * Helper function to include HTML files dynamically within templates.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

/**
 * Verifikasi password secara aman dari sisi server.
 * Default password: st1Jambi
 */
function verifyPasswordOnServer(passwordInput) {
  const SVR_PASSWORD = "st1Jambi";
  return passwordInput === SVR_PASSWORD;
}

/**
 * 7. Mengambil data referensi pegawai dari sheet 'Referensi_Pegawai'
 */
function getReferensiPegawai() {
  try {
    const ss = getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Referensi_Pegawai');
    if (!sheet) {
      sheet = ss.insertSheet('Referensi_Pegawai');
      sheet.appendRow(['No', 'Nama', 'NIP', 'Email_Korporat', 'Unit_Pelayanan', 'Sub_Unit_Pelayanan']);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const list = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      list.push({
        no: Number(data[i][0]) || i,
        nama: String(data[i][1]),
        nip: String(data[i][2]),
        email: String(data[i][3]),
        unit: String(data[i][4]),
        subUnit: String(data[i][5])
      });
    }
    return list;
  } catch (error) {
    Logger.log('Gagal mengambil data referensi pegawai: ' + error.toString());
    return [];
  }
}

/**
 * 8. Menyimpan data referensi pegawai ke sheet 'Referensi_Pegawai'
 */
function saveReferensiPegawai(refList) {
  try {
    const ss = getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Referensi_Pegawai');
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet('Referensi_Pegawai');
    sheet.appendRow(['No', 'Nama', 'NIP', 'Email_Korporat', 'Unit_Pelayanan', 'Sub_Unit_Pelayanan']);
    
    if (refList && refList.length > 0) {
      const rows = refList.map((e, idx) => [
        idx + 1,
        e.nama,
        e.nip,
        e.email,
        e.unit,
        e.subUnit
      ]);
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }
    return { success: true, count: refList ? refList.length : 0 };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 9. Mengambil seluruh data master CSA (Referensi Pegawai, Daftar Belum Survei, Target Unit) sekaligus untuk optimasi pemuatan client.
 */
function getCSAMasterData() {
  try {
    const employees = getReferensiPegawai();
    const uncompleted = getUncompletedEmployees();
    const targets = getUnitTargets();
    
    return {
      employees: employees,
      uncompleted: uncompleted,
      targets: targets
    };
  } catch (error) {
    return { error: error.toString() };
  }
}
