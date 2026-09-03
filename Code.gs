/**
 * SuaraWarga - Sistem Voting Berbasis Google Apps Script
 * Backend (Code.gs)
 * 
 * Cara Pemasangan:
 * 1. Buat Google Spreadsheet baru.
 * 2. Buka menu Extensions > Apps Script.
 * 3. Hapus semua kode default dan paste kode ini.
 * 4. Simpan proyek.
 * 5. Jalankan fungsi `initSheets` secara manual di editor. Setujui izin yang diminta.
 * 6. Lakukan deploy sebagai Web App (Menu Deploy > New deployment):
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Gunakan URL Web App yang dihasilkan untuk frontend Anda.
 */

// ==========================================
// SETUP & INISIALISASI
// ==========================================

/**
 * Membuat sheet yang dibutuhkan jika belum ada.
 * Dijalankan sekali saat setup.
 */
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Setup Sheet Calon
  var sheetCalon = ss.getSheetByName('Calon');
  if (!sheetCalon) {
    sheetCalon = ss.insertSheet('Calon');
    sheetCalon.appendRow(['noUrut', 'nama', 'visi', 'misi', 'foto', 'suara']);
  }
  
  // Setup Sheet Voters
  var sheetVoters = ss.getSheetByName('Voters');
  if (!sheetVoters) {
    sheetVoters = ss.insertSheet('Voters');
    sheetVoters.appendRow(['nama', 'kode', 'sudahVoting']);
  }
  
  // Setup Sheet Settings
  var sheetSettings = ss.getSheetByName('Settings');
  if (!sheetSettings) {
    sheetSettings = ss.insertSheet('Settings');
    sheetSettings.appendRow(['key', 'value']);
    sheetSettings.appendRow(['statusSesi', 'tutup']);
  }
}

// ==========================================
// WEB HANDLERS (GET & POST)
// ==========================================

function doGet(e) {
  try {
    var action = e.parameter.action;
    var result = {};
    
    switch(action) {
      case 'getCalon':
        result = handleGetCalon();
        break;
      case 'getStatusSesi':
        result = handleGetStatusSesi();
        break;
      case 'getVoters':
        result = handleGetVoters();
        break;
      default:
        result = {error: 'Action GET tidak valid.'};
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({error: error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.action) {
      // Fallback jika tidak menggunakan body JSON
      data = e.parameter;
    }
    
    var action = data.action;
    var result = {};
    
    switch(action) {
      case 'tambahCalon':
        result = handleTambahCalon(data);
        break;
      case 'editCalon':
        result = handleEditCalon(data);
        break;
      case 'hapusCalon':
        result = handleHapusCalon(data);
        break;
      case 'tambahVoter':
        result = handleTambahVoter(data);
        break;
      case 'editVoter':
        result = handleEditVoter(data);
        break;
      case 'hapusVoter':
        result = handleHapusVoter(data);
        break;
      case 'toggleSesi':
        result = handleToggleSesi();
        break;
      case 'verifikasiVoter':
        result = handleVerifikasiVoter(data);
        break;
      case 'vote':
        result = handleVote(data);
        break;
      default:
        result = {error: 'Action POST tidak valid.'};
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({error: error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function normalizeHeader(str) {
  var s = String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (s === 'namalengkap' || s === 'nama' || s === 'namawarga') return 'nama';
  if (s === 'kodeunik' || s === 'kode') return 'kode';
  if (s === 'sudahvoting' || s === 'statusvoting') return 'sudahVoting';
  if (s === 'nourut' || s === 'no') return 'noUrut';
  if (s === 'namacalon') return 'nama';
  if (s === 'visi') return 'visi';
  if (s === 'misi') return 'misi';
  if (s === 'foto' || s === 'fotourl' || s === 'gambar') return 'foto';
  if (s === 'suara' || s === 'jumlahsuara') return 'suara';
  return str;
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var objects = [];
  var sheetName = sheet.getName();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var rawHeader = headers[j];
      var normHeader = normalizeHeader(rawHeader);
      obj[rawHeader] = row[j];
      obj[normHeader] = row[j];
    }
    
    // Fallback otomatis berdasarkan posisi kolom jika header tidak standar
    if (sheetName === 'Voters') {
      if (obj.nama === undefined || obj.nama === '') obj.nama = row[0];
      if (obj.kode === undefined || obj.kode === '') obj.kode = row[1];
      if (obj.sudahVoting === undefined) obj.sudahVoting = row[2];
    } else if (sheetName === 'Calon') {
      if (obj.noUrut === undefined) obj.noUrut = row[0];
      if (obj.nama === undefined) obj.nama = row[1];
      if (obj.visi === undefined) obj.visi = row[2];
      if (obj.misi === undefined) obj.misi = row[3];
      if (obj.foto === undefined) obj.foto = row[4];
      if (obj.suara === undefined) obj.suara = row[5];
    }
    
    obj._rowIndex = i + 1; // Menyimpan baris asli untuk operasi update/delete
    objects.push(obj);
  }
  return objects;
}

function generateKode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // Tanpa I, O, 0, 1
  var result = '';
  for (var i = 0; i < 5; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function generateUniqueKode() {
  var sheetVoters = getSheet('Voters');
  if (!sheetVoters) return generateKode(); // Fallback jika sheet belum ada
  
  var voters = sheetToObjects(sheetVoters);
  var existingKodes = voters.map(function(v) { return v.kode; });
  
  var kode;
  var isUnique = false;
  var maxAttempts = 100;
  var attempts = 0;
  
  while (!isUnique && attempts < maxAttempts) {
    kode = generateKode();
    if (existingKodes.indexOf(kode) === -1) {
      isUnique = true;
    }
    attempts++;
  }
  
  return kode;
}

// ==========================================
// GET HANDLERS
// ==========================================

function handleGetCalon() {
  var sheet = getSheet('Calon');
  if (!sheet) throw new Error('Sheet Calon tidak ditemukan.');
  
  var calonList = sheetToObjects(sheet);
  
  var totalSuara = 0;
  for (var i = 0; i < calonList.length; i++) {
    var suara = parseInt(calonList[i].suara) || 0;
    calonList[i].suara = suara;
    totalSuara += suara;
  }
  
  for (var i = 0; i < calonList.length; i++) {
    if (totalSuara === 0) {
      calonList[i].persen = 0;
    } else {
      calonList[i].persen = Math.round((calonList[i].suara / totalSuara) * 100);
    }
    // Hapus internal index sebelum dikembalikan
    delete calonList[i]._rowIndex;
  }
  
  return {calon: calonList};
}

function handleGetStatusSesi() {
  var sheet = getSheet('Settings');
  if (!sheet) throw new Error('Sheet Settings tidak ditemukan.');
  
  var data = sheet.getDataRange().getValues();
  var status = 'tutup';
  var statusFound = false;
  
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0] || '').trim().toLowerCase();
    if (key === 'statussesi') {
      status = String(data[i][1] || '').trim().toLowerCase();
      statusFound = true;
      break;
    }
  }
  
  if (!statusFound) {
    sheet.appendRow(['statusSesi', 'tutup']);
  }
  
  return {status: status};
}

function handleGetVoters() {
  var sheet = getSheet('Voters');
  if (!sheet) throw new Error('Sheet Voters tidak ditemukan.');
  
  var voters = sheetToObjects(sheet);
  var result = [];
  
  for (var i = 0; i < voters.length; i++) {
    result.push({
      nama: voters[i].nama,
      kode: voters[i].kode,
      sudahVoting: (voters[i].sudahVoting === true || String(voters[i].sudahVoting).toLowerCase() === 'true')
    });
  }
  
  return {voters: result};
}

/// ==========================================
// PHOTO URL HELPER
// ==========================================

function formatPhotoUrl(url) {
  if (!url) return '';
  url = String(url).trim();
  
  // Format link Google Drive ke direct viewable URL
  var driveMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || 
                   url.match(/[?&]id=([a-zA-Z0-9_-]+)/) ||
                   url.match(/\/d\/([a-zA-Z0-9_-]+)/);
                   
  if (driveMatch && driveMatch[1]) {
    return 'https://lh3.googleusercontent.com/d/' + driveMatch[1];
  }
  return url;
}

// ==========================================
// CALON CRUD
// ==========================================

function handleTambahCalon(data) {
  var sheet = getSheet('Calon');
  if (!sheet) throw new Error('Sheet Calon tidak ditemukan.');
  
  var calonList = sheetToObjects(sheet);
  var maxNoUrut = 0;
  for (var i = 0; i < calonList.length; i++) {
    var noUrut = parseInt(calonList[i].noUrut) || 0;
    if (noUrut > maxNoUrut) maxNoUrut = noUrut;
  }
  
  var newNoUrut = maxNoUrut + 1;
  var fotoUrl = formatPhotoUrl(data.foto || '');
  
  sheet.appendRow([
    newNoUrut, 
    data.nama || '', 
    data.visi || '', 
    data.misi || '', 
    fotoUrl, 
    0 // suara awal
  ]);
  
  return {success: true, noUrut: newNoUrut, foto: fotoUrl};
}

function handleEditCalon(data) {
  var sheet = getSheet('Calon');
  var targetNoUrut = parseInt(data.noUrut);
  
  if (!sheet || isNaN(targetNoUrut)) throw new Error('Parameter tidak valid.');
  
  var calonList = sheetToObjects(sheet);
  var rowIndex = -1;
  
  // 1. Cari berdasarkan noUrut
  for (var i = 0; i < calonList.length; i++) {
    if (parseInt(calonList[i].noUrut) === targetNoUrut) {
      rowIndex = calonList[i]._rowIndex;
      break;
    }
  }
  
  // 2. Fallback: cari berdasarkan urutan baris
  if (rowIndex === -1 && targetNoUrut > 0 && targetNoUrut <= calonList.length) {
    rowIndex = calonList[targetNoUrut - 1]._rowIndex;
  }
  
  // 3. Fallback: cari berdasarkan nama calon
  if (rowIndex === -1 && data.nama) {
    var targetNama = String(data.nama).trim().toLowerCase();
    for (var j = 0; j < calonList.length; j++) {
      if (String(calonList[j].nama).trim().toLowerCase() === targetNama) {
        rowIndex = calonList[j]._rowIndex;
        break;
      }
    }
  }
  
  if (rowIndex === -1) throw new Error('Calon tidak ditemukan.');
  
  var headers = sheet.getDataRange().getValues()[0];
  var noUrutCol = -1, namaCol = -1, visiCol = -1, misiCol = -1, fotoCol = -1;
  for (var h = 0; h < headers.length; h++) {
    var norm = normalizeHeader(headers[h]);
    if (norm === 'nourut') noUrutCol = h + 1;
    if (norm === 'nama') namaCol = h + 1;
    if (norm === 'visi') visiCol = h + 1;
    if (norm === 'misi') misiCol = h + 1;
    if (norm === 'foto') fotoCol = h + 1;
  }
  if (noUrutCol === -1) noUrutCol = 1;
  if (namaCol === -1) namaCol = 2;
  if (visiCol === -1) visiCol = 3;
  if (misiCol === -1) misiCol = 4;
  if (fotoCol === -1) fotoCol = 5;
  
  // Pastikan noUrut terisi di kolom A
  sheet.getRange(rowIndex, noUrutCol).setValue(targetNoUrut);
  
  if (data.nama !== undefined) sheet.getRange(rowIndex, namaCol).setValue(data.nama);
  if (data.visi !== undefined) sheet.getRange(rowIndex, visiCol).setValue(data.visi);
  if (data.misi !== undefined) sheet.getRange(rowIndex, misiCol).setValue(data.misi);
  
  var fotoUrl = data.foto !== undefined ? formatPhotoUrl(data.foto) : undefined;
  if (fotoUrl !== undefined) sheet.getRange(rowIndex, fotoCol).setValue(fotoUrl);
  
  return {success: true, foto: fotoUrl};
}

function handleHapusCalon(data) {
  var sheet = getSheet('Calon');
  var targetNoUrut = parseInt(data.noUrut);
  
  if (!sheet || isNaN(targetNoUrut)) throw new Error('Parameter tidak valid.');
  
  var calonList = sheetToObjects(sheet);
  var rowIndex = -1;
  
  for (var i = 0; i < calonList.length; i++) {
    if (parseInt(calonList[i].noUrut) === targetNoUrut) {
      rowIndex = calonList[i]._rowIndex;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error('Calon tidak ditemukan.');
  
  sheet.deleteRow(rowIndex);
  return {success: true};
}

// ==========================================
// VOTER CRUD
// ==========================================

function handleTambahVoter(data) {
  var sheet = getSheet('Voters');
  if (!sheet) throw new Error('Sheet Voters tidak ditemukan.');
  
  var nama = (data.nama || '').trim();
  if (!nama) throw new Error('Nama voter harus diisi.');
  
  var kode = generateUniqueKode();
  
  sheet.appendRow([nama, kode, false]);
  return {success: true, kode: kode};
}

function handleEditVoter(data) {
  var sheet = getSheet('Voters');
  var targetKode = (data.kode || '').trim().toUpperCase();
  
  if (!sheet || !targetKode) throw new Error('Parameter tidak valid.');
  
  var voters = sheetToObjects(sheet);
  var rowIndex = -1;
  
  for (var i = 0; i < voters.length; i++) {
    if (String(voters[i].kode).trim().toUpperCase() === targetKode) {
      rowIndex = voters[i]._rowIndex;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error('Voter tidak ditemukan.');
  
  var headers = sheet.getDataRange().getValues()[0];
  var namaCol = -1;
  for (var h = 0; h < headers.length; h++) {
    if (normalizeHeader(headers[h]) === 'nama') {
      namaCol = h + 1;
      break;
    }
  }
  if (namaCol === -1) namaCol = 1;
  if (data.nama !== undefined) {
    sheet.getRange(rowIndex, namaCol).setValue(data.nama.trim());
  }
  
  return {success: true};
}

function handleHapusVoter(data) {
  var sheet = getSheet('Voters');
  var targetKode = (data.kode || '').trim().toUpperCase();
  
  if (!sheet || !targetKode) throw new Error('Parameter tidak valid.');
  
  var voters = sheetToObjects(sheet);
  var rowIndex = -1;
  
  for (var i = 0; i < voters.length; i++) {
    if (String(voters[i].kode).trim().toUpperCase() === targetKode) {
      rowIndex = voters[i]._rowIndex;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error('Voter tidak ditemukan.');
  
  sheet.deleteRow(rowIndex);
  return {success: true};
}

// ==========================================
// SESSION CONTROL
// ==========================================

function handleToggleSesi() {
  var sheet = getSheet('Settings');
  if (!sheet) throw new Error('Sheet Settings tidak ditemukan.');
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var currentStatus = 'tutup';
  
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0] || '').trim().toLowerCase();
    if (key === 'statussesi') {
      rowIndex = i + 1;
      currentStatus = String(data[i][1] || '').trim().toLowerCase();
      break;
    }
  }
  
  var newStatus = (currentStatus === 'buka') ? 'tutup' : 'buka';
  
  if (rowIndex === -1) {
    sheet.appendRow(['statusSesi', newStatus]);
  } else {
    sheet.getRange(rowIndex, 2).setValue(newStatus);
  }
  
  return {success: true, status: newStatus};
}

// ==========================================
// VOTING
// ==========================================

function checkStatusSesi() {
  var statusInfo = handleGetStatusSesi();
  return statusInfo.status === 'buka';
}

function handleVerifikasiVoter(data) {
  if (!checkStatusSesi()) {
    return {valid: false, pesan: 'Sesi pemilihan sedang ditutup.'};
  }
  
  var sheet = getSheet('Voters');
  var namaInput = (data.nama || '').trim().toLowerCase();
  var kodeInput = (data.kode || '').trim().toUpperCase();
  
  if (!namaInput || !kodeInput) {
    return {valid: false, pesan: 'Nama dan kode harus diisi.'};
  }
  
  var voters = sheetToObjects(sheet);
  var voterFound = null;
  
  for (var i = 0; i < voters.length; i++) {
    var vNama = String(voters[i].nama).trim().toLowerCase();
    var vKode = String(voters[i].kode).trim().toUpperCase();
    
    if (vNama === namaInput && vKode === kodeInput) {
      voterFound = voters[i];
      break;
    }
  }
  
  if (!voterFound) {
    return {valid: false, pesan: 'Nama atau kode tidak cocok.'};
  }
  
  var sudahVoting = (voterFound.sudahVoting === true || String(voterFound.sudahVoting).toLowerCase() === 'true');
  if (sudahVoting) {
    return {valid: false, pesan: 'Anda sudah melakukan pemilihan.'};
  }
  
  var calonInfo = handleGetCalon();
  return {
    valid: true, 
    calon: calonInfo.calon
  };
}

function handleVote(data) {
  // 1. Cek Sesi
  if (!checkStatusSesi()) {
    throw new Error('Sesi pemilihan sedang ditutup.');
  }
  
  // 2. Cek Voter
  var namaInput = (data.nama || '').trim().toLowerCase();
  var kodeInput = (data.kode || '').trim().toUpperCase();
  var calonId = parseInt(data.calonId);
  
  if (isNaN(calonId)) throw new Error('Pilihan tidak valid.');
  
  var sheetVoters = getSheet('Voters');
  var voters = sheetToObjects(sheetVoters);
  var voterIndex = -1;
  var voterFound = null;
  
  for (var i = 0; i < voters.length; i++) {
    var vNama = String(voters[i].nama).trim().toLowerCase();
    var vKode = String(voters[i].kode).trim().toUpperCase();
    
    if (vNama === namaInput && vKode === kodeInput) {
      voterFound = voters[i];
      voterIndex = voters[i]._rowIndex;
      break;
    }
  }
  
  if (!voterFound) throw new Error('Pemilih tidak valid.');
  
  var sudahVoting = (voterFound.sudahVoting === true || String(voterFound.sudahVoting).toLowerCase() === 'true');
  if (sudahVoting) throw new Error('Anda sudah pernah memilih.');
  
  // 3. Tambah Suara Calon
  var sheetCalon = getSheet('Calon');
  var calonList = sheetToObjects(sheetCalon);
  var calonIndex = -1;
  var currentSuara = 0;
  
  for (var i = 0; i < calonList.length; i++) {
    if (parseInt(calonList[i].noUrut) === calonId) {
      calonIndex = calonList[i]._rowIndex;
      currentSuara = parseInt(calonList[i].suara) || 0;
      break;
    }
  }
  
  if (calonIndex === -1) throw new Error('Kandidat tidak ditemukan.');
  
  // 4. Update Data (Execute)
  var headersCalon = sheetCalon.getDataRange().getValues()[0];
  var suaraCol = -1;
  for (var hc = 0; hc < headersCalon.length; hc++) {
    if (normalizeHeader(headersCalon[hc]) === 'suara') {
      suaraCol = hc + 1;
      break;
    }
  }
  if (suaraCol === -1) suaraCol = 6;
  sheetCalon.getRange(calonIndex, suaraCol).setValue(currentSuara + 1);
  
  var headersVoter = sheetVoters.getDataRange().getValues()[0];
  var sudahVotingCol = -1;
  for (var hv = 0; hv < headersVoter.length; hv++) {
    if (normalizeHeader(headersVoter[hv]) === 'sudahVoting') {
      sudahVotingCol = hv + 1;
      break;
    }
  }
  if (sudahVotingCol === -1) sudahVotingCol = 3;
  sheetVoters.getRange(voterIndex, sudahVotingCol).setValue(true);
  
  return {success: true};
}
