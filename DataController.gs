/**
 * DATA CONTROLLER - PUSAT KAWALAN DATA
 * Menguruskan pengambilan data dari Sheet dan pemprosesan logik backend.
 */

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. HELPER FUNCTIONS ---
  const getMap = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {}; 
    const data = sheet.getDataRange().getDisplayValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0] ? data[i][0].toString().trim() : "";
      const name = data[i][1] ? data[i][1].toString().trim() : "";
      if (id) map[id] = name;
    }
    return map;
  };

  const getMapVot = () => {
    const sheet = ss.getSheetByName('Vot') || ss.getSheetByName('vot');
    if (!sheet) return {};
    const data = sheet.getDataRange().getDisplayValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0] ? data[i][0].toString().trim() : "";
      const kod = data[i][1] || "";
      if (id) map[id] = kod;
    }
    return map;
  };

  const getMapBayaran = () => {
    const sheet = ss.getSheetByName('Perakuan Siap') || ss.getSheetByName('Perakuan_Siap') || ss.getSheetByName('Perakuan Siap Kerja');
    if (!sheet) return {};
    const data = sheet.getDataRange().getDisplayValues();
    const map = {};
    // Indeks Column: [0]UNIQUE_ID, [1]PROJEK_ID, [5]JUMLAH, [6]SST (Pastikan indeks ini betul dalam sheet anda)
    for (let i = 1; i < data.length; i++) {
      const projId = data[i][1] ? data[i][1].toString().trim() : ""; 
      let amount = data[i][5] ? parseFloat(data[i][5].toString().replace(/,/g, '')) : 0;
      let sst = data[i][6] ? parseFloat(data[i][6].toString().replace(/,/g, '')) : 0;
      if (projId) {
        if (!map[projId]) map[projId] = 0;
        map[projId] += (amount + sst); 
      }
    }
    return map;
  };

  // --- 2. LOAD MAPS (RUJUKAN) ---
  const mapBahagian = getMap('Bahagian') || getMap('bahagian');
  const mapSyarikat = getMap('Syarikat') || getMap('syarikat');
  const mapSeksyen  = getMap('Seksyen_Unit') || getMap('seksyen_unit');
  const mapPegawai  = getMap('REF_PEGAWAI') || {}; // Jika tiada sheet REF, ia kosong
  const mapVot      = getMapVot();
  const mapBayaran  = getMapBayaran();

  // --- 3. BACA DATA PROJEK ---
  const projSheet = ss.getSheetByName('Projek');
  if (!projSheet) return JSON.stringify([]);
  
  const data = projSheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rawRows = data.slice(1);

  // --- 4. UTILITY HELPERS ---
  const parseRM = (v) => v ? parseFloat(v.toString().replace(/[^0-9.-]+/g,"")) : 0;
  
  const parseDateObj = (d) => {
    if (!d) return null;
    const p = d.split('/'); // Format DD/MM/YYYY
    if (p.length < 3) return null;
    return new Date(p[2], p[1]-1, p[0]);
  };

  const today = new Date();

  const calcDateDiff = (start, end, type) => {
    if (!start || !end) return null;
    const diffTime = end.getTime() - start.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    if (diffDays < 0) return "0 Hari";
    if (type === 'MINGGU') {
      const weeks = Math.floor(diffDays / 7);
      const days = diffDays % 7;
      let s = `${weeks} Minggu`;
      if (days > 0) s += ` ${days} Hari`;
      return s;
    } else if (type === 'BULAN') {
      const nextDay = new Date(end);
      nextDay.setDate(nextDay.getDate() + 1);
      let months = (nextDay.getFullYear() - start.getFullYear()) * 12 + (nextDay.getMonth() - start.getMonth());
      if (nextDay.getDate() < start.getDate()) { months--; }
      if (months < 1 && diffDays > 0) return `${diffDays} Hari`;
      return `${months} Bulan`;
    }
    return "";
  };

  // --- 5. PEMPROSESAN UTAMA (MAPPING) ---
  const formatted = rawRows.map(row => {
    let obj = {};
    headers.forEach((h, i) => {
       const cleanHeader = h ? h.toString().trim() : "";
       if (cleanHeader) obj[cleanHeader] = row[i];
    });

    const uniqueId = obj['UNIQUE_ID'] ? obj['UNIQUE_ID'].toString().trim() : ""; 

    // Mapping Nama (ID -> Nama Sebenar)
    const idSyarikat = (obj['SYARIKAT'] || "").trim();
    const idBahagian = (obj['BAHAGIAN'] || "").trim();
    const idSeksyen  = (obj['SEKSYEN'] || "").trim();
    const idPegawai  = (obj['PEGAWAI BERTANGGUNGJAWAB'] || "").trim(); // ID Pegawai
    const idVot      = (obj['VOT'] || "").trim();

    // Nota: Jika ID Pegawai tiada dalam Map, guna nilai asal (Nama)
    obj['SYARIKAT_NAMA'] = mapSyarikat[idSyarikat] || idSyarikat || '-';
    obj['BAHAGIAN_NAMA'] = mapBahagian[idBahagian] || idBahagian || '-';
    obj['SEKSYEN_NAMA']  = mapSeksyen[idSeksyen]   || idSeksyen  || '-';
    obj['PEGAWAI_NAMA']  = mapPegawai[idPegawai]   || idPegawai  || '-'; 
    obj['VOT_NAMA']      = mapVot[idVot] || idVot || '-'; 
    obj['NAMA_PAPARAN']  = (obj['PROJEK'] && obj['PROJEK'].length > 3) ? obj['PROJEK'] : obj['SISTEM'];

    // Kewangan & Kos Akhir
    const valJumlah = parseRM(obj['JUMLAH']);
    const valVOAdd = parseRM(obj['VO TAMBAHAN']);
    const valVOLess = parseRM(obj['VO PENGURANGAN']);
    const valSST = parseRM(obj['SST']);
    const valSSTAdj = parseRM(obj['PELARASAN SST']);
    
    // Formula: (Asal + VO Tambah - VO Kurang) + (SST + Pelarasan SST)
    obj['KOS_AKHIR'] = (valJumlah + valVOAdd - valVOLess) + (valSST + valSSTAdj);
    obj['KOS_PECAHAN'] = { kontrak: (valJumlah + valVOAdd - valVOLess), sst: (valSST + valSSTAdj) };
    
    // Bayaran
    obj['JUMLAH_BAYARAN'] = mapBayaran[uniqueId] || 0;
    let progressPayment = 0;
    if (obj['KOS_AKHIR'] > 0) progressPayment = Math.round((obj['JUMLAH_BAYARAN'] / obj['KOS_AKHIR']) * 100);
    if (progressPayment > 100) progressPayment = 100;
    obj['PROGRESS_MASA'] = progressPayment;

    // Logik Tarikh & EOT
    const dateMula = parseDateObj(obj['TARIKH MULA']);
    const dateSiapAsal = parseDateObj(obj['TARIKH SIAP']);
    let dateSiapSemasa = dateSiapAsal ? new Date(dateSiapAsal) : null;
    
    const rawEOT = obj['EOT'] ? obj['EOT'].toString().trim() : "";
    
    if (dateSiapAsal && rawEOT) {
      const eotDate = parseDateObj(rawEOT);
      if (eotDate && eotDate > dateSiapAsal) {
        dateSiapSemasa = eotDate; // EOT jenis Tarikh
      } else if (!isNaN(rawEOT)) {
        const days = parseInt(rawEOT);
        if (days > 0) dateSiapSemasa.setDate(dateSiapSemasa.getDate() + days); // EOT jenis Hari
      }
    }

    obj['TARIKH_MULA_ISO'] = dateMula;
    obj['TARIKH_SIAP_ASAL_ISO'] = dateSiapAsal;
    obj['TARIKH_SIAP_SEMASA_ISO'] = dateSiapSemasa;
    obj['TAHUN_MULA'] = dateMula ? dateMula.getFullYear() : null;

    // Tempoh Paparan
    const cat = (obj['KATEGORI PERKHIDMATAN'] || "").toUpperCase();
    const isMaintenance = cat.includes("PENYELENGGARAAN");
    const mode = isMaintenance ? 'BULAN' : 'MINGGU';

    let displayTempoh = [];
    if (dateMula && dateSiapAsal) {
      const durAsal = calcDateDiff(dateMula, dateSiapAsal, mode);
      displayTempoh.push({ label: "Tempoh Asal", val: durAsal });

      if (dateSiapSemasa && dateSiapSemasa.getTime() > dateSiapAsal.getTime()) {
        const durEOT = calcDateDiff(dateSiapAsal, dateSiapSemasa, mode);
        displayTempoh.push({ label: "Tempoh EOT", val: durEOT });
        const durTotal = calcDateDiff(dateMula, dateSiapSemasa, mode);
        displayTempoh.push({ label: "Keseluruhan", val: durTotal });
      }
    }
    obj['TEMPOH_DATA'] = displayTempoh;

    // Status Projek (Logic Status)
    let status = 'Belum Mula';
    const tSelesai = parseDateObj(obj['TARIKH SELESAI']); 
    let infoLewat = "";

    if (obj['STATUS'] === 'Completed' || tSelesai) {
      status = 'Siap (CPC)';
      if (progressPayment === 0) progressPayment = 100; // Auto 100% jika status dah siap
      if (tSelesai && dateSiapSemasa && tSelesai > dateSiapSemasa) {
        status = 'Siap (Lewat)';
        const hariLewat = Math.ceil((tSelesai - dateSiapSemasa) / (1000 * 60 * 60 * 24));
        const mingguLewat = Math.floor(hariLewat / 7);
        const bakiHari = hariLewat % 7;
        infoLewat = `Lewat: ${mingguLewat} Minggu ${bakiHari} Hari`;
      }
    } else if (obj['STATUS'] === 'Cancel') {
      status = 'Batal';
    } else if (dateMula && dateSiapSemasa) {
      if (today > dateSiapSemasa) {
        status = 'Lewat (Overdue)';
        const hariLewat = Math.ceil((today - dateSiapSemasa) / (1000 * 60 * 60 * 24));
        const mingguLewat = Math.floor(hariLewat / 7);
        const bakiHari = hariLewat % 7;
        infoLewat = `Overdue: ${mingguLewat} Minggu ${bakiHari} Hari`;
      } else {
        status = 'Dalam Jadual';
      }
    }

    obj['LOGIK_SIAP'] = status;
    obj['INFO_LEWAT'] = infoLewat;

    return obj;
  });

  return JSON.stringify(formatted);
}

/**
 * --- FUNGSI UTAMA BARU (GAMIFICATION) ---
 * Menggabungkan Data Projek + Kemajuan + Perakuan Siap
 * Ini adalah fungsi yang perlu dipanggil oleh Frontend (App.js)
 */
function getIntegratedData() {
  // 1. Dapatkan Data Projek yang sudah diproses (guna fungsi getData di atas)
  var rawProjek = getData(); 
  var dataProjek = JSON.parse(rawProjek);

  // 2. Dapatkan Data Kemajuan (Raw)
  var dataKemajuan = getRawSheetData("Kemajuan");

  // 3. Dapatkan Data Perakuan Siap (Raw)
  var dataPerakuan = getRawSheetData("Perakuan_Siap"); // Pastikan nama sheet betul

  // 4. Lakukan JOIN Data
  var joinedData = dataProjek.map(function(p) {
    
    // Cari data Kemajuan (SDLC) yang sepadan dengan UNIQUE_ID
    var sdlc = dataKemajuan.find(function(k) { 
      return String(k.PROJEK) === String(p.UNIQUE_ID); 
    }) || {}; 

    // Cari semua Sijil Perakuan Siap yang sepadan dengan UNIQUE_ID
    var financials = dataPerakuan.filter(function(s) {
      return String(s.PROJEK) === String(p.UNIQUE_ID);
    }) || [];

    // Gabungkan ke dalam objek projek
    return {
      ...p,           // Data asal projek (siap dengan kos, status, EOT)
      sdlc: sdlc,     // Data Milestone (URS, UAT, Go Live)
      financials: financials // Senarai sijil pembayaran
    };
  });

  return JSON.stringify(joinedData);
}

// --- HELPER UNTUK MENGAMBIL DATA RAW DARI SHEET ---
function getRawSheetData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Cuba cari nama sheet dengan beberapa variasi
  var sheet = ss.getSheetByName(sheetName) || ss.getSheetByName(sheetName.replace('_', ' ')); 
  
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getDisplayValues();
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      // Trim header untuk elak masalah 'PROJEK ' vs 'PROJEK'
      var headerName = headers[j].toString().trim();
      obj[headerName] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

// --- FUNGSI SOKONGAN LAIN ---

function getKakitanganData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kakitangan");
  if (!sheet) return JSON.stringify([]); 
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rows = data.slice(1);
  const formatted = rows.map(row => {
    let obj = {};
    headers.forEach((h, i) => { obj[h.trim()] = row[i]; });
    return obj;
  });
  return JSON.stringify(formatted);
}

function getJawatanData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jawatan");
  if (!sheet) return JSON.stringify([]); 
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rows = data.slice(1);
  const formatted = rows.map(row => {
    let obj = {};
    headers.forEach((h, i) => { obj[h.trim()] = row[i]; });
    return obj;
  });
  return JSON.stringify(formatted);
}