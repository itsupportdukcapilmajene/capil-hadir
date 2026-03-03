//Version 28 on 3 Mar 2026, 11:50
//https://script.google.com/macros/s/AKfycbw6L7Y6W6daOdZz0q8nYQAWQtUsJv9S2MOfXVwYnqTmtnWYTXVC6wxFMkagZ6mkcL4P/exec

/************** KONFIGURASI INTI (SESUAIKAN) **************/
const APP_KEY        = 'absendukcapilmajene';
const SPREADSHEET_ID = '11lU4f6s5cMBMMEIftwr1B0mRQO4s8RFQ4kw82rWm1AI';
const FOLDER_ID      = '1XuISKcA79uKBEOOYmvZoGQ5YnXpRAN1o';
const TZ             = 'Asia/Makassar';

const ROLES = { REGULER:'reguler', PENJAGA_MALAM:'penjaga malam', KEBERSIHAN:'kebersihan' };

/************** ENTRY POINT **************/
function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
   

    // CONFIG fetch (tanpa NIP/Nama)
    if (p.config === 'true') {
      const cfg = getConfig_();
      const resp = cfg.ok ? cfg : { ok:false, message:'Config not OK', code:'CONFIG_ERROR' };
      return logAndReturn(resp, { aksi:'CONFIG', serverTime:tsNow_() });
    }

    // Open sheets
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const shPegawai = ss.getSheetByName('Pegawai');
    const shAbsen   = ss.getSheetByName('Absensi');
    if (!shPegawai || !shAbsen) {
      const resp = { ok:false, message:'❌ Sheet tidak ditemukan', code:'SHEET_NOT_FOUND', serverTime:tsNow_() };
      return logAndReturn(resp, { aksi:'SYSTEM' });
    }

    // Load config
    const CFG = getConfig_();
    if (!CFG.ok) {
      const resp = { ok:false, message:'❌ Config error: '+(CFG.error||''), code:'CONFIG_ERROR', serverTime:tsNow_() };
      return logAndReturn(resp, { aksi:'CONFIG' });
    }

    // Params
    const nip         = (p.nip || '').trim();
    // ===== RATE LIMIT =====
if(nip){
  if(!checkRateLimit_(nip)){
    return jsonOut({
      ok:false,
      message:'⛔ Terlalu banyak percobaan. Tunggu 1 menit.',
      code:'RATE_LIMIT',
      serverTime: tsNow_()
    });
  }
}
    const status      = (p.status || '').toLowerCase(); // izin/sakit/pulang cepat
    const mode       = (p.mode || '').toUpperCase(); // FULL / PARTIAL
const tglMulai   = p.tglMulai || '';
const tglSelesai = p.tglSelesai || '';
    const ket         = p.keterangan || '';
    const riwayatOnly = p.riwayat === 'true';
    const force       = p.force === 'true';
    const checkOnly   = p.checkOnly === 'true';
    const apel        = (p.apel || '').toLowerCase() === 'true';

    const lat      = safeNum_(p.lat);
    const lng      = safeNum_(p.lng);
    const accuracy = safeNum_(p.accuracy);

    const fileBase64 = p.file || '';
    const fileType   = (p.fileType || '').toLowerCase();
    const fileName   = p.fileName || '';

    const selfieBase64 = p.selfie || '';
    const selfieType   = (p.selfieType || '').toLowerCase();
    const selfieName   = p.selfieName || 'selfie.jpg';

    // Waktu
    const now = new Date();
    const jamFull = Utilities.formatDate(now, TZ, 'HH:mm:ss');
    const tanggalFull = Utilities.formatDate(now, TZ, 'EEEE, dd MMMM yyyy');
    const jamHHmm = Utilities.formatDate(now, TZ, 'HH:mm');
    const jamTotal = toMinutes_(jamHHmm);

    // Pegawai
    const pegData = shPegawai.getDataRange().getValues(); // A:NIP, B:Nama, G:Jadwal, H:Role
    const rowPeg  = pegData.slice(1).find(rec => String(rec[0]) === nip);
    const nama    = rowPeg ? String(rowPeg[1]||'') : '';
    
    if (!nama) {

  if(!checkNipFailLock_(nip,false)){
    return logAndReturn({
      ok:false,
      message:'⛔ NIP diblokir sementara (terlalu banyak percobaan).',
      code:'NIP_LOCKED',
      serverTime: jamFull
    }, { nip });
  }

  return logAndReturn({
    ok:false,
    message:'❌ NIP tidak terdaftar',
    code:'NIP_NOT_FOUND',
    serverTime: jamFull
  }, { nip });
}
checkNipFailLock_(nip,true);

    const jadwal   = (rowPeg[6]||'full').toLowerCase(); // full/tanggal genap/tanggal ganjil
    const role     = getRole_(rowPeg, ROLES);

    // Tanggal efektif
    const eff     = getTanggalEfektifDate_(now, TZ, role, ROLES);
    const tanggal = eff.tanggal;  // dd/MM/yyyy
    const hari    = eff.hari;

// === DINAS LUAR (override radius) ===
const dinasLuar = isDinasLuar_(tanggal, nip);
const IS_DINAS_LUAR = dinasLuar.active === true;

    const tanggalNum = parseInt(tanggal.split('/')[0],10);
    const isGenap = tanggalNum % 2 === 0;

    // Libur & jadwal
    const isKalenderLibur = isHariLibur_(tanggal) || ['Sabtu','Minggu'].includes(hari);
    const isKerjaTambahan = isKerjaTambahan_(tanggal);
    const isLiburKalenderFinal = isKerjaTambahan ? false : isKalenderLibur;

    const skipJadwal = (role===ROLES.PENJAGA_MALAM || role===ROLES.KEBERSIHAN);
    const bolehMasuk = skipJadwal ? true : (
      jadwal==='full' ||
      (jadwal==='tanggal genap' && (isGenap || ['Senin','Jumat'].includes(hari))) ||
      (jadwal==='tanggal ganjil' && (!isGenap || ['Senin','Jumat'].includes(hari)))
    );

    // RIWAYAT SAJA
    if (riwayatOnly) {
      if (role === ROLES.PENJAGA_MALAM) {
        const r1 = ambilRiwayatSingle_(ss, 'Absen_PenjagaMalam', nip, tanggal);
        const resp = { ok:true, message:r1.found?'📄 Riwayat berhasil dimuat.':'📭 Belum ada riwayat.', nama, role, riwayat:r1.riwayat, serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'RIWAYAT_SINGLE', lat, lng, accuracy });
      } else if (role === ROLES.KEBERSIHAN) {
        const r2 = ambilRiwayatSingle_(ss, 'Absen_Kebersihan', nip, tanggal);
        const resp = { ok:true, message:r2.found?'📄 Riwayat berhasil dimuat.':'📭 Belum ada riwayat.', nama, role, riwayat:r2.riwayat, serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'RIWAYAT_SINGLE', lat, lng, accuracy });
      } else {
        const r3 = ambilRiwayatHariIni_(shAbsen, nip, tanggal, TZ);
        r3.riwayat['Absen Apel'] = sudahApelHariIni_(ss, nip, tanggal) ? 'Sudah Apel' : '-';
        const resp = { ok:true, message:r3.found?'📄 Riwayat berhasil dimuat.':'📭 Belum ada riwayat.', nama, role, riwayat:r3.riwayat, serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'RIWAYAT_REGULER', lat, lng, accuracy });
      }
    }


// PENGAJUAN IZIN / SAKIT / PULANG CEPAT (SLOT-BASED)
if (['izin','sakit','pulang cepat'].includes(status)) {

  const shIzin = ss.getSheetByName('Izin') || ss.insertSheet('Izin');
  ensureHeader_(shIzin, [
    'Timestamp','Tanggal','NIP','Nama',
    'Slot','Jenis','Keterangan','BuktiURL','Status'
  ]);

  let inserted = 0;

  const buktiUrl = fileBase64
    ? uploadToDrive_(fileBase64, fileType||'application/octet-stream',
        `${nip}_${tanggal}_${sanitizeFileName_(fileName||'bukti')}`)
    : '';

  // ================================
  // PARTIAL / PULANG CEPAT
  // ================================
  if (mode === 'PARTIAL' || status === 'pulang cepat') {

    let slotIzin = (p.slot || '').toUpperCase();
    if (status === 'pulang cepat') slotIzin = 'PULANG';

    if (!['PAGI','SIANG','PULANG'].includes(slotIzin)) {
      return logAndReturn({
        ok:false,
        message:'❌ Slot tidak valid.',
        code:'INVALID_SLOT',
        serverTime:jamFull
      }, { nip, nama, aksi:'IZIN' });
    }

    if (!getIzinSlotAnyStatus_(tanggal, nip, slotIzin)) {

      shIzin.appendRow([
        new Date(),
        tanggal,
        nip,
        nama,
        slotIzin,
        status.toUpperCase(),
        ket,
        buktiUrl,
        'PENDING'
      ]);

      inserted++;
    }
  }

  // ================================
  // FULL (MULTI DAY SUPPORT)
  // ================================
  if (mode === 'FULL' && tglMulai && tglSelesai) {

const start = new Date(tglMulai + 'T00:00:00');
const end   = new Date(tglSelesai + 'T00:00:00');

    if (start > end) {
  return logAndReturn({
    ok:false,
    message:'❌ Tanggal selesai tidak boleh lebih kecil dari tanggal mulai.',
    code:'INVALID_DATE_RANGE',
    serverTime:jamFull
  }, { nip, nama, aksi:'IZIN' });
}

    for (let d = new Date(start); d <= end; d.setDate(d.getDate()+1)) {

const tglLoop = Utilities.formatDate(d, TZ, 'dd/MM/yyyy');
const hariLoop = Utilities.formatDate(d, TZ, 'EEEE');
const mapHari = {
  Sunday:'Minggu', Monday:'Senin', Tuesday:'Selasa',
  Wednesday:'Rabu', Thursday:'Kamis', Friday:'Jumat', Saturday:'Sabtu'
};
const hariId = mapHari[hariLoop];

// SKIP HARI LIBUR & WEEKEND
if (isHariLibur_(tglLoop) || ['Sabtu','Minggu'].includes(hariId)) {
  continue;
}

      ['PAGI','SIANG','PULANG'].forEach(slotLoop => {

        if (!getIzinSlotAnyStatus_(tglLoop, nip, slotLoop)) {

          shIzin.appendRow([
            new Date(),
            tglLoop,
            nip,
            nama,
            slotLoop,
            status.toUpperCase(),
            ket,
            buktiUrl,
            'PENDING'
          ]);

          inserted++;
        }

      });
    }
  }

  if (inserted === 0) {
    return logAndReturn({
      ok:false,
      message:'⚠️ Tidak ada data yang ditambahkan (mungkin sudah ada izin atau hari libur).',
      code:'IZIN_SKIPPED',
      serverTime:jamFull
    }, { nip, nama, aksi:'IZIN' });
  }

  return logAndReturn({
    ok:true,
    message:`📌 ${inserted} slot izin berhasil diajukan.`,
    code:'IZIN_MULTI_RECORDED',
    serverTime:jamFull
  }, { nip, nama, aksi:'IZIN_MULTI', note:buktiUrl||'' });
}


// VALIDASI GEO (DINAS LUAR = OVERRIDE)
if (!IS_DINAS_LUAR) {
  if (Number.isFinite(lat)&&Number.isFinite(lng)) {
    const distance = haversine_(lat,lng,CFG.office.lat,CFG.office.lng);
    if (Number.isFinite(accuracy) && CFG.geo.max_accuracy_m && accuracy>CFG.geo.max_accuracy_m && !force) {
      const resp = { ok:false, warning:true, code:'GEO_ACCURACY_WEAK', message:`📶 Akurasi GPS lemah (${Math.round(accuracy)} m).`, serverTime:jamFull };
      return logAndReturn(resp, { nip, nama, aksi:'GEO_CHECK', lat, lng, accuracy });
    }
    if (CFG.geo.max_radius_m && distance>CFG.geo.max_radius_m && !force) {
      const resp = { ok:false, warning:true, code:'GEO_OUT_OF_RADIUS', message:`📍 Di luar radius kantor (${Math.round(distance)} m).`, serverTime:jamFull };
      return logAndReturn(resp, { nip, nama, aksi:'GEO_CHECK', lat, lng, accuracy, note:'dist='+Math.round(distance) });
    }
  } else if (!force) {
    const resp = { ok:false, warning:true, code:'GEO_LOCATION_INVALID', message:'📍 Lokasi tidak valid.', serverTime:jamFull };
    return logAndReturn(resp, { nip, nama, aksi:'GEO_CHECK' });
  }
}

    // PRE-CHECK (tanpa tulis)
if (checkOnly) {

  if (apel) {

    if (role !== ROLES.REGULER) {
      const resp = { ok:false, warning:true, code:'APEL_ROLE_NOT_ALLOWED', message:'🚫 Apel tidak berlaku untuk peran ini.', serverTime:jamFull };
      return logAndReturn(resp, { nip, nama, aksi:'APEL_PRECHECK', lat, lng, accuracy });
    }

    const allowedDays = ['Senin','Selasa','Rabu','Kamis'];
    if (!allowedDays.includes(hari)) {
      const resp = { ok:false, warning:true, code:'APEL_DAY_NOT_ALLOWED', message:'🚫 Apel hanya Senin–Kamis.', serverTime:jamFull };
      return logAndReturn(resp, { nip, nama, aksi:'APEL_PRECHECK', lat, lng, accuracy });
    }

    if (!inWindow_(jamTotal, CFG.apel.start, CFG.apel.end)) {
      const jamApel = `${CFG.apel.start}–${CFG.apel.end}`;
      const resp = { ok:false, warning:true, code:'APEL_TIME_WINDOW', message:`🚫 Apel hanya pada ${jamApel} Wita.`, serverTime:jamFull };
      return logAndReturn(resp, { nip, nama, aksi:'APEL_PRECHECK', lat, lng, accuracy });
    }

  } else {

    if ((!bolehMasuk || isLiburKalenderFinal) && !skipJadwal && !force) {
      const namaLibur = isLiburKalenderFinal ? getNamaHariLibur_(tanggal) : '';

      const resp = {
        ok:false,
        warning:true,
        code:'OUT_OF_SCHEDULE',
        message:`⚠️ Hari ini ${hari}${namaLibur ? ' ('+namaLibur+')' : ''}, bukan jadwal kerja Anda.`,
        serverTime:jamFull
      };

      return logAndReturn(resp, { nip, nama, aksi:'REGULER_PRECHECK', lat, lng, accuracy });
    }
  }

  const resp = {
    ok:true,
    warning:false,
    code:'PRECHECK_OK',
    message:'✅ Valid, lanjutkan.',
    nama:nama,
    serverTime:jamFull,
    serverDate:tanggalFull
  };

  return logAndReturn(resp, { nip, nama, aksi: apel?'APEL_PRECHECK':'REGULER_PRECHECK', lat, lng, accuracy });
}


    // APEL (reguler; tanpa selfie)
    if (apel) {
      if (role!==ROLES.REGULER) {
        const resp = { ok:false, message:'🚫 Apel hanya untuk pegawai reguler.', code:'APEL_ROLE_NOT_ALLOWED', serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'APEL', lat, lng, accuracy });
      }
      const allowedDays = ['Senin','Selasa','Rabu','Kamis'];
      if (!allowedDays.includes(hari)) {
        const resp = { ok:false, message:'🚫 Apel hanya Senin–Kamis.', code:'APEL_DAY_NOT_ALLOWED', serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'APEL', lat, lng, accuracy });
      }
      if (!inWindow_(jamTotal, CFG.apel.start, CFG.apel.end)) {
        const jamApel = `${CFG.apel.start}–${CFG.apel.end}`;
        const resp = { ok:false, message:`🚫 Apel hanya pada ${jamApel} Wita.`, code:'APEL_TIME_WINDOW', serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'APEL', lat, lng, accuracy });
      }
      const shApel = ss.getSheetByName('Absen_Apel') || ss.insertSheet('Absen_Apel');
      ensureHeader_(shApel, ['Timestamp','NIP','Nama','Jam','Lat','Lng','Accuracy']);
      const sudahApel = sudahApelHariIni_(ss, nip, tanggal);
      if (sudahApel) {
        const resp = { ok:false, message:'⚠️ Anda sudah absen apel hari ini.', code:'APEL_ALREADY_DONE', serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'APEL', lat, lng, accuracy });
      }
      shApel.appendRow([new Date(), nip, nama, jamHHmm, lat||'', lng||'', accuracy||'' ]);
      const resp = { ok:true, message:`✅ Apel berhasil pukul ${jamHHmm}`, code:'APEL_OK', nama, role, serverTime:jamFull };
      return logAndReturn(resp, { nip, nama, aksi:'APEL', lat, lng, accuracy });
    }

    /************** ABSEN 1x (PENJAGA MALAM & KEBERSIHAN) **************/
    if (role === ROLES.PENJAGA_MALAM || role === ROLES.KEBERSIHAN) {
      const sheetName = (role === ROLES.PENJAGA_MALAM) ? 'Absen_PenjagaMalam' : 'Absen_Kebersihan';
      const shSingle = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
      ensureHeader_(shSingle, ['Timestamp','NIP','Nama','Jam','Lat','Lng','Accuracy','SelfieURL']);

      const dataSingle = shSingle.getDataRange().getValues();
      const sudahSingle = dataSingle.slice(1).some(rec => {
        const tgl = Utilities.formatDate(new Date(rec[0]), TZ, 'dd/MM/yyyy');
        return tgl === tanggal && String(rec[1]) === nip;
      });
      if (sudahSingle) {
        const resp = { ok:false, message:'⚠️ Anda sudah absen hari ini.', code:'ALREADY_CHECKED_IN', serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'SINGLE', lat, lng, accuracy });
      }

      if (CFG.upload.selfie_required && !isImageMime_(selfieType)) {
        const resp = { ok:false, message:'📸 Selfie harus berupa gambar.', code:'SELFIE_MIME_INVALID', serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'SINGLE', lat, lng, accuracy });
      }
      if (CFG.upload.selfie_required && tooBig_(selfieBase64, (CFG.upload.max_selfie_mb||5))) {
        const resp = { ok:false, message:`📸 Selfie terlalu besar (> ${CFG.upload.max_selfie_mb||5}MB).`, code:'SELFIE_TOO_LARGE', serverTime:jamFull };
        return logAndReturn(resp, { nip, nama, aksi:'SINGLE', lat, lng, accuracy });
      }

      const selfieUrl = selfieBase64
        ? uploadToDrive_(selfieBase64, selfieType||'image/jpeg', `${nip}_${tanggal}_SELFIE_${jamFull.replace(/[:]/g,'-')}_${sanitizeFileName_(selfieName)}`)
        : '';

      shSingle.appendRow([new Date(), nip, nama, jamHHmm, lat||'', lng||'', accuracy||'', selfieUrl ]);
      const resp = { ok:true, message:`✅ Kehadiran tercatat pukul ${jamHHmm}`, code:'SINGLE_OK', nama, role, riwayat:{ 'Kehadiran': jamHHmm }, serverTime:jamFull };
      return logAndReturn(resp, { nip, nama, aksi:'SINGLE', lat, lng, accuracy, note:selfieUrl||'' });
    }

/************** ABSEN REGULER **************/

return handleReguler_({
  ss,
  shAbsen,
  CFG,
  nip,
  nama,
  role,
  hari,
  tanggal,
  eff,
  jamHHmm,
  jamFull,
  jamTotal,
  lat,
  lng,
  accuracy,
  selfieBase64,
  selfieType,
  selfieName
});

  } catch (err) {
    return logAndReturn(
      {
        ok:false,
        message:'❌ Terjadi kesalahan server.',
        code:'SERVER_ERROR',
        serverTime: tsNow_()
      },
      { aksi:'SYSTEM_ERROR', note:String(err) }
    );
  }
}



/************** LOGGING **************/

function logAndReturn(resp, ctx) {
  try {
    logEvent_(
      (ctx && ctx.nip) || '',
      (ctx && ctx.nama) || '',
      (ctx && ctx.aksi) || '',
      resp.code || '',
      resp.message || '',
      !!resp.ok,
      (ctx && ctx.lat),
      (ctx && ctx.lng),
      (ctx && ctx.accuracy),
      resp.serverTime || (ctx && ctx.serverTime) || '',
      (ctx && ctx.note) || ''
    );
  } catch (_e) {}
  return jsonOut(resp);
}

function logEvent_(nip, nama, aksi, code, message, ok, lat, lng, accuracy, serverTime, note) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Logs') || ss.insertSheet('Logs');
  ensureHeader_(sh, ['Timestamp','NIP','Nama','Aksi','Code','Message','OK','Lat','Lng','Accuracy','ServerTime','Note']);
  sh.appendRow([new Date(), nip||'', nama||'', aksi||'', code||'', message||'', ok, lat||'', lng||'', accuracy||'', serverTime||'', note||'']);
}

function tsNow_(){ return Utilities.formatDate(new Date(), TZ, 'HH:mm:ss'); }

/************** CONFIG & WINDOWS **************/
function getConfig_(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Config');
  if (!sh) return { ok:false, error:"Sheet 'Config' tidak ditemukan" };

  const rows = sh.getDataRange().getValues();
  const kv = {};

  rows.forEach(r => {
    const k = String(r[0]||'').trim();
    let v = r[1];
    if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'HH:mm');
    else v = String(v||'').trim();
    if (k) kv[k] = v;
  });

  const toNum  = (s,d=null)=>{ const n=Number(s); return Number.isFinite(n)?n:d; };
  const toBool = (s,d=false)=>{
    const t=String(s||'').toLowerCase();
    if(['true','1','yes','y'].includes(t)) return true;
    if(['false','0','no','n'].includes(t)) return false;
    return d;
  };

const cfg = {
  ok: true,
  office: {
    lat: toNum(kv['OFFICE_LAT']),
    lng: toNum(kv['OFFICE_LNG'])
  },
  geo: {
    max_radius_m:   toNum(kv['MAX_RADIUS_M'], 50),
    max_accuracy_m: toNum(kv['MAX_ACCURACY_M'], 75)
  },
  apel: {
    start: kv['APEL_START'] || '07:30',
    end:   kv['APEL_END']   || '08:00'
  },
  upload: {
    selfie_required: toBool(kv['SELFIE_REQUIRED'], true),
    max_selfie_mb:   toNum(kv['MAX_SELFIE_MB'], 5),
    max_bukti_mb:    toNum(kv['MAX_BUKTI_MB'], 10)
  },

  // 🔐 TAMBAHAN INI
  policy: {
    allow_late_absen: toBool(kv['ALLOW_LATE_ABSEN'], false)
  }
};


  cfg.win_simple = {
    SENKAM: {
      masukStart:  kv['SENKAM_MASUK_START'],
      masukEnd:    kv['SENKAM_MASUK_END'],
      siangStart:  kv['SENKAM_SIANG_START'],
      siangEnd:    kv['SENKAM_SIANG_END'],
      pulangStart: kv['SENKAM_PULANG_START'],
      pulangEnd:   kv['SENKAM_PULANG_END']
    },
    JUMAT: {
      masukStart:  kv['JUM_MASUK_START'],
      masukEnd:    kv['JUM_MASUK_END'],
      siangStart:  kv['JUM_SIANG_START'],
      siangEnd:    kv['JUM_SIANG_END'],
      pulangStart: kv['JUM_PULANG_START'],
      pulangEnd:   kv['JUM_PULANG_END']
    }
  };

  if (cfg.office.lat == null || cfg.office.lng == null) {
    cfg.ok = false;
    cfg.error = "OFFICE_LAT / OFFICE_LNG belum diisi di sheet Config.";
  }

  return cfg;
}




function getWindowsFromConfig_(CFG, hari){
  const w = CFG.win_simple || {};
  const isJumat = (hari === 'Jumat');
  const src = isJumat ? w.JUMAT : w.SENKAM;

  const toM = v => v ? toMinutes_(v) : null;

  return {
    masukStart:  toM(src?.masukStart),
    masukEnd:    toM(src?.masukEnd),
    siangStart:  toM(src?.siangStart),
    siangEnd:    toM(src?.siangEnd),
    pulangStart: toM(src?.pulangStart),
    pulangEnd:   toM(src?.pulangEnd)
  };
}

function handleLateSlot_(params) {
  const {
    shAbsen, ss, nip, nama, tanggal, jamHHmm, jamFull,
    slotObj, rowIndex, existingRow, selfieBase64,
    selfieType, selfieName, CFG, lat, lng, accuracy
  } = params;

  const row = rowIndex + 1;

   // 🔒 ANTI DUPLICATE
  if (existingRow[slotObj.timeCol-1]) {
    return {
      error:'ALREADY_FILLED',
      msg:`⚠️ ${slotObj.jenis} sudah dilakukan sebelumnya.`
    };
  }

  if (!existingRow[slotObj.timeCol-1]) {
    shAbsen.getRange(row, slotObj.timeCol)
      .setNumberFormat('@STRING@')
      .setValue(jamHHmm);
  }

// 🔒 JANGAN UBAH STATUS JIKA SUDAH HADIR
if (!existingRow[6]) {
  shAbsen.getRange(row,7).setValue('TELAT');
}

  // Upload selfie (jika wajib)
  if (CFG.upload.selfie_required) {

    if (!isImageMime_(selfieType)) {
      return { error:'SELFIE_MIME_INVALID', msg:'📸 Selfie harus berupa gambar.' };
    }

    if (tooBig_(selfieBase64, (CFG.upload.max_selfie_mb||5))) {
      return { error:'SELFIE_TOO_LARGE', msg:`📸 Selfie terlalu besar (> ${CFG.upload.max_selfie_mb||5}MB).` };
    }

    const selfieUrl = selfieBase64
      ? uploadToDrive_(
          selfieBase64,
          selfieType || 'image/jpeg',
          `${nip}_${tanggal}_${slotObj.jenis.replace(' ','').toUpperCase()}_LATE_${jamFull.replace(/[:]/g,'-')}_${sanitizeFileName_(selfieName)}`
        )
      : '';

    if (selfieUrl) {
      shAbsen.getRange(row, slotObj.selfieCol).setValue(selfieUrl);
    }
  }

  const ketCell = shAbsen.getRange(row,8);
  const oldKet = String(ketCell.getValue() || '');
  const frag = `Telat absen (${slotObj.jenis})`;
  ketCell.setValue(oldKet ? oldKet + '; ' + frag : frag);

  const rAll = ambilRiwayatHariIni_(shAbsen, nip, tanggal, TZ);
  rAll.riwayat['Absen Apel'] =
    sudahApelHariIni_(ss, nip, tanggal) ? 'Sudah Apel' : '-';

  rAll.riwayat[slotObj.jenis] = jamHHmm;

  return {
    success:true,
    riwayat:rAll.riwayat
  };
}

function handleReguler_(ctx) {

  const {
    ss, shAbsen, CFG,
    nip, nama, role, hari, tanggal, eff,
    jamHHmm, jamFull, jamTotal,
    lat, lng, accuracy,
    selfieBase64, selfieType, selfieName
  } = ctx;

  const win = getWindowsFromConfig_(CFG, hari);

  let jenis='', timeCol=0, selfieCol=0;

  if (between_(jamTotal, win.masukStart, win.masukEnd)) {
    jenis='Absen Masuk'; timeCol=4; selfieCol=10;
  }
  else if (between_(jamTotal, win.siangStart, win.siangEnd)) {
    jenis='Absen Siang'; timeCol=5; selfieCol=11;
  }
  else if (between_(jamTotal, win.pulangStart, win.pulangEnd)) {
    jenis='Absen Pulang'; timeCol=6; selfieCol=12;
  }
  else {
    return handleRegulerOutsideWindow_(ctx, win);
  }

  // === DALAM WINDOW NORMAL ===

  let idx = findRowIndex_(shAbsen, nip, tanggal);
  if (idx < 0) {
    shAbsen.appendRow([eff.dateObj,nip,nama,'','','','','','','','','']);
    idx = shAbsen.getLastRow()-1;
  }

const row = idx + 1;

// 🔒 ANTI DUPLICATE (WINDOW NORMAL)
const existing = shAbsen.getRange(row,1,1,12).getValues()[0];

if (existing[timeCol-1]) {
  return logAndReturn(
    {
      ok:false,
      message:`⚠️ ${jenis} sudah dilakukan sebelumnya.`,
      code:'ALREADY_FILLED',
      serverTime:jamFull
    },
    { nip, nama, aksi:jenis, lat, lng, accuracy }
  );
}

// ✅ Baru tulis jika kosong
shAbsen.getRange(row,timeCol)
  .setNumberFormat('@STRING@')
  .setValue(jamHHmm);

// 🔒 Jangan ubah TELAT jadi HADIR
if (!existing[6]) {
  shAbsen.getRange(row,7).setValue('HADIR');
}

  if (CFG.upload.selfie_required) {

    if (!isImageMime_(selfieType))
      return logAndReturn(
        { ok:false, message:'📸 Selfie harus berupa gambar.', code:'SELFIE_MIME_INVALID', serverTime:jamFull },
        { nip, nama, aksi:jenis }
      );

    if (tooBig_(selfieBase64, (CFG.upload.max_selfie_mb||5)))
      return logAndReturn(
        { ok:false, message:`📸 Selfie terlalu besar (> ${CFG.upload.max_selfie_mb||5}MB).`, code:'SELFIE_TOO_LARGE', serverTime:jamFull },
        { nip, nama, aksi:jenis }
      );

    const selfieUrl = selfieBase64
      ? uploadToDrive_(selfieBase64, selfieType||'image/jpeg',
        `${nip}_${tanggal}_${jenis.replace(' ','').toUpperCase()}_${jamFull.replace(/[:]/g,'-')}_${sanitizeFileName_(selfieName)}`)
      : '';

    if (selfieUrl)
      shAbsen.getRange(row,selfieCol).setValue(selfieUrl);
  }

  const riwayat = ambilRiwayatHariIni_(shAbsen, nip, tanggal, TZ);
  riwayat.riwayat['Absen Apel'] =
    sudahApelHariIni_(ss, nip, tanggal) ? 'Sudah Apel' : '-';

  return logAndReturn(
    {
      ok:true,
      message:`✅ ${jenis} berhasil pukul ${jamHHmm}`,
      code:'REGULER_OK',
      nama,
      role,
      riwayat:riwayat.riwayat,
      serverTime:jamFull
    },
    { nip, nama, aksi:jenis, lat, lng, accuracy }
  );
}

function handleRegulerOutsideWindow_(ctx, win) {

  const {
    ss, shAbsen, CFG,
    nip, nama, role, hari, tanggal, eff,
    jamHHmm, jamFull, jamTotal,
    lat, lng, accuracy,
    selfieBase64, selfieType, selfieName
  } = ctx;

  // === VALIDASI GEO DULU ===
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
    return logAndReturn(
      { ok:false, message:`⏳ Belum waktunya absen.`, code:'OUT_OF_WINDOW_TOO_EARLY', serverTime:jamFull },
      { nip, nama, aksi:'REGULER', lat, lng, accuracy }
    );
  }

  const dist = haversine_(lat, lng, CFG.office.lat, CFG.office.lng);

  if (CFG.geo.max_radius_m && dist > CFG.geo.max_radius_m) {
    return logAndReturn(
      { ok:false, message:`📍 Di luar radius kantor.`, code:'GEO_OUT_OF_RADIUS', serverTime:jamFull },
      { nip, nama, aksi:'REGULER', lat, lng, accuracy }
    );
  }

  if (CFG.geo.max_accuracy_m && accuracy > CFG.geo.max_accuracy_m) {
    return logAndReturn(
      { ok:false, message:`📶 Akurasi GPS lemah (${Math.round(accuracy)} m).`, code:'GEO_ACCURACY_WEAK', serverTime:jamFull },
      { nip, nama, aksi:'REGULER', lat, lng, accuracy }
    );
  }

  // 🔒 GLOBAL LATE LOCK
if (CFG.policy.allow_late_absen !== true) {
  return logAndReturn(
    {
      ok:false,
      message:'⛔ Waktu absen sudah berakhir.',
      code:'WINDOW_CLOSED',
      serverTime: jamFull
    },
    { nip, nama, aksi:'REGULER', lat, lng, accuracy }
  );
}

  // === DEFINISI SLOT ===
  const slots = [
    { jenis:'Absen Masuk',  start:win.masukStart,  end:win.masukEnd,  timeCol:4, selfieCol:10 },
    { jenis:'Absen Siang',  start:win.siangStart,  end:win.siangEnd,  timeCol:5, selfieCol:11 },
    { jenis:'Absen Pulang', start:win.pulangStart, end:win.pulangEnd, timeCol:6, selfieCol:12 },
  ];

  const prevSlot = [...slots].reverse().find(s => s.end!=null && jamTotal > s.end);

  // === JIKA ADA SLOT SEBELUMNYA (TELAT) ===
  if (prevSlot) {

    let idx = findRowIndex_(shAbsen, nip, tanggal);
    if (idx < 0) {
      shAbsen.appendRow([eff.dateObj,nip,nama,'','','','','','','','','']);
      idx = shAbsen.getLastRow()-1;
    }

const existing = shAbsen.getRange(idx+1,1,1,12).getValues()[0];

// 🔒 GLOBAL SLOT LOCK
if (existing[3] && existing[4] && existing[5]) {
  return logAndReturn(
    {
      ok:false,
      message:'⚠️ Semua slot hari ini sudah terisi.',
      code:'ALL_SLOTS_FILLED',
      serverTime:jamFull
    },
    { nip, nama, aksi:'REGULER', lat, lng, accuracy }
  );
}
    const result = handleLateSlot_({
      shAbsen,
      ss,
      nip,
      nama,
      tanggal,
      jamHHmm,
      jamFull,
      slotObj: prevSlot,
      rowIndex: idx,
      existingRow: existing,
      selfieBase64,
      selfieType,
      selfieName,
      CFG,
      lat,
      lng,
      accuracy
    });

    if (result.error) {
      return logAndReturn(
        { ok:false, message:result.msg, code:result.error, serverTime:jamFull },
        { nip, nama, aksi:prevSlot.jenis, lat, lng, accuracy }
      );
    }

    return logAndReturn(
      {
        ok:true,
        message:`⏰ Dicatat sebagai Telat (${prevSlot.jenis}).`,
        code:'OUT_OF_WINDOW_MARKED_LATE',
        nama,
        role,
        riwayat: result.riwayat,
        serverTime: jamFull
      },
      { nip, nama, aksi:prevSlot.jenis, lat, lng, accuracy }
    );
  }

  // === SEBELUM SLOT PERTAMA ===
  const nextEarly = slots.find(s => s.start!=null && jamTotal < s.start);
  if (nextEarly) {
    const startStr = timeStrFromMinutes_(nextEarly.start);
    const endStr   = nextEarly.end ? ('–'+timeStrFromMinutes_(nextEarly.end)) : '';
    return logAndReturn(
      { ok:false, message:`⏳ Belum waktunya ${nextEarly.jenis}. Jam ${startStr}${endStr}.`, code:'OUT_OF_WINDOW_TOO_EARLY', serverTime:jamFull },
      { nip, nama, aksi:'REGULER', lat, lng, accuracy }
    );
  }

// 🔒 FINAL SAFETY LOCK
if (CFG.policy.allow_late_absen !== true) {
  return logAndReturn(
    {
      ok:false,
      message:'⛔ Waktu absen sudah berakhir.',
      code:'WINDOW_CLOSED',
      serverTime: jamFull
    },
    { nip, nama, aksi:'REGULER', lat, lng, accuracy }
  );
}

  const target = { jenis:'Absen Pulang', timeCol:6, selfieCol:12 };

  let idx2 = findRowIndex_(shAbsen, nip, tanggal);
  if (idx2 < 0) {
    shAbsen.appendRow([eff.dateObj,nip,nama,'','','','','','','','','']);
    idx2 = shAbsen.getLastRow()-1;
  }

  const existing2 = shAbsen.getRange(idx2+1,1,1,12).getValues()[0];

  const result2 = handleLateSlot_({
    shAbsen,
    ss,
    nip,
    nama,
    tanggal,
    jamHHmm,
    jamFull,
    slotObj: target,
    rowIndex: idx2,
    existingRow: existing2,
    selfieBase64,
    selfieType,
    selfieName,
    CFG,
    lat,
    lng,
    accuracy
  });

  if (result2.error) {
    return logAndReturn(
      { ok:false, message:result2.msg, code:result2.error, serverTime:jamFull },
      { nip, nama, aksi:target.jenis, lat, lng, accuracy }
    );
  }

  return logAndReturn(
    {
      ok:true,
      message:`⏰ Dicatat sebagai Telat (${target.jenis}).`,
      code:'OUT_OF_WINDOW_MARKED_LATE',
      nama,
      role,
      riwayat: result2.riwayat,
      serverTime: jamFull
    },
    { nip, nama, aksi:target.jenis, lat, lng, accuracy }
  );
}




function parseRange_(s){
  const arr = String(s||'').split('-').map(t=>t.trim());
  const a = arr[0] || ''; const b = arr[1] || '';
  return { start: toMinutes_(a), end: toMinutes_(b) };
}

const SLOT_MAP = {
  PAGI:   { col: 4, label: 'Absen Masuk' },
  SIANG:  { col: 5, label: 'Absen Siang' },
  PULANG: { col: 6, label: 'Absen Pulang' }
};

/************** HELPERS **************/
function jsonOut(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function safeNum_(v){ const n=Number(v); return Number.isFinite(n)?n:NaN; }
function toMinutes_(hhmm){
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm||'').trim());
  if(!m) return null;
  const h=Number(m[1]), mi=Number(m[2]);
  return h*60+mi;
}
function timeStrFromMinutes_(mins){
  const h = Math.floor(mins/60), m = mins%60;
  return (h<10?'0':'')+h+':' + (m<10?'0':'')+m;
}
function between_(val,start,end){
  if (start==null || end==null || val==null) return false;
  return val>=start && val<=end;
}
function inWindow_(val, startHHMM, endHHMM){
  return between_(val, toMinutes_(startHHMM), toMinutes_(endHHMM));
}

function getRole_(row,ROLES){
  const val=(row&&String(row[7]||'').trim().toLowerCase())||ROLES.REGULER;
  if (val===ROLES.PENJAGA_MALAM || val===ROLES.KEBERSIHAN) return val;
  return ROLES.REGULER;
}
function getTanggalEfektifDate_(now,TZ,role,ROLES){
  const d=new Date(now);
  const hour=Number(Utilities.formatDate(d,TZ,'H'));
  if(role===ROLES.PENJAGA_MALAM && hour<8) d.setDate(d.getDate()-1);
  const tanggal=Utilities.formatDate(d,TZ,'dd/MM/yyyy');
  const hariIng=Utilities.formatDate(d,TZ,'EEEE');
  const map={Sunday:'Minggu',Monday:'Senin',Tuesday:'Selasa',Wednesday:'Rabu',Thursday:'Kamis',Friday:'Jumat',Saturday:'Sabtu'};
  return {tanggal, hari:map[hariIng], dateObj:d};
}

function findRowIndex_(sheet, nip, tanggal){
  const data=sheet.getDataRange().getValues();
  const body=data.slice(1);
  const idxBody=body.findIndex(rec => Utilities.formatDate(new Date(rec[0]), TZ, 'dd/MM/yyyy') === tanggal && String(rec[1]) === nip);
  if (idxBody < 0) return -1;
  return idxBody + 1; // offset header
}

function ambilRiwayatHariIni_(sheet,nip,tanggal,TZ){
  const data=sheet.getDataRange().getValues();
  const body=data.slice(1);
  const idxBody=body.findIndex(rec=>Utilities.formatDate(new Date(rec[0]),TZ,'dd/MM/yyyy')===tanggal && String(rec[1])===nip);
let riwayat={'Absen Masuk':'-','Absen Siang':'-','Absen Pulang':'-'};
let found = idxBody >= 0;

// ⬅️ JIKA ADA BARIS ABSENSI, AMBIL JAM
if (idxBody >= 0) {
  const rowIndex = idxBody + 2;
  const rowData = sheet.getRange(rowIndex,1,1,12).getValues()[0];

  const ket = String(rowData[7] || '');

  const jamMasuk  = renderJamCell_(rowData[3],TZ);
  const jamSiang  = renderJamCell_(rowData[4],TZ);
  const jamPulang = renderJamCell_(rowData[5],TZ);

  riwayat = {
    'Absen Masuk':
      jamMasuk !== '-'
        ? (ket.includes('Absen Masuk') ? jamMasuk + ' (Telat)' : jamMasuk)
        : '-',

    'Absen Siang':
      jamSiang !== '-'
        ? (ket.includes('Absen Siang') ? jamSiang + ' (Telat)' : jamSiang)
        : '-',

    'Absen Pulang':
      jamPulang !== '-'
        ? (ket.includes('Absen Pulang') ? jamPulang + ' (Telat)' : jamPulang)
        : '-'
  };
}


// 🔁 OVERLAY IZIN (SELALU JALAN, ADA ABSEN ATAU TIDAK)
['PAGI','SIANG','PULANG'].forEach(s => {
  const izin = getIzinSlot_(tanggal, nip, s);
  if (izin && riwayat[SLOT_MAP[s].label] === '-') {
    riwayat[SLOT_MAP[s].label] = izin.jenis;
    found = true; // ⬅️ INI KUNCI
  }
});

return { found, riwayat };

}
function renderJamCell_(v,TZ){
  if (!v) return '-';
  if (Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v)) return Utilities.formatDate(v, TZ, 'HH:mm');
  if (typeof v === 'string') { const m=v.match(/^(\d{1,2}:\d{2})(?::\d{2})?$/); return m?m[1]:'-'; }
  return '-';
}

function ambilRiwayatSingle_(ss, sheetName, nip, tanggal) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return { found:false, riwayat:{ 'Kehadiran':'-' } };
  const data = sh.getDataRange().getValues();
  const body = data.slice(1);
  const row = body.find(rec => Utilities.formatDate(new Date(rec[0]), TZ, 'dd/MM/yyyy') === tanggal && String(rec[1]) === nip);
  if (!row) return { found:false, riwayat:{ 'Kehadiran':'-' } };
  const jam = (typeof row[3] === 'string') ? row[3] : Utilities.formatDate(new Date(row[0]), TZ, 'HH:mm');
  return { found:true, riwayat:{ 'Kehadiran': jam } };
}

function haversine_(lat1,lon1,lat2,lon2){
  const R=6371e3,toRad=deg=>deg*Math.PI/180;
  const dLat=toRad(lat2-lat1), dLon=toRad(lon2-lon1);
  const a=Math.sin(dLat/2)**2+Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*Math.sin(dLon/2)**2;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}

function getNamaHariLibur_(tanggal){
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Libur');
  if(!sh) return '';

  const data = sh.getDataRange().getValues();
  const row = data.slice(1).find(rec =>
    Utilities.formatDate(new Date(rec[0]), TZ, 'dd/MM/yyyy') === tanggal
  );

  return row ? String(row[1] || '').trim() : '';
}

function isHariLibur_(tanggal){
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Libur');
  if(!sh) return false;

  const data = sh.getDataRange().getValues();
  return data.slice(1).some(rec =>
    Utilities.formatDate(new Date(rec[0]), TZ, 'dd/MM/yyyy') === tanggal
  );
}


function isKerjaTambahan_(tanggal){
  const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Kerja_Tambahan');
  if(!sh) return false;
  const data=sh.getDataRange().getValues();
  return data.slice(1).some(rec=>Utilities.formatDate(new Date(rec[0]),TZ,'dd/MM/yyyy')===tanggal);
}

function sanitizeFileName_(name){
  name=(name||'file.bin').toString();
  name=name.replace(/[\/\\:*?"<>|]+/g,'_').replace(/[^\w.\-() @]+/g,'_');
  if(name.length>120) name=name.slice(0,120);
  if(!name.trim()) name='file.bin';
  return name;
}
function isImageMime_(mime){ return String(mime||'').toLowerCase().startsWith('image/'); }
function tooBig_(base64, maxMB){ if(!base64) return false; const approx=Math.floor((base64.length*3)/4); return approx > maxMB*1024*1024; }

function uploadToDrive_(base64, mime, fileName){
  const raw = Utilities.base64Decode(base64);
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const file = folder.createFile(Utilities.newBlob(raw, mime||'application/octet-stream', fileName));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}
  return file.getUrl();
}

function ensureHeader_(sh, headers) {
  const last = sh.getLastRow();
  if (last === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else {
    const num = Math.max(headers.length, sh.getLastColumn());
    const first = sh.getRange(1,1,1,num).getValues()[0];
    let same = true;
    for (let i=0;i<headers.length;i++){
      if (String(first[i]||'').toLowerCase() !== String(headers[i]||'').toLowerCase()) { same=false; break; }
    }
    if (!same) {
      sh.insertRows(1);
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }
}

function sudahApelHariIni_(ss, nip, tanggal){
  const sh = ss.getSheetByName('Absen_Apel');
  if (!sh) return false;
  const data = sh.getDataRange().getValues();
  return data.slice(1).some(rec=>{
    const tgl=Utilities.formatDate(new Date(rec[0]),TZ,'dd/MM/yyyy');
    return tgl===tanggal && String(rec[1])===nip;
  });
}

function getIzinSlot_(tanggal, nip, slot) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Izin');
  if (!sh) return null;

  const data = sh.getDataRange().getValues();
  const row = data.slice(1).find(r =>
    Utilities.formatDate(new Date(r[1]), TZ, 'dd/MM/yyyy') === tanggal &&
    String(r[2]) === nip &&
    String(r[4]).toUpperCase() === slot &&
    String(r[8]).toUpperCase() === 'DISETUJUI'
  );

  if (!row) return null;

  return {
    jenis: String(row[5]), // IZIN / SAKIT / PULANG CEPAT
    ket:   String(row[6] || '')
  };
}
function getIzinSlotAnyStatus_(tanggal, nip, slot) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Izin');
  if (!sh) return false;

  const data = sh.getDataRange().getValues();
  return data.slice(1).some(r =>
    Utilities.formatDate(new Date(r[1]), TZ, 'dd/MM/yyyy') === tanggal &&
    String(r[2]) === nip &&
    String(r[4]).toUpperCase() === slot &&
    String(r[8]).toUpperCase() !== 'DITOLAK'
  );
}



/************** DINAS LUAR (OVERRIDE GEO) **************/
function isDinasLuar_(tanggal, nip) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Dinas_Luar');
  if (!sh) return { active:false, keterangan:'' };

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const tgl = Utilities.formatDate(new Date(data[i][0]), TZ, 'dd/MM/yyyy');
    if (tgl === tanggal && String(data[i][1]) === nip) {
      return {
        active: true,
        keterangan: String(data[i][3] || '').trim()
      };
    }
  }
  return { active:false, keterangan:'' };
}

function checkRateLimit_(nip){
  const cache = CacheService.getScriptCache();
  const key = 'RL_' + nip;

  const current = Number(cache.get(key) || 0);

  if(current >= 5){
    return false;
  }

  cache.put(key, current + 1, 60); // 60 detik
  return true;
}

function checkNipFailLock_(nip, success){
  const cache = CacheService.getScriptCache();
  const key = 'FAIL_' + nip;

  if(success){
    cache.remove(key);
    return true;
  }

  const count = Number(cache.get(key) || 0) + 1;

  if(count >= 3){
    cache.put(key, count, 600); // lock 10 menit
    return false;
  }

  cache.put(key, count, 600);
  return true;
}