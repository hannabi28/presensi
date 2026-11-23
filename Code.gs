const SPREADSHEET_ID = '14OYCL9z0EswoIT_ma1YjvoTx5ffZtOAgVSEjHCsNa6s';
const SHEET_NAME = 'Absensi';
const DRIVE_FOLDER_ID = '1BjfDVr13USBx6qh4_gXffxy79UKWlIlh';

function authorizeServices() {
  SpreadsheetApp.openById(SPREADSHEET_ID).getName();
  DriveApp.getFolderById(DRIVE_FOLDER_ID).getName();
  return 'Spreadsheet dan Google Drive siap digunakan';
}

function doGet(e) {
  if (e && e.parameter) {
    if (e.parameter.action === 'cleanup') {
      const result = cleanupOldFiles30Days();
      return json({
        ok: true,
        message: `Pembersihan berhasil: ${result.deletedCount} berkas (>30 hari) telah dipindahkan ke tempat sampah Google Drive. Data rekapan absensi di Spreadsheet dan Dashboard tetap aman tersimpan.`,
        deletedCount: result.deletedCount,
        skippedCount: result.skippedCount
      });
    }
    if (e.parameter.action === 'setupTrigger') {
      const msg = setupDailyCleanupTrigger();
      return json({ ok: true, message: msg });
    }
  }

  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return json({ ok: true, records: [] });

  const headers = values.shift();
  const records = values.map((row, rowIndex) => {
    const record = headers.reduce((acc, header, index) => {
      acc[header] = row[index];
      return acc;
    }, {});
    record._row = rowIndex + 2;
    return record;
  }).map(record => {
    record.photo = normalizePhotoUrl(record.photo);
    if (!record.supportingDocument && record.proof) {
      record.supportingDocument = record.proof;
    }
    return record;
  });

  return json({ ok: true, records: records });
}

function doPost(event) {
  let data;
  try {
    data = JSON.parse(event.postData.contents);
  } catch (error) {
    return json({ ok: false, message: `Data permintaan tidak valid: ${error.message}` });
  }

  let rowNumber = 0;
  let rowWritten = false;
  try {
    const sheet = getSheet();
    const id = data.id || new Date().getTime();

    const row = [
      id,
      data.date || '',
      data.dateLabel || '',
      data.name || '',
      data.department || '',
      data.attendanceType || '',
      data.status || '',
      data.arrival || '',
      data.reason || '',
      formatLocation(data.location),
      '',
      data.createdAt || new Date().toISOString(),
      ''
    ];
    const lastRow = sheet.getLastRow();
    const existing = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext()
      : null;
    rowNumber = existing ? existing.getRow() : lastRow + 1;
    if (existing) {
      row[10] = sheet.getRange(rowNumber, 11).getValue();
      row[12] = sheet.getRange(rowNumber, 13).getValue();
    }
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    rowWritten = true;

    const photo = savePhoto(data.photo, data.name);
    if (!photo.url) throw new Error('Foto tidak valid atau gagal disimpan ke Google Drive');
    sheet.getRange(rowNumber, 11).setValue(photo.url);
    SpreadsheetApp.flush();
    if (sheet.getRange(rowNumber, 11).getValue() !== photo.url) {
      throw new Error('Tautan foto tidak berhasil ditulis ke kolom K spreadsheet');
    }

    if (data.supportingDocument) {
      const document = saveSupportingDocument(data.supportingDocument, data.supportingDocumentName, data.name);
      if (!document.url) throw new Error('Bukti pendukung tidak valid atau gagal disimpan ke Google Drive');
      sheet.getRange(rowNumber, 13).setValue(document.url);
      SpreadsheetApp.flush();
      if (sheet.getRange(rowNumber, 13).getValue() !== document.url) {
        throw new Error('Tautan bukti pendukung tidak berhasil ditulis ke kolom M spreadsheet');
      }
    }

    return json({ ok: true, row: rowNumber, message: 'Absensi berhasil disimpan' });
  } catch (error) {
    if (rowWritten) {
      return json({ ok: false, row: rowNumber, message: `Data absensi tersimpan di spreadsheet (baris ${rowNumber}), tetapi foto atau bukti pendukung gagal disimpan: ${error.message}` });
    }
    return json({ ok: false, message: error.message });
  }
}

function getSheet() {
  if (SPREADSHEET_ID === 'ISI_ID_GOOGLE_SHEET_DI_SINI') {
    throw new Error('SPREADSHEET_ID belum dikonfigurasi');
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'id', 'date', 'dateLabel', 'name', 'department', 'attendanceType',
      'status', 'arrival', 'reason', 'location', 'photo', 'createdAt', 'supportingDocument'
    ]);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('supportingDocument') === -1) {
      sheet.getRange(1, headers.length + 1).setValue('supportingDocument');
    }
  }

  return sheet;
}

function formatLocation(location) {
  if (!location) return '';
  const latitude = Number(location.latitude);
  const longitude = Number(location.longitude);
  const coordinates = Number.isFinite(latitude) && Number.isFinite(longitude)
    ? `Koordinat: ${latitude.toFixed(6)}, ${longitude.toFixed(6)}`
    : '';
  const label = String(location.label || '').trim();
  return [label, coordinates].filter(Boolean).join(' | ');
}

function savePhoto(photoData, name) {
  if (!photoData || !photoData.startsWith('data:image/')) return '';

  const match = photoData.match(/^data:(image\/[^;]+);base64,(.+)$/);
  if (!match) return '';

  const folder = DRIVE_FOLDER_ID
    ? DriveApp.getFolderById(DRIVE_FOLDER_ID)
    : DriveApp.getRootFolder();
  const extension = match[1].split('/')[1].replace('jpeg', 'jpg');
  const safeName = String(name || 'peserta').replace(/[^\w-]+/g, '-');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(match[2]),
    match[1],
    `${safeName}-${new Date().getTime()}.${extension}`
  );
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    Logger.log('Izin berbagi per berkas dilewati/dibatasi oleh organisasi: ' + error.message);
  }
  return { url: file.getUrl() };
}

function saveSupportingDocument(dataUrl, fileName, name) {
  const match = String(dataUrl || '').match(/^data:(application\/pdf|image\/(?:jpeg|png|webp));base64,([\s\S]+)$/);
  if (!match) throw new Error('Bukti pendukung harus berupa PDF, JPEG, PNG, atau WEBP.');

  const mimeType = match[1];
  const content = Utilities.base64Decode(match[2]);
  if (content.length > 5 * 1024 * 1024) throw new Error('Bukti pendukung tidak boleh lebih besar dari 5 MB.');
  const extension = mimeType === 'application/pdf' ? 'pdf' : mimeType.split('/')[1].replace('jpeg', 'jpg');
  const safeName = String(name || 'peserta').replace(/[^\w-]+/g, '-');
  const safeFileName = String(fileName || 'bukti-pendukung').replace(/[^\w.-]+/g, '-');
  const blob = Utilities.newBlob(
    content,
    mimeType,
    `${safeName}-bukti-${new Date().getTime()}-${safeFileName || `dokumen.${extension}`}`
  );
  const folder = DRIVE_FOLDER_ID
    ? DriveApp.getFolderById(DRIVE_FOLDER_ID)
    : DriveApp.getRootFolder();
  const file = folder.createFile(blob);
  return { url: file.getUrl() };
}

function normalizePhotoUrl(url) {
  if (!url) return '';
  const value = String(url);
  if (value.includes('drive.google.com/thumbnail')) return value;

  const match = value.match(/\/file\/d\/([^/]+)/) || value.match(/[?&]id=([^&]+)/);
  return match ? `https://drive.google.com/thumbnail?id=${match[1]}&sz=w1600` : value;
}

function json(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menghapus/memindahkan ke tempat sampah berkas foto dan bukti pendukung di Google Drive
 * yang telah berumur lebih dari 30 hari.
 * Catatan: Data rekapan absensi di Spreadsheet dan Dashboard TIDAK DIHAPUS dan tetap tersimpan aman.
 */
function cleanupOldFiles30Days() {
  const folder = DRIVE_FOLDER_ID
    ? DriveApp.getFolderById(DRIVE_FOLDER_ID)
    : DriveApp.getRootFolder();
  const files = folder.getFiles();
  const now = new Date().getTime();
  const thirtyDaysInMillis = 30 * 24 * 60 * 60 * 1000;
  let deletedCount = 0;
  let skippedCount = 0;
  const failures = [];

  while (files.hasNext()) {
    const file = files.next();
    const createdTime = file.getDateCreated().getTime();
    if (now - createdTime > thirtyDaysInMillis) {
      try {
        file.setTrashed(true);
        deletedCount++;
      } catch (error) {
        const failure = `${file.getName()}: ${error.message}`;
        failures.push(failure);
        Logger.log('Gagal memindahkan file ke sampah: ' + failure);
      }
    } else {
      skippedCount++;
    }
  }

  const summary = `Pembersihan Drive selesai: ${deletedCount} berkas (>30 hari) dipindahkan ke tempat sampah. ${skippedCount} berkas berumur <= 30 hari dipertahankan.`;
  Logger.log(summary);
  if (failures.length) {
    throw new Error(`${summary} ${failures.length} berkas gagal dibersihkan: ${failures.slice(0, 10).join('; ')}`);
  }
  return { ok: true, deletedCount: deletedCount, skippedCount: skippedCount };
}

/**
 * Memasang trigger otomatis harian di Google Apps Script (berjalan setiap hari jam 01.00 pagi)
 * sehingga berkas berumur lebih dari 30 hari di Google Drive otomatis terhapus secara berkala.
 */
function setupDailyCleanupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'cleanupOldFiles30Days') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('cleanupOldFiles30Days')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();

  Logger.log('Trigger harian pembersihan berkas >30 hari berhasil dipasang.');
  return 'Trigger pembersihan otomatis harian berhasil dipasang. Berkas di Google Drive yang melewati 30 hari akan otomatis dipindahkan ke sampah.';
}
