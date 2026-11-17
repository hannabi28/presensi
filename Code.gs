const SPREADSHEET_ID = '14OYCL9z0EswoIT_ma1YjvoTx5ffZtOAgVSEjHCsNa6s';
const SHEET_NAME = 'Absensi';
const DRIVE_FOLDER_ID = '1BjfDVr13USBx6qh4_gXffxy79UKWlIlh';

function authorizeServices() {
  SpreadsheetApp.openById(SPREADSHEET_ID).getName();
  DriveApp.getFolderById(DRIVE_FOLDER_ID).getName();
  return 'Spreadsheet dan Google Drive siap digunakan';
}

function doGet() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return json({ ok: true, records: [] });

  const headers = values.shift();
  const records = values.map(row => headers.reduce((record, header, index) => {
    record[header] = row[index];
    return record;
  }, {})).map(record => {
    record.photo = normalizePhotoUrl(record.photo);
    return record;
  });

  return json({ ok: true, records: records });
}

function doPost(event) {
  let rowNumber = 0;
  let rowWritten = false;
  try {
    const data = JSON.parse(event.postData.contents);
    const sheet = getSheet();
    const proofColumn = ensureProofColumn(sheet);
    const id = data.id || new Date().getTime();
    const needsMedicalProof = data.status === 'Ijin' && /\bsakit\b/i.test(String(data.reason || ''));
    if (needsMedicalProof && (!data.proof || !data.proof.data)) {
      throw new Error('Surat dokter wajib diunggah untuk izin sakit');
    }

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
      data.createdAt || new Date().toISOString()
    ];
    const lastRow = sheet.getLastRow();
    const existing = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext()
      : null;
    rowNumber = existing ? existing.getRow() : lastRow + 1;
    if (existing) row[10] = sheet.getRange(rowNumber, 11).getValue();
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

    let proof = null;
    if (data.proof) {
      proof = saveProof(data.proof, data.name);
      sheet.getRange(rowNumber, proofColumn).setValue(proof.url);
      SpreadsheetApp.flush();
      if (sheet.getRange(rowNumber, proofColumn).getValue() !== proof.url) {
        throw new Error('Tautan bukti tidak berhasil ditulis ke spreadsheet');
      }
    }

    const sharingErrors = [photo.sharingError, proof && proof.sharingError].filter(Boolean);
    if (sharingErrors.length) {
      return json({
        ok: true,
        row: rowNumber,
        warning: `File sudah disimpan, tetapi izin berbagi Drive gagal: ${sharingErrors.join('; ')}`
      });
    }

    return json({ ok: true, row: rowNumber, message: 'Absensi berhasil disimpan' });
  } catch (error) {
    if (rowWritten) {
      return json({ ok: false, row: rowNumber, message: `Data absensi tersimpan di spreadsheet (baris ${rowNumber}), tetapi penyimpanan foto atau bukti gagal: ${error.message}` });
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
      'status', 'arrival', 'reason', 'location', 'photo', 'createdAt', 'proof'
    ]);
  }

  ensureProofColumn(sheet);
  return sheet;
}

function ensureProofColumn(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const existing = headers.indexOf('proof');
  if (existing >= 0) return existing + 1;

  const column = sheet.getLastColumn() + 1;
  sheet.getRange(1, column).setValue('proof');
  return column;
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
  let sharingError = '';
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    sharingError = error.message;
  }
  return { url: file.getUrl(), sharingError: sharingError };
}

function saveProof(proofData, name) {
  if (!proofData || typeof proofData.data !== 'string') {
    throw new Error('Data bukti surat tidak valid');
  }

  const match = proofData.data.match(/^data:(application\/pdf|image\/[a-zA-Z0-9.+-]+);base64,([\s\S]+)$/);
  if (!match) throw new Error('Bukti harus berupa file PDF atau gambar');

  const safeName = String(name || 'peserta').replace(/[^\w-]+/g, '-');
  const safeOriginalName = String(proofData.name || 'bukti')
    .replace(/[^\w.-]+/g, '-')
    .replace(/^\.+/, '');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(match[2]),
    match[1],
    `${safeName}-bukti-${new Date().getTime()}-${safeOriginalName}`
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
