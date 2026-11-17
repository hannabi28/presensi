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
    if (data.action === 'adminDeleteMany') return adminDeleteMany(data);
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

    if (photo.sharingError) {
      return json({
        ok: true,
        row: rowNumber,
        warning: `Foto dan link sudah disimpan, tetapi izin berbagi Drive gagal: ${photo.sharingError}`
      });
    }

    return json({ ok: true, row: rowNumber, message: 'Absensi berhasil disimpan' });
  } catch (error) {
    if (rowWritten) {
      return json({ ok: false, row: rowNumber, message: `Data absensi tersimpan di spreadsheet (baris ${rowNumber}), tetapi foto gagal disimpan: ${error.message}` });
    }
    return json({ ok: false, message: error.message });
  }
}

function adminDeleteMany(data) {
  const expectedPassword = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expectedPassword) throw new Error('ADMIN_PASSWORD belum dikonfigurasi di Script Properties Apps Script.');
  if (String(data.password || '') !== expectedPassword) throw new Error('Password admin salah.');

  const ids = Array.isArray(data.ids) ? data.ids.map(id => String(id)) : [];
  if (!ids.length || ids.some(id => !id.trim())) throw new Error('Pilih setidaknya satu data absensi yang memiliki ID.');
  if (new Set(ids).size !== ids.length) throw new Error('Daftar data yang dipilih tidak valid.');

  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Tidak ada data absensi untuk dihapus.');

  const requestedIds = new Set(ids);
  const matchingRows = new Map(ids.map(id => [id, []]));
  sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach((row, index) => {
    const id = String(row[0] == null ? '' : row[0]);
    if (requestedIds.has(id)) matchingRows.get(id).push(index + 2);
  });

  const rows = [];
  ids.forEach(id => {
    const matches = matchingRows.get(id);
    if (matches.length !== 1) {
      throw new Error('Data yang dipilih tidak ditemukan atau ID-nya tidak unik. Muat ulang dashboard sebelum menghapus.');
    }
    rows.push(matches[0]);
  });

  rows.sort((a, b) => b - a);
  let groupTop = rows[0];
  let groupBottom = rows[0];
  for (let index = 1; index <= rows.length; index += 1) {
    const row = rows[index];
    if (row !== undefined && row === groupBottom - 1) {
      groupBottom = row;
      continue;
    }
    sheet.deleteRows(groupBottom, groupTop - groupBottom + 1);
    if (row !== undefined) {
      groupTop = row;
      groupBottom = row;
    }
  }
  SpreadsheetApp.flush();
  return json({ ok: true, deleted: ids.length });
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
      'status', 'arrival', 'reason', 'location', 'photo', 'createdAt'
    ]);
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
  let sharingError = '';
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    sharingError = error.message;
  }
  return { url: file.getUrl(), sharingError: sharingError };
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
