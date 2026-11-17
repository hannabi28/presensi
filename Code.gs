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
  let data;
  try {
    data = JSON.parse(event.postData.contents);
  } catch (error) {
    return json({ ok: false, message: `Data permintaan tidak valid: ${error.message}` });
  }

  if (data.action === 'delete') return deleteRecords(data);

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

function deleteRecords(data) {
  try {
    const configuredToken = PropertiesService.getScriptProperties().getProperty('#Admin_Pr3s3ensi');
    if (!configuredToken) {
      throw new Error('Token admin belum dikonfigurasi. Atur Script Property PRESENSI_DELETE_TOKEN terlebih dahulu.');
    }
    if (!data.adminToken || String(data.adminToken) !== configuredToken) {
      throw new Error('Token admin tidak valid.');
    }

    const ids = Array.isArray(data.ids) ? data.ids.map(String).filter(Boolean) : [];
    if (!ids.length) throw new Error('Pilih setidaknya satu data untuk dihapus.');

    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return json({ ok: true, deleted: 0 });

    const idValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    const requestedIds = new Set(ids);
    const rows = [];
    idValues.forEach((value, index) => {
      if (requestedIds.has(String(value[0]))) rows.push(index + 2);
    });
    rows.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));
    return json({ ok: true, deleted: rows.length });
  } catch (error) {
    return json({ ok: false, message: error.message });
  }
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
