const SHEET_NAME = 'Historial Pausas Activas';
const SUMMARY_SHEET = 'Resumen Mensual';
const PHOTO_FOLDER_NAME = 'Evidencias Pausas Activas SST';
const FORMAT_SHEET_PREFIX = 'Registro Fotografico ';
const SPREADSHEET_NAME = 'REGISTRO PAUSAS ACTIVAS SST EN LINEA';
const ERROR_SHEET = 'Errores Apps Script';
const APP_VERSION = '2026-09-09-05';
const SHEET_IMAGE_MAX_BYTES = 1900000;
const SHEET_IMAGE_MAX_PIXELS = 1000000;

function doGet(e) {
  try {
    const ss = getSpreadsheet();
    const mode = e && e.parameter ? e.parameter.mode || 'ping' : 'ping';
    if (mode === 'photo') {
      return jsonResponse(getPhotoData(e), e);
    }

    prepareSheet(ss);
    prepareSummarySheet(ss);
    preparePhotoSheet(ss, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'), 'PEAJE FRAGUA');
    if (mode === 'records') {
      return jsonResponse({
        ok: true,
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      version: APP_VERSION,
      records: getRecordsFromSheet(ss)
      }, e);
    }

    return jsonResponse({
      ok: true,
      message: 'API Pausas Activas SST lista',
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      version: APP_VERSION,
      sheets: [SHEET_NAME, SUMMARY_SHEET],
      photoFolder: PHOTO_FOLDER_NAME,
      mode
    }, e);
  } catch (error) {
    try {
      logError(null, 'doGet', error);
    } catch (innerError) {
      console.log(innerError.message);
    }
    return jsonResponse({ ok: false, version: APP_VERSION, error: error.message }, e);
  }
}

function doPost(e) {
  try {
    const rawPayload = e.parameter && e.parameter.payload
      ? e.parameter.payload
      : (e.postData && e.postData.contents ? e.postData.contents : '{}');
    const payload = JSON.parse(rawPayload);
    if (payload.action !== 'saveRecords') {
      throw new Error('Accion no soportada');
    }

    const records = Array.isArray(payload.records) ? payload.records : [];
    if (!records.length) {
      throw new Error('No llegaron registros');
    }

    const ss = getSpreadsheet();
    const sheet = prepareSheet(ss);
    const folder = getPhotoFolder();
    const now = new Date();
    const photoLinksById = {};
    const existingIds = getExistingIds(sheet);
    const newRecords = records.filter((record) => {
      return !record.id || !existingIds[record.id];
    });

    const rows = newRecords.map((record) => {
      const links = savePhotos(folder, record);
      photoLinksById[record.id || ''] = links;
      return [
        now,
        record.date || '',
        record.month || '',
        record.shift || '',
        record.workCenter || '',
        links.length,
        links.join('\n'),
        record.notes || '',
        record.id || ''
      ];
    });

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      try {
        insertRecordsInPhotoSheets(ss, newRecords);
      } catch (imageError) {
        logError(ss, 'insertRecordsInPhotoSheets', imageError);
      }
    }
    updateSummary(ss);

    return jsonResponse({
      ok: true,
      saved: rows.length,
      received: records.length,
      spreadsheetUrl: ss.getUrl(),
      version: APP_VERSION,
      photoLinksById
    });
  } catch (error) {
    try {
      logError(null, 'doPost', error);
    } catch (innerError) {
      console.log(innerError.message);
    }
    return jsonResponse({ ok: false, version: APP_VERSION, error: error.message });
  }
}

function getExistingIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const values = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  return values.reduce((map, row) => {
    if (row[0]) map[row[0]] = true;
    return map;
  }, {});
}

function getSpreadsheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return SpreadsheetApp.openById(file.getId());
    }
  }

  const created = SpreadsheetApp.create(SPREADSHEET_NAME);
  prepareSheet(created);
  preparePhotoSheet(created, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'), 'PEAJE FRAGUA');
  return created;
}

function getRecordsFromSheet(ss) {
  const sheet = prepareSheet(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  return values.map((row) => {
    const links = String(row[6] || '').split(/\n+/).filter(Boolean);
    return {
      id: row[8] || `online-${Utilities.getUuid()}`,
      date: normalizeDate(row[1]),
      month: row[2] || getMonthFromDate(normalizeDate(row[1])),
      shift: row[3] || '',
      workCenter: row[4] || '',
      photos: links.map((url, index) => ({
        name: `Foto ${index + 1}`,
        type: 'image/jpeg',
        dataUrl: '',
        url,
        thumbnailUrl: buildDriveThumbnailUrl(url)
      })),
      notes: row[7] || '',
      syncedAt: row[0] ? new Date(row[0]).toISOString() : ''
    };
  }).filter((record) => record.date);
}

function getPhotoData(e) {
  const params = e && e.parameter ? e.parameter : {};
  const fileId = params.id || extractDriveFileId(params.url || '');
  if (!fileId) {
    throw new Error('No llego el id de la foto');
  }

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const bytes = blob.getBytes();
  const contentType = blob.getContentType() || 'image/jpeg';
  const base64 = Utilities.base64Encode(bytes);
  return {
    ok: true,
    id: fileId,
    name: file.getName(),
    type: contentType,
    size: bytes.length,
    dataUrl: `data:${contentType};base64,${base64}`
  };
}

function extractDriveFileId(url) {
  const text = String(url || '');
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/
  ];
  for (let index = 0; index < patterns.length; index += 1) {
    const match = text.match(patterns[index]);
    if (match) return match[1];
  }
  return '';
}

function buildDriveThumbnailUrl(url) {
  const fileId = extractDriveFileId(url);
  return fileId ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w800` : url;
}

function normalizeDate(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').slice(0, 10);
}

function getOrCreateSheet(ss, name) {
  const existing = ss.getSheetByName(name);
  if (existing) return existing;

  const reusable = ss.getSheets().find((sheet) => {
    const sheetName = sheet.getName();
    const isDefaultName = /^Hoja\s*\d*$|^Sheet\s*\d*$/i.test(sheetName);
    return isDefaultName && sheet.getLastRow() <= 1 && sheet.getLastColumn() <= 1 && !String(sheet.getRange(1, 1).getValue() || '').trim();
  });

  if (reusable) {
    reusable.setName(name);
    return reusable;
  }

  return ss.insertSheet(name);
}

function prepareSheet(ss) {
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  const headers = [
    'Fecha de sincronizacion',
    'Fecha',
    'Mes',
    'Turno',
    'Centro de trabajo',
    'Cantidad de fotos',
    'Enlaces de fotos',
    'Observaciones',
    'Id local'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (sheet.getMaxColumns() > headers.length) {
      sheet.getRange(1, headers.length + 1, 1, sheet.getMaxColumns() - headers.length).clearContent();
    }
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setFontColor('#000000');
  sheet.setFrozenRows(1);

  return sheet;
}

function prepareSummarySheet(ss) {
  const summary = getOrCreateSheet(ss, SUMMARY_SHEET);
  if (summary.getLastRow() === 0) {
    summary.appendRow(['Mes', 'Turno', 'Registros', 'Fotos']);
  }
  summary.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setFontColor('#000000');
  return summary;
}

function logError(ss, source, error) {
  const target = ss || getSpreadsheet();
  const sheet = target.getSheetByName(ERROR_SHEET) || target.insertSheet(ERROR_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Fecha', 'Origen', 'Mensaje', 'Detalle']);
    sheet.getRange(1, 1, 1, 4)
      .setFontWeight('bold')
      .setBackground('#c82135')
      .setFontColor('#ffffff');
  }
  sheet.appendRow([
    new Date(),
    source,
    error && error.message ? error.message : String(error),
    error && error.stack ? error.stack : ''
  ]);
}

function savePhotos(folder, record) {
  const photos = Array.isArray(record.photos) ? record.photos : [];
  return photos.map((photo, index) => {
    if (photo.url && !photo.dataUrl) return photo.url;
    if (!photo.dataUrl) return '';

    const data = String(photo.dataUrl).split(',');
    const meta = data[0] || '';
    const content = data[1] || '';
    const mimeMatch = meta.match(/data:(.*);base64/);
    const mimeType = mimeMatch ? mimeMatch[1] : (photo.type || 'image/jpeg');
    const bytes = Utilities.base64Decode(content);
    const name = photo.name || buildServerPhotoName(record, index + 1, mimeType);
    const blob = Utilities.newBlob(bytes, mimeType, name);
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (error) {
      console.log(`No se pudo cambiar permisos de ${name}: ${error.message}`);
    }
    return file.getUrl();
  }).filter(Boolean);
}

function getPhotoFolder() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function insertRecordsInPhotoSheets(ss, records) {
  const grouped = records.reduce((map, record) => {
    const month = record.month || getMonthFromDate(record.date);
    if (!map[month]) map[month] = [];
    map[month].push(record);
    return map;
  }, {});

  Object.keys(grouped).forEach((month) => {
    const first = grouped[month][0] || {};
    const sheet = preparePhotoSheet(ss, month, first.workCenter || 'PEAJE FRAGUA');
    setPhotoRowHeights(sheet, grouped[month]);
    grouped[month].forEach((record) => insertRecordImages(sheet, record));
  });
}

function setPhotoRowHeights(sheet, records) {
  const byDay = {};
  records.forEach((record) => {
    const day = Number(String(record.date || '').split('-')[2]);
    if (!day) return;
    const key = `${day}|${record.shift}`;
    byDay[key] = (byDay[key] || 0) + (Array.isArray(record.photos) ? record.photos.length : 0);
  });

  for (let day = 1; day <= 31; day += 1) {
    const maxPhotos = Math.max(
      byDay[`${day}|A`] || 0,
      byDay[`${day}|B`] || 0,
      byDay[`${day}|C`] || 0,
      1
    );
    sheet.setRowHeight(13 + day, Math.max(190, maxPhotos * 170));
  }
}

function preparePhotoSheet(ss, month, workCenter) {
  const sheetName = `${FORMAT_SHEET_PREFIX}${month || 'Sin Mes'}`.slice(0, 99);
  const existingSheet = ss.getSheetByName(sheetName);
  const sheet = existingSheet || ss.insertSheet(sheetName);
  const days = getDaysInMonth(month);

  if (!existingSheet) {
    sheet.clear();
  }
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 95);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 260);

  sheet.getRange('A2:D12').breakApart();
  sheet.getRange('A2:A4').merge();
  sheet.getRange('B2:D4').merge();
  sheet.getRange('B5:D6').merge();
  sheet.getRange('A7:D9').merge();
  sheet.getRange('A10:D12').merge();

  sheet.getRange('B2').setValue('FORMATO REGISTRO FOTOGRAFICO DE PAUSAS ACTIVAS\nZIMA SEGURIDAD LTDA');
  sheet.getRange('A5').setValue('ESTACION DE RECAUDO');
  sheet.getRange('B5').setValue(`CENTRO DE TRABAJO: ${workCenter}`);
  sheet.getRange('A7').setValue(`MES: ${getMonthLabel(month)}`);
  sheet.getRange('A10').setValue('ANEXO FOTOGRAFICO');
  sheet.getRange('A13:D13').setValues([['FECHA', 'TURNO A', 'TURNO B', 'TURNO C']]);

  sheet.getRange('B2:D4')
    .setBackground('#d9d9d9')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.getRange('A5:D9').setFontWeight('bold').setVerticalAlignment('middle').setWrap(true);
  sheet.getRange('A10:D12')
    .setBackground('#d9d9d9')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('A13:D13')
    .setBackground('#ededed')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeights(2, 3, 24);
  sheet.setRowHeights(7, 6, 22);
  sheet.setRowHeight(13, 38);

  for (let day = 1; day <= days; day += 1) {
    const row = 13 + day;
    const date = `${month}-${String(day).padStart(2, '0')}`;
    sheet.getRange(row, 1).setValue(formatDateForSheet(date));
    sheet.setRowHeight(row, 190);
  }

  const lastRow = 13 + days;
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 4)
    .setBorder(true, true, true, true, true, true, '#1c2630', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');
  sheet.getRange(14, 1, days, 4).setHorizontalAlignment('center').setWrap(true);
  sheet.setFrozenRows(13);

  return sheet;
}

function insertRecordImages(sheet, record) {
  const day = Number(String(record.date || '').split('-')[2]);
  if (!day) return;
  const row = 13 + day;
  const column = getShiftColumn(record.shift);
  if (!column) return;

  const photos = Array.isArray(record.photos) ? record.photos : [];
  photos.forEach((photo, index) => {
    if (!photo.dataUrl) return;
    try {
      validateSheetImage(photo);
      const blob = dataUrlToBlob(photo.dataUrl, photo.name || buildServerPhotoName(record, index + 1, 'image/jpeg'), photo.type || 'image/jpeg');
      const image = sheet.insertImage(blob, column, row);
      const size = fitImage(Number(photo.width || 900), Number(photo.height || 700), 230, 155);
      image.setWidth(size.width);
      image.setHeight(size.height);
      image.setAnchorCell(sheet.getRange(row, column));
      image.setAnchorCellXOffset(14);
      image.setAnchorCellYOffset(12 + index * 165);
    } catch (error) {
      const message = `Foto ${index + 1} no insertada: ${error.message}`;
      appendCellMessage(sheet.getRange(row, column), message);
      logError(sheet.getParent(), 'insertRecordImages', new Error(`${record.date} turno ${record.shift}: ${message}`));
    }
  });

  if (record.notes) {
    sheet.getRange(row, column).setValue(record.notes);
  }
}

function validateSheetImage(photo) {
  const width = Number(photo.width || 0);
  const height = Number(photo.height || 0);
  const pixels = width * height;
  const bytes = Number(photo.size || estimateDataUrlBytes(photo.dataUrl));
  if (pixels > SHEET_IMAGE_MAX_PIXELS) {
    throw new Error(`supera 1 millon de pixeles (${pixels})`);
  }
  if (bytes > SHEET_IMAGE_MAX_BYTES) {
    throw new Error(`supera 2MB (${bytes} bytes)`);
  }
}

function estimateDataUrlBytes(dataUrl) {
  const content = String(dataUrl || '').split(',')[1] || '';
  return Math.ceil((content.length * 3) / 4);
}

function appendCellMessage(range, message) {
  const current = String(range.getValue() || '').trim();
  range.setValue(current ? `${current}\n${message}` : message);
}

function dataUrlToBlob(dataUrl, name, mimeType) {
  const data = String(dataUrl).split(',');
  const meta = data[0] || '';
  const content = data[1] || '';
  const mimeMatch = meta.match(/data:(.*);base64/);
  const type = mimeMatch ? mimeMatch[1] : mimeType;
  return Utilities.newBlob(Utilities.base64Decode(content), type, name);
}

function getShiftColumn(shift) {
  if (shift === 'A') return 2;
  if (shift === 'B') return 3;
  if (shift === 'C') return 4;
  return 0;
}

function getDaysInMonth(month) {
  if (!month) return 31;
  const parts = String(month).split('-').map(Number);
  return new Date(parts[0], parts[1], 0).getDate();
}

function getMonthFromDate(date) {
  return String(date || '').slice(0, 7);
}

function getMonthLabel(month) {
  if (!month) return '';
  const parts = String(month).split('-').map(Number);
  const names = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
  return `${names[parts[1] - 1]} DE ${parts[0]}`;
}

function formatDateForSheet(date) {
  const parts = String(date).split('-');
  return `${Number(parts[2])}/${Number(parts[1])}/${parts[0]}`;
}

function fitImage(width, height, maxWidth, maxHeight) {
  const ratio = Math.min(maxWidth / width, maxHeight / height, 1);
  return {
    width: Math.round(width * ratio),
    height: Math.round(height * ratio)
  };
}

function buildServerPhotoName(record, index, mimeType) {
  const extension = mimeType.split('/')[1] || 'jpg';
  const date = record.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const shift = record.shift || 'X';
  return `pausa-activa_${date}_turno-${shift}_${String(index).padStart(2, '0')}.${extension}`;
}

function updateSummary(ss) {
  const source = prepareSheet(ss);
  const summary = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  const values = source.getDataRange().getValues().slice(1);
  const map = {};

  values.forEach((row) => {
    const month = row[2];
    const shift = row[3];
    if (!month) return;
    const key = `${month}|${shift}`;
    if (!map[key]) {
      map[key] = { month, shift, records: 0, photos: 0 };
    }
    map[key].records += 1;
    map[key].photos += Number(row[5] || 0);
  });

  const rows = Object.values(map)
    .sort((a, b) => String(b.month).localeCompare(String(a.month)) || String(a.shift).localeCompare(String(b.shift)))
    .map((item) => [item.month, item.shift, item.records, item.photos]);

  summary.clear();
  summary.appendRow(['Mes', 'Turno', 'Registros', 'Fotos']);
  if (rows.length) {
    summary.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  summary.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setFontColor('#000000');
  summary.autoResizeColumns(1, 4);
}

function safe(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function jsonResponse(data, e) {
  const callback = e && e.parameter ? e.parameter.callback : '';
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(data)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
