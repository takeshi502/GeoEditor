function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Geo Editor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getGeoEditorBootstrap() {
  return apiCall_(function () { return getGeoEditorBootstrap_(); });
}

function getPlaceDetail(placeId) {
  return apiCall_(function () { return getPlaceDetail_(placeId); });
}

function createPlaceBundle(payload) {
  return apiCall_(function () { return createPlaceBundle_(payload); });
}

function savePlaceBundle(payload) {
  return apiCall_(function () { return savePlaceBundle_(payload); });
}

function mergePlaces(payload) {
  return apiCall_(function () { return mergePlaces_(payload); });
}

function resolveCandidate(payload) {
  return apiCall_(function () { return resolveCandidate_(payload); });
}

function apiCall_(callback) {
  try {
    var result = callback() || {};
    return {
      success: true,
      data: result.data === undefined ? result : result.data,
      warnings: result.warnings || [],
      error_code: null,
      message: result.message || '',
      server_revision: result.server_revision === undefined ? null : result.server_revision
    };
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return {
      success: false,
      data: null,
      warnings: error.warnings || [],
      error_code: error.code || 'INTERNAL_ERROR',
      message: error.message || '予期しないエラーが発生しました。',
      server_revision: error.serverRevision === undefined ? null : error.serverRevision
    };
  }
}

function geoError_(code, message, details) {
  var error = new Error(message);
  error.code = code;
  if (details) Object.keys(details).forEach(function (key) { error[key] = details[key]; });
  return error;
}

function spreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw geoError_('VALIDATION_ERROR', 'Script Properties に SPREADSHEET_ID を設定してください。');
  return SpreadsheetApp.openById(id);
}

function withWriteLock_(callback) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw geoError_('INTERNAL_ERROR', '保存ロックを取得できませんでした。');
  try { return callback(); } finally { lock.releaseLock(); }
}

function jsonValue_(value) {
  return value instanceof Date ? value.toISOString() : value;
}

function readTable_(ss, schemaKey) {
  var contract = GEO_SCHEMA[schemaKey];
  var sheet = ss.getSheetByName(contract.sheet);
  if (!sheet) throw geoError_('SCHEMA_MISMATCH', 'シートが見つかりません: ' + contract.sheet);
  var width = contract.headers.length;
  var actual = sheet.getRange(1, 1, 1, width).getValues()[0];
  validateHeaders_(contract.sheet, actual, contract.headers);
  var lastRow = sheet.getLastRow();
  var rows = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, width).getValues();
  return {
    sheet: sheet,
    headers: contract.headers.slice(),
    rows: rows.filter(function (row) { return row.some(function (cell) { return cell !== ''; }); }).map(function (row, index) {
      var item = { _row: index + 2 };
      contract.headers.forEach(function (header, col) { item[header] = jsonValue_(row[col]); });
      return item;
    })
  };
}

function objectToRow_(object, headers) {
  return headers.map(function (header) {
    var value = object[header];
    return value === undefined || value === null ? '' : value;
  });
}

function publicItem_(item) {
  var copy = {};
  Object.keys(item).forEach(function (key) { if (key !== '_row') copy[key] = item[key]; });
  return copy;
}

function id_(prefix) {
  return prefix + '_' + Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}
