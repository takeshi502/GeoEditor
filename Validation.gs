function validateHeaders_(sheetName, actual, expected) {
  if (actual.length !== expected.length || expected.some(function (header, index) { return actual[index] !== header; })) {
    throw geoError_('SCHEMA_MISMATCH', sheetName + ' のヘッダーが正式schemaと一致しません。');
  }
}

function requiredString_(value, label) {
  if (typeof value !== 'string' || !value.trim()) throw geoError_('VALIDATION_ERROR', label + ' は必須です。');
  return value.trim();
}

function finiteNumber_(value, label, min, max) {
  var number = Number(value);
  if (!isFinite(number) || number < min || number > max) {
    throw geoError_('VALIDATION_ERROR', label + ' は ' + min + '〜' + max + ' の数値で指定してください。');
  }
  return number;
}

function integer_(value, label, min, max) {
  var number = finiteNumber_(value, label, min, max);
  if (Math.floor(number) !== number) throw geoError_('VALIDATION_ERROR', label + ' は整数で指定してください。');
  return number;
}

function boolean_(value, label) {
  if (value === true || value === false) return value;
  throw geoError_('VALIDATION_ERROR', label + ' は真偽値で指定してください。');
}

function validatePlace_(place, isNew) {
  if (!isNew) requiredString_(place.place_id, 'place_id');
  requiredString_(place['地点名'], '地点名');
  integer_(place['優先度'] === '' || place['優先度'] == null ? 0 : place['優先度'], '優先度', -100000, 100000);
  boolean_(place['有効'], '有効');
  if (!isNew) integer_(place.revision, 'revision', 1, Number.MAX_SAFE_INTEGER);
}

function validatePoint_(point, expectedPlaceId, isNew) {
  if (!isNew) requiredString_(point.point_id, 'point_id');
  if (point.place_id && expectedPlaceId && point.place_id !== expectedPlaceId) {
    throw geoError_('VALIDATION_ERROR', 'point.place_id が親地点と一致しません。');
  }
  requiredString_(point['ポイント名'], 'ポイント名');
  integer_(point['表示順'], '表示順', 1, 100000);
  finiteNumber_(point['中心緯度'], '中心緯度', -90, 90);
  finiteNumber_(point['中心経度'], '中心経度', -180, 180);
  finiteNumber_(point['判定半径_m'], '判定半径_m', 1, 10000);
  finiteNumber_(point['最大GPS精度_m'], '最大GPS精度_m', 1, 10000);
  integer_(point['優先度'], '優先度', -100000, 100000);
  boolean_(point['有効'], '有効');
  if (!isNew) integer_(point.revision, 'revision', 1, Number.MAX_SAFE_INTEGER);
}

function aliases_(value) {
  if (Array.isArray(value)) return value.map(String).map(function (v) { return v.trim(); }).filter(Boolean);
  if (!value) return [];
  return String(value).split(/[\n,、]/).map(function (v) { return v.trim(); }).filter(Boolean);
}

function mergeAlias_(value, alias) {
  var list = aliases_(value);
  if (alias && list.indexOf(alias) < 0) list.push(alias);
  return list.join(', ');
}

function assertUniqueIds_(items, key) {
  var seen = {};
  items.forEach(function (item) {
    var id = item[key];
    if (id && seen[id]) throw geoError_('DUPLICATE_ID', key + ' が重複しています: ' + id);
    if (id) seen[id] = true;
  });
}
