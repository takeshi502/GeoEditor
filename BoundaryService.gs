var SMART_LOGGER_BOUNDARY_ENDPOINT = 'https://script.google.com/macros/s/AKfycbxTRBdftgH2xs0aGTSWRMOaL1v2rDKatiqhalIFDbdSWJSVNtdPN13XHHovygjJj1T0VQ/exec';
var GEO_BOUNDARY_CONFIG = {
  municipality: {
    version: 'n03-20260101-47-r1',
    sha256: '9d7a5c93b997f500cf2256c762a80dd825047021dfe3ba65c4d15c8e623c225f'
  },
  base: {
    version: 'okinawa-pref-gis-20210805-r1',
    sha256: 'cae36df8f4d944e5551c565d430ec8819d77aab45a2e7262eaf7cbcc083dd734'
  }
};

function getGeoBoundary_(kind) {
  var config = GEO_BOUNDARY_CONFIG[kind];
  if (!config) throw geoError_('VALIDATION_ERROR', '取得できない境界種別です。');
  var response;
  try {
    response = UrlFetchApp.fetch(SMART_LOGGER_BOUNDARY_ENDPOINT + '?boundary=' + encodeURIComponent(kind), {
      method: 'get', muteHttpExceptions: true, followRedirects: true
    });
  } catch (error) {
    throw geoError_('BOUNDARY_FETCH_FAILED', 'Smart Loggerの正式境界データを取得できませんでした。');
  }
  if (response.getResponseCode() !== 200) {
    throw geoError_('BOUNDARY_FETCH_FAILED', 'Smart Loggerの正式境界データを取得できませんでした (' + response.getResponseCode() + ')。');
  }
  var bytes = response.getBlob().getBytes();
  var actual = digestHex_(bytes);
  if (actual !== config.sha256) throw geoError_('BOUNDARY_INTEGRITY_ERROR', '正式境界データの整合性を確認できませんでした。');
  var geojson;
  try { geojson = JSON.parse(response.getContentText('UTF-8')); }
  catch (error) { throw geoError_('BOUNDARY_FORMAT_ERROR', '正式境界データを読み取れませんでした。'); }
  if (!geojson || geojson.type !== 'FeatureCollection' || !Array.isArray(geojson.features)) {
    throw geoError_('BOUNDARY_FORMAT_ERROR', '正式境界データの形式が不正です。');
  }
  return { kind:kind, version:config.version, sha256:actual, geojson:geojson };
}

function digestHex_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes).map(function (value) {
    var unsigned = value < 0 ? value + 256 : value;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}
