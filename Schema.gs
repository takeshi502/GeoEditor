var GEO_EDITOR_VERSION = '0.1.0';

var GEO_SCHEMA = Object.freeze({
  places: Object.freeze({
    sheet: '06_地点マスタ',
    headers: Object.freeze([
      'place_id', '地点名', 'base_area_id', '優先度', '別名', '登録元',
      '最終使用日時', '有効', '備考', '作成日時', '更新日時', 'revision'
    ])
  }),
  points: Object.freeze({
    sheet: '06_判定ポイントマスタ',
    headers: Object.freeze([
      'point_id', 'place_id', 'ポイント名', '表示順', '中心緯度', '中心経度',
      '判定半径_m', '最大GPS精度_m', '優先度', '最終使用日時', '登録元',
      '有効', '備考', '作成日時', '更新日時', 'revision'
    ])
  }),
  candidates: Object.freeze({
    sheet: '08_地点学習候補',
    headers: Object.freeze([
      'candidate_id', '候補名', '推定所属種別', '推定base_area_id', '推定town_id',
      '基準緯度', '基準経度', 'GPS精度_m', '確認回数', '初回確認日時', '最終確認日時',
      '初回source_type', '初回source_id', '最近source_type', '最近source_id',
      '最寄place_id', '最寄地点距離_m', '候補状態', '登録先place_id', 'revision'
    ])
  }),
  baseAreas: Object.freeze({
    sheet: '07_基地エリアマスタ',
    headers: Object.freeze([
      'base_area_id', '基地名', '判定方式', '中心緯度', '中心経度', '判定半径_m',
      '境界データID', '優先度', '別名', '登録元', '有効', '備考', '作成日時',
      '更新日時', 'revision'
    ])
  }),
  municipalities: Object.freeze({
    sheet: '09_市町村境界マスタ',
    headers: Object.freeze([
      'municipality_id', '行政区域コード', '市町村正式名', '市町村表示名',
      '境界データID', 'bbox_min_lat', 'bbox_min_lng', 'bbox_max_lat',
      'bbox_max_lng', '有効', 'データ基準日', 'revision'
    ])
  })
});

function schemaInfo_() {
  var result = {};
  Object.keys(GEO_SCHEMA).forEach(function (key) {
    result[key] = {
      sheet: GEO_SCHEMA[key].sheet,
      headers: GEO_SCHEMA[key].headers.slice()
    };
  });
  return { app_version: GEO_EDITOR_VERSION, schemas: result };
}
