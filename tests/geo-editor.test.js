'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const os = require('node:os');
const crypto = require('node:crypto');

const root = path.resolve(__dirname, '..');
let uuidCounter = 0;
const context = vm.createContext({
  console, Date, Math, Number, Object, String, Array, Boolean, JSON, Error, RegExp,
  isFinite, parseInt, parseFloat,
  Utilities: { getUuid: () => `uuid-${++uuidCounter}` }
});
for (const name of ['Schema.gs','Code.gs','Validation.gs','GeoMath.gs','PlaceService.gs','PlaceMergeService.gs','CandidateService.gs']) {
  vm.runInContext(fs.readFileSync(path.join(root, name), 'utf8'), context, { filename:name });
}
const spatialContext = vm.createContext({ window:{}, console });
const boundarySource = fs.readFileSync(path.join(root, 'BoundaryGeometry.html'), 'utf8').match(/^<script>\s*([\s\S]*?)\s*<\/script>\s*$/)[1];
vm.runInContext(boundarySource, spatialContext, { filename:'BoundaryGeometry.html' });
const spatialSource = fs.readFileSync(path.join(root, 'SpatialOrder.html'), 'utf8').match(/^<script>\s*([\s\S]*?)\s*<\/script>\s*$/)[1];
vm.runInContext(spatialSource, spatialContext, { filename:'SpatialOrder.html' });
const spatial = spatialContext.window.GeoPlaceOrder;
const boundary = spatialContext.window.GeoBoundary;

function boundaryIndex(features, idProperty='municipality_id') {
  return boundary.buildIndex({type:'FeatureCollection',features}, {idProperty,cellSize:.05});
}
function rectangleFeature(id, minLat, minLng, maxLat, maxLng, idProperty='municipality_id') {
  return {type:'Feature',properties:{[idProperty]:id},geometry:{type:'Polygon',coordinates:[[[minLng,minLat],[maxLng,minLat],[maxLng,maxLat],[minLng,maxLat],[minLng,minLat]]]}};
}

function point(overrides = {}) {
  return Object.assign({ point_id:'p1', place_id:'a', 'ポイント名':'main', '表示順':1, '中心緯度':26, '中心経度':127, '判定半径_m':20, '最大GPS精度_m':20, '優先度':100, '有効':true, revision:1 }, overrides);
}
function place(overrides = {}) {
  return Object.assign({ place_id:'a', '地点名':'A', base_area_id:'', '優先度':0, '別名':'', '登録元':'test', '最終使用日時':'', '有効':true, '備考':'', '作成日時':'x', '更新日時':'x', revision:1 }, overrides);
}
function baseArea(overrides = {}) {
  return Object.assign({ base_area_id:'base-1','基地名':'フォスター','判定方式':'circle','中心緯度':26,'中心経度':127,'判定半径_m':1000,'境界データID':'','優先度':0,'別名':'','登録元':'test','有効':true,'備考':'','作成日時':'x','更新日時':'x',revision:1 }, overrides);
}
function municipality(overrides = {}) {
  return Object.assign({ municipality_id:'municipality-1','行政区域コード':'47205','市町村正式名':'宜野湾市','市町村表示名':'宜野湾市','境界データID':'boundary-1',bbox_min_lat:26.24,bbox_min_lng:127.72,bbox_max_lat:26.31,bbox_max_lng:127.80,'有効':true,'データ基準日':'2026-01-01',revision:1 }, overrides);
}
function fakeSheet(headers, rows) {
  const sheet = {
    headers, rows: rows.map(r => Object.assign({}, r)), writes:[], appended:[],
    getLastRow() { return this.rows.length + 1; },
    getRange(row, col, height, width) {
      return {
        setValues: values => {
          if (Number.isInteger(sheet.failAfter)) {
            if (sheet.failAfter === 0) { sheet.failAfter = null; throw new Error('simulated write failure'); }
            sheet.failAfter -= 1;
          }
          values.forEach((valuesRow, offset) => {
            const object = {}; headers.forEach((h,i) => { object[h] = valuesRow[i]; });
            const index = row - 2 + offset;
            if (index < sheet.rows.length) sheet.rows[index] = object; else sheet.rows.push(object);
            sheet.writes.push({ row:row+offset, object });
          });
          return this;
        },
        clearContent: () => { sheet.writes.push({ row, cleared:true }); }
      };
    },
    appendRow(values) { const object={}; headers.forEach((h,i)=>{object[h]=values[i];}); this.rows.push(object); this.appended.push(object); }
  };
  return sheet;
}
function table(key, rows) {
  const headers = Array.from(context.GEO_SCHEMA[key].headers);
  const withRows = rows.map((r,i) => Object.assign({_row:i+2}, r));
  return { headers, rows:withRows, sheet:fakeSheet(headers, withRows) };
}

test('正式schemaの列順と列数を検証する', () => {
  assert.equal(context.GEO_SCHEMA.places.headers.length, 12);
  assert.equal(context.GEO_SCHEMA.points.headers.length, 16);
  assert.equal(context.GEO_SCHEMA.candidates.headers.length, 20);
  assert.equal(context.GEO_SCHEMA.municipalities.headers.length, 12);
  assert.doesNotThrow(() => context.validateHeaders_('x', Array.from(context.GEO_SCHEMA.places.headers), context.GEO_SCHEMA.places.headers));
  assert.throws(() => context.validateHeaders_('x', ['wrong'], context.GEO_SCHEMA.places.headers), e => e.code === 'SCHEMA_MISMATCH');
});

test('距離計算は既知の緯度差をメートルで返す', () => {
  assert.ok(Math.abs(context.distanceMeters_(0, 0, 0.001, 0) - 111.195) < 0.2);
});

test('登録地点はbase_area_idごとにまとまり、同一データなら安定した空間順になる', () => {
  const places=[
    place({place_id:'a','地点名':'A',base_area_id:'base-1'}),
    place({place_id:'b','地点名':'B',base_area_id:''}),
    place({place_id:'c','地点名':'C',base_area_id:'base-1'}),
    place({place_id:'d','地点名':'D',base_area_id:''})
  ];
  const points=[
    point({point_id:'pa',place_id:'a','中心緯度':26.30,'中心経度':127.76}),
    point({point_id:'pb',place_id:'b','中心緯度':26.31,'中心経度':127.77}),
    point({point_id:'pc',place_id:'c','中心緯度':26.32,'中心経度':127.78}),
    point({point_id:'pd',place_id:'d','中心緯度':27.20,'中心経度':128.10})
  ];
  const bases=[{base_area_id:'base-1','基地名':'フォスター'}];
  const first=Array.from(spatial.orderPlaces(places,points,bases), record => record.place.place_id);
  const second=Array.from(spatial.orderPlaces(places.slice().reverse(),points.slice().reverse(),bases), record => record.place.place_id);
  assert.deepEqual(first,second);
  assert.equal(Math.abs(first.indexOf('a')-first.indexOf('c')),1);
  assert.equal(spatial.orderPlaces(places,points,bases).filter(record => record.groupLabel === 'フォスター').length,2);
});

test('近い地点を連続させ、複数point地点は表示順先頭の有効pointで1件として扱う', () => {
  const places=[place({place_id:'near-1'}),place({place_id:'near-2'}),place({place_id:'far'})];
  const points=[
    point({point_id:'n1-secondary',place_id:'near-1','表示順':2,'中心緯度':27.20,'中心経度':128.10}),
    point({point_id:'n1-main',place_id:'near-1','表示順':1,'中心緯度':26.3000,'中心経度':127.7600}),
    point({point_id:'n2',place_id:'near-2','表示順':1,'中心緯度':26.3002,'中心経度':127.7602}),
    point({point_id:'far',place_id:'far','表示順':1,'中心緯度':27.2000,'中心経度':128.1000})
  ];
  assert.equal(spatial.representativePoint('near-1',points).point_id,'n1-main');
  const ordered=spatial.orderPlaces(places,points,[]);
  const ids=Array.from(ordered, record => record.place.place_id);
  assert.equal(ordered.filter(record => record.place.place_id === 'near-1').length,1);
  assert.equal(Math.abs(ids.indexOf('near-1')-ids.indexOf('near-2')),1);
});

test('基地外地点は代表pointの正式polygonでグループ化し、基地グループを先に並べる', () => {
  const places=[
    place({place_id:'base','地点名':'基地内',base_area_id:'base-1'}),
    place({place_id:'ginowan','地点名':'宜野湾',base_area_id:''}),
    place({place_id:'chatan','地点名':'北谷',base_area_id:''})
  ];
  const points=[
    point({point_id:'base-point',place_id:'base','中心緯度':26.30,'中心経度':127.76}),
    point({point_id:'ginowan-point',place_id:'ginowan','中心緯度':26.27,'中心経度':127.75}),
    point({point_id:'chatan-point',place_id:'chatan','中心緯度':26.32,'中心経度':127.76})
  ];
  const municipalities=[
    municipality(),
    municipality({municipality_id:'municipality-2','行政区域コード':'47326','市町村正式名':'北谷町','市町村表示名':'北谷町',bbox_min_lat:26.293,bbox_min_lng:127.742,bbox_max_lat:26.347,bbox_max_lng:127.786})
  ];
  const index=boundaryIndex([
    rectangleFeature('municipality-1',26.24,127.72,26.31,127.80),
    rectangleFeature('municipality-2',26.293,127.742,26.347,127.786)
  ]);
  const ordered=spatial.orderPlaces(places,points,[baseArea()],municipalities,index);
  assert.deepEqual(Array.from(ordered, record => record.groupLabel),['フォスター','宜野湾市','北谷町']);
  assert.equal(ordered[0].groupType,'base');
});

test('bbox候補が重なってもpolygon内側だけを選び、未判定地点も一覧に残す', () => {
  const broad=municipality({municipality_id:'broad','市町村表示名':'広域',bbox_min_lat:26.20,bbox_min_lng:127.70,bbox_max_lat:26.40,bbox_max_lng:127.90});
  const narrow=municipality({municipality_id:'narrow','市町村表示名':'狭域',bbox_min_lat:26.25,bbox_min_lng:127.74,bbox_max_lat:26.30,bbox_max_lng:127.79});
  const located=point({point_id:'located',place_id:'located','中心緯度':26.27,'中心経度':127.76});
  const index=boundaryIndex([
    rectangleFeature('broad',26.20,127.70,26.40,127.74),
    rectangleFeature('narrow',26.25,127.74,26.30,127.79)
  ]);
  assert.equal(spatial.municipalityForPoint(located,[broad,narrow],index).municipality_id,'narrow');
  const ordered=spatial.orderPlaces(
    [place({place_id:'located'}),place({place_id:'unknown'})],
    [located,point({point_id:'unknown',place_id:'unknown','中心緯度':30,'中心経度':130})],
    [],[broad,narrow],index
  );
  assert.deepEqual(Array.from(ordered, record => record.groupLabel),['狭域','市町村未判定']);
});

test('基地所属地点は市町村polygon内でも基地グループだけに1件表示する', () => {
  const places=[place({place_id:'inside-base',base_area_id:'base-1'})];
  const points=[point({point_id:'inside-base-point',place_id:'inside-base','中心緯度':26.27,'中心経度':127.76})];
  const index=boundaryIndex([rectangleFeature('municipality-1',26.24,127.72,26.31,127.80)]);
  const ordered=spatial.orderPlaces(places,points,[baseArea()],[municipality()],index);
  assert.equal(ordered.length,1);
  assert.equal(ordered[0].groupLabel,'フォスター');
  assert.equal(ordered[0].groupKey,'base:base-1');
});

test('Polygon/MultiPolygonは穴を外側とし境界線上を内側としてSmart Loggerと同じ判定をする', () => {
  const index=boundaryIndex([{type:'Feature',properties:{municipality_id:'m'},geometry:{type:'MultiPolygon',coordinates:[
    [[[127,26],[128,26],[128,27],[127,27],[127,26]],[[127.4,26.4],[127.6,26.4],[127.6,26.6],[127.4,26.6],[127.4,26.4]]],
    [[[129,26],[130,26],[130,27],[129,27],[129,26]]]
  ]}}]);
  assert.equal(index.find(26.2,127.2),'m');
  assert.equal(index.find(26.5,127.5),null);
  assert.equal(index.find(26.5,129.5),'m');
  assert.equal(index.find(26,127.5),'m');
});

test('正式polygonが未読込ならbboxだけで市町村を推測しない', () => {
  const p=point({'中心緯度':26.27,'中心経度':127.76});
  assert.equal(spatial.municipalityForPoint(p,[municipality()],null),null);
});

const officialMunicipalityAsset=path.join(os.homedir(),'.codex','.chatgpt-projects','g-p-6a86017f22e08191904accfa8aa448d3','SmartLoggerWeb','public','admin-boundaries','2026-01-01-r1','municipalities-okinawa-prefecture-2026-01-01-r1.geojson');
test('Smart Logger正式GeoJSONで第2ゲートは那覇市だけに含まれる', {skip:!fs.existsSync(officialMunicipalityAsset)}, () => {
  const collection=JSON.parse(fs.readFileSync(officialMunicipalityAsset,'utf8'));
  const index=boundary.buildIndex(collection,{idProperty:'municipality_id',cellSize:.05});
  assert.deepEqual(Array.from(index.matches(26.19565886378853,127.6590058207512)),['municipality_47201']);
});

test('判定円重複量を計算する', () => {
  const a=point({'中心緯度':0,'中心経度':0,'判定半径_m':70});
  const b=point({point_id:'p2',place_id:'b','中心緯度':0.001,'中心経度':0,'判定半径_m':60});
  assert.ok(Math.abs(context.overlapMeters_(a,b) - 18.805) < 0.3);
  assert.equal(context.circlesOverlap_(a,b), true);
});

test('同一place_idは重複警告せず、異なるplace_idは警告する', () => {
  const a=point({'中心緯度':0,'中心経度':0,'判定半径_m':100});
  const same=point({point_id:'p2','中心緯度':0.0001,'中心経度':0,'判定半径_m':100});
  const other=point({point_id:'p3',place_id:'b','中心緯度':0.0001,'中心経度':0,'判定半径_m':100});
  assert.equal(context.overlapWarnings_([a],[same]).length,0);
  assert.equal(context.overlapWarnings_([a],[other]).length,1);
});

test('推奨半径は安全余白5mを引き5m単位で安全側へ丸め、GPS精度も提案する', () => {
  const a=point({'中心緯度':0,'中心経度':0,'判定半径_m':50});
  const b=point({point_id:'p2',place_id:'b','中心緯度':0.0004,'中心経度':0,'判定半径_m':11});
  const suggestion=context.radiusSuggestion_(a,[b]);
  assert.equal(suggestion.radius_m,25);
  assert.equal(suggestion.max_accuracy_m,25);
});

test('revision競合を停止する', () => {
  assert.doesNotThrow(() => context.assertRevision_(2,2,'地点',2));
  assert.throws(() => context.assertRevision_(1,2,'地点',2), e => e.code === 'REVISION_CONFLICT' && e.serverRevision === 2);
});

test('地点名だけの保存では未変更ポイントを書き換えずrevisionを維持する', () => {
  const places=table('places',[place()]), points=table('points',[point()]);
  context.withWriteLock_=fn=>fn();
  context.spreadsheet_=()=>({});
  context.readTable_=(_ss,key)=>key === 'places' ? places : points;
  const result=context.savePlaceBundle_({
    place:place({'地点名':'A renamed'}),
    points:[point()]
  });
  assert.equal(result.data.points[0].revision,1);
  assert.equal(points.sheet.writes.length,0);
  assert.equal(places.sheet.writes.length,1);
});

test('数値由来のpoint名を保持して地点名だけ保存する', () => {
  const currentPlace=place({'地点名':5679,revision:7,'更新日時':'place-old'});
  const currentPoint=point({'ポイント名':5679,revision:5,'更新日時':'point-old'});
  const places=table('places',[currentPlace]), points=table('points',[currentPoint]);
  context.withWriteLock_=fn=>fn(); context.spreadsheet_=()=>({});
  context.readTable_=(_ss,key)=>key === 'places' ? places : points;
  const result=context.savePlaceBundle_({
    place:Object.assign({},currentPlace,{'地点名':'メイン'}),
    points:[Object.assign({},currentPoint,{'ポイント名':'5679'})]
  });
  assert.equal(result.data.place['地点名'],'メイン');
  assert.equal(result.data.place.revision,8);
  assert.notEqual(result.data.place['更新日時'],'place-old');
  assert.equal(String(result.data.points[0]['ポイント名']),'5679');
  assert.equal(result.data.points[0].revision,5);
  assert.equal(result.data.points[0]['更新日時'],'point-old');
  assert.equal(places.sheet.writes.length,1);
  assert.equal(points.sheet.writes.length,0);
});

test('地点名とpoint名を同時変更すると両方のrevisionを更新する', () => {
  const currentPlace=place({'地点名':'5679',revision:2,'更新日時':'place-old'});
  const currentPoint=point({'ポイント名':'5679',revision:4,'更新日時':'point-old'});
  const places=table('places',[currentPlace]), points=table('points',[currentPoint]);
  context.withWriteLock_=fn=>fn(); context.spreadsheet_=()=>({});
  context.readTable_=(_ss,key)=>key === 'places' ? places : points;
  const result=context.savePlaceBundle_({
    place:Object.assign({},currentPlace,{'地点名':'メイン'}),
    points:[Object.assign({},currentPoint,{'ポイント名':'正面入口'})]
  });
  assert.equal(result.data.place.revision,3);
  assert.equal(result.data.points[0].revision,5);
  assert.equal(places.sheet.writes.length,1);
  assert.equal(points.sheet.writes.length,1);
});

test('新規pointのポイント名が空ならvalidationを維持する', () => {
  const currentPlace=place(), currentPoint=point();
  const places=table('places',[currentPlace]), points=table('points',[currentPoint]);
  context.withWriteLock_=fn=>fn(); context.spreadsheet_=()=>({});
  context.readTable_=(_ss,key)=>key === 'places' ? places : points;
  assert.throws(() => context.savePlaceBundle_({
    place:currentPlace,
    points:[currentPoint,point({point_id:undefined,place_id:'a','ポイント名':'','表示順':2,revision:undefined})]
  }), error => error.code === 'VALIDATION_ERROR' && /ポイント名/.test(error.message));
});

test('数値由来の地点名を保持してpoint名だけ保存し、親placeを更新しない', () => {
  const currentPlace=place({'地点名':5665,revision:3,'更新日時':'place-old'});
  const currentPoint=point({'ポイント名':'メイン',revision:2,'更新日時':'point-old'});
  const places=table('places',[currentPlace]), points=table('points',[currentPoint]);
  context.withWriteLock_=fn=>fn();
  context.spreadsheet_=()=>({});
  context.readTable_=(_ss,key)=>key === 'places' ? places : points;
  const result=context.savePlaceBundle_({
    place:Object.assign({},currentPlace,{'地点名':'5665'}),
    points:[Object.assign({},currentPoint,{'ポイント名':'5665'})]
  });
  assert.equal(result.data.place['地点名'],5665);
  assert.equal(result.data.place.revision,3);
  assert.equal(result.data.place['更新日時'],'place-old');
  assert.equal(places.sheet.writes.length,0);
  assert.equal(result.data.points[0]['ポイント名'],'5665');
  assert.equal(result.data.points[0].revision,3);
  assert.notEqual(result.data.points[0]['更新日時'],'point-old');
  assert.equal(points.sheet.writes.length,1);
});

test('空の地点名による新規地点作成は引き続き拒否する', () => {
  const places=table('places',[]), points=table('points',[]);
  assert.throws(() => context.createPlaceBundleWithTables_({
    place:place({place_id:undefined,revision:undefined,'地点名':''}),
    points:[point({point_id:undefined,place_id:undefined,revision:undefined})]
  },places,points), error => error.code === 'VALIDATION_ERROR' && /地点名/.test(error.message));
});

test('保存payloadは画面表示可能な数値由来地点名を文字列として送る', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/function buildSaveBundle\(\)/);
  assert.match(client,/const place = clone\(state\.draftPlace\)/);
  assert.match(client,/place\['地点名'\] = String\(place\['地点名'\] == null \? '' : place\['地点名'\]\)/);
  assert.match(client,/copy\['ポイント名'\] = String\(copy\['ポイント名'\] == null \? '' : copy\['ポイント名'\]\)/);
  assert.match(client,/const \{ place, points \} = buildSaveBundle\(\)/);
  assert.match(client,/server\('savePlaceBundle', \{ place, points, preserve_old_name:preserve \}\)/);
});

test('新規地点と初期pointを同時に生成する', () => {
  const places=table('places',[]), points=table('points',[]);
  const result=context.createPlaceBundleWithTables_({ place:place({place_id:undefined,revision:undefined}), points:[point({point_id:undefined,place_id:undefined,revision:undefined})] },places,points);
  assert.match(result.data.place.place_id,/^place_/);
  assert.equal(result.data.points.length,1);
  assert.equal(result.data.points[0].place_id,result.data.place.place_id);
  assert.equal(result.data.place.revision,1);
});

function candidateHarness(status='pending', candidateOverrides={}, baseAreas=[]) {
  const candidate=Object.assign({ candidate_id:'c1','候補名':'候補','推定所属種別':'unknown','推定base_area_id':'','推定town_id':'','基準緯度':26,'基準経度':127,'GPS精度_m':10,'確認回数':2,'初回確認日時':'','最終確認日時':'','初回source_type':'','初回source_id':'','最近source_type':'','最近source_id':'','最寄place_id':'','最寄地点距離_m':'','候補状態':status,'登録先place_id':'',revision:1 },candidateOverrides);
  const tables={ candidates:table('candidates',[candidate]), places:table('places',[place()]), points:table('points',[point()]), baseAreas:table('baseAreas',baseAreas) };
  context.withWriteLock_=fn=>fn(); context.spreadsheet_=()=>({}); context.readTable_=(_ss,key)=>tables[key];
  return tables;
}

test('candidate promote', () => {
  const t=candidateHarness();
  const res=context.resolveCandidate_({action:'promote',candidate_id:'c1',revision:1,bundle:{place:place({place_id:undefined,revision:undefined,'地点名':'候補'}),points:[point({point_id:undefined,place_id:undefined,revision:undefined})]}});
  assert.match(res.data.place.place_id,/^place_/);
  assert.equal(t.candidates.sheet.rows[0]['候補状態'],'promoted');
});

test('candidate merge', () => {
  const t=candidateHarness();
  context.resolveCandidate_({action:'merge',candidate_id:'c1',revision:1,place_id:'a',place_revision:1,add_candidate_alias:true});
  assert.equal(t.candidates.sheet.rows[0]['候補状態'],'merged');
  assert.match(t.places.sheet.rows[0]['別名'],/候補/);
});

test('candidate add_point', () => {
  const t=candidateHarness();
  context.resolveCandidate_({action:'add_point',candidate_id:'c1',revision:1,place_id:'a',place_revision:1,point:point({point_id:undefined,place_id:undefined,revision:undefined,'ポイント名':'候補入口'})});
  assert.equal(t.points.sheet.appended.length,1);
  assert.equal(t.candidates.sheet.rows[0]['候補状態'],'merged');
});

test('candidate reject', () => {
  const t=candidateHarness();
  context.resolveCandidate_({action:'reject',candidate_id:'c1',revision:1});
  assert.equal(t.candidates.sheet.rows[0]['候補状態'],'rejected');
});

test('candidate promoteは調整した登録予定位置をpointへ使い候補元座標を保持する', () => {
  const t=candidateHarness();
  const res=context.resolveCandidate_({
    action:'promote', candidate_id:'c1', revision:1,
    registration_position:{latitude:26.1234,longitude:127.5678},
    bundle:{place:place({place_id:undefined,revision:undefined,'地点名':'調整済み候補'}),points:[point({point_id:undefined,place_id:undefined,revision:undefined})]}
  });
  assert.equal(res.data.points[0]['中心緯度'],26.1234);
  assert.equal(res.data.points[0]['中心経度'],127.5678);
  assert.equal(t.candidates.sheet.rows[0]['基準緯度'],26);
  assert.equal(t.candidates.sheet.rows[0]['基準経度'],127);
});

test('candidate add_pointは調整した登録予定位置だけを新pointへ反映する', () => {
  const t=candidateHarness();
  const res=context.resolveCandidate_({
    action:'add_point', candidate_id:'c1', revision:1, place_id:'a', place_revision:1,
    registration_position:{latitude:26.2222,longitude:127.3333},
    point:point({point_id:undefined,place_id:undefined,revision:undefined,'ポイント名':'調整入口'})
  });
  assert.equal(res.data.point['中心緯度'],26.2222);
  assert.equal(res.data.point['中心経度'],127.3333);
  assert.equal(t.candidates.sheet.rows[0]['基準緯度'],26);
  assert.equal(t.candidates.sheet.rows[0]['基準経度'],127);
});

test('candidate mergeとrejectはregistration_positionを解釈しない', () => {
  let t=candidateHarness();
  assert.doesNotThrow(() => context.resolveCandidate_({action:'merge',candidate_id:'c1',revision:1,place_id:'a',place_revision:1,registration_position:{latitude:999,longitude:999}}));
  assert.equal(t.points.sheet.appended.length,0);
  t=candidateHarness();
  assert.doesNotThrow(() => context.resolveCandidate_({action:'reject',candidate_id:'c1',revision:1,registration_position:{latitude:999,longitude:999}}));
  assert.equal(t.points.sheet.appended.length,0);
});

test('candidate登録予定位置はサーバーで緯度経度範囲を検証する', () => {
  const t=candidateHarness();
  assert.throws(() => context.resolveCandidate_({
    action:'add_point', candidate_id:'c1', revision:1, place_id:'a', place_revision:1,
    registration_position:{latitude:91,longitude:127}, point:point({'ポイント名':'不正位置'})
  }), error => error.code === 'VALIDATION_ERROR');
  assert.equal(t.points.sheet.appended.length,0);
  assert.equal(t.candidates.sheet.rows[0]['候補状態'],'pending');
  assert.throws(() => context.candidateRegistrationPosition_({registration_position:{latitude:null,longitude:127}}), error => error.code === 'VALIDATION_ERROR');
  assert.throws(() => context.candidateRegistrationPosition_({registration_position:{latitude:26,longitude:'NaN'}}), error => error.code === 'VALIDATION_ERROR');
});

test('candidate promoteは推定基地を後方互換の初期値として06へ保存する', () => {
  const t=candidateHarness('pending',{'推定base_area_id':'base-1'},[baseArea()]);
  const incomingPlace=place({place_id:undefined,revision:undefined,'地点名':'候補'});
  delete incomingPlace.base_area_id;
  const res=context.resolveCandidate_({action:'promote',candidate_id:'c1',revision:1,bundle:{place:incomingPlace,points:[point({point_id:undefined,place_id:undefined,revision:undefined})]}});
  assert.equal(res.data.place.base_area_id,'base-1');
  assert.equal(t.candidates.sheet.rows[0]['推定base_area_id'],'base-1');
});

test('candidate promoteは推定基地より明示選択を優先し08推定値を変えない', () => {
  const t=candidateHarness('pending',{'推定base_area_id':'base-1'},[baseArea(),baseArea({base_area_id:'base-2','基地名':'カデナ'})]);
  const res=context.resolveCandidate_({
    action:'promote',candidate_id:'c1',revision:1,registration_base_area_id:'base-2',
    bundle:{place:place({place_id:undefined,revision:undefined,'地点名':'候補',base_area_id:'base-2'}),points:[point({point_id:undefined,place_id:undefined,revision:undefined})]}
  });
  assert.equal(res.data.place.base_area_id,'base-2');
  assert.equal(t.candidates.sheet.rows[0]['推定base_area_id'],'base-1');
});

test('candidate promoteは基地なしを空欄として保存する', () => {
  const t=candidateHarness('pending',{'推定base_area_id':'base-1'},[baseArea()]);
  const res=context.resolveCandidate_({
    action:'promote',candidate_id:'c1',revision:1,registration_base_area_id:'',
    bundle:{place:place({place_id:undefined,revision:undefined,'地点名':'候補',base_area_id:''}),points:[point({point_id:undefined,place_id:undefined,revision:undefined})]}
  });
  assert.equal(res.data.place.base_area_id,'');
  assert.equal(t.candidates.sheet.rows[0]['推定base_area_id'],'base-1');
});

test('candidate promoteは位置調整と所属基地選択を同時に反映する', () => {
  const t=candidateHarness('pending',{'推定base_area_id':'base-1'},[baseArea(),baseArea({base_area_id:'base-2','基地名':'カデナ'})]);
  const res=context.resolveCandidate_({
    action:'promote',candidate_id:'c1',revision:1,registration_base_area_id:'base-2',registration_position:{latitude:26.4,longitude:127.8},
    bundle:{place:place({place_id:undefined,revision:undefined,'地点名':'候補',base_area_id:'base-2'}),points:[point({point_id:undefined,place_id:undefined,revision:undefined})]}
  });
  assert.equal(res.data.place.base_area_id,'base-2');
  assert.equal(res.data.points[0]['中心緯度'],26.4);
  assert.equal(res.data.points[0]['中心経度'],127.8);
  assert.equal(t.candidates.sheet.rows[0]['基準緯度'],26);
  assert.equal(t.candidates.sheet.rows[0]['推定base_area_id'],'base-1');
});

test('candidate promoteは不正または無効な所属基地を拒否する', () => {
  let t=candidateHarness('pending',{},[baseArea()]);
  const payload={action:'promote',candidate_id:'c1',revision:1,registration_base_area_id:'missing',bundle:{place:place({place_id:undefined,revision:undefined}),points:[point({point_id:undefined,place_id:undefined,revision:undefined})]}};
  assert.throws(() => context.resolveCandidate_(payload), error => error.code === 'VALIDATION_ERROR');
  assert.equal(t.candidates.sheet.rows[0]['候補状態'],'pending');
  t=candidateHarness('pending',{},[baseArea({base_area_id:'inactive','有効':false})]);
  assert.throws(() => context.resolveCandidate_(Object.assign({},payload,{registration_base_area_id:'inactive'})), error => error.code === 'VALIDATION_ERROR');
});

test('candidate add_pointは既存地点の所属基地を変更しない', () => {
  const t=candidateHarness();
  t.places.sheet.rows[0].base_area_id='base-1';
  t.places.rows[0].base_area_id='base-1';
  const res=context.resolveCandidate_({action:'add_point',candidate_id:'c1',revision:1,place_id:'a',place_revision:1,point:point({'ポイント名':'追加位置'})});
  assert.equal(res.data.place.base_area_id,'base-1');
  assert.equal(t.places.sheet.rows[0].base_area_id,'base-1');
});

function mergeHarness() {
  const places=table('places',[place(),place({place_id:'b','地点名':'B',revision:4})]);
  const points=table('points',[point(),point({point_id:'p2',place_id:'b','ポイント名':'裏側入口','表示順':1,revision:3})]);
  context.withWriteLock_=fn=>fn(); context.spreadsheet_=()=>({});
  context.readTable_=(_ss,key)=>key === 'places' ? places : points;
  return {places,points};
}

function mergePayload(overrides={}) {
  return Object.assign({
    source_place_id:'b', source_revision:4,
    target_place_id:'a', target_revision:1,
    preserve_source_name:true,
    point_revisions:[{point_id:'p1',revision:1},{point_id:'p2',revision:3}]
  },overrides);
}

test('登録地点どうしを統合しpoint_idを維持して移管・再採番する', () => {
  const t=mergeHarness();
  const result=context.mergePlaces_(mergePayload());
  assert.equal(result.data.place.place_id,'a');
  assert.equal(result.data.place.revision,2);
  assert.match(result.data.place['別名'],/(^|, )B($|,)/);
  assert.equal(result.data.merged_place.place_id,'b');
  assert.equal(result.data.merged_place['有効'],false);
  assert.equal(result.data.merged_place.revision,5);
  assert.match(result.data.merged_place['備考'],/aへ統合/);
  assert.deepEqual(Array.from(result.data.points,p=>p.point_id),['p1','p2']);
  assert.deepEqual(Array.from(result.data.points,p=>p.place_id),['a','a']);
  assert.deepEqual(Array.from(result.data.points,p=>p['表示順']),[1,2]);
  assert.equal(result.data.points[0].revision,1);
  assert.equal(result.data.points[1].revision,4);
});

test('地点統合は同名を別名へ重複追加せず、全revisionを検証する', () => {
  let t=mergeHarness();
  t.places.rows[1]['地点名']='A'; t.places.sheet.rows[1]['地点名']='A';
  const result=context.mergePlaces_(mergePayload());
  assert.equal(result.data.place['別名'],'');
  t=mergeHarness();
  assert.throws(() => context.mergePlaces_(mergePayload({source_revision:3})), e => e.code === 'REVISION_CONFLICT');
  t=mergeHarness();
  assert.throws(() => context.mergePlaces_(mergePayload({point_revisions:[{point_id:'p1',revision:1},{point_id:'p2',revision:2}]})), e => e.code === 'REVISION_CONFLICT');
});

test('地点統合の途中失敗はplaceとpointを補償復元する', () => {
  const t=mergeHarness();
  const beforePlaces=JSON.parse(JSON.stringify(t.places.sheet.rows));
  const beforePoints=JSON.parse(JSON.stringify(t.points.sheet.rows));
  t.points.sheet.failAfter=0;
  assert.throws(() => context.mergePlaces_(mergePayload()), /simulated write failure/);
  assert.deepEqual(t.places.sheet.rows.map(row => Object.fromEntries(t.places.headers.map(h => [h,row[h] == null ? '' : row[h]]))), beforePlaces.map(row => Object.fromEntries(t.places.headers.map(h => [h,row[h] == null ? '' : row[h]]))));
  assert.deepEqual(t.points.sheet.rows.map(row => Object.fromEntries(t.points.headers.map(h => [h,row[h] == null ? '' : row[h]]))), beforePoints.map(row => Object.fromEntries(t.points.headers.map(h => [h,row[h] == null ? '' : row[h]]))));
});

test('論理削除は有効=falseとして保持する', () => {
  const now='2026-01-01T00:00:00.000Z';
  const disabled=context.normalizePoint_(point({'有効':false}),'p1','a',now,false,point());
  assert.equal(disabled['有効'],false);
  assert.equal(disabled.point_id,'p1');
  assert.equal(disabled.revision,2);
});

test('地点専用編集レイアウトと地図は確定したGrid行の内側に収まる', () => {
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const index=fs.readFileSync(path.join(root,'Index.html'),'utf8');
  assert.doesNotMatch(index,/leaflet@1\.9\.4\/dist\/leaflet\.css/);
  assert.doesNotMatch(index,/unpkg\.com\/leaflet/);
  assert.match(index,/include\('Leaflet'\)/);
  assert.match(fs.readFileSync(path.join(root,'Leaflet.html'),'utf8'),/Leaflet 1\.9\.4/);
  assert.match(styles,/\.leaflet-pane[^}]*position:absolute;[^}]*left:0;[^}]*top:0;/);
  assert.match(styles,/\.leaflet-container \{[^}]*overflow:hidden;/);
  assert.match(styles,/html,body \{[^}]*height:100%;[^}]*overflow:hidden;/);
  assert.match(styles,/\.workspace \{[^}]*min-height:0;[^}]*grid-template-rows:minmax\(0,1fr\);[^}]*overflow:hidden;/);
  assert.doesNotMatch(styles,/\.workspace\.focus-mode/);
  assert.match(styles,/\.studio-grid \{[^}]*grid-template-columns:minmax\(420px,1fr\) 390px/);
  assert.match(styles,/\.panel \{[^}]*min-height:0;[^}]*overflow:auto;/);
  assert.match(styles,/\.map-panel \{[^}]*min-height:0;[^}]*overflow:hidden;/);
  assert.match(styles,/#map \{[^}]*height:100%;[^}]*min-height:0;[^}]*overflow:hidden;/);
  assert.match(client,/renderAll\(\);\s*await new Promise\(resolve => requestAnimationFrame\(resolve\)\);\s*map\.invalidateSize\(\{ pan:false \}\);\s*fitData\(\);/);
  assert.ok((client.match(/map\.invalidateSize/g)||[]).length >= 2);
});

test('地図候補マーカーと左一覧は共通の候補選択処理を使う', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/data-candidate[^\n]+selectCandidate\(el\.dataset\.candidate\)/);
  assert.match(client,/marker\.on\('click', \(\) => selectCandidate\(c\.candidate_id\)\)/);
  assert.match(client,/markerElement\.dataset\.candidate = c\.candidate_id/);
  assert.match(client,/markerElement\.setAttribute\('aria-label', `候補: \$\{c\['候補名'\]\}`\)/);
  assert.match(client,/state\.selectedCandidateId = candidateId/);
  assert.match(client,/if \(state\.selectedCandidateId\) return renderCandidateEditor\(\)/);
});

test('候補選択は左一覧へ同期し、正式地点の地図コンテキストを維持する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  assert.match(client,/c\.candidate_id === state\.selectedCandidateId \? 'active'/);
  assert.match(client,/requestAnimationFrame\(scrollSelectedCandidateIntoView\)/);
  assert.match(client,/state\.mapContextPlaceId = state\.selectedPlaceId/);
  assert.match(client,/const activePlaceId = state\.selectedPlaceId \|\| state\.mapContextPlaceId/);
  assert.match(client,/candidate-icon\$\{selected \? ' selected' : ''\}/);
  assert.match(styles,/\.candidate-icon\.selected/);
});

test('候補の登録予定位置はクライアントdraftでのみ移動し候補元座標を保持する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  assert.match(client,/candidateDraft: null/);
  assert.match(client,/originalLatitude:location\['中心緯度'\]/);
  assert.match(client,/draggable:!!\(candidateDraft && candidateDraft\.editing\)/);
  assert.match(client,/marker\.on\('dragend', event => previewCandidatePosition/);
  assert.match(client,/id="moveCandidatePosition"/);
  assert.match(client,/id="confirmCandidatePosition"/);
  assert.match(client,/id="restoreCandidateOriginal"/);
  assert.match(client,/id="cancelCandidateMove"/);
  assert.match(client,/draft\.latitude = draft\.originalLatitude/);
  assert.match(client,/draft\.longitude = draft\.originalLongitude/);
  assert.match(styles,/\.candidate-icon\.adjusting/);
});

test('候補位置調整はpromoteとadd_pointだけregistration_positionへ渡す', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/if \(action === 'promote' \|\| action === 'add_point'\) payload\.registration_position/);
  assert.match(client,/function hasUnsavedChanges\(\) \{ return state\.dirty \|\| candidateDraftChanged\(\); \}/);
  assert.match(client,/候補の登録予定位置または認識範囲が変更されています。変更を破棄して移動しますか？/);
  assert.match(client,/window\.addEventListener\('beforeunload',[\s\S]*?hasUnsavedChanges\(\)/);
});

test('候補の予定円と500m以内の既存point円を表示し異なる地点の重複だけ警告する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  assert.match(client,/const CANDIDATE_NEARBY_DISTANCE_M = 500/);
  assert.match(client,/candidateCircle = L\.circle\(\[latitude,longitude\],[\s\S]*?radius:Number\(candidateDraft\.form\.radius\)/);
  assert.match(client,/function candidateNearbyPoints\(point\)[\s\S]*?item\.distance <= CANDIDATE_NEARBY_DISTANCE_M \|\| item\.overlap > 0/);
  assert.match(client,/const samePlace = !!point\.place_id && p\.place_id === point\.place_id/);
  assert.match(client,/const conflicting = !samePlace && overlap > 0/);
  assert.match(client,/同じ地点（重なり可）/);
  assert.match(client,/異なる地点との重なり \$\{conflicts\.length\}件/);
  assert.match(styles,/\.candidate-nearby-item\.overlap/);
});

test('候補ドラッグと半径入力は円・推奨値・重複表示へ即時反映し取消でdraftを復元する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/marker\.on\('drag', event => previewCandidateDrag\(event\.target\.getLatLng\(\),candidateCircle\)\)/);
  assert.match(client,/function previewCandidateDrag\(latlng,circle\)[\s\S]*?circle\.setLatLng\(latlng\)[\s\S]*?updateCandidateProximity\(\)/);
  assert.match(client,/candidate-radius'\)\.addEventListener\('input',[\s\S]*?renderMap\(\); updateCandidateProximity\(\); updateSaveState\(\)/);
  assert.match(client,/推奨認識範囲：\$\{suggestion\.radius\}m \/ 現在設定：\$\{currentRadius\}m/);
  assert.match(client,/draft\.form\.radius = draft\.moveSession\.radius/);
  assert.match(client,/state\.candidateDraft = createCandidateDraft\(candidate\)/);
  assert.doesNotMatch(client,/function recalculateCandidateRecommendation/);
});

test('候補promote前に有効基地または基地なしを選択して確認できる', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  assert.match(client,/state\.baseAreas\.find\(base => base\.base_area_id === candidate\['推定base_area_id'\] && base\['有効'\] !== false\)/);
  assert.match(client,/baseAreaId:inferredBaseArea \? inferredBaseArea\.base_area_id : ''/);
  assert.match(client,/state\.baseAreas\.filter\(base => base\['有効'\] !== false\)/);
  assert.match(client,/<option value="">基地なし<\/option>/);
  assert.match(client,/id="candidate-base"/);
  assert.match(client,/登録内容を確認/);
  assert.match(client,/id="cancelPromoteCandidate"/);
  assert.match(client,/id="confirmPromoteCandidate"/);
  assert.match(client,/この内容で登録/);
  assert.match(styles,/\.confirm-grid/);
});

test('候補promote payloadは選択基地を明示し位置移動では基地選択を変えない', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const startMove=client.match(/function startCandidateMove\(\) \{([\s\S]*?)\n  \}/);
  assert.ok(startMove);
  assert.doesNotMatch(startMove[1],/baseAreaId/);
  assert.match(client,/payload\.registration_base_area_id = draft\.form\.baseAreaId/);
  assert.match(client,/base_area_id:draft\.form\.baseAreaId/);
  assert.match(client,/draft\.form\.baseAreaId = e\.target\.value/);
});

test('候補マーカー選択は未保存変更保護を通り、正式point選択を維持する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const selectCandidateBody=client.match(/function selectCandidate\(candidateId\) \{([\s\S]*?)\n  \}\n\n  function scrollSelectedCandidateIntoView/);
  assert.ok(selectCandidateBody);
  assert.match(selectCandidateBody[1],/!canLeave\(\)/);
  assert.match(client,/marker\.on\('click', \(\) => inContext \? selectPoint\(id\) : \(point\.place_id && \(state\.mergeFlow \? chooseMergeTarget\(point\.place_id\) : selectPlace\(point\.place_id, point\.point_id\)\)\)\)/);
  assert.match(client,/candidateRead\('candidate_id',c\.candidate_id\)/);
  assert.match(client,/a\.distance-b\.distance \|\| String\(a\.place\['地点名'\]\)\.localeCompare\(String\(b\.place\['地点名'\]\),'ja'\)/);
});

test('左地点一覧を常設し、戻る画面遷移なしで編集対象を切り替える', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const index=fs.readFileSync(path.join(root,'Index.html'),'utf8');
  assert.match(index,/class="left-panel panel"/);
  assert.match(index,/id="pointNavigator"/);
  assert.doesNotMatch(index,/id="pointPanel"/);
  assert.doesNotMatch(index,/id="backToListBtn"/);
  assert.doesNotMatch(client,/focus-mode/);
  assert.doesNotMatch(client,/function backToList/);
  assert.match(client,/function showPlaceOverview\(\)[\s\S]*?state\.selectedPointId = null/);
  assert.match(client,/id="placeOverviewBtn" class="target-card overview/);
  assert.match(client,/id="addPointBtn"[^>]*>＋ 乗降位置/);
});

test('地点モード左カラムは学習候補を登録地点より上に表示する', () => {
  const index=fs.readFileSync(path.join(root,'Index.html'),'utf8');
  const candidatePosition=index.indexOf('id="candidateList"');
  const placePosition=index.indexOf('id="placeList"');
  assert.ok(candidatePosition >= 0 && placePosition > candidatePosition);
  assert.match(index,/id="placeModeList"[\s\S]*?学習候補[\s\S]*?candidateList[\s\S]*?登録地点[\s\S]*?placeList/);
});

test('地点一覧は安定した空間順をキャッシュし、常設中も検索とスクロールを保持する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const index=fs.readFileSync(path.join(root,'Index.html'),'utf8');
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  assert.match(index,/include\('SpatialOrder'\)/);
  assert.match(client,/placeOrderCache: \{ signature:null, records:\[\], baseSignature:null, baseRecords:\[\] \}/);
  assert.match(client,/GeoPlaceOrder\.signature\(places, state\.points, state\.baseAreas, state\.municipalities, municipalityIndex\)/);
  assert.match(client,/GeoPlaceOrder\.orderPlaces\(places,state\.points,state\.baseAreas,state\.municipalities,municipalityIndex\)/);
  assert.match(styles,/\.place-group-label[^}]*width:100%[^}]*font-size:16px/);
  assert.match(client,/renderBasePlaceList\(matchingPlaces, q\)/);
  assert.match(client,/renderMunicipalityPlaceList\(matchingPlaces, q\)/);
  assert.match(client,/expandedPlaceGroups: new Set\(\)/);
  assert.match(client,/const expanded=searching\|\|state\.expandedPlaceGroups\.has\(group\.key\)/);
  assert.match(client,/data-place-group="\$\{esc\(group\.key\)\}" aria-expanded="\$\{expanded\}"/);
  assert.match(client,/expanded\?'▼':'▶'/);
  assert.match(client,/function togglePlaceGroup\(groupKey\)[\s\S]*?state\.expandedPlaceGroups\.delete\(groupKey\)[\s\S]*?state\.expandedPlaceGroups\.add\(groupKey\)[\s\S]*?renderLists\(\)/);
  assert.match(client,/renderPlaceGroups\(\$\('basePlaceList'\)[\s\S]*?Boolean\(q\.trim\(\)\)/);
  assert.match(client,/renderPlaceGroups\(container[\s\S]*?Boolean\(context\.q\.trim\(\)\)/);
  assert.match(client,/state\.municipalities = response\.data\.municipalities/);
  assert.match(client,/listView: \{ scrollTop:0, lastPlaceId:null, lastCandidateId:null \}/);
  assert.match(client,/state\.listView\.scrollTop = panel\.scrollTop/);
  assert.match(client,/left-panel'\)\.scrollTop = state\.listView\.scrollTop/);
  assert.match(client,/left-panel'\)\.addEventListener\('scroll'/);
  assert.match(client,/state\.query=e\.target\.value/);
  assert.match(styles,/\.list-item\.recent/);
});

test('市町村GeoJSON準備中は未判定グループを作らず、判定中表示を出す', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const loadingGate=client.indexOf("if(resource.status!=='ready')");
  const orderCall=client.indexOf('renderPlaceGroups(container,orderedPlaceRecords()', loadingGate);
  assert.ok(loadingGate >= 0 && orderCall > loadingGate);
  assert.match(client,/市町村を判定中…/);
  assert.match(client,/const municipalityReady = loadBoundary\('municipality'\);[\s\S]*?renderAll\(\)/);
  assert.match(client,/municipalityReady\.catch\(\(\) => \{\}\)/);
  assert.match(client,/renderBasePlaceList\(matchingPlaces, q\);[\s\S]*?renderMunicipalityPlaceList\(matchingPlaces, q\)/);
});

test('市町村GeoJSON取得失敗は未判定扱いにせず再試行できる', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/resource\.status==='error'[\s\S]*?市町村境界を読み込めませんでした[\s\S]*?retryMunicipalityBtn/);
  assert.match(client,/retryMunicipalityBtn'\)\)\$\('retryMunicipalityBtn'\)\.addEventListener\('click',retryMunicipalityBoundary\)/);
  assert.match(client,/function retryMunicipalityBoundary\(\)[\s\S]*?loadBoundary\('municipality'\)/);
});

test('境界ロードはreadyキャッシュを再利用し、古い世代の完了結果を破棄する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/if \(resource\.status==='ready'\) return resource/);
  assert.match(client,/if \(resource\.promise\) return resource\.promise/);
  assert.match(client,/const generation=\+\+resource\.generation/);
  assert.ok((client.match(/if\s*\(generation!==resource\.generation\)\s*return resource/g)||[]).length>=2);
});

test('境界assetはmanifestとIndexedDB永続キャッシュを使いversion・SHA変更を分離する', () => {
  const cache=fs.readFileSync(path.join(root,'BoundaryCache.html'),'utf8');
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const code=fs.readFileSync(path.join(root,'Code.gs'),'utf8');
  const service=fs.readFileSync(path.join(root,'BoundaryService.gs'),'utf8');
  assert.match(cache,/const DB_NAME = 'geo-editor-boundary-assets'/);
  assert.match(cache,/\[manifest\.kind, manifest\.version, manifest\.sha256\]\.join\(':'\)/);
  assert.match(cache,/indexedDB\.open/);
  assert.match(cache,/crypto\.subtle\.digest\('SHA-256'/);
  assert.match(cache,/metadataMatches[\s\S]*?digestGeoJson\(record\.geojson\)/);
  assert.match(client,/server\('getGeoBoundaryManifest',kind\)/);
  assert.match(client,/GeoBoundaryCache\.get\(manifest\)/);
  assert.match(client,/GeoBoundaryCache\.put\(manifest,geojson\)/);
  assert.match(client,/data\.version!==manifest\.version\|\|data\.sha256!==manifest\.sha256/);
  assert.match(code,/function getGeoBoundaryManifest\(kind\)/);
  assert.match(service,/function getGeoBoundaryManifest_\(kind\)/);
});

test('破損キャッシュは削除して正式取得へ戻りbbox fallbackを使わない', () => {
  const cache=fs.readFileSync(path.join(root,'BoundaryCache.html'),'utf8');
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(cache,/!metadataMatches \|\| !validGeoJson\(record\.geojson\)[\s\S]*?remove\(manifest\)/);
  assert.match(cache,/digestGeoJson\(record\.geojson\) !== record\.cacheSha256[\s\S]*?remove\(manifest\)/);
  assert.match(client,/catch\(error\)\{try\{await GeoBoundaryCache\.remove\(manifest\);\}catch\(ignore\)\{\}geojson=null;index=null;\}/);
  assert.match(client,/if\(!geojson\)[\s\S]*?server\('getGeoBoundary',kind\)/);
});

test('基地・市町村グループは初期全閉で複数展開でき、検索中だけ該当グループを自動展開する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  assert.match(client,/expandedPlaceGroups: new Set\(\)/);
  assert.match(client,/Boolean\(q\.trim\(\)\)/);
  assert.match(client,/const expanded=searching\|\|state\.expandedPlaceGroups\.has\(group\.key\)/);
  assert.match(client,/if\(!expanded\)return heading/);
  assert.match(client,/aria-expanded="\$\{expanded\}"/);
  assert.match(client,/expanded\?'▼':'▶'/);
  assert.match(client,/state\.expandedPlaceGroups\.has\(groupKey\)[\s\S]*?delete\(groupKey\)[\s\S]*?add\(groupKey\)/);
  assert.doesNotMatch(client,/expandedPlaceGroups\s*=\s*new Set/);
  assert.match(client,/if \(panel\) state\.listView\.scrollTop = panel\.scrollTop;[\s\S]*?renderLists\(\)/);
  assert.match(styles,/\.place-group-label[^}]*width:100%[^}]*cursor:pointer/);
  assert.match(styles,/\.place-group-items[^}]*display:flex/);
});

test('全乗降位置を常時一覧表示し、一覧と地図は共通のpoint選択処理を使う', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/state\.draftPoints\.map\(point =>/);
  assert.match(client,/data-edit-point="\$\{esc\(id\)\}"/);
  assert.match(client,/data-edit-point[^\n]+selectPoint\(el\.dataset\.editPoint\)/);
  assert.match(client,/marker\.on\('click', \(\) => inContext \? selectPoint\(id\)/);
  assert.match(client,/letter-icon\$\{isActive \? ' selected' : ''\}/);
  assert.match(client,/navigator\.innerHTML = `<div class="target-heading">/);
  assert.match(client,/target-card \$\{id === state\.selectedPointId \? 'active'/);
});

test('左一覧から別地点と候補へ直接切り替える際は未保存変更保護を通る', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const selectPlaceBody=client.match(/function selectPlace\(placeId, pointId\) \{([\s\S]*?)\n  \}\n\n  function selectPoint/);
  const selectCandidateBody=client.match(/function selectCandidate\(candidateId\) \{([\s\S]*?)\n  \}\n\n  function scrollSelectedCandidateIntoView/);
  assert.ok(selectPlaceBody);
  assert.ok(selectCandidateBody);
  assert.match(selectPlaceBody[1],/placeId === state\.selectedPlaceId/);
  assert.match(selectPlaceBody[1],/if \(!canLeave\(\)\) return/);
  assert.match(selectCandidateBody[1],/!canLeave\(\)/);
  assert.match(client,/state\.selectedPlaceId = placeId/);
  assert.match(client,/state\.selectedCandidateId = candidateId/);
});

test('point切替はdraftを維持し、移動操作中だけ確認して地図と同期する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const selectPointBody=client.match(/function selectPoint\(pointId\) \{([\s\S]*?)\n  \}\n\n  function showPlaceOverview/);
  assert.ok(selectPointBody);
  assert.match(selectPointBody[1],/state\.moveSession/);
  assert.match(selectPointBody[1],/state\.selectedPointId = pointId/);
  assert.doesNotMatch(selectPointBody[1],/state\.draftPoints = \[\]/);
  assert.match(selectPointBody[1],/scrollRightPanelToEditor\(\)/);
  assert.match(selectPointBody[1],/map\.panTo/);
  assert.match(client,/marker\.on\('click', \(\) => inContext \? selectPoint\(id\)/);
});

test('右カラム上部を固定せず全体を1つの縦スクロール領域にする', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  const navigatorRule=styles.match(/\.point-navigator \{([^}]*)\}/);
  const panelRule=styles.match(/\.right-panel \{([^}]*)\}/);
  assert.ok(navigatorRule);
  assert.ok(panelRule);
  assert.match(navigatorRule[1],/position:static/);
  assert.doesNotMatch(navigatorRule[1],/position:(?:sticky|fixed)/);
  assert.match(panelRule[1],/min-height:0/);
  assert.match(panelRule[1],/overflow-y:auto/);
  assert.match(client,/function resetRightPanelScroll\(\)[\s\S]*?right-panel'\)\.scrollTop = 0/);
  assert.match(client,/function scrollRightPanelToEditor\(\)[\s\S]*?editor\.getBoundingClientRect\(\)\.top - panel\.getBoundingClientRect\(\)\.top/);
  assert.match(client,/function selectPlace\(placeId, pointId\)[\s\S]*?resetRightPanelScroll\(\)/);
  assert.match(client,/function selectCandidate\(candidateId\)[\s\S]*?resetRightPanelScroll\(\)/);
  assert.match(client,/function showPlaceOverview\(\)[\s\S]*?scrollRightPanelToEditor\(\)/);
});

test('乗降位置追加は位置指定・内容確認・保存のガイドを表示する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const index=fs.readFileSync(path.join(root,'Index.html'),'utf8');
  assert.match(index,/id="placementGuide"/);
  assert.match(client,/1 地図で位置を指定/);
  assert.match(client,/2 名称と認識範囲/);
  assert.match(client,/3 保存/);
  assert.match(client,/state\.placement = 'add-point'/);
  assert.match(client,/state\.addFlowPointId = point\._temp_id/);
  assert.match(client,/function newPoint[\s\S]*?'ポイント名':''/);
});

test('位置移動は確定と元に戻すを備え、移動前座標を保持する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/state\.moveSession = \{ pointId:id, lat:point\['中心緯度'\], lng:point\['中心経度'\], wasDirty:state\.dirty \}/);
  assert.match(client,/id="confirmMove"/);
  assert.match(client,/id="cancelMove"/);
  assert.match(client,/point\['中心緯度'\] = session\.lat/);
  assert.match(client,/point\['中心経度'\] = session\.lng/);
});

test('認識範囲はスライダーと数値入力を同期し地図へ即時反映する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/id="point-radius-range" type="range"/);
  assert.match(client,/range\.addEventListener\('input'/);
  assert.match(client,/number\.addEventListener\('input'/);
  assert.match(client,/point\['判定半径_m'\] = radius/);
  assert.match(client,/renderMap\(\)/);
});

test('乗降位置の基本設定と詳細設定を分離し、親地点を常時表示する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/編集中の地点：\$\{esc\(state\.draftPlace\['地点名'\]/);
  assert.match(client,/<h3>基本設定<\/h3>/);
  assert.match(client,/<details class="advanced"><summary>詳細設定<\/summary>/);
  assert.match(client,/GPS許容誤差（m）/);
  assert.match(client,/判定優先度/);
});

test('候補処理は利用者向け用語で既存actionを維持する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/>乗降位置として追加<\/button>/);
  assert.match(client,/resolveCandidateAction\('add_point'\)/);
  assert.match(client,/resolveCandidateAction\('promote'\)/);
  assert.match(client,/resolveCandidateAction\('merge'\)/);
  assert.match(client,/resolveCandidateAction\('reject'\)/);
  assert.match(client,/② 既存地点にまとめる/);
  assert.match(client,/同じ場所・同じ乗降位置として扱います/);
  assert.match(client,/③ 既存地点の乗降位置として追加/);
});

test('地点をまとめるUIは近隣・検索・地図選択と確認を共通処理へつなぐ', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/id="startMergePlace">この地点を別の地点にまとめる/);
  assert.match(client,/function startMergeFlow\(\)/);
  assert.match(client,/function chooseMergeTarget\(placeId\)/);
  assert.match(client,/state\.mergeFlow\?chooseMergeTarget\(el\.dataset\.place\):selectPlace/);
  assert.match(client,/state\.mergeFlow \? chooseMergeTarget\(point\.place_id\) : selectPlace/);
  assert.match(client,/近くの地点/);
  assert.match(client,/残す地点/);
  assert.match(client,/まとめる地点/);
  assert.match(client,/この内容でまとめる/);
  assert.match(client,/server\('mergePlaces', payload\)/);
});

test('無効化済みの旧地点は通常一覧と地図から除外する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/state\.places\.filter\(p => p\['有効'\] !== false/);
  assert.match(client,/isPlaceActive\(p\.place_id\)/);
});

test('地点・基地エリア・市町村境界の3モードとread-only境界診断を備える', () => {
  const index=fs.readFileSync(path.join(root,'Index.html'),'utf8');
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const service=fs.readFileSync(path.join(root,'BoundaryService.gs'),'utf8');
  assert.match(index,/id="modePlaces"/); assert.match(index,/id="modeBases"/); assert.match(index,/id="modeMunicipalities"/);
  assert.match(index,/include\('BoundaryGeometry'\)/);
  assert.match(client,/GeoBoundary\.buildIndex/);
  assert.match(client,/GeoBoundary\.bboxCandidates/);
  assert.match(client,/if \(state\.mode !== 'places'\) \{ inspectBoundaryPoint/);
  assert.match(client,/readonly-badge/);
  assert.match(client,/市町村bbox候補/);
  assert.match(client,/市町村Polygon/);
  assert.match(service,/SMART_LOGGER_BOUNDARY_ENDPOINT/);
  assert.match(service,/BOUNDARY_INTEGRITY_ERROR/);
  assert.match(service,/Utilities\.computeDigest/);
  assert.doesNotMatch(service,/SpreadsheetApp|setValue|setValues|appendRow/);
});

test('境界取得サービスは種別をallowlistしSHA不一致を拒否する', () => {
  const sample=Buffer.from(JSON.stringify({type:'FeatureCollection',features:[]}));
  const serviceContext=vm.createContext({
    console, JSON, Array, String, Error, encodeURIComponent,
    geoError_:(code,message)=>Object.assign(new Error(message),{code}),
    UrlFetchApp:{fetch:()=>({getResponseCode:()=>200,getBlob:()=>({getBytes:()=>Array.from(sample)}),getContentText:()=>sample.toString('utf8')})},
    Utilities:{DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(_alg,bytes)=>Array.from(crypto.createHash('sha256').update(Buffer.from(bytes)).digest(),value=>value>127?value-256:value)}
  });
  vm.runInContext(fs.readFileSync(path.join(root,'BoundaryService.gs'),'utf8'),serviceContext,{filename:'BoundaryService.gs'});
  assert.throws(()=>serviceContext.getGeoBoundary_('town'),error=>error.code==='VALIDATION_ERROR');
  assert.throws(()=>serviceContext.getGeoBoundary_('base'),error=>error.code==='BOUNDARY_INTEGRITY_ERROR');
  serviceContext.GEO_BOUNDARY_CONFIG.base.sha256=crypto.createHash('sha256').update(sample).digest('hex');
  const result=serviceContext.getGeoBoundary_('base');
  assert.equal(result.geojson.type,'FeatureCollection');
  assert.equal(result.sha256,serviceContext.GEO_BOUNDARY_CONFIG.base.sha256);
});
