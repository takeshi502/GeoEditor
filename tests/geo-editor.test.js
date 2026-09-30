'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
let uuidCounter = 0;
const context = vm.createContext({
  console, Date, Math, Number, Object, String, Array, Boolean, JSON, Error, RegExp,
  isFinite, parseInt, parseFloat,
  Utilities: { getUuid: () => `uuid-${++uuidCounter}` }
});
for (const name of ['Schema.gs','Code.gs','Validation.gs','GeoMath.gs','PlaceService.gs','CandidateService.gs']) {
  vm.runInContext(fs.readFileSync(path.join(root, name), 'utf8'), context, { filename:name });
}

function point(overrides = {}) {
  return Object.assign({ point_id:'p1', place_id:'a', 'ポイント名':'main', '表示順':1, '中心緯度':26, '中心経度':127, '判定半径_m':20, '最大GPS精度_m':20, '優先度':100, '有効':true, revision:1 }, overrides);
}
function place(overrides = {}) {
  return Object.assign({ place_id:'a', '地点名':'A', base_area_id:'', '優先度':0, '別名':'', '登録元':'test', '最終使用日時':'', '有効':true, '備考':'', '作成日時':'x', '更新日時':'x', revision:1 }, overrides);
}
function fakeSheet(headers, rows) {
  const sheet = {
    headers, rows: rows.map(r => Object.assign({}, r)), writes:[], appended:[],
    getLastRow() { return this.rows.length + 1; },
    getRange(row, col, height, width) {
      return {
        setValues: values => {
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
  assert.doesNotThrow(() => context.validateHeaders_('x', Array.from(context.GEO_SCHEMA.places.headers), context.GEO_SCHEMA.places.headers));
  assert.throws(() => context.validateHeaders_('x', ['wrong'], context.GEO_SCHEMA.places.headers), e => e.code === 'SCHEMA_MISMATCH');
});

test('距離計算は既知の緯度差をメートルで返す', () => {
  assert.ok(Math.abs(context.distanceMeters_(0, 0, 0.001, 0) - 111.195) < 0.2);
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
  assert.match(client,/const place = clone\(state\.draftPlace\)/);
  assert.match(client,/place\['地点名'\] = String\(place\['地点名'\] == null \? '' : place\['地点名'\]\)/);
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

function candidateHarness(status='pending') {
  const candidate={ candidate_id:'c1','候補名':'候補','推定所属種別':'unknown','推定base_area_id':'','推定town_id':'','基準緯度':26,'基準経度':127,'GPS精度_m':10,'確認回数':2,'初回確認日時':'','最終確認日時':'','初回source_type':'','初回source_id':'','最近source_type':'','最近source_id':'','最寄place_id':'','最寄地点距離_m':'','候補状態':status,'登録先place_id':'',revision:1 };
  const tables={ candidates:table('candidates',[candidate]), places:table('places',[place()]), points:table('points',[point()]), baseAreas:table('baseAreas',[]) };
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

test('論理削除は有効=falseとして保持する', () => {
  const now='2026-01-01T00:00:00.000Z';
  const disabled=context.normalizePoint_(point({'有効':false}),'p1','a',now,false,point());
  assert.equal(disabled['有効'],false);
  assert.equal(disabled.point_id,'p1');
  assert.equal(disabled.revision,2);
});

test('3カラムと地図は確定したGrid行の内側に収まる', () => {
  const styles=fs.readFileSync(path.join(root,'Styles.html'),'utf8');
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const index=fs.readFileSync(path.join(root,'Index.html'),'utf8');
  assert.doesNotMatch(index,/leaflet@1\.9\.4\/dist\/leaflet\.css/);
  assert.match(styles,/\.leaflet-pane[^}]*position:absolute;[^}]*left:0;[^}]*top:0;/);
  assert.match(styles,/\.leaflet-container \{[^}]*overflow:hidden;/);
  assert.match(styles,/html, body \{[^}]*height:100%;[^}]*overflow:hidden;/);
  assert.match(styles,/\.workspace \{[^}]*min-height:0;[^}]*grid-template-rows:minmax\(0, 1fr\);[^}]*overflow:hidden;/);
  assert.match(styles,/\.panel \{[^}]*min-height:0;[^}]*overflow:auto;/);
  assert.match(styles,/\.map-panel \{[^}]*min-height:0;[^}]*overflow:hidden;/);
  assert.match(styles,/#map \{[^}]*height:100%;[^}]*min-height:0;[^}]*overflow:hidden;/);
  assert.match(client,/renderAll\(\);\s*await new Promise\(resolve => requestAnimationFrame\(resolve\)\);\s*map\.invalidateSize\(\{ pan:false \}\);\s*fitData\(\);/);
  assert.equal((client.match(/map\.invalidateSize/g)||[]).length,1);
});

test('地図候補マーカーと左一覧は共通の候補選択処理を使う', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  assert.match(client,/data-candidate[^\n]+selectCandidate\(el\.dataset\.candidate\)/);
  assert.match(client,/marker\.on\('click',[\s\S]*?selectCandidate\(c\.candidate_id\)/);
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

test('候補マーカー選択は未保存変更保護を通り、正式point選択を維持する', () => {
  const client=fs.readFileSync(path.join(root,'Client.html'),'utf8');
  const selectCandidateBody=client.match(/function selectCandidate\(candidateId\) \{([\s\S]*?)\n  \}\n\n  function scrollSelectedCandidateIntoView/);
  assert.ok(selectCandidateBody);
  assert.match(selectCandidateBody[1],/if \(!canLeave\(\)\) return/);
  assert.match(client,/marker\.on\('click', \(\) => point\.place_id && selectPlace\(point\.place_id, point\.point_id\)\)/);
  assert.match(client,/candidateRead\('candidate_id',c\.candidate_id\)/);
  assert.match(client,/String\(a\['地点名'\]\)\.localeCompare\(String\(b\['地点名'\]\),'ja'\)/);
});
