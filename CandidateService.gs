function resolveCandidate_(payload) {
  return withWriteLock_(function () {
    payload = payload || {};
    var action = payload.action;
    if (['promote', 'merge', 'add_point', 'reject'].indexOf(action) < 0) {
      throw geoError_('VALIDATION_ERROR', '未対応の候補処理です: ' + action);
    }
    var ss = spreadsheet_();
    var candidatesTable = readTable_(ss, 'candidates');
    var placesTable = readTable_(ss, 'places');
    var pointsTable = readTable_(ss, 'points');
    var baseAreasTable = readTable_(ss, 'baseAreas');
    assertUniqueIds_(candidatesTable.rows, 'candidate_id');
    assertUniqueIds_(placesTable.rows, 'place_id');
    assertUniqueIds_(pointsTable.rows, 'point_id');
    assertUniqueIds_(baseAreasTable.rows, 'base_area_id');
    var candidate = findById_(candidatesTable.rows, 'candidate_id', requiredString_(payload.candidate_id, 'candidate_id'));
    if (!candidate) throw geoError_('NOT_FOUND', '学習候補が見つかりません。');
    assertRevision_(payload.revision, candidate.revision, '学習候補', candidate.revision);
    if (['promoted', 'merged', 'rejected'].indexOf(candidate['候補状態']) >= 0) {
      throw geoError_('VALIDATION_ERROR', 'この学習候補は処理済みです。');
    }

    var result;
    if (action === 'promote') {
      var promotePosition = candidateRegistrationPosition_(payload);
      var bundle = payload.bundle || {};
      bundle.place = bundle.place || {};
      if (!bundle.place['地点名']) bundle.place['地点名'] = candidate['候補名'];
      bundle.place.base_area_id = candidatePromotionBaseArea_(payload, bundle.place, candidate, baseAreasTable.rows);
      bundle.points = bundle.points && bundle.points.length ? bundle.points : [candidatePoint_(candidate, payload.point || {})];
      applyCandidateRegistrationPosition_(bundle.points[0], promotePosition);
      var placeStart = placesTable.sheet.getLastRow() + 1;
      var pointStart = pointsTable.sheet.getLastRow() + 1;
      result = createPlaceBundleWithTables_(bundle, placesTable, pointsTable);
      try {
        updateCandidate_(candidatesTable, candidate, 'promoted', result.data.place.place_id);
      } catch (promoteError) {
        try {
          placesTable.sheet.getRange(placeStart, 1, 1, placesTable.headers.length).clearContent();
          pointsTable.sheet.getRange(pointStart, 1, result.data.points.length, pointsTable.headers.length).clearContent();
          candidatesTable.sheet.getRange(candidate._row, 1, 1, candidatesTable.headers.length)
            .setValues([objectToRow_(candidate, candidatesTable.headers)]);
        } catch (rollbackError) { console.error('Rollback failed: ' + rollbackError.message); }
        throw promoteError;
      }
    } else if (action === 'reject') {
      updateCandidate_(candidatesTable, candidate, 'rejected', '');
      result = { data: { candidate_id: candidate.candidate_id, status: 'rejected' }, warnings: [] };
    } else {
      var placeId = requiredString_(payload.place_id, '登録先place_id');
      var place = findById_(placesTable.rows, 'place_id', placeId);
      if (!place) throw geoError_('NOT_FOUND', '登録先地点が見つかりません。');
      assertRevision_(payload.place_revision, place.revision, '登録先地点', place.revision);
      var now = nowIso_();
      var updatedPlace = normalizePlace_(place, place.place_id, now, false, place);
      if (payload.add_candidate_alias === true) updatedPlace['別名'] = mergeAlias_(updatedPlace['別名'], candidate['候補名']);
      var warnings = [];
      var createdPoint = null;
      if (action === 'add_point') {
        var sourcePoint = candidatePoint_(candidate, payload.point || {});
        applyCandidateRegistrationPosition_(sourcePoint, candidateRegistrationPosition_(payload));
        validatePoint_(sourcePoint, placeId, true);
        createdPoint = normalizePoint_(sourcePoint, id_('point'), placeId, now, true);
        warnings = overlapWarnings_([createdPoint], pointsTable.rows.filter(function (point) { return point.place_id !== placeId; }));
      }
      var appendedPointRow = createdPoint ? pointsTable.sheet.getLastRow() + 1 : null;
      try {
        placesTable.sheet.getRange(place._row, 1, 1, placesTable.headers.length)
          .setValues([objectToRow_(updatedPlace, placesTable.headers)]);
        if (createdPoint) pointsTable.sheet.appendRow(objectToRow_(createdPoint, pointsTable.headers));
        updateCandidate_(candidatesTable, candidate, 'merged', placeId);
      } catch (mergeError) {
        try {
          placesTable.sheet.getRange(place._row, 1, 1, placesTable.headers.length)
            .setValues([objectToRow_(place, placesTable.headers)]);
          candidatesTable.sheet.getRange(candidate._row, 1, 1, candidatesTable.headers.length)
            .setValues([objectToRow_(candidate, candidatesTable.headers)]);
          if (appendedPointRow) pointsTable.sheet.getRange(appendedPointRow, 1, 1, pointsTable.headers.length).clearContent();
        } catch (rollbackError) { console.error('Rollback failed: ' + rollbackError.message); }
        throw mergeError;
      }
      result = { data: { place: updatedPlace, point: createdPoint, candidate_id: candidate.candidate_id }, warnings: warnings, server_revision: updatedPlace.revision };
    }
    return result;
  });
}

function candidatePromotionBaseArea_(payload, place, candidate, baseAreas) {
  var hasExplicitSelection = Object.prototype.hasOwnProperty.call(payload || {}, 'registration_base_area_id');
  var selected = hasExplicitSelection
    ? payload.registration_base_area_id
    : (place && place.base_area_id ? place.base_area_id : candidate['推定base_area_id']);
  if (selected == null || selected === '') return '';
  if (typeof selected !== 'string') throw geoError_('VALIDATION_ERROR', '所属基地の指定が不正です。');
  selected = selected.trim();
  if (!selected) return '';
  var baseArea = findById_(baseAreas || [], 'base_area_id', selected);
  if (!baseArea || baseArea['有効'] === false) {
    throw geoError_('VALIDATION_ERROR', '選択された所属基地は有効な基地マスタに存在しません。');
  }
  return selected;
}

function candidateRegistrationPosition_(payload) {
  var position = payload && payload.registration_position;
  if (position == null) return null;
  if (typeof position !== 'object') throw geoError_('VALIDATION_ERROR', '登録予定位置の形式が不正です。');
  if (position.latitude == null || position.latitude === '' || position.longitude == null || position.longitude === '') {
    throw geoError_('VALIDATION_ERROR', '登録予定位置の緯度・経度は必須です。');
  }
  return {
    latitude: finiteNumber_(position.latitude, '登録予定位置の緯度', -90, 90),
    longitude: finiteNumber_(position.longitude, '登録予定位置の経度', -180, 180)
  };
}

function applyCandidateRegistrationPosition_(point, position) {
  if (!position) return point;
  point['中心緯度'] = position.latitude;
  point['中心経度'] = position.longitude;
  return point;
}

function candidatePoint_(candidate, overrides) {
  return {
    'ポイント名': overrides['ポイント名'] || 'メイン',
    '表示順': overrides['表示順'] == null ? 1 : overrides['表示順'],
    '中心緯度': overrides['中心緯度'] == null ? candidate['基準緯度'] : overrides['中心緯度'],
    '中心経度': overrides['中心経度'] == null ? candidate['基準経度'] : overrides['中心経度'],
    '判定半径_m': overrides['判定半径_m'] == null ? 50 : overrides['判定半径_m'],
    '最大GPS精度_m': overrides['最大GPS精度_m'] == null ? Math.min(50, Number(overrides['判定半径_m'] || 50)) : overrides['最大GPS精度_m'],
    '優先度': overrides['優先度'] == null ? 100 : overrides['優先度'],
    '登録元': 'learning_candidate',
    '有効': overrides['有効'] === false ? false : true,
    '備考': overrides['備考'] || ''
  };
}

function updateCandidate_(table, current, status, placeId) {
  var updated = publicItem_(current);
  updated['候補状態'] = status;
  updated['登録先place_id'] = placeId || '';
  updated.revision = Number(current.revision) + 1;
  table.sheet.getRange(current._row, 1, 1, table.headers.length)
    .setValues([objectToRow_(updated, table.headers)]);
  return updated;
}
