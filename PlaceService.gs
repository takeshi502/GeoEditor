function getGeoEditorBootstrap_() {
  var ss = spreadsheet_();
  var places = readTable_(ss, 'places');
  var points = readTable_(ss, 'points');
  var candidates = readTable_(ss, 'candidates');
  var baseAreas = readTable_(ss, 'baseAreas');
  return {
    data: {
      places: places.rows.map(publicItem_),
      points: points.rows.map(publicItem_),
      candidates: candidates.rows.map(publicItem_),
      base_areas: baseAreas.rows.map(publicItem_),
      schema: schemaInfo_()
    }
  };
}

function getPlaceDetail_(placeId) {
  placeId = requiredString_(placeId, 'place_id');
  var ss = spreadsheet_();
  var places = readTable_(ss, 'places');
  var points = readTable_(ss, 'points');
  var place = findById_(places.rows, 'place_id', placeId);
  if (!place) throw geoError_('NOT_FOUND', '地点が見つかりません: ' + placeId);
  return {
    data: {
      place: publicItem_(place),
      points: points.rows.filter(function (point) { return point.place_id === placeId; }).map(publicItem_),
      revision: place.revision,
      updated_at: place['更新日時']
    },
    server_revision: place.revision
  };
}

function createPlaceBundle_(payload) {
  return withWriteLock_(function () {
    var ss = spreadsheet_();
    var placesTable = readTable_(ss, 'places');
    var pointsTable = readTable_(ss, 'points');
    return createPlaceBundleWithTables_(payload, placesTable, pointsTable);
  });
}

function createPlaceBundleWithTables_(payload, placesTable, pointsTable) {
  payload = payload || {};
  var sourcePlace = payload.place || {};
  var sourcePoints = payload.points || [];
  if (!Array.isArray(sourcePoints) || !sourcePoints.length) {
    throw geoError_('VALIDATION_ERROR', '新規地点には初期判定ポイントが1件以上必要です。');
  }
  validatePlace_(sourcePlace, true);
  assertUniqueIds_(placesTable.rows, 'place_id');
  assertUniqueIds_(pointsTable.rows, 'point_id');
  var now = nowIso_();
  var placeId = id_('place');
  var place = normalizePlace_(sourcePlace, placeId, now, true);
  var points = sourcePoints.map(function (source) {
    validatePoint_(source, placeId, true);
    return normalizePoint_(source, id_('point'), placeId, now, true);
  });
  assertUniqueIds_(points, 'point_id');
  var warnings = overlapWarnings_(points, pointsTable.rows);
  var placeRow = placesTable.sheet.getLastRow() + 1;
  var pointStartRow = pointsTable.sheet.getLastRow() + 1;
  try {
    placesTable.sheet.getRange(placeRow, 1, 1, placesTable.headers.length)
      .setValues([objectToRow_(place, placesTable.headers)]);
    pointsTable.sheet.getRange(pointStartRow, 1, points.length, pointsTable.headers.length)
      .setValues(points.map(function (point) { return objectToRow_(point, pointsTable.headers); }));
  } catch (error) {
    try { placesTable.sheet.getRange(placeRow, 1, 1, placesTable.headers.length).clearContent(); } catch (ignored) {}
    try { pointsTable.sheet.getRange(pointStartRow, 1, points.length, pointsTable.headers.length).clearContent(); } catch (ignored2) {}
    throw error;
  }
  return { data: { place: place, points: points }, warnings: warnings, server_revision: 1 };
}

function savePlaceBundle_(payload) {
  return withWriteLock_(function () {
    payload = payload || {};
    var incomingPlace = payload.place || {};
    var incomingPoints = payload.points || [];
    validatePlace_(incomingPlace, false);
    if (!Array.isArray(incomingPoints)) throw geoError_('VALIDATION_ERROR', 'points は配列で指定してください。');
    var ss = spreadsheet_();
    var placesTable = readTable_(ss, 'places');
    var pointsTable = readTable_(ss, 'points');
    assertUniqueIds_(placesTable.rows, 'place_id');
    assertUniqueIds_(pointsTable.rows, 'point_id');
    assertUniqueIds_(incomingPoints, 'point_id');
    var currentPlace = findById_(placesTable.rows, 'place_id', incomingPlace.place_id);
    if (!currentPlace) throw geoError_('NOT_FOUND', '地点が見つかりません: ' + incomingPlace.place_id);
    assertRevision_(incomingPlace.revision, currentPlace.revision, '地点', currentPlace.revision);

    var now = nowIso_();
    var placeChanged = !placeEditableEquals_(incomingPlace, currentPlace);
    var place = placeChanged ? normalizePlace_(incomingPlace, currentPlace.place_id, now, false, currentPlace) : publicItem_(currentPlace);
    if (placeChanged && String(place['地点名']) !== String(currentPlace['地点名']) && payload.preserve_old_name !== false) {
      place['別名'] = mergeAlias_(place['別名'], currentPlace['地点名']);
    }
    var existingForPlace = pointsTable.rows.filter(function (point) { return point.place_id === place.place_id; });
    var existingById = indexBy_(existingForPlace, 'point_id');
    var points = incomingPoints.map(function (incoming) {
      var isNew = !incoming.point_id;
      validatePoint_(incoming, place.place_id, isNew);
      if (!isNew) {
        var current = existingById[incoming.point_id];
        if (!current) throw geoError_('NOT_FOUND', '判定ポイントが見つかりません: ' + incoming.point_id);
        assertRevision_(incoming.revision, current.revision, '判定ポイント', current.revision);
        if (pointEditableEquals_(incoming, current)) return publicItem_(current);
        return normalizePoint_(incoming, current.point_id, place.place_id, now, false, current);
      }
      return normalizePoint_(incoming, id_('point'), place.place_id, now, true);
    });
    assertUniqueIds_(points, 'point_id');
    if (!points.some(function (point) { return point['有効'] === true; })) {
      throw geoError_('VALIDATION_ERROR', '地点には有効な判定ポイントが1件以上必要です。');
    }
    var outsidePoints = pointsTable.rows.filter(function (point) { return point.place_id !== place.place_id; });
    var warnings = overlapWarnings_(points, outsidePoints);

    var appendedRows = [];
    try {
      if (placeChanged) {
        placesTable.sheet.getRange(currentPlace._row, 1, 1, placesTable.headers.length)
          .setValues([objectToRow_(place, placesTable.headers)]);
      }
      points.forEach(function (point) {
        var current = existingById[point.point_id];
        if (current) {
          if (Number(point.revision) !== Number(current.revision)) {
            pointsTable.sheet.getRange(current._row, 1, 1, pointsTable.headers.length)
              .setValues([objectToRow_(point, pointsTable.headers)]);
          }
        } else {
          var row = pointsTable.sheet.getLastRow() + 1;
          pointsTable.sheet.appendRow(objectToRow_(point, pointsTable.headers));
          appendedRows.push(row);
        }
      });
    } catch (error) {
      try {
        if (placeChanged) {
          placesTable.sheet.getRange(currentPlace._row, 1, 1, placesTable.headers.length)
            .setValues([objectToRow_(currentPlace, placesTable.headers)]);
        }
        existingForPlace.forEach(function (point) {
          pointsTable.sheet.getRange(point._row, 1, 1, pointsTable.headers.length)
            .setValues([objectToRow_(point, pointsTable.headers)]);
        });
        appendedRows.forEach(function (row) {
          pointsTable.sheet.getRange(row, 1, 1, pointsTable.headers.length).clearContent();
        });
      } catch (rollbackError) { console.error('Rollback failed: ' + rollbackError.message); }
      throw error;
    }
    return { data: { place: place, points: points }, warnings: warnings, server_revision: place.revision };
  });
}

function normalizePlace_(source, placeId, now, isNew, current) {
  return {
    place_id: placeId,
    '地点名': String(source['地点名']).trim(),
    base_area_id: source.base_area_id || '',
    '優先度': Number(source['優先度'] || 0),
    '別名': aliases_(source['別名']).join(', '),
    '登録元': isNew ? (source['登録元'] || 'geo_editor') : (current['登録元'] || source['登録元'] || 'geo_editor'),
    '最終使用日時': isNew ? (source['最終使用日時'] || '') : current['最終使用日時'],
    '有効': source['有効'],
    '備考': source['備考'] || '',
    '作成日時': isNew ? now : current['作成日時'],
    '更新日時': now,
    revision: isNew ? 1 : Number(current.revision) + 1
  };
}

function normalizePoint_(source, pointId, placeId, now, isNew, current) {
  return {
    point_id: pointId,
    place_id: placeId,
    'ポイント名': String(source['ポイント名']).trim(),
    '表示順': Number(source['表示順']),
    '中心緯度': Number(source['中心緯度']),
    '中心経度': Number(source['中心経度']),
    '判定半径_m': Number(source['判定半径_m']),
    '最大GPS精度_m': Number(source['最大GPS精度_m']),
    '優先度': Number(source['優先度']),
    '最終使用日時': isNew ? (source['最終使用日時'] || '') : current['最終使用日時'],
    '登録元': isNew ? (source['登録元'] || 'geo_editor') : (current['登録元'] || source['登録元'] || 'geo_editor'),
    '有効': source['有効'],
    '備考': source['備考'] || '',
    '作成日時': isNew ? now : current['作成日時'],
    '更新日時': now,
    revision: isNew ? 1 : Number(current.revision) + 1
  };
}

function pointEditableEquals_(incoming, current) {
  return String(incoming['ポイント名']).trim() === String(current['ポイント名']).trim() &&
    Number(incoming['表示順']) === Number(current['表示順']) &&
    Number(incoming['中心緯度']) === Number(current['中心緯度']) &&
    Number(incoming['中心経度']) === Number(current['中心経度']) &&
    Number(incoming['判定半径_m']) === Number(current['判定半径_m']) &&
    Number(incoming['最大GPS精度_m']) === Number(current['最大GPS精度_m']) &&
    Number(incoming['優先度']) === Number(current['優先度']) &&
    incoming['有効'] === current['有効'] &&
    String(incoming['備考'] || '') === String(current['備考'] || '');
}

function placeEditableEquals_(incoming, current) {
  return String(incoming['地点名']).trim() === String(current['地点名']).trim() &&
    String(incoming.base_area_id || '') === String(current.base_area_id || '') &&
    Number(incoming['優先度'] || 0) === Number(current['優先度'] || 0) &&
    aliases_(incoming['別名']).join(', ') === aliases_(current['別名']).join(', ') &&
    incoming['有効'] === current['有効'] &&
    String(incoming['備考'] || '') === String(current['備考'] || '');
}

function assertRevision_(incoming, current, label, serverRevision) {
  if (Number(incoming) !== Number(current)) {
    throw geoError_('REVISION_CONFLICT', label + 'は別の処理で更新されています。再読込してください。', { serverRevision: serverRevision });
  }
}

function findById_(rows, key, value) {
  for (var i = 0; i < rows.length; i += 1) if (rows[i][key] === value) return rows[i];
  return null;
}

function indexBy_(rows, key) {
  var result = {};
  rows.forEach(function (row) { result[row[key]] = row; });
  return result;
}
