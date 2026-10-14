function disablePlace_(payload) {
  return withWriteLock_(function () {
    payload = payload || {};
    var placeId = requiredString_(payload.place_id, 'place_id');
    var pointRevisions = payload.point_revisions;
    if (!Array.isArray(pointRevisions)) {
      throw geoError_('VALIDATION_ERROR', 'point_revisions は配列で指定してください。');
    }

    var ss = spreadsheet_();
    var placesTable = readTable_(ss, 'places');
    var pointsTable = readTable_(ss, 'points');
    assertUniqueIds_(placesTable.rows, 'place_id');
    assertUniqueIds_(pointsTable.rows, 'point_id');
    var currentPlace = findById_(placesTable.rows, 'place_id', placeId);
    if (!currentPlace) throw geoError_('NOT_FOUND', '地点が見つかりません: ' + placeId);
    var currentPoints = pointsTable.rows.filter(function (point) { return point.place_id === placeId; });

    var activeCurrentPoints = currentPoints.filter(function (point) { return point['有効'] !== false; });
    if (currentPlace['有効'] === false && !activeCurrentPoints.length) {
      return {
        data: { place:publicItem_(currentPlace), points:currentPoints.map(publicItem_), disabled_point_count:0, already_disabled:true },
        server_revision: currentPlace.revision
      };
    }

    if (currentPlace['有効'] !== false) {
      assertRevision_(payload.place_revision, currentPlace.revision, '地点', currentPlace.revision);
    }
    var revisions = {};
    pointRevisions.forEach(function (item) {
      if (item && item.point_id) revisions[String(item.point_id)] = item.revision;
    });
    currentPoints.forEach(function (point) {
      if (!Object.prototype.hasOwnProperty.call(revisions, point.point_id)) {
        throw geoError_('REVISION_CONFLICT', '乗降位置の状態が更新されています。再読込してください。', { serverRevision:point.revision });
      }
      assertRevision_(revisions[point.point_id], point.revision, '乗降位置', point.revision);
    });

    var now = nowIso_();
    var placeChanged = currentPlace['有効'] !== false;
    var disabledPlace = publicItem_(currentPlace);
    if (placeChanged) {
      var placeInput = publicItem_(currentPlace);
      placeInput['有効'] = false;
      disabledPlace = normalizePlace_(placeInput, placeId, now, false, currentPlace);
    }
    var disabledPointCount = 0;
    var disabledPoints = currentPoints.map(function (current) {
      if (current['有効'] === false) return publicItem_(current);
      var incoming = publicItem_(current);
      incoming['有効'] = false;
      disabledPointCount += 1;
      return normalizePoint_(incoming, current.point_id, placeId, now, false, current);
    });

    var placeSnapshot = publicItem_(currentPlace);
    var pointSnapshots = currentPoints.map(publicItem_);
    try {
      if (placeChanged) {
        placesTable.sheet.getRange(currentPlace._row, 1, 1, placesTable.headers.length)
          .setValues([objectToRow_(disabledPlace, placesTable.headers)]);
      }
      disabledPoints.forEach(function (point) {
        var current = findById_(currentPoints, 'point_id', point.point_id);
        if (Number(point.revision) !== Number(current.revision)) {
          pointsTable.sheet.getRange(current._row, 1, 1, pointsTable.headers.length)
            .setValues([objectToRow_(point, pointsTable.headers)]);
        }
      });
    } catch (error) {
      try {
        if (placeChanged) {
          placesTable.sheet.getRange(currentPlace._row, 1, 1, placesTable.headers.length)
            .setValues([objectToRow_(placeSnapshot, placesTable.headers)]);
        }
        currentPoints.forEach(function (point, index) {
          pointsTable.sheet.getRange(point._row, 1, 1, pointsTable.headers.length)
            .setValues([objectToRow_(pointSnapshots[index], pointsTable.headers)]);
        });
      } catch (rollbackError) { console.error('Place disable rollback failed: ' + rollbackError.message); }
      throw error;
    }

    return {
      data: { place:disabledPlace, points:disabledPoints, disabled_point_count:disabledPointCount, already_disabled:false },
      server_revision: disabledPlace.revision
    };
  });
}
