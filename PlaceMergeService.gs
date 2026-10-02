function mergePlaces_(payload) {
  return withWriteLock_(function () {
    payload = payload || {};
    var sourceId = requiredString_(payload.source_place_id, 'まとめる地点');
    var targetId = requiredString_(payload.target_place_id, '残す地点');
    if (sourceId === targetId) throw geoError_('VALIDATION_ERROR', '同じ地点どうしはまとめられません。');

    var ss = spreadsheet_();
    var placesTable = readTable_(ss, 'places');
    var pointsTable = readTable_(ss, 'points');
    assertUniqueIds_(placesTable.rows, 'place_id');
    assertUniqueIds_(pointsTable.rows, 'point_id');
    var source = findById_(placesTable.rows, 'place_id', sourceId);
    var target = findById_(placesTable.rows, 'place_id', targetId);
    if (!source || !target) throw geoError_('NOT_FOUND', '統合対象の地点が見つかりません。');
    if (source['有効'] === false || target['有効'] === false) throw geoError_('VALIDATION_ERROR', '無効な地点は統合対象にできません。');
    assertRevision_(payload.source_revision, source.revision, 'まとめる地点', source.revision);
    assertRevision_(payload.target_revision, target.revision, '残す地点', target.revision);

    var relevantPoints = pointsTable.rows.filter(function (point) {
      return point.place_id === sourceId || point.place_id === targetId;
    });
    var revisions = {};
    (payload.point_revisions || []).forEach(function (item) { revisions[item.point_id] = item.revision; });
    relevantPoints.forEach(function (point) {
      if (!Object.prototype.hasOwnProperty.call(revisions, point.point_id)) {
        throw geoError_('REVISION_CONFLICT', '乗降位置の状態が更新されています。再読込してください。', { serverRevision: point.revision });
      }
      assertRevision_(revisions[point.point_id], point.revision, '乗降位置', point.revision);
    });

    var now = nowIso_();
    var updatedTarget = normalizePlace_(target, targetId, now, false, target);
    if (payload.preserve_source_name !== false && String(source['地点名']).trim() !== String(target['地点名']).trim()) {
      updatedTarget['別名'] = mergeAlias_(updatedTarget['別名'], source['地点名']);
    }
    var updatedSourceInput = publicItem_(source);
    updatedSourceInput['有効'] = false;
    updatedSourceInput['備考'] = appendAuditNote_(updatedSourceInput['備考'], targetId + 'へ統合', now);
    var updatedSource = normalizePlace_(updatedSourceInput, sourceId, now, false, source);

    var targetPoints = relevantPoints.filter(function (point) { return point.place_id === targetId; })
      .sort(pointOrder_);
    var sourcePoints = relevantPoints.filter(function (point) { return point.place_id === sourceId; })
      .sort(pointOrder_);
    var mergedPoints = targetPoints.concat(sourcePoints).map(function (current, index) {
      var incoming = publicItem_(current);
      incoming.place_id = targetId;
      incoming['表示順'] = index + 1;
      if (incoming.place_id === current.place_id && Number(incoming['表示順']) === Number(current['表示順'])) return publicItem_(current);
      return normalizePoint_(incoming, current.point_id, targetId, now, false, current);
    });
    var outside = pointsTable.rows.filter(function (point) { return point.place_id !== sourceId && point.place_id !== targetId; });
    var warnings = overlapWarnings_(mergedPoints.filter(function (point) { return point['有効'] !== false; }), outside);

    var placeSnapshots = [publicItem_(target), publicItem_(source)];
    var pointSnapshots = relevantPoints.map(publicItem_);
    try {
      placesTable.sheet.getRange(target._row, 1, 1, placesTable.headers.length)
        .setValues([objectToRow_(updatedTarget, placesTable.headers)]);
      placesTable.sheet.getRange(source._row, 1, 1, placesTable.headers.length)
        .setValues([objectToRow_(updatedSource, placesTable.headers)]);
      mergedPoints.forEach(function (point) {
        var current = findById_(relevantPoints, 'point_id', point.point_id);
        if (point.place_id !== current.place_id || Number(point['表示順']) !== Number(current['表示順'])) {
          pointsTable.sheet.getRange(current._row, 1, 1, pointsTable.headers.length)
            .setValues([objectToRow_(point, pointsTable.headers)]);
        }
      });
    } catch (error) {
      try {
        placesTable.sheet.getRange(target._row, 1, 1, placesTable.headers.length)
          .setValues([objectToRow_(placeSnapshots[0], placesTable.headers)]);
        placesTable.sheet.getRange(source._row, 1, 1, placesTable.headers.length)
          .setValues([objectToRow_(placeSnapshots[1], placesTable.headers)]);
        relevantPoints.forEach(function (point, index) {
          pointsTable.sheet.getRange(point._row, 1, 1, pointsTable.headers.length)
            .setValues([objectToRow_(pointSnapshots[index], pointsTable.headers)]);
        });
      } catch (rollbackError) { console.error('Place merge rollback failed: ' + rollbackError.message); }
      throw error;
    }
    return {
      data: { place: updatedTarget, merged_place: updatedSource, points: mergedPoints },
      warnings: warnings,
      server_revision: updatedTarget.revision
    };
  });
}

function pointOrder_(a, b) {
  return Number(a['表示順']) - Number(b['表示順']) || String(a.point_id).localeCompare(String(b.point_id));
}

function appendAuditNote_(note, message, now) {
  var entry = '[' + now + '] ' + message;
  return note ? String(note).trim() + '\n' + entry : entry;
}
