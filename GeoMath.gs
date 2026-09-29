var EARTH_RADIUS_M = 6371008.8;

function distanceMeters_(lat1, lon1, lat2, lon2) {
  var toRad = Math.PI / 180;
  var phi1 = Number(lat1) * toRad;
  var phi2 = Number(lat2) * toRad;
  var dPhi = (Number(lat2) - Number(lat1)) * toRad;
  var dLambda = (Number(lon2) - Number(lon1)) * toRad;
  var a = Math.sin(dPhi / 2) * Math.sin(dPhi / 2) +
    Math.cos(phi1) * Math.cos(phi2) * Math.sin(dLambda / 2) * Math.sin(dLambda / 2);
  return 2 * EARTH_RADIUS_M * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function overlapMeters_(pointA, pointB) {
  var distance = distanceMeters_(pointA['中心緯度'], pointA['中心経度'], pointB['中心緯度'], pointB['中心経度']);
  return Math.max(0, Number(pointA['判定半径_m']) + Number(pointB['判定半径_m']) - distance);
}

function circlesOverlap_(pointA, pointB) {
  return overlapMeters_(pointA, pointB) > 0;
}

function radiusSuggestion_(point, allPoints) {
  var constraints = (allPoints || []).filter(function (other) {
    return other['有効'] !== false && other.place_id !== point.place_id && other.point_id !== point.point_id;
  }).map(function (other) {
    var distance = distanceMeters_(point['中心緯度'], point['中心経度'], other['中心緯度'], other['中心経度']);
    return {
      point_id: other.point_id,
      place_id: other.place_id,
      distance_m: distance,
      max_radius_m: distance - Number(other['判定半径_m']) - 5
    };
  }).sort(function (a, b) { return a.max_radius_m - b.max_radius_m; });

  var raw = constraints.length ? Math.min(50, constraints[0].max_radius_m) : 50;
  var rounded = Math.max(5, Math.floor(raw / 5) * 5);
  return {
    radius_m: rounded,
    max_accuracy_m: Math.min(50, rounded),
    raw_radius_m: raw,
    density: rounded <= 15 ? 'critical' : rounded < 20 ? 'warning' : 'normal',
    nearest_constraints: constraints.slice(0, 5)
  };
}

function overlapWarnings_(points, comparisonPoints) {
  var warnings = [];
  (points || []).forEach(function (point) {
    if (point['有効'] === false) return;
    (comparisonPoints || []).forEach(function (other) {
      if (other['有効'] === false || point.place_id === other.place_id || point.point_id === other.point_id) return;
      var overlap = overlapMeters_(point, other);
      if (overlap > 0) warnings.push({
        code: 'CIRCLE_OVERLAP',
        point_id: point.point_id,
        other_point_id: other.point_id,
        other_place_id: other.place_id,
        overlap_m: Math.round(overlap * 10) / 10,
        message: '異なる地点の判定円が ' + (Math.round(overlap * 10) / 10) + 'm 重複しています。specific place判定が曖昧になる可能性があります。'
      });
    });
  });
  return warnings;
}
