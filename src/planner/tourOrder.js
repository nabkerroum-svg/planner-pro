/**
 * Ordre des chantiers sur une journée : proximité + point de départ / arrivée optionnels.
 */

import { haversine } from "./transport.js";

export function sortByCoords(siteList, coords) {
  if (siteList.length <= 1) return siteList;
  var withC = siteList.filter(function (s) {
    return coords[s.id];
  });
  var withoutC = siteList.filter(function (s) {
    return !coords[s.id];
  });
  if (!withC.length) return siteList;

  var visited = new Set();
  var ordered = [];
  var cur = withC[0];
  visited.add(cur.id);
  ordered.push(cur);

  while (ordered.length < withC.length) {
    var near = null;
    var minD = Infinity;
    for (var i = 0; i < withC.length; i++) {
      var s = withC[i];
      if (visited.has(s.id)) continue;
      var d = haversine(
        coords[cur.id].lat,
        coords[cur.id].lon,
        coords[s.id].lat,
        coords[s.id].lon
      );
      if (d < minD) {
        minD = d;
        near = s;
      }
    }
    if (!near) break;
    visited.add(near.id);
    ordered.push(near);
    cur = near;
  }
  return ordered.concat(withoutC);
}

function nearestNeighbourFromSeed(pool, coords, lat0, lon0) {
  if (!pool.length) return [];
  var withCoords = pool.filter(function (s) {
    return coords[s.id];
  });
  var withoutCoords = pool.filter(function (s) {
    return !coords[s.id];
  });
  if (!withCoords.length) return pool.slice();

  var visited = new Set();
  var ordered = [];
  var cur = null;
  var minD = Infinity;
  for (var i0 = 0; i0 < withCoords.length; i0++) {
    var s0 = withCoords[i0];
    var d0 = haversine(lat0, lon0, coords[s0.id].lat, coords[s0.id].lon);
    if (d0 < minD) {
      minD = d0;
      cur = s0;
    }
  }
  if (!cur) return sortByCoords(pool, coords);
  visited.add(cur.id);
  ordered.push(cur);

  while (ordered.length < withCoords.length) {
    var near = null;
    minD = Infinity;
    for (var i = 0; i < withCoords.length; i++) {
      var s = withCoords[i];
      if (visited.has(s.id)) continue;
      var d = haversine(
        coords[cur.id].lat,
        coords[cur.id].lon,
        coords[s.id].lat,
        coords[s.id].lon
      );
      if (d < minD) {
        minD = d;
        near = s;
      }
    }
    if (!near) break;
    visited.add(near.id);
    ordered.push(near);
    cur = near;
  }
  return ordered.concat(withoutCoords);
}

/**
 * @param {object} anchors — startSiteId, endSiteId, startSeedCoords?: { lat, lon }
 */
export function orderSitesWithAnchors(siteList, coords, anchors) {
  if (!siteList.length) return [];
  anchors = anchors || {};
  var startSiteId = anchors.startSiteId || "";
  var endSiteId = anchors.endSiteId || "";
  var seed = anchors.startSeedCoords;

  var start = startSiteId ? siteList.find(function (s) { return s.id === startSiteId; }) : null;
  var end = endSiteId ? siteList.find(function (s) { return s.id === endSiteId; }) : null;

  var pool = siteList.slice();
  if (start) pool = pool.filter(function (s) { return s.id !== start.id; });
  if (end) pool = pool.filter(function (s) { return s.id !== end.id; });

  var orderedMiddle;
  if (start && coords[start.id]) {
    orderedMiddle = nearestNeighbourFromSeed(pool, coords, coords[start.id].lat, coords[start.id].lon);
  } else if (seed && typeof seed.lat === "number" && typeof seed.lon === "number") {
    orderedMiddle = nearestNeighbourFromSeed(pool, coords, seed.lat, seed.lon);
  } else {
    orderedMiddle = sortByCoords(pool, coords);
  }

  var out = [];
  if (start) out.push(start);
  for (var j = 0; j < orderedMiddle.length; j++) out.push(orderedMiddle[j]);
  if (end && (!start || end.id !== start.id)) {
    var last = out[out.length - 1];
    if (!last || last.id !== end.id) out.push(end);
  }
  return out;
}

export async function sortByProximity(siteList, batchGeocode) {
  if (siteList.length <= 1) return siteList;
  var coords = await batchGeocode(siteList);
  return sortByCoords(siteList, coords);
}
