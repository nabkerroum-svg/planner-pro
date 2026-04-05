/**
 * Géocodage, routage OSRM et durées de trajet — modes voiture / TC / vélo / à pied.
 */

const LS_GEO = "planner_geo_v1";
const LS_ROUTE = "planner_route_v1";

function lsGet(k, def) {
  try {
    const r = localStorage.getItem(k);
    return r ? JSON.parse(r) : def;
  } catch (e) {
    return def;
  }
}
function lsSet(k, v) {
  try {
    localStorage.setItem(k, JSON.stringify(v));
  } catch (e) {}
}

let geoCache = lsGet(LS_GEO, {});
let routeCache = lsGet(LS_ROUTE, {});

/** Facteurs appliqués à la durée OSRM « driving » (vélo / TC). À pied : profil foot dédié. */
export const TRANSPORT_FACTORS = { car: 1, pt: 1.8, bike: 1.4, walk: 1 };

export function transportEmoji(mode) {
  if (mode === "bike") return "\uD83D\uDEB2";
  if (mode === "pt") return "\uD83D\uDE8C";
  if (mode === "walk") return "\uD83D\uDEB6";
  return "\uD83D\uDE97";
}

export function transportLabelFr(mode) {
  if (mode === "bike") return "Velo";
  if (mode === "pt") return "Transports en commun";
  if (mode === "walk") return "A pied";
  return "Voiture";
}

export function haversine(aLat, aLon, bLat, bLon) {
  const R = 6371;
  const dLat = ((bLat - aLat) * Math.PI) / 180;
  const dLon = ((bLon - aLon) * Math.PI) / 180;
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos((aLat * Math.PI) / 180) *
      Math.cos((bLat * Math.PI) / 180) *
      Math.sin(dLon / 2) *
      Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

export function siteHasStoredCoords(site) {
  if (!site || site.lat == null || site.lon == null) return false;
  const lat = Number(site.lat);
  const lon = Number(site.lon);
  return Number.isFinite(lat) && Number.isFinite(lon);
}

export function seedGeoCacheForAddress(address, lat, lon) {
  const q = address && String(address).trim();
  if (!q || lat == null || lon == null) return;
  const c = { lat: Number(lat), lon: Number(lon) };
  geoCache[q] = c;
  lsSet(LS_GEO, geoCache);
}

export async function geocodeAddress(addr) {
  const q = addr && addr.trim();
  if (!q) throw new Error("Adresse vide");
  if (geoCache[q]) return geoCache[q];
  const url =
    "https://nominatim.openstreetmap.org/search?format=json&q=" +
    encodeURIComponent(q) +
    "&limit=1";
  const res = await fetch(url, {
    headers: {
      Accept: "application/json",
      "Accept-Language": "fr",
      "User-Agent": "PlannerPro/1.0 (geocodage chantiers)",
    },
  });
  if (!res.ok) throw new Error("Nominatim " + res.status);
  const data = await res.json();
  if (!data || !data.length) throw new Error("Adresse introuvable: " + q);
  const coords = { lat: parseFloat(data[0].lat), lon: parseFloat(data[0].lon) };
  geoCache[q] = coords;
  lsSet(LS_GEO, geoCache);
  return coords;
}

/**
 * Coordonnées pour un chantier : priorité aux lat/lon enregistrés (autocomplete), sinon géocodage.
 */
export async function resolveSiteCoords(site) {
  if (!site || !site.address || !String(site.address).trim()) throw new Error("Adresse vide");
  if (siteHasStoredCoords(site)) {
    const c = { lat: Number(site.lat), lon: Number(site.lon) };
    seedGeoCacheForAddress(site.address, c.lat, c.lon);
    return c;
  }
  return geocodeAddress(site.address);
}

/**
 * @param {"driving"|"foot"} profile
 */
export async function fetchRoute(A, B, profile) {
  const prof = profile === "foot" ? "foot" : "driving";
  const key =
    prof +
    ":" +
    A.lat.toFixed(4) +
    "," +
    A.lon.toFixed(4) +
    "-" +
    B.lat.toFixed(4) +
    "," +
    B.lon.toFixed(4);
  if (routeCache[key] !== undefined) return routeCache[key];
  const url =
    "https://router.project-osrm.org/route/v1/" +
    prof +
    "/" +
    A.lon +
    "," +
    A.lat +
    ";" +
    B.lon +
    "," +
    B.lat +
    "?overview=false";
  const res = await fetch(url, { headers: { Accept: "application/json" } });
  if (!res.ok) throw new Error("OSRM " + res.status);
  const data = await res.json();
  if (!data || !data.routes || !data.routes[0]) throw new Error("Pas de route");
  const minutes = data.routes[0].duration / 60;
  routeCache[key] = minutes;
  lsSet(LS_ROUTE, routeCache);
  return minutes;
}

export function estimateTravelByDistance(A, B, mode) {
  const distKm = haversine(A.lat, A.lon, B.lat, B.lon);
  let speedKmH = 25;
  if (mode === "bike") speedKmH = 13;
  else if (mode === "pt") speedKmH = 18;
  else if (mode === "walk") speedKmH = 5;
  return Math.max(2, Math.round((distKm / speedKmH) * 60));
}

export async function travelMinutes(siteA, siteB, mode) {
  if (!siteA || !siteA.address || !siteB || !siteB.address) return 3;
  if (siteA.address.trim() === siteB.address.trim()) return 0;

  try {
    const A = await resolveSiteCoords(siteA);
    const B = await resolveSiteCoords(siteB);

    if (mode === "walk") {
      try {
        const base = await fetchRoute(A, B, "foot");
        return Math.max(1, Math.round(base * (TRANSPORT_FACTORS.walk || 1)));
      } catch (e2) {
        return estimateTravelByDistance(A, B, "walk");
      }
    }

    const factor = TRANSPORT_FACTORS[mode] || 1;
    try {
      const base = await fetchRoute(A, B, "driving");
      const result = Math.round(base * factor);
      return Math.max(1, result);
    } catch (e2) {
      return estimateTravelByDistance(A, B, mode);
    }
  } catch (e) {
    const cpA = (siteA.address || "").match(/\b13\d{3}\b/);
    const cpB = (siteB.address || "").match(/\b13\d{3}\b/);
    if (cpA && cpB) {
      const diff = Math.abs(parseInt(cpA[0], 10) - parseInt(cpB[0], 10));
      if (diff === 0) return mode === "walk" ? 5 : 3;
      if (diff <= 2) return mode === "walk" ? 12 : 7;
      if (diff <= 5) return mode === "walk" ? 22 : 12;
      return mode === "walk" ? 35 : 18;
    }
    return mode === "walk" ? 15 : 8;
  }
}

export async function batchGeocode(siteList) {
  const coords = {};
  await Promise.all(
    siteList.map(async function (s) {
      if (!s || !s.address) return;
      try {
        if (siteHasStoredCoords(s)) {
          coords[s.id] = { lat: Number(s.lat), lon: Number(s.lon) };
          seedGeoCacheForAddress(s.address, s.lat, s.lon);
        } else {
          coords[s.id] = await geocodeAddress(s.address);
        }
      } catch (e) {}
    })
  );
  return coords;
}
