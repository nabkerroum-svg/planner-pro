/**
 * Suggestions d'adresses via l'API de recherche Nominatim (usage raisonné : debounce côté UI).
 */

const USER_AGENT = "PlannerPro/1.0 (chantiers planning; usage modere)";

function pickCity(addr) {
  if (!addr || typeof addr !== "object") return "";
  return (
    addr.city ||
    addr.town ||
    addr.village ||
    addr.hamlet ||
    addr.municipality ||
    addr.suburb ||
    addr.county ||
    ""
  );
}

/**
 * @param {string} query
 * @returns {Promise<Array<{ displayName: string, postcode: string, city: string, lat: number, lon: number, placeId: string }>>}
 */
export async function fetchAddressSuggestions(query) {
  const q = query && String(query).trim();
  if (!q || q.length < 3) return [];

  const params = new URLSearchParams({
    format: "json",
    q: q,
    limit: "5",
    addressdetails: "1",
    countrycodes: "fr",
  });

  const url = "https://nominatim.openstreetmap.org/search?" + params.toString();
  const res = await fetch(url, {
    headers: {
      Accept: "application/json",
      "Accept-Language": "fr",
      "User-Agent": USER_AGENT,
    },
  });

  if (!res.ok) throw new Error("Nominatim " + res.status);

  const data = await res.json();
  if (!Array.isArray(data)) return [];

  return data.map(function (item) {
    const a = item.address || {};
    return {
      placeId: String(
        item.place_id != null ? item.place_id : [item.osm_type, item.osm_id, item.lat, item.lon].join(":")
      ),
      displayName: item.display_name || "",
      postcode: a.postcode || "",
      city: pickCity(a),
      lat: parseFloat(item.lat),
      lon: parseFloat(item.lon),
    };
  });
}
