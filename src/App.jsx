import React, { useState, useEffect, useRef, useMemo } from "react";
import {
  travelMinutes,
  batchGeocode,
  geocodeAddress,
  transportEmoji,
  transportLabelFr,
} from "./planner/transport.js";
import { parseFreq, computeOccurrenceDayIndices } from "./planner/recurrence.js";
import { orderSitesWithAnchors } from "./planner/tourOrder.js";
import { adjustStartForTimeConstraint } from "./planner/timeConstraints.js";
import { AddressAutocomplete } from "./planner/AddressAutocomplete.jsx";

// ─── Helpers ─────────────────────────────────────────────────────────────────
function pad2(n) { return String(n).padStart(2, "0"); }
function parseHM(str) {
  if (!str) return 0;
  const parts = String(str).split(":");
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  return h * 60 + m;
}
function formatHM(min) {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return pad2(h) + ":" + pad2(m);
}
function splitIsoDateTokens(text) {
  return String(text || "")
    .split(/[\n,;]+/)
    .map(function (x) {
      return x.trim();
    })
    .filter(Boolean);
}

function todayISO() {
  const d = new Date();
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}
function getWeekDates(isoDate) {
  const d = new Date(isoDate);
  const wd = (d.getDay() + 6) % 7;
  const mon = new Date(d);
  mon.setDate(d.getDate() - wd);
  const out = [];
  for (let i = 0; i < 7; i++) {
    const dd = new Date(mon);
    dd.setDate(mon.getDate() + i);
    out.push(dd.getFullYear() + "-" + pad2(dd.getMonth() + 1) + "-" + pad2(dd.getDate()));
  }
  return out;
}
function weekdayIdx(iso) { return (new Date(iso).getDay() + 6) % 7; }
function getMonthDates(isoDate) {
  const parts = isoDate.split("-").map(Number);
  const Y = parts[0]; const M = parts[1];
  const out = [];
  let d = new Date(Y, M - 1, 1);
  while (d.getFullYear() === Y && d.getMonth() === M - 1) {
    out.push(d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

// ─── LocalStorage ─────────────────────────────────────────────────────────────
const LS_KEYS = {
  EMP: "planner_emp_v1",
  SITES: "planner_sites_v1",
  SVC: "planner_svc_v1",
  TASKS: "planner_tasks_v1",
  GRP: "planner_grp_v1",
  FREQ: "planner_freq_v1",
  GEO: "planner_geo_v1",
  ROUTE: "planner_route_v1",
  START_TIME: "planner_start_time_v1",
  END_TIME: "planner_end_time_v1",
  TOUR_START_SITE: "planner_tour_start_site_v1",
  TOUR_START_ADDR: "planner_tour_start_addr_v1",
  TOUR_START_GEO: "planner_tour_start_geo_v1",
  TOUR_END_SITE: "planner_tour_end_site_v1",
  TOUR_END_ADDR: "planner_tour_end_addr_v1"
};
function lsGet(k, def) {
  try { const r = localStorage.getItem(k); return r ? JSON.parse(r) : def; } catch (e) { return def; }
}
function lsSet(k, v) {
  try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {}
}

// Estimation de capacite basee sur la duree de prestation uniquement (sans biais 10 min)
// Le trajet sera recalcule ensuite de facon reelle
function estimateCapacity(occList, allServices) {
  return occList.reduce(function(total, occ) {
    var svc = allServices.find(function(s) { return s.id === occ.svcId; });
    // Estimation trajet : 0 pour le 1er, 8 min moyenne pour les suivants (Marseille)
    return total + (svc ? svc.duration : 60);
  }, 0) + (occList.length > 1 ? (occList.length - 1) * 8 : 0);
}


// ─── Build tasks ─────────────────────────────────────────────────────────────
//
// ALGORITHME EN 6 ETAPES :
// 1. Generer TOUTES les occurrences (generateTargetDays - calendrier strict)
// 2. Assigner chaque occurrence au jour le plus proche disponible
// 3. PHASE D EQUILIBRAGE : redistribuer les chantiers pour lisser la charge
// 4. Tri par proximite (batchGeocode + nearest neighbour)
// 5. Calculer les horaires (trajet avant, prestation complete)
// 6. Rapport de validation et d equilibrage
//
async function buildTasks(opts) {
  var planDate     = opts.planDate;
  var empId        = opts.empId;
  var transport    = opts.transport;
  var siteIds      = opts.siteIds;
  var allSites     = opts.allSites;
  var allServices  = opts.allServices;
  var freqs        = opts.freqs;
  var siteServices = opts.siteServices;
  var existing     = opts.existing;
  var tourStartSiteId = opts.tourStartSiteId || "";
  var tourStartAddress = opts.tourStartAddress || "";
  var tourStartLat = opts.tourStartLat;
  var tourStartLon = opts.tourStartLon;
  var tourEndSiteId = opts.tourEndSiteId || "";

  var START = parseHM(opts.startTime || "06:00");
  var END   = parseHM(opts.endTime   || "12:00");
  var SOFT_CAP = END - START;
  var HARD_CAP = SOFT_CAP + 240; // 600 min max absolu

  var curY = parseInt(planDate.split("-")[0], 10);
  var curM = new Date(planDate + "T12:00:00").getMonth() + 1;

  // ── 1. Jours ouvres lun-sam ───────────────────────────────────────────────
  var workdays = getMonthDates(curY + "-" + pad2(curM) + "-01").filter(function(d) {
    return weekdayIdx(d) <= 5;
  });
  var NB = workdays.length;
  if (NB === 0) return existing.slice();

  // ── 2. Generer TOUTES les occurrences (calendrier strict, pas de ratio) ───
  var occurrences = [];
  var expectedPerSite = {};

  for (var i = 0; i < siteIds.length; i++) {
    var sid   = siteIds[i];
    var site  = allSites.find(function(s) { return s.id === sid; });
    if (!site) continue;
    var svcId = siteServices[sid] || site.defaultServiceId;
    if (!svcId) continue;

    var freq  = freqs.find(function(f) { return f.siteId === sid; });
    var fp    = freq ? freq : { type: "once", timesPerMonth: 1 };
    var ftype = fp.type || "once";

    var svc0  = allServices.find(function(s) { return s.id === svcId; });
    var dur0  = svc0 ? svc0.duration : 60;

    var targetDays = computeOccurrenceDayIndices(workdays, fp);

    if (!expectedPerSite[sid]) expectedPerSite[sid] = 0;
    expectedPerSite[sid] += targetDays.length;

    for (var ti = 0; ti < targetDays.length; ti++) {
      occurrences.push({
        site:     site,
        svcId:    svcId,
        siteIdx:  i,
        target:   targetDays[ti],
        dur:      dur0,
        freqType: ftype,
        fixed:    ftype === "fixed"
      });
    }
  }

  var expectedTotal = occurrences.length;

  // ── 3. Capacite initiale des jours existants ──────────────────────────────
  // Estimation de la charge d un jour : somme des durees + 8 min par trajet
  function dayCapacity(ocList) {
    if (!ocList || ocList.length === 0) return 0;
    return ocList.reduce(function(s, o) { return s + o.dur; }, 0) +
           Math.max(0, ocList.length - 1) * 8;
  }

  var dayLoad = new Array(NB).fill(0);
  existing.forEach(function(t) {
    if (t.employeeId !== empId) return;
    var di = workdays.indexOf(t.date);
    if (di < 0) return;
    var svc = allServices.find(function(s) { return s.id === t.serviceId; });
    dayLoad[di] += (svc ? svc.duration : 60) + (t.travelMinutes || 8);
  });

  // ── 4. Assignation initiale (radius search depuis le jour cible) ──────────
  occurrences.sort(function(a, b) {
    return a.target !== b.target ? a.target - b.target : a.siteIdx - b.siteIdx;
  });

  var dayAssignments = [];
  for (var di2 = 0; di2 < NB; di2++) dayAssignments.push([]);

  occurrences.forEach(function(occ) {
    var needed = occ.dur + 8;
    var target = occ.target;
    var assigned = -1;

    for (var radius = 0; radius < NB && assigned === -1; radius++) {
      var cands = radius === 0 ? [target] : [target + radius, target - radius];
      for (var ci = 0; ci < cands.length; ci++) {
        var d = cands[ci];
        if (d < 0 || d >= NB) continue;
        if (dayLoad[d] + needed <= SOFT_CAP) {
          assigned = d; break;
        }
      }
    }
    if (assigned === -1) {
      for (var radius2 = 0; radius2 < NB && assigned === -1; radius2++) {
        var cands2 = radius2 === 0 ? [target] : [target + radius2, target - radius2];
        for (var ci2 = 0; ci2 < cands2.length; ci2++) {
          var d2 = cands2[ci2];
          if (d2 < 0 || d2 >= NB) continue;
          if (dayLoad[d2] + needed <= HARD_CAP) {
            assigned = d2; break;
          }
        }
      }
    }
    if (assigned === -1) {
      var minL = Infinity;
      for (var d3 = 0; d3 < NB; d3++) {
        if (dayLoad[d3] < minL) { minL = dayLoad[d3]; assigned = d3; }
      }
    }

    dayAssignments[assigned].push(occ);
    dayLoad[assigned] += occ.dur + 8;
  });

  // ── 5. PHASE D EQUILIBRAGE ────────────────────────────────────────────────
  // Objectif : toutes les journees entre START et END (06h-12h)
  // Methode  : deplacer iterativement des chantiers du jour le plus charge
  //            vers le meilleur destinataire, jusqu a convergence.
  //
  // Chantier transferable : non fixe, deplacement <= 10 jours
  // Seuil d equilibrage   : si max - min > 60 min → rééquilibrer

  var beforeLoads = dayAssignments.map(function(dl) { return dayCapacity(dl); });

  var movedTasks = [];
  var BALANCE_ITERATIONS = 120;
  var DIST_MAX = 15; // max 3 semaines de decalage pour un transfert

  for (var iter = 0; iter < BALANCE_ITERATIONS; iter++) {
    var loads = dayAssignments.map(function(dl) { return dayCapacity(dl); });

    // Trouver le jour le plus surchargé et le moins chargé
    var maxIdx = 0;
    var minLoad = Infinity;
    for (var di3 = 1; di3 < NB; di3++) {
      if (loads[di3] > loads[maxIdx]) maxIdx = di3;
    }
    for (var di3b = 0; di3b < NB; di3b++) {
      if (loads[di3b] < minLoad) minLoad = loads[di3b];
    }

    // Arreter si le jour max est raisonnable ET l'ecart max-min est acceptable (<= 90 min)
    if (loads[maxIdx] <= 370 && loads[maxIdx] - minLoad <= 90) break;

    // Chantiers mobiles du jour surchargé
    var mobile = dayAssignments[maxIdx].filter(function(o) { return !o.fixed; });
    if (!mobile.length) break;

    // Chercher le meilleur (occ, dst) qui minimise le nouveau max
    var bestMove = null;
    var bestScore = Infinity;

    for (var mi = 0; mi < mobile.length; mi++) {
      var occ = mobile[mi];
      var newSrcLoad = loads[maxIdx] - occ.dur - 8;

      for (var dst = 0; dst < NB; dst++) {
        if (dst === maxIdx) continue;
        var dist = Math.abs(dst - maxIdx);
        if (dist > DIST_MAX) continue; // trop loin temporellement

        var newDstLoad = loads[dst] + occ.dur + 8;
        if (360 + newDstLoad > END + 60) continue; // destinataire pas trop surchargé

        var newMax = Math.max(newSrcLoad, newDstLoad);
        var score  = newMax + dist; // favoriser les jours proches
        if (score < bestScore) {
          bestScore = score;
          bestMove  = { occ: occ, src: maxIdx, dst: dst };
        }
      }
    }

    if (!bestMove) break;

    // Effectuer le deplacement
    var occIdx = dayAssignments[bestMove.src].indexOf(bestMove.occ);
    dayAssignments[bestMove.src].splice(occIdx, 1);
    dayAssignments[bestMove.dst].push(bestMove.occ);
    movedTasks.push({
      siteId: bestMove.occ.site.id,
      name:   bestMove.occ.site.name || bestMove.occ.site.id,
      from:   workdays[bestMove.src],
      to:     workdays[bestMove.dst]
    });
  }

  var afterLoads = dayAssignments.map(function(dl) { return dayCapacity(dl); });

  // ── 6. Placer les taches jour par jour ────────────────────────────────────
  var acc = existing.slice();

  for (var di4 = 0; di4 < NB; di4++) {
    var day    = workdays[di4];
    var ocList = dayAssignments[di4];
    if (!ocList || ocList.length === 0) continue;

    var sObjs = ocList.map(function(occ) { return occ.site; });
    var sortedSites;
    try {
      var dayCoords = await batchGeocode(sObjs);
      var startSeed = null;
      if (!tourStartSiteId) {
        if (
          tourStartLat != null &&
          tourStartLon != null &&
          Number.isFinite(Number(tourStartLat)) &&
          Number.isFinite(Number(tourStartLon))
        ) {
          startSeed = { lat: Number(tourStartLat), lon: Number(tourStartLon) };
        } else if (tourStartAddress && tourStartAddress.trim()) {
          try {
            var gSeed = await geocodeAddress(tourStartAddress.trim());
            startSeed = { lat: gSeed.lat, lon: gSeed.lon };
          } catch (eSeed) {}
        }
      }
      sortedSites = orderSitesWithAnchors(sObjs, dayCoords, {
        startSiteId: tourStartSiteId || null,
        endSiteId: tourEndSiteId || null,
        startSeedCoords: startSeed || undefined
      });
    } catch(e) {
      sortedSites = sObjs;
    }

    var dayEx = acc.filter(function(t) {
      return t.employeeId === empId && t.date === day;
    }).sort(function(a, b) { return parseHM(a.startAt) - parseHM(b.startAt); });

    var cursor   = START;
    var prevSite = null;

    if (dayEx.length > 0) {
      var lastEx  = dayEx[dayEx.length - 1];
      var lastSvc = allServices.find(function(s) { return s.id === lastEx.serviceId; });
      cursor   = parseHM(lastEx.startAt) + (lastSvc ? lastSvc.duration : 60);
      prevSite = allSites.find(function(s) { return s.id === lastEx.siteId; }) || null;
    } else {
      if (tourStartSiteId) {
        prevSite = allSites.find(function(s) { return s.id === tourStartSiteId; }) || null;
      } else if (tourStartAddress && tourStartAddress.trim()) {
        prevSite = {
          id: "__tour_start__",
          name: "Depart",
          address: tourStartAddress.trim(),
          lat: tourStartLat,
          lon: tourStartLon
        };
      }
    }

    for (var si = 0; si < sortedSites.length; si++) {
      var curSite = sortedSites[si];
      var occ2    = ocList.find(function(o) { return o.site.id === curSite.id; });
      if (!occ2) continue;

      var svcId3 = occ2.svcId;
      var svc3   = allServices.find(function(s) { return s.id === svcId3; });
      var dur3   = svc3 ? svc3.duration : 60;

      var trav = 0;
      if (prevSite) {
        try { trav = await travelMinutes(prevSite, curSite, transport); }
        catch(e) {
          var cpA = (prevSite.address || "").match(/\b13\d{3}\b/);
          var cpB = (curSite.address || "").match(/\b13\d{3}\b/);
          if (cpA && cpB) {
            var diff = Math.abs(parseInt(cpA[0]) - parseInt(cpB[0]));
            trav = diff === 0 ? 3 : diff <= 2 ? 7 : diff <= 5 ? 12 : 18;
          } else { trav = 8; }
        }
      }

      var rawStart = cursor + trav;
      var startMin = adjustStartForTimeConstraint(
        rawStart,
        dur3,
        cursor + trav,
        curSite.timeConstraint || null
      );
      var endMin   = startMin + dur3;

      cursor   = endMin;
      prevSite = curSite;

      acc.push({
        id:                String(Math.random()),
        date:              day,
        employeeId:        empId,
        siteId:            curSite.id,
        serviceId:         svcId3,
        transport:         transport,
        startAt:           formatHM(startMin),
        travelMinutes:     Math.round(trav),
        effectiveDuration: dur3,
        lateFlag:          endMin > END,
        timeConstraint:    curSite.timeConstraint || undefined
      });
    }
  }

  // ── 7. Rapport final ──────────────────────────────────────────────────────
  var newTasks = acc.filter(function(t) {
    return t.employeeId === empId && !existing.find(function(e) { return e.id === t.id; });
  });

  var plannedPerSite = {};
  var totalTrav = 0; var travDist = {}; var exactly8 = 0; var exactly10 = 0;

  newTasks.forEach(function(t) {
    if (!plannedPerSite[t.siteId]) plannedPerSite[t.siteId] = 0;
    plannedPerSite[t.siteId]++;
    var tr = Math.round(t.travelMinutes || 0);
    totalTrav += tr;
    if (tr === 8) exactly8++;
    if (tr === 10) exactly10++;
    var bkt = tr <= 5 ? "0-5" : tr <= 10 ? "6-10" : tr <= 15 ? "11-15" : tr <= 20 ? "16-20" : "21+";
    if (!travDist[bkt]) travDist[bkt] = 0;
    travDist[bkt]++;
  });

  var plannedTotal = Object.values(plannedPerSite).reduce(function(s, v) { return s + v; }, 0);
  var missing      = expectedTotal - plannedTotal;
  var avgTrav      = newTasks.length > 0 ? Math.round(totalTrav / newTasks.length) : 0;

  // Stats d equilibrage
  var underDays = []; var overDays = [];
  afterLoads.forEach(function(l, i) {
    var endMin = 360 + l;
    if (endMin > END) overDays.push({ day: workdays[i], endTime: formatHM(endMin) });
    if (endMin < 660) underDays.push({ day: workdays[i], endTime: formatHM(endMin) });
  });

  var perSiteErrors = [];
  siteIds.forEach(function(sid4) {
    var exp = expectedPerSite[sid4] || 0;
    var got = plannedPerSite[sid4] || 0;
    if (exp !== got) {
      var s2 = allSites.find(function(x) { return x.id === sid4; });
      perSiteErrors.push({ name: s2 ? s2.name : sid4, expected: exp, placed: got });
    }
  });

  console.log("[PlannerPro] ══ RAPPORT PLANNING + EQUILIBRAGE ══");
  console.log("  Occurrences attendues : " + expectedTotal);
  console.log("  Interventions placees : " + plannedTotal);
  console.log("  Manquantes            : " + missing);
  console.log("  Deplacements equilibrage : " + movedTasks.length);
  console.log("  Jours surcharges (>12h) : " + overDays.length);
  console.log("  Jours sous-charges (<11h) : " + underDays.length);
  console.log("  Trajet moyen : " + avgTrav + " min");
  console.log("  Trajets = 8min (fallback) : " + exactly8);
  console.log("  Trajets = 10min (biais?) : " + exactly10);
  console.log("  Distribution trajets : " + JSON.stringify(travDist));

  if (missing > 0) {
    console.error("[PlannerPro] ERREUR : " + missing + " intervention(s) manquante(s) !");
    perSiteErrors.forEach(function(r) {
      console.error("  >> " + r.name + " attendu=" + r.expected + " place=" + r.placed);
    });
  }
  if (missing === 0 && movedTasks.length === 0) {
    console.log("[PlannerPro] Planning equilibre sans deplacement necessaire");
  }
  if (movedTasks.length > 0) {
    console.log("[PlannerPro] Deplacements effectues:");
    movedTasks.slice(0, 5).forEach(function(m) {
      console.log("  >> " + m.name + " : " + m.from + " → " + m.to);
    });
    if (movedTasks.length > 5) {
      console.log("  ... et " + (movedTasks.length - 5) + " autres");
    }
  }

  return acc;
}


const THEME = {
  primary:   "#0F0F0F",
  gold:      "#B08D57",
  goldLight: "#C9A96E",
  goldPale:  "#F7F2E8",
  goldBorder:"#D4B87A",
  dark:      "#1A1A1A",
  bg:        "#F2F0EC",
  card:      "#FFFFFF",
  cardBorder:"#E8E2D8",
  pale:      "#F7F2E8",
  grid:      "#E8E2D8",
  textMuted: "#6B6355",
  danger:    "#C0392B",
  success:   "#2E7D32",
  shadow:    "0 2px 8px rgba(15,10,0,0.08)",
  shadowMd:  "0 4px 16px rgba(15,10,0,0.12)",
  radius:    "12px",
  radiusSm:  "8px",
};
const LOGO = "data:image/png;base64,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";

// ─── UI Components ────────────────────────────────────────────────────────────
function AppBox(props) {
  return (
    <div style={{
      border: "1px solid " + THEME.cardBorder,
      borderRadius: THEME.radius,
      padding: "16px 18px",
      background: THEME.card,
      boxShadow: THEME.shadow,
      ...(props.style || {})
    }}>
      {(props.title || props.right) && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, paddingBottom: 10, borderBottom: "1px solid " + THEME.grid }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: THEME.primary, letterSpacing: 0.3, textTransform: "uppercase" }}>{props.title}</div>
          <div>{props.right}</div>
        </div>
      )}
      {props.children}
    </div>
  );
}

function AppRow(props) {
  return (
    <div style={{ display: "flex", gap: props.gap || 8, alignItems: props.align || "center", flexWrap: props.wrap === false ? "nowrap" : "wrap", ...(props.style || {}) }}>
      {props.children}
    </div>
  );
}

function AppInput(props) {
  const { style, ...rest } = props;
  return (
    <input
      {...rest}
      value={props.value === undefined || props.value === null ? "" : props.value}
      style={{
        padding: "8px 11px",
        borderRadius: THEME.radiusSm,
        border: "1px solid " + THEME.cardBorder,
        background: "#FAFAF8",
        color: THEME.primary,
        fontSize: 13,
        minWidth: 0,
        outline: "none",
        transition: "border 0.2s",
        ...(style || {})
      }}
    />
  );
}

function AppSelect(props) {
  const { style, ...rest } = props;
  return (
    <select
      {...rest}
      value={props.value === undefined || props.value === null ? "" : props.value}
      style={{
        padding: "8px 11px",
        borderRadius: THEME.radiusSm,
        border: "1px solid " + THEME.cardBorder,
        background: "#FAFAF8",
        color: THEME.primary,
        fontSize: 13,
        cursor: "pointer",
        ...(style || {})
      }}
    />
  );
}

function AppButton(props) {
  const { kind, style, children, ...rest } = props;
  let bg, color, border;
  if (kind === "primary") {
    bg = THEME.primary; color = "#FFFFFF"; border = "none";
  } else if (kind === "gold") {
    bg = THEME.gold; color = "#FFFFFF"; border = "none";
  } else if (kind === "danger") {
    bg = THEME.danger; color = "#FFFFFF"; border = "none";
  } else {
    bg = "#FFFFFF"; color = THEME.primary; border = "1px solid " + THEME.cardBorder;
  }
  return (
    <button {...rest} style={{
      padding: "7px 14px",
      borderRadius: THEME.radiusSm,
      border: border,
      background: bg,
      color: color,
      cursor: "pointer",
      fontSize: 12.5,
      fontWeight: 500,
      letterSpacing: 0.2,
      transition: "opacity 0.15s, transform 0.1s",
      ...(style || {})
    }}>
      {children}
    </button>
  );
}

function AppBadge(props) {
  return (
    <span style={{
      background: THEME.goldPale,
      border: "1px solid " + THEME.goldBorder,
      color: THEME.gold,
      borderRadius: 999,
      padding: "2px 10px",
      fontSize: 11,
      fontWeight: 600,
      letterSpacing: 0.5,
    }}>
      {props.children}
    </span>
  );
}

function GoldDivider() {
  return <div style={{ height: 1, background: "linear-gradient(90deg, transparent, " + THEME.gold + ", transparent)", margin: "10px 0" }} />;
}

// ─── Import Panel ─────────────────────────────────────────────────────────────
function ImportPanel(props) {
  const [status, setStatus] = useState(null);
  const [preview, setPreview] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const fileRef = useRef();

  function loadSheetJS() {
    return new Promise(function(resolve, reject) {
      if (window.XLSX) return resolve(window.XLSX);
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      s.onload = function() { resolve(window.XLSX); };
      s.onerror = function() { reject(new Error("SheetJS indisponible")); };
      document.head.appendChild(s);
    });
  }

  function normalizeKey(k) {
    return String(k).trim().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  }

  function getCol(row) {
    const keys = Object.keys(row);
    function find() {
      const searchKeys = Array.prototype.slice.call(arguments);
      for (let i = 0; i < searchKeys.length; i++) {
        const nk = normalizeKey(searchKeys[i]);
        const found = keys.find(function(rk) { return normalizeKey(rk).indexOf(nk) !== -1; });
        if (found !== undefined && String(row[found]).trim() !== "") return String(row[found]).trim();
      }
      return "";
    }
    return {
      address: find("adresse / chantier", "adresse chantier", "adresse", "chantier", "address"),
      employee: find("salarie", "employe", "employee", "name"),
      freq: find("frequence", "frequency", "freq", "rythme"),
      service: find("prestation", "service", "type"),
      duration: find("duree", "duration", "temps"),
      siteName: find("nom chantier", "nom", "name site")
    };
  }

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setStatus("Lecture en cours...");
    setPreview([]);
    setRawRows([]);
    try {
      let rows = [];
      if (/\.(xlsx|xls|ods)$/i.test(file.name)) {
        setStatus("Chargement SheetJS...");
        const XLSX = await loadSheetJS();
        const ab = await file.arrayBuffer();
        const wb = XLSX.read(new Uint8Array(ab), { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
        rows = json.map(function(r) {
          const out = {};
          Object.keys(r).forEach(function(k) { out[normalizeKey(k)] = String(r[k] || "").trim(); });
          return out;
        });
      } else {
        const ab = await file.arrayBuffer();
        let text;
        try { text = new TextDecoder("utf-8", { fatal: true }).decode(ab); }
        catch (err) { text = new TextDecoder("latin-1").decode(ab); }
        const lines = text.split(/\r?\n/).filter(function(l) { return l.trim(); });
        if (lines.length < 2) { setStatus("Fichier vide."); return; }
        const sep = lines[0].indexOf(";") !== -1 ? ";" : ",";
        const headers = lines[0].split(sep).map(function(h) { return normalizeKey(h.replace(/['"]/g, "")); });
        for (let i = 1; i < lines.length; i++) {
          const cells = lines[i].split(sep).map(function(c) { return c.trim().replace(/^["']|["']$/g, ""); });
          if (cells.every(function(c) { return !c; })) continue;
          const row = {};
          headers.forEach(function(h, idx) { row[h] = cells[idx] || ""; });
          rows.push(row);
        }
      }
      if (rows.length === 0) { setStatus("Aucune ligne trouvee."); return; }
      const cols = Object.keys(rows[0]).join(" | ");
      setRawRows(rows);
      setPreview(rows.slice(0, 5));
      setStatus(rows.length + " lignes detectees - Colonnes : " + cols);
    } catch (err) {
      setStatus("ERREUR : " + err.message);
      console.error(err);
    }
    e.target.value = "";
  }

  function handleImport() {
    if (!rawRows.length) return;
    let newSites = props.sites.slice();
    let newEmps = props.employees.slice();
    let newSvcs = props.services.slice();
    let newFreqs = props.siteFrequencies.slice();
    let imported = 0;
    let skipped = 0;

    rawRows.forEach(function(row) {
      const cols = getCol(row);
      const address = cols.address;
      const empName = cols.employee;
      const freqStr = cols.freq;
      const svcName = cols.service;
      const dureeStr = cols.duration;
      const siteNameCol = cols.siteName;

      if (!address) { skipped++; return; }

      const siteName = siteNameCol || address.substring(0, 30).toUpperCase();

      let durMin = 60;
      if (dureeStr) {
        const hm = dureeStr.match(/^(\d+):(\d+)$/);
        if (hm) durMin = parseInt(hm[1], 10) * 60 + parseInt(hm[2], 10);
        else if (!isNaN(parseFloat(dureeStr))) durMin = parseFloat(dureeStr);
      }

      if (empName && !newEmps.find(function(e) { return e.name.toLowerCase().trim() === empName.toLowerCase().trim(); })) {
        newEmps.push({ id: String(Math.random()), name: empName.trim() });
      }

      let svcId;
      if (svcName) {
        let svc = newSvcs.find(function(s) {
          return s.name.toLowerCase().indexOf(svcName.toLowerCase()) !== -1 || svcName.toLowerCase().indexOf(s.name.toLowerCase()) !== -1;
        });
        if (!svc) { svc = { id: String(Math.random()), name: svcName.trim(), duration: durMin }; newSvcs.push(svc); }
        svcId = svc.id;
      }

      let site = newSites.find(function(s) { return s.address && s.address.toLowerCase().trim() === address.toLowerCase().trim(); });
      if (!site) {
        site = { id: String(Math.random()), name: siteName, address: address.trim(), defaultServiceId: svcId };
        newSites.push(site);
      } else if (svcId && !site.defaultServiceId) {
        const updatedSite = Object.assign({}, site, { defaultServiceId: svcId });
        newSites = newSites.map(function(s) { return s.id === site.id ? updatedSite : s; });
        site = updatedSite;
      }

      if (freqStr) {
        const fp = parseFreq(freqStr);
        const fi = newFreqs.findIndex(function(f) { return f.siteId === site.id; });
        if (fi === -1) newFreqs.push(Object.assign({ siteId: site.id }, fp));
        else newFreqs[fi] = Object.assign({}, newFreqs[fi], fp);
      }
      imported++;
    });

    props.setSites(newSites);
    props.setEmployees(newEmps);
    props.setServices(newSvcs);
    props.setSiteFrequencies(newFreqs);
    setStatus(imported + " chantier(s) importes" + (skipped > 0 ? ", " + skipped + " ignores" : ""));
    setRawRows([]);
    setPreview([]);
  }

  const isErr = status && (status.indexOf("ERREUR") !== -1 || status.indexOf("vide") !== -1);
  const isOk = status && (status.indexOf("importes") !== -1 || status.indexOf("detectees") !== -1);

  return (
    <div style={{ border: "2px dashed " + THEME.gold, borderRadius: 12, padding: 16, background: "#fffdf9", marginBottom: 16 }}>
      <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 6, color: THEME.gold }}>Import Excel / CSV</div>
      <div style={{ fontSize: 12, color: "#777", marginBottom: 10 }}>
        Colonnes : <strong>Adresse / Chantier</strong> | <strong>Salarie</strong> | <strong>Frequence</strong> | <strong>Prestation</strong> | <strong>Duree</strong>
        <br />Frequence : nombre de passages par mois (ex: 2, 4, 8...)
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
        <AppButton onClick={function() {
          const csv = "Adresse / Chantier;Salarie;Frequence;Prestation;Duree\n1 rue Linne, 13004 Marseille;CHAABANI HABIB;2;COMPLET : Nettoyage complet;01:00\n114 Rue Corse, Marseille;CHAABANI HABIB;4;COMMUNS : Entretien des parties communes;00:45\n";
          const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url; a.download = "modele_import.csv"; a.click();
          setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
        }} style={{ fontSize: 12 }}>Modele CSV</AppButton>
        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />
        <AppButton kind="primary" onClick={function() { if (fileRef.current) fileRef.current.click(); }} style={{ fontSize: 12 }}>
          Choisir fichier (.xlsx / .csv)
        </AppButton>
        {rawRows.length > 0 && (
          <AppButton onClick={handleImport} style={{ fontSize: 12, background: THEME.gold, border: "none", color: "#fff", fontWeight: 700 }}>
            Importer {rawRows.length} ligne(s)
          </AppButton>
        )}
      </div>
      {status && (
        <div style={{ fontSize: 12, padding: "6px 10px", borderRadius: 8, marginBottom: 8, background: isErr ? "#fef2f2" : isOk ? "#f0fdf4" : "#fffbeb", color: isErr ? "#dc2626" : isOk ? "#16a34a" : "#92400e" }}>
          {status}
        </div>
      )}
      {preview.length > 0 && (
        <div style={{ overflowX: "auto" }}>
          <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 4 }}>Apercu :</div>
          <table style={{ borderCollapse: "collapse", fontSize: 11 }}>
            <thead>
              <tr>
                {Object.keys(preview[0]).map(function(h) {
                  return <th key={h} style={{ border: "1px solid #ddd", padding: "4px 8px", background: "#f5f5f5", textAlign: "left", whiteSpace: "nowrap" }}>{h}</th>;
                })}
              </tr>
            </thead>
            <tbody>
              {preview.map(function(row, i) {
                return (
                  <tr key={i}>
                    {Object.values(row).map(function(v, j) {
                      return <td key={j} style={{ border: "1px solid #ddd", padding: "4px 8px", whiteSpace: "nowrap", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis" }}>{v}</td>;
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("planning");
  const [isAdding, setIsAdding] = useState(false);
  const [searchEmps, setSearchEmps] = useState("");
  const [searchSites, setSearchSites] = useState("");
  const [searchSvcs, setSearchSvcs] = useState("");

  const [employees, setEmployees] = useState(function() { return lsGet(LS_KEYS.EMP, [{ id: "e1", name: "Frederic Martin" }, { id: "e2", name: "Elizabeth Rey" }]); });
  const [sites, setSites] = useState(function() { return lsGet(LS_KEYS.SITES, [{ id: "s1", name: "106 BRETEUIL", address: "106 Avenue de Breteuil, Marseille" }, { id: "s2", name: "114 CORSE", address: "114 Rue Corse, Marseille" }]); });
  const [groups, setGroups] = useState(function() { return lsGet(LS_KEYS.GRP, []); });
  const [freqs, setFreqs] = useState(function() { return lsGet(LS_KEYS.FREQ, []); });
  const [services, setServices] = useState(function() { return lsGet(LS_KEYS.SVC, [{ id: "p1", name: "Entretien", duration: 60 }, { id: "p2", name: "Vitrerie", duration: 30 }]); });
  const [tasks, setTasks] = useState(function() { return lsGet(LS_KEYS.TASKS, []); });

  useEffect(function() { lsSet(LS_KEYS.EMP, employees); }, [employees]);
  useEffect(function() { lsSet(LS_KEYS.SITES, sites); }, [sites]);
  useEffect(function() { lsSet(LS_KEYS.GRP, groups); }, [groups]);
  useEffect(function() { lsSet(LS_KEYS.FREQ, freqs); }, [freqs]);
  useEffect(function() { lsSet(LS_KEYS.SVC, services); }, [services]);
  useEffect(function() { lsSet(LS_KEYS.TASKS, tasks); }, [tasks]);

  const [planDate, setPlanDate] = useState(todayISO());
  const [planEmpId, setPlanEmpId] = useState("");
  const [planSiteIds, setPlanSiteIds] = useState([]);
  const [planGrpIds, setPlanGrpIds] = useState([]);
  const [siteServices, setSiteServices] = useState({});
  const [planTransport, setPlanTransport] = useState("car");
  const [planStartTime, setPlanStartTime] = useState(function() { return lsGet(LS_KEYS.START_TIME, "06:00"); });
  const [planEndTime,   setPlanEndTime]   = useState(function() { return lsGet(LS_KEYS.END_TIME,   "12:00"); });
  const [planTourStartSiteId, setPlanTourStartSiteId] = useState(function() { return lsGet(LS_KEYS.TOUR_START_SITE, ""); });
  const [planTourStartGeo, setPlanTourStartGeo] = useState(function() {
    var g = lsGet(LS_KEYS.TOUR_START_GEO, null);
    if (g && typeof g === "object" && typeof g.address === "string") {
      return { address: g.address, lat: g.lat, lon: g.lon };
    }
    return { address: lsGet(LS_KEYS.TOUR_START_ADDR, "") || "", lat: undefined, lon: undefined };
  });
  const [planTourEndSiteId, setPlanTourEndSiteId] = useState(function() { return lsGet(LS_KEYS.TOUR_END_SITE, ""); });
  const [planTourEndAddress, setPlanTourEndAddress] = useState(function() { return lsGet(LS_KEYS.TOUR_END_ADDR, ""); });
  const [viewMode, setViewMode] = useState("day");

  useEffect(function() { lsSet(LS_KEYS.START_TIME, planStartTime); }, [planStartTime]);
  useEffect(function() { lsSet(LS_KEYS.END_TIME,   planEndTime);   }, [planEndTime]);
  useEffect(function() { lsSet(LS_KEYS.TOUR_START_SITE, planTourStartSiteId); }, [planTourStartSiteId]);
  useEffect(function() {
    lsSet(LS_KEYS.TOUR_START_GEO, planTourStartGeo);
    lsSet(LS_KEYS.TOUR_START_ADDR, planTourStartGeo.address || "");
  }, [planTourStartGeo]);
  useEffect(function() { lsSet(LS_KEYS.TOUR_END_SITE, planTourEndSiteId); }, [planTourEndSiteId]);
  useEffect(function() { lsSet(LS_KEYS.TOUR_END_ADDR, planTourEndAddress); }, [planTourEndAddress]);
  const [editingFreq, setEditingFreq] = useState(null);

  const [newEmpName, setNewEmpName] = useState("");
  const [newSiteName, setNewSiteName] = useState("");
  const [newSiteGeo, setNewSiteGeo] = useState({ address: "", lat: undefined, lon: undefined });
  const [newGrpName, setNewGrpName] = useState("");
  const [newSvcName, setNewSvcName] = useState("");
  const [newSvcDur, setNewSvcDur] = useState("60");

  const DAYS_LABELS = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"];
  const MONTHS_LABELS = ["Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"];

  const curEmp = employees.find(function(e) { return e.id === planEmpId; }) || null;
  const weekDates = useMemo(function() { return getWeekDates(planDate); }, [planDate]);

  useEffect(function() {
    const allIds = planSiteIds.concat(planGrpIds.reduce(function(acc, gid) {
      return acc.concat(sites.filter(function(s) { return s.groupId === gid; }).map(function(s) { return s.id; }));
    }, []));
    const ns = {};
    allIds.forEach(function(sid) {
      const site = sites.find(function(s) { return s.id === sid; });
      if (site && site.defaultServiceId && !siteServices[sid]) ns[sid] = site.defaultServiceId;
    });
    if (Object.keys(ns).length > 0) setSiteServices(function(p) { return Object.assign({}, p, ns); });
  }, [planSiteIds, planGrpIds, sites]);

  const tasksForDay = useMemo(function() {
    if (!curEmp) return [];
    return tasks.filter(function(t) { return t.date === planDate && t.employeeId === curEmp.id; })
      .sort(function(a, b) { return parseHM(a.startAt) - parseHM(b.startAt); });
  }, [tasks, planDate, curEmp]);

  const tasksByWeek = useMemo(function() {
    if (!curEmp) return {};
    const map = {};
    weekDates.forEach(function(d) { map[d] = []; });
    tasks.forEach(function(t) {
      if (t.employeeId === curEmp.id && weekDates.indexOf(t.date) !== -1) map[t.date].push(t);
    });
    Object.values(map).forEach(function(l) { l.sort(function(a, b) { return parseHM(a.startAt) - parseHM(b.startAt); }); });
    return map;
  }, [tasks, weekDates, curEmp]);

  const tasksByMonth = useMemo(function() {
    if (!curEmp || !planDate) return {};
    const parts = planDate.split("-").map(Number);
    const Y = parts[0]; const M = parts[1];
    const map = {};
    tasks.forEach(function(t) {
      if (t.employeeId !== curEmp.id || !t.date) return;
      const p2 = t.date.split("-").map(Number);
      if (p2[0] === Y && p2[1] === M) {
        if (!map[t.date]) map[t.date] = [];
        map[t.date].push(t);
      }
    });
    Object.values(map).forEach(function(l) { l.sort(function(a, b) { return parseHM(a.startAt) - parseHM(b.startAt); }); });
    return map;
  }, [tasks, curEmp, planDate]);

  const allSelectedSiteIds = useMemo(function() {
    const fromGroups = planGrpIds.reduce(function(acc, gid) {
      return acc.concat(sites.filter(function(s) { return s.groupId === gid; }).map(function(s) { return s.id; }));
    }, []);
    return [...new Set(planSiteIds.concat(fromGroups))];
  }, [planSiteIds, planGrpIds, sites]);

  const sitesNeedingSvc = useMemo(function() {
    return allSelectedSiteIds.filter(function(sid) {
      const site = sites.find(function(s) { return s.id === sid; });
      return !((site && site.defaultServiceId) || siteServices[sid]);
    });
  }, [allSelectedSiteIds, sites, siteServices]);

  async function addTask() {
    if (!planEmpId) { alert("Selectionne un salarie"); return; }
    const ids = allSelectedSiteIds;
    if (!ids.length) { alert("Selectionne au moins un chantier"); return; }
    for (let i = 0; i < ids.length; i++) {
      const sid = ids[i];
      const site = sites.find(function(s) { return s.id === sid; });
      if (!siteServices[sid] && !(site && site.defaultServiceId)) {
        alert("Prestation manquante pour " + (site ? site.name : sid)); return;
      }
    }
    setIsAdding(true);
    try {
      // Construire les frequences effectives :
      // - Si frequence configuree : l'utiliser
      // - Si pas de frequence ET vue mois : repartir sur le mois entier (4x/mois par defaut)
      // - Si pas de frequence ET vue jour/semaine : planifier ce jour uniquement
      var effectiveFreqs = freqs.slice();
      ids.forEach(function(sid) {
        var hasFreq = effectiveFreqs.find(function(f) { return f.siteId === sid; });
        if (!hasFreq) {
          if (viewMode === "month") {
            // Pas de freq configuree + vue mois = 4 fois par mois par defaut
            effectiveFreqs.push({ siteId: sid, type: "monthly", timesPerMonth: 4 });
          } else {
            effectiveFreqs.push({ siteId: sid, type: "once", timesPerMonth: 1 });
          }
        }
      });
      const acc = await buildTasks({
        planDate,
        empId: planEmpId,
        transport: planTransport,
        siteIds: ids,
        allSites: sites,
        allServices: services,
        freqs: effectiveFreqs,
        siteServices,
        existing: tasks,
        startTime: planStartTime,
        endTime: planEndTime,
        tourStartSiteId: planTourStartSiteId,
        tourStartAddress: planTourStartGeo.address,
        tourStartLat: planTourStartGeo.lat,
        tourStartLon: planTourStartGeo.lon,
        tourEndSiteId: planTourEndSiteId
      });
      setTasks(acc);
      setPlanSiteIds([]);
      setPlanGrpIds([]);
      setSiteServices({});
    } catch (err) {
      alert("Erreur : " + err.message);
    } finally {
      setIsAdding(false);
    }
  }

  function updateTask(id, patch) {
    setTasks(function(prev) { return prev.map(function(t) { return t.id === id ? Object.assign({}, t, patch) : t; }); });
  }

  function deleteTask(id) {
    setTasks(function(prev) { return prev.filter(function(t) { return t.id !== id; }); });
  }

  function deleteView() {
    if (!curEmp) return;
    let ids = [];
    if (viewMode === "day") ids = tasksForDay.map(function(t) { return t.id; });
    else if (viewMode === "week") Object.values(tasksByWeek).forEach(function(l) { l.forEach(function(t) { ids.push(t.id); }); });
    else Object.values(tasksByMonth).forEach(function(l) { l.forEach(function(t) { ids.push(t.id); }); });
    if (!ids.length) { alert("Aucune intervention a supprimer"); return; }
    setTasks(function(prev) { return prev.filter(function(t) { return ids.indexOf(t.id) === -1; }); });
  }

  async function optimizeView() {
    if (!curEmp) return;
    let toOpt = [];
    if (viewMode === "day") toOpt = tasksForDay.slice();
    else if (viewMode === "week") Object.values(tasksByWeek).forEach(function(l) { toOpt = toOpt.concat(l); });
    else Object.values(tasksByMonth).forEach(function(l) { toOpt = toOpt.concat(l); });
    if (toOpt.length < 2) { alert("Pas assez de chantiers"); return; }
    setIsAdding(true);
    try {
      const byDate = {};
      toOpt.forEach(function(t) { if (!byDate[t.date]) byDate[t.date] = []; byDate[t.date].push(t); });
      const optimized = [];
      const dateKeys = Object.keys(byDate);
      for (let di = 0; di < dateKeys.length; di++) {
        const dayTasks = byDate[dateKeys[di]];
        const sObjs = dayTasks.map(function(t) { return sites.find(function(s) { return s.id === t.siteId; }); }).filter(Boolean);
        const coords = await batchGeocode(sObjs);
        var optSeed = null;
        if (!planTourStartSiteId) {
          if (planTourStartGeo.lat != null && planTourStartGeo.lon != null) {
            optSeed = { lat: Number(planTourStartGeo.lat), lon: Number(planTourStartGeo.lon) };
          } else if (planTourStartGeo.address && planTourStartGeo.address.trim()) {
            try {
              var gOpt = await geocodeAddress(planTourStartGeo.address.trim());
              optSeed = { lat: gOpt.lat, lon: gOpt.lon };
            } catch (eOpt) {}
          }
        }
        const sorted = orderSitesWithAnchors(sObjs, coords, {
          startSiteId: planTourStartSiteId || null,
          endSiteId: planTourEndSiteId || null,
          startSeedCoords: optSeed || undefined
        });
        const orderedTasks = sorted.map(function(s) { return dayTasks.find(function(t) { return t.siteId === s.id; }); }).filter(Boolean)
          .concat(dayTasks.filter(function(t) { return !sorted.find(function(s) { return s.id === t.siteId; }); }));
        let cursor = null;
        let prev = null;
        for (let ti = 0; ti < orderedTasks.length; ti++) {
          const t = orderedTasks[ti];
          const site = sites.find(function(s) { return s.id === t.siteId; });
          const svc = services.find(function(s) { return s.id === t.serviceId; });
          const totalSlot2 = svc ? svc.duration : 60;
          let startMin;
          let trav2 = 0;
          var earliestBound;
          if (cursor === null) {
            startMin = parseHM(t.startAt || planStartTime || "06:00");
            earliestBound = startMin;
            trav2 = 0;
          } else {
            trav2 = await travelMinutes(prev, site, t.transport || planTransport);
            startMin = cursor + trav2;
            earliestBound = startMin;
          }
          startMin = adjustStartForTimeConstraint(
            startMin,
            totalSlot2,
            earliestBound,
            site ? site.timeConstraint || null : null
          );
          const effDur2 = totalSlot2;
          cursor = startMin + effDur2;
          prev = site;
          optimized.push(Object.assign({}, t, { startAt: formatHM(startMin), travelMinutes: Math.round(trav2), effectiveDuration: effDur2 }));
        }
      }
      setTasks(function(prev) { return prev.map(function(t) { const f = optimized.find(function(o) { return o.id === t.id; }); return f ? f : t; }); });
      alert("Optimise par proximite !");
    } finally {
      setIsAdding(false);
    }
  }

  function exportICS(taskList, filename) {
    if (!curEmp || !taskList.length) { alert("Aucune intervention"); return; }
    function esc(s) { return String(s).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/,/g, "\\,").replace(/;/g, "\\;"); }
    function dt(d, hm) { const p = d.split("-").map(Number); const t = hm.split(":").map(Number); return p[0] + pad2(p[1]) + pad2(p[2]) + "T" + pad2(t[0]) + pad2(t[1]) + "00"; }
    const lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlannerPro//FR", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"];
    taskList.forEach(function(t) {
      const site = sites.find(function(s) { return s.id === t.siteId; });
      const svc = services.find(function(s) { return s.id === t.serviceId; });
      const dur = svc ? svc.duration : 60;
      lines.push("BEGIN:VEVENT", "UID:" + Math.random() + "@pp", "DTSTART;TZID=Europe/Paris:" + dt(t.date, t.startAt), "DTEND;TZID=Europe/Paris:" + dt(t.date, formatHM(parseHM(t.startAt) + dur)), "SUMMARY:" + esc(curEmp.name + " - " + (site ? site.name : "")), "DESCRIPTION:" + esc(svc ? svc.name : ""), "DTSTAMP:" + new Date().toISOString().replace(/[-:]/g, "").split(".")[0] + "Z", "END:VEVENT");
    });
    lines.push("END:VCALENDAR");
    const blob = new Blob([lines.join("\r\n")], { type: "text/calendar" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
    setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
  }

  function exportPDF(taskList, title, subtitle, endTime) {
    if (!curEmp || !taskList.length) { alert("Aucune intervention"); return; }
    const byDate = {};
    taskList.forEach(function(t) { if (!byDate[t.date]) byDate[t.date] = []; byDate[t.date].push(t); });
    let rows = "";
    Object.keys(byDate).sort().forEach(function(d) {
      rows += "<tr style='background:#f5f5f5'><td colspan='5' style='font-weight:bold;padding:8px'>" + new Date(d).toLocaleDateString("fr-FR", { weekday: "long", year: "numeric", month: "long", day: "numeric" }) + "</td></tr>";
      byDate[d].sort(function(a, b) { return parseHM(a.startAt) - parseHM(b.startAt); }).forEach(function(t) {
        const site = sites.find(function(s) { return s.id === t.siteId; });
        const svc = services.find(function(s) { return s.id === t.serviceId; });
        const mode = transportLabelFr(t.transport || "car");
        rows += "<tr><td>" + t.startAt + "</td><td>" + (site ? site.name : "") + "</td><td>" + (svc ? svc.name : "") + "</td><td>" + (svc ? svc.duration : 60) + " min</td><td>" + mode + "</td></tr>";
      });
    });
    const w = window.open("", "_blank");
    if (!w) { alert("Pop-up bloque"); return; }
    w.document.write("<html><head><title>" + title + "</title><style>@page{size:A4 landscape;margin:10mm}body{font-family:Arial}table{width:100%;border-collapse:collapse;margin-top:12px}th,td{border:1px solid #ccc;padding:6px;font-size:12px}th{background:#eee}</style></head><body><h2>" + title + "</h2><p>" + subtitle + "</p><table><thead><tr><th>Heure</th><th>Chantier</th><th>Prestation</th><th>Duree</th><th>Transport</th></tr></thead><tbody>" + rows + "</tbody></table></body></html>");
    w.document.close(); w.focus(); setTimeout(function() { w.print(); }, 100);
  }

  // ── Export Word (.docx) ─────────────────────────────────────────────────────
  function exportWord(taskList, title, subtitle, endTime) {
    if (!curEmp || !taskList.length) { alert("Aucune intervention a exporter"); return; }

    var LOGO = "data:image/png;base64,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";
    var GOLD = "#B08D57";
    var DARK = "#1a1208";
    var GOLD_LIGHT = "#f9f4ec";

    function hm(min) {
      return String(Math.floor(min/60)).padStart(2,"0") + "h" + String(min%60).padStart(2,"0");
    }
    function dur2str(min) {
      if (min < 60) return min + " min";
      return Math.floor(min/60) + "h" + (min%60 > 0 ? String(min%60).padStart(2,"0") : "00");
    }

    // Grouper + trier
    var byDate = {};
    taskList.forEach(function(t) {
      if (!byDate[t.date]) byDate[t.date] = [];
      byDate[t.date].push(t);
    });
    Object.keys(byDate).forEach(function(d) {
      byDate[d].sort(function(a,b){ return parseHM(a.startAt)-parseHM(b.startAt); });
    });
    var sortedDates = Object.keys(byDate).sort();

    // Récap global
    var grandNb=0, grandWork=0, grandTravel=0;
    sortedDates.forEach(function(d) {
      byDate[d].forEach(function(t) {
        var svc = services.find(function(s){return s.id===t.serviceId;});
        grandNb++;
        grandWork   += t.effectiveDuration || (svc ? svc.duration : 60);
        grandTravel += t.travelMinutes || 0;
      });
    });

    var today = new Date().toLocaleDateString("fr-FR",{weekday:"long",year:"numeric",month:"long",day:"numeric"});
    var periodParts = planDate.split("-");
    var periodLabel = viewMode === "month"
      ? new Date(planDate+"T12:00:00").toLocaleDateString("fr-FR",{month:"long",year:"numeric"})
      : subtitle;

    var CSS = `
@page { size: 297mm 210mm; mso-page-orientation: landscape; margin: 1cm 1.5cm; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: Arial, Calibri, sans-serif;
  font-size: 9pt;
  color: ${DARK};
  background: #ffffff;
  margin: 0;
  padding: 0;
}
.page { max-width: 100%; margin: 0 auto; padding: 0.6cm 1cm; }

/* ── HEADER ── */
.doc-header {
  display: flex;
  align-items: center;
  gap: 16pt;
  padding-bottom: 8pt;
  margin-bottom: 4pt;
  border-bottom: 2pt solid ${GOLD};
}
.doc-header img { width: 52pt; height: auto; flex-shrink: 0; }
.doc-header-text {}
.doc-company {
  font-family: Georgia, "Times New Roman", serif;
  font-size: 14pt;
  font-weight: bold;
  color: ${DARK};
  letter-spacing: 1pt;
  line-height: 1.1;
}
.doc-company span { color: ${GOLD}; }
.doc-subtitle {
  font-size: 8pt;
  color: #888;
  margin-top: 2pt;
  letter-spacing: 0.3pt;
  text-transform: uppercase;
}
.doc-meta {
  margin-top: 3pt;
  font-size: 9pt;
  color: ${DARK};
}
.doc-meta b { color: ${GOLD}; }

/* ── RECAP GLOBAL ── */
.global-recap {
  background: ${GOLD_LIGHT};
  border: 1pt solid #e0c97a;
  border-left: 3pt solid ${GOLD};
  border-radius: 2pt;
  padding: 6pt 10pt;
  margin: 8pt 0;
  display: flex;
  gap: 0;
  align-items: stretch;
}
.recap-stat {
  flex: 1;
  text-align: center;
  border-right: 1pt solid #e0c97a;
  padding: 0 8pt;
}
.recap-stat:last-child { border-right: none; }
.recap-stat .val {
  display: block;
  font-size: 12pt;
  font-weight: bold;
  color: ${DARK};
  line-height: 1.2;
}
.recap-stat .lbl {
  display: block;
  font-size: 7pt;
  color: #a08050;
  text-transform: uppercase;
  letter-spacing: 0.5pt;
  margin-top: 2pt;
}

/* ── BLOC JOUR ── */
.day-block { margin-bottom: 10pt; page-break-inside: avoid; }
.day-header { margin-bottom: 4pt; padding-bottom: 3pt; border-bottom: 1.5pt solid ${GOLD}; }
.day-title {
  font-family: Georgia, "Times New Roman", serif;
  font-size: 10.5pt;
  font-weight: bold;
  color: ${DARK};
  letter-spacing: 0.5pt;
  text-transform: uppercase;
}
.day-info {
  font-size: 8pt;
  color: #666;
  margin-top: 1pt;
}
.day-info .dstart { color: ${DARK}; font-weight: bold; }
.day-info .dend   { color: ${GOLD}; font-weight: bold; }
.day-info .dlate  { color: #c62828; font-weight: bold; }

/* ── INTERVENTION ── */
.interv {
  display: flex;
  gap: 8pt;
  align-items: flex-start;
  padding: 3pt 7pt;
  margin-bottom: 3pt;
  background: #fdfcfa;
  border-left: 2.5pt solid #e0c97a;
}
.interv-time-col {
  min-width: 50pt;
  flex-shrink: 0;
  text-align: center;
}
.interv-tstart {
  font-size: 10pt;
  font-weight: bold;
  color: ${DARK};
  line-height: 1;
}
.interv-arrow { font-size: 8pt; color: #bbb; margin: 1pt 0; }
.interv-tend {
  font-size: 9.5pt;
  font-weight: bold;
  color: ${GOLD};
  line-height: 1;
}
.interv-tend-late { color: #c62828 !important; }
.interv-body { flex: 1; }
.interv-site {
  font-size: 9pt;
  font-weight: bold;
  color: ${DARK};
  line-height: 1.2;
}
.interv-addr {
  font-size: 8pt;
  color: #777;
  margin-top: 1pt;
}
.interv-tags {
  display: flex;
  gap: 8pt;
  margin-top: 2pt;
  flex-wrap: wrap;
}
.tag {
  font-size: 7.5pt;
  padding: 1pt 5pt;
  border-radius: 2pt;
  font-weight: normal;
}
.tag-prest { background: #edf5ed; color: #2e6b2e; border: 0.5pt solid #b0d4b0; }
.tag-trajet { background: #fdf3e6; color: #8a5a00; border: 0.5pt solid #e8c880; }

/* ── RÉCAP JOURNÉE ── */
.day-recap {
  display: flex;
  gap: 0;
  background: #f7f3ec;
  border: 0.5pt solid #e0c97a;
  border-radius: 2pt;
  padding: 3pt 10pt;
  margin-top: 3pt;
  font-size: 8pt;
}
.day-recap-item { flex: 1; color: #5a4000; }
.day-recap-item b { color: ${DARK}; }
.day-recap-total { font-weight: bold; color: ${DARK}; text-align: right; }

/* ── SÉPARATEUR ── */
.day-sep {
  border: none;
  border-top: 0.5pt solid #e0c97a;
  margin: 8pt 0;
}

/* ── PIED DE PAGE ── */
.doc-footer {
  margin-top: 14pt;
  padding-top: 5pt;
  border-top: 1.5pt solid ${GOLD};
  display: flex;
  justify-content: space-between;
  font-size: 8pt;
  color: #aaa;
}
.doc-footer b { color: ${GOLD}; }
`;

    var html = "";
    html += `<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office" `;
    html += `xmlns:w="urn:schemas-microsoft-com:office:word" `;
    html += `xmlns="http://www.w3.org/TR/REC-html40">`;
    html += `<head><meta charset="utf-8"><style>${CSS}</style></head><body><div class="page">`;

    // ── HEADER ──
    html += `<div class="doc-header">`;
    html += `<img src="${LOGO}" alt="Marie-Eugenie" />`;
    html += `<div class="doc-header-text">`;
    html += `<div class="doc-company">Marie<span>-</span>Eug&eacute;nie</div>`;
    html += `<div class="doc-subtitle">Planning d&apos;interventions &mdash; Nettoyage professionnel</div>`;
    html += `<div class="doc-meta">Agent : <b>${curEmp.name}</b> &nbsp;&bull;&nbsp; P&eacute;riode : <b>${periodLabel}</b></div>`;
    html += `</div></div>`;

    // ── RECAP GLOBAL ──
    html += `<div class="global-recap">`;
    html += `<div class="recap-stat"><span class="val">${grandNb}</span><span class="lbl">Interventions</span></div>`;
    html += `<div class="recap-stat"><span class="val">${dur2str(grandWork)}</span><span class="lbl">Sur place</span></div>`;
    html += `<div class="recap-stat"><span class="val">${dur2str(grandTravel)}</span><span class="lbl">Trajets</span></div>`;
    html += `<div class="recap-stat"><span class="val">${dur2str(grandWork+grandTravel)}</span><span class="lbl">Total</span></div>`;
    html += `<div class="recap-stat"><span class="val">${sortedDates.length}</span><span class="lbl">Jours</span></div>`;
    html += `</div>`;

    // ── HELPER : génère le HTML d'un bloc jour ──────────────────────────────
    function renderDayBlock(d) {
      var list2 = byDate[d];
      var dateObj = new Date(d+"T12:00:00");
      var dayLabel = dateObj.toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"});
      dayLabel = dayLabel.charAt(0).toUpperCase() + dayLabel.slice(1);

      var dayWork=0, dayTravel=0;
      list2.forEach(function(t) {
        var svc=services.find(function(s){return s.id===t.serviceId;});
        dayWork   += t.effectiveDuration||(svc?svc.duration:60);
        dayTravel += t.travelMinutes||0;
      });

      var firstT  = list2[0];
      var lastT   = list2[list2.length-1];
      var lastSvc = services.find(function(s){return s.id===lastT.serviceId;});
      var lastDur = lastT.effectiveDuration||(lastSvc?lastSvc.duration:60);
      var dayEnd  = parseHM(lastT.startAt)+lastDur;
      var isLateDay = dayEnd > parseHM(endTime || "12:00");

      var bh = "";
      bh += `<div class="day-block">`;
      bh += `<div class="day-header">`;
      bh += `<div class="day-title">${dayLabel}</div>`;
      bh += `<div class="day-info">`;
      bh += `D&eacute;but&nbsp;<span class="dstart">${firstT.startAt}</span>`;
      bh += `&nbsp;&bull;&nbsp;Fin&nbsp;<span class="${isLateDay?"dlate":"dend"}">${hm(dayEnd)}${isLateDay?" &#9888;":""}</span>`;
      bh += `&nbsp;&bull;&nbsp;<b>${list2.length}</b> intervention${list2.length>1?"s":""}`;
      bh += `</div></div>`;

      list2.forEach(function(t) {
        var site    = sites.find(function(s){return s.id===t.siteId;});
        var svc     = services.find(function(s){return s.id===t.serviceId;});
        var dur     = t.effectiveDuration||(svc?svc.duration:60);
        var trav    = t.travelMinutes||0;
        var endMin  = parseHM(t.startAt)+dur;
        var isLate  = endMin>parseHM(endTime || "12:00");
        var siteName= site?site.name:"";
        var addr    = (site&&site.address&&site.address!==siteName)?site.address:"";
        var svcName = svc?svc.name:"";

        bh += `<div class="interv">`;
        bh += `<div class="interv-time-col">`;
        bh += `<div class="interv-tstart">${t.startAt}</div>`;
        bh += `<div class="interv-arrow">&#8595;</div>`;
        bh += `<div class="interv-tend${isLate?" interv-tend-late":""}">${hm(endMin)}${isLate?"&nbsp;&#9888;":""}</div>`;
        bh += `</div>`;
        bh += `<div class="interv-body">`;
        bh += `<div class="interv-site">&#128205; ${siteName}</div>`;
        if (addr) bh += `<div class="interv-addr">${addr}</div>`;
        bh += `<div class="interv-tags">`;
        if (svcName) bh += `<span class="tag tag-prest">&#129529; ${svcName} &mdash; ${dur2str(dur)}</span>`;
        if (trav>0)  bh += `<span class="tag tag-trajet">${transportEmoji(t.transport || "car")} Trajet : ${trav} min</span>`;
        bh += `</div></div></div>`;
      });

      bh += `<div class="day-recap">`;
      bh += `<div class="day-recap-item">&#129529; Sur place&nbsp;: <b>${dur2str(dayWork)}</b></div>`;
      bh += `<div class="day-recap-item">&#128663; Trajets&nbsp;: <b>${dur2str(dayTravel)}</b></div>`;
      bh += `<div class="day-recap-total">Total&nbsp;: ${dur2str(dayWork+dayTravel)}</div>`;
      bh += `</div>`;
      bh += `</div>`; // .day-block
      return bh;
    }

    // ── BLOCS PAR JOUR – 2 colonnes côte à côte ──────────────────────────────
    for (var pairIdx = 0; pairIdx < sortedDates.length; pairIdx += 2) {
      html += `<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;table-layout:fixed;width:100%">`;
      html += `<tr>`;
      html += `<td style="width:50%;vertical-align:top;padding-right:7pt;">`;
      html += renderDayBlock(sortedDates[pairIdx]);
      html += `</td>`;
      html += `<td style="width:50%;vertical-align:top;padding-left:7pt;border-left:1pt solid #e0c97a;">`;
      if (pairIdx+1 < sortedDates.length) html += renderDayBlock(sortedDates[pairIdx+1]);
      html += `</td>`;
      html += `</tr></table>`;
      if (pairIdx+2 < sortedDates.length) html += `<hr class="day-sep">`;
    }

    // ── PIED DE PAGE ──
    html += `<div class="doc-footer">`;
    html += `<span><b>Marie-Eug&eacute;nie</b> &mdash; Nettoyage Professionnel &mdash; Depuis 1975</span>`;
    html += `<span>G&eacute;n&eacute;r&eacute; le ${today}</span>`;
    html += `</div>`;

    html += `</div></body></html>`;

    var blob = new Blob(["\uFEFF"+html], { type: "application/msword;charset=utf-8" });
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement("a");
    var p    = planDate.split("-");
    a.href   = url;
    a.download = "planning_" + curEmp.name + "_" + (viewMode==="month"?p[0]+"-"+p[1]:viewMode==="week"?"sem_"+weekDates[0]:planDate) + ".doc";
    a.click();
    setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
  }


  function exportView(fmt) {
    if (!curEmp) { alert("Selectionne un salarie"); return; }
    var list = []; var fn = ""; var title = "Planning Interventions"; var sub = curEmp.name;
    if (viewMode === "day") {
      list = tasksForDay;
      fn = "planning_" + curEmp.name + "_" + planDate + "." + fmt;
      sub += " - " + new Date(planDate + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
    } else if (viewMode === "week") {
      Object.values(tasksByWeek).forEach(function(l) { list = list.concat(l); });
      fn = "planning_" + curEmp.name + "_sem_" + weekDates[0] + "." + fmt;
      sub += " - Semaine du " + weekDates[0] + " au " + weekDates[6];
    } else {
      Object.values(tasksByMonth).forEach(function(l) { list = list.concat(l); });
      var p = planDate.split("-");
      fn = "planning_" + curEmp.name + "_" + p[0] + "-" + p[1] + "." + fmt;
      sub += " - " + new Date(planDate + "T12:00:00").toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
    }
    if (fmt === "ics") exportICS(list, fn);
    else if (fmt === "word") exportWord(list, title, sub, planEndTime);
    else exportPDF(list, title, sub, planEndTime);
  }

  function addEmp() { if (!newEmpName.trim()) return; setEmployees(function(p) { return p.concat([{ id: String(Math.random()), name: newEmpName.trim() }]); }); setNewEmpName(""); }
  function delEmp(id) { setEmployees(function(p) { return p.filter(function(e) { return e.id !== id; }); }); setTasks(function(p) { return p.filter(function(t) { return t.employeeId !== id; }); }); if (planEmpId === id) setPlanEmpId(""); }
  function addSite() {
    if (!newSiteName.trim()) return;
    var o = { id: String(Math.random()), name: newSiteName.trim(), address: (newSiteGeo.address || "").trim() };
    if (newSiteGeo.lat != null && newSiteGeo.lon != null && Number.isFinite(Number(newSiteGeo.lat)) && Number.isFinite(Number(newSiteGeo.lon))) {
      o.lat = Number(newSiteGeo.lat);
      o.lon = Number(newSiteGeo.lon);
    }
    setSites(function(p) { return p.concat([o]); });
    setNewSiteName("");
    setNewSiteGeo({ address: "", lat: undefined, lon: undefined });
  }
  function addGrp() { if (!newGrpName.trim()) return; setGroups(function(p) { return p.concat([{ id: String(Math.random()), name: newGrpName.trim() }]); }); setNewGrpName(""); }
  function updSite(id, patch) { setSites(function(p) { return p.map(function(s) { return s.id === id ? Object.assign({}, s, patch) : s; }); }); }
  function delSite(id) { setSites(function(p) { return p.filter(function(s) { return s.id !== id; }); }); setTasks(function(p) { return p.filter(function(t) { return t.siteId !== id; }); }); setFreqs(function(p) { return p.filter(function(f) { return f.siteId !== id; }); }); setPlanSiteIds(function(p) { return p.filter(function(s) { return s !== id; }); }); }
  function delGrp(id) { setGroups(function(p) { return p.filter(function(g) { return g.id !== id; }); }); setSites(function(p) { return p.map(function(s) { return s.groupId === id ? Object.assign({}, s, { groupId: undefined }) : s; }); }); }
  function updFreq(id, patch) {
    setFreqs(function(prev) {
      const i = prev.findIndex(function(f) { return f.siteId === id; });
      if (i === -1) return prev.concat([Object.assign({ siteId: id, type: "once", months: [], days: [] }, patch)]);
      return prev.map(function(f, idx) { return idx === i ? Object.assign({}, f, patch) : f; });
    });
  }
  function addSvc() { if (!newSvcName.trim()) return; setServices(function(p) { return p.concat([{ id: String(Math.random()), name: newSvcName.trim(), duration: parseInt(newSvcDur, 10) || 60 }]); }); setNewSvcName(""); setNewSvcDur("60"); }
  function delSvc(id) { setServices(function(p) { return p.filter(function(s) { return s.id !== id; }); }); setTasks(function(p) { return p.filter(function(t) { return t.serviceId !== id; }); }); }

  const filteredEmps = employees.filter(function(e) { return e.name.toLowerCase().indexOf(searchEmps.toLowerCase()) !== -1; });
  const filteredSites = sites.filter(function(s) { return s.name.toLowerCase().indexOf(searchSites.toLowerCase()) !== -1; });
  const filteredSvcs = services.filter(function(s) { return s.name.toLowerCase().indexOf(searchSvcs.toLowerCase()) !== -1; });

  const groupedSites = useMemo(function() {
    const g = { ungrouped: [] };
    filteredSites.forEach(function(s) {
      if (s.groupId) { if (!g[s.groupId]) g[s.groupId] = []; g[s.groupId].push(s); }
      else g.ungrouped.push(s);
    });
    return g;
  }, [filteredSites]);

  function renderPlanningRight() {
    if (!curEmp) return <div style={{ fontSize: 14, color: "#666" }}>Selectionne un salarie.</div>;
    if (viewMode === "day") {
      if (!tasksForDay.length) return <div style={{ fontSize: 14, color: "#666" }}>Aucune intervention ce jour.</div>;
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 420, overflow: "auto" }}>
          {tasksForDay.map(function(t, idx) {
            const site = sites.find(function(s) { return s.id === t.siteId; });
            const svc = services.find(function(s) { return s.id === t.serviceId; });
            return (
              <div key={t.id} style={{ border: "1px solid " + THEME.cardBorder, borderLeft: "3px solid " + THEME.gold, borderRadius: THEME.radiusSm, padding: "9px 12px", background: THEME.card, boxShadow: THEME.shadow, marginBottom: 5 }}>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                  <AppInput type="time" style={{ width: 90 }} value={t.startAt} onChange={function(e) { updateTask(t.id, { startAt: e.target.value }); }} />
                  <AppSelect value={t.siteId} onChange={function(e) { updateTask(t.id, { siteId: e.target.value }); }}>
                    {sites.map(function(s) { return <option key={s.id} value={s.id}>{s.name}</option>; })}
                  </AppSelect>
                  <AppSelect value={t.serviceId} onChange={function(e) { updateTask(t.id, { serviceId: e.target.value }); }}>
                    {services.map(function(s) { return <option key={s.id} value={s.id}>{s.name}</option>; })}
                  </AppSelect>
                  <AppSelect value={t.transport || "car"} onChange={function(e) { updateTask(t.id, { transport: e.target.value }); }}>
                    <option value="car">Voiture</option>
                    <option value="pt">Transports</option>
                    <option value="bike">Velo</option>
                    <option value="walk">A pied</option>
                  </AppSelect>
                  <AppButton onClick={function() { deleteTask(t.id); }}>Supprimer</AppButton>
                </div>
                <div style={{ fontSize: 12, color: "#777", marginTop: 4 }}>{site ? site.name : ""} - {svc ? svc.duration + " min" : "?"}</div>
                {idx > 0 && <div style={{ fontSize: 11, color: "#aaa", marginTop: 2 }}>Trajet calcule depuis le chantier precedent</div>}
              </div>
            );
          })}
        </div>
      );
    }
    if (viewMode === "week") {
      return (
        <div style={{ display: "flex", gap: 8, overflowX: "auto" }}>
          {weekDates.map(function(d, i) {
            const list = tasksByWeek[d] || [];
            return (
              <div key={d} style={{ flex: "0 0 180px", border: "1px solid " + THEME.cardBorder, borderTop: "3px solid " + THEME.gold, borderRadius: THEME.radiusSm, padding: 10, background: THEME.card, boxShadow: THEME.shadow }}>
                <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 4, color: "#333" }}>{DAYS_LABELS[i]} {d.slice(8, 10)}/{d.slice(5, 7)}</div>
                {!list.length && <div style={{ fontSize: 12, color: "#aaa" }}>-</div>}
                {list.map(function(t) {
                  const site = sites.find(function(s) { return s.id === t.siteId; });
                  const svc = services.find(function(s) { return s.id === t.serviceId; });
                  return (
                    <div key={t.id} style={{ fontSize: 11, padding: "5px 8px", borderRadius: 5, background: THEME.goldPale, marginBottom: 4, border: "1px solid " + THEME.goldBorder }}>
                      <div style={{ fontWeight: 600, marginBottom: 2 }}>{t.startAt} - {site ? site.name : ""}</div>
                      <div style={{ fontSize: 10, color: "#666", marginBottom: 3 }}>{svc ? svc.name : ""}</div>
                      <AppButton onClick={function() { deleteTask(t.id); }} style={{ padding: "3px 7px", fontSize: 10, background: "#FFF0F0", border: "1px solid #FFBCBC", color: THEME.danger, borderRadius: 5 }}>Supprimer</AppButton>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      );
    }
    if (viewMode === "month") {
      const dates = Object.keys(tasksByMonth).sort();
      if (!dates.length) return <div style={{ fontSize: 14, color: "#666" }}>Aucune intervention ce mois.</div>;
      return (
        <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 520, overflow: "auto" }}>
          {dates.map(function(d) {
            const list = tasksByMonth[d];
            const totalTravel = list.reduce(function(acc2, t) { return acc2 + (t.travelMinutes || 0); }, 0);
            const totalWork = list.reduce(function(acc2, t) {
              const svc = services.find(function(s) { return s.id === t.serviceId; });
              const slot = svc ? svc.duration : 60;
              const trav = t.travelMinutes || 0;
              return acc2 + (t.effectiveDuration || slot); // prestation complete
            }, 0);
            const totalMin = totalWork + totalTravel;
            const lastTask = list[list.length - 1];
            const lastEnd = lastTask ? parseHM(lastTask.startAt) + (function() { const sv = services.find(function(s) { return s.id === lastTask.serviceId; }); return sv ? sv.duration : 60; }()) : 0;
            const overTime = lastEnd > parseHM("12:00");
            return (
              <div key={d} style={{ border: "1px solid " + (overTime ? "#E57373" : THEME.cardBorder), borderLeft: "3px solid " + (overTime ? THEME.danger : THEME.gold), borderRadius: THEME.radiusSm, padding: "10px 12px", background: overTime ? "#FFF5F5" : THEME.card, boxShadow: THEME.shadow }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                  <div style={{ fontSize: 12, fontWeight: 800, color: "#1a1a2e" }}>
                    <span style={{ fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, fontSize: 12 }}>{DAYS_LABELS[weekdayIdx(d)]} {d.slice(8, 10)}/{d.slice(5, 7)}</span>
                    {overTime && <span style={{ marginLeft: 6, fontSize: 10, color: "#dc2626", fontWeight: 700 }}>⚠ depasse 12h</span>}
                  </div>
                  <div style={{ fontSize: 10, color: "#555", textAlign: "right" }}>
                    <span style={{ color: "#0f3460", fontWeight: 600 }}>{Math.floor(totalWork/60)}h{String(totalWork%60).padStart(2,"0")} sur place</span>
                    {totalTravel > 0 && <span style={{ color: "#a87a3d", marginLeft: 4 }}>+ {totalTravel}min trajet</span>}
                    <span style={{ color: "#888", marginLeft: 4 }}>= {Math.floor(totalMin/60)}h{String(totalMin%60).padStart(2,"0")} total</span>
                  </div>
                </div>
                {list.map(function(t, idx) {
                  const site = sites.find(function(s) { return s.id === t.siteId; });
                  const svc = services.find(function(s) { return s.id === t.serviceId; });
                  const mode = transportLabelFr(t.transport || "car");
                  const totalSlot3 = svc ? svc.duration : 60;
                  const travel = t.travelMinutes || 0;
                  const durMin = t.effectiveDuration || totalSlot3; // prestation complete
                  const endMin = parseHM(t.startAt) + durMin;
                  const isLate = endMin > parseHM(planEndTime);
                  return (
                    <div key={t.id}>
                      {travel > 0 && (
                        <div style={{ display: "flex", alignItems: "center", gap: 4, padding: "2px 8px", margin: "2px 0", fontSize: 10, color: THEME.gold, background: THEME.goldPale, borderRadius: 4, border: "1px solid " + THEME.goldBorder }}>
                          <span>{transportEmoji(t.transport || "car")}</span>
                          <span style={{ fontWeight: 600 }}>{travel} min trajet</span>
                          {idx === 0 && <span style={{ color: "#bbb" }}>(depuis le domicile)</span>}
                          {travel === 10 && <span style={{ color: "#ccc", fontSize: 9 }}>(estime)</span>}
                        </div>
                      )}
                      <div style={{ fontSize: 11, marginBottom: 2, padding: "6px 8px", background: isLate ? "#fff0f0" : "#f4f6ff", borderRadius: 6, border: "1px solid " + (isLate ? "#fca5a5" : "#dde3f0"), display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <div>
                          <span style={{ fontWeight: 700, color: "#0f3460" }}>{t.startAt}</span>
                          <span style={{ color: "#999" }}> → </span>
                          <span style={{ fontWeight: 600, color: isLate ? "#dc2626" : "#333" }}>{formatHM(endMin)}</span>
                          <span style={{ color: "#111", marginLeft: 6, fontWeight: 600 }}>{site ? site.name : ""}</span>
                          <span style={{ color: "#888", marginLeft: 4, fontSize: 10 }}>
                            ({svc ? svc.name : ""} · <span style={{ color: "#0f3460" }}>{durMin}min sur place</span>{travel > 0 ? " · " + transportEmoji(t.transport || "car") + " " + travel + " min" : ""} · {mode})
                          </span>
                        </div>
                        <AppButton onClick={function() { deleteTask(t.id); }} style={{ padding: "2px 6px", fontSize: 10, background: "#FFF0F0", border: "1px solid #FFBCBC", color: THEME.danger, marginLeft: 6, flexShrink: 0, borderRadius: 4 }}>×</AppButton>
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      );
    }
    return null;
  }

  return (
    <div style={{ minHeight: "100vh", background: THEME.bg, color: THEME.primary, fontFamily: "Arial, Calibri, sans-serif", position: "relative" }}>
      <img
        src={LOGO}
        alt=""
        aria-hidden="true"
        style={{
          position: "fixed",
          top: "50%",
          left: "50%",
          transform: "translate(-50%, -50%)",
          width: "min(55vw, 640px)",
          maxWidth: "90vw",
          height: "auto",
          opacity: 0.055,
          pointerEvents: "none",
          zIndex: 0,
          userSelect: "none",
        }}
      />

      {/* ── Overlay calcul ── */}
      {isAdding && (
        <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.65)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ background: THEME.card, borderRadius: 16, padding: "36px 48px", textAlign: "center", boxShadow: THEME.shadowMd, borderTop: "3px solid " + THEME.gold }}>
            <div style={{ width: 48, height: 48, margin: "0 auto 14px", border: "3px solid " + THEME.goldBorder, borderTopColor: THEME.gold, borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
            <div style={{ fontWeight: 700, fontSize: 15, color: THEME.primary, marginBottom: 5 }}>Calcul en cours...</div>
            <div style={{ fontSize: 12, color: THEME.textMuted }}>Geocodage et optimisation des tournees</div>
          </div>
        </div>
      )}

      {/* ── HEADER ── */}
      <header style={{
        background: THEME.dark,
        borderBottom: "2px solid " + THEME.gold,
        position: "sticky", top: 0, zIndex: 100,
        boxShadow: "0 2px 12px rgba(0,0,0,0.3)",
        isolation: "isolate"
      }}>
        <div style={{ maxWidth: 1200, margin: "0 auto", padding: "0 20px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 58 }}>
          {/* Logo + nom */}
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <img src={LOGO} alt="Marie-Eugenie" style={{ height: 38, width: "auto", filter: "brightness(1.1)" }} />
            <div>
              <div style={{ fontSize: 14, fontWeight: 700, color: "#FFFFFF", letterSpacing: 1, lineHeight: 1.1 }}>Marie-Eug&eacute;nie</div>
              <div style={{ fontSize: 10, color: THEME.gold, letterSpacing: 1.5, textTransform: "uppercase", fontWeight: 500 }}>Planning Interventions</div>
            </div>
          </div>
          {/* Onglets nav */}
          <nav style={{ display: "flex", gap: 2 }}>
            {[["planning","Planning"],["employees","Salariés"],["sites","Chantiers"],["services","Prestations"]].map(function(item) {
              var isActive = tab === item[0];
              return (
                <button key={item[0]} onClick={function() { setTab(item[0]); }} style={{
                  padding: "8px 18px",
                  background: isActive ? THEME.gold : "transparent",
                  color: isActive ? "#FFFFFF" : "rgba(255,255,255,0.65)",
                  border: "none",
                  borderRadius: 6,
                  cursor: "pointer",
                  fontSize: 13,
                  fontWeight: isActive ? 600 : 400,
                  letterSpacing: 0.3,
                  transition: "background 0.2s, color 0.2s",
                }}>
                  {item[1]}
                </button>
              );
            })}
          </nav>
        </div>
      </header>

      <div style={{ maxWidth: 1200, margin: "0 auto", padding: "20px 20px", position: "relative", zIndex: 1 }}>
        <style>{`
          @keyframes spin { to { transform: rotate(360deg); } }
          button:hover { opacity: 0.88; }
          input:focus, select:focus { border-color: ${THEME.gold} !important; box-shadow: 0 0 0 2px ${THEME.goldPale}; }
        `}</style>

        {tab === "planning" && (
          <div style={{ display: "grid", gridTemplateColumns: "1.2fr 0.8fr", gap: 16 }}>
            <AppBox title="Creer une intervention" right={
              <div style={{ display: "flex", gap: 6 }}>
                <AppButton onClick={optimizeView}>Optimiser</AppButton>
                <AppButton kind="danger" onClick={deleteView}>Tout supprimer</AppButton>
                <AppButton onClick={function() { exportView("ics"); }}>ICS</AppButton>
                <AppButton onClick={function() { exportView("pdf"); }}>PDF</AppButton>
                <AppButton onClick={function() { exportView("word"); }} kind="gold">Word ↓</AppButton>
              </div>
            }>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
                <div>
                  <div style={{ fontSize: 12, color: "#555", marginBottom: 2 }}>
                    {viewMode === "month" ? "Mois" : viewMode === "week" ? "Semaine" : "Date"}
                  </div>
                  {viewMode === "month" ? (
                    <AppInput
                      type="month"
                      value={planDate.substring(0, 7)}
                      onChange={function(e) { setPlanDate(e.target.value + "-01"); }}
                    />
                  ) : (
                    <AppInput type="date" value={planDate} onChange={function(e) { setPlanDate(e.target.value); }} />
                  )}
                </div>
                <div>
                  <div style={{ fontSize: 12, color: "#555", marginBottom: 2 }}>Salarie</div>
                  <AppSelect value={planEmpId} onChange={function(e) { setPlanEmpId(e.target.value); }}>
                    <option value="">-- choisir --</option>
                    {employees.map(function(e) { return <option key={e.id} value={e.id}>{e.name}</option>; })}
                  </AppSelect>
                </div>
              </div>

              <div style={{ display: "flex", gap: 16, marginBottom: 10 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>Groupes</div>
                  <div style={{ border: "1px solid " + THEME.grid, borderRadius: 10, padding: 8, maxHeight: 90, overflow: "auto", background: "#fafafa" }}>
                    {!groups.length ? <div style={{ fontSize: 12, color: "#aaa" }}>Aucun groupe</div> :
                      groups.map(function(g) {
                        return (
                          <label key={g.id} style={{ display: "flex", alignItems: "center", gap: 6, padding: 3, cursor: "pointer", fontSize: 13 }}>
                            <input type="checkbox" checked={planGrpIds.indexOf(g.id) !== -1} onChange={function(e) { if (e.target.checked) setPlanGrpIds(function(p) { return p.concat([g.id]); }); else setPlanGrpIds(function(p) { return p.filter(function(id) { return id !== g.id; }); }); }} />
                            {g.name}
                          </label>
                        );
                      })}
                  </div>
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>Chantiers</div>
                  <div style={{ border: "1px solid " + THEME.grid, borderRadius: 10, padding: 8, maxHeight: 90, overflow: "auto", background: "#fafafa" }}>
                    {sites.map(function(s) {
                      return (
                        <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 6, padding: 3, cursor: "pointer", fontSize: 13 }}>
                          <input type="checkbox" checked={planSiteIds.indexOf(s.id) !== -1} onChange={function(e) { if (e.target.checked) setPlanSiteIds(function(p) { return p.concat([s.id]); }); else setPlanSiteIds(function(p) { return p.filter(function(id) { return id !== s.id; }); }); }} />
                          {s.name}
                        </label>
                      );
                    })}
                  </div>
                </div>
              </div>

              {sitesNeedingSvc.length > 0 && (
                <div style={{ marginBottom: 10, border: "1px solid " + THEME.grid, borderRadius: 10, padding: 10, background: "#fafafa" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 6 }}>Prestations a selectionner</div>
                  {sitesNeedingSvc.map(function(sid) {
                    const site = sites.find(function(s) { return s.id === sid; });
                    return (
                      <div key={sid} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 6 }}>
                        <div style={{ fontSize: 12, fontWeight: 600, minWidth: 120 }}>{site ? site.name : sid}</div>
                        <AppSelect value={siteServices[sid] || ""} onChange={function(e) { setSiteServices(function(p) { const n = Object.assign({}, p); n[sid] = e.target.value; return n; }); }} style={{ flex: 1 }}>
                          <option value="">-- prestation --</option>
                          {services.map(function(s) { return <option key={s.id} value={s.id}>{s.name} ({s.duration} min)</option>; })}
                        </AppSelect>
                      </div>
                    );
                  })}
                </div>
              )}

              <div style={{ display: "flex", gap: 8, alignItems: "flex-start", marginBottom: 10, flexWrap: "wrap" }}>
                <div>
                  <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>Heure de début</div>
                  <AppInput
                    type="time"
                    value={planStartTime}
                    onChange={function(e) { setPlanStartTime(e.target.value); }}
                    style={{ width: 90 }}
                  />
                </div>
                <div>
                  <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>Heure de fin</div>
                  <AppInput
                    type="time"
                    value={planEndTime}
                    onChange={function(e) { setPlanEndTime(e.target.value); }}
                    style={{ width: 90 }}
                  />
                </div>
                <div>
                  <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>Transport</div>
                  <AppSelect value={planTransport} onChange={function(e) { setPlanTransport(e.target.value); }}>
                    <option value="car">Voiture</option>
                    <option value="pt">Transports en commun</option>
                    <option value="bike">Velo</option>
                    <option value="walk">A pied</option>
                  </AppSelect>
                </div>
                <AppButton kind="primary" onClick={addTask} disabled={isAdding} style={{ marginTop: 22 }}>
                  {isAdding ? "Calcul..." : viewMode === "month" ? "Planifier le mois" : "Ajouter au planning"}
                </AppButton>
              </div>
              {viewMode === "month" && !isAdding && (
                <div style={{ fontSize: 11, color: THEME.gold, marginTop: 6, padding: "7px 12px", background: THEME.goldPale, borderRadius: THEME.radiusSm, border: "1px solid " + THEME.goldBorder }}>
                  Distribue les chantiers sur tous les jours ouvres du mois selon leur frequence. Sans frequence configuree = 4 passages/mois.
                </div>
              )}

              <div style={{ marginTop: 14, paddingTop: 12, borderTop: "1px dashed " + THEME.grid }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: THEME.primary, marginBottom: 8, letterSpacing: 0.3 }}>Tournee (optionnel)</div>
                <div style={{ fontSize: 11, color: THEME.textMuted, marginBottom: 8 }}>
                  Point de depart : premier trajet depuis ce lieu. Laisse vide pour le comportement habituel. L&apos;optimiseur en tient compte.
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
                  <div>
                    <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>Depart — chantier</div>
                    <AppSelect value={planTourStartSiteId} onChange={function(e) { setPlanTourStartSiteId(e.target.value); }} style={{ minWidth: 200 }}>
                      <option value="">(aucun)</option>
                      {sites.map(function(s) { return <option key={s.id} value={s.id}>{s.name}</option>; })}
                    </AppSelect>
                  </div>
                  <div style={{ flex: "1 1 220px", minWidth: 200 }}>
                    <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>ou adresse libre</div>
                    <AddressAutocomplete
                      value={planTourStartGeo.address}
                      onChange={setPlanTourStartGeo}
                      placeholder="Adresse de depart"
                      theme={{ grid: THEME.grid, pale: THEME.pale, text: THEME.primary, textMuted: THEME.textMuted, shadow: THEME.shadow }}
                    />
                  </div>
                  <div>
                    <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>Arrivee — chantier (base)</div>
                    <AppSelect value={planTourEndSiteId} onChange={function(e) { setPlanTourEndSiteId(e.target.value); }} style={{ minWidth: 200 }}>
                      <option value="">(aucun)</option>
                      {sites.map(function(s) { return <option key={s.id} value={s.id}>{s.name}</option>; })}
                    </AppSelect>
                  </div>
                  <div style={{ flex: "1 1 200px", minWidth: 180 }}>
                    <div style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>Fin de tournee — adresse (reserve)</div>
                    <AppInput placeholder="Pour evolution ulterieure" value={planTourEndAddress} onChange={function(e) { setPlanTourEndAddress(e.target.value); }} />
                  </div>
                </div>
              </div>

              <div style={{ borderTop: "1px solid " + THEME.grid, paddingTop: 10, marginTop: 4 }}>
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <span style={{ fontSize: 12, color: "#555" }}>Vue :</span>
                  <AppButton kind={viewMode === "day" ? "primary" : "default"} onClick={function() { setViewMode("day"); }}>Jour</AppButton>
                  <AppButton kind={viewMode === "week" ? "primary" : "default"} onClick={function() { setViewMode("week"); }}>Semaine</AppButton>
                  <AppButton kind={viewMode === "month" ? "primary" : "default"} onClick={function() { setViewMode("month"); }}>Mois</AppButton>
                </div>
              </div>
            </AppBox>

            <AppBox title={"Vue " + (viewMode === "day" ? "journee" : viewMode === "week" ? "semaine" : "mois") + (curEmp ? " (" + curEmp.name + ")" : "")}>
              {renderPlanningRight()}
            </AppBox>
          </div>
        )}

        {tab === "employees" && (
          <AppBox title="Salaries">
            <AppInput placeholder="Rechercher..." value={searchEmps} onChange={function(e) { setSearchEmps(e.target.value); }} style={{ marginBottom: 12, width: "100%" }} />
            <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
              <AppInput placeholder="Nom du salarie" value={newEmpName} onChange={function(e) { setNewEmpName(e.target.value); }} />
              <AppButton kind="primary" onClick={addEmp}>Ajouter</AppButton>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {filteredEmps.map(function(e) {
                return (
                  <div key={e.id} style={{ display: "flex", gap: 8 }}>
                    <AppInput value={e.name} onChange={function(ev) { setEmployees(function(p) { return p.map(function(x) { return x.id === e.id ? Object.assign({}, x, { name: ev.target.value }) : x; }); }); }} style={{ flex: 1 }} />
                    <AppButton onClick={function() { delEmp(e.id); }}>Supprimer</AppButton>
                  </div>
                );
              })}
              {!filteredEmps.length && <div style={{ fontSize: 14, color: "#777" }}>Aucun salarie.</div>}
            </div>
          </AppBox>
        )}

        {tab === "sites" && (
          <AppBox title="Chantiers">
            <ImportPanel sites={sites} setSites={setSites} employees={employees} setEmployees={setEmployees} services={services} setServices={setServices} setSiteFrequencies={setFreqs} siteFrequencies={freqs} />
            <AppInput placeholder="Rechercher..." value={searchSites} onChange={function(e) { setSearchSites(e.target.value); }} style={{ marginBottom: 12, width: "100%" }} />
            <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
              <AppInput placeholder="Nouveau groupe" value={newGrpName} onChange={function(e) { setNewGrpName(e.target.value); }} />
              <AppButton kind="primary" onClick={addGrp}>Creer groupe</AppButton>
            </div>
            {groups.length > 0 && (
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 12 }}>
                {groups.map(function(g) {
                  return (
                    <div key={g.id} style={{ border: "1px solid " + THEME.goldBorder, background: THEME.goldPale, padding: "4px 10px", borderRadius: 6, display: "flex", alignItems: "center", gap: 6, fontSize: 12 }}>
                      {g.name}
                      <AppButton onClick={function() { delGrp(g.id); }} style={{ padding: "1px 6px", fontSize: 11, background: "#FFF0F0", border: "none", color: THEME.danger, borderRadius: 3, cursor: "pointer" }}>×</AppButton>
                    </div>
                  );
                })}
              </div>
            )}
            <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
              <AppInput placeholder="Nom du chantier" value={newSiteName} onChange={function(e) { setNewSiteName(e.target.value); }} />
              <AddressAutocomplete
                value={newSiteGeo.address}
                onChange={setNewSiteGeo}
                placeholder="Adresse complete"
                theme={{ grid: THEME.grid, pale: THEME.pale, text: THEME.primary, textMuted: THEME.textMuted, shadow: THEME.shadow }}
              />
              <AppButton kind="primary" onClick={addSite}>Ajouter</AppButton>
            </div>

            {Object.entries(groupedSites).map(function(entry) {
              const gid = entry[0]; const list = entry[1];
              const gname = gid === "ungrouped" ? "Sans groupe" : (groups.find(function(g) { return g.id === gid; }) || {}).name;
              return (
                <div key={gid} style={{ marginBottom: 16 }}>
                  {gid !== "ungrouped" && <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 8, paddingLeft: 8, borderLeft: "3px solid " + THEME.gold }}>{gname}</div>}
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {list.map(function(s) {
                      const freq = freqs.find(function(f) { return f.siteId === s.id; });
                      const isEdit = editingFreq === s.id;
                      return (
                        <div key={s.id} style={{ border: "1px solid " + THEME.cardBorder, borderLeft: "3px solid " + THEME.gold, borderRadius: THEME.radiusSm, padding: "9px 12px", background: THEME.card, boxShadow: THEME.shadow, marginBottom: 5 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
                            <div style={{ flex: 1 }}>
                              <AppInput value={s.name} onChange={function(ev) { updSite(s.id, { name: ev.target.value }); }} style={{ width: "100%", marginBottom: 4 }} />
                              <AddressAutocomplete
                                value={s.address || ""}
                                onChange={function(p) {
                                  var patch = { address: p.address };
                                  if (p.lat != null && p.lon != null) {
                                    patch.lat = p.lat;
                                    patch.lon = p.lon;
                                  } else {
                                    patch.lat = undefined;
                                    patch.lon = undefined;
                                  }
                                  updSite(s.id, patch);
                                }}
                                placeholder="Adresse"
                                style={{ marginBottom: 4 }}
                                theme={{ grid: THEME.grid, pale: THEME.pale, text: THEME.primary, textMuted: THEME.textMuted, shadow: THEME.shadow }}
                              />
                              {groups.length > 0 && (
                                <AppSelect value={s.groupId || ""} onChange={function(e) { updSite(s.id, { groupId: e.target.value || undefined }); }} style={{ width: "100%", fontSize: 12, marginBottom: 4 }}>
                                  <option value="">Sans groupe</option>
                                  {groups.map(function(g) { return <option key={g.id} value={g.id}>{g.name}</option>; })}
                                </AppSelect>
                              )}
                              <AppSelect value={s.defaultServiceId || ""} onChange={function(e) { updSite(s.id, { defaultServiceId: e.target.value || undefined }); }} style={{ width: "100%", fontSize: 12, marginBottom: 6 }}>
                                <option value="">Pas de prestation par defaut</option>
                                {services.map(function(svc) { return <option key={svc.id} value={svc.id}>{svc.name} ({svc.duration} min)</option>; })}
                              </AppSelect>
                              <div style={{ fontSize: 11, fontWeight: 600, color: THEME.textMuted, marginBottom: 4 }}>Contrainte horaire (optionnel)</div>
                              <AppSelect
                                value={(s.timeConstraint && s.timeConstraint.kind) || "none"}
                                onChange={function(e) {
                                  var k = e.target.value;
                                  if (k === "none") updSite(s.id, { timeConstraint: undefined });
                                  else if (k === "fixed") updSite(s.id, { timeConstraint: { kind: "fixed", at: "09:30" } });
                                  else updSite(s.id, { timeConstraint: { kind: "window", start: "08:00", end: "10:00" } });
                                }}
                                style={{ width: "100%", fontSize: 12, marginBottom: 4 }}
                              >
                                <option value="none">Aucune</option>
                                <option value="fixed">Heure fixe</option>
                                <option value="window">Plage autorisee</option>
                              </AppSelect>
                              {s.timeConstraint && s.timeConstraint.kind === "fixed" && (
                                <AppInput
                                  type="time"
                                  value={s.timeConstraint.at || "09:30"}
                                  onChange={function(e) {
                                    updSite(s.id, { timeConstraint: Object.assign({}, s.timeConstraint, { at: e.target.value }) });
                                  }}
                                  style={{ width: 120, marginBottom: 4 }}
                                />
                              )}
                              {s.timeConstraint && s.timeConstraint.kind === "window" && (
                                <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                                  <AppInput
                                    type="time"
                                    value={s.timeConstraint.start || "08:00"}
                                    onChange={function(e) {
                                      updSite(s.id, { timeConstraint: Object.assign({}, s.timeConstraint, { start: e.target.value }) });
                                    }}
                                    style={{ width: 100 }}
                                  />
                                  <span style={{ fontSize: 12, color: "#888" }}>a</span>
                                  <AppInput
                                    type="time"
                                    value={s.timeConstraint.end || "10:00"}
                                    onChange={function(e) {
                                      updSite(s.id, { timeConstraint: Object.assign({}, s.timeConstraint, { end: e.target.value }) });
                                    }}
                                    style={{ width: 100 }}
                                  />
                                </div>
                              )}
                            </div>
                            <div style={{ marginLeft: 8, display: "flex", flexDirection: "column", gap: 4 }}>
                              <AppButton onClick={function() { setEditingFreq(isEdit ? null : s.id); }}>{isEdit ? "Fermer" : "Frequence"}</AppButton>
                              <AppButton onClick={function() { delSite(s.id); }}>Supprimer</AppButton>
                            </div>
                          </div>
                          {isEdit && (
                            <div style={{ borderTop: "1px solid " + THEME.grid, paddingTop: 10, marginTop: 8 }}>
                              <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, color: "#333" }}>Frequence et occurrences</div>
                              <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "flex-end", marginBottom: 10 }}>
                                <div>
                                  <div style={{ fontSize: 11, color: "#666", marginBottom: 4 }}>Mode</div>
                                  <AppSelect
                                    value={
                                      freq && freq.customDateMode
                                        ? "custom_iso"
                                        : (freq && freq.type) || "once"
                                    }
                                    onChange={function(e) {
                                      var v = e.target.value;
                                      var base = { fixedDate: undefined, customDateMode: false };
                                      if (v === "custom_iso") {
                                        updFreq(
                                          s.id,
                                          Object.assign(base, {
                                            type: "monthly",
                                            timesPerMonth: Math.max(1, (freq && freq.timesPerMonth) || 4),
                                            customDates: (freq && freq.customDates) || [],
                                            customDateMode: true
                                          })
                                        );
                                      } else if (v === "fixed") {
                                        updFreq(s.id, Object.assign(base, { type: "fixed", timesPerMonth: 1, fixedDate: todayISO(), customDates: [] }));
                                      } else if (v === "monthly") {
                                        updFreq(s.id, Object.assign(base, { type: "monthly", timesPerMonth: Math.max(1, (freq && freq.timesPerMonth) || 4), customDates: [] }));
                                      } else {
                                        updFreq(s.id, Object.assign(base, {
                                          type: v,
                                          timesPerMonth: v === "daily" ? 26 : v === "weekly" ? 4 : v === "fortnightly" ? 2 : 1,
                                          customDates: []
                                        }));
                                      }
                                    }}
                                    style={{ minWidth: 200 }}
                                  >
                                    <option value="once">Unique (1er jour ouvre)</option>
                                    <option value="weekly">Chaque semaine (4 / mois)</option>
                                    <option value="fortnightly">Toutes les 2 semaines</option>
                                    <option value="monthly">X fois par mois</option>
                                    <option value="daily">Chaque jour ouvre</option>
                                    <option value="fixed">Date fixe</option>
                                    <option value="customWeekdays">Jours precis (lun–sam)</option>
                                    <option value="custom_iso">Dates personnalisees (ISO)</option>
                                  </AppSelect>
                                </div>
                                {freq && freq.type === "monthly" && !freq.customDateMode && (
                                  <div style={{ display: "flex", alignItems: "center", gap: 8, background: THEME.goldPale, border: "1px solid " + THEME.goldBorder, borderRadius: THEME.radiusSm, padding: "8px 14px" }}>
                                    <AppInput
                                      type="number"
                                      min="1"
                                      max="31"
                                      value={(freq && freq.timesPerMonth) || 1}
                                      onChange={function(e) {
                                        var n = Math.max(1, parseInt(e.target.value, 10) || 1);
                                        updFreq(s.id, { type: "monthly", timesPerMonth: n });
                                      }}
                                      style={{ width: 60, textAlign: "center", fontWeight: 700, fontSize: 18, border: "none", background: "transparent", padding: "0 4px" }}
                                    />
                                    <span style={{ fontSize: 13, color: "#555", whiteSpace: "nowrap" }}>fois / mois</span>
                                  </div>
                                )}
                                {freq && freq.type === "fixed" && (
                                  <div>
                                    <div style={{ fontSize: 11, color: "#666", marginBottom: 4 }}>Date</div>
                                    <AppInput
                                      type="date"
                                      value={(freq && freq.fixedDate) || todayISO()}
                                      onChange={function(e) { updFreq(s.id, { fixedDate: e.target.value }); }}
                                    />
                                  </div>
                                )}
                              </div>
                              {freq && freq.type === "customWeekdays" && (
                                <div style={{ marginBottom: 10 }}>
                                  <div style={{ fontSize: 11, color: "#666", marginBottom: 6 }}>Jours autorises</div>
                                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                                    {DAYS_LABELS.map(function(label, wd) {
                                      var arr = (freq && freq.allowedWeekdays) || [];
                                      var on = arr.indexOf(wd) !== -1;
                                      return (
                                        <label key={wd} style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 4, cursor: "pointer" }}>
                                          <input
                                            type="checkbox"
                                            checked={on}
                                            onChange={function() {
                                              var next = on ? arr.filter(function(x) { return x !== wd; }) : arr.concat([wd]);
                                              updFreq(s.id, { allowedWeekdays: next });
                                            }}
                                          />
                                          {label}
                                        </label>
                                      );
                                    })}
                                  </div>
                                </div>
                              )}
                              <div style={{ display: "flex", flexWrap: "wrap", gap: 14, marginBottom: 10 }}>
                                <div>
                                  <div style={{ fontSize: 11, color: "#666", marginBottom: 4 }}>Max. occurrences (vide = illimite)</div>
                                  <AppInput
                                    type="number"
                                    min="1"
                                    max="99"
                                    placeholder="—"
                                    value={(freq && freq.occurrenceCount != null && freq.occurrenceCount > 0) ? String(freq.occurrenceCount) : ""}
                                    onChange={function(e) {
                                      var raw = e.target.value.trim();
                                      if (!raw) updFreq(s.id, { occurrenceCount: undefined });
                                      else updFreq(s.id, { occurrenceCount: Math.max(1, parseInt(raw, 10) || 1) });
                                    }}
                                    style={{ width: 80 }}
                                  />
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, color: "#666", marginBottom: 4 }}>Jours interdits (lun–dim)</div>
                                  <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                                    {DAYS_LABELS.concat(["Dim"]).map(function(label, idx) {
                                      var wd = idx === 6 ? 6 : idx;
                                      var arr = (freq && freq.blockedWeekdays) || [];
                                      var on = arr.indexOf(wd) !== -1;
                                      return (
                                        <label key={"b" + wd} style={{ fontSize: 11, display: "flex", alignItems: "center", gap: 2, cursor: "pointer" }}>
                                          <input
                                            type="checkbox"
                                            checked={on}
                                            onChange={function() {
                                              var next = on ? arr.filter(function(x) { return x !== wd; }) : arr.concat([wd]);
                                              updFreq(s.id, { blockedWeekdays: next });
                                            }}
                                          />
                                          {label}
                                        </label>
                                      );
                                    })}
                                  </div>
                                </div>
                              </div>
                              <div style={{ marginBottom: 8 }}>
                                <div style={{ fontSize: 11, color: "#666", marginBottom: 4 }}>Dates interdites (YYYY-MM-DD, une par ligne)</div>
                                <textarea
                                  rows={2}
                                  style={{ width: "100%", fontSize: 12, border: "1px solid " + THEME.grid, borderRadius: 8, padding: 8 }}
                                  placeholder={"2026-04-01\n2026-04-15"}
                                  defaultValue={(freq && freq.blockedDates && freq.blockedDates.join("\n")) || ""}
                                  onBlur={function(e) { updFreq(s.id, { blockedDates: splitIsoDateTokens(e.target.value) }); }}
                                />
                              </div>
                              {freq && freq.customDateMode && (
                                <div style={{ marginBottom: 8 }}>
                                  <div style={{ fontSize: 11, color: "#666", marginBottom: 4 }}>Dates prevues (ISO, une par ligne) — prioritaire sur le mode</div>
                                  <textarea
                                    key={(freq.customDates || []).join(",")}
                                    rows={3}
                                    style={{ width: "100%", fontSize: 12, border: "1px solid " + THEME.grid, borderRadius: 8, padding: 8 }}
                                    placeholder={"2026-04-05\n2026-04-12"}
                                    defaultValue={(freq && freq.customDates && freq.customDates.join("\n")) || ""}
                                    onBlur={function(e) { updFreq(s.id, { customDates: splitIsoDateTokens(e.target.value) }); }}
                                  />
                                </div>
                              )}
                              {freq && freq.type === "monthly" && freq.timesPerMonth > 0 && !freq.customDateMode && (
                                <div style={{ fontSize: 11, color: "#a87a3d", fontStyle: "italic" }}>
                                  soit ~1 passage tous les {Math.round(22 / (freq.timesPerMonth || 1))} jours ouvres (estimation)
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
            {!filteredSites.length && <div style={{ fontSize: 14, color: "#777" }}>Aucun chantier.</div>}
          </AppBox>
        )}

        {tab === "services" && (
          <AppBox title="Prestations">
            <AppInput placeholder="Rechercher..." value={searchSvcs} onChange={function(e) { setSearchSvcs(e.target.value); }} style={{ marginBottom: 12, width: "100%" }} />
            <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
              <AppInput placeholder="Intitule" value={newSvcName} onChange={function(e) { setNewSvcName(e.target.value); }} />
              <AppInput placeholder="Duree (min)" value={newSvcDur} onChange={function(e) { setNewSvcDur(e.target.value); }} style={{ width: 120 }} />
              <AppButton kind="primary" onClick={addSvc}>Ajouter</AppButton>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {filteredSvcs.map(function(s) {
                return (
                  <div key={s.id} style={{ display: "flex", gap: 8 }}>
                    <AppInput value={s.name} onChange={function(ev) { setServices(function(p) { return p.map(function(x) { return x.id === s.id ? Object.assign({}, x, { name: ev.target.value }) : x; }); }); }} style={{ flex: 1 }} />
                    <AppInput value={String(s.duration)} onChange={function(ev) { setServices(function(p) { return p.map(function(x) { return x.id === s.id ? Object.assign({}, x, { duration: parseInt(ev.target.value, 10) || 0 }) : x; }); }); }} style={{ width: 100 }} />
                    <AppButton onClick={function() { delSvc(s.id); }}>Supprimer</AppButton>
                  </div>
                );
              })}
              {!filteredSvcs.length && <div style={{ fontSize: 14, color: "#777" }}>Aucune prestation.</div>}
            </div>
          </AppBox>
        )}
      </div>

      <div style={{ textAlign: "center", fontSize: 12, color: "#888", padding: 12, position: "relative", zIndex: 1 }}>
        Donnees sauvegardees localement · Trajets OSRM · Proximite Nearest Neighbour · Import Excel/CSV · Export PDF / ICS / Word
      </div>
    </div>
  );
}
