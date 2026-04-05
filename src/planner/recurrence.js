/**
 * Fréquences / occurrences : parsing libre + règles avancées (jours, plage, perso).
 */

function weekdayIdxFromIso(iso) {
  return (new Date(iso + "T12:00:00").getDay() + 6) % 7;
}

export function parseFreq(str) {
  if (!str) return { type: "once", timesPerMonth: 1 };
  var s = String(str).toLowerCase().trim();
  if (s.indexOf("semaine") !== -1 || s === "hebdomadaire" || s === "weekly")
    return { type: "weekly", timesPerMonth: 4 };
  if (
    s.indexOf("15 jours") !== -1 ||
    s.indexOf("quinzaine") !== -1 ||
    s.indexOf("2 semaines") !== -1 ||
    s.indexOf("bimensuel") !== -1
  )
    return { type: "fortnightly", timesPerMonth: 2 };
  if (s.indexOf("2 fois") !== -1 || s.indexOf("deux fois") !== -1)
    return { type: "fortnightly", timesPerMonth: 2 };
  if (
    s.indexOf("1 fois") !== -1 ||
    s.indexOf("une fois") !== -1 ||
    s === "mensuel" ||
    s === "monthly"
  )
    return { type: "monthly", timesPerMonth: 1 };
  var match = s.match(/^(\d+)/);
  if (match) {
    var n2 = parseInt(match[1], 10);
    if (n2 >= 1) return { type: "monthly", timesPerMonth: n2 };
  }
  if (s === "quotidien" || s === "daily") return { type: "daily", timesPerMonth: 26 };
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return { type: "fixed", timesPerMonth: 1, fixedDate: s };
  return { type: "once", timesPerMonth: 1 };
}

export function generateWeekdayOnlyIndices(workdays, allowedWeekdays) {
  if (!workdays.length) return [];
  if (!allowedWeekdays || !allowedWeekdays.length) return [0];
  var out = [];
  for (var i = 0; i < workdays.length; i++) {
    var wd = weekdayIdxFromIso(workdays[i]);
    if (allowedWeekdays.indexOf(wd) !== -1) out.push(i);
  }
  return out.length ? out : [0];
}

export function generateTargetDays(freqType, n, workdays) {
  var nb = workdays.length;
  if (nb === 0) return [];

  if (freqType === "once") {
    return [0];
  }

  if (freqType === "daily") {
    var all = [];
    for (var i = 0; i < nb; i++) all.push(i);
    return all;
  }

  if (freqType === "weekly") {
    var seenWeeks = {};
    var targets = [];
    for (var i2 = 0; i2 < nb; i2++) {
      var d = workdays[i2];
      var dateObj = new Date(d + "T12:00:00");
      var dayOfMonth = dateObj.getDate();
      var weekNum = Math.floor((dayOfMonth - 1) / 7);
      if (!seenWeeks[weekNum]) {
        seenWeeks[weekNum] = true;
        targets.push(i2);
      }
      if (targets.length >= 4) break;
    }
    while (targets.length < 4 && targets.length < nb) {
      var spacing = Math.floor(nb / 4);
      var pos = targets.length * spacing;
      pos = Math.min(pos, nb - 1);
      if (targets.indexOf(pos) === -1) targets.push(pos);
      else targets.push(Math.min(pos + 1, nb - 1));
    }
    return targets.slice(0, 4);
  }

  if (freqType === "fortnightly") {
    var t1 = -1;
    var t2 = -1;
    for (var i3 = 0; i3 < nb; i3++) {
      var day = new Date(workdays[i3] + "T12:00:00").getDate();
      if (t1 === -1 && day >= 5) {
        t1 = i3;
      }
      if (t1 !== -1 && t2 === -1 && day >= 19) {
        t2 = i3;
        break;
      }
    }
    if (t1 === -1) t1 = 0;
    if (t2 === -1 || t2 === t1) {
      var d1 = new Date(workdays[t1] + "T12:00:00");
      for (var i4 = t1 + 1; i4 < nb; i4++) {
        var d2 = new Date(workdays[i4] + "T12:00:00");
        var gap = (d2 - d1) / (1000 * 60 * 60 * 24);
        if (gap >= 13) {
          t2 = i4;
          break;
        }
      }
      if (t2 === -1 || t2 === t1) t2 = nb - 1;
    }
    return [t1, t2];
  }

  if (freqType === "monthly") {
    if (n >= nb) {
      var allIdx = [];
      for (var i5 = 0; i5 < nb; i5++) allIdx.push(i5);
      return allIdx;
    }
    var usedIdx = {};
    var result = [];
    for (var k = 0; k < n; k++) {
      var spacing2 = nb / n;
      var idealPos = Math.round(k * spacing2 + spacing2 / 2) - 1;
      idealPos = Math.max(0, Math.min(nb - 1, idealPos));
      var pos2 = idealPos;
      var found = false;
      for (var r = 0; r < nb && !found; r++) {
        var candidates = r === 0 ? [idealPos] : [idealPos + r, idealPos - r];
        for (var ci = 0; ci < candidates.length && !found; ci++) {
          var p = candidates[ci];
          if (p >= 0 && p < nb && !usedIdx[p]) {
            pos2 = p;
            found = true;
          }
        }
      }
      usedIdx[pos2] = true;
      result.push(pos2);
    }
    result.sort(function (a, b) {
      return a - b;
    });
    return result;
  }

  return [0];
}

/**
 * Filtre les indices de jours ouvrés selon jours autorisés / interdits / dates bloquées.
 */
export function applyRecurrenceFilters(workdays, indices, fp) {
  if (!indices || !indices.length) return [];
  var allowed = fp.allowedWeekdays;
  var hasAllowed = allowed && allowed.length > 0;
  var blockedW = fp.blockedWeekdays || [];
  var blockedSet = {};
  (fp.blockedDates || []).forEach(function (d) {
    if (d) blockedSet[String(d)] = true;
  });

  var out = indices.filter(function (ii) {
    var iso = workdays[ii];
    if (!iso) return false;
    var wd = weekdayIdxFromIso(iso);
    if (hasAllowed && allowed.indexOf(wd) === -1) return false;
    if (blockedW.indexOf(wd) !== -1) return false;
    if (blockedSet[iso]) return false;
    return true;
  });

  return out;
}

/**
 * Indices de jours ouvrés (lun–sam) pour un chantier sur le mois courant.
 */
export function computeOccurrenceDayIndices(workdays, fp) {
  var ftype = fp.type || "once";
  var n = Math.max(1, fp.timesPerMonth || 1);
  var indices = [];

  if (fp.customDateMode || (fp.customDates && fp.customDates.length > 0)) {
    if (!fp.customDates || !fp.customDates.length) return [];
    fp.customDates.forEach(function (iso) {
      var i = workdays.indexOf(iso);
      if (i >= 0) indices.push(i);
    });
    indices.sort(function (a, b) {
      return a - b;
    });
    var seen = {};
    indices = indices.filter(function (x) {
      if (seen[x]) return false;
      seen[x] = true;
      return true;
    });
  } else if (ftype === "customWeekdays") {
    indices = generateWeekdayOnlyIndices(workdays, fp.allowedWeekdays || []);
  } else if (ftype === "fixed" && fp.fixedDate) {
    var fi = workdays.indexOf(fp.fixedDate);
    indices = [fi >= 0 ? fi : 0];
  } else {
    indices = generateTargetDays(ftype, n, workdays);
  }

  indices = applyRecurrenceFilters(workdays, indices, fp);

  if (fp.occurrenceCount != null && fp.occurrenceCount > 0 && indices.length > fp.occurrenceCount) {
    indices = indices.slice(0, fp.occurrenceCount);
  }

  return indices;
}
