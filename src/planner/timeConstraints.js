/**
 * Contraintes horaires optionnelles sur un chantier (structure évolutive).
 * kind: 'none' | 'fixed' | 'window'
 */

function parseHM(str) {
  if (!str) return 0;
  var parts = String(str).split(":");
  var h = parseInt(parts[0], 10) || 0;
  var m = parseInt(parts[1], 10) || 0;
  return h * 60 + m;
}

/**
 * @param {number} startMin — heure de début souhaitée (minutes depuis minuit)
 * @param {number} durMin — durée prestation
 * @param {number} earliestMin — minimum absolu (après trajet depuis le précédent)
 * @param {object|null} timeConstraint — { kind, at?, start?, end? }
 */
export function adjustStartForTimeConstraint(startMin, durMin, earliestMin, timeConstraint) {
  if (!timeConstraint || !timeConstraint.kind || timeConstraint.kind === "none") {
    return Math.max(startMin, earliestMin);
  }
  if (timeConstraint.kind === "fixed" && timeConstraint.at) {
    var ft = parseHM(timeConstraint.at);
    return Math.max(earliestMin, ft);
  }
  if (timeConstraint.kind === "window" && timeConstraint.start && timeConstraint.end) {
    var ws = parseHM(timeConstraint.start);
    var we = parseHM(timeConstraint.end);
    var latestStart = we - durMin;
    if (latestStart < ws) return Math.max(startMin, earliestMin);
    var s = Math.max(earliestMin, ws);
    s = Math.max(s, startMin);
    return Math.min(s, latestStart);
  }
  return Math.max(startMin, earliestMin);
}

export { parseHM as parseHMMinutes };
