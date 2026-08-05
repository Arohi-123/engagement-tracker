// The 5-role model: SuperAdmin, GlobalUser, GlobalViewer, RegionalUser, RegionalViewer.
const GLOBAL_ROLES = new Set(['SuperAdmin', 'GlobalUser', 'GlobalViewer']);
const WRITE_ROLES = new Set(['SuperAdmin', 'GlobalUser', 'RegionalUser']);

// Regional roles can hold more than one region now — user.regions is always an array.
function hasRegion(user, regionKey) {
  return Array.isArray(user.regions) && user.regions.some(r => String(r).toUpperCase() === String(regionKey).toUpperCase());
}

function canReadRegion(user, regionKey) {
  if (!user || !user.role) return false;
  if (GLOBAL_ROLES.has(user.role)) return true;
  return hasRegion(user, regionKey);
}

function canWriteRegion(user, regionKey) {
  if (!user || !user.role || !WRITE_ROLES.has(user.role)) return false;
  if (user.role === 'SuperAdmin' || user.role === 'GlobalUser') return true;
  return hasRegion(user, regionKey); // RegionalUser — only their assigned region(s)
}

module.exports = { canReadRegion, canWriteRegion };
