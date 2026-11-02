// Short id handles.
//
// Graph ids are ~85 opaque chars, paid for in every list result and every
// follow-up argument. Each maps to a short handle derived from it (kind prefix
// + base36 hash) and persisted, so it resolves across restarts. Handles minted
// through the collision ladder are the exception to determinism: they depend
// on what else the minting process had seen.
import fs from 'fs';

/** FNV-1a, rendered base36 (4-7 chars for a 32-bit hash). */
export function fnv1a36(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(36).padStart(4, '0');
}

// Handles are lowercase alphanumerics; real Graph ids always contain '!', '-'
// or uppercase, so this never matches one.
const HANDLE_RE = /^(nb|sg|s|p)[0-9a-z]{4,}$/;
export const isHandle = (s) => typeof s === 'string' && HANDLE_RE.test(s);

/** Handle store. Omit filePath for an in-memory one (tests). */
export function createHandleStore({ filePath } = {}) {
  const readDisk = () => {
    if (!filePath) return {};
    try {
      if (fs.existsSync(filePath)) return JSON.parse(fs.readFileSync(filePath, 'utf8'));
    } catch { /* a corrupt or unreadable map is not fatal — handles just re-mint */ }
    return {};
  };

  let map = readDisk();
  let flushTimer = null;

  function flush() {
    if (!filePath) return;
    try {
      // Merge with disk so concurrent clients converge rather than clobber.
      map = { ...readDisk(), ...map };
      fs.writeFileSync(filePath, JSON.stringify(map));
    } catch { /* best-effort persistence; the in-memory map still works */ }
  }

  function scheduleFlush() {
    if (flushTimer) return;
    flushTimer = setTimeout(() => { flushTimer = null; flush(); }, 250);
    if (flushTimer.unref) flushTimer.unref(); // do not hold the process open
  }

  /** Real id -> handle, minting and persisting on first sight. */
  function handle(realId, kind) {
    if (!realId || typeof realId !== 'string') return realId;
    const base = fnv1a36(realId);
    // Widen one char at a time; base is 4-7 chars, so slicing past its length
    // would just retry the same candidate.
    for (let len = 4; len <= base.length; len++) {
      const h = kind + base.slice(0, len);
      if (!map[h]) { map[h] = realId; scheduleFlush(); return h; }
      if (map[h] === realId) return h;
    }
    for (let n = 2; ; n++) {
      const h = kind + base + n;
      if (!map[h]) { map[h] = realId; scheduleFlush(); return h; }
      if (map[h] === realId) return h;
    }
  }

  /** Handle (or raw Graph id) -> real id. Raw ids pass through. */
  function resolveId(handleOrId) {
    if (!isHandle(handleOrId)) return handleOrId;
    if (map[handleOrId]) return map[handleOrId];
    // Another client may have minted this handle after we loaded the map.
    map = { ...readDisk(), ...map };
    if (map[handleOrId]) return map[handleOrId];
    throw new Error(`Unknown id handle "${handleOrId}". Re-run a list/find/tree call to refresh handles.`);
  }

  return { handle, resolveId, flush };
}
