// Shared MSAL configuration, used by both the server and authenticate.js so the
// two cannot drift apart. No import-time side effects: .env is the entry
// point's job, and the client id resolves when a client is built.
import { PublicClientApplication } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Repo root — state files live beside package.json, not inside src/.
export const ROOT = path.join(__dirname, '..');

export const tokenFilePath = path.join(ROOT, '.access-token.txt');
export const msalCachePath = path.join(ROOT, '.msal-cache.json');
export const idMapPath = path.join(ROOT, '.id-map.json');

// Graph Explorer's public client id. ONENOTE_CLIENT_ID overrides it, so a fork
// keeps working if Microsoft retires this one.
export const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';

/** The client id to use: the environment's, or Microsoft's public one. */
export const clientId = () => process.env.ONENOTE_CLIENT_ID || DEFAULT_CLIENT_ID;

// Personal MSA only supports the non-".All" OneNote scopes.
export const SCOPES = [
  'https://graph.microsoft.com/Notes.ReadWrite',
  'https://graph.microsoft.com/Notes.Create',
  'https://graph.microsoft.com/User.Read',
];

/**
 * Write a file that holds a credential.
 *
 * The MSAL cache carries a refresh token valid for ~90 days, so it must not be
 * readable by every account on the machine — which is exactly what the default
 * 0644 gives you from writeFileSync. `mode` only applies when the file is
 * created, so an existing file is chmod'ed as well. chmod is a no-op on
 * Windows, hence the guard.
 */
export function writeSecret(filePath, contents) {
  fs.writeFileSync(filePath, contents, { mode: 0o600 });
  try {
    fs.chmodSync(filePath, 0o600);
  } catch { /* platforms without POSIX modes */ }
}

/**
 * MSAL cache plugin persistence — what makes silent auto-refresh work across
 * restarts (~90 days without re-login).
 */
export function createCachePlugin(cachePath = msalCachePath) {
  return {
    beforeCacheAccess: async (ctx) => {
      try {
        if (fs.existsSync(cachePath)) {
          ctx.tokenCache.deserialize(fs.readFileSync(cachePath, 'utf8'));
        }
      } catch { /* a missing or corrupt cache just means re-authenticating */ }
    },
    afterCacheAccess: async (ctx) => {
      if (ctx.cacheHasChanged) writeSecret(cachePath, ctx.tokenCache.serialize());
    },
  };
}

/** A PublicClientApplication wired to the persisted cache. */
export function createClient(cachePath = msalCachePath) {
  return new PublicClientApplication({
    auth: { clientId: clientId(), authority: 'https://login.microsoftonline.com/common' },
    cache: { cachePlugin: createCachePlugin(cachePath) },
  });
}
