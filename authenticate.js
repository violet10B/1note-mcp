// Device-code login for the OneNote MCP server.
//
// The MSAL client, client id, scopes and cache plugin all come from
// src/auth.js — the same module the server uses — so `npm run auth` and the
// `authenticate` tool always agree on where the token cache lives.
import dotenv from 'dotenv';
import path from 'path';
import { msalCachePath, createClient, SCOPES, ROOT } from './src/auth.js';

// Resolved against the repo root, not the cwd, so `npm run auth` works from
// anywhere in the tree.
dotenv.config({ path: path.join(ROOT, '.env') });

const pca = createClient();

async function authenticate() {
  try {
    // Check for an existing valid session first.
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      try {
        const r = await pca.acquireTokenSilent({ account: accounts[0], scopes: SCOPES });
        console.log(`\nAlready authenticated as ${accounts[0].username}.`);
        console.log(`Token valid until: ${new Date(r.expiresOn).toLocaleString()}`);
        console.log('No re-login needed — token will auto-refresh.');
        return;
      } catch { /* expired or revoked — fall through to device code */ }
    }

    console.log('Starting authentication...');
    console.log('A URL and code will appear shortly...\n');

    const result = await pca.acquireTokenByDeviceCode({
      deviceCodeCallback: (info) => { console.log(info.message); },
      scopes: SCOPES,
    });

    console.log(`\nAuthentication successful!`);
    console.log(`Signed in as: ${result.account.username}`);
    console.log(`Token cache saved to: ${msalCachePath}`);
    console.log('Token will auto-refresh silently — no hourly re-login needed.');
  } catch (error) {
    console.error('Authentication error:', error.message);
    process.exit(1);
  }
}

authenticate();
