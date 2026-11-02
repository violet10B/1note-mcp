// Graph HTTP client: request building, error surfacing, retry policy.
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * The actionable part of a Graph error. The wire format wraps it in an envelope
 * with a request id, a timestamp and a nested innerError, none of which a caller
 * can use. Bodies that are not that envelope are passed through as-is.
 */
export function graphMessage(detail) {
  try {
    const err = JSON.parse(detail).error;
    if (err && err.message) return (err.code ? `${err.code}: ` : '') + err.message;
  } catch { /* not the JSON envelope — the raw body is the best we have */ }
  return detail;
}

/**
 * Build a `graph(method, path, opts)` function.
 *
 * `getToken` is injected rather than imported so this module holds no auth
 * state, and `fetchImpl` so a test can stub the transport without touching
 * globals.
 */
export function createGraph({ getToken, baseUrl = GRAPH_BASE, fetchImpl = globalThis.fetch } = {}) {
  if (typeof getToken !== 'function') throw new Error('createGraph needs a getToken function.');

  async function graph(method, apiPath, opts = {}) {
    const token = await getToken();
    const url = apiPath.startsWith('http') ? apiPath : baseUrl + apiPath;
    const headers = { Authorization: `Bearer ${token}` };
    let body;
    if (opts.formData !== undefined) {
      // Multipart — let fetch/undici set Content-Type with the boundary.
      body = opts.formData;
    } else if (opts.body !== undefined) {
      if (opts.json) {
        headers['Content-Type'] = opts.contentType || 'application/json';
        body = JSON.stringify(opts.body);
      } else {
        if (opts.contentType) headers['Content-Type'] = opts.contentType;
        body = opts.body;
      }
    }
    const res = await fetchImpl(url, { method, headers, body });
    if (!res.ok) {
      let detail = '';
      try { detail = await res.text(); } catch { /* the status alone still tells us enough */ }
      // OneNote throttles edits to one page/section aggressively: 429, and 409
      // with code 30103 ("too many simultaneous requests to the same location").
      // Both are transient, so back off and retry rather than surfacing a failure
      // the caller can only fix by trying again themselves.
      const throttled = res.status === 429 || (res.status === 409 && detail.includes('30103'));
      // A page's content endpoint 404s for a moment after the page is created:
      // create-then-read is a race in OneNote. Retrying a GET absorbs it; a page
      // that is genuinely gone still fails, just three short attempts later.
      const notReadyYet = res.status === 404 && method === 'GET' && /\/content(\?|$)/.test(apiPath);
      const attempt = (opts._attempt || 0) + 1;
      // Multipart bodies are streams and cannot be replayed, so those go straight
      // to the caller rather than being retried with a consumed body.
      if ((throttled || notReadyYet) && attempt <= 3 && opts.formData === undefined) {
        const retryAfter = Number(res.headers.get('Retry-After')) || 0;
        const waitMs = retryAfter ? retryAfter * 1000 : 1000 * 2 ** attempt;
        await new Promise((r) => setTimeout(r, waitMs));
        return graph(method, apiPath, { ...opts, _attempt: attempt });
      }
      const err = new Error(
        `Graph ${method} ${apiPath} -> ${res.status} ${res.statusText} ${graphMessage(detail).slice(0, 300)}`
      );
      // Structured, so a caller can branch on the failure instead of matching
      // the message text — which is what a 501 fallback needs to do safely.
      err.status = res.status;
      throw err;
    }
    if (opts.raw) return res.text();
    if (res.status === 202) {
      // Not a payload — a shape this module made up. Tag it as such rather than
      // hanging a flag off whatever Graph happened to return, which a real
      // response could in principle also carry.
      return {
        kind: 'accepted',
        operationLocation: res.headers.get('Operation-Location') || res.headers.get('Location'),
      };
    }
    if (res.status === 204) return null;
    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }

  /** Follow @odata.nextLink, bounded so a bad paging value cannot loop forever. */
  async function graphList(apiPath) {
    let out = [], next = apiPath, guard = 0;
    while (next && guard++ < 50) {
      const r = await graph('GET', next);
      if (r && Array.isArray(r.value)) out = out.concat(r.value);
      next = r && r['@odata.nextLink'] ? r['@odata.nextLink'] : null;
    }
    return out;
  }

  return { graph, graphList };
}
