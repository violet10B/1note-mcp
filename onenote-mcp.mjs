#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { InteractionRequiredAuthError } from '@azure/msal-node';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';

import {
  createClient, writeSecret, tokenFilePath, idMapPath, SCOPES, ROOT,
} from './src/auth.js';
import { createGraph } from './src/graph.js';
import { assertFetchableUrl } from './src/images.js';
import { createHandleStore } from './src/ids.js';
import { checkMathInContent } from './src/mathml.js';
import { renderGuide, TOPICS as GUIDE_TOPICS } from './src/guide.js';
import {
  parseHtml, contentContainer, sectionSpan, spliceSection, textNodesUnder, replaceInTextNodes,
  pickOne, checkboxesIn,
} from './src/dom.js';
import {
  htmlToMarkdown, htmlToText, outlineOf, toOneNoteHtml, renderTable, tableToGrid,
  editGrid, assertNoLostImages, escapeHtml, toXhtml,
} from './src/content.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Resolved against the repo root, not the cwd: an MCP client launches the
// server from whatever directory it likes.
dotenv.config({ path: path.join(ROOT, '.env') });

const VERSION = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8')).version;

const GRAPH = 'https://graph.microsoft.com/v1.0';

// Client id, scopes and the persisted token cache live in src/auth.js, shared
// with authenticate.js so the two entry points cannot drift apart.
const pca = createClient();

async function getToken() {
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const r = await pca.acquireTokenSilent({ account: accounts[0], scopes: SCOPES });
      return r.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        throw new Error('Session expired. Run the "authenticate" tool or `npm run auth`.');
      }
      console.error('Silent token refresh failed:', e.message);
    }
  }
  try {
    if (fs.existsSync(tokenFilePath)) {
      const t = JSON.parse(fs.readFileSync(tokenFilePath, 'utf8')).token;
      if (t) return t;
    }
  } catch { /* fall through to the env var / error below */ }
  if (process.env.GRAPH_ACCESS_TOKEN) return process.env.GRAPH_ACCESS_TOKEN;
  throw new Error('Not authenticated. Run the "authenticate" tool or `npm run auth`.');
}

// ---------------------------------------------------------------------------
// Microsoft Graph helpers
//
// The HTTP client lives in src/graph.js: request building, error surfacing and
// the retry policy are the parts a caller depends on and that are worth testing
// on their own, and they need no OneNote knowledge to do their job.
// ---------------------------------------------------------------------------
const { graph, graphList } = createGraph({ getToken, baseUrl: GRAPH });

/**
 * Run `fn` over `items` a few at a time — enough parallelism to be quick, not
 * enough for OneNote to start returning 429s. `fn` gets the item's index in the
 * whole list, not in its batch.
 */
async function inBatches(items, size, fn) {
  const out = [];
  for (let i = 0; i < items.length; i += size) {
    out.push(...await Promise.all(items.slice(i, i + size).map((item, j) => fn(item, i + j))));
  }
  return out;
}

async function awaitOperation(operationUrl) {
  if (!operationUrl) return { status: 'initiated' };
  // Poll immediately — a small copy is often already done, and sleeping first
  // charged every caller 3 s it did not need.
  for (let i = 0; i < 20; i++) {
    const op = await graph('GET', operationUrl);
    if (op && op.status === 'completed') return op;
    if (op && op.status === 'failed') {
      throw new Error(`Operation failed: ${op.error ? op.error.message : JSON.stringify(op)}`);
    }
    await new Promise((r) => setTimeout(r, 3000));
  }
  throw new Error('Operation timed out after 60 s.');
}

// Short-lived read caches. Writes invalidate the entry, so a stale read is
// never used for an edit.
const PAGE_TTL_MS = 60_000;
const LIST_TTL_MS = 60_000;
const pageCache = new Map();
const listCache = new Map();

/** Drop every cached variant of a page. */
const invalidatePage = (pageId) => {
  pageCache.delete(`${pageId}|true`);
  pageCache.delete(`${pageId}|false`);
};

/** Page content HTML, cached. includeIds is part of the key. */
async function getPageHtml(pageId, { includeIds = true, fresh = false } = {}) {
  const key = `${pageId}|${includeIds}`;
  const hit = pageCache.get(key);
  if (!fresh && hit && Date.now() - hit.at < PAGE_TTL_MS) return hit.html;
  const html = await graph(
    'GET', `/me/onenote/pages/${pageId}/content${includeIds ? '?includeIDs=true' : ''}`, { raw: true }
  );
  pageCache.set(key, { html, at: Date.now() });
  return html;
}

/** Page content as a parsed DOM (ids present). */
async function fetchPageDom(pageId, opts) {
  const html = await getPageHtml(pageId, opts);
  const doc = await parseHtml(html);
  return { html, doc, body: doc.body };
}

async function cachedList(key, fetcher) {
  const hit = listCache.get(key);
  if (hit && Date.now() - hit.at < LIST_TTL_MS) return hit.value;
  const value = await fetcher();
  listCache.set(key, { value, at: Date.now() });
  return value;
}

const SECTIONS_QUERY =
  '/me/onenote/sections?$select=id,displayName&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)&$top=100';
const allSections = () => cachedList('sections', () => graphList(SECTIONS_QUERY));
const sectionPages = (sectionId) => cachedList(
  `pages:${sectionId}`, () => graphList(`/me/onenote/sections/${sectionId}/pages?$select=id,title&$top=100`)
);
const invalidatePageLists = () => {
  for (const key of [...listCache.keys()]) if (key.startsWith('pages:')) listCache.delete(key);
};
const allNotebooks = () =>
  cachedList('notebooks', () => graphList('/me/onenote/notebooks?$select=id,displayName,isDefault'));

// Tool result helpers
const ok = (data) => ({
  content: [{ type: 'text', text: typeof data === 'string' ? data : JSON.stringify(data) }],
});
const fail = (err) => ({
  content: [{ type: 'text', text: 'Error: ' + (err && err.message ? err.message : String(err)) }],
  isError: true,
});

const ids = createHandleStore({ filePath: idMapPath });
const { handle, resolveId } = ids;

/** Convert declared-format content to OneNote XHTML, refusing invalid MathML. */
async function prepareContent(content, format, what = 'page') {
  const { html, warnings } = await toOneNoteHtml(content, format);
  const mc = await checkMathInContent(html);
  if (mc.hardError) throw new Error(`Invalid MathML, ${what} not written: ${mc.hardError}`);
  return { html, warnings: [...warnings, ...mc.warnings] };
}

const withWarnings = (result, warnings) =>
  warnings && warnings.length ? { ...(typeof result === 'string' ? { status: result } : result), warnings } : result;

/**
 * Resolve an element id from a text fragment, preferring the most specific
 * match (a paragraph over the div wrapping it). OneNote's includeIDs adds
 * id="p:{guid}{n}", not data-id.
 */
async function findElementByText(pageId, text, { fresh = false } = {}) {
  const { doc } = await fetchPageDom(pageId, { fresh });
  const needle = text.toLowerCase();
  const matches = [...doc.body.querySelectorAll('[id]')]
    .filter((el) => el.textContent.toLowerCase().includes(needle))
    .sort((a, b) => a.textContent.length - b.textContent.length);
  if (!matches.length) throw new Error(`No element found containing "${text}".`);
  return { el: matches[0], elementId: matches[0].getAttribute('id'), doc };
}

/** Render page HTML as the requested format. */
const asFormat = async (html, fmt) =>
  fmt === 'html' ? html
    : fmt === 'text' ? await htmlToText(html)
      : fmt === 'outline' ? await outlineOf(html)
        : await htmlToMarkdown(html);

/**
 * PATCH one page, invalidating its cache. Commands built from a parsed page
 * serialise through jsdom, which writes HTML5 void tags; OneNote wants XHTML.
 */
async function patchPage(pageId, command) {
  const body = command.map((c) => (c.content ? { ...c, content: toXhtml(c.content) } : c));
  await graph('PATCH', `/me/onenote/pages/${pageId}/content`, { body, json: true });
  invalidatePage(pageId);
}

// OneNote wants raw binary in a multipart part, never base64, so the caller
// passes a path or URL and the server reads the bytes.
const IMAGE_TYPES = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff',
};
const MAX_IMAGE_BYTES = Math.floor(3.5 * 1024 * 1024); // stay under Graph's 4 MB request cap

async function loadImageBytes(filePath, url) {
  if (filePath) {
    if (!fs.existsSync(filePath)) throw new Error(`File not found: ${filePath}`);
    const ext = path.extname(filePath).slice(1).toLowerCase();
    const type = IMAGE_TYPES[ext];
    if (!type) throw new Error(`Unsupported image extension ".${ext}". Use: ${Object.keys(IMAGE_TYPES).join(', ')}.`);
    // Size from the directory entry, before the file is read.
    const { size } = fs.statSync(filePath);
    if (size > MAX_IMAGE_BYTES) {
      throw new Error(`Image is ${(size / 1048576).toFixed(1)} MB; the Graph request limit is ~4 MB.`);
    }
    return { buf: fs.readFileSync(filePath), type };
  }
  const safeUrl = assertFetchableUrl(url);
  const res = await fetch(safeUrl);
  if (!res.ok) throw new Error(`Could not fetch image URL: ${res.status} ${res.statusText}`);
  const ct = (res.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
  let type = Object.values(IMAGE_TYPES).includes(ct) ? ct : null;
  if (!type) type = IMAGE_TYPES[(url.split('?')[0].split('.').pop() || '').toLowerCase()];
  if (!type) throw new Error(`Could not determine a supported image type from the URL (content-type: ${ct || 'unknown'}).`);
  return { buf: Buffer.from(await res.arrayBuffer()), type };
}

// Sent once at init rather than per tool call, so it stays short: this is paid
// for on every request. Anything the server can do FOR the model belongs in
// code, not in prose the model has to apply by hand.
const SERVER_INSTRUCTIONS = `IDS — list/tree/search tools return short id handles (e.g. p1a2b, s3f9c, nb07xy) in the "id" field. Pass them straight back as-is; raw Microsoft Graph ids also work. Never invent one.

CONTENT — read and write pages as MARKDOWN; it is the default everywhere. The server converts it and applies the house style, so do NOT hand-write styled HTML. GFM tables, nested lists, emphasis, links, code fences and task lists (- [ ] / - [x] become real check boxes) all work, and raw HTML or MathML passes through inline. A write that returns "warnings" did not fully survive the conversion — read them.

EDITING — prefer the narrowest tool; each avoids reading or re-sending the page:
  editTable       set cells, add/delete rows and columns; by index or by header/row text
  editPageSection replace, append to, prepend to or delete one heading's content
  setCheckbox     tick or untick check boxes by their text
  findAndReplace  swap an exact HTML fragment, or a run of plain text
  insertNearText  insert before/after the paragraph matching a text fragment
  replaceElement  replace one block element by its id
  addToPage       append or prepend Markdown
  replacePage     rewrite the whole body — last resort

READING — getPage(format:"outline") maps the headings, table headers and row labels for a fraction of the page's tokens, and is enough to aim an edit. getPageSection reads one heading; getMathExpressions reads only the formulas; searchPages(inContent:true) finds a page by what is in it, within a section or notebook; batchUpdatePages changes many pages in one call.

IMAGES — addImageToPage takes a local filePath or url; never paste base64.

GUIDE — call getAuthoringGuide ONCE before authoring or restructuring a page: how to structure one and when a table beats a list, plus the OneNote traps for equations, tables, note tags and styling that are silent and not recoverable after the write.`;

// ---------------------------------------------------------------------------
// Server
// ---------------------------------------------------------------------------
const server = new McpServer(
  { name: 'onenote', version: VERSION, description: 'OneNote MCP Server' },
  { capabilities: { tools: { listChanged: true } }, instructions: SERVER_INSTRUCTIONS }
);

// Shared tool arguments. Terse on purpose — every one is re-sent per request.
const FORMAT = z.enum(['markdown', 'text', 'html']).optional().describe('Default "markdown".');
const CONTENT = (what) => z.string().describe(`${what} (Markdown by default).`);
// z.number().int() writes a 9007199254740991 maximum into the schema that no
// caller reads. A refinement validates the same thing without emitting it.
const whole = (schema) => schema.refine(Number.isInteger, 'Must be a whole number.');
// A row or column: a 0-based index, or text to match it by. Matching by text is
// what lets a caller edit a cell it saw in an outline without fetching the
// table to count rows first.
const REF = z.union([whole(z.number().min(0)), z.string()]);
// Clients use readOnlyHint to decide what needs a confirmation prompt.
const READ_ONLY = { readOnlyHint: true };

// Auth

server.tool(
  'authenticate',
  'Start the Microsoft device-code login. Tokens then auto-refresh for ~90 days.',
  async () => {
    try {
      const accounts = await pca.getTokenCache().getAllAccounts();
      if (accounts.length > 0) {
        try {
          await pca.acquireTokenSilent({ account: accounts[0], scopes: SCOPES });
          return ok('Already authenticated with a valid cached token. No re-login needed.');
        } catch { /* cached account can no longer refresh — fall through to device code */ }
      }
      const result = await pca.acquireTokenByDeviceCode({
        deviceCodeCallback: (info) => { console.error('\n' + info.message + '\n'); },
        scopes: SCOPES,
      });
      return ok(`Authenticated as ${result.account.username}. Tokens will auto-refresh silently.`);
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'saveAccessToken',
  'Save a raw Graph access token. Expires in ~1h with no refresh — prefer "authenticate".',
  { token: z.string().describe('A raw Microsoft Graph access token string.') },
  async ({ token }) => {
    try {
      // Contains a live bearer token — not readable by other accounts.
      writeSecret(tokenFilePath, JSON.stringify({ token }));
      return ok('Token saved. Expires in ~1h — use "authenticate" for auto-refresh.');
    } catch (e) {
      return fail(e);
    }
  }
);

// Guide

server.tool(
  'getAuthoringGuide',
  'The long-form authoring guide: the house style, and every OneNote gotcha for headings/tables/math/tags/images, found by writing real content through the live service (SERVER_INSTRUCTIONS above is the terse, always-sent version). Call ONCE, early, before authoring or restructuring a page.',
  { topic: z.enum(GUIDE_TOPICS).optional().describe('One section only. Omit for the whole guide.') },
  READ_ONLY,
  async ({ topic }) => {
    try {
      return ok(renderGuide(topic));
    } catch (e) {
      return fail(e);
    }
  }
);

// Notebooks

server.tool(
  'listNotebooks',
  'List notebooks, optionally filtered by name (case-insensitive contains).',
  { query: z.string().optional().describe('Name fragment. Omit to list all.') },
  READ_ONLY,
  async ({ query }) => {
    try {
      let v = await allNotebooks();
      if (query) {
        const needle = query.toLowerCase();
        v = v.filter((n) => n.displayName.toLowerCase().includes(needle));
      }
      return ok(v.map((n) => ({ id: handle(n.id, 'nb'), name: n.displayName, isDefault: n.isDefault })));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'createNotebook',
  'Create a new OneNote notebook.',
  { displayName: z.string().describe('Name for the new notebook.') },
  async ({ displayName }) => {
    try {
      const res = await graph('POST', '/me/onenote/notebooks', { body: { displayName }, json: true });
      listCache.delete('notebooks');
      return ok({ id: handle(res.id, 'nb'), name: res.displayName, webUrl: res.links?.oneNoteWebUrl?.href ?? null });
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'getNotebookTree',
  'The whole notebook / section-group / section hierarchy as one nested tree.',
  { notebookId: z.string().optional().describe('Just this notebook; omit for every notebook.') },
  READ_ONLY,
  async ({ notebookId }) => {
    try {
      notebookId = notebookId ? resolveId(notebookId) : undefined;
      const [notebooks, sections, groups] = await Promise.all([
        allNotebooks(),
        allSections(),
        cachedList('sectionGroups', () => graphList(
          '/me/onenote/sectionGroups?$select=id,displayName&$expand=parentNotebook($select=id),parentSectionGroup($select=id)&$top=100'
        )),
      ]);

      const targets = notebookId ? notebooks.filter((n) => n.id === notebookId) : notebooks;

      function buildGroup(g) {
        return {
          id: handle(g.id, 'sg'),
          name: g.displayName,
          sections: sections
            .filter((s) => s.parentSectionGroup?.id === g.id)
            .map((s) => ({ id: handle(s.id, 's'), name: s.displayName })),
          sectionGroups: groups.filter((child) => child.parentSectionGroup?.id === g.id).map(buildGroup),
        };
      }

      const tree = targets.map((nb) => ({
        id: handle(nb.id, 'nb'),
        name: nb.displayName,
        isDefault: nb.isDefault,
        sections: sections
          .filter((s) => s.parentNotebook?.id === nb.id && !s.parentSectionGroup)
          .map((s) => ({ id: handle(s.id, 's'), name: s.displayName })),
        sectionGroups: groups
          .filter((g) => g.parentNotebook?.id === nb.id && !g.parentSectionGroup)
          .map(buildGroup),
      }));

      return ok(notebookId ? (tree[0] ?? null) : tree);
    } catch (e) {
      return fail(e);
    }
  }
);

// Sections

server.tool(
  'listSections',
  'List sections with notebook and group attribution. Pass a query to return only matching names.',
  {
    query: z.string().optional().describe('Name fragment. Omit to list all.'),
    notebookId: z.string().optional().describe('Limit to this notebook.'),
  },
  READ_ONLY,
  async ({ query, notebookId }) => {
    try {
      notebookId = notebookId ? resolveId(notebookId) : undefined;
      let v = await allSections();
      if (notebookId) v = v.filter((s) => s.parentNotebook?.id === notebookId);
      if (query) {
        const needle = query.toLowerCase();
        v = v.filter((s) => s.displayName.toLowerCase().includes(needle));
      }
      return ok(v.map((s) => ({
        id: handle(s.id, 's'),
        name: s.displayName,
        notebook: s.parentNotebook?.displayName ?? null,
        sectionGroup: s.parentSectionGroup?.displayName ?? null,
      })));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'createSection',
  'Create a section under a notebook or section group.',
  {
    displayName: z.string().describe('Name of the new section.'),
    notebookId: z.string().optional().describe('Create directly under this notebook.'),
    sectionGroupId: z.string().optional().describe('Create under this section group instead.'),
  },
  async ({ displayName, notebookId, sectionGroupId }) => {
    try {
      if (!notebookId && !sectionGroupId) throw new Error('Provide either notebookId or sectionGroupId.');
      const parent = sectionGroupId
        ? `/me/onenote/sectionGroups/${resolveId(sectionGroupId)}/sections`
        : `/me/onenote/notebooks/${resolveId(notebookId)}/sections`;
      const res = await graph('POST', parent, { body: { displayName }, json: true });
      listCache.delete('sections');
      return ok({ id: handle(res.id, 's'), name: res.displayName });
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'createSectionGroup',
  'Create a section group under a notebook or another section group.',
  {
    displayName: z.string().describe('Name of the new section group.'),
    notebookId: z.string().optional().describe('Create directly under this notebook.'),
    sectionGroupId: z.string().optional().describe('Create nested under this section group instead.'),
  },
  async ({ displayName, notebookId, sectionGroupId }) => {
    try {
      if (!notebookId && !sectionGroupId) throw new Error('Provide either notebookId or sectionGroupId.');
      const parent = sectionGroupId
        ? `/me/onenote/sectionGroups/${resolveId(sectionGroupId)}/sectionGroups`
        : `/me/onenote/notebooks/${resolveId(notebookId)}/sectionGroups`;
      const res = await graph('POST', parent, { body: { displayName }, json: true });
      listCache.delete('sectionGroups');
      return ok({ id: handle(res.id, 'sg'), name: res.displayName });
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'copySectionToNotebook',
  'Copy a section to another notebook. Async — polls up to 60 s.',
  {
    sectionId: z.string().describe('The section id to copy.'),
    notebookId: z.string().describe('The destination notebook id.'),
    renameAs: z.string().optional().describe('Optional new name for the copied section.'),
  },
  async ({ sectionId, notebookId, renameAs }) => {
    try {
      const body = { id: resolveId(notebookId) };
      if (renameAs) body.renameAs = renameAs;
      let r;
      try {
        r = await graph('POST', `/me/onenote/sections/${resolveId(sectionId)}/copyToNotebook`, { body, json: true });
      } catch (e) {
        // Same account limitation as copyToSection. There is no cheap stand-in
        // for a whole section, so this one says what to do instead of guessing.
        if (e.status !== 501) throw e;
        throw new Error(
          'This account does not implement copyToNotebook (personal Microsoft accounts do not). '
          + 'Create the section with createSection, then copy its pages one at a time with '
          + 'copyPageToSection, which falls back to re-creating each page.'
        );
      }
      listCache.delete('sections');
      if (r && r.kind === 'accepted' && r.operationLocation) {
        const op = await awaitOperation(r.operationLocation);
        return ok({ status: 'completed', id: op.resourceId ? handle(op.resourceId, 's') : null });
      }
      return ok({ status: 'initiated' });
    } catch (e) {
      return fail(e);
    }
  }
);

// Pages — reading

server.tool(
  'listPages',
  'List pages in a section, newest first. previewChars adds a snippet per page.',
  {
    sectionId: z.string().describe('The section id (from listSections).'),
    top: whole(z.number().min(1).max(100)).optional().describe('Max pages (default 50; 20 with previews).'),
    previewChars: whole(z.number().min(1).max(1000)).optional().describe('Chars of content preview per page.'),
    includeDates: z.boolean().optional().describe('Include created/lastModified. Default false.'),
  },
  READ_ONLY,
  async ({ sectionId, top, previewChars, includeDates }) => {
    try {
      sectionId = resolveId(sectionId);
      const select = includeDates ? 'id,title,createdDateTime,lastModifiedDateTime' : 'id,title';
      const limit = top || (previewChars ? 20 : 50);
      // Single request capped at `limit` — do NOT follow nextLink, or `limit`
      // becomes a page size instead of a hard cap (sequential slow calls).
      const r = await graph(
        'GET',
        `/me/onenote/sections/${sectionId}/pages?$select=${select}&$orderby=lastModifiedDateTime desc&$top=${limit}`
      );
      const pages = (r && r.value) || [];
      const base = pages.map((p) => ({
        id: handle(p.id, 'p'),
        title: p.title,
        ...(includeDates ? { created: p.createdDateTime, lastModified: p.lastModifiedDateTime } : {}),
      }));
      if (!previewChars) return ok(base);

      // Previews are bounded to the listed pages rather than every page in the
      // section.
      await inBatches(pages, 4, async (p, i) => {
        try {
          base[i].preview = (await htmlToText(await getPageHtml(p.id, { includeIds: false }))).slice(0, previewChars);
        } catch {
          base[i].preview = null;
        }
      });
      return ok(base);
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'getPage',
  'Read a page. "markdown" (default) can be written straight back; "outline" is a cheap map of it. Use "html" only for element ids.',
  {
    pageId: z.string().describe('The page id.'),
    format: z.enum(['markdown', 'outline', 'text', 'html']).optional().describe('Default "markdown". "outline" maps the headings and tables; "html" is raw XHTML.'),
    maxChars: whole(z.number().min(1)).optional().describe('Truncate output at this many characters. Applies to the serialised result, outline included.'),
    includeIds: z.boolean().optional().describe('With format "html", add element ids for replaceElement targets.'),
  },
  READ_ONLY,
  async ({ pageId, format, maxChars, includeIds }) => {
    try {
      pageId = resolveId(pageId);
      const fmt = format ?? 'markdown';
      const html = await getPageHtml(pageId, { includeIds: fmt === 'html' ? !!includeIds : false });
      const out = await asFormat(html, fmt);
      if (!maxChars) return ok(out);
      // An outline is an object, so truncating the raw value would do nothing at
      // all. Serialise first, then cut — the caller asked for a byte ceiling.
      return ok((typeof out === 'string' ? out : JSON.stringify(out)).slice(0, maxChars));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'getPageSection',
  'Read only the content under one heading, down to the next heading of equal or higher level.',
  {
    pageId: z.string().describe('The page id.'),
    heading: z.string().describe('Text fragment identifying the heading.'),
    format: FORMAT,
  },
  READ_ONLY,
  async ({ pageId, heading, format }) => {
    try {
      pageId = resolveId(pageId);
      const { body } = await fetchPageDom(pageId);
      const { nodes, start, end } = sectionSpan(contentContainer(body), heading);
      const fragment = nodes.slice(start, end).map((n) => n.outerHTML).join('');
      const fmt = format ?? 'markdown';
      return ok(await asFormat(fmt === 'html' ? fragment : `<body>${fragment}</body>`, fmt));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'getMathExpressions',
  'Read only the <math> blocks on a page, each with text context. Pass a returned "mathml" to findAndReplace to fix one formula.',
  {
    pageId: z.string().describe('The page id.'),
    contextChars: whole(z.number().min(0).max(500)).optional().describe('Context chars per formula (default 160).'),
  },
  READ_ONLY,
  async ({ pageId, contextChars }) => {
    try {
      pageId = resolveId(pageId);
      const { doc } = await fetchPageDom(pageId);
      const cc = contextChars ?? 160;
      const expressions = [...doc.querySelectorAll('math')].map((m, i) => {
        let el = m, elementId = null;
        while (el) {
          if (el.getAttribute && el.getAttribute('id')) { elementId = el.getAttribute('id'); break; }
          el = el.parentElement;
        }
        const container = m.closest('p,h1,h2,h3,h4,li,td') || m;
        const context = (container.textContent || '').replace(/\s+/g, ' ').trim().slice(0, cc);
        return { index: i, elementId, mathml: m.outerHTML, context };
      });
      return ok({ count: expressions.length, expressions });
    } catch (e) {
      return fail(e);
    }
  }
);

// Personal accounts get no server-side full-text search, so matching content
// means reading every page in scope. Bounded, so a wide search degrades into a
// truthful partial answer instead of a hundred-page stall.
const MAX_CONTENT_SCAN = 150;

/** A preview centred on the hit — the top of the page rarely says why it matched. */
const snippetAround = (text, needle, chars) => {
  const at = text.toLowerCase().indexOf(needle);
  if (at < 0) return text.slice(0, chars);
  const from = Math.max(0, at - Math.floor(chars / 3));
  return (from ? '…' : '') + text.slice(from, from + chars) + (from + chars < text.length ? '…' : '');
};

server.tool(
  'searchPages',
  'Find pages by title, or by their content with inContent. Scope with sectionId or notebookId.',
  {
    query: z.string().describe('Text to match.'),
    sectionId: z.string().optional().describe('Limit search to this section.'),
    notebookId: z.string().optional().describe('Limit search to sections in this notebook.'),
    inContent: z.boolean().optional().describe('Also match page bodies. Needs a section or notebook scope.'),
    previewChars: whole(z.number().min(1).max(2000)).optional().describe('Chars of content preview per match.'),
  },
  READ_ONLY,
  async ({ query, sectionId, notebookId, inContent, previewChars }) => {
    try {
      const needle = query.toLowerCase();
      let sections;
      if (sectionId) {
        sections = [{ id: resolveId(sectionId), displayName: null, parentNotebook: null }];
      } else {
        sections = await allSections();
        if (notebookId) {
          const nb = resolveId(notebookId);
          sections = sections.filter((s) => s.parentNotebook?.id === nb);
        }
      }
      if (inContent && !sectionId && !notebookId) {
        throw new Error('inContent needs sectionId or notebookId: matching content reads every page in scope.');
      }

      const listed = await inBatches(sections, 8, async (s) => {
        try {
          return { s, pages: await sectionPages(s.id) };
        } catch {
          return { s, pages: [] };
        }
      });

      // Attribution is omitted rather than sent as null.
      const found = (s, p, match) => ({
        _rid: p.id, // real id, kept only for the preview fetch below
        id: handle(p.id, 'p'),
        title: p.title,
        ...(inContent ? { match } : {}),
        sectionId: handle(s.id, 's'),
        ...(s.displayName ? { section: s.displayName } : {}),
        ...(s.parentNotebook?.displayName ? { notebook: s.parentNotebook.displayName } : {}),
      });

      const results = [];
      const unmatched = [];
      for (const { s, pages } of listed) {
        for (const p of pages) {
          if (p.title && p.title.toLowerCase().includes(needle)) results.push(found(s, p, 'title'));
          else if (inContent) unmatched.push({ s, p });
        }
      }

      const scanned = unmatched.slice(0, MAX_CONTENT_SCAN);
      await inBatches(scanned, 4, async ({ s, p }) => {
        try {
          const text = await htmlToText(await getPageHtml(p.id, { includeIds: false }));
          if (!text.toLowerCase().includes(needle)) return;
          results.push({ ...found(s, p, 'content'), preview: snippetAround(text, needle, previewChars ?? 200) });
        } catch { /* a page that cannot be read is not a match */ }
      });

      // Content hits already carry a snippet; only title hits need one.
      if (previewChars) {
        await inBatches(results.filter((r) => r.preview === undefined), 4, async (match) => {
          try {
            match.preview = (await htmlToText(await getPageHtml(match._rid, { includeIds: false }))).slice(0, previewChars);
          } catch {
            match.preview = null;
          }
        });
      }

      for (const r of results) delete r._rid;
      const out = { matches: results.length, scannedSections: sections.length, pages: results };
      if (inContent) out.scannedPages = scanned.length;
      if (unmatched.length > scanned.length) {
        out.truncated = `Only the first ${MAX_CONTENT_SCAN} pages were read — narrow the scope to search the rest.`;
      }
      return ok(out);
    } catch (e) {
      return fail(e);
    }
  }
);

// Pages — writing

server.tool(
  'createPage',
  'Create a page. The server applies the house style to Markdown content.',
  {
    sectionId: z.string().describe('The target section id.'),
    title: z.string().describe('The page title.'),
    content: z.string().optional().describe('Body content (Markdown by default). Default: empty.'),
    format: FORMAT,
    lang: z.string().optional().describe('Language tag for spell-check, e.g. "de-CH".'),
  },
  async ({ sectionId, title, content, format, lang }) => {
    try {
      sectionId = resolveId(sectionId);
      const { html: body, warnings } = await prepareContent(content ?? '', format ?? 'markdown', 'page');
      const html =
        `<!DOCTYPE html><html${lang ? ` lang="${escapeHtml(lang)}"` : ''}><head>` +
        `<title>${escapeHtml(title)}</title>` +
        `<meta name="created" content="${new Date().toISOString()}"/></head>` +
        `<body>${body || '<p></p>'}</body></html>`;
      const res = await graph('POST', `/me/onenote/sections/${sectionId}/pages`, {
        body: html,
        contentType: 'application/xhtml+xml',
      });
      invalidatePageLists();
      return ok(withWarnings({
        id: handle(res.id, 'p'),
        title: res.title,
        createdDateTime: res.createdDateTime,
        webUrl: res.links?.oneNoteWebUrl?.href ?? null,
      }, warnings));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'replacePage',
  'Replace a page\'s entire body. The widest edit — prefer a narrower tool for a partial change.',
  {
    pageId: z.string().describe('The page id.'),
    content: CONTENT('New body content'),
    format: FORMAT,
    allowImageLoss: z.boolean().optional().describe('Permit dropping existing images (refused by default).'),
  },
  async ({ pageId, content, format, allowImageLoss }) => {
    try {
      pageId = resolveId(pageId);
      const { html, warnings } = await prepareContent(content, format ?? 'markdown', 'page');
      const existing = await getPageHtml(pageId, { includeIds: false });
      assertNoLostImages(html, existing, { allowImageLoss });
      await patchPage(pageId, [{ target: 'body', action: 'replace', content: html }]);
      return ok(withWarnings('Page body replaced.', warnings));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'addToPage',
  'Append or prepend content to a page body.',
  {
    pageId: z.string().describe('The page id.'),
    content: CONTENT('Content to add'),
    where: z.enum(['end', 'start']).optional().describe('"end" (append, default) or "start" (prepend).'),
    format: FORMAT,
  },
  async ({ pageId, content, where, format }) => {
    try {
      pageId = resolveId(pageId);
      const { html, warnings } = await prepareContent(content, format ?? 'markdown', 'content');
      const action = where === 'start' ? 'prepend' : 'append';
      await patchPage(pageId, [{ target: 'body', action, content: html }]);
      return ok(withWarnings(`Content added at the ${where === 'start' ? 'start' : 'end'} of the page.`, warnings));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'setPageTitle',
  'Change the title of a page.',
  {
    pageId: z.string().describe('The page id.'),
    newTitle: z.string().describe('The new title.'),
  },
  async ({ pageId, newTitle }) => {
    try {
      pageId = resolveId(pageId);
      // patchPage normalises void tags but does not escape, and this content is
      // a bare title rather than converted markup — so escape it here or a
      // title containing "<" reaches the PATCH body verbatim.
      await patchPage(pageId, [{ target: 'title', action: 'replace', content: escapeHtml(newTitle) }]);
      return ok(`Title updated to "${newTitle}".`);
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'insertNearText',
  'Insert content before or after the paragraph containing a text fragment.',
  {
    pageId: z.string().describe('The page id.'),
    nearText: z.string().describe('Text fragment identifying the target element.'),
    content: CONTENT('Content to insert'),
    position: z.enum(['before', 'after']).optional().describe('"after" (default) or "before" the matched element.'),
    format: FORMAT,
  },
  async ({ pageId, nearText, content, position, format }) => {
    try {
      pageId = resolveId(pageId);
      const { html, warnings } = await prepareContent(content, format ?? 'markdown', 'content');
      const { elementId } = await findElementByText(pageId, nearText);
      await patchPage(pageId, [{ target: elementId, action: 'insert', position: position ?? 'after', content: html }]);
      return ok(withWarnings(`Content inserted ${position ?? 'after'} the element containing "${nearText}".`, warnings));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'replaceElement',
  'Replace one block element by its id (from getMathExpressions, or getPage html+includeIds). Ids are snapshot-bound.',
  {
    pageId: z.string().describe('The page id.'),
    elementId: z.string().describe('Element id, e.g. "p:{guid}{12}".'),
    content: CONTENT('Replacement content'),
    format: FORMAT,
  },
  async ({ pageId, elementId, content, format }) => {
    try {
      pageId = resolveId(pageId);
      const { html, warnings } = await prepareContent(content, format ?? 'markdown', 'element');
      await patchPage(pageId, [{ target: elementId, action: 'replace', content: html }]);
      return ok(withWarnings('Element replaced.', warnings));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'findAndReplace',
  'Swap an exact HTML fragment (e.g. one formula from getMathExpressions) or a run of plain text. Errors if it is absent rather than corrupting the page.',
  {
    pageId: z.string().describe('The page id.'),
    find: z.string().describe('Exact HTML fragment, or a run of visible text.'),
    replace: z.string().describe('Replacement. Raw HTML, not Markdown; text-only when "find" matched text.'),
    all: z.boolean().optional().describe('Replace all occurrences (default true). False = first only.'),
  },
  async ({ pageId, find, replace, all }) => {
    try {
      pageId = resolveId(pageId);
      const mc = await checkMathInContent(replace);
      if (mc.hardError) throw new Error('Replacement contains invalid MathML, not written: ' + mc.hardError);

      const { doc } = await fetchPageDom(pageId);

      // Only elements OneNote can actually 'replace' are eligible: an
      // absolute-positioned div (the page's top-level content container) is
      // NOT replaceable, and targeting one corrupts the page rather than
      // failing loudly. See the supported target/action matrix in
      // https://learn.microsoft.com/en-us/graph/onenote-update-page
      const isReplaceable = (el) => {
        const tag = el.tagName.toLowerCase();
        if (/^(p|h[1-6]|li|ol|ul|table)$/.test(tag)) return true;
        return tag === 'div' && el.parentElement !== doc.body; // nested divs only
      };
      // Scope the write to the smallest replaceable block containing the change,
      // so a one-word fix does not have the whole page as its blast radius.
      const blockOf = (node) => {
        let el = node.nodeType === 1 ? node : node.parentElement;
        while (el && el !== doc.body) {
          if (el.getAttribute('id') && isReplaceable(el)) return el;
          el = el.parentElement;
        }
        return null;
      };
      const commandFor = (blocks) => {
        // Nothing replaceable holds the change (e.g. a block-level equation
        // sitting directly in the page's absolute div), or it spans several
        // blocks: fall back to the whole body.
        if (blocks.length !== 1 || !blocks[0]) {
          return { scope: 'body', command: [{ target: 'body', action: 'replace', content: doc.body.innerHTML }] };
        }
        const target = blocks[0];
        return {
          scope: target.tagName.toLowerCase(),
          command: [{ target: target.getAttribute('id'), action: 'replace', content: target.outerHTML }],
        };
      };

      const countIn = (text) => text.split(find).length - 1;

      let mode, occurrences, onPage, scope, command;
      if (doc.body.innerHTML.includes(find)) {
        // Markup mode: "find" is a literal slice of the page's HTML.
        mode = 'html';
        onPage = countIn(doc.body.innerHTML); // count before mutating
        const target = [...doc.body.querySelectorAll('[id]')]
          .filter((el) => el.innerHTML.includes(find) && isReplaceable(el))
          .sort((a, b) => a.innerHTML.length - b.innerHTML.length)[0];
        // Only the written block changes, so the count has to come from it and
        // not from the page, or a partial fix would report as a complete one.
        occurrences = all === false ? 1 : countIn(target ? target.innerHTML : doc.body.innerHTML);
        const swap = (h) => (all === false ? h.replace(find, replace) : h.split(find).join(replace));
        if (target) target.innerHTML = swap(target.innerHTML);
        else doc.body.innerHTML = swap(doc.body.innerHTML);
        ({ scope, command } = commandFor([target || null]));
      } else {
        // Text mode: the fragment is visible text, so entities, tag boundaries
        // and OneNote's own markup do not have to be guessed at.
        const hits = textNodesUnder(doc.body).filter((t) => t.data.includes(find));
        if (!hits.length) {
          const mathCount = (doc.body.innerHTML.match(/<math/gi) || []).length;
          throw new Error(
            `"${find}" is on the page neither as markup nor as text within a single run ` +
            `(page has ${mathCount} <math> block(s)). For a formula, copy the exact string from ` +
            `getMathExpressions — OneNote normalises MathML on write, so the stored markup differs ` +
            `from what was sent. For text, check that it is not split by bold or a link.`
          );
        }
        if (/<[a-z][\s\S]*>/i.test(replace)) {
          throw new Error(
            `"find" matched plain text, but the replacement contains markup, which would be ` +
            `inserted literally. Use the exact HTML fragment from getPage(format:"html") as "find".`
          );
        }
        mode = 'text';
        onPage = hits.reduce((n, node) => n + countIn(node.data), 0);
        const done = replaceInTextNodes(hits, find, replace, { all: all !== false });
        occurrences = done.occurrences;
        ({ scope, command } = commandFor([...new Set(done.touched.map(blockOf))]));
      }

      await patchPage(pageId, command);
      const result = { status: 'replaced', occurrences, mode, scope };
      // Left behind either by all:false or by scoping the write to one block.
      if (onPage > occurrences) result.remaining = onPage - occurrences;
      return ok(withWarnings(result, mc.warnings));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'editTable',
  'Edit a table in place: set cells, add or delete rows and columns. Untouched cells keep their formatting.',
  {
    pageId: z.string().describe('The page id.'),
    nearText: z.string().optional().describe('Find the table by text in it, or just above it.'),
    tableIndex: whole(z.number().min(0)).optional().describe('By position (0-based), when nearText is omitted.'),
    setCells: z.array(z.object({
      row: REF.describe('Row index (0 is the header), or text in the row.'),
      col: REF.describe('Column index, or header name.'),
      value: z.string().describe('Inline Markdown.'),
    })).optional().describe('Cell updates.'),
    addRows: z.array(z.array(z.string())).optional().describe('Rows of cells to add.'),
    atRow: REF.optional().describe('Insert addRows here. Default: append.'),
    deleteRows: z.array(REF).optional().describe('Rows to delete.'),
    addColumns: z.array(z.array(z.string())).optional().describe('Columns of cells to add; first cell is the header.'),
    atCol: REF.optional().describe('Insert addColumns here. Default: append.'),
    deleteColumns: z.array(REF).optional().describe('Columns to delete.'),
  },
  // Everything after the three locator arguments is a grid edit, handed to
  // editGrid as one object rather than re-listed here.
  async ({ pageId, nearText, tableIndex, ...edits }) => {
    try {
      pageId = resolveId(pageId);
      if (!['setCells', 'addRows', 'deleteRows', 'addColumns', 'deleteColumns'].some((k) => edits[k])) {
        throw new Error('Nothing to do: pass setCells, addRows, deleteRows, addColumns and/or deleteColumns.');
      }
      const { doc } = await fetchPageDom(pageId, { fresh: true });
      const tables = [...doc.body.querySelectorAll('table')];
      if (!tables.length) throw new Error('This page has no tables.');

      let table;
      if (nearText) {
        const needle = nearText.toLowerCase();
        table = tables.find((t) => t.textContent.toLowerCase().includes(needle));
        if (!table) {
          // Not inside a table — try the first table after a matching element.
          const anchor = [...doc.body.querySelectorAll('*')]
            .filter((el) => el.textContent.toLowerCase().includes(needle) && !el.querySelector('table'))
            .sort((a, b) => a.textContent.length - b.textContent.length)[0];
          if (anchor) {
            table = tables.find((t) => anchor.compareDocumentPosition(t) & 4 /* FOLLOWING */);
          }
        }
        if (!table) throw new Error(`No table found in or after an element containing "${nearText}".`);
      } else {
        table = tables[tableIndex ?? 0];
        if (!table) throw new Error(`tableIndex ${tableIndex} is out of range (page has ${tables.length} table(s)).`);
      }

      const elementId = table.getAttribute('id');
      if (!elementId) {
        throw new Error('That table has no generated element id, so OneNote cannot replace it. Rewrite the section with replaceElement or replacePage instead.');
      }

      // Cells carry their HTML across, so editing one cell does not flatten the
      // formatting of the rest — OneNote cannot patch a <td>, so the whole
      // table is always rewritten.
      const { grid, hasHeader } = tableToGrid(table);
      const { grid: next, width } = editGrid(grid, hasHeader, edits);

      const html = renderTable({
        headers: hasHeader ? next[0] : null,
        rows: hasHeader ? next.slice(1) : next,
        dataId: table.getAttribute('data-id') || undefined,
      });
      const mc = await checkMathInContent(html);
      if (mc.hardError) throw new Error(`Invalid MathML, table not written: ${mc.hardError}`);
      await patchPage(pageId, [{ target: elementId, action: 'replace', content: html }]);
      return ok(withWarnings(
        { status: 'table updated', rows: next.length, columns: width, header: hasHeader }, mc.warnings
      ));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'setCheckbox',
  'Tick or untick check boxes by their text. The page is never read out.',
  {
    pageId: z.string().describe('The page id.'),
    text: z.union([z.string(), z.array(z.string())]).describe('Text of the check box(es).'),
    checked: z.boolean().optional().describe('Default true; false unticks.'),
  },
  async ({ pageId, text, checked }) => {
    try {
      pageId = resolveId(pageId);
      const wanted = Array.isArray(text) ? text : [text];
      if (!wanted.length) throw new Error('Pass the text of at least one check box.');
      // Fresh: the write targets generated element ids, which a page edit
      // elsewhere in the conversation would already have invalidated.
      const { doc } = await fetchPageDom(pageId, { fresh: true });
      const boxes = checkboxesIn(doc.body);
      if (!boxes.length) throw new Error('This page has no check boxes. Write them as Markdown task lists: "- [ ] …".');

      const ticking = checked !== false;
      // Keyed by element id: two fragments can name one box, and writing that
      // box twice in one PATCH fails the whole command.
      const blocks = new Map();
      const done = [];
      for (const fragment of wanted) {
        const box = pickOne(boxes, fragment, 'check box', 'check boxes');
        if (!box.block) {
          throw new Error(`The check box "${box.text}" is not inside an element OneNote can replace. Use editPageSection instead.`);
        }
        box.tagged.setAttribute('data-tag', ticking ? 'to-do:completed' : 'to-do');
        const id = box.block.getAttribute('id');
        if (!blocks.has(id)) done.push(box.text);
        blocks.set(id, box.block);
      }
      // Nothing is written until every fragment has resolved, so a typo in the
      // third one leaves the page as it was rather than half-ticked.
      await patchPage(pageId, [...blocks].map(([target, el]) => ({
        target, action: 'replace', content: el.outerHTML,
      })));
      return ok({ status: ticking ? 'ticked' : 'unticked', items: done });
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'editPageSection',
  'Replace, add to or delete the content under one heading. The body is never read out or re-sent.',
  {
    pageId: z.string().describe('The page id.'),
    heading: z.string().describe('Text fragment identifying the heading.'),
    content: CONTENT('Content for the section').optional(),
    mode: z.enum(['replace', 'append', 'prepend', 'delete']).optional().describe('Default "replace"; keeps the heading. "delete" removes the heading and its content.'),
    format: FORMAT,
    allowImageLoss: z.boolean().optional().describe('Permit dropping images inside the section (refused by default).'),
  },
  async ({ pageId, heading, content, mode, format, allowImageLoss }) => {
    try {
      pageId = resolveId(pageId);
      const how = mode ?? 'replace';
      if (how !== 'delete' && content == null) throw new Error('content is required unless mode is "delete".');
      const { html: fragment, warnings } = how === 'delete'
        ? { html: '', warnings: [] }
        : await prepareContent(content, format ?? 'markdown', 'section');
      // Fresh: this is a read-modify-write of the whole body, so a stale read
      // would revert edits made from the OneNote app in the meantime.
      const { html: existing, doc } = await fetchPageDom(pageId, { fresh: true });
      spliceSection(doc, heading, fragment, how);
      assertNoLostImages(doc.body.innerHTML, existing, { allowImageLoss });
      await patchPage(pageId, [{ target: 'body', action: 'replace', content: doc.body.innerHTML }]);
      const past = { replace: 'replaced', append: 'appended', prepend: 'prepended', delete: 'deleted' }[how];
      return ok(withWarnings(`Section "${heading}" ${past}.`, warnings));
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'batchUpdatePages',
  'Apply one update to many pages. Returns a per-page success/error list.',
  {
    updates: z.array(z.object({
      pageId: z.string().describe('The page id.'),
      content: CONTENT('Content to apply'),
      action: z.enum(['append', 'prepend', 'replace', 'insert']).optional().describe('Default "append".'),
      target: z.string().optional().describe('Default "body"; or "title" or an element id.'),
      format: FORMAT,
    })).min(1).max(50),
  },
  async ({ updates }) => {
    try {
      const results = await inBatches(updates, 4, async (u) => {
        try {
          const pid = resolveId(u.pageId);
          const { html, warnings } = await prepareContent(u.content, u.format ?? 'markdown', 'page');
          await patchPage(pid, [{ target: u.target || 'body', action: u.action || 'append', content: html }]);
          return warnings.length ? { pageId: u.pageId, ok: true, warnings } : { pageId: u.pageId, ok: true };
        } catch (e) {
          return { pageId: u.pageId, ok: false, error: (e.message || String(e)).slice(0, 160) };
        }
      });
      return ok({ total: results.length, succeeded: results.filter((r) => r.ok).length, results });
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'addImageToPage',
  'Upload an image from a local filePath or url. Appends at the end, or near nearText. Max ~3.5 MB; png/jpg/gif/bmp/tiff.',
  {
    pageId: z.string().describe('The page id.'),
    filePath: z.string().optional().describe('Absolute path to an image file.'),
    url: z.string().optional().describe('Image URL; the server fetches it.'),
    nearText: z.string().optional().describe('Place it near the paragraph containing this text.'),
    position: z.enum(['before', 'after']).optional().describe('With nearText: "after" (default) or "before".'),
    width: whole(z.number().min(1)).optional().describe('Display width in px.'),
    height: whole(z.number().min(1)).optional().describe('Display height in px.'),
    altText: z.string().optional().describe('Alt text.'),
  },
  async ({ pageId, filePath, url, nearText, position, width, height, altText }) => {
    try {
      pageId = resolveId(pageId);
      if (!filePath && !url) throw new Error('Provide either filePath or url.');
      const { buf, type } = await loadImageBytes(filePath, url);
      // loadImageBytes already enforces the cap for a local file; a fetched URL
      // can only be measured once it has arrived.
      if (buf.length > MAX_IMAGE_BYTES) {
        throw new Error(`Image is ${(buf.length / 1048576).toFixed(1)} MB; the Graph request limit is ~4 MB.`);
      }
      const partName = 'img' + Date.now();
      const attrs = [`src="name:${partName}"`, `alt="${escapeHtml(altText || '')}"`];
      if (width) attrs.push(`width="${width}"`);
      if (height) attrs.push(`height="${height}"`);
      const imgHtml = `<img ${attrs.join(' ')} />`;

      const command = nearText
        ? [{
            target: (await findElementByText(pageId, nearText)).elementId,
            action: 'insert',
            position: position || 'after',
            content: imgHtml,
          }]
        : [{ target: 'body', action: 'append', content: imgHtml }];

      const fd = new FormData();
      fd.append('Commands', new Blob([JSON.stringify(command)], { type: 'application/json' }));
      fd.append(partName, new Blob([buf], { type }));
      await graph('PATCH', `/me/onenote/pages/${pageId}/content`, { formData: fd });
      invalidatePage(pageId);
      return ok(`Image added (${(buf.length / 1024).toFixed(0)} KB, ${type})${nearText ? ` ${position || 'after'} "${nearText}"` : ''}.`);
    } catch (e) {
      return fail(e);
    }
  }
);

/**
 * Rebuild a page in another section from its own stored HTML — the stand-in
 * for copyToSection on accounts that do not implement it.
 *
 * This is a re-creation, not a copy, and says so: the new page gets a new id
 * and a new creation date. Images are the one thing that cannot survive, since
 * they live as authenticated resources that cannot be re-submitted, so a page
 * carrying any is refused rather than silently arriving without them.
 */
async function recreateElsewhere(pid, sectionId, { renameAs, move, allowImageLoss }) {
  const pageHtml = await getPageHtml(pid, { includeIds: false });
  const doc = await parseHtml(pageHtml);
  const images = doc.querySelectorAll('img').length;
  if (images && !allowImageLoss) {
    throw new Error(
      `This account has no copy API, so the page has to be re-created from its HTML — but it ` +
      `holds ${images} image(s), which are authenticated resources that cannot be re-submitted. ` +
      `Pass allowImageLoss:true to proceed without them, or re-add them with addImageToPage ` +
      `after the copy.`
    );
  }
  const title = renameAs || doc.querySelector('title')?.textContent || 'Untitled';
  // The absolute-positioned wrapper is the page's own furniture, not content.
  const inner = contentContainer(doc.body).innerHTML;
  const res = await graph('POST', `/me/onenote/sections/${sectionId}/pages`, {
    body: `<!DOCTYPE html><html><head><title>${escapeHtml(title)}</title>`
      + `<meta name="created" content="${new Date().toISOString()}"/></head>`
      + `<body>${toXhtml(inner) || '<p></p>'}</body></html>`,
    contentType: 'application/xhtml+xml',
  });
  invalidatePageLists();
  const out = {
    status: move ? 're-created and original deleted' : 're-created',
    id: handle(res.id, 'p'),
    note: 'This account does not implement copyToSection, so the page was rebuilt from its HTML: '
      + 'new id, new creation date, and any ink or embedded files were not carried over.',
  };
  if (images) out.imagesDropped = images;
  if (!move) return out;
  try {
    await graph('DELETE', `/me/onenote/pages/${pid}`);
    invalidatePage(pid);
  } catch (e) {
    out.status = 're-created but original NOT deleted';
    out.warning = `The page now exists in both sections. Deleting the original failed: ${e.message}`;
  }
  return out;
}

// Pages — managing

server.tool(
  'deletePage',
  'Delete a page by id.',
  { pageId: z.string().describe('The page id to delete.') },
  { destructiveHint: true, idempotentHint: true },
  async ({ pageId }) => {
    try {
      const pid = resolveId(pageId);
      await graph('DELETE', `/me/onenote/pages/${pid}`);
      invalidatePage(pid);
      invalidatePageLists();
      return ok('Page deleted.');
    } catch (e) {
      return fail(e);
    }
  }
);

server.tool(
  'copyPageToSection',
  'Copy a page to another section, or move it. Falls back to re-creating the page when the account has no copy API.',
  {
    pageId: z.string().describe('The page id.'),
    sectionId: z.string().describe('The destination section id.'),
    renameAs: z.string().optional().describe('Optional new title for the copy. Ignored when move is true.'),
    move: z.boolean().optional().describe('Delete the original after a successful copy (default false).'),
    allowImageLoss: z.boolean().optional().describe('Permit the fallback to drop images (refused by default).'),
  },
  async ({ pageId, sectionId, renameAs, move, allowImageLoss }) => {
    try {
      const pid = resolveId(pageId);
      const target = resolveId(sectionId);
      const body = { id: target };
      if (renameAs && !move) body.renameAs = renameAs;
      let r;
      try {
        r = await graph('POST', `/me/onenote/pages/${pid}/copyToSection`, { body, json: true });
      } catch (e) {
        // Personal Microsoft accounts answer copyToSection with 501, so the
        // copy API simply is not there. Rebuilding the page from its own HTML
        // is the same thing by hand.
        if (e.status !== 501) throw e;
        return ok(await recreateElsewhere(pid, target, { renameAs, move, allowImageLoss }));
      }
      invalidatePageLists();
      let newId = null;
      if (r && r.kind === 'accepted' && r.operationLocation) {
        const op = await awaitOperation(r.operationLocation);
        newId = op.resourceId ? handle(op.resourceId, 'p') : null;
      } else if (!move) {
        return ok({ status: 'initiated' });
      }
      if (!move) return ok({ status: 'copied', id: newId });

      // Copy succeeded; the delete is a separate call, so report honestly if it
      // fails rather than claiming an atomic move.
      try {
        await graph('DELETE', `/me/onenote/pages/${pid}`);
      } catch (e) {
        return ok({
          status: 'copied but original NOT deleted',
          id: newId,
          warning: `The page now exists in both sections. Deleting the original failed: ${e.message}`,
        });
      }
      invalidatePage(pid);
      return ok({ status: 'moved', id: newId });
    } catch (e) {
      return fail(e);
    }
  }
);

// Start
async function main() {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error(`OneNote MCP server v${VERSION} started (short id handles enabled).`);
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length === 0) {
      console.error('No cached session — run the "authenticate" tool or `npm run auth`.');
    } else {
      console.error(`Cached session for ${accounts[0].username}. Token will auto-refresh as needed.`);
    }
    // Persisting handles is debounced by 250 ms and the timer is unref'd, so a
    // handle minted for a just-created page is still only in memory when the
    // process ends. Without this, the id handed back by the very last call of
    // a session can come back "Unknown id handle" on the next one.
    process.on('exit', () => ids.flush());
    process.on('SIGINT', () => process.exit(0));
    process.on('SIGTERM', () => process.exit(0));
  } catch (error) {
    console.error('Error starting server:', error);
    process.exit(1);
  }
}

main();
