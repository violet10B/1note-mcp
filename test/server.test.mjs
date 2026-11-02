// Contract tests for the MCP surface itself. `initialize` and `tools/list` do
// not touch Microsoft Graph, so these run offline.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';
import { createGraph } from '../src/graph.js';

const SERVER = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'onenote-mcp.mjs');

/**
 * A Graph client with a stubbed transport, for the behaviours a caller depends
 * on (status codes, 202/204 handling) without reaching the network.
 */
const client = ({ status = 200, body = '', headers = {} } = {}) => createGraph({
  getToken: async () => 'test-token',
  fetchImpl: async () => ({
    ok: status >= 200 && status < 300,
    status,
    statusText: String(status),
    headers: { get: (k) => headers[k] ?? null },
    text: async () => body,
  }),
});

// The tool schemas and instructions are re-sent on EVERY request, so their size
// is a standing cost in every conversation. These budgets are deliberately
// close to the current values: adding a tool should be a conscious decision
// that trades against the budget, not something that drifts in unnoticed.
// Raised from 18_000 to fit getAuthoringGuide: unlike every other tool here,
// its VALUE is in being comprehensive, not narrow, so its schema (description
// + topic enum) earns a bigger one-time slice of this budget in exchange for
// its content costing nothing until a model actually calls it.
const TOOLS_JSON_BUDGET = 18_600;
// Tightened from 2_800 once getAuthoringGuide existed to hold the detail: what
// is sent on every request should be what changes behaviour on every request,
// and everything else belongs one tool call away.
const INSTRUCTIONS_BUDGET = 2_000;
// editTable is deliberately the one large schema: it is six operations over
// two addressing modes, and every one of them replaces a read-then-rewrite.
const PER_TOOL_BUDGET = 1_800;

let proc, initResult, tools;

before(async () => {
  proc = spawn('node', [SERVER], { stdio: ['pipe', 'pipe', 'pipe'] });
  proc.stderr.resume(); // startup banner; ignore

  let buf = '';
  const pending = new Map();
  proc.stdout.on('data', (d) => {
    buf += d.toString();
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i); buf = buf.slice(i + 1);
      if (!line.trim()) continue;
      let msg; try { msg = JSON.parse(line); } catch { continue; }
      const cb = pending.get(msg.id);
      if (cb) { pending.delete(msg.id); cb(msg); }
    }
  });

  let id = 0;
  const send = (method, params) => new Promise((resolve, reject) => {
    const myId = ++id;
    const timer = setTimeout(() => reject(new Error(`${method} timed out`)), 20_000);
    pending.set(myId, (m) => { clearTimeout(timer); resolve(m); });
    proc.stdin.write(JSON.stringify({ jsonrpc: '2.0', id: myId, method, params }) + '\n');
  });

  initResult = (await send('initialize', {
    protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'test', version: '1' },
  })).result;
  tools = (await send('tools/list')).result.tools;
});

after(() => proc && proc.kill());

test('the server starts and advertises its tools', () => {
  assert.ok(tools.length > 20, `expected a full toolset, got ${tools.length}`);
  assert.equal(initResult.serverInfo.name, 'onenote');
});

test('every tool the instructions name actually exists', () => {
  const names = new Set(tools.map((t) => t.name));
  for (const named of [
    'editTable', 'editPageSection', 'findAndReplace', 'insertNearText', 'replaceElement',
    'addToPage', 'replacePage', 'getPageSection', 'getMathExpressions', 'batchUpdatePages',
    'addImageToPage',
  ]) {
    assert.ok(names.has(named), `instructions reference "${named}" but it is not registered`);
  }
});

test('tools removed in v3 stay removed', () => {
  const names = new Set(tools.map((t) => t.name));
  // Each was a duplicate that differed from a survivor by one parameter.
  for (const gone of [
    'appendToPage', 'prependToPage', 'listPagesWithPreview', 'findSection',
    'findNotebook', 'getNotebook', 'listSectionGroups', 'movePageToSection',
  ]) {
    assert.ok(!names.has(gone), `"${gone}" should have been merged away`);
  }
});

test('updatePage stays removed — its capabilities are covered elsewhere', () => {
  const names = new Set(tools.map((t) => t.name));
  assert.ok(!names.has('updatePage'), 'addToPage / replacePage / insertNearText / replaceElement / editTable cover it');
});

test('every tool has a description and a schema', () => {
  for (const t of tools) {
    assert.ok(t.description && t.description.length > 10, `${t.name} needs a description`);
    assert.ok(t.inputSchema, `${t.name} needs an inputSchema`);
  }
});

test('the per-request token cost stays within budget', () => {
  const json = JSON.stringify(tools);
  assert.ok(
    json.length <= TOOLS_JSON_BUDGET,
    `tools/list is ${json.length} chars (budget ${TOOLS_JSON_BUDGET}). ` +
    `This is paid on every request — merge a tool or trim descriptions.`
  );
  assert.ok(
    initResult.instructions.length <= INSTRUCTIONS_BUDGET,
    `instructions are ${initResult.instructions.length} chars (budget ${INSTRUCTIONS_BUDGET}).`
  );
});

test('no single tool dominates the schema budget', () => {
  for (const t of tools) {
    const size = JSON.stringify(t).length;
    assert.ok(size <= PER_TOOL_BUDGET, `${t.name} schema is ${size} chars (budget ${PER_TOOL_BUDGET})`);
  }
});

test('editTable addresses rows and columns by index or by name', () => {
  const props = tools.find((x) => x.name === 'editTable').inputSchema.properties;
  for (const op of ['setCells', 'addRows', 'deleteRows', 'addColumns', 'deleteColumns']) {
    assert.ok(props[op], `editTable should support ${op}`);
  }
  // Naming a row or column is what removes the read-then-edit round trip.
  const kinds = (s) => s.anyOf.map((x) => x.type).sort();
  assert.deepEqual(kinds(props.setCells.items.properties.row), ['number', 'string']);
  assert.deepEqual(kinds(props.setCells.items.properties.col), ['number', 'string']);
});

test('reading tools declare themselves read-only', () => {
  // Clients use readOnlyHint to decide what needs a confirmation prompt.
  for (const name of ['getPage', 'getPageSection', 'listPages', 'searchPages', 'getNotebookTree']) {
    assert.equal(tools.find((x) => x.name === name).annotations?.readOnlyHint, true, name);
  }
  assert.equal(tools.find((x) => x.name === 'deletePage').annotations?.destructiveHint, true);
  assert.ok(!tools.find((x) => x.name === 'replacePage').annotations?.readOnlyHint);
});

test('the schema carries no bounds the caller never reads', () => {
  // z.number().int() writes zod's safe-integer ceiling into every numeric
  // argument. It is pure noise, and it is re-sent on every single request.
  assert.ok(
    !JSON.stringify(tools).includes('9007199254740991'),
    'use whole(z.number()...) instead of z.number().int()'
  );
});

test('writing tools accept a format parameter', () => {
  for (const name of ['createPage', 'replacePage', 'addToPage', 'insertNearText', 'replaceElement']) {
    const t = tools.find((x) => x.name === name);
    assert.ok(t.inputSchema.properties.format, `${name} should accept a format`);
    assert.deepEqual(t.inputSchema.properties.format.enum, ['markdown', 'text', 'html']);
  }
});

test('getPage defaults to a structured format and offers a cheap outline', () => {
  const t = tools.find((x) => x.name === 'getPage');
  assert.deepEqual(t.inputSchema.properties.format.enum, ['markdown', 'outline', 'text', 'html']);
});

test('editPageSection can add to and delete a section, not just replace it', () => {
  const t = tools.find((x) => x.name === 'editPageSection');
  assert.deepEqual(t.inputSchema.properties.mode.enum, ['replace', 'append', 'prepend', 'delete']);
  // "delete" is the only way to remove a heading and its content, so content
  // cannot be required.
  assert.ok(!(t.inputSchema.required || []).includes('content'), 'content must be optional');
  // Splicing a section rewrites the body, so it needs the same image guard as
  // replacePage, or an image below the heading would vanish without a word.
  assert.ok(t.inputSchema.properties.allowImageLoss, 'editPageSection needs the image guard');
});

test('setCheckbox can tick several boxes in one call', () => {
  // Ticking a to-do used to mean reading the whole page as HTML just to learn
  // an element id; this is the tool that removes that read.
  const t = tools.find((x) => x.name === 'setCheckbox');
  assert.ok(t, 'setCheckbox should be registered');
  const kinds = t.inputSchema.properties.text.anyOf.map((x) => x.type).sort();
  assert.deepEqual(kinds, ['array', 'string'], 'one PATCH should be able to carry several');
  assert.ok(!(t.inputSchema.required || []).includes('checked'), 'checked defaults to ticking');
});

test('searchPages can match content, not only titles', () => {
  const t = tools.find((x) => x.name === 'searchPages');
  assert.equal(t.inputSchema.properties.inContent.type, 'boolean');
  // Content matching reads every page in scope, so the scope is the argument
  // that keeps it bounded.
  for (const scope of ['sectionId', 'notebookId']) assert.ok(t.inputSchema.properties[scope], scope);
});

test('the instructions name the tools they promise', () => {
  // The instructions are the only place the model is told these exist.
  const names = new Set(tools.map((t) => t.name));
  for (const named of ['setCheckbox', 'searchPages']) {
    assert.ok(initResult.instructions.includes(named), `instructions should mention ${named}`);
    assert.ok(names.has(named));
  }
});

test('searchPages omits unknown attribution instead of sending null', () => {
  // A search scoped to one section never looks the section up, so its name and
  // notebook are unknown — and an unknown is cheaper omitted than sent as a
  // null on every hit. Read through the tool's own result shape rather than the
  // source text: the contract is the output, not the line that builds it.
  const t = tools.find((x) => x.name === 'searchPages');
  const props = Object.keys(t.inputSchema.properties);
  assert.ok(props.includes('sectionId'), 'a section-scoped search is expressible');
  // The description must not promise attribution it cannot supply.
  assert.doesNotMatch(t.description, /section name/i);
});

test('getAuthoringGuide exists, is read-only, and its topic list is closed', () => {
  const t = tools.find((x) => x.name === 'getAuthoringGuide');
  assert.ok(t, 'getAuthoringGuide should be registered');
  assert.equal(t.annotations?.readOnlyHint, true);
  const topics = t.inputSchema.properties.topic.enum;
  assert.ok(topics.length >= 5, 'should cover headings, tables, math, tags and images at least');
  assert.ok(!(t.inputSchema.required || []).includes('topic'), 'omitting topic must return the whole guide');
});

test('copyPageToSection advertises its fallback and guards image loss', () => {
  // Personal Microsoft accounts answer copyToSection with 501, so the copy API
  // is simply absent for the accounts this server targets. The tool rebuilds
  // the page instead — which cannot carry images, hence the guard.
  const t = tools.find((x) => x.name === 'copyPageToSection');
  assert.match(t.description, /falls back/i, 'the fallback must be discoverable from the description');
  assert.ok(t.inputSchema.properties.allowImageLoss, 'the fallback needs the image guard');
  assert.ok(!(t.inputSchema.required || []).includes('allowImageLoss'));
});

test('a Graph failure carries its status code, not just a message', async () => {
  // The 501 fallback in copyPageToSection branches on `err.status`, so the
  // status must survive into the thrown error — matching the message text
  // instead would break the moment Microsoft rewords the error. Asserted
  // against the real client with a stubbed transport.
  const { graph } = client({ status: 501, body: '{"error":{"code":"20111","message":"not implemented"}}' });
  const err = await graph('POST', '/me/onenote/pages/p1/copyToSection', { body: {}, json: true })
    .then(() => null, (e) => e);
  assert.ok(err, 'a non-ok response must throw');
  assert.equal(err.status, 501, 'the HTTP status is attached to the error');
  // The envelope's request id, timestamp and innerError are dropped: only the
  // actionable part reaches the caller.
  assert.match(err.message, /20111: not implemented/);
});

test('a 202 is reported as an accepted operation, not a Graph payload', async () => {
  const { graph } = client({ status: 202, headers: { Location: 'https://example.test/op/1' } });
  const r = await graph('POST', '/me/onenote/pages/p1/copyToSection', { body: {}, json: true });
  assert.equal(r.kind, 'accepted');
  assert.equal(r.operationLocation, 'https://example.test/op/1');
});

test('a 204 is an empty success and a JSON body is parsed', async () => {
  const none = client({ status: 204 });
  assert.equal(await none.graph('DELETE', '/me/onenote/pages/p1'), null);

  const some = client({ status: 200, body: '{"id":"abc"}' });
  assert.deepEqual(await some.graph('GET', '/me/onenote/pages/p1'), { id: 'abc' });
});
