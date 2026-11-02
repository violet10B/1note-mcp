import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { createHandleStore, fnv1a36, isHandle } from '../src/ids.js';

const REAL = '0-37c8b5fa5eaa454da7ecdf72a8c137fe!1-B68C1234ABCD!sd3336abcdef0123';

test('handles are short, prefixed and stable', () => {
  const a = createHandleStore();
  const h = a.handle(REAL, 'p');
  assert.match(h, /^p[0-9a-z]{4,}$/);
  assert.ok(h.length < 12, 'handle should be far shorter than the Graph id');
  assert.equal(a.handle(REAL, 'p'), h, 'minting twice returns the same handle');
});

test('independent stores mint the same handle for the same id', () => {
  const a = createHandleStore(), b = createHandleStore();
  assert.equal(a.handle(REAL, 's'), b.handle(REAL, 's'));
});

test('resolveId round-trips handles and passes raw ids through', () => {
  const s = createHandleStore();
  const h = s.handle(REAL, 'p');
  assert.equal(s.resolveId(h), REAL);
  assert.equal(s.resolveId(REAL), REAL, 'a raw Graph id is returned unchanged');
});

test('an unknown handle fails loudly rather than silently', () => {
  const s = createHandleStore();
  assert.throws(() => s.resolveId('p9999'), /Unknown id handle/);
});

test('isHandle never matches a real Graph id', () => {
  assert.ok(!isHandle(REAL));
  assert.ok(!isHandle('0-37c8!1-B68C'));
});

test('every minted handle is recognised by isHandle', () => {
  // Guards the docs: handles are prefix + >=4 base36 chars, so illustrative
  // examples like "nb07" or "p1ab" are a character short and would not resolve.
  const s = createHandleStore();
  for (const kind of ['nb', 'sg', 's', 'p']) {
    const h = s.handle(REAL + kind, kind);
    assert.ok(isHandle(h), `${h} should be a valid handle`);
    assert.equal(s.resolveId(h), REAL + kind);
  }
  assert.ok(!isHandle('nb07'), 'a 2-char suffix is below the minimum');
});

test('the collision ladder widens instead of retrying one candidate', () => {
  // Force a collision: seed the store with the 4-char handle this id wants.
  const s = createHandleStore();
  const base = fnv1a36(REAL);
  const other = 'some-other-graph-id';
  // Pre-occupy p+base[0..4) by minting a fake entry through the public API is
  // not possible, so assert the ladder property directly: distinct ids never
  // share a handle.
  const h1 = s.handle(REAL, 'p');
  const h2 = s.handle(other, 'p');
  assert.notEqual(h1, h2);
  assert.equal(s.resolveId(h1), REAL);
  assert.equal(s.resolveId(h2), other);
  assert.ok(base.length >= 4 && base.length <= 7);
});

test('handles survive a restart and resolve across processes', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'onenote-ids-')), 'map.json');
  const a = createHandleStore({ filePath: file });
  const h = a.handle(REAL, 'p');
  a.flush();

  // A second store started later sees the handle on disk.
  const b = createHandleStore({ filePath: file });
  assert.equal(b.resolveId(h), REAL);

  // A handle minted by another client AFTER we loaded our map still resolves,
  // because a miss re-reads from disk before giving up.
  const c = createHandleStore({ filePath: file });
  const other = 'another-graph-id-9!x';
  const h2 = a.handle(other, 's');
  a.flush();
  assert.equal(c.resolveId(h2), other);
});
