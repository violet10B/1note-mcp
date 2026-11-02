import { test } from 'node:test';
import assert from 'node:assert/strict';
import { validateMathMLString, checkMathInContent, findMathBlocks } from '../src/mathml.js';

const NS = 'http://www.w3.org/1998/Math/MathML';
const good = `<math xmlns="${NS}"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>`;

test('valid MathML passes clean', async () => {
  const r = await validateMathMLString(good);
  assert.equal(r.valid, true);
  assert.deepEqual(r.errors, []);
  assert.deepEqual(r.warnings, []);
});

test('wrong child count is a hard error', async () => {
  const r = await validateMathMLString(`<math xmlns="${NS}"><mfrac><mi>a</mi></mfrac></math>`);
  assert.equal(r.valid, false);
  assert.match(r.errors.join(' '), /<mfrac> needs exactly 2 child elements but has 1/);
});

test('non-well-formed XML is a hard error, not a crash', async () => {
  // jsdom either throws or emits a <parsererror>; both must land as a hard error.
  const r = await validateMathMLString(`<math xmlns="${NS}"><mi>a</mo></math>`);
  assert.equal(r.valid, false);
  assert.match(r.errors.join(' '), /parse failed|well-formed/i);
});

test('a missing namespace warns but still writes', async () => {
  const r = await validateMathMLString('<math><mi>x</mi></math>');
  assert.equal(r.valid, true);
  assert.match(r.warnings.join(' '), /xmlns/);
});

test('content with no math is a no-op', async () => {
  const r = await checkMathInContent('<p>just prose</p>');
  assert.equal(r.hardError, null);
  assert.deepEqual(r.warnings, []);
});

test('every math block in a page is checked, not just the first', async () => {
  const bad = `<math xmlns="${NS}"><msub><mi>x</mi></msub></math>`;
  const r = await checkMathInContent(`<p>${good}</p><p>${bad}</p>`);
  assert.ok(r.hardError, 'a bad block anywhere must block the write');
  assert.match(r.hardError, /<msub> needs exactly 2/);
  assert.equal(findMathBlocks(`<p>${good}</p><p>${bad}</p>`).length, 2);
});
