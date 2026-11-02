// The url argument to addImageToPage is model-supplied, so it is reachable from
// a prompt-injected page. These pin the guard that stops this server being used
// as a proxy into the local network.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { assertFetchableUrl } from '../src/images.js';

test('ordinary http(s) image URLs are allowed', () => {
  for (const url of [
    'https://example.com/cat.png',
    'http://example.com/cat.png',
    'https://cdn.example.org/a/b/c.jpg?w=100',
  ]) {
    assert.ok(assertFetchableUrl(url), url);
  }
});

test('loopback and private addresses are refused', () => {
  for (const url of [
    'http://localhost/x.png',
    'http://127.0.0.1/x.png',
    'http://127.1.2.3/x.png',
    'http://0.0.0.0/x.png',
    'http://10.0.0.5/x.png',
    'http://192.168.1.1/x.png',
    'http://172.16.0.1/x.png',
    'http://172.31.255.255/x.png',
    'http://169.254.169.254/latest/meta-data/', // cloud metadata
    'http://box.localhost/x.png',
    'http://thing.internal/x.png',
  ]) {
    assert.throws(() => assertFetchableUrl(url), /Refusing to fetch/, url);
  }
});

test('a public address just outside the private range is still allowed', () => {
  assert.ok(assertFetchableUrl('http://172.32.0.1/x.png'));
});

test('non-http schemes are refused', () => {
  for (const url of ['file:///etc/passwd', 'ftp://example.com/x.png', 'data:image/png;base64,AAAA']) {
    assert.throws(() => assertFetchableUrl(url), /Only http and https/, url);
  }
});

test('a malformed URL is refused rather than reaching fetch', () => {
  assert.throws(() => assertFetchableUrl('not a url'), /valid URL/);
});
