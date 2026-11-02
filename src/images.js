// URL guard for image fetching. Kept in src/ rather than in the server so it can
// be tested without starting a server.

/**
 * Reject URLs that would make this process reach something it should not.
 *
 * The url argument is model-supplied, so a prompt-injected page could aim the
 * fetch at a loopback or link-local address (cloud metadata lives at
 * 169.254.169.254) and use this server as a proxy into the local network.
 * Only http(s) is allowed, and private ranges are refused by host.
 */
export function assertFetchableUrl(raw) {
  let u;
  try {
    u = new URL(raw);
  } catch {
    throw new Error(`Not a valid URL: ${raw}`);
  }
  if (u.protocol !== 'http:' && u.protocol !== 'https:') {
    throw new Error(`Only http and https URLs are supported (got "${u.protocol}").`);
  }
  const host = u.hostname.replace(/^\[|\]$/g, '').toLowerCase();
  const isIpv4 = /^\d{1,3}(\.\d{1,3}){3}$/.test(host);
  const isIpv6 = host.includes(':');
  const blocked = host === 'localhost'
    || host === '0.0.0.0'
    || host.endsWith('.localhost')
    || host.endsWith('.internal')
    || (isIpv4 && (
      host.startsWith('127.')
      || host.startsWith('10.')
      || host.startsWith('192.168.')
      || host.startsWith('169.254.')
      || /^172\.(1[6-9]|2\d|3[01])\./.test(host)
    ))
    // ::1 and the unique-local / link-local IPv6 ranges
    || (isIpv6 && (host === '::1' || /^(fc|fd|fe80)/.test(host.replace(/:/g, ''))));
  if (blocked) {
    throw new Error(
      `Refusing to fetch "${u.hostname}": loopback, private and link-local addresses are not allowed. `
      + 'Download the image and pass filePath instead.'
    );
  }
  return u;
}
