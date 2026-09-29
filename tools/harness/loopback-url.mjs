// Discovery files choose a local session, never an arbitrary network destination.
export function loopbackUrl(port, pathname = '/') {
  const number = Number(port);
  if (!Number.isInteger(number) || number < 1 || number > 65535) {
    throw new Error('Invalid loopback port');
  }
  const url = new URL('http://127.0.0.1/');
  url.port = String(number);
  url.pathname = pathname;
  return url.href;
}

export function localHttpUrl(value) {
  const url = new URL(value);
  if (url.protocol !== 'http:' || url.hostname !== '127.0.0.1' || url.username || url.password) {
    throw new Error('The harness only connects to HTTP on 127.0.0.1');
  }
  return url.href;
}

export function localRoute(base, route) {
  return localHttpUrl(new URL(route, localHttpUrl(base)).href);
}
