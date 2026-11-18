/*
 * Service worker of the SCORM player host.
 *
 * player.js extracts each SCORM zip into the Cache API under
 *   <host>/content/<packageId>/<path inside zip>
 * and this worker answers every request below /content/ from that cache, so the
 * SCO behaves exactly as if it were deployed on a web server (relative links,
 * dynamic script loading, XHR/fetch of JSON, media, ...).
 */
'use strict';

var CONTENT_PREFIX = new URL('./content/', self.location).pathname;
var CACHE_PREFIX = 'scorm-pkg-';
var indexCache = {}; // packageId -> { lowercasePath: realPath }

self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });

self.addEventListener('fetch', function (event) {
  var url = new URL(event.request.url);
  if (url.origin !== self.location.origin || url.pathname.indexOf(CONTENT_PREFIX) !== 0) { return; }
  if (event.request.method !== 'GET' && event.request.method !== 'HEAD') { return; }
  event.respondWith(serve(event.request, url));
});

function serve(request, url) {
  var rest = url.pathname.substring(CONTENT_PREFIX.length); // "<pkg>/<path>"
  var slash = rest.indexOf('/');
  var pkg = slash > 0 ? rest.substring(0, slash) : rest;
  var key = url.origin + url.pathname;

  return caches.open(CACHE_PREFIX + pkg).then(function (cache) {
    return cache.match(key).then(function (hit) {
      if (hit) { return hit; }
      // Fallback 1: directory -> index.html
      if (/\/$/.test(key)) { return cache.match(key + 'index.html'); }
      // Fallback 2: case-insensitive lookup (content authored on Windows servers)
      return lookupIndex(cache, pkg, url).then(function (real) { return real ? cache.match(real) : undefined; });
    });
  }).then(function (response) {
    if (!response) {
      return new Response('Not found in SCORM package: ' + url.pathname, { status: 404, headers: { 'Content-Type': 'text/plain' } });
    }
    var range = request.headers.get('Range');
    return range ? rangeResponse(response, range) : response;
  });
}

function lookupIndex(cache, pkg, url) {
  var load = indexCache[pkg]
    ? Promise.resolve(indexCache[pkg])
    : cache.match(url.origin + CONTENT_PREFIX + pkg + '/__index.json')
      .then(function (r) { return r ? r.json() : {}; })
      .then(function (idx) { indexCache[pkg] = idx; return idx; });
  return load.then(function (idx) {
    var p = decodeURIComponent(url.pathname.substring(CONTENT_PREFIX.length + pkg.length + 1)).toLowerCase();
    return idx[p] ? url.origin + CONTENT_PREFIX + pkg + '/' + encodePath(idx[p]) : undefined;
  });
}

function encodePath(p) { return p.split('/').map(encodeURIComponent).join('/'); }

// Video/audio seeking needs byte-range support.
function rangeResponse(response, rangeHeader) {
  return response.blob().then(function (blob) {
    var m = /bytes=(\d*)-(\d*)/.exec(rangeHeader);
    if (!m) { return new Response(blob, { status: 200, headers: response.headers }); }
    var size = blob.size;
    var start = m[1] === '' ? Math.max(0, size - (+m[2])) : +m[1];
    var end = m[1] !== '' && m[2] !== '' ? Math.min(+m[2], size - 1) : size - 1;
    if (start >= size) {
      return new Response('', { status: 416, headers: { 'Content-Range': 'bytes */' + size } });
    }
    var headers = new Headers(response.headers);
    headers.set('Content-Range', 'bytes ' + start + '-' + end + '/' + size);
    headers.set('Content-Length', String(end - start + 1));
    headers.set('Accept-Ranges', 'bytes');
    return new Response(blob.slice(start, end + 1), { status: 206, headers: headers });
  });
}
