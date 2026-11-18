/*
 * SCORM player host page logic.
 *
 * 1. Registers the service worker (sw.js) that serves package files.
 * 2. Tells the SharePoint web part (parent iframe or window.opener) it is ready.
 * 3. On "scorm:load": extracts the zip into the Cache API (once per package),
 *    exposes the SCORM API on this window and launches the SCO in an iframe.
 * 4. Forwards progress (CMI data), navigation requests and API logs to the web part.
 */
(function () {
  'use strict';

  var HOST_VERSION = '1.0.0';
  var config = window.SCORM_PLAYER_CONFIG || { ALLOWED_PARENT_ORIGINS: [] };
  var CONTENT_BASE = new URL('./content/', window.location.href);
  var CACHE_PREFIX = 'scorm-pkg-';

  var parentWin = window.opener || (window.parent !== window ? window.parent : null);
  var parentOrigin = null;     // set from the first valid load message
  var runtime = null;
  var currentScoId = null;
  var frameEl = null;
  var statusEl = document.getElementById('status');
  var stageEl = document.getElementById('stage');

  var MIME = {
    html: 'text/html', htm: 'text/html', xhtml: 'application/xhtml+xml', js: 'text/javascript', mjs: 'text/javascript',
    css: 'text/css', json: 'application/json', xml: 'application/xml', xsd: 'application/xml', txt: 'text/plain',
    svg: 'image/svg+xml', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp',
    ico: 'image/x-icon', bmp: 'image/bmp', mp3: 'audio/mpeg', wav: 'audio/wav', ogg: 'audio/ogg', m4a: 'audio/mp4',
    mp4: 'video/mp4', webm: 'video/webm', ogv: 'video/ogg', vtt: 'text/vtt', woff: 'font/woff', woff2: 'font/woff2',
    ttf: 'font/ttf', otf: 'font/otf', eot: 'application/vnd.ms-fontobject', pdf: 'application/pdf',
    swf: 'application/x-shockwave-flash', wasm: 'application/wasm'
  };

  function setStatus(text, isError) {
    statusEl.textContent = text || '';
    statusEl.className = text ? (isError ? 'status error' : 'status') : 'status hidden';
  }

  function post(msg) {
    if (parentWin) { parentWin.postMessage(msg, parentOrigin || '*'); }
  }

  function log(level, message) {
    if (level === 'error') { console.error('[SCORM]', message); }
    if (parentOrigin) { post({ type: 'scorm:log', level: level, message: message }); }
  }

  function originAllowed(origin) {
    return (config.ALLOWED_PARENT_ORIGINS || []).some(function (rule) {
      return rule.charAt(0) === '^' ? new RegExp(rule, 'i').test(origin) : rule.toLowerCase() === origin.toLowerCase();
    });
  }

  function encodePath(p) { return p.split('/').map(encodeURIComponent).join('/'); }

  // --------------------------------------------------------- service worker
  var swReady = ('serviceWorker' in navigator)
    ? navigator.serviceWorker.register('sw.js', { scope: './' }).then(function () { return navigator.serviceWorker.ready; })
    : Promise.reject(new Error('Service workers are not supported by this browser.'));

  swReady.catch(function (e) {
    setStatus('This browser blocked the SCORM player (' + e.message + '). Try "Open in new window" or allow third-party storage for the player site.', true);
  });

  // ------------------------------------------------------- package extraction
  function extractPackage(packageId, zipBuffer) {
    var cacheName = CACHE_PREFIX + packageId;
    var base = CONTENT_BASE.href + packageId + '/';
    return caches.open(cacheName).then(function (cache) {
      return cache.match(base + '__complete').then(function (done) {
        if (done) { return { fromCache: true, fileCount: -1 }; }
        setStatus('Unpacking course…');
        return JSZip.loadAsync(zipBuffer).then(function (zip) {
          var paths = Object.keys(zip.files).filter(function (p) { return !zip.files[p].dir; });
          var index = {}, i = 0;
          paths.forEach(function (p) { index[p.toLowerCase()] = p; });

          function next() {
            var batch = paths.slice(i, i + 25);
            i += batch.length;
            if (!batch.length) { return Promise.resolve(); }
            setStatus('Unpacking course… ' + Math.round((i / paths.length) * 100) + '%');
            return Promise.all(batch.map(function (p) {
              return zip.files[p].async('arraybuffer').then(function (buf) {
                var ext = (p.split('.').pop() || '').toLowerCase();
                return cache.put(base + encodePath(p), new Response(buf, {
                  headers: { 'Content-Type': MIME[ext] || 'application/octet-stream', 'Content-Length': String(buf.byteLength) }
                }));
              });
            })).then(next);
          }

          return next().then(function () {
            return Promise.all([
              cache.put(base + '__index.json', new Response(JSON.stringify(index), { headers: { 'Content-Type': 'application/json' } })),
              cache.put(base + '__complete', new Response(String(Date.now())))
            ]);
          }).then(function () { return { fromCache: false, fileCount: paths.length }; });
        });
      });
    });
  }

  function cleanupOldPackages() {
    var maxAge = (config.CACHE_MAX_AGE_DAYS || 14) * 86400000;
    caches.keys().then(function (keys) {
      keys.filter(function (k) { return k.indexOf(CACHE_PREFIX) === 0; }).forEach(function (k) {
        var pkg = k.substring(CACHE_PREFIX.length);
        caches.open(k).then(function (c) { return c.match(CONTENT_BASE.href + pkg + '/__complete'); }).then(function (r) {
          return r ? r.text() : '0';
        }).then(function (ts) {
          if (Date.now() - (+ts || 0) > maxAge) { caches.delete(k); }
        });
      });
    });
  }

  // ------------------------------------------------------------- SCO launch
  function unloadCurrent() {
    if (runtime) { runtime.abandon(); }
    runtime = null;
    delete window.API;
    delete window.API_1484_11;
    if (frameEl) {
      frameEl.src = 'about:blank';
      frameEl.parentNode.removeChild(frameEl);
      frameEl = null;
    }
  }

  function launch(msg) {
    unloadCurrent();
    currentScoId = msg.scoId;
    setStatus('Preparing course…');

    swReady.then(function () {
      return extractPackage(msg.packageId, msg.zip);
    }).then(function (result) {
      var scoId = msg.scoId;
      runtime = new window.ScormRuntime({
        version: msg.version,
        cmi: msg.cmi || {},
        learner: msg.learner,
        masteryScore: msg.masteryScore,
        scaledPassingScore: msg.scaledPassingScore,
        completionThreshold: msg.completionThreshold,
        launchData: msg.launchData,
        onProgress: function (cmi, reason) { post({ type: 'scorm:progress', scoId: scoId, reason: reason, cmi: cmi }); },
        onNav: function (request) { post({ type: 'scorm:nav', scoId: scoId, request: request }); },
        onLog: log
      });
      var api = runtime.createApi();
      if (msg.version === '1.2') { window.API = api; } else { window.API_1484_11 = api; }

      var src = new URL(msg.launchPath, CONTENT_BASE.href + msg.packageId + '/').href;
      frameEl = document.createElement('iframe');
      frameEl.className = 'sco';
      frameEl.title = msg.title || 'Course content';
      frameEl.setAttribute('allow', 'autoplay; fullscreen; clipboard-write');
      frameEl.addEventListener('load', function () { setStatus(''); });
      frameEl.src = src;
      stageEl.appendChild(frameEl);

      log('info', 'Launched ' + msg.launchPath + ' (SCORM ' + msg.version + ', ' + (result.fromCache ? 'from cache' : result.fileCount + ' files unpacked') + ')');
      post({ type: 'scorm:loaded', scoId: scoId, fromCache: result.fromCache, fileCount: result.fileCount });
    }).catch(function (e) {
      var text = 'Could not start the course: ' + (e && e.message ? e.message : e);
      setStatus(text, true);
      post({ type: 'scorm:error', message: text });
    });
  }

  // ---------------------------------------------------------------- messages
  window.addEventListener('message', function (e) {
    var msg = e.data;
    if (!msg || typeof msg.type !== 'string' || msg.type.indexOf('scorm:') !== 0) { return; }
    if (e.source !== parentWin) { return; }
    if (!originAllowed(e.origin)) {
      console.warn('[SCORM] Rejected message from ' + e.origin + '. Add it to ALLOWED_PARENT_ORIGINS in config.js.');
      return;
    }
    parentOrigin = e.origin;

    switch (msg.type) {
      case 'scorm:load': launch(msg); break;
      case 'scorm:unload': unloadCurrent(); setStatus('Select a lesson to start.'); break;
      case 'scorm:ping': post({ type: 'scorm:ready', hostVersion: HOST_VERSION }); break;
    }
  });

  window.addEventListener('pagehide', function () { if (runtime) { runtime.abandon(); } });

  if (!parentWin) {
    setStatus('SCORM player host is running. Add the "SCORM Player" web part to a SharePoint page and point it to this URL.');
  } else {
    setStatus('Waiting for course…');
    swReady.then(function () { post({ type: 'scorm:ready', hostVersion: HOST_VERSION }); });
  }
  cleanupOldPackages();
})();
