/**
 * Netlify Serverless Function — Helix Attend API Proxy
 *
 * Routes: Browser → /api (Netlify, same origin) → Apps Script (server-side, no CORS)
 *
 * SETUP: Set APPS_SCRIPT_URL in Netlify → Site configuration → Environment variables
 */

const https = require('https');

exports.handler = async function(event) {
  const SCRIPT_URL = process.env.APPS_SCRIPT_URL;

  // ── Environment variable not set ──
  if (!SCRIPT_URL) {
    return respond(500, {
      ok: false,
      error: 'APPS_SCRIPT_URL not set. Go to Netlify → Site configuration → Environment variables and add it.'
    });
  }

  // ── Parse action + body from incoming PWA request ──
  const action = (event.queryStringParameters && event.queryStringParameters.action) || '';
  let body = {};
  try { if (event.body) body = JSON.parse(event.body); } catch(e) {}

  // ── Build Apps Script URL with action param ──
  const url = SCRIPT_URL.replace(/\/$/, '') + '?action=' + encodeURIComponent(action);

  // ── Call Apps Script ──
  try {
    const raw = await httpsPost(url, body);

    // Validate it looks like JSON before returning
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch(e) {
      // Apps Script returned HTML or error page — surface it clearly
      const preview = raw.substring(0, 200).replace(/\n/g, ' ');
      return respond(502, {
        ok: false,
        error: 'Apps Script returned non-JSON. Check your deployment URL and re-deploy the Apps Script. Preview: ' + preview
      });
    }

    return respond(200, parsed);

  } catch(err) {
    return respond(502, { ok: false, error: 'Proxy error: ' + err.message });
  }
};

/** POST JSON to a URL, following up to 5 redirects, always with GET after first redirect */
function httpsPost(url, body, redirectCount) {
  redirectCount = redirectCount || 0;
  if (redirectCount > 5) return Promise.reject(new Error('Too many redirects'));

  const bodyStr = JSON.stringify(body);
  const parsed  = new URL(url);

  return new Promise((resolve, reject) => {
    const options = {
      hostname: parsed.hostname,
      path:     parsed.pathname + parsed.search,
      method:   redirectCount === 0 ? 'POST' : 'GET',
      headers:  redirectCount === 0
        ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(bodyStr) }
        : {}
    };

    const req = https.request(options, (res) => {
      // Follow redirect — Apps Script always redirects to googleusercontent.com
      if ([301,302,303,307,308].includes(res.statusCode) && res.headers.location) {
        const next = res.headers.location.startsWith('http')
          ? res.headers.location
          : 'https://' + parsed.hostname + res.headers.location;
        // Consume response body to free socket
        res.resume();
        return httpsPost(next, body, redirectCount + 1).then(resolve).catch(reject);
      }

      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data));
      res.on('error', reject);
    });

    req.on('error', reject);

    // Only send body on first (POST) request
    if (redirectCount === 0) req.write(bodyStr);
    req.end();
  });
}

/** Standard JSON response helper */
function respond(statusCode, body) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    },
    body: typeof body === 'string' ? body : JSON.stringify(body)
  };
}
