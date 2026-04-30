/**
 * Netlify Serverless Function — Helix Attend API Proxy
 * 
 * All calls from the PWA go to /.netlify/functions/api
 * This function forwards them to Google Apps Script server-side,
 * completely bypassing browser CORS restrictions.
 * 
 * Browser → Netlify Function (same origin, no CORS) → Apps Script (server-to-server, no CORS)
 */

const https = require('https');
const http  = require('http');

exports.handler = async function(event) {
  // Read the Apps Script URL from Netlify environment variable
  // Set this in: Netlify Dashboard → Site → Environment Variables → APPS_SCRIPT_URL
  const SCRIPT_URL = process.env.APPS_SCRIPT_URL;

  if (!SCRIPT_URL) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ok: false, error: 'APPS_SCRIPT_URL environment variable not set in Netlify' })
    };
  }

  // Parse incoming request from PWA
  let body = {};
  try {
    if (event.body) body = JSON.parse(event.body);
  } catch(e) {}

  const action  = event.queryStringParameters?.action || body.action || '';
  const payload = body.payload || body;

  // Build the Apps Script URL
  const scriptUrl = new URL(SCRIPT_URL);
  scriptUrl.searchParams.set('action', action);

  // Forward to Apps Script as server-side POST (no CORS issues server-to-server)
  try {
    const result = await fetchUrl(scriptUrl.toString(), {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(payload)
    });

    return {
      statusCode: 200,
      headers: {
        'Content-Type':                'application/json',
        'Access-Control-Allow-Origin': '*'   // Allow PWA to read this response
      },
      body: result
    };
  } catch(err) {
    return {
      statusCode: 502,
      headers: {
        'Content-Type':                'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({ ok: false, error: 'Proxy error: ' + err.message })
    };
  }
};

/** Simple HTTP/HTTPS fetch for Node.js (no node-fetch needed) */
function fetchUrl(url, options = {}) {
  return new Promise((resolve, reject) => {
    const parsed  = new URL(url);
    const lib     = parsed.protocol === 'https:' ? https : http;
    const bodyStr = options.body || '';

    const req = lib.request({
      hostname: parsed.hostname,
      path:     parsed.pathname + parsed.search,
      method:   options.method || 'GET',
      headers:  {
        ...options.headers,
        'Content-Length': Buffer.byteLength(bodyStr)
      }
    }, (res) => {
      // Follow redirects (Apps Script always redirects)
      if ((res.statusCode === 301 || res.statusCode === 302 || res.statusCode === 307) && res.headers.location) {
        return fetchUrl(res.headers.location, options).then(resolve).catch(reject);
      }

      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data));
    });

    req.on('error', reject);
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}
