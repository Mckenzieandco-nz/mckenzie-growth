/**
 * /api/qaPing — diagnostic endpoint. Calls Anthropic from inside Azure using
 * Node's built-in https module only (no SDK dependency). Returns key metadata
 * + raw HTTPS result.
 *
 * No dependencies on npm packages — runs without node_modules.
 */

const https = require('https');

function rawAnthropicCall(apiKey) {
  return new Promise((resolve) => {
    const body = JSON.stringify({
      model: 'claude-haiku-4-5',
      max_tokens: 50,
      messages: [{ role: 'user', content: 'hi' }],
    });
    const req = https.request({
      hostname: 'api.anthropic.com',
      path: '/v1/messages',
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body),
      },
      timeout: 30000,
    }, (res) => {
      let chunks = '';
      res.on('data', (c) => { chunks += c; });
      res.on('end', () => resolve({ status: res.statusCode, body: chunks, headers: res.headers }));
    });
    req.on('error', (err) => resolve({ status: 0, body: '', error: err.message }));
    req.on('timeout', () => { req.destroy(); resolve({ status: 0, body: '', error: 'timeout' }); });
    req.write(body);
    req.end();
  });
}

module.exports = async function (context, req) {
  const rawKey = process.env.ANTHROPIC_API_KEY;
  const cleanedKey = (rawKey || '').trim();

  const keyInfo = rawKey ? {
    rawLength: rawKey.length,
    trimmedLength: cleanedKey.length,
    hadWhitespace: rawKey.length !== cleanedKey.length,
    startsCorrectly: cleanedKey.startsWith('sk-ant-'),
    firstChars: cleanedKey.slice(0, 12),
    lastChars: cleanedKey.slice(-4),
  } : null;

  if (!cleanedKey) {
    context.res = {
      status: 503,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: false, stage: 'env_var', error: 'ANTHROPIC_API_KEY is empty/missing on the server.', keyInfo },
    };
    return;
  }

  if (!cleanedKey.startsWith('sk-ant-')) {
    context.res = {
      status: 502,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: false,
        stage: 'env_var',
        error: `ANTHROPIC_API_KEY value does not start with "sk-ant-".`,
        keyInfo,
      },
    };
    return;
  }

  const rawResult = await rawAnthropicCall(cleanedKey);
  let rawParsed = null;
  try { rawParsed = JSON.parse(rawResult.body); } catch { /* leave null */ }

  const ok = rawResult.status === 200;
  const out = {
    ok,
    keyInfo,
    rawHttps: {
      ok,
      status: rawResult.status,
      requestId: rawResult.headers && rawResult.headers['request-id'],
      reply: rawParsed && rawParsed.content ? (rawParsed.content[0] || {}).text : null,
      bodySnippet: rawResult.body ? rawResult.body.slice(0, 300) : null,
      error: rawResult.error || null,
    },
    sdk: { ok: ok, skipped: true, reason: 'Using built-in https module only; no SDK installed' },
  };
  context.log.info('qaPing result:', JSON.stringify({ ok: out.ok, status: rawResult.status, requestId: out.rawHttps.requestId }));
  context.res = {
    status: ok ? 200 : 502,
    headers: { 'Content-Type': 'application/json' },
    body: out,
  };
};
