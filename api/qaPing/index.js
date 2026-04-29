/**
 * /api/qaPing — diagnostic endpoint. Runs a minimal Claude call from inside
 * Azure and returns the full result (or full error detail). Use this to
 * confirm the ANTHROPIC_API_KEY env var works from the Functions runtime.
 *
 * No body required. GET or POST works.
 */

const { Anthropic } = require('@anthropic-ai/sdk');
const https = require('https');

/* Raw HTTPS call to Anthropic — matches the PowerShell test exactly.
   Bypasses the SDK so we can isolate SDK-related issues from network issues. */
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

  // Step 1: env var diagnostics — never reveal the secret, just metadata
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
        error: `ANTHROPIC_API_KEY value does not start with "sk-ant-". First 12 chars: "${keyInfo.firstChars}". Re-paste the key in Azure Configuration.`,
        keyInfo,
      },
    };
    return;
  }

  // Step 2: try raw HTTPS first — bypasses the SDK, matches PowerShell exactly.
  const rawResult = await rawAnthropicCall(cleanedKey);
  let rawParsed = null;
  try { rawParsed = JSON.parse(rawResult.body); } catch { /* leave null */ }

  // Step 3: try the SDK — same call shape, but goes through @anthropic-ai/sdk.
  let sdkResult = { ok: false };
  try {
    const client = new Anthropic({ apiKey: cleanedKey });
    const resp = await client.messages.create({
      model: 'claude-haiku-4-5',
      max_tokens: 50,
      messages: [{ role: 'user', content: 'hi' }],
    });
    const text = (resp.content.find(b => b.type === 'text') || {}).text || '';
    sdkResult = { ok: true, model: resp.model, reply: text.trim() };
  } catch (err) {
    sdkResult = {
      ok: false,
      error: err.message || 'Unknown error',
      errorClass: err && err.constructor ? err.constructor.name : typeof err,
      status: err.status,
      requestId: err.request_id || (err.headers && (err.headers['request-id'] || err.headers['x-request-id'])),
      anthropicError: err.error,
    };
  }

  const overallOk = rawResult.status === 200 || sdkResult.ok;
  const sdkVersion = (() => {
    try { return require('@anthropic-ai/sdk/package.json').version; } catch { return 'unknown'; }
  })();

  const out = {
    ok: overallOk,
    keyInfo,
    sdkVersion,
    rawHttps: {
      ok: rawResult.status === 200,
      status: rawResult.status,
      requestId: rawResult.headers && rawResult.headers['request-id'],
      reply: rawParsed && rawParsed.content ? (rawParsed.content[0] || {}).text : null,
      bodySnippet: rawResult.body ? rawResult.body.slice(0, 300) : null,
      error: rawResult.error || null,
    },
    sdk: sdkResult,
  };
  context.log.info('qaPing result:', JSON.stringify(out));
  context.res = {
    status: overallOk ? 200 : 502,
    headers: { 'Content-Type': 'application/json' },
    body: out,
  };
};
