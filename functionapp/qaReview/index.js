/**
 * /api/qaReview — runs an AI QA review on an uploaded document via Claude.
 *
 * Uses ONLY Node's built-in https module — no @anthropic-ai/sdk or mammoth
 * dependency, so deployment doesn't depend on node_modules being present.
 *
 * Supported file types: PDF (sent natively as document block), plain text
 * (.txt, .md, .csv). DOCX is not supported in this build — convert to PDF.
 */

const https = require('https');

// ---------- Constants ----------

const ALLOWED_MODELS = new Set([
  'claude-opus-4-7',
  'claude-sonnet-4-6',
  'claude-haiku-4-5',
]);

const MAX_TOKENS_CAP = 32000;
const MAX_FILE_BYTES = 25 * 1024 * 1024;

const QA_CHECKS = {
  grammar:      { title: 'Grammar & Spelling',     desc: 'Typos, subject-verb agreement, punctuation' },
  clarity:      { title: 'Clarity & Readability',  desc: 'Plain English, sentence length, jargon' },
  technical:    { title: 'Technical Accuracy',     desc: 'Units, calculations, standards references' },
  brand:        { title: 'Brand & Style',          desc: 'McKenzie & Co voice, tone, formatting' },
  consistency:  { title: 'Internal Consistency',   desc: 'Terminology, figures match text, cross-refs' },
  completeness: { title: 'Completeness',           desc: 'Required sections, appendices, sign-offs' },
  citations:    { title: 'Citations & References', desc: 'Correct format, all sources listed' },
  compliance:   { title: 'Regulatory Compliance',  desc: 'NES, RMA, district plan alignment' },
  audience:     { title: 'Audience & Tone',        desc: 'Appropriate for planner / client / public' },
  structure:    { title: 'Structure & Flow',       desc: 'Logical ordering, headings, summary at top' },
};

const CHECK_IDS = Object.keys(QA_CHECKS);

const REVIEW_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['overall', 'breakdown', 'issues'],
  properties: {
    overall: { type: 'integer', description: '0-100 submission readiness score' },
    breakdown: {
      type: 'object',
      additionalProperties: false,
      properties: Object.fromEntries(CHECK_IDS.map(k => [k, { type: 'integer' }])),
    },
    issues: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['category', 'categoryId', 'title', 'severity', 'description', 'quote', 'fix'],
        properties: {
          category:    { type: 'string' },
          categoryId:  { type: 'string', enum: CHECK_IDS },
          title:       { type: 'string' },
          severity:    { type: 'string', enum: ['low', 'medium', 'high'] },
          description: { type: 'string' },
          quote:       { type: 'string' },
          fix:         { type: 'string' },
        },
      },
    },
  },
};

const DEFAULT_SYSTEM_PROMPT = `You are an expert technical document reviewer for McKenzie & Co — a New Zealand engineering and planning consultancy.

Your job is to review the provided document and produce a structured QA report.

GUIDELINES
- Apply NZ English spelling and McKenzie & Co house style.
- Reference the relevant standards (RMA, NES, AS/NZS, district plans) where applicable.
- Be constructive: every issue must include a concrete suggested fix.
- Prioritise issues that would block submission to council or a client over stylistic nits.
- Calibrate the overall score: 90+ = ready to submit; 75-89 = minor revisions; 60-74 = revise before submitting; <60 = significant rework.

OUTPUT
Return JSON that matches the supplied schema.

For each issue:
- "categoryId" must be one of: grammar, clarity, technical, brand, consistency, completeness, citations, compliance, audience, structure
- "category" must be the matching human-readable check name
- "severity" must be one of: low, medium, high
- "quote" is a short verbatim excerpt or location pointer (e.g. "§3.2") so the author can find it
- "fix" is specific and actionable

For "breakdown", return a 0-100 integer for each requested check. Omit checks not requested.
"overall" is the weighted submission readiness score, 0-100.`;

// ---------- Helpers ----------

function fillTemplate(tpl, vars) {
  return tpl.replace(/\{(\w+)\}/g, (m, k) => (vars[k] !== undefined && vars[k] !== null ? String(vars[k]) : m));
}

function badRequest(message, status = 400) {
  const e = new Error(message);
  e.status = status;
  return e;
}

function extractDocument(fileName, mimeType, base64) {
  const buffer = Buffer.from(base64, 'base64');
  if (buffer.byteLength > MAX_FILE_BYTES) {
    throw badRequest('File exceeds 25 MB limit', 413);
  }
  const lower = (fileName || '').toLowerCase();

  if (lower.endsWith('.pdf') || mimeType === 'application/pdf') {
    return {
      kind: 'pdf',
      block: {
        type: 'document',
        source: { type: 'base64', media_type: 'application/pdf', data: base64 },
        ...(fileName ? { title: fileName } : {}),
      },
    };
  }

  if (lower.endsWith('.txt') || lower.endsWith('.md') || lower.endsWith('.csv') || (mimeType || '').startsWith('text/')) {
    return { kind: 'text', text: buffer.toString('utf-8') };
  }

  throw badRequest(
    `Unsupported file type "${lower.split('.').pop() || 'unknown'}". This deploy supports PDF and plain text only. Convert .docx/.doc/.rtf/.odt to PDF first.`
  );
}

function buildUserContent(extracted, meta, checksList) {
  const metadata =
`DOCUMENT METADATA
File: ${meta.fileName || 'untitled'}
Purpose: ${meta.purpose}
Project: ${meta.projectNumber || 'unspecified'}${meta.projectName ? ' — ' + meta.projectName : ''}
Version: v${meta.version || 1}
Reviewer context: ${meta.context || 'none provided'}

REQUESTED QA CHECKS
${checksList}
`;

  const blocks = [];
  if (extracted.kind === 'pdf') {
    blocks.push({ type: 'text', text: metadata + '\nDocument content is attached as a PDF.' });
    blocks.push(extracted.block);
    blocks.push({ type: 'text', text: 'Produce the JSON QA review per the schema.' });
  } else {
    const truncated = extracted.text.length > 600000 ? extracted.text.slice(0, 600000) + '\n\n[…truncated]' : extracted.text;
    blocks.push({
      type: 'text',
      text: metadata + '\nDOCUMENT CONTENT\n---\n' + truncated + '\n---\n\nProduce the JSON QA review per the schema.',
    });
  }
  return blocks;
}

/* Raw HTTPS POST to Anthropic /v1/messages.
   Resolves with parsed JSON on 2xx, rejects with Error (status/requestId/body) on non-2xx. */
function callAnthropic(apiKey, payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
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
      timeout: 290000,   // just under Function App's 5-min timeout
    }, (res) => {
      let chunks = '';
      res.on('data', c => { chunks += c; });
      res.on('end', () => {
        let parsed;
        try { parsed = JSON.parse(chunks); } catch { parsed = { _rawBody: chunks }; }
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve({ data: parsed, status: res.statusCode, requestId: res.headers['request-id'] });
        } else {
          const apiMsg = parsed && parsed.error ? (parsed.error.message || JSON.stringify(parsed.error)) : chunks.slice(0, 500);
          const err = new Error(`Anthropic ${res.statusCode}: ${apiMsg}`);
          err.status = res.statusCode;
          err.requestId = res.headers['request-id'];
          err.body = parsed;
          reject(err);
        }
      });
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(new Error('timeout')); });
    req.write(body);
    req.end();
  });
}

// ---------- Handler ----------

module.exports = async function (context, req) {
  try {
    const apiKey = (process.env.ANTHROPIC_API_KEY || '').trim();
    if (!apiKey) {
      context.res = { status: 503, headers: { 'Content-Type': 'application/json' }, body: { error: 'ANTHROPIC_API_KEY is not configured on the server.' } };
      return;
    }
    if (!apiKey.startsWith('sk-ant-')) {
      context.res = { status: 502, headers: { 'Content-Type': 'application/json' }, body: { error: 'ANTHROPIC_API_KEY value is not a valid Anthropic key.' } };
      return;
    }

    const body = req.body;
    if (!body || typeof body !== 'object') throw badRequest('Request body must be JSON');

    const {
      fileBase64, fileName, mimeType,
      projectNumber, projectName, purpose, context: docContext, checks,
      reviewer, version, systemPrompt, checkPrompts,
      model: requestedModel, maxTokens: requestedMaxTokens,
    } = body;

    if (!fileBase64) throw badRequest('fileBase64 is required');
    if (!purpose) throw badRequest('purpose is required');
    if (!Array.isArray(checks) || !checks.length) throw badRequest('checks must be a non-empty array');

    const validChecks = checks.filter(c => CHECK_IDS.includes(c));
    if (!validChecks.length) throw badRequest('No valid checks supplied. Allowed: ' + CHECK_IDS.join(', '));

    const model = ALLOWED_MODELS.has(requestedModel) ? requestedModel : 'claude-opus-4-7';
    const maxTokens = Math.min(Math.max(parseInt(requestedMaxTokens) || 16000, 1024), MAX_TOKENS_CAP);

    const extracted = extractDocument(fileName, mimeType, fileBase64);

    const checksList = validChecks.map(id => {
      const c = QA_CHECKS[id];
      const extra = checkPrompts && checkPrompts[id] ? '\n  Additional guidance: ' + checkPrompts[id] : '';
      return `- ${c.title}: ${c.desc}${extra}`;
    }).join('\n');

    const renderedSystem = fillTemplate(
      systemPrompt && systemPrompt.trim() ? systemPrompt : DEFAULT_SYSTEM_PROMPT,
      {
        filename: fileName || 'untitled',
        purpose,
        projectNumber: projectNumber || 'unspecified',
        projectName: projectName || '',
        context: docContext || 'none provided',
        checks: checksList,
        version: version || 1,
      }
    );

    const userContent = buildUserContent(
      extracted,
      { fileName, purpose, projectNumber, projectName, version, context: docContext },
      checksList
    );

    const payload = {
      model,
      max_tokens: maxTokens,
      thinking: { type: 'adaptive' },
      output_config: {
        effort: 'high',
        format: { type: 'json_schema', schema: REVIEW_SCHEMA },
      },
      cache_control: { type: 'ephemeral' },
      system: renderedSystem,
      messages: [{ role: 'user', content: userContent }],
    };

    const result = await callAnthropic(apiKey, payload);
    const message = result.data;

    const textBlock = (message.content || []).find(b => b.type === 'text');
    if (!textBlock || !textBlock.text) {
      throw new Error('Claude returned no text content (stop_reason: ' + message.stop_reason + ')');
    }

    let parsed;
    try {
      parsed = JSON.parse(textBlock.text);
    } catch (e) {
      throw new Error('Claude response was not valid JSON: ' + e.message + '. First 500 chars: ' + textBlock.text.slice(0, 500));
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ...parsed,
        model: message.model,
        usage: message.usage,
        stop_reason: message.stop_reason,
        requestId: result.requestId,
      },
    };
  } catch (err) {
    const detail = {
      error: err.message || 'Unknown error',
      errorClass: err && err.constructor ? err.constructor.name : typeof err,
      status: err.status,
      requestId: err.requestId,
      anthropicError: err.body && err.body.error ? err.body.error : undefined,
    };
    context.log.error('qaReview error:', JSON.stringify(detail), err.stack);

    context.res = {
      status: err.status || 500,
      headers: { 'Content-Type': 'application/json' },
      body: detail,
    };
  }
};
