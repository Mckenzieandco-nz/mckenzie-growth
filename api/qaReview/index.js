/**
 * /api/qaReview — runs an AI QA review on an uploaded document via Claude.
 *
 * Body shape (JSON):
 * {
 *   fileBase64: string,         // base64 of the file bytes
 *   fileName: string,
 *   mimeType?: string,
 *   projectNumber?: string,
 *   projectName?: string,
 *   purpose: string,
 *   context?: string,           // free-text reviewer context
 *   checks: string[],           // QA check IDs (see QA_CHECKS keys below)
 *   reviewer?: string,
 *   version?: number,
 *   systemPrompt?: string,      // template override from Admin tab
 *   checkPrompts?: { [id]: string },
 *   model?: string,             // whitelisted Claude model alias
 *   maxTokens?: number          // capped server-side
 * }
 *
 * Returns: { overall, breakdown, issues, model, usage }
 *
 * Auth: ANTHROPIC_API_KEY must be set in Azure SWA Configuration. The frontend
 * never sees the key.
 */

const { Anthropic } = require('@anthropic-ai/sdk');
const mammoth = require('mammoth');

// ---------- Constants ----------

const ALLOWED_MODELS = new Set([
  'claude-opus-4-7',
  'claude-sonnet-4-6',
  'claude-haiku-4-5',
]);

const MAX_TOKENS_CAP = 32000;
const MAX_FILE_BYTES = 25 * 1024 * 1024;   // 25 MB raw file

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

// JSON Schema for the structured output. Matches the frontend's review shape.
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
You will receive document metadata, the QA checks to run, and the document content (as a PDF document block or extracted text). Return JSON that matches the supplied schema.

For each issue:
- "categoryId" must be one of: grammar, clarity, technical, brand, consistency, completeness, citations, compliance, audience, structure
- "category" must be the matching human-readable check name
- "severity" must be one of: low, medium, high
- "quote" is a short verbatim excerpt or location pointer (e.g. "§3.2") so the author can find it
- "fix" is specific and actionable

For "breakdown", return a 0-100 integer for each check that was requested. Omit checks that were not requested.
"overall" is the weighted submission readiness score, 0-100.`;

// ---------- Helpers ----------

function fillTemplate(tpl, vars) {
  return tpl.replace(/\{(\w+)\}/g, (m, k) => (vars[k] !== undefined && vars[k] !== null ? String(vars[k]) : m));
}

async function extractDocument(fileName, mimeType, base64) {
  const buffer = Buffer.from(base64, 'base64');
  if (buffer.byteLength > MAX_FILE_BYTES) {
    throw Object.assign(new Error('File exceeds 25 MB limit'), { status: 413 });
  }
  const lower = (fileName || '').toLowerCase();

  // PDFs go to Claude as a document block — native, no client-side parsing.
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

  // DOCX → extract via mammoth.
  if (lower.endsWith('.docx')) {
    const r = await mammoth.extractRawText({ buffer });
    return { kind: 'text', text: r.value };
  }

  // Plain text-like — decode utf-8.
  if (lower.endsWith('.txt') || lower.endsWith('.md') || lower.endsWith('.csv') || (mimeType || '').startsWith('text/')) {
    return { kind: 'text', text: buffer.toString('utf-8') };
  }

  // Best-effort fallback. .doc / .rtf / .odt aren't well-supported by browsers
  // anyway — give a clear hint instead of returning garbage.
  throw Object.assign(new Error(
    `Unsupported file type "${lower.split('.').pop() || 'unknown'}". Supported: PDF, DOCX, TXT. Convert .doc/.rtf/.odt to PDF first.`
  ), { status: 400 });
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
    // Text — inline. Trim if absurdly long; Claude has 1M context but we're paying per token.
    const truncated = extracted.text.length > 600000 ? extracted.text.slice(0, 600000) + '\n\n[…truncated]' : extracted.text;
    blocks.push({
      type: 'text',
      text: metadata + '\nDOCUMENT CONTENT\n---\n' + truncated + '\n---\n\nProduce the JSON QA review per the schema.',
    });
  }
  return blocks;
}

function badRequest(message, status = 400) {
  const e = new Error(message);
  e.status = status;
  return e;
}

// ---------- Handler ----------

module.exports = async function (context, req) {
  try {
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      context.res = {
        status: 503,
        headers: { 'Content-Type': 'application/json' },
        body: { error: 'ANTHROPIC_API_KEY is not configured on the server. Set it in Azure SWA → Configuration.' },
      };
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

    // Extract / prepare document
    const extracted = await extractDocument(fileName, mimeType, fileBase64);

    // Render system prompt — preserves admin placeholders for backward compat.
    // For maximum cache hit rate, prefer the default prompt (no placeholders).
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

    const client = new Anthropic({ apiKey });

    // Stream to avoid HTTP timeouts on long reviews; collect via finalMessage().
    const stream = client.messages.stream({
      model,
      max_tokens: maxTokens,
      thinking: { type: 'adaptive' },
      output_config: {
        effort: 'high',
        format: { type: 'json_schema', schema: REVIEW_SCHEMA },
      },
      // Top-level cache_control auto-caches the last cacheable block (system prompt)
      // — repeated calls with the same system prompt + model hit the cache.
      cache_control: { type: 'ephemeral' },
      system: renderedSystem,
      messages: [{ role: 'user', content: userContent }],
    });

    const finalMessage = await stream.finalMessage();

    // Extract structured JSON output
    const textBlock = finalMessage.content.find(b => b.type === 'text');
    if (!textBlock || !textBlock.text) {
      throw new Error('Claude returned no text content (stop_reason: ' + finalMessage.stop_reason + ')');
    }

    let parsed;
    try {
      parsed = JSON.parse(textBlock.text);
    } catch (e) {
      // Should not happen with structured outputs, but be safe.
      throw new Error('Claude response was not valid JSON: ' + e.message);
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ...parsed,
        model: finalMessage.model,
        usage: {
          input_tokens: finalMessage.usage.input_tokens,
          output_tokens: finalMessage.usage.output_tokens,
          cache_read_input_tokens: finalMessage.usage.cache_read_input_tokens || 0,
          cache_creation_input_tokens: finalMessage.usage.cache_creation_input_tokens || 0,
        },
        stop_reason: finalMessage.stop_reason,
      },
    };
  } catch (err) {
    context.log.error('qaReview error:', err.message, err.stack);

    let status = err.status || 500;
    if (err instanceof Anthropic.AuthenticationError) status = 401;
    else if (err instanceof Anthropic.PermissionDeniedError) status = 403;
    else if (err instanceof Anthropic.BadRequestError) status = 400;
    else if (err instanceof Anthropic.NotFoundError) status = 404;
    else if (err instanceof Anthropic.RateLimitError) status = 429;
    else if (err instanceof Anthropic.APIError) status = err.status || 500;

    context.res = {
      status,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message || 'Unknown error' },
    };
  }
};
