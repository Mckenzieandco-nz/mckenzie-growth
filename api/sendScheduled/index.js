/**
 * sendScheduled — HTTP trigger called by Power Automate on a schedule.
 * Query param: ?freq=monthly|quarterly|annual
 *
 * Reads KPIs and all data from Cosmos, computes metrics for the
 * appropriate completed period, builds an HTML report, and emails it
 * via SMTP (same env vars as sendReport).
 */

const { database } = require('../shared/cosmos');
const nodemailer = require('nodemailer');

// ─── KPI computation (mirrors client-side computeKpi) ─────────────────────────

function inPeriod(date, from, to) {
  return (!from || date >= from) && (!to || date <= to);
}

function computeKpi(metric, data, from, to) {
  const users = data.users || [];
  const n = users.length || 1;
  switch (metric) {
    case 'headcount': return users.length;
    case 'cpd_total': return (data.cpdLog || []).filter(e => inPeriod(e.date, from, to)).reduce((a, e) => a + (e.hours || 0), 0);
    case 'cpd_avg': {
      const t = (data.cpdLog || []).filter(e => inPeriod(e.date, from, to)).reduce((a, e) => a + (e.hours || 0), 0);
      return parseFloat((t / n).toFixed(1));
    }
    case 'review_coverage': {
      const s = new Set((data.reviews || []).filter(r => inPeriod(r.date, from, to)).map(r => r.userId));
      return users.length ? Math.round((s.size / users.length) * 100) : 0;
    }
    case 'oto_coverage': {
      const ids = new Set(
        (data.onetoones || []).filter(o => inPeriod(o.date, from, to))
          .flatMap(o => [o.userId, o.withId])
          .filter(id => users.find(u => u.id === id))
      );
      return users.length ? Math.round((ids.size / users.length) * 100) : 0;
    }
    case 'training_complete': {
      const due = (data.assignments || []).filter(a => !a.dueDate || a.dueDate <= to);
      const done = due.filter(a => a.status === 'completed');
      return due.length ? Math.round((done.length / due.length) * 100) : 100;
    }
    case 'overdue_training': {
      const today = new Date().toISOString().split('T')[0];
      return (data.assignments || []).filter(a => a.status !== 'completed' && a.dueDate && a.dueDate < today).length;
    }
    case 'skills_avg': {
      const t = users.reduce((a, u) => a + Object.keys(u.skills || {}).length, 0);
      return parseFloat((t / n).toFixed(1));
    }
    case 'disclosures_open': return (data.disclosures || []).filter(d => d.status !== 'resolved').length;
    default: return 0;
  }
}

const LABELS = {
  headcount: 'Headcount', cpd_total: 'Total CPD Hours', cpd_avg: 'Avg CPD Hours/Person',
  review_coverage: 'Review Coverage', oto_coverage: '1:1 Coverage',
  training_complete: 'Training Completion', overdue_training: 'Overdue Assignments',
  skills_avg: 'Avg Skills Assessed', disclosures_open: 'Open Disclosures',
};
const UNITS = {
  headcount: 'people', cpd_total: 'hrs', cpd_avg: 'hrs',
  review_coverage: '%', oto_coverage: '%', training_complete: '%',
  overdue_training: 'items', skills_avg: 'skills', disclosures_open: 'items',
};

// ─── Period helpers ────────────────────────────────────────────────────────────

function getPeriod(freq, now) {
  const y = now.getFullYear(), m = now.getMonth(); // 0-based
  if (freq === 'monthly') {
    const pm = m === 0 ? 12 : m, py = m === 0 ? y - 1 : y;
    const from = `${py}-${String(pm).padStart(2, '0')}-01`;
    const last = new Date(py, pm, 0).getDate();
    const to = `${py}-${String(pm).padStart(2, '0')}-${String(last).padStart(2, '0')}`;
    return { from, to, label: new Date(py, pm - 1, 1).toLocaleDateString('en-NZ', { month: 'long', year: 'numeric' }) };
  }
  if (freq === 'quarterly') {
    const qm = m === 0 ? 11 : m - 1;
    const q = Math.floor(qm / 3) + 1, qy = m === 0 ? y - 1 : y;
    const sm = (q - 1) * 3 + 1, em = sm + 2;
    const from = `${qy}-${String(sm).padStart(2, '0')}-01`;
    const last = new Date(qy, em, 0).getDate();
    const to = `${qy}-${String(em).padStart(2, '0')}-${String(last).padStart(2, '0')}`;
    return { from, to, label: `Q${q} ${qy}` };
  }
  if (freq === 'annual') {
    const py = y - 1;
    return { from: `${py}-01-01`, to: `${py}-12-31`, label: String(py) };
  }
  return null;
}

// ─── HTML builder ──────────────────────────────────────────────────────────────

function buildHtml(period, freq, rows) {
  const metCount = rows.filter(r => r.met).length;
  const e = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const cap = s => s ? s.charAt(0).toUpperCase() + s.slice(1) : '';
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body{font-family:Arial,sans-serif;max-width:620px;margin:0 auto;padding:24px;color:#3d3935}
h1{color:#48773c;margin:0 0 4px}.sub{color:#888;font-size:.88rem;margin-bottom:20px}
.summary{background:#f5f3ee;border-radius:8px;padding:16px 24px;margin-bottom:24px}
.score{font-size:2rem;font-weight:800;color:#48773c}.sl{font-size:.85rem;color:#666}
table{width:100%;border-collapse:collapse;margin-bottom:20px}
th{background:#3d3935;color:#fff;padding:8px 12px;text-align:left;font-size:.78rem}
td{padding:8px 12px;border-bottom:1px solid #e8e4dc;font-size:.83rem}
tr:nth-child(even)td{background:#faf9f6}
.met{color:#48773c;font-weight:700}.notmet{color:#c0392b;font-weight:700}
.foot{font-size:.7rem;color:#aaa;border-top:1px solid #e8e4dc;padding-top:12px;margin-top:8px}
</style></head><body>
<h1>McKenzie Growth — KPI Report</h1>
<div class="sub">${e(period.label)} · ${cap(freq)} report · Auto-generated ${new Date().toLocaleDateString('en-NZ', { day: 'numeric', month: 'long', year: 'numeric' })}</div>
<div class="summary"><div class="score">${metCount}/${rows.length}</div><div class="sl">KPIs on target</div></div>
<table><thead><tr><th>KPI</th><th>Metric</th><th>Result</th><th>Target</th><th>Status</th></tr></thead><tbody>
${rows.map(r => {
    const unit = UNITS[r.k.metric] || '';
    return `<tr><td><strong>${e(r.k.name)}</strong>${r.k.description ? `<br><span style="color:#888;font-size:.73rem">${e(r.k.description)}</span>` : ''}</td><td>${LABELS[r.k.metric] || r.k.metric}</td><td>${r.val}${unit === '%' ? '%' : ' ' + unit}</td><td>${r.k.targetDirection === 'lte' ? '≤' : '≥'} ${r.k.target}${unit === '%' ? '%' : ' ' + unit}</td><td class="${r.met ? 'met' : 'notmet'}">${r.met ? '✅ Met' : '❌ Not met'}</td></tr>`;
  }).join('')}
</tbody></table>
<div class="foot">McKenzie Growth Platform — Automatically generated by Power Automate schedule.</div>
</body></html>`;
}

// ─── Main handler ──────────────────────────────────────────────────────────────

module.exports = async function (context, req) {
  const freq = (req.query.freq || '').toLowerCase();
  if (!['monthly', 'quarterly', 'annual'].includes(freq)) {
    context.res = { status: 400, body: { error: 'freq must be monthly, quarterly, or annual' } };
    return;
  }

  try {
    // Load all data from Cosmos
    const cols = ['users', 'materials', 'assignments', 'reviews', 'onetoones', 'cpdLog', 'disclosures', 'kpis', 'settings'];
    const data = {};
    for (const name of cols) {
      try {
        const { resource } = await database.container(name).item('all', 'all').read();
        data[name] = resource ? resource.data : (name === 'settings' ? {} : []);
      } catch (e) { data[name] = name === 'settings' ? {} : []; }
    }

    const settings = data.settings || {};
    const kpis = data.kpis || [];
    const recipients = (settings.reportRecipients || '').split(',').map(e => e.trim()).filter(Boolean);

    if (!recipients.length) {
      context.res = { status: 200, body: { ok: false, reason: 'No recipients configured' } };
      return;
    }

    const period = getPeriod(freq, new Date());
    if (!period) {
      context.res = { status: 400, body: { error: 'Could not determine period' } };
      return;
    }

    const activeKpis = kpis.filter(k => k.active);
    if (!activeKpis.length) {
      context.res = { status: 200, body: { ok: false, reason: 'No active KPIs' } };
      return;
    }

    const rows = activeKpis.map(k => {
      const val = computeKpi(k.metric, data, period.from, period.to);
      const met = k.targetDirection === 'lte' ? val <= k.target : val >= k.target;
      return { k, val, met };
    });

    const metCount = rows.filter(r => r.met).length;
    const html = buildHtml(period, freq, rows);
    const subject = `McKenzie Growth KPI Report — ${period.label} (${metCount}/${rows.length} on target)`;

    // Send via SMTP
    const smtpHost = process.env.SMTP_HOST, smtpUser = process.env.SMTP_USER, smtpPass = process.env.SMTP_PASS;
    const smtpPort = parseInt(process.env.SMTP_PORT || '587', 10);
    const smtpFrom = process.env.SMTP_FROM || smtpUser;

    if (!smtpHost || !smtpUser || !smtpPass) {
      context.res = { status: 500, body: { error: 'SMTP not configured. Set SMTP_HOST, SMTP_USER, SMTP_PASS in Azure App Settings.' } };
      return;
    }

    const nodemailer = require('nodemailer');
    const transporter = nodemailer.createTransport({
      host: smtpHost, port: smtpPort, secure: smtpPort === 465,
      auth: { user: smtpUser, pass: smtpPass },
      tls: { rejectUnauthorized: false }
    });

    await transporter.sendMail({ from: smtpFrom, to: recipients.join(', '), subject, html });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, freq, period: period.label, sent: recipients.length, kpisReported: rows.length, onTarget: metCount }
    };
  } catch (err) {
    context.log.error('sendScheduled error:', err.message);
    context.res = { status: 500, body: { error: err.message } };
  }
};
