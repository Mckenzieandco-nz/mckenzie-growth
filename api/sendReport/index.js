const nodemailer = require('nodemailer');

module.exports = async function (context, req) {
  try {
    const { subject, to, html } = req.body || {};

    if (!subject || !to || !html) {
      context.res = { status: 400, body: { error: 'subject, to, and html are required' } };
      return;
    }

    const recipients = Array.isArray(to) ? to : [to];
    if (!recipients.length) {
      context.res = { status: 400, body: { error: 'No recipients provided' } };
      return;
    }

    // SMTP config from Azure App Settings environment variables
    const smtpHost = process.env.SMTP_HOST;
    const smtpPort = parseInt(process.env.SMTP_PORT || '587', 10);
    const smtpUser = process.env.SMTP_USER;
    const smtpPass = process.env.SMTP_PASS;
    const smtpFrom = process.env.SMTP_FROM || smtpUser;

    if (!smtpHost || !smtpUser || !smtpPass) {
      context.res = {
        status: 500,
        body: { error: 'SMTP not configured. Set SMTP_HOST, SMTP_USER, SMTP_PASS in Azure App Settings.' }
      };
      return;
    }

    const transporter = nodemailer.createTransport({
      host: smtpHost,
      port: smtpPort,
      secure: smtpPort === 465,
      auth: { user: smtpUser, pass: smtpPass },
      tls: { rejectUnauthorized: false }
    });

    await transporter.sendMail({
      from: smtpFrom,
      to: recipients.join(', '),
      subject,
      html
    });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, sent: recipients.length }
    };
  } catch (err) {
    context.log.error('sendReport error:', err.message);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message }
    };
  }
};
