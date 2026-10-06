#!/usr/bin/env node
// Daily digest of AVSF requests that have been open too long.
// Usage: node send-digest.js [--dry-run] [--days 7] [--send-empty]
// Reads the same Firestore project as avsf_dashboard.html over the public REST API
// (the dashboard itself runs without sign-in), then emails the digest through SMTP.
'use strict';
const fs = require('fs');
const path = require('path');

const OPEN = new Set(['PROCESSING', 'ON_HOLD', 'PENDING_BUYER', 'IDOC_ERROR']);
const LABEL = { PROCESSING: 'Processing', ON_HOLD: 'On hold', PENDING_BUYER: 'Pending buyer', IDOC_ERROR: 'IDOC error' };
const esc = v => String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const reqNo = n => String(n || '').replace(/^0+/, '') || '0';

function loadFirebaseConfig() {
  let serviceAccount = null;
  if (process.env.FIREBASE_SERVICE_ACCOUNT) {
    serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
    if (process.env.FIREBASE_PROJECT_ID || serviceAccount.project_id) {
      return { projectId: process.env.FIREBASE_PROJECT_ID || serviceAccount.project_id, serviceAccount };
    }
  }
  if (process.env.FIREBASE_PROJECT_ID) {
    return { projectId: process.env.FIREBASE_PROJECT_ID, apiKey: process.env.FIREBASE_API_KEY || '', serviceAccount };
  }
  const html = fs.readFileSync(path.join(__dirname, '..', 'avsf_dashboard.html'), 'utf8');
  const grab = k => (html.match(new RegExp(k + '\\s*:\\s*"([^"]+)"')) || [])[1];
  const cfg = { projectId: grab('projectId'), apiKey: grab('apiKey') };
  if (!cfg.projectId) throw new Error('Could not read the Firebase config from avsf_dashboard.html');
  return { ...cfg, serviceAccount };
}

// Service-account auth (bypasses security rules, like the Admin SDK) with no extra dependency:
// sign an RS256 JWT with Node's crypto and exchange it for an OAuth access token.
async function getAccessToken(sa, tokenUrl) {
  const b64 = o => Buffer.from(JSON.stringify(o)).toString('base64url');
  const now = Math.floor(Date.now() / 1000);
  const head = b64({ alg: 'RS256', typ: 'JWT' });
  const claims = b64({ iss: sa.client_email, scope: 'https://www.googleapis.com/auth/datastore', aud: sa.token_uri || tokenUrl, iat: now, exp: now + 3300 });
  const sig = require('crypto').createSign('RSA-SHA256').update(head + '.' + claims).sign(sa.private_key).toString('base64url');
  const res = await fetch(tokenUrl || sa.token_uri, {
    method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: `${head}.${claims}.${sig}` })
  });
  if (!res.ok) throw new Error(`Service account token request failed: HTTP ${res.status} ${(await res.text()).slice(0, 200)}`);
  return (await res.json()).access_token;
}

function fromValue(v) {
  if (v == null) return null;
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return v.doubleValue;
  if ('booleanValue' in v) return v.booleanValue;
  if ('timestampValue' in v) return v.timestampValue;
  if ('nullValue' in v) return null;
  if ('mapValue' in v) return Object.fromEntries(Object.entries(v.mapValue.fields || {}).map(([k, x]) => [k, fromValue(x)]));
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(fromValue);
  return null;
}

async function fetchAvsfs(cfg, base) {
  const root = base || process.env.FIRESTORE_BASE || 'https://firestore.googleapis.com';
  const out = [];
  const headers = {};
  if (cfg.serviceAccount) headers.authorization = 'Bearer ' + await getAccessToken(cfg.serviceAccount, process.env.TOKEN_URL);
  let token = '';
  do {
    const qs = new URLSearchParams({ pageSize: '300' });
    if (cfg.apiKey && !cfg.serviceAccount) qs.set('key', cfg.apiKey);
    if (token) qs.set('pageToken', token);
    const res = await fetch(`${root}/v1/projects/${cfg.projectId}/databases/(default)/documents/avsfs?${qs}`, { headers });
    if (!res.ok) {
      throw new Error(`Firestore read failed: HTTP ${res.status} ${(await res.text()).replace(/\s+/g, ' ').slice(0, 200)}` +
        (res.status === 403 && !cfg.serviceAccount ? '\nYour Firestore rules block anonymous reads. Set the FIREBASE_SERVICE_ACCOUNT secret (a service account JSON key).' : ''));
    }
    const body = await res.json();
    (body.documents || []).forEach(d => out.push(Object.fromEntries(Object.entries(d.fields || {}).map(([k, v]) => [k, fromValue(v)]))));
    token = body.nextPageToken || '';
  } while (token);
  return out;
}

function ageDays(a, now) {
  const d = String(a.entryDate || '').split('T')[0];
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return null;
  const today = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate());
  return Math.max(0, Math.floor((today - Date.parse(d + 'T00:00:00Z')) / 86400000));
}

const TEAM = [['ROSANA', 'MUNOZ'], ['BRANDON', 'RAMOS'], ['YANNICK', 'ROJAS']];
function isTeamMember(name) {
  const toks = String(name || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().split(/[^A-Z]+/).filter(Boolean);
  return TEAM.some(([f, l]) => toks.includes(f) && toks.includes(l));
}
// Processing in the dashboard, but SAP says another approver still holds it.
function waitingOnOthers(a) {
  if (a.status === 'ON_HOLD' || a.status === 'PENDING_BUYER') return true;
  const ap = String(a.approverName || '').trim();
  return a.status === 'PROCESSING' && !!ap && !/^VM TEAM/i.test(ap) && !isTeamMember(ap);
}

function buildDigest(records, opts = {}) {
  const days = opts.days || 7, now = opts.now || new Date(), url = opts.url || '';
  const open = records.filter(a => OPEN.has(a.status) && !(a.status === 'IDOC_ERROR' && a.completionDate));
  const stuck = open.map(a => ({ ...a, age: ageDays(a, now) })).filter(a => a.age != null && a.age >= days)
    .sort((a, b) => b.age - a.age);
  const ours = stuck.filter(a => !waitingOnOthers(a));
  const others = stuck.filter(waitingOnOthers);
  const since = new Date(now.getTime() - 86400000).toISOString().split('T')[0];
  const doneYesterday = records.filter(a => a.status === 'COMPLETED' && String(a.completionDate || '').split('T')[0] >= since).length;
  const subject = stuck.length
    ? `AVSF digest: ${stuck.length} request${stuck.length === 1 ? '' : 's'} open ${days}+ days (oldest ${stuck[0].age}d)`
    : `AVSF digest: nothing open ${days}+ days`;

  const rowText = a => `  #${reqNo(a.requestNo)}  ${String(a.age).padStart(3)}d  ${LABEL[a.status] || a.status}  CC ${a.companyCode || '-'}  ${a.vendorName || ''}${a.approverName ? '  (' + a.approverName + ')' : ''}`;
  const text = [
    subject, '',
    `Open: ${open.length} | Stuck ${days}+ days: ${stuck.length} | Completed in the last 24h: ${doneYesterday}`, '',
    ours.length ? `IN OUR QUEUE (${ours.length})\n${ours.map(rowText).join('\n')}\n` : '',
    others.length ? `WAITING ON OTHERS (${others.length})\n${others.map(rowText).join('\n')}\n` : '',
    url ? `Open the dashboard: ${url}` : ''
  ].filter(x => x !== '').join('\n');

  const th = 'text-align:left;padding:6px 10px;font-size:11px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e5e7eb;';
  const td = 'padding:7px 10px;font-size:13px;border-bottom:1px solid #f1f5f9;vertical-align:top;';
  const table = (title, list, color) => list.length ? `
    <h3 style="margin:22px 0 6px;font-size:14px;color:${color};">${esc(title)} (${list.length})</h3>
    <table cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;"><thead><tr>
      <th style="${th}">Age</th><th style="${th}">Request</th><th style="${th}">Vendor</th><th style="${th}">CC</th><th style="${th}">Status</th><th style="${th}">With</th></tr></thead><tbody>
      ${list.map(a => `<tr><td style="${td}"><b style="color:${a.age >= 14 ? '#b91c1c' : '#b45309'};">${a.age}d</b></td><td style="${td}">#${esc(reqNo(a.requestNo))} ${esc(a.requestType || '')}</td>
        <td style="${td}">${esc(a.vendorName)}</td><td style="${td}">${esc(a.companyCode || '-')}</td><td style="${td}">${esc(LABEL[a.status] || a.status)}</td><td style="${td}">${esc(a.approverName || '')}</td></tr>`).join('')}
    </tbody></table>` : '';
  const html = `<!doctype html><html><body style="font-family:Segoe UI,Arial,sans-serif;color:#18171a;max-width:760px;margin:0 auto;padding:16px;">
    <h2 style="margin:0 0 4px;">AVSF daily digest</h2>
    <div style="color:#6b7280;font-size:13px;">${esc(now.toISOString().split('T')[0])} · ${open.length} open · <b>${stuck.length}</b> open ${days}+ days · ${doneYesterday} completed in the last 24h</div>
    ${stuck.length ? table('In our queue', ours, '#0f766e') + table('Waiting on others', others, '#b45309')
      : '<p style="margin-top:20px;">Nothing has been open for ' + days + '+ days. 🎉</p>'}
    ${url ? `<p style="margin-top:24px;"><a href="${esc(url)}" style="background:#0f766e;color:#fff;padding:9px 16px;border-radius:8px;text-decoration:none;font-weight:600;">Open the dashboard</a></p>` : ''}
    <p style="color:#9ca3af;font-size:11px;margin-top:24px;">Sent automatically. Change the threshold with STUCK_DAYS or the recipients with DIGEST_TO in the workflow settings.</p>
  </body></html>`;
  return { subject, text, html, stuck, count: stuck.length };
}

async function main(argv = process.argv.slice(2), env = process.env, deps = {}) {
  const flag = n => argv.includes(n);
  const val = n => { const i = argv.indexOf(n); return i >= 0 ? argv[i + 1] : undefined; };
  const days = +(val('--days') || env.STUCK_DAYS || 7);
  const cfg = deps.cfg || loadFirebaseConfig();
  const records = deps.records || await fetchAvsfs(cfg);
  const digest = buildDigest(records, { days, url: env.DASHBOARD_URL || '' });
  console.log(`Loaded ${records.length} requests; ${digest.count} open ${days}+ days.`);
  if (!digest.count && !flag('--send-empty') && env.DIGEST_SEND_EMPTY !== '1') {
    console.log('Nothing to report - no email sent.');
    return { sent: false, digest };
  }
  if (flag('--dry-run') || env.DRY_RUN === '1') {
    const f = path.join(env.DIGEST_OUT_DIR || process.cwd(), 'digest-preview.html');
    fs.writeFileSync(f, digest.html);
    console.log('\n' + digest.text + '\n\nDry run: wrote ' + f);
    return { sent: false, digest };
  }
  const to = (env.DIGEST_TO || '').split(/[,;]/).map(s => s.trim()).filter(Boolean);
  if (!to.length) throw new Error('DIGEST_TO is not set (comma-separated recipient list).');
  const transport = deps.transport || require('nodemailer').createTransport({
    host: env.SMTP_HOST, port: +(env.SMTP_PORT || 587), secure: env.SMTP_SECURE === '1' || +(env.SMTP_PORT) === 465,
    auth: env.SMTP_USER ? { user: env.SMTP_USER, pass: env.SMTP_PASS } : undefined
  });
  const info = await transport.sendMail({ from: env.DIGEST_FROM || env.SMTP_USER, to, subject: digest.subject, text: digest.text, html: digest.html });
  console.log('Sent to ' + to.join(', ') + (info && info.messageId ? ' (' + info.messageId + ')' : ''));
  return { sent: true, digest, info };
}

module.exports = { waitingOnOthers, buildDigest, fetchAvsfs, fromValue, getAccessToken, main };
if (require.main === module) main().catch(e => { console.error(e.message); process.exit(1); });
