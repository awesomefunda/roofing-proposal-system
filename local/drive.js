// ============================================================
// local/drive.js — Google Drive API operations
// ============================================================

const { google } = require('googleapis');
const { getAuthenticatedClient, loadConfig } = require('./auth');

function getDrive() {
  return google.drive({ version: 'v3', auth: getAuthenticatedClient() });
}

// ── Internal helpers ──────────────────────────────────────────

async function _createFolder(drive, name) {
  const res = await drive.files.create({
    requestBody: { name, mimeType: 'application/vnd.google-apps.folder' },
    fields: 'id, webViewLink',
  });
  return { id: res.data.id, url: res.data.webViewLink };
}

async function _createDoc(drive, name, content, parentId) {
  const res = await drive.files.create({
    requestBody: {
      name,
      mimeType: 'application/vnd.google-apps.document',
      parents: [parentId],
    },
    media: { mimeType: 'text/plain', body: content },
    fields: 'id, webViewLink',
  });
  return { id: res.data.id, url: res.data.webViewLink };
}

// ── Template text ─────────────────────────────────────────────
// These placeholders are replaced at runtime by Code.gs.
// Open the links in your browser after setup to format them
// (add logo, change fonts, etc.) — the {{TAGS}} stay as-is.

const PROPOSAL_TEMPLATE = `ESTIMATE AND PROPOSAL FOR ROOF REPLACEMENT
==========================================
{{COMPANY_NAME}}
{{COMPANY_PHONE}}  |  {{COMPANY_EMAIL}}
{{COMPANY_LICENSE}}  |  {{COMPANY_TAGLINE}}

Date:    {{DATE}}
Job #:   {{ROW_ID}}

──────────────────────────────────────────
CLIENT
──────────────────────────────────────────
Name:     {{CLIENT_NAME}}
Address:  {{CLIENT_ADDRESS}}
Phone:    {{CLIENT_PHONE}}
Email:    {{CLIENT_EMAIL}}

──────────────────────────────────────────
PROJECT
──────────────────────────────────────────
Roof Type:    {{ROOF_TYPE}}
Square Feet:  {{SQ_FT}}
Warranty:     {{WARRANTY}}
Color Choice: {{COLOR_CHOICE}}

──────────────────────────────────────────
SCOPE OF WORK
──────────────────────────────────────────
{{LINE_ITEMS}}

══════════════════════════════════════════
TOTAL:  {{TOTAL}}
══════════════════════════════════════════

Additional rotted wood will be charged at {{ROTTED_WOOD_RATE}} upon homeowner approval.
The work area will be kept clean at the end of every work day.

If you agree to this estimate, please sign and date below.

Homeowner signature:  ___________________________  Date: __________

By signing, you authorize {{COMPANY_NAME}} to perform the above work
and agree to pay the total amount upon completion.

{{NOTES}}
`;

const INVOICE_TEMPLATE = `INVOICE
=======
{{COMPANY_NAME}}
{{COMPANY_PHONE}}  |  {{COMPANY_EMAIL}}
{{COMPANY_LICENSE}}  |  {{COMPANY_TAGLINE}}

Invoice Date:  {{DATE}}
Invoice #:     {{ROW_ID}}

──────────────────────────────────────────
BILL TO
──────────────────────────────────────────
{{CLIENT_NAME}}
{{CLIENT_ADDRESS}}
{{CLIENT_PHONE}}
{{CLIENT_EMAIL}}

──────────────────────────────────────────
SERVICES RENDERED
──────────────────────────────────────────
{{LINE_ITEMS}}

══════════════════════════════════════════
TOTAL DUE:  {{TOTAL}}
══════════════════════════════════════════

Payment is due within 30 days of invoice date.
Please make checks payable to {{COMPANY_NAME}}.

Thank you for your business!

{{NOTES}}
`;

// ── Public: one-time setup ────────────────────────────────────
// Creates the Proposals folder + Templates folder + both Doc
// templates in the user's Google Drive. Called automatically
// right after OAuth login completes — no extra clicks needed.

async function setupDriveResources(companyName) {
  const drive = getDrive();
  const co    = companyName || 'Roofing Co.';

  const proposalsFolder = await _createFolder(drive, `${co} — Proposals`);
  const templatesFolder = await _createFolder(drive, `${co} — Templates`);

  const proposalDoc = await _createDoc(
    drive,
    `[TEMPLATE] Roof Proposal — ${co}`,
    PROPOSAL_TEMPLATE,
    templatesFolder.id
  );

  const invoiceDoc = await _createDoc(
    drive,
    `[TEMPLATE] Roof Invoice — ${co}`,
    INVOICE_TEMPLATE,
    templatesFolder.id
  );

  return {
    folderId:            proposalsFolder.id,
    folderUrl:           proposalsFolder.url,
    folderName:          `${co} — Proposals`,
    templatesUrl:        templatesFolder.url,
    proposalTemplateId:  proposalDoc.id,
    proposalTemplateUrl: proposalDoc.url,
    invoiceTemplateId:   invoiceDoc.id,
    invoiceTemplateUrl:  invoiceDoc.url,
  };
}

// ── Public: create a proposal doc for local testing ───────────
// Mimics what Code.gs does in production: fills in the template
// and saves a Google Doc to the Proposals folder.

async function createProposalDoc(payload) {
  const drive = getDrive();
  const cfg   = loadConfig();

  if (!cfg.folderId) {
    throw new Error('No Drive folder — complete onboarding first.');
  }

  const today    = new Date().toISOString().split('T')[0];
  const safeName = payload.clientName.replace(/[^a-zA-Z0-9]/g, '_');
  const content  = buildProposalText(payload, cfg);

  return _createDoc(
    drive,
    `[LOCAL] Proposal_${safeName}_${today}`,
    content,
    cfg.folderId
  );
}

function buildProposalText(p, cfg) {
  const lines = [
    'ESTIMATE AND PROPOSAL FOR ROOF REPLACEMENT',
    '==========================================',
    `[LOCAL DEV — production generates a branded PDF via GAS]`,
    '',
    `Company:  ${cfg.companyName || 'Roofing Co.'}`,
    `Date:     ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`,
    '',
    `Client:   ${p.clientName}`,
    `Address:  ${p.clientAddress}`,
    p.clientPhone  ? `Phone:    ${p.clientPhone}` : null,
    p.clientEmail  ? `Email:    ${p.clientEmail}` : null,
    '',
    `Roof Type:  ${p.roofType   || 'N/A'}`,
    p.sqft         ? `Sq Ft:      ${Number(p.sqft).toLocaleString()}` : null,
    `Warranty:   ${p.warranty   || '7-year'}`,
    p.colorChoice  ? `Color:      ${p.colorChoice}` : null,
    '',
    'SCOPE OF WORK',
    '─────────────',
    ...p.lineItems.map((item, i) => {
      const price = item.price ? `  —  $${Number(item.price).toLocaleString()}` : '  (included)';
      return `${String(i + 1).padStart(2)}. ${item.description}${price}`;
    }),
    '',
    '─────────────────────────────────────────',
    `TOTAL:  $${Number(p.total).toLocaleString()}`,
    '─────────────────────────────────────────',
    '',
    'Additional rotted wood charged at $26.00/ft upon homeowner approval.',
    'Area will be kept clean every day.',
    '',
    'Homeowner signature:  ___________________________  Date: ________',
    '',
    p.notes ? `Notes: ${p.notes}` : null,
  ].filter(l => l !== null);

  return lines.join('\n');
}

// ── Get the connected account email ───────────────────────────

async function getAccountEmail() {
  const oauth2 = google.oauth2({ version: 'v2', auth: getAuthenticatedClient() });
  const res    = await oauth2.userinfo.get();
  return res.data.email;
}

module.exports = { setupDriveResources, createProposalDoc, getAccountEmail };
