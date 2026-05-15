#!/usr/bin/env node
// ============================================================
// scripts/setup-drive.js
// Creates the three Google Drive resources required by Code.gs:
//   1. Proposals folder  → PROPOSALS_FOLDER_ID
//   2. Proposal template → PROPOSAL_TEMPLATE_ID
//   3. Invoice template  → INVOICE_TEMPLATE_ID
// Then patches gas/Code.gs with all three IDs automatically.
//
// Usage:
//   npm run setup-drive
//
// Prerequisite: complete the local dev server onboarding first
//   npm run dev  → open http://localhost:3000 → follow setup steps
// ============================================================

const { google } = require('googleapis');
const fs         = require('fs');
const path       = require('path');
const http       = require('http');
const { exec }   = require('child_process');

const AUTH_PATH  = path.join(__dirname, '..', 'local', '.dev-config.json');
const CODE_GS    = path.join(__dirname, '..', 'gas', 'Code.gs');
const REDIRECT   = 'http://localhost:3001/oauth2callback';
const SCOPES     = [
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/userinfo.email',
];

// ── Helpers ───────────────────────────────────────────────────

function loadConfig() {
  if (!fs.existsSync(AUTH_PATH)) return {};
  try { return JSON.parse(fs.readFileSync(AUTH_PATH, 'utf8')); } catch { return {}; }
}

function saveConfig(patch) {
  const existing = loadConfig();
  fs.writeFileSync(AUTH_PATH, JSON.stringify({ ...existing, ...patch }, null, 2));
}

function log(msg)    { console.log(`  ${msg}`); }
function ok(msg)     { console.log(`  ✅ ${msg}`); }
function warn(msg)   { console.log(`  ⚠️  ${msg}`); }
function header(msg) { console.log(`\n${'─'.repeat(52)}\n  ${msg}\n${'─'.repeat(52)}`); }

// ── OAuth2 ────────────────────────────────────────────────────

function getClient(cfg) {
  return new google.auth.OAuth2(cfg.clientId, cfg.clientSecret, REDIRECT);
}

async function ensureAuthenticated() {
  const cfg = loadConfig();

  if (!cfg.clientId || !cfg.clientSecret) {
    console.error(`
  ❌ No Google API credentials found.

  Please complete the local dev server setup first:
    1. npm run dev
    2. Open http://localhost:3000
    3. Follow the setup steps (enter Client ID + Secret)
    4. Then re-run: npm run setup-drive
`);
    process.exit(1);
  }

  // Re-use existing tokens if available
  if (cfg.tokens && (cfg.tokens.access_token || cfg.tokens.refresh_token)) {
    log('Using existing OAuth tokens from .dev-config.json');
    const client = getClient(cfg);
    client.setCredentials(cfg.tokens);
    client.on('tokens', tokens => saveConfig({ tokens: { ...loadConfig().tokens, ...tokens } }));
    return client;
  }

  // Otherwise do a fresh OAuth dance on port 3001
  return new Promise((resolve, reject) => {
    const client  = getClient(cfg);
    const authUrl = client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: 'consent' });

    const server = http.createServer(async (req, res) => {
      if (!req.url.startsWith('/oauth2callback')) return;
      const code = new URL(req.url, 'http://localhost:3001').searchParams.get('code');
      res.end('<h2>✅ Authenticated! You can close this tab.</h2>');
      server.close();
      try {
        const { tokens } = await client.getToken(code);
        client.setCredentials(tokens);
        saveConfig({ tokens });
        ok('OAuth tokens saved');
        resolve(client);
      } catch (err) { reject(err); }
    });

    server.listen(3001, () => {
      warn('Opening browser for Google login…');
      const cmd = process.platform === 'win32'
        ? `start "" "${authUrl}"`
        : process.platform === 'darwin' ? `open "${authUrl}"` : `xdg-open "${authUrl}"`;
      exec(cmd);
      console.log(`\n  If the browser didn't open, visit:\n  ${authUrl}\n`);
    });
  });
}

// ── Drive helpers ─────────────────────────────────────────────

function getDrive(auth)  { return google.drive({ version: 'v3', auth }); }

async function createFolder(drive, name) {
  const res = await drive.files.create({
    requestBody: { name, mimeType: 'application/vnd.google-apps.folder' },
    fields: 'id, webViewLink',
  });
  return { id: res.data.id, url: res.data.webViewLink };
}

async function createDoc(drive, name, content, parentId) {
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

// ── Template content ──────────────────────────────────────────

const PROPOSAL_TEMPLATE = `ESTIMATE AND PROPOSAL FOR ROOF REPLACEMENT
==========================================
{{COMPANY_NAME}}
{{COMPANY_PHONE}}  |  {{COMPANY_EMAIL}}
{{COMPANY_LICENSE}}  |  {{COMPANY_TAGLINE}}

Date:    {{DATE}}
Job #:   {{ROW_ID}}

CLIENT INFORMATION
──────────────────
Name:     {{CLIENT_NAME}}
Address:  {{CLIENT_ADDRESS}}
Phone:    {{CLIENT_PHONE}}
Email:    {{CLIENT_EMAIL}}

PROJECT DETAILS
───────────────
Roof Type:    {{ROOF_TYPE}}
Square Feet:  {{SQ_FT}}
Warranty:     {{WARRANTY}}
Color Choice: {{COLOR_CHOICE}}

SCOPE OF WORK
─────────────
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

Notes:
{{NOTES}}
`;

const INVOICE_TEMPLATE = `INVOICE
=======
{{COMPANY_NAME}}
{{COMPANY_PHONE}}  |  {{COMPANY_EMAIL}}
{{COMPANY_LICENSE}}  |  {{COMPANY_TAGLINE}}

Invoice Date:  {{DATE}}
Invoice #:     {{ROW_ID}}

BILL TO
───────
{{CLIENT_NAME}}
{{CLIENT_ADDRESS}}
{{CLIENT_PHONE}}
{{CLIENT_EMAIL}}

SERVICES RENDERED
─────────────────
{{LINE_ITEMS}}

══════════════════════════════════════════
TOTAL DUE:  {{TOTAL}}
══════════════════════════════════════════

Payment is due within 30 days of invoice date.
Please make checks payable to {{COMPANY_NAME}}.

Thank you for your business!

Notes:
{{NOTES}}
`;

// ── Patch Code.gs ─────────────────────────────────────────────

function patchCodeGs(ids) {
  let src = fs.readFileSync(CODE_GS, 'utf8');

  if (ids.folderId) {
    src = src.replace(
      /PROPOSALS_FOLDER_ID:\s*'[^']*'/,
      `PROPOSALS_FOLDER_ID:  '${ids.folderId}'`
    );
  }
  if (ids.proposalTemplateId) {
    src = src.replace(
      /PROPOSAL_TEMPLATE_ID:\s*'[^']*'/,
      `PROPOSAL_TEMPLATE_ID: '${ids.proposalTemplateId}'`
    );
  }
  if (ids.invoiceTemplateId) {
    src = src.replace(
      /INVOICE_TEMPLATE_ID:\s*'[^']*'/,
      `INVOICE_TEMPLATE_ID:  '${ids.invoiceTemplateId}'`
    );
  }

  fs.writeFileSync(CODE_GS, src);
}

// ── Main ──────────────────────────────────────────────────────

async function main() {
  header('Roofing Proposal System — Drive Setup');

  const auth  = await ensureAuthenticated();
  const drive = getDrive(auth);
  const cfg   = loadConfig();
  const company = cfg.companyName || 'Roofing Co.';

  // ── 1. Proposals folder ─────────────────────────────────────
  let folderId = cfg.folderId;
  let folderUrl;

  if (folderId) {
    ok(`Proposals folder already exists (reusing)`);
    folderUrl = cfg.folderUrl || `https://drive.google.com/drive/folders/${folderId}`;
  } else {
    log(`Creating Proposals folder: "${company} — Proposals"…`);
    const folder = await createFolder(drive, `${company} — Proposals`);
    folderId  = folder.id;
    folderUrl = folder.url;
    saveConfig({ folderId, folderUrl });
    ok(`Proposals folder created`);
  }

  // ── 2. Templates folder (sibling of Proposals) ──────────────
  log(`Creating Templates folder: "${company} — Templates"…`);
  const tmplFolder = await createFolder(drive, `${company} — Templates`);
  ok(`Templates folder created`);

  // ── 3. Proposal template doc ─────────────────────────────────
  log(`Creating Proposal template Google Doc…`);
  const proposalDoc = await createDoc(
    drive,
    `[TEMPLATE] Roof Proposal — ${company}`,
    PROPOSAL_TEMPLATE,
    tmplFolder.id
  );
  ok(`Proposal template created`);

  // ── 4. Invoice template doc ──────────────────────────────────
  log(`Creating Invoice template Google Doc…`);
  const invoiceDoc = await createDoc(
    drive,
    `[TEMPLATE] Roof Invoice — ${company}`,
    INVOICE_TEMPLATE,
    tmplFolder.id
  );
  ok(`Invoice template created`);

  // ── 5. Patch Code.gs ─────────────────────────────────────────
  log(`Patching gas/Code.gs with new IDs…`);
  patchCodeGs({
    folderId,
    proposalTemplateId: proposalDoc.id,
    invoiceTemplateId:  invoiceDoc.id,
  });
  ok(`Code.gs updated`);

  // ── 6. Save IDs to .dev-config.json ──────────────────────────
  saveConfig({
    proposalTemplateId: proposalDoc.id,
    invoiceTemplateId:  invoiceDoc.id,
    templatesUrl:       tmplFolder.url,
  });

  // ── Summary ───────────────────────────────────────────────────
  header('Setup Complete');
  console.log(`
  📁 Proposals Folder:
     ${folderUrl}

  📄 Proposal Template:
     ${proposalDoc.url}

  🧾 Invoice Template:
     ${invoiceDoc.url}

  ─────────────────────────────────────────────────────
  ✅ gas/Code.gs has been patched with all three IDs.

  NEXT STEPS
  ──────────
  1. Open the Proposal and Invoice templates in the links
     above. Add your logo, colors, and any extra formatting
     you want — the {{MERGE_TAGS}} will still be replaced.

  2. Fill in your company branding in gas/Code.gs CONFIG:
       COMPANY_NAME, COMPANY_PHONE, COMPANY_EMAIL,
       COMPANY_LICENSE, COMPANY_TAGLINE

  3. Push to Google Apps Script:
       npm run push

  4. In the Apps Script editor:
       Deploy → New Deployment → Web App
       Execute as: Me | Who has access: Anyone
       Copy the Web App URL
  `);
}

main().catch(err => {
  console.error('\n  ❌ Error:', err.message);
  process.exit(1);
});
