// ============================================================
// local/provision.js — one-shot contractor provisioning
// ============================================================
// Uses the contractor's own Google OAuth token to create
// everything inside their Google account. The platform owner
// has zero access after this function completes.
//
// What gets created in the contractor's Google account:
//   1. Google Sheet  (Catalog + Leads pre-populated)
//   2. Proposals folder  (Drive)
//   3. Templates folder  (Drive)
//   4. Logo file         (Drive, if uploaded)
//   5. Proposal template Google Doc
//   6. Invoice template Google Doc
//   7. Apps Script project (bound to the Sheet)
//   8. Web App deployment
//
// Returns: { webAppUrl, spreadsheetUrl, folderUrl, ... }
// ============================================================

const { google } = require('googleapis');
const fs         = require('fs');
const path       = require('path');
const { Readable } = require('stream');

const CODE_GS_PATH    = path.join(__dirname, '..', 'gas', 'Code.gs');
const INDEX_HTML_PATH = path.join(__dirname, '..', 'gas', 'Index.html');

// GAS manifest — defines the web app settings
const GAS_MANIFEST = {
  timeZone: 'America/Los_Angeles',
  exceptionLogging: 'STACKDRIVER',
  runtimeVersion: 'V8',
  webapp: {
    executeAs: 'USER_DEPLOYING', // runs as the contractor — not the platform
    access: 'ANYONE',
  },
  oauthScopes: [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/userinfo.email',
  ],
};

// Catalog line items pre-loaded into the Leads sheet
const CATALOG_ITEMS = [
  ['Item', 'Description', 'Price', 'Unit', 'Category', 'Active'],
  ['permits',        'Obtain all building permits & licenses required by the city',      0,    'included in total',      'Permits & Admin',    'Yes'],
  ['remove_flat',    'Remove and dispose of tar and gravel (flat roof)',                1800, 'per job',                'Removal',            'Yes'],
  ['remove_tile',    'Remove and dispose of concrete tile layer',                       4200, 'per job',                'Removal',            'Yes'],
  ['remove_shingle', 'Remove and dispose of composition shingles',                      2400, 'per layer',              'Removal',            'Yes'],
  ['osb_716',        'Install APA 7/16 OSB plywood',                                   3750, 'per job',                'Decking',            'Yes'],
  ['osb_18k',        'Install 7/16 OSB plywood (18,000 sqft)',                          7800, 'per job',                'Decking',            'Yes'],
  ['peel_stick',     'Install Peel-N-Stick titanium PSU30 underlayment',               1800, 'per job',                'Underlayment',       'Yes'],
  ['synthetic_under','Install Certainteed Diamond Deck synthetic underlayment',         1200, 'per job',                'Underlayment',       'Yes'],
  ['base_sheet',     'Install 3-ply: 1 base sheet + 2 torch layers (flat roof)',        2800, 'per job · 1,500 sqft',  'Underlayment',       'Yes'],
  ['tile_standard',  'Install villa tile standard (18,000 sqft)',                       5400, 'customer choice',        'Roofing Material',   'Yes'],
  ['shingle_30yr',   'Install CertainTeed 30-year asphalt composition shingles',        8200, 'customer choice',        'Roofing Material',   'Yes'],
  ['shingle_40yr',   'Install 40-year manufactured warranty composition shingles',      9800, 'customer choice',        'Roofing Material',   'Yes'],
  ['ohagin_3',       "Install (3) O'Hagin vents low profile",                           450,  'per job',                'Vents & Flashing',   'Yes'],
  ['ohagin_6',       "Install (6) O'Hagin vents low profile",                           900,  'per job',                'Vents & Flashing',   'Yes'],
  ['metal_flashing', 'Prime and paint all new pipe metal flashing',                       0,  'included',               'Vents & Flashing',   'Yes'],
  ['drip_edge',      'Install metal drip edge around eaves 2x2',                          0,  'included',               'Vents & Flashing',   'Yes'],
  ['chimney_saddle', 'Furnish and install new chimney saddle',                           320, 'per unit',               'Vents & Flashing',   'Yes'],
  ['valley_flash',   'Install new valley metal flashing',                                  0,  'included',               'Vents & Flashing',   'Yes'],
  ['gutter_95ft',    'Replace 95 ft rain gutter',                                       1330, 'customer choice',        'Gutters',            'Yes'],
  ['gutter_180ft',   'Replace 180 ft rain gutter and downspout',                        2940, 'customer choice',        'Gutters',            'Yes'],
  ['gutter_334ft',   'Replace 334 ft rain gutter',                                      3674, 'customer choice',        'Gutters',            'Yes'],
  ['cleanup',        'Clean all debris and take to recycling center',                      0,  'daily + end of job',     'Cleanup & Warranty', 'Yes'],
  ['rotted_wood',    'Additional rotted wood charged at $26.00 per foot',                 0,  'upon homeowner approval','Cleanup & Warranty', 'Yes'],
];

const LEADS_HEADERS = [
  ['Timestamp','Client Name','Address','Phone','Email',
   'Roof Type','Sq Ft','Warranty','Color Choice',
   'Line Items (JSON)','Total ($)','Status',
   'Proposal URL','DocuSign Status','Invoice URL','Notes'],
];

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

// ── Main provisioning function ────────────────────────────────

async function provision(authClient, { companyName, companyPhone, companyEmail, companyLicense, companyTagline, logoBase64 }) {
  const drive   = google.drive({ version: 'v3', auth: authClient });
  const sheets  = google.sheets({ version: 'v4', auth: authClient });
  const script  = google.script({ version: 'v1', auth: authClient });
  const co      = companyName.trim();
  const results = { companyName: co };

  log(`Provisioning for: ${co}`);

  // ── 1. Proposals folder ─────────────────────────────────────
  log('Creating Proposals folder…');
  const pFolder = await drive.files.create({
    requestBody: { name: `${co} — Proposals`, mimeType: 'application/vnd.google-apps.folder' },
    fields: 'id, webViewLink',
  });
  results.folderId  = pFolder.data.id;
  results.folderUrl = pFolder.data.webViewLink;

  // ── 2. Templates folder ──────────────────────────────────────
  log('Creating Templates folder…');
  const tFolder = await drive.files.create({
    requestBody: { name: `${co} — Templates`, mimeType: 'application/vnd.google-apps.folder' },
    fields: 'id, webViewLink',
  });
  results.templatesFolderId = tFolder.data.id;

  // ── 3. Logo (if provided) ────────────────────────────────────
  results.logoFileId = '';
  results.logoUrl    = '';
  if (logoBase64) {
    log('Uploading logo…');
    const base64Data = logoBase64.includes(',') ? logoBase64.split(',')[1] : logoBase64;
    const logoBuffer = Buffer.from(base64Data, 'base64');
    const logoStream = Readable.from(logoBuffer);

    const logoRes = await drive.files.create({
      requestBody: {
        name: `${co}_logo.png`,
        mimeType: 'image/png',
        parents: [results.templatesFolderId],
      },
      media: { mimeType: 'image/png', body: logoStream },
      fields: 'id',
    });
    results.logoFileId = logoRes.data.id;
    // Make logo publicly readable so GAS UrlFetchApp can access it
    await drive.permissions.create({
      fileId: results.logoFileId,
      requestBody: { role: 'reader', type: 'anyone' },
    });
    results.logoUrl = `https://drive.google.com/uc?export=view&id=${results.logoFileId}`;
  }

  // ── 4. Google Sheet ──────────────────────────────────────────
  log('Creating Google Sheet…');
  const ss = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: `${co} — Proposal System` },
      sheets: [
        { properties: { title: 'Catalog', index: 0 } },
        { properties: { title: 'Leads',   index: 1 } },
      ],
    },
  });
  results.spreadsheetId  = ss.data.spreadsheetId;
  results.spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${results.spreadsheetId}/edit`;

  // ── 5. Populate Catalog ──────────────────────────────────────
  log('Populating Catalog tab…');
  await sheets.spreadsheets.values.update({
    spreadsheetId: results.spreadsheetId,
    range: 'Catalog!A1',
    valueInputOption: 'RAW',
    requestBody: { values: CATALOG_ITEMS },
  });

  // ── 6. Populate Leads headers ────────────────────────────────
  log('Setting up Leads tab…');
  await sheets.spreadsheets.values.update({
    spreadsheetId: results.spreadsheetId,
    range: 'Leads!A1',
    valueInputOption: 'RAW',
    requestBody: { values: LEADS_HEADERS },
  });

  // ── 7. Proposal template Doc ─────────────────────────────────
  log('Creating Proposal template…');
  const propDoc = await drive.files.create({
    requestBody: {
      name: `[TEMPLATE] Roof Proposal — ${co}`,
      mimeType: 'application/vnd.google-apps.document',
      parents: [results.templatesFolderId],
    },
    media: { mimeType: 'text/plain', body: PROPOSAL_TEMPLATE },
    fields: 'id',
  });
  results.proposalTemplateId  = propDoc.data.id;
  results.proposalTemplateUrl = `https://docs.google.com/document/d/${propDoc.data.id}/edit`;

  // ── 8. Invoice template Doc ──────────────────────────────────
  log('Creating Invoice template…');
  const invDoc = await drive.files.create({
    requestBody: {
      name: `[TEMPLATE] Roof Invoice — ${co}`,
      mimeType: 'application/vnd.google-apps.document',
      parents: [results.templatesFolderId],
    },
    media: { mimeType: 'text/plain', body: INVOICE_TEMPLATE },
    fields: 'id',
  });
  results.invoiceTemplateId  = invDoc.data.id;
  results.invoiceTemplateUrl = `https://docs.google.com/document/d/${invDoc.data.id}/edit`;

  // ── 9. Generate Code.gs with bootstrap data injected ─────────
  log('Generating Code.gs…');
  const bootstrap = {
    SHEET_ID:             results.spreadsheetId,
    PROPOSALS_FOLDER_ID:  results.folderId,
    PROPOSAL_TEMPLATE_ID: results.proposalTemplateId,
    INVOICE_TEMPLATE_ID:  results.invoiceTemplateId,
    COMPANY_NAME:         co,
    COMPANY_PHONE:        (companyPhone   || '').trim(),
    COMPANY_EMAIL:        (companyEmail   || '').trim(),
    COMPANY_LICENSE:      (companyLicense || '').trim(),
    COMPANY_TAGLINE:      (companyTagline || 'Licensed & Insured').trim(),
    LOGO_FILE_ID:         results.logoFileId,
    LOGO_URL:             results.logoUrl,
  };

  const codeTemplate   = fs.readFileSync(CODE_GS_PATH, 'utf8');
  // Inject bootstrap as a JSON string literal (single placeholder, try/catch in Code.gs handles raw template)
  const codeGsContent  = codeTemplate.replace('%%BOOTSTRAP_JSON%%', JSON.stringify(bootstrap));
  const indexHtmlContent = fs.readFileSync(INDEX_HTML_PATH, 'utf8');

  // ── 10. Create Apps Script project ───────────────────────────
  log('Creating Apps Script project…');
  const project = await script.projects.create({
    requestBody: {
      title: `${co} — Proposal System`,
      parentId: results.spreadsheetId,
    },
  });
  results.scriptId  = project.data.scriptId;
  results.scriptUrl = `https://script.google.com/home/projects/${results.scriptId}/edit`;

  // ── 11. Upload Code.gs + Index.html + manifest ───────────────
  log('Uploading script files…');
  await script.projects.updateContent({
    scriptId: results.scriptId,
    requestBody: {
      files: [
        { name: 'Code',        type: 'SERVER_JS', source: codeGsContent },
        { name: 'Index',       type: 'HTML',      source: indexHtmlContent },
        { name: 'appsscript', type: 'JSON',       source: JSON.stringify(GAS_MANIFEST, null, 2) },
      ],
    },
  });

  // ── 12. Create version ───────────────────────────────────────
  log('Creating script version…');
  const version = await script.projects.versions.create({
    scriptId: results.scriptId,
    requestBody: { description: 'Provisioned by Roofing Proposal Platform' },
  });
  results.versionNumber = version.data.versionNumber;

  // ── 13. Deploy as Web App ────────────────────────────────────
  log('Deploying Web App…');
  try {
    const deployment = await script.projects.deployments.create({
      scriptId: results.scriptId,
      requestBody: {
        versionNumber: results.versionNumber,
        manifestFileName: 'appsscript',
        description: 'Web App — ' + co,
      },
    });
    results.deploymentId = deployment.data.deploymentId;
    results.webAppUrl    = `https://script.google.com/macros/s/${results.deploymentId}/exec`;
    results.deployed     = true;
  } catch (err) {
    // Deployment via API sometimes requires manual first-time authorization.
    // Fall back to showing the script URL with one-click deploy instructions.
    log('Auto-deploy skipped (' + err.message + ') — showing manual step');
    results.deployed  = false;
    results.deployErr = err.message;
  }

  log('✅ Provisioning complete');
  return results;
}

function log(msg) { console.log('  [provision] ' + msg); }

module.exports = { provision };
