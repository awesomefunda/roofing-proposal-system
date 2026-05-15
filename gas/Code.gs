// ============================================================
// ROOFING PROPOSAL SYSTEM — Google Apps Script Backend
// ============================================================
// Zero-database proposal + invoice system for roofing
// contractors. Runs entirely on Google Workspace — free.
//
// PRIVACY MODEL
// ─────────────
// This script is deployed by the CONTRACTOR from their own
// Google account (Extensions → Apps Script from their Sheet).
// All Drive folders, templates, PDFs, and emails belong to
// the contractor. The developer has zero access to any data.
//
// FIRST-RUN SETUP
// ───────────────
// No IDs are hardcoded. On first open the web app shows a
// setup screen. The contractor fills in their company details
// and clicks "Set up". The script creates:
//   • Proposals folder in their Drive
//   • Proposal template Google Doc
//   • Invoice template Google Doc
// IDs are stored in Script Properties — no code changes needed.
//
// DEPLOYMENT (contractor does this once)
// ──────────────────────────────────────
//   1. Open your Google Sheet → Extensions → Apps Script
//   2. Paste Code.gs and Index.html (or use clasp push)
//   3. Deploy → New Deployment → Web App
//        Execute as: Me
//        Who has access: Anyone
//   4. Open the Web App URL → fill in company details → done
//   5. Bookmark the URL on your phone
// ============================================================

// ── Bootstrap — injected by the platform during provisioning ──
// This block is auto-generated. All values are stored into
// Script Properties on first run, then this block is never
// used again. Edit via Extensions → Apps Script →
// Project Settings → Script Properties.

// %%BOOTSTRAP_JSON%% is replaced with a JSON string by provision.js during onboarding.
// The try/catch means the raw template file is valid GAS — _BOOTSTRAP stays {} until provisioned.
let _BOOTSTRAP = {};
try { _BOOTSTRAP = JSON.parse('%%BOOTSTRAP_JSON%%'); } catch(e) {}

function _initProps() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('_init') === '1') return; // already done
  props.setProperties(_BOOTSTRAP);
  props.setProperty('_init', '1');
}

// ── Script Properties — the single source of truth ───────────

function P(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

function setProps(obj) {
  PropertiesService.getScriptProperties().setProperties(obj);
}

// ── Runtime config built from Script Properties ───────────────

function getConfig() {
  return {
    SHEET_ID:             P('SHEET_ID'),
    PROPOSAL_TEMPLATE_ID: P('PROPOSAL_TEMPLATE_ID'),
    INVOICE_TEMPLATE_ID:  P('INVOICE_TEMPLATE_ID'),
    PROPOSALS_FOLDER_ID:  P('PROPOSALS_FOLDER_ID'),
    COMPANY_NAME:         P('COMPANY_NAME')    || 'Roofing Co.',
    COMPANY_PHONE:        P('COMPANY_PHONE')   || '',
    COMPANY_EMAIL:        P('COMPANY_EMAIL')   || '',
    COMPANY_LICENSE:      P('COMPANY_LICENSE') || '',
    COMPANY_TAGLINE:      P('COMPANY_TAGLINE') || 'Licensed & Insured',
    TIMEZONE:             P('TIMEZONE')        || 'America/Los_Angeles',
    DEFAULT_WARRANTY:     P('DEFAULT_WARRANTY')|| '7-year',
    ROTTED_WOOD_RATE:     P('ROTTED_WOOD_RATE')|| '$26.00 per foot',
    ROOF_TYPES:     (P('ROOF_TYPES')     || 'Flat,Tile,Composition,Shingle,Metal,Other').split(','),
    WARRANTY_OPTIONS:(P('WARRANTY_OPTIONS')|| '5-year,7-year,10-year,Manufacturer warranty').split(','),
  };
}

function isSetupComplete() {
  return !!(P('PROPOSALS_FOLDER_ID') && P('PROPOSAL_TEMPLATE_ID') && P('INVOICE_TEMPLATE_ID'));
}

// ============================================================
// WEB APP ENTRY POINTS
// ============================================================

function doGet(e) {
  _initProps(); // runs once on first open, no-op after that
  const action = e.parameter.action;

  if (action === 'catalog') return serveCatalog();
  if (action === 'config')  return serveConfig();

  const cfg = getConfig();

  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.config = JSON.stringify({
    setupMode:      !isSetupComplete(),
    companyName:    cfg.COMPANY_NAME,
    companyPhone:   cfg.COMPANY_PHONE,
    companyEmail:   cfg.COMPANY_EMAIL,
    companyLicense: cfg.COMPANY_LICENSE,
    roofTypes:      cfg.ROOF_TYPES,
    warrantyOptions: cfg.WARRANTY_OPTIONS,
    defaultWarranty: cfg.DEFAULT_WARRANTY,
    gasUrl:         ScriptApp.getService().getUrl(),
  });

  return tmpl.evaluate()
    .setTitle(cfg.COMPANY_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  _initProps();
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'first_setup')    return runFirstSetup(data.payload);
    if (action === 'submit_job')     return submitJob(data.payload);
    if (action === 'generate_invoice') return generateInvoiceFromRow(data.rowId);

    return jsonOut({ error: 'Unknown action: ' + action });
  } catch (err) {
    Logger.log(err.stack);
    return jsonOut({ error: err.message });
  }
}

// ============================================================
// FIRST-RUN SETUP
// Runs as the contractor (whoever deployed the script).
// Creates all Drive resources in their own account.
// ============================================================

function runFirstSetup(payload) {
  const company  = (payload.companyName || 'Roofing Co.').trim();
  const phone    = (payload.companyPhone   || '').trim();
  const email    = (payload.companyEmail   || '').trim();
  const license  = (payload.companyLicense || '').trim();
  const tagline  = (payload.companyTagline || 'Licensed & Insured').trim();

  // ── Get the Sheet ID from the bound spreadsheet ─────────────
  let sheetId = '';
  try {
    sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  } catch (e) {
    Logger.log('Could not get active spreadsheet: ' + e.message);
  }

  // ── Create Proposals folder ──────────────────────────────────
  const proposalsFolder = DriveApp.createFolder(company + ' — Proposals');

  // ── Create Templates folder ──────────────────────────────────
  const templatesFolder = DriveApp.createFolder(company + ' — Templates');

  // ── Create Proposal template ─────────────────────────────────
  const proposalDoc = DocumentApp.create('[TEMPLATE] Roof Proposal — ' + company);
  proposalDoc.getBody().setText(PROPOSAL_TEMPLATE_TEXT);
  proposalDoc.saveAndClose();
  DriveApp.getFileById(proposalDoc.getId()).moveTo(templatesFolder);

  // ── Create Invoice template ───────────────────────────────────
  const invoiceDoc = DocumentApp.create('[TEMPLATE] Roof Invoice — ' + company);
  invoiceDoc.getBody().setText(INVOICE_TEMPLATE_TEXT);
  invoiceDoc.saveAndClose();
  DriveApp.getFileById(invoiceDoc.getId()).moveTo(templatesFolder);

  // ── Save everything to Script Properties ─────────────────────
  setProps({
    SHEET_ID:             sheetId,
    PROPOSALS_FOLDER_ID:  proposalsFolder.getId(),
    PROPOSAL_TEMPLATE_ID: proposalDoc.getId(),
    INVOICE_TEMPLATE_ID:  invoiceDoc.getId(),
    COMPANY_NAME:         company,
    COMPANY_PHONE:        phone,
    COMPANY_EMAIL:        email,
    COMPANY_LICENSE:      license,
    COMPANY_TAGLINE:      tagline,
  });

  return jsonOut({
    success:             true,
    folderUrl:           proposalsFolder.getUrl(),
    proposalTemplateUrl: 'https://docs.google.com/document/d/' + proposalDoc.getId() + '/edit',
    invoiceTemplateUrl:  'https://docs.google.com/document/d/' + invoiceDoc.getId() + '/edit',
    templatesUrl:        templatesFolder.getUrl(),
  });
}

// ── Template text (merge tags replaced at runtime) ────────────

const PROPOSAL_TEMPLATE_TEXT = [
  'ESTIMATE AND PROPOSAL FOR ROOF REPLACEMENT',
  '==========================================',
  '{{COMPANY_NAME}}',
  '{{COMPANY_PHONE}}  |  {{COMPANY_EMAIL}}',
  '{{COMPANY_LICENSE}}  |  {{COMPANY_TAGLINE}}',
  '',
  'Date:    {{DATE}}',
  'Job #:   {{ROW_ID}}',
  '',
  '──────────────────────────────────────────',
  'CLIENT',
  '──────────────────────────────────────────',
  'Name:     {{CLIENT_NAME}}',
  'Address:  {{CLIENT_ADDRESS}}',
  'Phone:    {{CLIENT_PHONE}}',
  'Email:    {{CLIENT_EMAIL}}',
  '',
  '──────────────────────────────────────────',
  'PROJECT',
  '──────────────────────────────────────────',
  'Roof Type:    {{ROOF_TYPE}}',
  'Square Feet:  {{SQ_FT}}',
  'Warranty:     {{WARRANTY}}',
  'Color Choice: {{COLOR_CHOICE}}',
  '',
  '──────────────────────────────────────────',
  'SCOPE OF WORK',
  '──────────────────────────────────────────',
  '{{LINE_ITEMS}}',
  '',
  '══════════════════════════════════════════',
  'TOTAL:  {{TOTAL}}',
  '══════════════════════════════════════════',
  '',
  'Additional rotted wood will be charged at {{ROTTED_WOOD_RATE}} upon homeowner approval.',
  'The work area will be kept clean at the end of every work day.',
  '',
  'Homeowner signature:  ___________________________  Date: __________',
  '',
  '{{NOTES}}',
].join('\n');

const INVOICE_TEMPLATE_TEXT = [
  'INVOICE',
  '=======',
  '{{COMPANY_NAME}}',
  '{{COMPANY_PHONE}}  |  {{COMPANY_EMAIL}}',
  '{{COMPANY_LICENSE}}  |  {{COMPANY_TAGLINE}}',
  '',
  'Invoice Date:  {{DATE}}',
  'Invoice #:     {{ROW_ID}}',
  '',
  '──────────────────────────────────────────',
  'BILL TO',
  '──────────────────────────────────────────',
  '{{CLIENT_NAME}}',
  '{{CLIENT_ADDRESS}}',
  '{{CLIENT_PHONE}}',
  '{{CLIENT_EMAIL}}',
  '',
  '──────────────────────────────────────────',
  'SERVICES RENDERED',
  '──────────────────────────────────────────',
  '{{LINE_ITEMS}}',
  '',
  '══════════════════════════════════════════',
  'TOTAL DUE:  {{TOTAL}}',
  '══════════════════════════════════════════',
  '',
  'Payment is due within 30 days of invoice date.',
  'Please make checks payable to {{COMPANY_NAME}}.',
  '',
  'Thank you for your business!',
  '',
  '{{NOTES}}',
].join('\n');

// ============================================================
// CONFIG + CATALOG ENDPOINTS
// ============================================================

function serveConfig() {
  const cfg = getConfig();
  return jsonOut({
    companyName:     cfg.COMPANY_NAME,
    companyPhone:    cfg.COMPANY_PHONE,
    companyLicense:  cfg.COMPANY_LICENSE,
    roofTypes:       cfg.ROOF_TYPES,
    warrantyOptions: cfg.WARRANTY_OPTIONS,
  });
}

function serveCatalog() {
  const cfg   = getConfig();
  const sheet = getSheet('Catalog', cfg);
  const rows  = sheet.getDataRange().getValues();

  const items = rows.slice(1)
    .filter(r => String(r[5]).toUpperCase() === 'YES' || String(r[5]).toUpperCase() === 'Y')
    .map(r => ({
      item:        String(r[0]),
      description: String(r[1]),
      price:       Number(r[2]) || 0,
      unit:        String(r[3]),
      category:    String(r[4]),
    }));

  return jsonOut({ items });
}

// ============================================================
// SUBMIT JOB
// ============================================================

function submitJob(payload) {
  const cfg = getConfig();

  const rowId  = logToLeads(payload, cfg);
  const { pdfBlob, fileName } = generatePDF(payload, rowId, 'proposal', cfg);

  const folder  = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  const file    = folder.createFile(pdfBlob.setName(fileName));
  const fileUrl = file.getUrl();

  getSheet('Leads', cfg).getRange(rowId, 12).setValue(fileUrl);

  if (payload.clientEmail) sendProposalEmail(payload, pdfBlob, fileName, cfg);
  sendOwnerNotification(payload, fileUrl, rowId, cfg);

  return jsonOut({ success: true, rowId, fileUrl });
}

// ============================================================
// LOG TO LEADS SHEET
// ============================================================

function logToLeads(p, cfg) {
  const sheet = getSheet('Leads', cfg);
  sheet.appendRow([
    new Date(),
    p.clientName,
    p.clientAddress,
    p.clientPhone      || '',
    p.clientEmail      || '',
    p.roofType         || '',
    p.sqft             || '',
    p.warranty         || cfg.DEFAULT_WARRANTY,
    p.colorChoice      || '',
    JSON.stringify(p.lineItems),
    p.total,
    'Proposal Sent',
    '',   // Proposal URL — filled after PDF saved
    'Pending',
    '',   // Invoice URL
    p.notes || '',
  ]);
  return sheet.getLastRow();
}

// ============================================================
// PDF GENERATION
// ============================================================

function generatePDF(p, rowId, type, cfg) {
  const templateId = (type === 'invoice')
    ? cfg.INVOICE_TEMPLATE_ID
    : cfg.PROPOSAL_TEMPLATE_ID;

  const folder   = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  const copyName = 'DRAFT_' + type + '_' + p.clientName + '_' + rowId;
  const copy     = DriveApp.getFileById(templateId).makeCopy(copyName, folder);
  const doc      = DocumentApp.openById(copy.getId());
  const body     = doc.getBody();

  const tags = {
    '{{COMPANY_NAME}}':     cfg.COMPANY_NAME,
    '{{COMPANY_PHONE}}':    cfg.COMPANY_PHONE,
    '{{COMPANY_EMAIL}}':    cfg.COMPANY_EMAIL,
    '{{COMPANY_LICENSE}}':  cfg.COMPANY_LICENSE,
    '{{COMPANY_TAGLINE}}':  cfg.COMPANY_TAGLINE,
    '{{CLIENT_NAME}}':      p.clientName,
    '{{CLIENT_ADDRESS}}':   p.clientAddress,
    '{{CLIENT_EMAIL}}':     p.clientEmail   || '',
    '{{CLIENT_PHONE}}':     p.clientPhone   || '',
    '{{DATE}}':             Utilities.formatDate(new Date(), cfg.TIMEZONE, 'MMMM dd, yyyy'),
    '{{ROOF_TYPE}}':        p.roofType      || '',
    '{{SQ_FT}}':            p.sqft ? Number(p.sqft).toLocaleString() : '',
    '{{WARRANTY}}':         p.warranty      || cfg.DEFAULT_WARRANTY,
    '{{COLOR_CHOICE}}':     p.colorChoice   || '_______________',
    '{{TOTAL}}':            '$' + Number(p.total).toLocaleString(),
    '{{ROW_ID}}':           String(rowId),
    '{{ROTTED_WOOD_RATE}}': cfg.ROTTED_WOOD_RATE,
    '{{NOTES}}':            p.notes         || '',
  };

  Object.entries(tags).forEach(([tag, val]) => body.replaceText(tag, val));

  const lineItems = p.lineItems.map((item, i) => {
    const price = item.price ? ' — $' + Number(item.price).toLocaleString() : '';
    return (i + 1) + '. ' + item.description + price;
  }).join('\n');
  body.replaceText('{{LINE_ITEMS}}', lineItems);

  doc.saveAndClose();

  const pdfBlob  = copy.getAs('application/pdf');
  copy.setTrashed(true);

  const safeName  = p.clientName.replace(/[^a-zA-Z0-9]/g, '_');
  const dateStr   = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'yyyy-MM-dd');
  const typeLabel = (type === 'invoice') ? 'Invoice' : 'Proposal';
  const fileName  = cfg.COMPANY_NAME.replace(/\s/g, '_') + '_' + typeLabel + '_' + safeName + '_' + dateStr + '.pdf';

  return { pdfBlob, fileName };
}

// ============================================================
// EMAIL TO CLIENT
// ============================================================

function sendProposalEmail(p, pdfBlob, fileName, cfg) {
  const subject = 'Your Roof Replacement Proposal — ' + cfg.COMPANY_NAME;
  const html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">'
    + '<div style="background:#1a3a6b;padding:24px;text-align:center;">'
    + '<h1 style="color:white;margin:0;font-size:22px;">' + cfg.COMPANY_NAME + '</h1>'
    + '<p style="color:#aac4ff;margin:6px 0 0;">' + cfg.COMPANY_TAGLINE + (cfg.COMPANY_LICENSE ? ' · ' + cfg.COMPANY_LICENSE : '') + '</p>'
    + '</div>'
    + '<div style="padding:24px;background:#f9f9f9;">'
    + '<p>Dear ' + p.clientName + ',</p>'
    + '<p>Thank you for the opportunity. Please find your proposal attached.</p>'
    + '<div style="background:white;border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin:20px 0;">'
    + '<p style="margin:0;font-size:12px;color:#666;text-transform:uppercase;">Project Address</p>'
    + '<p style="margin:4px 0 14px;font-weight:bold;">' + p.clientAddress + '</p>'
    + '<p style="margin:0;font-size:12px;color:#666;text-transform:uppercase;">Total Estimate</p>'
    + '<p style="margin:4px 0 0;font-weight:800;font-size:26px;color:#1a3a6b;">$' + Number(p.total).toLocaleString() + '</p>'
    + '</div>'
    + '<p>' + cfg.COMPANY_NAME + ' · ' + cfg.COMPANY_PHONE + '</p>'
    + '</div></div>';

  GmailApp.sendEmail(p.clientEmail, subject,
    'Dear ' + p.clientName + ', please find your roof replacement proposal attached.',
    { htmlBody: html, attachments: [pdfBlob.setName(fileName)], name: cfg.COMPANY_NAME, replyTo: cfg.COMPANY_EMAIL }
  );
}

// ============================================================
// NOTIFY THE CONTRACTOR
// ============================================================

function sendOwnerNotification(p, fileUrl, rowId, cfg) {
  if (!cfg.COMPANY_EMAIL) return;
  const subject = '✅ New Proposal #' + rowId + ' — ' + p.clientName + ' ($' + Number(p.total).toLocaleString() + ')';
  const body = [
    'New proposal created — Job #' + rowId,
    '',
    'Client:   ' + p.clientName,
    'Address:  ' + p.clientAddress,
    'Phone:    ' + (p.clientPhone  || 'N/A'),
    'Email:    ' + (p.clientEmail  || 'NOT PROVIDED'),
    'Roof:     ' + (p.roofType || 'N/A') + (p.sqft ? ' · ' + p.sqft + ' sqft' : ''),
    'Warranty: ' + (p.warranty || cfg.DEFAULT_WARRANTY),
    'Total:    $' + Number(p.total).toLocaleString(),
    '',
    'Drive:    ' + fileUrl,
    '',
    p.notes ? 'Notes: ' + p.notes : '',
  ].join('\n');

  GmailApp.sendEmail(cfg.COMPANY_EMAIL, subject, body);
}

// ============================================================
// GENERATE INVOICE FROM LEADS ROW
// ============================================================

function generateInvoiceFromRow(rowId) {
  const cfg   = getConfig();
  const sheet = getSheet('Leads', cfg);
  const row   = sheet.getRange(rowId, 1, 1, 16).getValues()[0];

  const payload = {
    clientName:    row[1],
    clientAddress: row[2],
    clientPhone:   row[3],
    clientEmail:   row[4],
    roofType:      row[5],
    sqft:          row[6],
    warranty:      row[7],
    colorChoice:   row[8],
    lineItems:     JSON.parse(row[9] || '[]'),
    total:         row[10],
    notes:         row[15],
  };

  const { pdfBlob, fileName } = generatePDF(payload, rowId, 'invoice', cfg);
  const folder   = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  const file     = folder.createFile(pdfBlob.setName(fileName));
  const fileUrl  = file.getUrl();

  sheet.getRange(rowId, 15).setValue(fileUrl);
  sheet.getRange(rowId, 12).setValue('Invoiced');

  return jsonOut({ success: true, invoiceUrl: fileUrl });
}

// ============================================================
// SHEETS CUSTOM MENU
// ============================================================

function onOpen() {
  const cfg  = getConfig();
  const name = cfg.COMPANY_NAME || 'Roofing';
  SpreadsheetApp.getUi()
    .createMenu('🏠 ' + name)
    .addItem('Generate Invoice',  'invoiceSelectedRow')
    .addItem('Resend Proposal',   'resendProposalForRow')
    .addSeparator()
    .addItem('Open Web App',      'openWebApp')
    .addToUi();
}

function invoiceSelectedRow() {
  const sheet  = SpreadsheetApp.getActiveSheet();
  const rowId  = sheet.getActiveCell().getRow();
  if (rowId < 2) { SpreadsheetApp.getUi().alert('Select a data row in the Leads sheet first.'); return; }
  const result = generateInvoiceFromRow(rowId);
  const data   = JSON.parse(result.getContent());
  SpreadsheetApp.getUi().alert(data.success
    ? '✅ Invoice created:\n' + data.invoiceUrl
    : '❌ Error: ' + data.error);
}

function resendProposalForRow() {
  SpreadsheetApp.getUi().alert('Select the Proposal URL from column L and open it to resend.');
}

function openWebApp() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('Web App URL:\n\n' + url + '\n\nCopy this and bookmark it on your phone.');
}

// ============================================================
// UTILITIES
// ============================================================

function getSheet(name, cfg) {
  const id = (cfg && cfg.SHEET_ID) ? cfg.SHEET_ID : P('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id).getSheetByName(name);
  // Fallback: container-bound spreadsheet
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
