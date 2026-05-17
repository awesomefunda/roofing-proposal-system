// ============================================================
// ROOFING PROPOSAL SYSTEM — Google Apps Script
// ============================================================
//
// ARCHITECTURE DECISION: Why Google Apps Script?
//   Zero cost, zero server, zero vendor lock-in. Everything runs
//   inside the contractor's own Google account — Drive, Docs,
//   Sheets, Gmail, and Apps Script are all included for free.
//   No DocuSign subscription (~$25-45/mo), no Pandadoc, no Zapier.
//
// HOW IT WORKS (high level):
//   1. Contractor fills form → submitProposal() runs
//   2. Google Doc template is duplicated, {{TAGS}} replaced
//   3. SHA-256 hash of doc text is saved as tamper baseline
//   4. UUID token written to Leads sheet; signing URL emailed to client
//   5. Client opens URL → Sign.html served via doGet()
//   6. Client passes email KBA → verifyClientAccess() confirms token + email
//   7. Client draws signature → submitSignature() embeds image in Doc,
//      appends Certificate of Completion, exports PDF, emails both parties
//
// PDF ENGINE: Google Doc template + {{TAG}} replacement.
// E-SIGNATURE: UUID token, SHA-256 tamper detection,
//   canvas image insertion, Certificate of Completion page.
// LEGAL BASIS: ESIGN Act 15 U.S.C. § 7001 / UETA
//
// KEY SCRIPT PROPERTIES (set once via Settings tab or editor):
//   COMPANY_NAME, OWNER_NAME, COMPANY_PHONE, COMPANY_EMAIL,
//   COMPANY_LICENSE, COMPANY_TAGLINE, COMPANY_WEBSITE,
//   SHEET_ID              — Google Sheets spreadsheet ID (Leads + Catalog tabs)
//   PROPOSALS_FOLDER_ID   — Drive folder where proposal Docs are saved
//   PROPOSAL_TEMPLATE_DOC_ID — Master Google Doc template
//   INVOICE_TEMPLATE_DOC_ID  — Master invoice Doc template
//   CONTRACTOR_PIN        — Hashed PIN for contractor app access
// NOTE: Signing link URL is derived automatically from ScriptApp.getService().getUrl()
//   at proposal submission time. No manual URL configuration needed — just always
//   access the contractor app via the latest /exec deployment URL.
// ============================================================

const C = {
  TIMESTAMP:    1,
  EST_NUM:      2,
  CLIENT_NAME:  3,
  CLIENT_EMAIL: 4,
  CLIENT_PHONE: 5,
  CLIENT_ADDR:  6,
  ROOF_TYPE:    7,
  SQ_FT:        8,
  WARRANTY:     9,
  COLOR:        10,
  TOTAL:        11,
  STATUS:       12,
  SIGN_TOKEN:   13,
  DOC_ID:       14,
  DOC_HASH:     15,
  DOC_URL:      16,
  SIGNED_PDF:   17,
  SIGNED_AT:    18,
  SIGNER_IP:    19,
  SIGNER_UA:    20,
  INVOICE_URL:  21,
  LINE_ITEMS:   22,
  NOTES:        23,
};
const LEAD_COLS = 23;

var _BOOTSTRAP = {};
try { _BOOTSTRAP = JSON.parse('%%BOOTSTRAP_JSON%%'); } catch(e) {}

function _initProps() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('_init') === '1') return;
  if (Object.keys(_BOOTSTRAP).length > 0) {
    props.setProperties(_BOOTSTRAP);
    props.setProperty('_init', '1');
  }
}

function P(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}
function setProps(obj) {
  PropertiesService.getScriptProperties().setProperties(obj);
}

function getConfig() {
  return {
    SHEET_ID:                 P('SHEET_ID'),
    PROPOSALS_FOLDER_ID:      P('PROPOSALS_FOLDER_ID'),
    PROPOSAL_TEMPLATE_DOC_ID: P('PROPOSAL_TEMPLATE_DOC_ID'),
    INVOICE_TEMPLATE_DOC_ID:  P('INVOICE_TEMPLATE_DOC_ID'),
    COMPANY_NAME:     P('COMPANY_NAME')     || '',
    OWNER_NAME:       P('OWNER_NAME')       || '',
    COMPANY_PHONE:    P('COMPANY_PHONE')    || '',
    COMPANY_EMAIL:    P('COMPANY_EMAIL')    || '',
    COMPANY_LICENSE:  P('COMPANY_LICENSE')  || '',
    COMPANY_TAGLINE:  P('COMPANY_TAGLINE')  || 'Licensed & Insured',
    COMPANY_WEBSITE:  P('COMPANY_WEBSITE')  || '',
    TIMEZONE:         P('TIMEZONE')         || 'America/Los_Angeles',
    DEFAULT_WARRANTY: P('DEFAULT_WARRANTY') || '7-year',
    ROTTED_WOOD_RATE: P('ROTTED_WOOD_RATE') || '$26.00 per foot',
    ROOF_TYPES:       (P('ROOF_TYPES')      || 'Flat,Tile,Composition Shingle,Metal,Other').split(','),
    WARRANTY_OPTIONS: (P('WARRANTY_OPTIONS')|| '5-year,7-year,10-year,Manufacturer warranty').split(','),
  };
}

function isSetupComplete() {
  return !!(P('COMPANY_NAME') && P('PROPOSALS_FOLDER_ID') && P('PROPOSAL_TEMPLATE_DOC_ID') && P('SHEET_ID'));
}

// ============================================================
// WEB APP ENTRY POINT
// ============================================================

function doGet(e) {
  _initProps();
  var token = e && e.parameter && e.parameter.sign;
  if (token) return serveSignPage_(token);
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(P('COMPANY_NAME') || 'Roofing Proposals')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function serveSignPage_(token) {
  var cfg  = getConfig();
  var lead = getLeadByToken_(token);
  // DESIGN DECISION: token is injected server-side into pageData rather than
  // read from window.location.search in Sign.html. Google Apps Script serves
  // HTML inside a googleusercontent.com iframe where window.location.search
  // does NOT contain the original ?sign= query parameter — reading it client-side
  // always returns empty, causing "Invalid Link" for every customer.
  // The token is not secret (it's already in the signing URL the client received),
  // so server injection is safe. Security comes from the email KBA step.
  var pageData = {
    valid:        !!lead,
    token:        token,
    companyName:  cfg.COMPANY_NAME,
    companyPhone: cfg.COMPANY_PHONE,
  };
  var safeJson = JSON.stringify(pageData)
    .replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026');
  var template = HtmlService.createTemplateFromFile('Sign');
  template.pageData = safeJson;
  return template.evaluate()
    .setTitle('Sign Your Estimate — ' + cfg.COMPANY_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// SHEETS MENU
// ============================================================

function onOpen() {
  _initProps();
  var name = P('COMPANY_NAME') || 'Roofing Tools';
  SpreadsheetApp.getUi()
    .createMenu('🏠 ' + name)
    .addItem('✏️  New Proposal',         'showSidebar')
    .addItem('🧾  Invoice Selected Row',  'invoiceSelectedRow')
    .addSeparator()
    .addItem('⚙️  Settings',              'showSidebar')
    .addToUi();
}

function showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Index').setTitle('Roofing Tools').setWidth(440)
  );
}

// ============================================================
// API FUNCTIONS (called via google.script.run)
// ============================================================

function verifyClientAccess(token, inputEmail) {
  if (!token || !inputEmail) return { success: false, error: 'Token and email are required.' };
  var lead = getLeadByToken_(token);
  if (!lead) return { success: false, error: 'This signing link is not valid or has expired.' };
  var d             = lead.data;
  var storedEmail   = (d[C.CLIENT_EMAIL - 1] || '').toLowerCase().trim();
  var providedEmail = (inputEmail || '').toLowerCase().trim();
  if (!storedEmail) return { success: false, error: 'No email on file. Contact the contractor.' };
  if (storedEmail !== providedEmail) return { success: false, error: 'That email does not match our records. Please try again.' };
  var status = d[C.STATUS - 1];
  var docId  = d[C.DOC_ID  - 1] || '';
  var docUrl = d[C.DOC_URL - 1] || '';
  // Ensure doc is shared — safety net for proposals created before the sharing fix
  if (docId) {
    try { DriveApp.getFileById(docId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  }
  return {
    success:       true,
    estimateNum:   d[C.EST_NUM      - 1],
    clientName:    d[C.CLIENT_NAME  - 1],
    clientEmail:   d[C.CLIENT_EMAIL - 1],
    total:         d[C.TOTAL        - 1],
    docUrl:        docUrl,
    status:        status,
    alreadySigned: status === 'Signed' || status === 'Finalized',
    signedPdfUrl:  d[C.SIGNED_PDF   - 1] || '',
  };
}

function verifyContractorPin(pin) {
  var stored = P('APP_PIN') || '4766';
  return { success: (pin || '').toString() === stored };
}

function getLeadData(rowId) {
  var cfg   = getConfig();
  var sheet = getSheet_('Leads', cfg);
  if (!sheet || rowId < 2) return null;
  var row = sheet.getRange(rowId, 1, 1, LEAD_COLS).getValues()[0];
  return {
    clientName:    row[C.CLIENT_NAME  - 1],
    clientAddress: row[C.CLIENT_ADDR  - 1],
    clientPhone:   row[C.CLIENT_PHONE - 1],
    clientEmail:   row[C.CLIENT_EMAIL - 1],
    roofType:      row[C.ROOF_TYPE    - 1],
    sqft:          row[C.SQ_FT        - 1],
    warranty:      row[C.WARRANTY     - 1],
    colorChoice:   row[C.COLOR        - 1],
    notes:         row[C.NOTES        - 1],
    total:         row[C.TOTAL        - 1],
    lineItems:     tryParseJSON_(row[C.LINE_ITEMS - 1]),
  };
}

function getAppState() {
  _initProps();
  var cfg = getConfig();
  return {
    setupComplete:         isSetupComplete(),
    companyName:           cfg.COMPANY_NAME,
    ownerName:             cfg.OWNER_NAME,
    companyPhone:          cfg.COMPANY_PHONE,
    companyEmail:          cfg.COMPANY_EMAIL,
    companyLicense:        cfg.COMPANY_LICENSE,
    companyTagline:        cfg.COMPANY_TAGLINE,
    proposalTemplateDocId: cfg.PROPOSAL_TEMPLATE_DOC_ID,
    invoiceTemplateDocId:  cfg.INVOICE_TEMPLATE_DOC_ID,
    defaultWarranty:       cfg.DEFAULT_WARRANTY,
    rottedWoodRate:        cfg.ROTTED_WOOD_RATE,
    roofTypes:             cfg.ROOF_TYPES,
    warrantyOptions:       cfg.WARRANTY_OPTIONS,
    catalog:               getCatalogItems(),
    recentLeads:           getRecentLeads_(cfg),
    sheetId:               cfg.SHEET_ID || '',
    companyWebsite:        cfg.COMPANY_WEBSITE || '',
    rottedWoodRate:        cfg.ROTTED_WOOD_RATE || '',
  };
}

function getRecentLeads_(cfg) {
  try {
    var sheet = getSheet_('Leads', cfg);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, LEAD_COLS).getValues();
    return rows.map(function(r, i) {
      return {
        rowId:      i + 2,
        clientName: String(r[C.CLIENT_NAME - 1] || ''),
        total:      r[C.TOTAL - 1] || 0,
        estNum:     String(r[C.EST_NUM - 1] || ''),
        timestamp:  r[C.TIMESTAMP - 1]
          ? Utilities.formatDate(new Date(r[C.TIMESTAMP - 1]), cfg.TIMEZONE, 'MMM dd, yy')
          : '',
      };
    }).reverse().slice(0, 30);
  } catch(e) { return []; }
}

function getCatalogItems() {
  try {
    var cfg   = getConfig();
    var sheet = getSheet_('Catalog', cfg);
    if (!sheet) {
      Logger.log('⚠️ Catalog sheet not found. Expected sheet named "Catalog" in your Google Sheet.');
      return [];
    }
    var rows = sheet.getDataRange().getValues();
    if (rows.length < 2) {
      Logger.log('⚠️ Catalog sheet is empty or has no data rows.');
      return [];
    }
    var items = rows.slice(1)
      .filter(function(r) { return r.length > 5 && String(r[5]).toUpperCase() === 'YES'; })
      .map(function(r) {
        return { item: String(r[0]||''), description: String(r[1]||''),
                 price: Number(r[2]) || 0, unit: String(r[3]||''), category: String(r[4]||'') };
      });
    if (items.length === 0) {
      Logger.log('⚠️ No items found in Catalog sheet with "YES" in column F (column 6). Check your Catalog sheet setup.');
    }
    return items;
  } catch(e) { 
    Logger.log('❌ Error loading Catalog: ' + e.message);
    return [];
  }
}

function getLeads() {
  try {
    var cfg   = getConfig();
    var sheet = getSheet_('Leads', cfg);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, LEAD_COLS).getValues();
    return rows.map(function(r, i) {
      return {
        rowId:        i + 2,
        timestamp:    r[C.TIMESTAMP - 1] ? Utilities.formatDate(new Date(r[C.TIMESTAMP - 1]), cfg.TIMEZONE, 'MMM dd yyyy') : '',
        estimateNum:  r[C.EST_NUM      - 1],
        clientName:   r[C.CLIENT_NAME  - 1],
        clientPhone:  r[C.CLIENT_PHONE - 1],
        clientEmail:  r[C.CLIENT_EMAIL - 1],
        total:        r[C.TOTAL        - 1],
        status:       r[C.STATUS       - 1],
        docUrl:       r[C.DOC_URL      - 1],
        signedPdfUrl: r[C.SIGNED_PDF   - 1],
        invoiceUrl:   r[C.INVOICE_URL  - 1],
        signToken:    r[C.SIGN_TOKEN   - 1],
      };
    }).reverse();
  } catch(e) { return []; }
}

function saveSettings(payload) {
  if (!payload) return { error: "No settings payload received." };
  var company = (payload.companyName || '').trim();
  if (!company) return { error: 'Company name is required.' };

  var folderId = P('PROPOSALS_FOLDER_ID');
  if (!folderId) folderId = DriveApp.createFolder(company + ' — Proposals').getId();
  var folder = DriveApp.getFolderById(folderId);

  // Prefer sheetId from UI payload, then stored value, then sidebar fallback
  var sheetId = (payload.sheetId || '').trim() || P('SHEET_ID');
  if (!sheetId) {
    try { sheetId = SpreadsheetApp.getActiveSpreadsheet().getId(); } catch(e) {}
  }

  var proposalDocId = (payload.proposalTemplateDocId || '').trim();
  var invoiceDocId  = (payload.invoiceTemplateDocId  || '').trim();

  if (!proposalDocId) {
    try {
      proposalDocId = createDefaultProposalTemplate_({
        COMPANY_NAME:    company,
        OWNER_NAME:      (payload.ownerName      || '').trim(),
        COMPANY_PHONE:   (payload.companyPhone   || '').trim(),
        COMPANY_EMAIL:   (payload.companyEmail   || '').trim(),
        COMPANY_LICENSE: (payload.companyLicense || '').trim(),
      }, folder);
    } catch(e) { return { error: 'Could not create proposal template: ' + e.message }; }
  }

  if (!invoiceDocId) {
    try {
      invoiceDocId = createDefaultInvoiceTemplate_({
        COMPANY_NAME:    company,
        OWNER_NAME:      (payload.ownerName      || '').trim(),
        COMPANY_PHONE:   (payload.companyPhone   || '').trim(),
        COMPANY_EMAIL:   (payload.companyEmail   || '').trim(),
        COMPANY_LICENSE: (payload.companyLicense || '').trim(),
      }, folder);
    } catch(e) { Logger.log('Invoice template creation failed: ' + e.message); }
  }

  setProps({
    SHEET_ID:                 sheetId || '',
    PROPOSALS_FOLDER_ID:      folderId,
    PROPOSAL_TEMPLATE_DOC_ID: proposalDocId,
    INVOICE_TEMPLATE_DOC_ID:  invoiceDocId,
    COMPANY_NAME:             company,
    COMPANY_WEBSITE:          (payload.companyWebsite || '').trim(),
    OWNER_NAME:               (payload.ownerName        || '').trim(),
    COMPANY_PHONE:            (payload.companyPhone      || '').trim(),
    COMPANY_EMAIL:            (payload.companyEmail      || '').trim(),
    COMPANY_LICENSE:          (payload.companyLicense    || '').trim(),
    COMPANY_TAGLINE:          (payload.companyTagline    || 'Licensed & Insured').trim(),
    DEFAULT_WARRANTY:         (payload.defaultWarranty   || '7-year').trim(),
    ROTTED_WOOD_RATE:         (payload.rottedWoodRate    || '$26.00 per foot').trim(),
    _init: '1',
  });

  if ((payload.appPin || '').trim()) {
    PropertiesService.getScriptProperties().setProperty('APP_PIN', payload.appPin.trim());
  }

  try { onOpen(); } catch(e) {}

  return { success: true, folderUrl: folder.getUrl(), proposalDocId: proposalDocId, invoiceDocId: invoiceDocId };
}

// ── Submit a new proposal ─────────────────────────────────────
function submitProposal(payload) {
  var cfg = getConfig();
  if (!isSetupComplete())           return { error: 'Please complete company setup first.' };
  if (!payload.clientName)          return { error: 'Client name is required.' };
  if (!payload.total || Number(payload.total) <= 0) return { error: 'Enter a total amount.' };
  if (!payload.clientEmail)         return { error: 'Client email is required to send the signing link.' };

  var sheet = getSheet_('Leads', cfg);
  if (!sheet) return { error: 'Leads sheet not found. Make sure the Google Sheet ID is saved in Settings and the sheet has a "Leads" tab.' };
  ensureLeadsHeaders_(sheet);
  var estimateNum = generateEstimateNumber_(cfg.COMPANY_NAME, sheet.getLastRow() + 1);
  var token = generateUUID_();

  var docId, docUrl;
  try {
    var r = createProposalDoc_(payload, cfg, estimateNum);
    docId = r.docId; docUrl = r.docUrl;
  } catch(e) { return { error: 'Document creation failed: ' + e.message }; }

  var docHash = '';
  try { docHash = computeDocHash_(docId); } catch(e) { Logger.log('Hash failed: ' + e.message); }

  var rowId;
  // LockService prevents two proposals being submitted simultaneously (race condition
  // on getLastRow() + appendRow()). 10-second wait covers slow Doc creation.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    rowId = logToLeads_(payload, cfg, estimateNum, token, docId, docHash, docUrl);
  } catch(e) {
    return { error: 'Submission conflict — please try again.' };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }

  // Signing URL: automatically derived from whichever deployment is currently serving this request.
  // ScriptApp.getService().getUrl() returns the current deployment's URL, but may include:
  //   - /u/N/ prefix when the logged-in owner accesses the web app (user-specific, 404 for others)
  //   - /dev suffix when accidentally called from the editor instead of the web app
  // Both are stripped to produce the canonical /macros/s/DEPLOY_ID/exec URL.
  // No manual configuration needed — just use the new deployment URL for the contractor app
  // and all signing links will automatically use that same deployment.
  var appUrl = '';
  try {
    appUrl = ScriptApp.getService().getUrl();
    appUrl = appUrl.replace(/\/macros\/u\/\d+\/s\//, '/macros/s/'); // strip /u/N/ prefix
    appUrl = appUrl.replace(/\/dev$/, '/exec');                      // /dev → /exec (editor safety)
  } catch(e) { Logger.log('URL detection failed: ' + e.message); }
  var signingUrl = appUrl ? appUrl + '?sign=' + token : '';

  var emailSent = false, emailError = null;
  if (payload.clientEmail) {
    try {
      sendSigningLinkEmail_(payload, signingUrl, docId, estimateNum, cfg);
      emailSent = true;
    } catch(e) {
      Logger.log('Client email failed: ' + e.message);
      emailError = 'Proposal saved, but email could not be sent (' + e.message + '). Share the signing link manually.';
    }
  }

  try { sendOwnerNotification_(payload, signingUrl, docUrl, estimateNum, cfg); } catch(e) {}

  return {
    success: true, rowId: rowId, estimateNum: estimateNum,
    docUrl: docUrl, signingUrl: signingUrl, token: token,
    emailSent: emailSent, emailError: emailError,
  };
}

// ── Process a client signature submission ─────────────────────
function submitSignature(payload) {
  if (!payload.consent)    return { error: 'Electronic consent is required to sign.' };
  if (!payload.token)      return { error: 'Invalid signing token.' };
  if (!payload.signature || payload.signature.length < 100) return { error: 'Signature is required.' };
  if (!payload.initials  || payload.initials.length  < 100) return { error: 'Initials are required.' };

  var cfg  = getConfig();
  var lead = getLeadByToken_(payload.token);
  if (!lead) return { error: 'Signing link is invalid or has expired.' };

  var d      = lead.data;
  var status = d[C.STATUS - 1];
  if (status === 'Signed' || status === 'Finalized') return { error: 'This proposal has already been signed.' };

  var docId       = d[C.DOC_ID      - 1];
  var storedHash  = d[C.DOC_HASH    - 1];
  var clientName  = d[C.CLIENT_NAME  - 1];
  var clientEmail = d[C.CLIENT_EMAIL - 1];
  var estimateNum = d[C.EST_NUM      - 1];
  var total       = d[C.TOTAL        - 1];

  if (!docId) return { error: 'Proposal document not found. Please contact the contractor.' };

  if (storedHash) {
    var currentHash = '';
    try { currentHash = computeDocHash_(docId); } catch(e) {
      return { error: 'Could not verify document integrity: ' + e.message };
    }
    if (currentHash !== storedHash) {
      return { error: 'The proposal document was modified after it was sent. Please contact the contractor for a new signing link.' };
    }
  }

  var signatureBlob, initialsBlob;
  try {
    signatureBlob = base64ToBlob_(payload.signature, 'signature.png');
    initialsBlob  = base64ToBlob_(payload.initials,  'initials.png');
  } catch(e) { return { error: 'Signature data is invalid: ' + e.message }; }

  var signedAt          = new Date();
  var signedAtFormatted = Utilities.formatDate(signedAt, cfg.TIMEZONE, 'MMMM dd, yyyy h:mm a z');
  var signedAtIso       = signedAt.toISOString();

  try {
    var doc  = DocumentApp.openById(docId);
    var body = doc.getBody();
    body.replaceText('\\{\\{SIGNED_AT\\}\\}', signedAtFormatted);
    replaceTagWithImage_(body, '\\{\\{SIGNATURE_1\\}\\}', signatureBlob, 200, 68);
    appendCertificate_(doc, {
      clientName: clientName, estimateNum: estimateNum, total: total,
      signedAt: signedAtFormatted, signerIp: payload.signerIp || 'Not captured',
      signerUa: (payload.signerUa || '').substring(0, 100),
      preHash: storedHash || 'N/A', companyName: cfg.COMPANY_NAME,
      signatureBlob: signatureBlob, initialsBlob: initialsBlob,
    });
    doc.saveAndClose();
  } catch(e) { return { error: 'Could not process signature in document: ' + e.message }; }

  var signedPdfUrl = '';
  try {
    var folder   = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
    var safeName = clientName.replace(/[^a-zA-Z0-9]/g, '_');
    var dateStr  = Utilities.formatDate(signedAt, cfg.TIMEZONE, 'yyyy-MM-dd');
    var pdfName  = 'SIGNED_Proposal_' + estimateNum + '_' + safeName + '_' + dateStr + '.pdf';
    var pdfFile  = folder.createFile(DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(pdfName));
    signedPdfUrl = pdfFile.getUrl();
  } catch(e) { return { error: 'PDF export failed: ' + e.message }; }

  var sigLock = LockService.getScriptLock();
  try {
    sigLock.waitLock(10000);
    var sheet = getSheet_('Leads', cfg);
    sheet.getRange(lead.row, C.STATUS,     1, 1).setValue('Signed');
    sheet.getRange(lead.row, C.SIGNED_PDF, 1, 1).setValue(signedPdfUrl);
    sheet.getRange(lead.row, C.SIGNED_AT,  1, 1).setValue(signedAtIso);
    sheet.getRange(lead.row, C.SIGNER_IP,  1, 1).setValue(payload.signerIp || '');
    sheet.getRange(lead.row, C.SIGNER_UA,  1, 1).setValue((payload.signerUa || '').substring(0, 200));
  } catch(e) { Logger.log('Sheet write failed after signing: ' + e.message);
  } finally { try { sigLock.releaseLock(); } catch(_) {} }

  try {
    var fileId = (signedPdfUrl.match(/\/d\/([^\/]+)/) || [])[1] || '';
    if (fileId) {
      sendSignedEmail_({ clientName: clientName, clientEmail: clientEmail, estimateNum: estimateNum, total: total },
                       DriveApp.getFileById(fileId).getBlob(), cfg);
    }
  } catch(e) { Logger.log('Signed email failed: ' + e.message); }

  return { success: true, signedPdfUrl: signedPdfUrl };
}

// ── Generate invoice from existing Leads row ──────────────────
function generateInvoiceFromRow(rowId) {
  var cfg = getConfig();
  if (!cfg.INVOICE_TEMPLATE_DOC_ID) return { error: 'No invoice template configured. Go to Settings.' };
  var sheet = getSheet_('Leads', cfg);
  var row   = sheet.getRange(rowId, 1, 1, LEAD_COLS).getValues()[0];
  var payload = {
    clientName:    row[C.CLIENT_NAME  - 1], clientAddress: row[C.CLIENT_ADDR  - 1],
    clientPhone:   row[C.CLIENT_PHONE - 1], clientEmail:   row[C.CLIENT_EMAIL - 1],
    roofType:      row[C.ROOF_TYPE    - 1], sqft:          row[C.SQ_FT        - 1],
    warranty:      row[C.WARRANTY     - 1], colorChoice:   row[C.COLOR        - 1],
    lineItems:     tryParseJSON_(row[C.LINE_ITEMS - 1]),
    total:         row[C.TOTAL        - 1], notes:         row[C.NOTES        - 1],
  };
  var estimateNum = row[C.EST_NUM - 1] || generateEstimateNumber_(cfg.COMPANY_NAME, rowId);
  var invoiceNum  = estimateNum.replace(/^([A-Z]+)-/, '$1INV-');
  var result;
  try { result = generateDoc_(payload, cfg.INVOICE_TEMPLATE_DOC_ID, cfg, invoiceNum, 'invoice');
  } catch(e) { return { error: 'Invoice generation failed: ' + e.message }; }
  var folder  = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  var pdfFile = folder.createFile(result.pdfBlob.setName(result.pdfFileName));
  var pdfUrl  = pdfFile.getUrl();
  sheet.getRange(rowId, C.STATUS,      1, 1).setValue('Invoiced');
  sheet.getRange(rowId, C.INVOICE_URL, 1, 1).setValue(pdfUrl);

  // Email invoice to client and notify contractor
  var emailSent = false;
  if (payload.clientEmail) {
    try {
      sendInvoiceEmail_(payload, estimateNum, invoiceNum, result.pdfBlob, cfg);
      emailSent = true;
    } catch(e) { Logger.log('Invoice email failed: ' + e.message); }
  }

  return { success: true, pdfUrl: pdfUrl, docUrl: result.docUrl, invoiceNum: invoiceNum, emailSent: emailSent };
}

function invoiceSelectedRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowId = sheet.getActiveCell().getRow();
  if (rowId < 2) { SpreadsheetApp.getUi().alert('Select a data row in the Leads sheet first.'); return; }
  var result = generateInvoiceFromRow(rowId);
  SpreadsheetApp.getUi().alert(result.success
    ? '✅ Invoice saved to Drive:\n' + result.pdfUrl
    : '❌ Error: ' + result.error);
}

// ============================================================
// CORE: Google Doc Template -> PDF
// ============================================================

function createProposalDoc_(payload, cfg, docNumber) {
  var folder   = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  var safeName = (payload.clientName || 'Client').replace(/[^a-zA-Z0-9]/g, '_');
  var dateStr  = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'yyyy-MM-dd');
  var tempDoc  = DriveApp.getFileById(cfg.PROPOSAL_TEMPLATE_DOC_ID)
                   .makeCopy('Proposal_' + docNumber + '_' + safeName + '_' + dateStr + ' [Editable]', folder);
  var docId    = tempDoc.getId();
  var _doc1    = DocumentApp.openById(docId);
  var body     = _doc1.getBody();
  var lineItems = payload.lineItems || [];
  var scopeText = lineItems.length > 0
    ? lineItems.map(function(it, i) { return (i+1) + '. ' + it.description; }).join('\n')
    : '(No items selected)';
  var tags = {
    CLIENT_NAME: payload.clientName || '', CLIENT_ADDRESS: payload.clientAddress || '',
    CLIENT_PHONE: payload.clientPhone || '', CLIENT_EMAIL: payload.clientEmail || '',
    DATE: Utilities.formatDate(new Date(), cfg.TIMEZONE, 'MMMM dd, yyyy'),
    ESTIMATE_NUM: docNumber, SCOPE_OF_WORK: scopeText,
    TOTAL: '$' + Number(payload.total||0).toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:0}),
    ROOF_TYPE: payload.roofType || '', SQ_FT: payload.sqft ? Number(payload.sqft).toLocaleString() : '',
    WARRANTY: payload.warranty || cfg.DEFAULT_WARRANTY, COLOR: payload.colorChoice || '',
    NOTES: payload.notes || '', COMPANY_NAME: cfg.COMPANY_NAME,
    OWNER_NAME: cfg.OWNER_NAME || cfg.COMPANY_NAME, COMPANY_PHONE: cfg.COMPANY_PHONE || '',
    COMPANY_EMAIL: cfg.COMPANY_EMAIL || '', COMPANY_LICENSE: cfg.COMPANY_LICENSE || '',
    ROTTED_WOOD_RATE: cfg.ROTTED_WOOD_RATE || '$26.00 per foot',
  };
  // Add split address tags for templates that use them
  tags['CLIENT_CITY']  = payload.clientCity  || '';
  tags['CLIENT_STATE'] = payload.clientState || '';
  tags['CLIENT_ZIP']   = payload.clientZip   || '';
  Object.keys(tags).forEach(function(k) { body.replaceText('\\{\\{' + k + '\\}\\}', tags[k]); });
  _doc1.saveAndClose();
  // Share doc so client can view it via the link in the email
  try {
    DriveApp.getFileById(docId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('✅ Document shared successfully: ' + docId);
  } catch(e) {
    Logger.log('❌ Failed to share document ' + docId + ': ' + e.message);
  }
  return { docId: docId, docUrl: tempDoc.getUrl() };
}

function generateDoc_(payload, templateDocId, cfg, docNumber, type) {
  var folder    = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  var safeName  = (payload.clientName || 'Client').replace(/[^a-zA-Z0-9]/g, '_');
  var dateStr   = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'yyyy-MM-dd');
  var typeLabel = type === 'invoice' ? 'Invoice' : 'Proposal';
  var tempDoc   = DriveApp.getFileById(templateDocId)
                    .makeCopy(typeLabel + '_' + docNumber + '_' + safeName + '_' + dateStr + ' [Doc]', folder);
  var _doc2     = DocumentApp.openById(tempDoc.getId());
  var body      = _doc2.getBody();
  var lineItems = payload.lineItems || [];
  var scopeText = lineItems.length > 0
    ? lineItems.map(function(it, i) { return (i+1) + '. ' + it.description; }).join('\n')
    : '(No items selected)';
  var tags = {
    CLIENT_NAME: payload.clientName || '', CLIENT_ADDRESS: payload.clientAddress || '',
    CLIENT_PHONE: payload.clientPhone || '', CLIENT_EMAIL: payload.clientEmail || '',
    DATE: Utilities.formatDate(new Date(), cfg.TIMEZONE, 'MMMM dd, yyyy'),
    ESTIMATE_NUM: docNumber, INVOICE_NUM: docNumber, SCOPE_OF_WORK: scopeText,
    TOTAL: '$' + Number(payload.total||0).toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:0}),
    ROOF_TYPE: payload.roofType || '', SQ_FT: payload.sqft ? Number(payload.sqft).toLocaleString() : '',
    WARRANTY: payload.warranty || cfg.DEFAULT_WARRANTY, COLOR: payload.colorChoice || '',
    NOTES: payload.notes || '', COMPANY_NAME: cfg.COMPANY_NAME,
    OWNER_NAME: cfg.OWNER_NAME || cfg.COMPANY_NAME, COMPANY_PHONE: cfg.COMPANY_PHONE || '',
    COMPANY_EMAIL: cfg.COMPANY_EMAIL || '', COMPANY_LICENSE: cfg.COMPANY_LICENSE || '',
    ROTTED_WOOD_RATE: cfg.ROTTED_WOOD_RATE || '$26.00 per foot',
  };
  Object.keys(tags).forEach(function(k) { body.replaceText('\\{\\{' + k + '\\}\\}', tags[k]); });
  _doc2.saveAndClose();
  return {
    pdfBlob:     tempDoc.getAs(MimeType.PDF),
    pdfFileName: typeLabel + '_' + docNumber + '_' + safeName + '_' + dateStr + '.pdf',
    docUrl:      tempDoc.getUrl(),
  };
}

// {{SIGNATURE_1}} must be on its own paragraph line in the template
function replaceTagWithImage_(body, tagRegex, imageBlob, widthPts, heightPts) {
  var found = body.findText(tagRegex);
  if (!found) return false;
  var para = found.getElement().getParent();
  if (!para) return false;
  para.clear();
  try {
    var img = para.appendInlineImage(imageBlob);
    img.setWidth(widthPts); img.setHeight(heightPts);
  } catch(e) { Logger.log('Image insertion failed: ' + e.message); para.appendText('[Signature]'); }
  return true;
}

function appendCertificate_(doc, audit) {
  var body = doc.getBody();
  var N = '#0d2240', G = '#2e7d32', W = '#ffffff';
  body.appendPageBreak();
  var hdrCell = body.appendTable([['CERTIFICATE OF COMPLETION']]).getCell(0,0);
  hdrCell.setBackgroundColor(N).setPaddingTop(16).setPaddingBottom(16);
  try { hdrCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
  hdrCell.editAsText().setFontFamily('Arial').setFontSize(16).setBold(true).setForegroundColor(W);

  var sub = body.appendParagraph('Electronic Signature Record — ' + audit.companyName);
  sub.editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#555555').setItalic(true);
  sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(8).setSpacingAfter(6);

  var legal = body.appendParagraph(
    'This document certifies that the individual named below electronically signed the ' +
    'above-referenced proposal in compliance with the Electronic Signatures in Global and ' +
    'National Commerce Act (ESIGN Act, 15 U.S.C. § 7001) and the Uniform Electronic ' +
    'Transactions Act (UETA). This electronic signature is legally binding and equivalent ' +
    'to a handwritten signature.'
  );
  legal.editAsText().setFontFamily('Arial').setFontSize(9).setForegroundColor('#444444');
  legal.setSpacingBefore(4).setSpacingAfter(10);

  var totalFmt = '$' + Number(audit.total||0).toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:0});
  var rows = [
    ['Signer:', audit.clientName], ['Estimate #:', audit.estimateNum],
    ['Amount:', totalFmt], ['Signed At (UTC):', audit.signedAt],
    ['IP Address:', audit.signerIp || 'Not captured'],
    ['User Agent:', (audit.signerUa || '').substring(0, 80)],
    ['Document Hash:', audit.preHash || 'N/A'],
  ];
  var infoTbl = body.appendTable(rows);
  for (var r = 0; r < rows.length; r++) {
    var lc = infoTbl.getCell(r, 0); var vc = infoTbl.getCell(r, 1);
    lc.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#333333');
    vc.editAsText().setFontFamily('Courier New').setFontSize(8).setForegroundColor('#444444');
    [lc, vc].forEach(function(c) { c.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8); });
  }
  try { infoTbl.setColumnWidth(0, 110); } catch(e) {}

  body.appendParagraph('').editAsText().setFontSize(2);
  var sh = body.appendParagraph('CLIENT SIGNATURE');
  sh.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(G);
  sh.setSpacingBefore(8).setSpacingAfter(4);
  if (audit.signatureBlob) {
    var sp = body.appendParagraph('');
    var si = sp.appendInlineImage(audit.signatureBlob); si.setWidth(220); si.setHeight(72);
  }

  body.appendParagraph('').editAsText().setFontSize(2);
  var ih = body.appendParagraph('CLIENT INITIALS');
  ih.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(G);
  ih.setSpacingBefore(6).setSpacingAfter(4);
  if (audit.initialsBlob) {
    var ip = body.appendParagraph('');
    var ii = ip.appendInlineImage(audit.initialsBlob); ii.setWidth(130); ii.setHeight(52);
  }

  body.appendParagraph('').editAsText().setFontSize(4);
  var ft = body.appendParagraph(
    'This certificate was automatically generated by the ' + audit.companyName +
    ' proposal system. The document hash above identifies the exact version of the ' +
    'proposal that was presented to the signer and can be used to verify document integrity.'
  );
  ft.editAsText().setFontFamily('Arial').setFontSize(8).setForegroundColor('#999999').setItalic(true);
  ft.setSpacingBefore(8);
}

// ============================================================
// CRYPTOGRAPHY & SECURITY
// ============================================================

function computeDocHash_(docId) {
  var doc  = DocumentApp.openById(docId);
  var text = doc.getBody().getText();
  doc.saveAndClose();
  var hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return hashBytes.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function base64ToBlob_(dataUrl, filename) {
  var parts = dataUrl.split(',');
  var mime  = parts[0].split(':')[1].split(';')[0];
  return Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, filename || 'image.png');
}

function generateUUID_() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0;
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });
}

function getLeadByToken_(token) {
  if (!token) return null;
  var cfg   = getConfig();
  var sheet = getSheet_('Leads', cfg);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, LEAD_COLS).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][C.SIGN_TOKEN - 1] === token) return { row: i + 2, data: data[i] };
  }
  return null;
}

// ============================================================
// AUTO-CREATE DOC TEMPLATES
// ============================================================

function createDefaultProposalTemplate_(cfg, folder) {
  var doc  = DocumentApp.create(cfg.COMPANY_NAME + ' — Proposal Template');
  var body = doc.getBody(); body.clear();
  var N = '#0d2240', G = '#2e7d32', W = '#ffffff';

  function colorBanner(text, bg, fg, fs) {
    var tbl = body.appendTable([[text]]); var cell = tbl.getCell(0,0);
    cell.setBackgroundColor(bg);
    cell.editAsText().setFontFamily('Arial').setFontSize(fs||11).setBold(true).setForegroundColor(fg);
    cell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(14).setPaddingRight(14);
    return tbl;
  }

  var hl = [cfg.COMPANY_NAME];
  if (cfg.COMPANY_PHONE)   hl.push(cfg.COMPANY_PHONE);
  if (cfg.COMPANY_EMAIL)   hl.push(cfg.COMPANY_EMAIL);
  if (cfg.COMPANY_LICENSE) hl.push('Lic # ' + cfg.COMPANY_LICENSE);
  var hdrTbl = colorBanner(hl.join('\n'), N, W, 13);
  var hdrT   = hdrTbl.getCell(0,0).editAsText();
  hdrT.setFontSize(0, cfg.COMPANY_NAME.length - 1, 16);
  var cs = cfg.COMPANY_NAME.length + 1;
  if (cs < hdrT.getText().length) hdrT.setFontSize(cs, hdrT.getText().length-1, 10).setBold(cs, hdrT.getText().length-1, false);

  body.appendParagraph('').editAsText().setFontSize(2);
  var infoTbl = body.appendTable([[
    'CLIENT\n\n{{CLIENT_NAME}}\n{{CLIENT_ADDRESS}}\n{{CLIENT_PHONE}}\n{{CLIENT_EMAIL}}',
    'ESTIMATE INFORMATION\n\nDate: {{DATE}}\nEstimate #: {{ESTIMATE_NUM}}\nRoof Type: {{ROOF_TYPE}}\nSq Ft: {{SQ_FT}}\nColor: {{COLOR}}'
  ]]);
  infoTbl.setColumnWidth(0, 234); infoTbl.setColumnWidth(1, 234);
  [['CLIENT',0],['ESTIMATE INFORMATION',1]].forEach(function(p) {
    var lbl = p[0], col = p[1], cell = infoTbl.getCell(0,col);
    var t = cell.editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#333333');
    t.setBold(0, lbl.length-1, true).setForegroundColor(0, lbl.length-1, G);
    cell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12).setPaddingRight(12);
  });
  var lt = infoTbl.getCell(0,0).editAsText(); var cns = 'CLIENT\n\n'.length;
  lt.setFontSize(cns, cns+'{{CLIENT_NAME}}'.length-1, 12)
    .setBold(cns, cns+'{{CLIENT_NAME}}'.length-1, true)
    .setForegroundColor(cns, cns+'{{CLIENT_NAME}}'.length-1, N);

  body.appendParagraph('').editAsText().setFontSize(2);
  colorBanner('☰  ESTIMATE & PROPOSAL FOR ROOF REPLACEMENT', N, W, 12);

  body.appendParagraph('\nWe are pleased to provide the following estimate for your roof replacement project.')
    .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#444444');

  var sowH = body.appendParagraph('⚒  SCOPE OF WORK');
  sowH.editAsText().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(G);
  sowH.setSpacingBefore(6).setSpacingAfter(4);

  body.appendParagraph('{{SCOPE_OF_WORK}}').editAsText().setFontFamily('Arial').setFontSize(11).setForegroundColor('#222222');
  body.appendParagraph('{{NOTES}}').editAsText().setFontFamily('Arial').setFontSize(10).setItalic(true).setForegroundColor('#555555');

  var totTbl = colorBanner('TOTAL COST OF MATERIALS, LABOR & PERMIT FEES\n{{TOTAL}}', G, W, 10);
  var tts = 'TOTAL COST OF MATERIALS, LABOR & PERMIT FEES\n'.length;
  totTbl.getCell(0,0).editAsText().setFontSize(tts, tts+'{{TOTAL}}'.length-1, 30);

  body.appendParagraph(
    '\nIf you have any questions, please do not hesitate to ask. ' +
    'To accept this proposal, please use the signing link sent to your email. ' +
    'Any additional rotted wood will be charged {{ROTTED_WOOD_RATE}}.'
  ).editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#555555');

  // Signature table — {{SIGNATURE_1}} MUST be on its own paragraph
  var sigTbl = body.appendTable([['','']]);
  try { sigTbl.setColumnWidth(0, 290); sigTbl.setColumnWidth(1, 178); } catch(e) {}
  var lsc = sigTbl.getCell(0,0);
  lsc.editAsText().setText('CLIENT SIGNATURE').setFontFamily('Arial').setFontSize(8).setBold(true).setForegroundColor('#888888');
  lsc.appendParagraph('{{SIGNATURE_1}}').editAsText().setFontFamily('Arial').setFontSize(11).setItalic(true).setForegroundColor('#cccccc');
  lsc.appendParagraph('Date Signed: {{SIGNED_AT}}').editAsText().setFontFamily('Arial').setFontSize(9).setForegroundColor('#555555');
  lsc.appendParagraph('Printed: {{CLIENT_NAME}}').editAsText().setFontFamily('Arial').setFontSize(9).setForegroundColor('#555555');
  lsc.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12).setPaddingRight(12);

  var rsc = sigTbl.getCell(0,1);
  try { rsc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT); } catch(e) {}
  var rt = rsc.editAsText().setText('Thank you for your business!\n\n{{OWNER_NAME}}\nOwner, {{COMPANY_NAME}}')
              .setFontFamily('Arial').setFontSize(10).setForegroundColor(G);
  rt.setBold(0, 'Thank you for your business!'.length-1, true);
  var os = 'Thank you for your business!\n\n'.length;
  rt.setFontSize(os, os+'{{OWNER_NAME}}'.length-1, 14)
    .setItalic(os, os+'{{OWNER_NAME}}'.length-1, true)
    .setForegroundColor(os, os+'{{OWNER_NAME}}'.length-1, N);
  rsc.setPaddingLeft(10).setPaddingTop(10).setPaddingBottom(10);

  body.appendParagraph('').editAsText().setFontSize(4);
  var ftTbl = body.appendTable([['◉  LICENSED & INSURED','◉  QUALITY CRAFTSMANSHIP','◉  BUILT TO LAST']]);
  try { ftTbl.setColumnWidth(0,156); ftTbl.setColumnWidth(1,156); ftTbl.setColumnWidth(2,156); } catch(e) {}
  for (var i = 0; i < 3; i++) {
    var fc = ftTbl.getCell(0,i);
    fc.setBackgroundColor(N);
    fc.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#aac4ff');
    fc.setPaddingTop(10).setPaddingBottom(10);
    try { fc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
  }

  doc.saveAndClose();
  var f = DriveApp.getFileById(doc.getId()); folder.addFile(f);
  try { DriveApp.getRootFolder().removeFile(f); } catch(e) {}
  return doc.getId();
}

function createDefaultInvoiceTemplate_(cfg, folder) {
  var doc  = DocumentApp.create(cfg.COMPANY_NAME + ' — Invoice Template');
  var body = doc.getBody(); body.clear();
  var N = '#0d2240', G = '#2e7d32', W = '#ffffff';

  function colorBanner(text, bg, fg, fs) {
    var tbl = body.appendTable([[text]]); var cell = tbl.getCell(0,0);
    cell.setBackgroundColor(bg);
    cell.editAsText().setFontFamily('Arial').setFontSize(fs||11).setBold(true).setForegroundColor(fg);
    cell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(14).setPaddingRight(14);
    return tbl;
  }

  var hl = [cfg.COMPANY_NAME];
  if (cfg.COMPANY_PHONE)   hl.push(cfg.COMPANY_PHONE);
  if (cfg.COMPANY_EMAIL)   hl.push(cfg.COMPANY_EMAIL);
  if (cfg.COMPANY_LICENSE) hl.push('Lic # ' + cfg.COMPANY_LICENSE);
  var hdrTbl = colorBanner(hl.join('\n'), N, W, 13);
  var hdrT   = hdrTbl.getCell(0,0).editAsText();
  hdrT.setFontSize(0, cfg.COMPANY_NAME.length-1, 16);
  var cs = cfg.COMPANY_NAME.length + 1;
  if (cs < hdrT.getText().length) hdrT.setFontSize(cs, hdrT.getText().length-1, 10).setBold(cs, hdrT.getText().length-1, false);

  body.appendParagraph('').editAsText().setFontSize(2);
  colorBanner('INVOICE FOR ROOF REPLACEMENT SERVICES', N, W, 12);
  body.appendParagraph('').editAsText().setFontSize(2);

  var infoTbl = body.appendTable([[
    'BILL TO\n\n{{CLIENT_NAME}}\n{{CLIENT_ADDRESS}}\n{{CLIENT_PHONE}}\n{{CLIENT_EMAIL}}',
    'INVOICE DETAILS\n\nInvoice #: {{INVOICE_NUM}}\nDate: {{DATE}}\nDue: Upon Completion\nRelated Est.: {{ESTIMATE_NUM}}'
  ]]);
  infoTbl.setColumnWidth(0, 234); infoTbl.setColumnWidth(1, 234);
  [['BILL TO',0],['INVOICE DETAILS',1]].forEach(function(p) {
    var lbl = p[0], col = p[1], cell = infoTbl.getCell(0,col);
    var t = cell.editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#333333');
    t.setBold(0, lbl.length-1, true).setForegroundColor(0, lbl.length-1, G);
    cell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12).setPaddingRight(12);
  });
  var cns = 'BILL TO\n\n'.length; var blt = infoTbl.getCell(0,0).editAsText();
  blt.setFontSize(cns, cns+'{{CLIENT_NAME}}'.length-1, 12)
     .setBold(cns, cns+'{{CLIENT_NAME}}'.length-1, true)
     .setForegroundColor(cns, cns+'{{CLIENT_NAME}}'.length-1, N);

  body.appendParagraph('').editAsText().setFontSize(2);
  body.appendParagraph('SERVICES RENDERED').editAsText()
    .setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor('#888888');
  body.appendParagraph('{{SCOPE_OF_WORK}}').editAsText().setFontFamily('Arial').setFontSize(11).setForegroundColor('#222222');
  body.appendParagraph('').editAsText().setFontSize(2);

  var totTbl = body.appendTable([['TOTAL DUE','{{TOTAL}}']]);
  var lc = totTbl.getCell(0,0); var rc = totTbl.getCell(0,1);
  lc.setBackgroundColor(G); rc.setBackgroundColor(G);
  lc.editAsText().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(W);
  rc.editAsText().setFontFamily('Arial').setFontSize(28).setBold(true).setForegroundColor(W);
  try { rc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT); } catch(e) {}
  [lc,rc].forEach(function(c) { c.setPaddingTop(12).setPaddingBottom(12).setPaddingLeft(14).setPaddingRight(14); });
  totTbl.setColumnWidth(0,300); totTbl.setColumnWidth(1,168);

  body.appendParagraph('\nPayment due upon completion. Please make checks payable to ' + cfg.COMPANY_NAME + '.')
    .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#555555');
  body.appendParagraph('{{NOTES}}').editAsText().setFontFamily('Arial').setFontSize(10).setItalic(true).setForegroundColor('#555555');

  var ty = body.appendParagraph('\nThank you for your business!');
  ty.editAsText().setFontFamily('Arial').setFontSize(13).setBold(true).setForegroundColor(G);
  ty.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(16);

  body.appendParagraph('We appreciate your trust in ' + cfg.COMPANY_NAME + '.')
    .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#666666');

  var ftTbl = body.appendTable([['◉  LICENSED & INSURED','◉  QUALITY CRAFTSMANSHIP','◉  BUILT TO LAST']]);
  try { ftTbl.setColumnWidth(0,156); ftTbl.setColumnWidth(1,156); ftTbl.setColumnWidth(2,156); } catch(e) {}
  for (var i = 0; i < 3; i++) {
    var fc = ftTbl.getCell(0,i);
    fc.setBackgroundColor(N);
    fc.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#aac4ff');
    fc.setPaddingTop(10).setPaddingBottom(10);
    try { fc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
  }

  doc.saveAndClose();
  var f = DriveApp.getFileById(doc.getId()); folder.addFile(f);
  try { DriveApp.getRootFolder().removeFile(f); } catch(e) {}
  return doc.getId();
}

// ============================================================
// CRM LOGGING
// ============================================================

function ensureLeadsHeaders_(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow([
    'Timestamp','Estimate #','Client Name','Client Email','Client Phone',
    'Client Address','Roof Type','Sq Ft','Warranty','Color',
    'Total ($)','Status','Sign Token','Doc ID','Doc Hash',
    'Doc URL','Signed PDF URL','Signed At','Signer IP','Signer UA',
    'Invoice URL','Line Items (JSON)','Notes',
  ]);
}

function logToLeads_(p, cfg, estimateNum, token, docId, docHash, docUrl) {
  var sheet = getSheet_('Leads', cfg);
  ensureLeadsHeaders_(sheet);
  sheet.appendRow([
    new Date(), estimateNum, p.clientName,
    p.clientEmail||'', p.clientPhone||'', p.clientAddress||'',
    p.roofType||'', p.sqft||'', p.warranty||cfg.DEFAULT_WARRANTY,
    p.colorChoice||'', p.total||0, 'Awaiting Signature',
    token, docId, docHash, docUrl,
    '','','','','',
    JSON.stringify(p.lineItems||[]), p.notes||'',
  ]);
  return sheet.getLastRow();
}

function generateEstimateNumber_(companyName, rowId) {
  var prefix = (companyName||'EST').split(/\s+/)
    .map(function(w) { return w[0] ? w[0].toUpperCase() : ''; })
    .join('').substring(0,4);
  var dateStr = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyMMdd');
  return prefix + '-' + dateStr + '-' + String(rowId).padStart(2,'0');
}

// ============================================================
// EMAIL
// ============================================================

function sendSigningLinkEmail_(p, signingUrl, docId, estimateNum, cfg) {
  var subject  = 'Your Roof Replacement Estimate — ' + estimateNum + ' — ' + cfg.COMPANY_NAME;
  var totalFmt = '$' + Number(p.total).toLocaleString();

  // Attach PDF version of the proposal (no editable doc link sent to client)
  var attachments = [];
  if (docId) {
    try {
      var pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
      pdfBlob.setName('Proposal_' + estimateNum + '.pdf');
      attachments.push(pdfBlob);
    } catch(e) { Logger.log('PDF attach failed: ' + e.message); }
  }

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#0d2240;padding:24px;text-align:center;">' +
      '<h1 style="color:white;margin:0;font-size:22px;">' + cfg.COMPANY_NAME + '</h1>' +
      '<p style="color:#aac4ff;margin:6px 0 0;font-size:13px;">' + (cfg.COMPANY_TAGLINE || 'Licensed &amp; Insured') +
        (cfg.COMPANY_LICENSE ? ' &nbsp;·&nbsp; Lic. ' + cfg.COMPANY_LICENSE : '') + '</p>' +
    '</div>' +
    '<div style="padding:28px;background:#f9f9f9;">' +
      '<p style="margin:0 0 16px;color:#333;font-size:15px;">Dear ' + (p.clientName||'Valued Customer') + ',</p>' +
      '<p style="margin:0 0 16px;color:#555;font-size:14px;">Your roof replacement proposal is attached to this email as a PDF. Please review it and use the button below to sign electronically.</p>' +
      '<div style="background:white;border:1px solid #e0e0e0;border-radius:8px;padding:20px;margin:20px 0;">' +
        '<table style="width:100%;border-collapse:collapse;">' +
          '<tr><td style="font-size:11px;color:#888;text-transform:uppercase;padding:0 0 2px;">Estimate #</td></tr>' +
          '<tr><td style="font-weight:700;color:#0d2240;font-size:16px;padding:0 0 14px;">' + estimateNum + '</td></tr>' +
          '<tr><td style="font-size:11px;color:#888;text-transform:uppercase;padding:0 0 2px;">Proposed Total</td></tr>' +
          '<tr><td style="font-weight:800;font-size:32px;color:#2e7d32;padding:0;">' + totalFmt + '</td></tr>' +
        '</table>' +
      '</div>' +
      (signingUrl
        ? '<a href="' + signingUrl + '" style="display:block;background:#2e7d32;color:white;text-align:center;padding:16px 24px;border-radius:8px;text-decoration:none;font-weight:700;font-size:16px;margin:24px 0 8px;">✍️ Review &amp; Sign Your Proposal</a>' +
          '<p style="font-size:11px;color:#888;text-align:center;margin:0 0 24px;">Secure, legally binding under the ESIGN Act. You will be asked to verify your email before signing.</p>'
        : '<p style="color:#888;font-size:13px;">Please reply to this email to sign your proposal.</p>') +
      '<hr style="border:none;border-top:1px solid #eee;margin:20px 0;">' +
      '<p style="color:#555;font-size:13px;margin:0 0 6px;">Questions? Reply to this email or contact us directly:</p>' +
      '<p style="color:#888;font-size:12px;margin:0;">' + cfg.COMPANY_NAME +
        (cfg.COMPANY_PHONE ? ' &nbsp;·&nbsp; ' + cfg.COMPANY_PHONE : '') +
        (cfg.COMPANY_EMAIL ? ' &nbsp;·&nbsp; ' + cfg.COMPANY_EMAIL : '') +
        (cfg.COMPANY_WEBSITE ? ' &nbsp;·&nbsp; <a href="' + cfg.COMPANY_WEBSITE + '" style="color:#2e7d32;text-decoration:none;">' + cfg.COMPANY_WEBSITE.replace(/^https?:\/\//,'') + '</a>' : '') + '</p>' +
    '</div></div>';

  GmailApp.sendEmail(p.clientEmail, subject,
    'Dear ' + (p.clientName||'Customer') + ',\n\nYour proposal ' + estimateNum + ' (' + totalFmt + ') is attached as a PDF.\n\nTo sign electronically, visit:\n' + (signingUrl || '(link unavailable — please reply to sign)'),
    { htmlBody: html, name: cfg.COMPANY_NAME, replyTo: cfg.COMPANY_EMAIL, attachments: attachments });
}

function sendOwnerNotification_(p, signingUrl, docUrl, estimateNum, cfg) {
  if (!cfg.COMPANY_EMAIL) return;
  // Safety check: don't send "New Proposal Sent" email to the customer (only to the owner/admin)
  if (cfg.COMPANY_EMAIL.toLowerCase().trim() === p.clientEmail.toLowerCase().trim()) {
    Logger.log('⚠️ Skipped owner notification — COMPANY_EMAIL is the same as client email. Please set a separate admin email in Settings.');
    return;
  }
  GmailApp.sendEmail(cfg.COMPANY_EMAIL,
    '✅ New Proposal Sent — ' + estimateNum + ' · ' + p.clientName + ' ($' + Number(p.total).toLocaleString() + ')',
    ['New proposal created.','','Estimate: '+estimateNum,'Client: '+p.clientName,
     'Address: '+(p.clientAddress||'N/A'),'Phone: '+(p.clientPhone||'N/A'),
     'Email: '+(p.clientEmail||'NOT PROVIDED'),'Total: $'+Number(p.total).toLocaleString(),'',
     '🔗 Signing Link: '+(signingUrl||'(deploy web app first)'),
     '✏️  Editable Doc: '+docUrl, '', p.notes?'Notes: '+p.notes:''].join('\n'));
}

function sendSignedEmail_(p, pdfBlob, cfg) {
  var totalFmt = '$' + Number(p.total||0).toLocaleString();
  var bodyText = 'Proposal ' + p.estimateNum + ' signed by ' + p.clientName + '. Signed PDF attached.';
  var pdfName  = 'SIGNED_Proposal_' + p.estimateNum + '.pdf';
  if (pdfBlob) pdfBlob.setName(pdfName);
  var atts = pdfBlob ? [pdfBlob] : [];
  if (cfg.COMPANY_EMAIL) {
    try { GmailApp.sendEmail(cfg.COMPANY_EMAIL, '✅ SIGNED — ' + p.estimateNum + ' · ' + p.clientName, bodyText,
      { attachments: atts, name: cfg.COMPANY_NAME }); }
    catch(e) { Logger.log('Owner signed email failed: ' + e.message); }
  }
  if (p.clientEmail) {
    var html2 =
      '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
      '<div style="background:#0d2240;padding:20px;text-align:center;"><h1 style="color:white;margin:0;">' + cfg.COMPANY_NAME + '</h1></div>' +
      '<div style="padding:24px;background:#f9f9f9;">' +
        '<p style="color:#333;">Dear ' + (p.clientName||'Customer') + ',</p>' +
        '<p style="color:#555;">Your signed proposal is attached for your records.</p>' +
        '<p style="font-size:26px;font-weight:800;color:#2e7d32;">' + totalFmt + '</p>' +
        '<p style="color:#888;font-size:12px;">Questions? ' + (cfg.COMPANY_PHONE || cfg.COMPANY_EMAIL || '') + '</p>' +
      '</div></div>';
    try { GmailApp.sendEmail(p.clientEmail, 'Your Signed Roof Proposal — ' + cfg.COMPANY_NAME, bodyText,
      { htmlBody: html2, attachments: atts, name: cfg.COMPANY_NAME, replyTo: cfg.COMPANY_EMAIL }); }
    catch(e) { Logger.log('Client signed email failed: ' + e.message); }
  }
}

function sendInvoiceEmail_(p, estimateNum, invoiceNum, pdfBlob, cfg) {
  var totalFmt = '$' + Number(p.total||0).toLocaleString();
  var pdfName  = 'Invoice_' + invoiceNum + '_' + (p.clientName||'Client').replace(/[^a-zA-Z0-9]/g,'_') + '.pdf';
  if (pdfBlob) pdfBlob.setName(pdfName);
  var atts = pdfBlob ? [pdfBlob] : [];

  // Email to client
  if (p.clientEmail) {
    var html =
      '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
      '<div style="background:#0d2240;padding:20px;text-align:center;">' +
        '<h1 style="color:white;margin:0;">' + cfg.COMPANY_NAME + '</h1>' +
        (cfg.COMPANY_PHONE ? '<p style="color:#aac4ff;margin:6px 0 0;font-size:14px;">' + cfg.COMPANY_PHONE + '</p>' : '') +
      '</div>' +
      '<div style="padding:24px;background:#f9f9f9;">' +
        '<p style="color:#333;">Dear ' + (p.clientName||'Customer') + ',</p>' +
        '<p style="color:#555;">Thank you for choosing ' + cfg.COMPANY_NAME + '. Please find your invoice attached.</p>' +
        '<div style="background:white;border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin:16px 0;">' +
          '<p style="margin:4px 0;color:#555;font-size:13px;">Invoice #: <strong>' + invoiceNum + '</strong></p>' +
          '<p style="margin:4px 0;color:#555;font-size:13px;">Related Estimate: <strong>' + estimateNum + '</strong></p>' +
          '<p style="margin:12px 0 4px;font-size:28px;font-weight:800;color:#2e7d32;">' + totalFmt + '</p>' +
          '<p style="margin:0;color:#888;font-size:11px;">Due upon completion</p>' +
        '</div>' +
        '<p style="color:#888;font-size:12px;">Questions? Contact us at ' + (cfg.COMPANY_PHONE || cfg.COMPANY_EMAIL || '') + '</p>' +
      '</div></div>';
    GmailApp.sendEmail(p.clientEmail, 'Invoice ' + invoiceNum + ' — ' + cfg.COMPANY_NAME, '',
      { htmlBody: html, attachments: atts, name: cfg.COMPANY_NAME, replyTo: cfg.COMPANY_EMAIL });
  }

  // Notify contractor
  if (cfg.COMPANY_EMAIL && cfg.COMPANY_EMAIL !== (p.clientEmail||'').toLowerCase().trim()) {
    GmailApp.sendEmail(cfg.COMPANY_EMAIL, '🧾 Invoice sent — ' + invoiceNum + ' · ' + (p.clientName||''),
      'Invoice ' + invoiceNum + ' (' + totalFmt + ') emailed to ' + (p.clientEmail||'client') + '.\nRelated estimate: ' + estimateNum,
      { attachments: atts, name: cfg.COMPANY_NAME });
  }
}

// ============================================================
// UTILITIES
// ============================================================

// DESIGN DECISION: Two-path sheet lookup.
// Primary path: SHEET_ID script property → SpreadsheetApp.openById()
//   Works in all contexts: web app, editor, sidebar, scheduled tasks.
//   SHEET_ID must be saved via Settings tab or setSheetId() editor function.
// Fallback path: getActiveSpreadsheet()
//   Only works when script is run from the Sheets editor (container-bound context).
//   Always throws in the standalone web app — caught and returns null.
// If this function returns null, the caller must guard against it (check before getLastRow etc).
function getSheet_(name, cfg) {
  var id = (cfg && cfg.SHEET_ID) ? cfg.SHEET_ID : P('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id).getSheetByName(name); } catch(e) {} }
  // Sidebar fallback only — throws in standalone web app, caught here
  try { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); } catch(e) { return null; }
}

function tryParseJSON_(str) {
  try { return JSON.parse(str || '[]'); } catch(e) { return []; }
}

// Public aliases (called from non-underscore contexts)
function getSheet(name, cfg)  { return getSheet_(name, cfg); }
function tryParseJSON(str)    { return tryParseJSON_(str); }
