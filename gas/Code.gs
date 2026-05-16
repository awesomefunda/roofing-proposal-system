// ============================================================
// ROOFING PROPOSAL SYSTEM — Google Apps Script
// ============================================================
// PDF Engine: Google Doc template + {{TAG}} replacement.
//   Duplicates master Doc → replaces text tags → exports PDF.
//   E-Signature: UUID token → SHA-256 tamper detection →
//   canvas image insertion → Certificate of Completion page.
//
// Legal basis: ESIGN Act 15 U.S.C. § 7001 / UETA
// Sheets: Leads (23 cols) | Catalog | Script Properties
// ============================================================

// ── Column Constants (1-indexed, matches spreadsheet columns) ─
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
  STATUS:       12,  // Awaiting Signature | Signed | Invoiced | Finalized
  SIGN_TOKEN:   13,  // UUID v4 — used in signing URL ?sign=[token]
  DOC_ID:       14,  // Google Doc ID of the editable proposal copy
  DOC_HASH:     15,  // SHA-256 of doc body text (before signing) for tamper detection
  DOC_URL:      16,  // Editable Doc URL (for Omar to edit/view)
  SIGNED_PDF:   17,  // Signed PDF URL (after client signs)
  SIGNED_AT:    18,  // ISO timestamp of signing
  SIGNER_IP:    19,  // Client IP address at time of signing
  SIGNER_UA:    20,  // Client user-agent string
  INVOICE_URL:  21,  // Invoice PDF URL
  LINE_ITEMS:   22,  // JSON string of line items
  NOTES:        23,  // Proposal notes
};
const LEAD_COLS = 23;

// ── Bootstrap (injected by provisioner, ignored otherwise) ────
let _BOOTSTRAP = {};
try { _BOOTSTRAP = JSON.parse('%%BOOTSTRAP_JSON%%'); } catch(e) {}

function _initProps() {
  const props = PropertiesService.getScriptProperties();
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

// ── Config ────────────────────────────────────────────────────
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
    TIMEZONE:         P('TIMEZONE')         || 'America/Los_Angeles',
    DEFAULT_WARRANTY: P('DEFAULT_WARRANTY') || '7-year',
    ROTTED_WOOD_RATE: P('ROTTED_WOOD_RATE') || '$26.00 per foot',
    ROOF_TYPES:       (P('ROOF_TYPES')      || 'Flat,Tile,Composition Shingle,Metal,Other').split(','),
    WARRANTY_OPTIONS: (P('WARRANTY_OPTIONS')|| '5-year,7-year,10-year,Manufacturer warranty').split(','),
  };
}

function isSetupComplete() {
  return !!(P('COMPANY_NAME') && P('PROPOSALS_FOLDER_ID') && P('PROPOSAL_TEMPLATE_DOC_ID'));
}

// ============================================================
// WEB APP ENTRY POINT
// ============================================================

function doGet(e) {
  _initProps();
  const token = e && e.parameter && e.parameter.sign;
  if (token) {
    return serveSignPage_(token);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle((P('COMPANY_NAME') || 'Roofing Proposals'))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Serve Sign.html with MINIMAL injection — no sensitive proposal data upfront.
// Total, docUrl, clientName are withheld until verifyClientAccess() succeeds.
function serveSignPage_(token) {
  const cfg  = getConfig();
  const lead = getLeadByToken_(token);

  // Only non-sensitive bootstrap data in the initial page
  const pageData = {
    valid:       !!lead,
    companyName: cfg.COMPANY_NAME,
    companyPhone:cfg.COMPANY_PHONE,
    // token is NOT injected — client reads it from window.location.search
    // sensitive fields (total, docUrl, clientName) come after KBA
  };

  // Escape for safe injection into a <script> tag
  const safeJson = JSON.stringify(pageData)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  const template = HtmlService.createTemplateFromFile('Sign');
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
  const name = P('COMPANY_NAME') || 'Roofing Tools';
  SpreadsheetApp.getUi()
    .createMenu('🏠 ' + name)
    .addItem('✏️  New Proposal',         'showSidebar')
    .addItem('🧾  Invoice Selected Row',  'invoiceSelectedRow')
    .addSeparator()
    .addItem('⚙️  Settings',             'showSidebar')
    .addToUi();
}

function showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Index').setTitle('Roofing Tools').setWidth(440)
  );
}

// ============================================================
// API FUNCTIONS  (called via google.script.run)
// ============================================================

// ── Client KBA — called from Sign.html after email entry ──────
// Requires BOTH the unguessable UUID token AND the exact email on file.
// Sensitive proposal data is only returned on success.
function verifyClientAccess(token, inputEmail) {
  if (!token || !inputEmail) return { success: false, error: 'Token and email are required.' };

  const lead = getLeadByToken_(token);
  if (!lead) return { success: false, error: 'This signing link is not valid or has expired.' };

  const d             = lead.data;
  const storedEmail   = (d[C.CLIENT_EMAIL - 1] || '').toLowerCase().trim();
  const providedEmail = (inputEmail || '').toLowerCase().trim();

  if (!storedEmail) {
    return { success: false, error: 'No email address on file. Please contact the contractor.' };
  }
  if (storedEmail !== providedEmail) {
    return { success: false, error: 'That email does not match our records. Please try again.' };
  }

  const status = d[C.STATUS - 1];
  return {
    success:       true,
    estimateNum:   d[C.EST_NUM      - 1],
    clientName:    d[C.CLIENT_NAME  - 1],
    clientEmail:   d[C.CLIENT_EMAIL - 1],
    total:         d[C.TOTAL        - 1],
    docUrl:        d[C.DOC_URL      - 1],
    status:        status,
    alreadySigned: status === 'Signed' || status === 'Finalized',
    signedPdfUrl:  d[C.SIGNED_PDF   - 1] || '',
  };
}

// ── Contractor PIN — server-side only, PIN never reaches client ─
function verifyContractorPin(pin) {
  const stored = P('APP_PIN') || '4766';
  return { success: (pin || '').toString() === stored };
}

// ── Load previous estimate data for form re-hydration ─────────
// Returns historical prices from stored LINE_ITEMS JSON —
// NOT current catalog prices. This preserves the original quote.
function getLeadData(rowId) {
  const cfg   = getConfig();
  const sheet = getSheet_('Leads', cfg);
  if (!sheet || rowId < 2) return null;
  const row = sheet.getRange(rowId, 1, 1, LEAD_COLS).getValues()[0];
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
    lineItems:     tryParseJSON_(row[C.LINE_ITEMS - 1]), // historical prices
  };
}

function getAppState() {
  _initProps();
  const cfg = getConfig();
  return {
    setupComplete:           isSetupComplete(),
    companyName:             cfg.COMPANY_NAME,
    ownerName:               cfg.OWNER_NAME,
    companyPhone:            cfg.COMPANY_PHONE,
    companyEmail:            cfg.COMPANY_EMAIL,
    companyLicense:          cfg.COMPANY_LICENSE,
    companyTagline:          cfg.COMPANY_TAGLINE,
    proposalTemplateDocId:   cfg.PROPOSAL_TEMPLATE_DOC_ID,
    invoiceTemplateDocId:    cfg.INVOICE_TEMPLATE_DOC_ID,
    defaultWarranty:         cfg.DEFAULT_WARRANTY,
    rottedWoodRate:          cfg.ROTTED_WOOD_RATE,
    roofTypes:               cfg.ROOF_TYPES,
    warrantyOptions:         cfg.WARRANTY_OPTIONS,
    catalog:                 getCatalogItems(),
    recentLeads:             getRecentLeads_(cfg),  // for "Load Previous" dropdown
  };
}

// Returns simplified lead summaries for the Load Previous dropdown.
// Does NOT include sensitive signing tokens or document hashes.
function getRecentLeads_(cfg) {
  try {
    const sheet = getSheet_('Leads', cfg);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, LEAD_COLS).getValues();
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
    const cfg   = getConfig();
    const sheet = getSheet_('Catalog', cfg);
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues();
    return rows.slice(1)
      .filter(r => String(r[5]).toUpperCase() === 'YES')
      .map(r => ({
        item:        String(r[0]),
        description: String(r[1]),
        price:       Number(r[2]) || 0,
        unit:        String(r[3]),
        category:    String(r[4]),
      }));
  } catch(e) { return []; }
}

function getLeads() {
  try {
    const cfg   = getConfig();
    const sheet = getSheet_('Leads', cfg);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, LEAD_COLS).getValues();
    return rows.map(function(r, i) {
      return {
        rowId:        i + 2,
        timestamp:    r[C.TIMESTAMP    - 1] ? Utilities.formatDate(new Date(r[C.TIMESTAMP - 1]), cfg.TIMEZONE, 'MMM dd yyyy') : '',
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
  const company = (payload.companyName || '').trim();
  if (!company) return { error: 'Company name is required.' };

  let folderId = P('PROPOSALS_FOLDER_ID');
  if (!folderId) {
    folderId = DriveApp.createFolder(company + ' — Proposals').getId();
  }
  const folder = DriveApp.getFolderById(folderId);

  let sheetId = P('SHEET_ID');
  if (!sheetId) {
    try { sheetId = SpreadsheetApp.getActiveSpreadsheet().getId(); } catch(e) {}
  }

  let proposalDocId = (payload.proposalTemplateDocId || '').trim();
  let invoiceDocId  = (payload.invoiceTemplateDocId  || '').trim();

  if (!proposalDocId) {
    try {
      proposalDocId = createDefaultProposalTemplate_({
        COMPANY_NAME:    company,
        OWNER_NAME:      (payload.ownerName     || '').trim(),
        COMPANY_PHONE:   (payload.companyPhone  || '').trim(),
        COMPANY_EMAIL:   (payload.companyEmail  || '').trim(),
        COMPANY_LICENSE: (payload.companyLicense|| '').trim(),
      }, folder);
    } catch(e) {
      return { error: 'Could not create proposal template: ' + e.message };
    }
  }

  if (!invoiceDocId) {
    try {
      invoiceDocId = createDefaultInvoiceTemplate_({
        COMPANY_NAME:    company,
        OWNER_NAME:      (payload.ownerName     || '').trim(),
        COMPANY_PHONE:   (payload.companyPhone  || '').trim(),
        COMPANY_EMAIL:   (payload.companyEmail  || '').trim(),
        COMPANY_LICENSE: (payload.companyLicense|| '').trim(),
      }, folder);
    } catch(e) {
      Logger.log('Invoice template creation failed: ' + e.message);
    }
  }

  setProps({
    SHEET_ID:                  sheetId || '',
    PROPOSALS_FOLDER_ID:       folderId,
    PROPOSAL_TEMPLATE_DOC_ID:  proposalDocId,
    INVOICE_TEMPLATE_DOC_ID:   invoiceDocId,
    COMPANY_NAME:              company,
    OWNER_NAME:                (payload.ownerName        || '').trim(),
    COMPANY_PHONE:             (payload.companyPhone      || '').trim(),
    COMPANY_EMAIL:             (payload.companyEmail      || '').trim(),
    COMPANY_LICENSE:           (payload.companyLicense    || '').trim(),
    COMPANY_TAGLINE:           (payload.companyTagline    || 'Licensed & Insured').trim(),
    DEFAULT_WARRANTY:          (payload.defaultWarranty   || '7-year').trim(),
    ROTTED_WOOD_RATE:          (payload.rottedWoodRate    || '$26.00 per foot').trim(),
    _init: '1',
  });

  // Update PIN only if a new one was provided (blank = keep existing)
  if ((payload.appPin || '').trim()) {
    PropertiesService.getScriptProperties().setProperty('APP_PIN', payload.appPin.trim());
  }

  try { onOpen(); } catch(e) {}

  return {
    success:       true,
    folderUrl:     folder.getUrl(),
    proposalDocId: proposalDocId,
    invoiceDocId:  invoiceDocId,
  };
}

// ── Submit a new proposal ─────────────────────────────────────
function submitProposal(payload) {
  const cfg = getConfig();
  if (!isSetupComplete())           return { error: 'Please complete company setup first.' };
  if (!payload.clientName)          return { error: 'Client name is required.' };
  if (!payload.total || Number(payload.total) <= 0)
                                    return { error: 'Enter a total amount.' };
  if (!payload.clientEmail)         return { error: 'Client email is required to send the signing link.' };

  // 1. Generate estimate number (needs a placeholder row first)
  const sheet      = getSheet_('Leads', cfg);
  ensureLeadsHeaders_(sheet);
  const estimateNum = generateEstimateNumber_(cfg.COMPANY_NAME, sheet.getLastRow() + 1);

  // 2. Generate unique signing token (not guessable — never use estimate # in URL)
  const token = generateUUID_();

  // 3. Duplicate template → replace text tags → save
  let docId, docUrl;
  try {
    const result = createProposalDoc_(payload, cfg, estimateNum);
    docId  = result.docId;
    docUrl = result.docUrl;
  } catch(e) {
    return { error: 'Document creation failed: ' + e.message };
  }

  // 4. Compute SHA-256 hash of final doc text BEFORE signing
  //    (captures state shown to client — tamper detection baseline)
  let docHash = '';
  try { docHash = computeDocHash_(docId); } catch(e) { Logger.log('Hash failed: ' + e.message); }

  // 5. Log to Leads sheet — use LockService to prevent race conditions
  //    if two proposals are submitted simultaneously.
  let rowId;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    rowId = logToLeads_(payload, cfg, estimateNum, token, docId, docHash, docUrl);
  } catch(e) {
    return { error: 'Submission conflict — another proposal was being saved. Please try again.' };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }

  // 6. Build signing URL
  let appUrl = '';
  try { appUrl = ScriptApp.getService().getUrl(); } catch(e) {}
  const signingUrl = appUrl ? appUrl + '?sign=' + token : '';

  // 7. Email client the signing link. Email failure is non-fatal:
  //    the proposal is already saved and the signing link is valid.
  //    We return emailError so the UI can warn without resetting the form.
  let emailSent  = false;
  let emailError = null;
  if (payload.clientEmail) {
    try {
      sendSigningLinkEmail_(payload, signingUrl, docUrl, estimateNum, cfg);
      emailSent = true;
    } catch(e) {
      Logger.log('Client email failed: ' + e.message);
      emailError = 'Proposal saved, but email could not be sent (' + e.message + '). Share the signing link manually.';
    }
  }

  // 8. Notify Omar (best-effort — don't block on failure)
  try { sendOwnerNotification_(payload, signingUrl, docUrl, estimateNum, cfg); } catch(e) {}

  return {
    success:     true,
    rowId:       rowId,
    estimateNum: estimateNum,
    docUrl:      docUrl,
    signingUrl:  signingUrl,
    token:       token,
    emailSent:   emailSent,
    emailError:  emailError,   // non-null = saved OK but email failed
  };
}

// ── Process a client signature submission ─────────────────────
function submitSignature(payload) {
  // payload: { token, consent, initials (dataUrl), signature (dataUrl), signerIp, signerUa }

  if (!payload.consent)  return { error: 'Electronic consent is required to sign.' };
  if (!payload.token)    return { error: 'Invalid signing token.' };
  if (!payload.signature || payload.signature.length < 100)
                         return { error: 'Signature is required.' };
  if (!payload.initials  || payload.initials.length  < 100)
                         return { error: 'Initials are required.' };

  const cfg  = getConfig();
  const lead = getLeadByToken_(payload.token);
  if (!lead) return { error: 'Signing link is invalid or has expired.' };

  const d      = lead.data;
  const status = d[C.STATUS - 1];

  // Double-signing prevention
  if (status === 'Signed' || status === 'Finalized') {
    return { error: 'This proposal has already been signed.' };
  }

  const docId      = d[C.DOC_ID    - 1];
  const storedHash = d[C.DOC_HASH  - 1];
  const docUrl     = d[C.DOC_URL   - 1];
  const clientName = d[C.CLIENT_NAME- 1];
  const clientEmail= d[C.CLIENT_EMAIL-1];
  const estimateNum= d[C.EST_NUM   - 1];
  const total      = d[C.TOTAL     - 1];

  if (!docId) return { error: 'Proposal document not found. Please contact the contractor.' };

  // ── Tamper detection ──────────────────────────────────────
  if (storedHash) {
    let currentHash = '';
    try { currentHash = computeDocHash_(docId); } catch(e) {
      return { error: 'Could not verify document integrity: ' + e.message };
    }
    if (currentHash !== storedHash) {
      return { error: 'The proposal document was modified after it was sent. Please contact the contractor for a new signing link.' };
    }
  }

  // ── Convert base64 canvas images → Blobs ─────────────────
  let signatureBlob, initialsBlob;
  try {
    signatureBlob = base64ToBlob_(payload.signature, 'signature.png');
    initialsBlob  = base64ToBlob_(payload.initials,  'initials.png');
  } catch(e) {
    return { error: 'Signature data is invalid: ' + e.message };
  }

  // ── Open the editable Doc and insert signatures ───────────
  const signedAt = new Date();
  const signedAtFormatted = Utilities.formatDate(signedAt, cfg.TIMEZONE, 'MMMM dd, yyyy h:mm a z');
  const signedAtIso       = signedAt.toISOString();

  try {
    const doc  = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Replace text-based signature tags
    body.replaceText('\\{\\{SIGNED_AT\\}\\}', signedAtFormatted);

    // Insert signature image at {{SIGNATURE_1}} placeholder
    replaceTagWithImage_(body, '\\{\\{SIGNATURE_1\\}\\}', signatureBlob, 200, 68);

    // Append Certificate of Completion on a new page
    appendCertificate_(doc, {
      clientName:    clientName,
      estimateNum:   estimateNum,
      total:         total,
      signedAt:      signedAtFormatted,
      signerIp:      payload.signerIp || 'Not captured',
      signerUa:      (payload.signerUa || '').substring(0, 100),
      preHash:       storedHash || 'N/A',
      companyName:   cfg.COMPANY_NAME,
      signatureBlob: signatureBlob,
      initialsBlob:  initialsBlob,
    });

    doc.saveAndClose();
  } catch(e) {
    return { error: 'Could not process signature in document: ' + e.message };
  }

  // ── Export signed PDF ─────────────────────────────────────
  let signedPdfUrl = '';
  try {
    const folder     = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
    const safeName   = clientName.replace(/[^a-zA-Z0-9]/g, '_');
    const dateStr    = Utilities.formatDate(signedAt, cfg.TIMEZONE, 'yyyy-MM-dd');
    const pdfName    = 'SIGNED_Proposal_' + estimateNum + '_' + safeName + '_' + dateStr + '.pdf';
    const pdfBlob    = DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(pdfName);
    const pdfFile    = folder.createFile(pdfBlob);
    signedPdfUrl     = pdfFile.getUrl();
  } catch(e) {
    return { error: 'PDF export failed: ' + e.message };
  }

  // ── Update Leads row (locked to prevent concurrent overwrites) ─
  const sigLock = LockService.getScriptLock();
  try {
    sigLock.waitLock(10000);
    const sheet = getSheet_('Leads', cfg);
    // Batch update: write all signing data in one range operation
    sheet.getRange(lead.row, C.STATUS, 1, 1).setValue('Signed');
    sheet.getRange(lead.row, C.SIGNED_PDF, 1, 1).setValue(signedPdfUrl);
    sheet.getRange(lead.row, C.SIGNED_AT,  1, 1).setValue(signedAtIso);
    sheet.getRange(lead.row, C.SIGNER_IP,  1, 1).setValue(payload.signerIp || '');
    sheet.getRange(lead.row, C.SIGNER_UA,  1, 1).setValue((payload.signerUa || '').substring(0, 200));
  } catch(e) {
    Logger.log('Sheet write failed after signing: ' + e.message);
    // Non-fatal: PDF was already exported — return success with a note
  } finally {
    try { sigLock.releaseLock(); } catch(_) {}
  }

  // ── Email signed PDF to both parties (best-effort) ────────
  if (signedPdfUrl) {
    try {
      const fileId = (signedPdfUrl.match(/\/d\/([^\/]+)/) || [])[1] || '';
      const signedPdfBlob = fileId ? DriveApp.getFileById(fileId).getBlob() : null;
      if (signedPdfBlob) {
        sendSignedEmail_({ clientName, clientEmail, estimateNum, total }, signedPdfBlob, cfg);
      }
    } catch(e) {
      Logger.log('Signed email failed: ' + e.message);
    }
  }

  return { success: true, signedPdfUrl: signedPdfUrl };
}

// ── Generate invoice from existing Leads row ──────────────────
function generateInvoiceFromRow(rowId) {
  const cfg = getConfig();
  if (!cfg.INVOICE_TEMPLATE_DOC_ID) return { error: 'No invoice template configured. Go to Settings.' };

  const sheet = getSheet_('Leads', cfg);
  const row   = sheet.getRange(rowId, 1, 1, LEAD_COLS).getValues()[0];

  const payload = {
    clientName:    row[C.CLIENT_NAME  - 1],
    clientAddress: row[C.CLIENT_ADDR  - 1],
    clientPhone:   row[C.CLIENT_PHONE - 1],
    clientEmail:   row[C.CLIENT_EMAIL - 1],
    roofType:      row[C.ROOF_TYPE    - 1],
    sqft:          row[C.SQ_FT        - 1],
    warranty:      row[C.WARRANTY     - 1],
    colorChoice:   row[C.COLOR        - 1],
    lineItems:     tryParseJSON_(row[C.LINE_ITEMS - 1]),
    total:         row[C.TOTAL        - 1],
    notes:         row[C.NOTES        - 1],
  };

  const estimateNum = row[C.EST_NUM - 1] || generateEstimateNumber_(cfg.COMPANY_NAME, rowId);
  const invoiceNum  = estimateNum.replace(/^([A-Z]+)-/, '$1INV-');

  let result;
  try {
    result = generateDoc_(payload, cfg.INVOICE_TEMPLATE_DOC_ID, cfg, invoiceNum, 'invoice');
  } catch(e) {
    return { error: 'Invoice generation failed: ' + e.message };
  }

  const folder  = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  const pdfFile = folder.createFile(result.pdfBlob.setName(result.pdfFileName));
  const pdfUrl  = pdfFile.getUrl();

  sheet.getRange(rowId, C.STATUS,      1, 1).setValue('Invoiced');
  sheet.getRange(rowId, C.INVOICE_URL, 1, 1).setValue(pdfUrl);

  return { success: true, pdfUrl, docUrl: result.docUrl, invoiceNum };
}

function invoiceSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rowId = sheet.getActiveCell().getRow();
  if (rowId < 2) {
    SpreadsheetApp.getUi().alert('Select a data row in the Leads sheet first.');
    return;
  }
  const result = generateInvoiceFromRow(rowId);
  SpreadsheetApp.getUi().alert(result.success
    ? '✅ Invoice saved to Drive:\n' + result.pdfUrl
    : '❌ Error: ' + result.error);
}

// ============================================================
// CORE: Google Doc Template → PDF
// ============================================================

// Creates a proposal Doc copy, replaces text tags, keeps as editable copy.
// Does NOT replace signature image tags (those come at signing time).
function createProposalDoc_(payload, cfg, docNumber) {
  const folder    = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  const safeName  = (payload.clientName || 'Client').replace(/[^a-zA-Z0-9]/g, '_');
  const dateStr   = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'yyyy-MM-dd');
  const baseName  = 'Proposal_' + docNumber + '_' + safeName + '_' + dateStr;

  const tempDoc = DriveApp.getFileById(cfg.PROPOSAL_TEMPLATE_DOC_ID).makeCopy(baseName + ' [Editable]', folder);
  const docId   = tempDoc.getId();
  const doc     = DocumentApp.openById(docId);
  const body    = doc.getBody();

  const lineItems  = payload.lineItems || [];
  const scopeText  = lineItems.length > 0
    ? lineItems.map(function(item, i) { return (i + 1) + '. ' + item.description; }).join('\n')
    : '(No items selected)';
  const totalFmt   = '$' + Number(payload.total || 0).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  const dateFmt    = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'MMMM dd, yyyy');

  const tags = {
    CLIENT_NAME:      payload.clientName    || '',
    CLIENT_ADDRESS:   payload.clientAddress || '',
    CLIENT_PHONE:     payload.clientPhone   || '',
    CLIENT_EMAIL:     payload.clientEmail   || '',
    DATE:             dateFmt,
    ESTIMATE_NUM:     docNumber,
    SCOPE_OF_WORK:    scopeText,
    TOTAL:            totalFmt,
    ROOF_TYPE:        payload.roofType      || '',
    SQ_FT:            payload.sqft ? Number(payload.sqft).toLocaleString() : '',
    WARRANTY:         payload.warranty      || cfg.DEFAULT_WARRANTY,
    COLOR:            payload.colorChoice   || '',
    NOTES:            payload.notes         || '',
    COMPANY_NAME:     cfg.COMPANY_NAME,
    OWNER_NAME:       cfg.OWNER_NAME        || cfg.COMPANY_NAME,
    COMPANY_PHONE:    cfg.COMPANY_PHONE     || '',
    COMPANY_EMAIL:    cfg.COMPANY_EMAIL     || '',
    COMPANY_LICENSE:  cfg.COMPANY_LICENSE   || '',
    ROTTED_WOOD_RATE: cfg.ROTTED_WOOD_RATE  || '$26.00 per foot',
    // Note: {{SIGNATURE_1}}, {{INITIALS_1}}, {{SIGNED_AT}} are intentionally
    // NOT replaced here — they are handled at signing time.
  };

  Object.keys(tags).forEach(function(key) {
    body.replaceText('\\{\\{' + key + '\\}\\}', tags[key]);
  });

  doc.saveAndClose();
  return { docId: docId, docUrl: tempDoc.getUrl() };
}

// Generic Doc → PDF helper (for invoices and other types)
function generateDoc_(payload, templateDocId, cfg, docNumber, type) {
  const folder    = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  const safeName  = (payload.clientName || 'Client').replace(/[^a-zA-Z0-9]/g, '_');
  const dateStr   = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'yyyy-MM-dd');
  const typeLabel = type === 'invoice' ? 'Invoice' : 'Proposal';
  const baseName  = typeLabel + '_' + docNumber + '_' + safeName + '_' + dateStr;

  const tempDoc = DriveApp.getFileById(templateDocId).makeCopy(baseName + ' [Doc]', folder);
  const doc     = DocumentApp.openById(tempDoc.getId());
  const body    = doc.getBody();

  const lineItems = payload.lineItems || [];
  const scopeText = lineItems.length > 0
    ? lineItems.map(function(item, i) { return (i + 1) + '. ' + item.description; }).join('\n')
    : '(No items selected)';
  const totalFmt  = '$' + Number(payload.total || 0).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  const dateFmt   = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'MMMM dd, yyyy');

  const tags = {
    CLIENT_NAME:      payload.clientName    || '',
    CLIENT_ADDRESS:   payload.clientAddress || '',
    CLIENT_PHONE:     payload.clientPhone   || '',
    CLIENT_EMAIL:     payload.clientEmail   || '',
    DATE:             dateFmt,
    ESTIMATE_NUM:     docNumber,
    INVOICE_NUM:      docNumber,
    SCOPE_OF_WORK:    scopeText,
    TOTAL:            totalFmt,
    ROOF_TYPE:        payload.roofType      || '',
    SQ_FT:            payload.sqft ? Number(payload.sqft).toLocaleString() : '',
    WARRANTY:         payload.warranty      || cfg.DEFAULT_WARRANTY,
    COLOR:            payload.colorChoice   || '',
    NOTES:            payload.notes         || '',
    COMPANY_NAME:     cfg.COMPANY_NAME,
    OWNER_NAME:       cfg.OWNER_NAME        || cfg.COMPANY_NAME,
    COMPANY_PHONE:    cfg.COMPANY_PHONE     || '',
    COMPANY_EMAIL:    cfg.COMPANY_EMAIL     || '',
    COMPANY_LICENSE:  cfg.COMPANY_LICENSE   || '',
    ROTTED_WOOD_RATE: cfg.ROTTED_WOOD_RATE  || '$26.00 per foot',
  };

  Object.keys(tags).forEach(function(key) {
    body.replaceText('\\{\\{' + key + '\\}\\}', tags[key]);
  });

  doc.saveAndClose();
  return {
    pdfBlob:     tempDoc.getAs(MimeType.PDF),
    pdfFileName: baseName + '.pdf',
    docUrl:      tempDoc.getUrl(),
  };
}

// ── Insert an image at a tag placeholder paragraph ────────────
// The tag MUST be on its own line (its own Paragraph element) in the template.
// Clears the paragraph and inserts the image blob inline.
function replaceTagWithImage_(body, tagRegex, imageBlob, widthPts, heightPts) {
  var found = body.findText(tagRegex);
  if (!found) return false;
  var textEl = found.getElement();
  var para   = textEl.getParent();
  if (!para) return false;

  para.clear();
  try {
    var img = para.appendInlineImage(imageBlob);
    img.setWidth(widthPts);
    img.setHeight(heightPts);
  } catch(e) {
    Logger.log('Image insertion failed: ' + e.message);
    para.appendText('[Signature]');
  }
  return true;
}

// ── Append a Certificate of Completion page to the signed Doc ─
function appendCertificate_(doc, audit) {
  const body = doc.getBody();
  const N = '#0d2240', G = '#2e7d32', W = '#ffffff';

  body.appendPageBreak();

  // Navy header banner
  const hdrTbl  = body.appendTable([['CERTIFICATE OF COMPLETION']]);
  const hdrCell = hdrTbl.getCell(0, 0);
  hdrCell.setBackgroundColor(N);
  hdrCell.editAsText()
    .setFontFamily('Arial').setFontSize(16).setBold(true).setForegroundColor(W);
  hdrCell.setTextAlignment(DocumentApp.TextAlignment.CENTER);
  hdrCell.setPaddingTop(16).setPaddingBottom(16);

  // Subtitle
  const sub = body.appendParagraph('Electronic Signature Record — ' + audit.companyName);
  sub.editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#555555').setItalic(true);
  sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(8).setSpacingAfter(6);

  // Legal statement
  const legal = body.appendParagraph(
    'This document certifies that the individual named below electronically signed the ' +
    'above-referenced proposal in compliance with the Electronic Signatures in Global and ' +
    'National Commerce Act (ESIGN Act, 15 U.S.C. § 7001) and the Uniform Electronic ' +
    'Transactions Act (UETA). This electronic signature is legally binding and equivalent ' +
    'to a handwritten signature.'
  );
  legal.editAsText().setFontFamily('Arial').setFontSize(9).setForegroundColor('#444444');
  legal.setSpacingBefore(4).setSpacingAfter(10);

  // Audit details table
  const totalFmt = '$' + Number(audit.total || 0).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  const rows = [
    ['Signer:',           audit.clientName],
    ['Estimate #:',       audit.estimateNum],
    ['Amount:',           totalFmt],
    ['Signed At (UTC):',  audit.signedAt],
    ['IP Address:',       audit.signerIp || 'Not captured'],
    ['User Agent:',       (audit.signerUa || '').substring(0, 80)],
    ['Document Hash:',    audit.preHash || 'N/A'],
  ];

  const infoTbl = body.appendTable(rows);
  for (var r = 0; r < rows.length; r++) {
    var labelCell = infoTbl.getCell(r, 0);
    var valueCell = infoTbl.getCell(r, 1);
    labelCell.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#333333');
    valueCell.editAsText().setFontFamily('Courier New').setFontSize(8).setForegroundColor('#444444');
    [labelCell, valueCell].forEach(function(c) {
      c.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8);
    });
  }
  try { infoTbl.setColumnWidth(0, 110); } catch(e) {}

  // Signature section
  body.appendParagraph('').editAsText().setFontSize(2);
  var sigHeader = body.appendParagraph('CLIENT SIGNATURE');
  sigHeader.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(G);
  sigHeader.setSpacingBefore(8).setSpacingAfter(4);

  if (audit.signatureBlob) {
    var sigPara = body.appendParagraph('');
    var sigImg  = sigPara.appendInlineImage(audit.signatureBlob);
    sigImg.setWidth(220).setHeight(72);
  }

  // Initials section
  body.appendParagraph('').editAsText().setFontSize(2);
  var initHeader = body.appendParagraph('CLIENT INITIALS');
  initHeader.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(G);
  initHeader.setSpacingBefore(6).setSpacingAfter(4);

  if (audit.initialsBlob) {
    var initPara = body.appendParagraph('');
    var initImg  = initPara.appendInlineImage(audit.initialsBlob);
    initImg.setWidth(130).setHeight(52);
  }

  // Footer
  body.appendParagraph('').editAsText().setFontSize(4);
  var footer = body.appendParagraph(
    'This certificate was automatically generated by the ' + audit.companyName +
    ' proposal system. The document hash above identifies the exact version of the ' +
    'proposal that was presented to the signer and can be used to verify document integrity.'
  );
  footer.editAsText().setFontFamily('Arial').setFontSize(8).setForegroundColor('#999999').setItalic(true);
  footer.setSpacingBefore(8);
}

// ============================================================
// CRYPTOGRAPHY & SECURITY
// ============================================================

// SHA-256 hex digest of a Google Doc's body text.
// Used as a tamper-detection baseline: computed before sending,
// verified before signing.  If they differ, the doc was edited.
function computeDocHash_(docId) {
  var doc  = DocumentApp.openById(docId);
  var text = doc.getBody().getText();
  doc.saveAndClose();

  var hashBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text,
    Utilities.Charset.UTF_8
  );
  return hashBytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

// Convert a canvas data URL (data:image/png;base64,...) to a Blob.
function base64ToBlob_(dataUrl, filename) {
  var parts  = dataUrl.split(',');
  var mime   = parts[0].split(':')[1].split(';')[0];
  var bytes  = Utilities.base64Decode(parts[1]);
  return Utilities.newBlob(bytes, mime, filename || 'image.png');
}

// RFC 4122 UUID v4 (cryptographically random enough for signing tokens)
function generateUUID_() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0;
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });
}

// Scan Leads sheet for a row whose SIGN_TOKEN matches the given token.
// Returns { row: <1-indexed row number>, data: <row array> } or null.
function getLeadByToken_(token) {
  if (!token) return null;
  var cfg   = getConfig();
  var sheet = getSheet_('Leads', cfg);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, LEAD_COLS).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][C.SIGN_TOKEN - 1] === token) {
      return { row: i + 2, data: data[i] };
    }
  }
  return null;
}

// ============================================================
// AUTO-CREATE DOC TEMPLATES (runs once on first saveSettings)
// ============================================================

function createDefaultProposalTemplate_(cfg, folder) {
  const doc  = DocumentApp.create(cfg.COMPANY_NAME + ' — Proposal Template');
  const body = doc.getBody();
  body.clear();

  const N = '#0d2240', G = '#2e7d32', W = '#ffffff';

  function colorBanner(text, bg, fg, fontSize) {
    const tbl  = body.appendTable([[text]]);
    const cell = tbl.getCell(0, 0);
    cell.setBackgroundColor(bg);
    cell.editAsText()
      .setFontFamily('Arial').setFontSize(fontSize || 11)
      .setBold(true).setForegroundColor(fg);
    cell.setPaddingTop(10).setPaddingBottom(10)
        .setPaddingLeft(14).setPaddingRight(14);
    return tbl;
  }

  // ── Company header (navy) ─────────────────────────────────
  const headerLines = [cfg.COMPANY_NAME];
  if (cfg.COMPANY_PHONE)   headerLines.push(cfg.COMPANY_PHONE);
  if (cfg.COMPANY_EMAIL)   headerLines.push(cfg.COMPANY_EMAIL);
  if (cfg.COMPANY_LICENSE) headerLines.push('Lic # ' + cfg.COMPANY_LICENSE);
  const hdrTbl = colorBanner(headerLines.join('\n'), N, W, 13);
  const hdrT   = hdrTbl.getCell(0, 0).editAsText();
  hdrT.setFontSize(0, cfg.COMPANY_NAME.length - 1, 16);
  const cs = cfg.COMPANY_NAME.length + 1;
  if (cs < hdrT.getText().length) {
    hdrT.setFontSize(cs, hdrT.getText().length - 1, 10)
        .setBold(cs, hdrT.getText().length - 1, false);
  }

  body.appendParagraph('').editAsText().setFontSize(2);

  // ── Client + Estimate info (two columns) ─────────────────
  const infoTbl = body.appendTable([[
    'CLIENT\n\n{{CLIENT_NAME}}\n{{CLIENT_ADDRESS}}\n{{CLIENT_PHONE}}\n{{CLIENT_EMAIL}}',
    'ESTIMATE INFORMATION\n\nDate: {{DATE}}\nEstimate #: {{ESTIMATE_NUM}}\nProject: Roof Replacement\nRoof Type: {{ROOF_TYPE}}\nSq Ft: {{SQ_FT}}\nColor: {{COLOR}}'
  ]]);
  infoTbl.setColumnWidth(0, 234);
  infoTbl.setColumnWidth(1, 234);
  [['CLIENT', 0], ['ESTIMATE INFORMATION', 1]].forEach(function(pair) {
    const label = pair[0], col = pair[1];
    const cell  = infoTbl.getCell(0, col);
    const t     = cell.editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#333333');
    t.setBold(0, label.length - 1, true).setForegroundColor(0, label.length - 1, G);
    cell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12).setPaddingRight(12);
  });
  const leftT   = infoTbl.getCell(0, 0).editAsText();
  const cnStart = 'CLIENT\n\n'.length;
  leftT.setFontSize(cnStart, cnStart + '{{CLIENT_NAME}}'.length - 1, 12)
       .setBold(cnStart, cnStart + '{{CLIENT_NAME}}'.length - 1, true)
       .setForegroundColor(cnStart, cnStart + '{{CLIENT_NAME}}'.length - 1, N);

  body.appendParagraph('').editAsText().setFontSize(2);

  // ── Proposal title banner ─────────────────────────────────
  colorBanner('☰  ESTIMATE & PROPOSAL FOR ROOF REPLACEMENT', N, W, 12);

  const intro = body.appendParagraph('\nWe are pleased to provide the following estimate for your roof replacement project.');
  intro.editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#444444');
  intro.setSpacingBefore(4).setSpacingAfter(8);

  // ── Scope of Work ─────────────────────────────────────────
  const sowH = body.appendParagraph('⚒  SCOPE OF WORK');
  sowH.editAsText().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(G);
  sowH.setSpacingBefore(4).setSpacingAfter(4);

  const sow = body.appendParagraph('{{SCOPE_OF_WORK}}');
  sow.editAsText().setFontFamily('Arial').setFontSize(11).setForegroundColor('#222222');
  sow.setSpacingBefore(2).setSpacingAfter(6);

  const notesPara = body.appendParagraph('{{NOTES}}');
  notesPara.editAsText().setFontFamily('Arial').setFontSize(10).setItalic(true).setForegroundColor('#555555');
  notesPara.setSpacingAfter(10);

  // ── Total bar (green) ─────────────────────────────────────
  const totalTbl     = colorBanner('TOTAL COST OF MATERIALS, LABOR & PERMIT FEES\n{{TOTAL}}', G, W, 10);
  const totalTagStart = 'TOTAL COST OF MATERIALS, LABOR & PERMIT FEES\n'.length;
  totalTbl.getCell(0, 0).editAsText()
    .setFontSize(totalTagStart, totalTagStart + '{{TOTAL}}'.length - 1, 30);

  // ── Disclaimer ────────────────────────────────────────────
  const disc = body.appendParagraph(
    '\nIf you have any questions, please do not hesitate to ask. ' +
    'To accept this proposal, please use the signing link sent to your email. ' +
    'Any additional rotted wood will be charged {{ROTTED_WOOD_RATE}}.'
  );
  disc.editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#555555');
  disc.setSpacingBefore(10).setSpacingAfter(12);

  // ── Signature section ─────────────────────────────────────
  // IMPORTANT: {{SIGNATURE_1}} MUST remain on its own paragraph line
  // so replaceTagWithImage_() can find and replace it with the canvas image.
  const sigTbl = body.appendTable([['', '']]);
  try { sigTbl.setColumnWidth(0, 290); sigTbl.setColumnWidth(1, 178); } catch(e) {}

  // Left cell: client signature area
  const leftSigCell = sigTbl.getCell(0, 0);
  leftSigCell.editAsText()
    .setText('CLIENT SIGNATURE')
    .setFontFamily('Arial').setFontSize(8).setBold(true).setForegroundColor('#888888');
  // Signature placeholder — this ENTIRE paragraph is replaced with image at signing
  leftSigCell.appendParagraph('{{SIGNATURE_1}}')
    .editAsText().setFontFamily('Arial').setFontSize(11).setItalic(true).setForegroundColor('#cccccc');
  leftSigCell.appendParagraph('Date Signed: {{SIGNED_AT}}')
    .editAsText().setFontFamily('Arial').setFontSize(9).setForegroundColor('#555555');
  leftSigCell.appendParagraph('Printed: {{CLIENT_NAME}}')
    .editAsText().setFontFamily('Arial').setFontSize(9).setForegroundColor('#555555');
  leftSigCell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12).setPaddingRight(12);

  // Right cell: contractor sign-off
  const rightSigCell = sigTbl.getCell(0, 1);
  rightSigCell.setTextAlignment(DocumentApp.TextAlignment.RIGHT);
  const rText = rightSigCell.editAsText()
    .setText('Thank you for your business!\n\n{{OWNER_NAME}}\nOwner, {{COMPANY_NAME}}')
    .setFontFamily('Arial').setFontSize(10).setForegroundColor(G);
  rText.setBold(0, 'Thank you for your business!'.length - 1, true);
  const ownerStart = 'Thank you for your business!\n\n'.length;
  rText.setFontSize(ownerStart, ownerStart + '{{OWNER_NAME}}'.length - 1, 14)
       .setItalic(ownerStart, ownerStart + '{{OWNER_NAME}}'.length - 1, true)
       .setForegroundColor(ownerStart, ownerStart + '{{OWNER_NAME}}'.length - 1, N);
  rightSigCell.setPaddingLeft(10).setPaddingTop(10).setPaddingBottom(10);

  body.appendParagraph('').editAsText().setFontSize(4);

  // ── Footer (navy, three columns) ──────────────────────────
  const footTbl = body.appendTable([['◉  LICENSED & INSURED', '◉  QUALITY CRAFTSMANSHIP', '◉  BUILT TO LAST']]);
  try { footTbl.setColumnWidth(0, 156); footTbl.setColumnWidth(1, 156); footTbl.setColumnWidth(2, 156); } catch(e) {}
  for (var i = 0; i < 3; i++) {
    const fc = footTbl.getCell(0, i);
    fc.setBackgroundColor(N);
    fc.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#aac4ff');
    fc.setTextAlignment(DocumentApp.TextAlignment.CENTER);
    fc.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(4).setPaddingRight(4);
  }

  doc.saveAndClose();

  const f = DriveApp.getFileById(doc.getId());
  folder.addFile(f);
  try { DriveApp.getRootFolder().removeFile(f); } catch(e) {}

  return doc.getId();
}

function createDefaultInvoiceTemplate_(cfg, folder) {
  const doc  = DocumentApp.create(cfg.COMPANY_NAME + ' — Invoice Template');
  const body = doc.getBody();
  body.clear();

  const N = '#0d2240', G = '#2e7d32', W = '#ffffff';

  function colorBanner(text, bg, fg, fontSize) {
    const tbl  = body.appendTable([[text]]);
    const cell = tbl.getCell(0, 0);
    cell.setBackgroundColor(bg);
    cell.editAsText().setFontFamily('Arial').setFontSize(fontSize || 11).setBold(true).setForegroundColor(fg);
    cell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(14).setPaddingRight(14);
    return tbl;
  }

  const hdrLines = [cfg.COMPANY_NAME];
  if (cfg.COMPANY_PHONE)   hdrLines.push(cfg.COMPANY_PHONE);
  if (cfg.COMPANY_EMAIL)   hdrLines.push(cfg.COMPANY_EMAIL);
  if (cfg.COMPANY_LICENSE) hdrLines.push('Lic # ' + cfg.COMPANY_LICENSE);
  const hdrTbl = colorBanner(hdrLines.join('\n'), N, W, 13);
  const hdrT   = hdrTbl.getCell(0, 0).editAsText();
  hdrT.setFontSize(0, cfg.COMPANY_NAME.length - 1, 16);
  const cs = cfg.COMPANY_NAME.length + 1;
  if (cs < hdrT.getText().length)
    hdrT.setFontSize(cs, hdrT.getText().length - 1, 10).setBold(cs, hdrT.getText().length - 1, false);

  body.appendParagraph('').editAsText().setFontSize(2);
  colorBanner('INVOICE FOR ROOF REPLACEMENT SERVICES', N, W, 12);
  body.appendParagraph('').editAsText().setFontSize(2);

  const infoTbl = body.appendTable([[
    'BILL TO\n\n{{CLIENT_NAME}}\n{{CLIENT_ADDRESS}}\n{{CLIENT_PHONE}}\n{{CLIENT_EMAIL}}',
    'INVOICE DETAILS\n\nInvoice #: {{INVOICE_NUM}}\nDate: {{DATE}}\nDue: Upon Completion\nRelated Est.: {{ESTIMATE_NUM}}'
  ]]);
  infoTbl.setColumnWidth(0, 234).setColumnWidth(1, 234);
  [['BILL TO', 0], ['INVOICE DETAILS', 1]].forEach(function(pair) {
    const label = pair[0], col = pair[1];
    const cell  = infoTbl.getCell(0, col);
    const t     = cell.editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#333333');
    t.setBold(0, label.length - 1, true).setForegroundColor(0, label.length - 1, G);
    cell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12).setPaddingRight(12);
  });
  const cnStart = 'BILL TO\n\n'.length;
  const blt = infoTbl.getCell(0, 0).editAsText();
  blt.setFontSize(cnStart, cnStart + '{{CLIENT_NAME}}'.length - 1, 12)
     .setBold(cnStart, cnStart + '{{CLIENT_NAME}}'.length - 1, true)
     .setForegroundColor(cnStart, cnStart + '{{CLIENT_NAME}}'.length - 1, N);

  body.appendParagraph('').editAsText().setFontSize(2);

  const srvH = body.appendParagraph('SERVICES RENDERED');
  srvH.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor('#888888');
  srvH.setSpacingBefore(6).setSpacingAfter(4);

  body.appendParagraph('{{SCOPE_OF_WORK}}')
    .editAsText().setFontFamily('Arial').setFontSize(11).setForegroundColor('#222222');

  body.appendParagraph('').editAsText().setFontSize(2);

  const totalTbl  = body.appendTable([['TOTAL DUE', '{{TOTAL}}']]);
  const leftCell  = totalTbl.getCell(0, 0);
  const rightCell = totalTbl.getCell(0, 1);
  leftCell.setBackgroundColor(G); rightCell.setBackgroundColor(G);
  leftCell.editAsText().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(W);
  rightCell.editAsText().setFontFamily('Arial').setFontSize(28).setBold(true).setForegroundColor(W);
  rightCell.setTextAlignment(DocumentApp.TextAlignment.RIGHT);
  [leftCell, rightCell].forEach(function(c) {
    c.setPaddingTop(12).setPaddingBottom(12).setPaddingLeft(14).setPaddingRight(14);
  });
  totalTbl.setColumnWidth(0, 300).setColumnWidth(1, 168);

  body.appendParagraph('\nPayment due upon completion. Please make checks payable to ' + cfg.COMPANY_NAME + '.')
    .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#555555');

  body.appendParagraph('{{NOTES}}')
    .editAsText().setFontFamily('Arial').setFontSize(10).setItalic(true).setForegroundColor('#555555');

  const ty = body.appendParagraph('\nThank you for your business!');
  ty.editAsText().setFontFamily('Arial').setFontSize(13).setBold(true).setForegroundColor(G);
  ty.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(16);

  body.appendParagraph('We appreciate your trust in ' + cfg.COMPANY_NAME + '.')
    .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor('#666666');

  const footTbl = body.appendTable([['◉  LICENSED & INSURED', '◉  QUALITY CRAFTSMANSHIP', '◉  BUILT TO LAST']]);
  footTbl.setColumnWidth(0, 156).setColumnWidth(1, 156).setColumnWidth(2, 156);
  for (var i = 0; i < 3; i++) {
    const fc = footTbl.getCell(0, i);
    fc.setBackgroundColor(N);
    fc.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#aac4ff');
    fc.setTextAlignment(DocumentApp.TextAlignment.CENTER);
    fc.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(4).setPaddingRight(4);
  }

  doc.saveAndClose();
  const f = DriveApp.getFileById(doc.getId());
  folder.addFile(f);
  try { DriveApp.getRootFolder().removeFile(f); } catch(e) {}
  return doc.getId();
}

// ============================================================
// CRM LOGGING
// ============================================================

function ensureLeadsHeaders_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow([
    'Timestamp', 'Estimate #', 'Client Name', 'Client Email', 'Client Phone',
    'Client Address', 'Roof Type', 'Sq Ft', 'Warranty', 'Color',
    'Total ($)', 'Status', 'Sign Token', 'Doc ID', 'Doc Hash',
    'Doc URL', 'Signed PDF URL', 'Signed At', 'Signer IP', 'Signer UA',
    'Invoice URL', 'Line Items (JSON)', 'Notes',
  ]);
}

function logToLeads_(p, cfg, estimateNum, token, docId, docHash, docUrl) {
  const sheet = getSheet_('Leads', cfg);
  ensureLeadsHeaders_(sheet);
  sheet.appendRow([
    new Date(),              // 1  TIMESTAMP
    estimateNum,             // 2  EST_NUM
    p.clientName,            // 3  CLIENT_NAME
    p.clientEmail   || '',   // 4  CLIENT_EMAIL
    p.clientPhone   || '',   // 5  CLIENT_PHONE
    p.clientAddress || '',   // 6  CLIENT_ADDR
    p.roofType      || '',   // 7  ROOF_TYPE
    p.sqft          || '',   // 8  SQ_FT
    p.warranty      || cfg.DEFAULT_WARRANTY, // 9 WARRANTY
    p.colorChoice   || '',   // 10 COLOR
    p.total         || 0,    // 11 TOTAL
    'Awaiting Signature',    // 12 STATUS
    token,                   // 13 SIGN_TOKEN
    docId,                   // 14 DOC_ID
    docHash,                 // 15 DOC_HASH
    docUrl,                  // 16 DOC_URL
    '',                      // 17 SIGNED_PDF
    '',                      // 18 SIGNED_AT
    '',                      // 19 SIGNER_IP
    '',                      // 20 SIGNER_UA
    '',                      // 21 INVOICE_URL
    JSON.stringify(p.lineItems || []), // 22 LINE_ITEMS
    p.notes         || '',   // 23 NOTES
  ]);
  return sheet.getLastRow();
}

function generateEstimateNumber_(companyName, rowId) {
  const prefix = (companyName || 'EST').split(/\s+/)
    .map(function(w) { return w[0] ? w[0].toUpperCase() : ''; })
    .join('').substring(0, 4);
  const dateStr = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyMMdd');
  return prefix + '-' + dateStr + '-' + String(rowId).padStart(2, '0');
}

// ============================================================
// EMAIL
// ============================================================

function sendSigningLinkEmail_(p, signingUrl, docUrl, estimateNum, cfg) {
  const subject = 'Your Roof Replacement Estimate is Ready to Sign — ' + cfg.COMPANY_NAME;
  const totalFmt = '$' + Number(p.total).toLocaleString();

  const html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#0d2240;padding:24px;text-align:center;">' +
      '<h1 style="color:white;margin:0;font-size:22px;">' + cfg.COMPANY_NAME + '</h1>' +
      '<p style="color:#aac4ff;margin:6px 0 0;font-size:13px;">' + cfg.COMPANY_TAGLINE +
        (cfg.COMPANY_LICENSE ? ' · Lic. ' + cfg.COMPANY_LICENSE : '') + '</p>' +
    '</div>' +
    '<div style="padding:28px;background:#f9f9f9;">' +
      '<p style="margin:0 0 16px;color:#333;">Dear ' + (p.clientName || 'Valued Customer') + ',</p>' +
      '<p style="margin:0 0 16px;color:#555;font-size:14px;">Thank you for considering ' + cfg.COMPANY_NAME + '. Your roof replacement proposal is ready for your review and signature.</p>' +
      '<div style="background:white;border:1px solid #e0e0e0;border-radius:8px;padding:20px;margin:20px 0;">' +
        '<p style="margin:0 0 4px;font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.5px;">Estimate #</p>' +
        '<p style="margin:0 0 16px;font-weight:700;color:#0d2240;font-size:15px;">' + estimateNum + '</p>' +
        '<p style="margin:0 0 4px;font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.5px;">Total Estimate</p>' +
        '<p style="margin:0;font-weight:800;font-size:30px;color:#2e7d32;">' + totalFmt + '</p>' +
      '</div>' +
      (signingUrl
        ? '<a href="' + signingUrl + '" style="display:block;background:#2e7d32;color:white;text-align:center;padding:16px;border-radius:8px;text-decoration:none;font-weight:700;font-size:16px;margin:20px 0;">✍️ Review &amp; Sign Proposal</a>' +
          '<p style="font-size:11px;color:#888;text-align:center;margin:0 0 20px;">This signing link is unique to you and is legally binding under the ESIGN Act.</p>'
        : '<p style="color:#555;margin:16px 0;">Please contact us to review your proposal.</p>'
      ) +
      (docUrl
        ? '<p style="text-align:center;margin:0 0 20px;"><a href="' + docUrl + '" style="color:#0d2240;font-size:13px;">📄 View Full Proposal Document</a></p>'
        : ''
      ) +
      '<hr style="border:none;border-top:1px solid #eee;margin:20px 0;">' +
      '<p style="color:#888;font-size:12px;margin:0;">' + cfg.COMPANY_NAME +
        (cfg.COMPANY_PHONE ? ' · ' + cfg.COMPANY_PHONE : '') +
        (cfg.COMPANY_EMAIL ? ' · ' + cfg.COMPANY_EMAIL : '') + '</p>' +
    '</div></div>';

  GmailApp.sendEmail(p.clientEmail, subject,
    'Dear ' + (p.clientName || 'Customer') + ', your roof replacement estimate ' + estimateNum +
    ' (' + totalFmt + ') is ready to sign. Visit: ' + signingUrl,
    {
      htmlBody: html,
      name:     cfg.COMPANY_NAME,
      replyTo:  cfg.COMPANY_EMAIL,
    }
  );
}

function sendOwnerNotification_(p, signingUrl, docUrl, estimateNum, cfg) {
  if (!cfg.COMPANY_EMAIL) return;
  const subject = '✅ New Proposal Sent — ' + estimateNum + ' · ' + p.clientName +
                  ' ($' + Number(p.total).toLocaleString() + ')';
  const body = [
    'New proposal created and signing link sent.',
    '',
    'Estimate:  ' + estimateNum,
    'Client:    ' + p.clientName,
    'Address:   ' + (p.clientAddress || 'N/A'),
    'Phone:     ' + (p.clientPhone   || 'N/A'),
    'Email:     ' + (p.clientEmail   || 'NOT PROVIDED'),
    'Total:     $' + Number(p.total).toLocaleString(),
    '',
    '🔗 Client Signing Link: ' + (signingUrl || '(deploy as web app first)'),
    '✏️  Editable Doc:         ' + docUrl,
    '',
    p.notes ? 'Notes: ' + p.notes : '',
  ].join('\n');
  GmailApp.sendEmail(cfg.COMPANY_EMAIL, subject, body);
}

function sendSignedEmail_(p, pdfBlob, cfg) {
  const totalFmt   = '$' + Number(p.total || 0).toLocaleString();
  const subjectOwn = '✅ SIGNED — Proposal ' + p.estimateNum + ' · ' + p.clientName;
  const subjectCli = 'Your Signed Roof Proposal — ' + cfg.COMPANY_NAME;

  const bodyText = 'Proposal ' + p.estimateNum + ' has been electronically signed by ' +
                   p.clientName + '. The signed PDF is attached.';

  const pdfName = 'SIGNED_Proposal_' + p.estimateNum + '.pdf';
  if (pdfBlob) pdfBlob.setName(pdfName);

  const attachments = pdfBlob ? [pdfBlob] : [];

  // Notify Omar
  if (cfg.COMPANY_EMAIL) {
    try {
      GmailApp.sendEmail(cfg.COMPANY_EMAIL, subjectOwn, bodyText,
        { attachments: attachments, name: cfg.COMPANY_NAME }
      );
    } catch(e) { Logger.log('Owner signed email failed: ' + e.message); }
  }

  // Email client their copy
  if (p.clientEmail) {
    const clientHtml =
      '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
      '<div style="background:#0d2240;padding:20px;text-align:center;">' +
        '<h1 style="color:white;margin:0;font-size:20px;">' + cfg.COMPANY_NAME + '</h1>' +
      '</div>' +
      '<div style="padding:24px;background:#f9f9f9;">' +
        '<p>Dear ' + (p.clientName || 'Customer') + ',</p>' +
        '<p>Thank you for signing your roofing proposal. Please find your signed copy attached for your records.</p>' +
        '<div style="background:white;border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin:16px 0;">' +
          '<p style="margin:0 0 4px;color:#888;font-size:11px;">ESTIMATE</p>' +
          '<p style="margin:0 0 12px;font-weight:700;color:#0d2240;">' + p.estimateNum + '</p>' +
          '<p style="margin:0 0 4px;color:#888;font-size:11px;">TOTAL</p>' +
          '<p style="margin:0;font-size:26px;font-weight:800;color:#2e7d32;">' + totalFmt + '</p>' +
        '</div>' +
        '<p style="color:#666;font-size:12px;">Questions? Contact us at ' +
          (cfg.COMPANY_PHONE || cfg.COMPANY_EMAIL) + '</p>' +
      '</div></div>';
    try {
      GmailApp.sendEmail(p.clientEmail, subjectCli, bodyText,
        { htmlBody: clientHtml, attachments: attachments, name: cfg.COMPANY_NAME, replyTo: cfg.COMPANY_EMAIL }
      );
    } catch(e) { Logger.log('Client signed email failed: ' + e.message); }
  }
}

// ============================================================
// UTILITIES
// ============================================================

function getSheet_(name, cfg) {
  const id = (cfg && cfg.SHEET_ID) ? cfg.SHEET_ID : P('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id).getSheetByName(name); } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function tryParseJSON_(str) {
  try { return JSON.parse(str || '[]'); } catch(e) { return []; }
}

// Legacy aliases for any menu references
function getSheet(name, cfg) { return getSheet_(name, cfg); }
function tryParseJSON(str)   { return tryParseJSON_(str); }
