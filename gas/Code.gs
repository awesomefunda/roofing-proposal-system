// ============================================================
// ROOFING PROPOSAL SYSTEM — Google Apps Script
// ============================================================
// Runs entirely inside the contractor's Google account.
// No server, no Railway, no developer approval needed.
//
// HOW TO SET UP (one time, ~3 minutes)
// ─────────────────────────────────────
//   1. Open the Google Sheet → Extensions → Apps Script
//   2. Deploy → New Deployment → Web App
//        Execute as: Me   |   Access: Anyone
//   3. Click Authorize → Allow
//   4. Open the sheet → click "🏠 Roofing Tools" menu
//   5. Fill in company details → Save
//
// PRIVACY: Everything runs as YOU. PDFs go to YOUR Drive.
//          Emails come from YOUR Gmail. Zero third-party access.
// ============================================================

// ── Bootstrap (injected by provisioner, ignored otherwise) ───
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

// ── Script Properties helpers ─────────────────────────────────

function P(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

function setProps(obj) {
  PropertiesService.getScriptProperties().setProperties(obj);
}

// ── Config ────────────────────────────────────────────────────

function getConfig() {
  return {
    SHEET_ID:         P('SHEET_ID'),
    PROPOSALS_FOLDER_ID: P('PROPOSALS_FOLDER_ID'),
    COMPANY_NAME:     P('COMPANY_NAME')    || '',
    COMPANY_PHONE:    P('COMPANY_PHONE')   || '',
    COMPANY_EMAIL:    P('COMPANY_EMAIL')   || '',
    COMPANY_LICENSE:  P('COMPANY_LICENSE') || '',
    COMPANY_TAGLINE:  P('COMPANY_TAGLINE') || 'Licensed & Insured',
    LOGO_FILE_ID:     P('LOGO_FILE_ID')    || '',
    TIMEZONE:         P('TIMEZONE')        || 'America/Los_Angeles',
    DEFAULT_WARRANTY: P('DEFAULT_WARRANTY')|| '7-year',
    ROTTED_WOOD_RATE: P('ROTTED_WOOD_RATE')|| '$26.00 per foot',
    ROOF_TYPES:       (P('ROOF_TYPES')     || 'Flat,Tile,Composition Shingle,Metal,Other').split(','),
    WARRANTY_OPTIONS: (P('WARRANTY_OPTIONS')|| '5-year,7-year,10-year,Manufacturer warranty').split(','),
  };
}

function isSetupComplete() {
  return !!(P('COMPANY_NAME') && P('PROPOSALS_FOLDER_ID'));
}

// ============================================================
// MENU
// ============================================================

function onOpen() {
  _initProps();
  const name = P('COMPANY_NAME') || 'Roofing Tools';
  SpreadsheetApp.getUi()
    .createMenu('🏠 ' + name)
    .addItem('✏️  New Proposal',          'showProposalSidebar')
    .addItem('🧾  Invoice Selected Row',   'invoiceSelectedRow')
    .addSeparator()
    .addItem('⚙️  Company Settings',       'showSettingsSidebar')
    .addToUi();
}

function showProposalSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('New Proposal')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showSettingsSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Settings')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// SIDEBAR API  (called via google.script.run)
// ============================================================

function getAppState() {
  _initProps();
  const cfg   = getConfig();
  const items = getCatalogItems();
  return {
    setupComplete:   isSetupComplete(),
    companyName:     cfg.COMPANY_NAME,
    companyPhone:    cfg.COMPANY_PHONE,
    companyEmail:    cfg.COMPANY_EMAIL,
    companyLicense:  cfg.COMPANY_LICENSE,
    companyTagline:  cfg.COMPANY_TAGLINE,
    roofTypes:       cfg.ROOF_TYPES,
    warrantyOptions: cfg.WARRANTY_OPTIONS,
    defaultWarranty: cfg.DEFAULT_WARRANTY,
    catalog:         items,
  };
}

function getCatalogItems() {
  try {
    const cfg   = getConfig();
    const sheet = getSheet('Catalog', cfg);
    const rows  = sheet.getDataRange().getValues();
    return rows.slice(1)
      .filter(r => String(r[5]).toUpperCase() === 'YES')
      .map(r => ({
        item:        String(r[0]),
        description: String(r[1]),
        price:       Number(r[2]) || 0,
        unit:        String(r[3]),
        category:    String(r[4]),
      }));
  } catch(e) {
    return [];
  }
}

// ── Save company settings (first-run or update) ───────────────

function saveSettings(payload) {
  const company = (payload.companyName || '').trim();
  if (!company) return { error: 'Company name is required.' };

  // Create Proposals folder if it doesn't exist yet
  let folderId = P('PROPOSALS_FOLDER_ID');
  if (!folderId) {
    const folder = DriveApp.createFolder(company + ' — Proposals');
    folderId = folder.getId();
  }

  // Get the bound Sheet ID
  let sheetId = P('SHEET_ID');
  if (!sheetId) {
    try { sheetId = SpreadsheetApp.getActiveSpreadsheet().getId(); } catch(e) {}
  }

  setProps({
    SHEET_ID:         sheetId,
    PROPOSALS_FOLDER_ID: folderId,
    COMPANY_NAME:     company,
    COMPANY_PHONE:    (payload.companyPhone   || '').trim(),
    COMPANY_EMAIL:    (payload.companyEmail   || '').trim(),
    COMPANY_LICENSE:  (payload.companyLicense || '').trim(),
    COMPANY_TAGLINE:  (payload.companyTagline || 'Licensed & Insured').trim(),
    _init:            '1',
  });

  // Rebuild the menu with the new company name
  onOpen();

  return { success: true, folderUrl: DriveApp.getFolderById(folderId).getUrl() };
}

// ── Submit a proposal ─────────────────────────────────────────

function submitProposal(payload) {
  const cfg   = getConfig();
  if (!isSetupComplete()) return { error: 'Please complete company setup first.' };

  const rowId = logToLeads(payload, cfg);
  const { pdfBlob, fileName, estimateNum } = generatePDF(payload, rowId, 'proposal', cfg);

  const folder  = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  const file    = folder.createFile(pdfBlob.setName(fileName));
  const fileUrl = file.getUrl();

  getSheet('Leads', cfg).getRange(rowId, 13).setValue(fileUrl);

  if (payload.clientEmail) sendProposalEmail(payload, pdfBlob, fileName, estimateNum, cfg);
  sendOwnerNotification(payload, fileUrl, rowId, estimateNum, cfg);

  return { success: true, rowId, fileUrl, estimateNum, fileName };
}

// ── Generate invoice from a Leads row ─────────────────────────

function generateInvoiceFromRow(rowId) {
  const cfg   = getConfig();
  const sheet = getSheet('Leads', cfg);
  const row   = sheet.getRange(rowId, 1, 1, 16).getValues()[0];

  const payload = {
    clientName:    row[1],  clientAddress: row[2],
    clientPhone:   row[3],  clientEmail:   row[4],
    roofType:      row[5],  sqft:          row[6],
    warranty:      row[7],  colorChoice:   row[8],
    lineItems:     JSON.parse(row[9] || '[]'),
    total:         row[10], notes:         row[15],
  };

  const { pdfBlob, fileName } = generatePDF(payload, rowId, 'invoice', cfg);
  const folder  = DriveApp.getFolderById(cfg.PROPOSALS_FOLDER_ID);
  const file    = folder.createFile(pdfBlob.setName(fileName));
  const fileUrl = file.getUrl();

  sheet.getRange(rowId, 15).setValue(fileUrl);
  sheet.getRange(rowId, 12).setValue('Invoiced');

  return { success: true, invoiceUrl: fileUrl };
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
    ? '✅ Invoice saved to Drive:\n' + result.invoiceUrl
    : '❌ Error: ' + result.error);
}

// ============================================================
// LOG TO LEADS SHEET
// ============================================================

function logToLeads(p, cfg) {
  // Ensure Leads sheet has headers
  const sheet = getSheet('Leads', cfg);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp','Client Name','Address','Phone','Email',
      'Roof Type','Sq Ft','Warranty','Color Choice',
      'Line Items (JSON)','Total ($)','Status',
      'Proposal URL','DocuSign Status','Invoice URL','Notes',
    ]);
  }
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
    JSON.stringify(p.lineItems || []),
    p.total            || 0,
    'Proposal Sent',
    '',   // Proposal URL — filled after PDF saved
    'Pending',
    '',   // Invoice URL
    p.notes || '',
  ]);
  return sheet.getLastRow();
}

// ============================================================
// PDF GENERATION  (HTML → Drive → PDF)
// ============================================================

function generateEstimateNumber(companyName, rowId) {
  const prefix = companyName.split(/\s+/)
    .map(function(w) { return w[0] ? w[0].toUpperCase() : ''; })
    .join('').substring(0, 4);
  const dateStr = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyMMdd');
  return prefix + '-' + dateStr + '-' + String(rowId).padStart(2, '0');
}

function getLogoBase64(cfg) {
  const logoId = cfg.LOGO_FILE_ID;
  if (!logoId) return '';
  try {
    const file  = DriveApp.getFileById(logoId);
    const mime  = file.getMimeType() || 'image/png';
    const bytes = file.getBlob().getBytes();
    return 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes);
  } catch(e) {
    Logger.log('Logo fetch failed: ' + e.message);
    return '';
  }
}

function generatePDF(p, rowId, type, cfg) {
  const estimateNum = generateEstimateNumber(cfg.COMPANY_NAME, rowId);
  const logoBase64  = getLogoBase64(cfg);

  const html = (type === 'invoice')
    ? buildInvoiceHtml(p, rowId, estimateNum, logoBase64, cfg)
    : buildProposalHtml(p, rowId, estimateNum, logoBase64, cfg);

  const blob    = Utilities.newBlob(html, 'text/html', 'temp.html');
  const tmpFile = DriveApp.createFile(blob);
  const pdfBlob = tmpFile.getAs('application/pdf');
  tmpFile.setTrashed(true);

  const safeName  = p.clientName.replace(/[^a-zA-Z0-9]/g, '_');
  const dateStr   = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'yyyy-MM-dd');
  const typeLabel = (type === 'invoice') ? 'Invoice' : 'Proposal';
  const fileName  = cfg.COMPANY_NAME.replace(/\s/g, '_') + '_' + typeLabel + '_' + safeName + '_' + dateStr + '.pdf';

  return { pdfBlob, fileName, estimateNum };
}

// ── HTML proposal builder ─────────────────────────────────────

function buildProposalHtml(p, rowId, estimateNum, logoBase64, cfg) {
  const NAVY  = '#0d2240';
  const GREEN = '#2e7d32';
  const LGREY = '#f5f5f5';
  const dateFormatted = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'MMMM dd, yyyy');
  const totalFmt = '$' + Number(p.total).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});

  const logoHtml = logoBase64
    ? '<img src="' + logoBase64 + '" style="max-height:80px;max-width:160px;object-fit:contain;" />'
    : '<div style="font-size:26px;font-weight:900;color:' + NAVY + ';">' + esc(cfg.COMPANY_NAME) + '</div>';

  const itemRows = (p.lineItems || []).map(function(item, i) {
    const price = item.price ? '$' + Number(item.price).toLocaleString('en-US', {minimumFractionDigits:2}) : 'Included';
    return '<tr>'
      + '<td style="padding:10px 8px;vertical-align:top;width:36px;">'
      +   '<div style="background:' + GREEN + ';color:white;border-radius:50%;width:24px;height:24px;'
      +   'text-align:center;line-height:24px;font-size:11px;font-weight:700;margin:auto;">' + (i+1) + '</div>'
      + '</td>'
      + '<td style="padding:10px 8px;border-bottom:1px dotted #ccc;">'
      +   '<span style="font-weight:600;color:#222;">' + esc(item.description) + '</span>'
      +   (item.unit ? '<br><span style="font-size:11px;color:#888;">' + esc(item.unit) + '</span>' : '')
      + '</td>'
      + '<td style="padding:10px 8px;text-align:right;white-space:nowrap;border-bottom:1px dotted #ccc;font-weight:600;">'
      +   price + '</td>'
      + '</tr>';
  }).join('');

  const notesHtml = p.notes
    ? '<div style="margin-top:18px;padding:12px 16px;background:#fffde7;border-left:4px solid #f9a825;border-radius:4px;">'
    +   '<p style="margin:0;font-size:11px;font-weight:700;color:#555;text-transform:uppercase;">Notes</p>'
    +   '<p style="margin:6px 0 0;font-size:13px;color:#333;">' + esc(p.notes) + '</p></div>'
    : '';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<style>body{font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333;margin:0;padding:0;}table{border-collapse:collapse;width:100%;}.page{width:750px;margin:0 auto;padding:0 0 40px;}</style>'
    + '</head><body><div class="page">'

    // Header
    + '<table style="padding:24px 30px 20px;" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="width:180px;vertical-align:middle;">' + logoHtml + '</td>'
    + '<td style="vertical-align:middle;padding-left:24px;">'
    +   '<div style="font-size:17px;font-weight:800;color:' + NAVY + ';">' + esc(cfg.COMPANY_NAME) + '</div>'
    +   (cfg.COMPANY_PHONE ? '<div style="font-size:12px;color:#555;margin-top:4px;">📞 ' + esc(cfg.COMPANY_PHONE) + '</div>' : '')
    +   (cfg.COMPANY_EMAIL ? '<div style="font-size:12px;color:#555;margin-top:2px;">✉️ ' + esc(cfg.COMPANY_EMAIL) + '</div>' : '')
    +   (cfg.COMPANY_LICENSE ? '<div style="font-size:12px;color:#555;margin-top:2px;">🪪 Lic. ' + esc(cfg.COMPANY_LICENSE) + '</div>' : '')
    + '</td></tr></table>'

    // Navy banner
    + '<div style="background:' + NAVY + ';padding:14px 30px;">'
    + '<span style="color:white;font-size:15px;font-weight:800;letter-spacing:1px;text-transform:uppercase;">Estimate &amp; Proposal for Roof Replacement</span></div>'

    // Client + Estimate columns
    + '<table style="background:' + LGREY + ';" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="width:50%;padding:18px 30px;vertical-align:top;border-right:2px solid #ddd;">'
    +   '<div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">Prepared For</div>'
    +   '<div style="font-size:15px;font-weight:800;color:' + NAVY + ';">' + esc(p.clientName) + '</div>'
    +   '<div style="margin-top:6px;font-size:12px;color:#444;line-height:1.6;">' + esc(p.clientAddress)
    +   (p.clientPhone ? '<br>📞 ' + esc(p.clientPhone) : '')
    +   (p.clientEmail ? '<br>✉️ ' + esc(p.clientEmail) : '') + '</div>'
    + '</td>'
    + '<td style="width:50%;padding:18px 30px;vertical-align:top;">'
    +   '<div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">Estimate Details</div>'
    +   '<table style="font-size:12px;" cellpadding="3" cellspacing="0">'
    +     '<tr><td style="color:#888;width:110px;">Estimate #</td><td style="font-weight:700;color:' + NAVY + ';">' + esc(estimateNum) + '</td></tr>'
    +     '<tr><td style="color:#888;">Date</td><td>' + dateFormatted + '</td></tr>'
    +     '<tr><td style="color:#888;">Roof Type</td><td>' + esc(p.roofType || '—') + '</td></tr>'
    +     (p.sqft ? '<tr><td style="color:#888;">Square Feet</td><td>' + Number(p.sqft).toLocaleString() + ' sqft</td></tr>' : '')
    +     '<tr><td style="color:#888;">Warranty</td><td>' + esc(p.warranty || cfg.DEFAULT_WARRANTY) + '</td></tr>'
    +     (p.colorChoice ? '<tr><td style="color:#888;">Color</td><td>' + esc(p.colorChoice) + '</td></tr>' : '')
    +   '</table>'
    + '</td></tr></table>'

    // Scope of work
    + '<div style="padding:20px 30px 0;">'
    + '<div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px;border-bottom:2px solid ' + NAVY + ';padding-bottom:6px;">Scope of Work</div>'
    + '<table cellpadding="0" cellspacing="0"><tbody>' + itemRows + '</tbody></table></div>'

    // Total bar
    + '<div style="background:' + GREEN + ';margin:20px 30px 0;padding:14px 20px;border-radius:6px;">'
    + '<table style="width:100%;" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="color:white;font-size:14px;font-weight:700;">TOTAL ESTIMATE</td>'
    + '<td style="text-align:right;color:white;font-size:22px;font-weight:900;">' + totalFmt + '</td>'
    + '</tr></table></div>'

    // Disclaimer
    + '<div style="padding:14px 30px 0;font-size:11px;color:#666;line-height:1.6;">'
    + '• Additional rotted wood charged at ' + esc(cfg.ROTTED_WOOD_RATE) + ' upon homeowner approval.<br>'
    + '• Work area kept clean at the end of every work day.</div>'

    + (notesHtml ? '<div style="padding:0 30px;">' + notesHtml + '</div>' : '')

    // Signature
    + '<div style="margin:28px 30px 0;padding:18px 20px;border:1px solid #ddd;border-radius:6px;background:#fafafa;">'
    + '<p style="margin:0 0 14px;font-size:12px;color:#555;">By signing, you authorize <strong>' + esc(cfg.COMPANY_NAME) + '</strong> to perform the above work and agree to pay the total amount upon completion.</p>'
    + '<table style="width:100%;" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="width:60%;border-top:1px solid #333;padding-top:6px;font-size:11px;color:#666;">Homeowner Signature</td>'
    + '<td style="width:8%;"></td>'
    + '<td style="width:32%;border-top:1px solid #333;padding-top:6px;font-size:11px;color:#666;">Date</td>'
    + '</tr></table>'
    + '<p style="margin:18px 0 0;text-align:center;font-size:13px;font-weight:700;color:' + GREEN + ';">Thank you for your business!</p></div>'

    // Footer
    + '<div style="background:' + NAVY + ';margin-top:30px;padding:12px 30px;text-align:center;">'
    + '<span style="color:#aac4ff;font-size:11px;letter-spacing:.8px;text-transform:uppercase;">'
    + 'Licensed &amp; Insured &nbsp;|&nbsp; Quality Craftsmanship &nbsp;|&nbsp; Built to Last'
    + (cfg.COMPANY_TAGLINE && cfg.COMPANY_TAGLINE !== 'Licensed & Insured' ? ' &nbsp;|&nbsp; ' + esc(cfg.COMPANY_TAGLINE) : '')
    + '</span></div>'
    + '</div></body></html>';
}

// ── HTML invoice builder ──────────────────────────────────────

function buildInvoiceHtml(p, rowId, estimateNum, logoBase64, cfg) {
  const NAVY  = '#0d2240';
  const GREEN = '#2e7d32';
  const LGREY = '#f5f5f5';
  const dateFormatted = Utilities.formatDate(new Date(), cfg.TIMEZONE, 'MMMM dd, yyyy');
  const totalFmt = '$' + Number(p.total).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
  const invoiceNum = estimateNum.replace(/^([A-Z]+)-/, '$1INV-');

  const logoHtml = logoBase64
    ? '<img src="' + logoBase64 + '" style="max-height:80px;max-width:160px;object-fit:contain;" />'
    : '<div style="font-size:26px;font-weight:900;color:' + NAVY + ';">' + esc(cfg.COMPANY_NAME) + '</div>';

  const itemRows = (p.lineItems || []).map(function(item, i) {
    const price = item.price ? '$' + Number(item.price).toLocaleString('en-US', {minimumFractionDigits:2}) : 'Included';
    return '<tr>'
      + '<td style="padding:10px 8px;vertical-align:top;width:36px;">'
      +   '<div style="background:' + GREEN + ';color:white;border-radius:50%;width:24px;height:24px;text-align:center;line-height:24px;font-size:11px;font-weight:700;margin:auto;">' + (i+1) + '</div>'
      + '</td>'
      + '<td style="padding:10px 8px;border-bottom:1px dotted #ccc;">'
      +   '<span style="font-weight:600;color:#222;">' + esc(item.description) + '</span>'
      +   (item.unit ? '<br><span style="font-size:11px;color:#888;">' + esc(item.unit) + '</span>' : '')
      + '</td>'
      + '<td style="padding:10px 8px;text-align:right;white-space:nowrap;border-bottom:1px dotted #ccc;font-weight:600;">' + price + '</td>'
      + '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<style>body{font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333;margin:0;padding:0;}table{border-collapse:collapse;width:100%;}.page{width:750px;margin:0 auto;padding:0 0 40px;}</style>'
    + '</head><body><div class="page">'

    + '<table style="padding:24px 30px 20px;" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="width:180px;vertical-align:middle;">' + logoHtml + '</td>'
    + '<td style="vertical-align:middle;padding-left:24px;">'
    +   '<div style="font-size:17px;font-weight:800;color:' + NAVY + ';">' + esc(cfg.COMPANY_NAME) + '</div>'
    +   (cfg.COMPANY_PHONE ? '<div style="font-size:12px;color:#555;margin-top:4px;">📞 ' + esc(cfg.COMPANY_PHONE) + '</div>' : '')
    +   (cfg.COMPANY_EMAIL ? '<div style="font-size:12px;color:#555;margin-top:2px;">✉️ ' + esc(cfg.COMPANY_EMAIL) + '</div>' : '')
    +   (cfg.COMPANY_LICENSE ? '<div style="font-size:12px;color:#555;margin-top:2px;">🪪 Lic. ' + esc(cfg.COMPANY_LICENSE) + '</div>' : '')
    + '</td>'
    + '<td style="vertical-align:middle;text-align:right;">'
    +   '<div style="font-size:28px;font-weight:900;color:' + NAVY + ';">INVOICE</div>'
    +   '<div style="font-size:12px;color:#888;"># ' + invoiceNum + '</div>'
    + '</td></tr></table>'

    + '<div style="background:' + NAVY + ';padding:14px 30px;">'
    + '<span style="color:white;font-size:15px;font-weight:800;letter-spacing:1px;text-transform:uppercase;">Invoice for Roof Replacement Services</span></div>'

    + '<table style="background:' + LGREY + ';" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="width:50%;padding:18px 30px;vertical-align:top;border-right:2px solid #ddd;">'
    +   '<div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">Bill To</div>'
    +   '<div style="font-size:15px;font-weight:800;color:' + NAVY + ';">' + esc(p.clientName) + '</div>'
    +   '<div style="margin-top:6px;font-size:12px;color:#444;line-height:1.6;">' + esc(p.clientAddress)
    +   (p.clientPhone ? '<br>📞 ' + esc(p.clientPhone) : '')
    +   (p.clientEmail ? '<br>✉️ ' + esc(p.clientEmail) : '') + '</div>'
    + '</td>'
    + '<td style="width:50%;padding:18px 30px;vertical-align:top;">'
    +   '<div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">Invoice Details</div>'
    +   '<table style="font-size:12px;" cellpadding="3" cellspacing="0">'
    +     '<tr><td style="color:#888;width:110px;">Invoice #</td><td style="font-weight:700;color:' + NAVY + ';">' + invoiceNum + '</td></tr>'
    +     '<tr><td style="color:#888;">Date</td><td>' + dateFormatted + '</td></tr>'
    +     '<tr><td style="color:#888;">Due</td><td style="font-weight:700;color:#c62828;">Upon Completion</td></tr>'
    +     '<tr><td style="color:#888;">Related Est.</td><td>' + estimateNum + '</td></tr>'
    +   '</table>'
    + '</td></tr></table>'

    + '<div style="padding:20px 30px 0;">'
    + '<div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px;border-bottom:2px solid ' + NAVY + ';padding-bottom:6px;">Services Rendered</div>'
    + '<table cellpadding="0" cellspacing="0"><tbody>' + itemRows + '</tbody></table></div>'

    + '<div style="background:' + GREEN + ';margin:20px 30px 0;padding:14px 20px;border-radius:6px;">'
    + '<table style="width:100%;" cellpadding="0" cellspacing="0"><tr>'
    + '<td style="color:white;font-size:14px;font-weight:700;">TOTAL DUE</td>'
    + '<td style="text-align:right;color:white;font-size:22px;font-weight:900;">' + totalFmt + '</td>'
    + '</tr></table></div>'

    + '<div style="padding:14px 30px 0;font-size:11px;color:#666;">Payment due upon completion. Please make checks payable to <strong>' + esc(cfg.COMPANY_NAME) + '</strong>.</div>'

    + (p.notes ? '<div style="padding:14px 30px 0;"><div style="padding:12px 16px;background:#fffde7;border-left:4px solid #f9a825;border-radius:4px;font-size:12px;color:#333;">' + esc(p.notes) + '</div></div>' : '')

    + '<div style="margin:28px 30px 0;text-align:center;padding:18px;background:#f0f7f0;border-radius:6px;">'
    + '<p style="margin:0;font-size:14px;font-weight:700;color:' + GREEN + ';">Thank you for your business!</p>'
    + '<p style="margin:6px 0 0;font-size:12px;color:#666;">We appreciate your trust in ' + esc(cfg.COMPANY_NAME) + '.</p></div>'

    + '<div style="background:' + NAVY + ';margin-top:30px;padding:12px 30px;text-align:center;">'
    + '<span style="color:#aac4ff;font-size:11px;letter-spacing:.8px;text-transform:uppercase;">'
    + 'Licensed &amp; Insured &nbsp;|&nbsp; Quality Craftsmanship &nbsp;|&nbsp; Built to Last'
    + (cfg.COMPANY_TAGLINE && cfg.COMPANY_TAGLINE !== 'Licensed & Insured' ? ' &nbsp;|&nbsp; ' + esc(cfg.COMPANY_TAGLINE) : '')
    + '</span></div>'
    + '</div></body></html>';
}

// ── HTML escape helper ────────────────────────────────────────

function esc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ============================================================
// EMAIL
// ============================================================

function sendProposalEmail(p, pdfBlob, fileName, estimateNum, cfg) {
  const subject = 'Your Roof Replacement Estimate — ' + cfg.COMPANY_NAME + ' (' + estimateNum + ')';
  const html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">'
    + '<div style="background:#0d2240;padding:24px;text-align:center;">'
    + '<h1 style="color:white;margin:0;font-size:22px;">' + cfg.COMPANY_NAME + '</h1>'
    + '<p style="color:#aac4ff;margin:6px 0 0;">' + cfg.COMPANY_TAGLINE + (cfg.COMPANY_LICENSE ? ' · Lic. ' + cfg.COMPANY_LICENSE : '') + '</p>'
    + '</div>'
    + '<div style="padding:24px;background:#f9f9f9;">'
    + '<p>Dear ' + p.clientName + ',</p>'
    + '<p>Thank you for the opportunity. Please find your estimate attached.</p>'
    + '<div style="background:white;border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin:20px 0;">'
    + '<p style="margin:0;font-size:11px;color:#666;text-transform:uppercase;">Estimate #</p>'
    + '<p style="margin:4px 0 14px;font-weight:bold;color:#0d2240;">' + estimateNum + '</p>'
    + '<p style="margin:0;font-size:11px;color:#666;text-transform:uppercase;">Project Address</p>'
    + '<p style="margin:4px 0 14px;font-weight:bold;">' + p.clientAddress + '</p>'
    + '<p style="margin:0;font-size:11px;color:#666;text-transform:uppercase;">Total Estimate</p>'
    + '<p style="margin:4px 0 0;font-weight:800;font-size:26px;color:#2e7d32;">$' + Number(p.total).toLocaleString() + '</p>'
    + '</div>'
    + '<p style="color:#666;font-size:12px;">' + cfg.COMPANY_NAME + ' · ' + cfg.COMPANY_PHONE + '</p>'
    + '</div></div>';

  GmailApp.sendEmail(p.clientEmail, subject,
    'Dear ' + p.clientName + ', please find your roof estimate attached.',
    { htmlBody: html, attachments: [pdfBlob.setName(fileName)], name: cfg.COMPANY_NAME, replyTo: cfg.COMPANY_EMAIL }
  );
}

function sendOwnerNotification(p, fileUrl, rowId, estimateNum, cfg) {
  if (!cfg.COMPANY_EMAIL) return;
  const subject = '✅ New Proposal ' + estimateNum + ' — ' + p.clientName + ' ($' + Number(p.total).toLocaleString() + ')';
  const body = [
    'New proposal created — ' + estimateNum,
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
// UTILITIES
// ============================================================

function getSheet(name, cfg) {
  const id = (cfg && cfg.SHEET_ID) ? cfg.SHEET_ID : P('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id).getSheetByName(name); } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
