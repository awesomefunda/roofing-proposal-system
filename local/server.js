// ============================================================
// local/server.js — Roofing Proposal Platform
// ============================================================
// A self-service onboarding platform. Each contractor:
//   1. Visits this page
//   2. Enters company name + logo
//   3. Clicks "Sign in with Google"
//   4. Gets their own private Web App URL
//
// The platform owner has ZERO access to contractor data.
// Everything is created inside the contractor's Google account
// using their own OAuth token, which is discarded after setup.
//
// DEVELOPER SETUP (one-time):
//   • Copy local/app-credentials.json.example
//     → local/app-credentials.json
//   • Fill in your Google OAuth Client ID + Secret
//     (console.cloud.google.com → OAuth 2.0 Client ID)
//   • Redirect URI: http://localhost:3000/auth/callback
//   • npm run dev
// ============================================================

const express    = require('express');
const fs         = require('fs');
const path       = require('path');
const crypto     = require('crypto');
const auth       = require('./auth');
const { provision } = require('./provision');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '5mb' }));
app.use(express.urlencoded({ extended: true, limit: '5mb' }));

// In-memory session store (keyed by random state token)
// { [state]: { companyName, companyPhone, companyEmail, companyLicense, companyTagline, logoBase64 } }
const pendingSessions = new Map();

// In-memory job store — keyed by jobId, holds provisioning status + result
const provisionJobs = new Map();

// ============================================================
// DEVELOPER GATE — shown only if app-credentials.json missing
// ============================================================

function devGatePage(error) {
  return page('Developer Setup', `
    <div class="dev-card">
      <div class="dev-badge">Developer Setup</div>
      <h1>One-time configuration needed</h1>
      <p class="sub">Create a Google OAuth app so contractors can sign in. This screen is only visible to you — contractors never see it.</p>

      <div class="steps">
        <div class="step"><span>1</span><div>Go to <a href="https://console.cloud.google.com" target="_blank">console.cloud.google.com</a> → create or select a project</div></div>
        <div class="step"><span>2</span><div><strong>APIs &amp; Services → Library</strong> → enable <strong>Google Drive API</strong>, <strong>Google Sheets API</strong>, and <strong>Apps Script API</strong></div></div>
        <div class="step"><span>3</span><div><strong>Credentials → + Create → OAuth 2.0 Client ID</strong><br>Type: Web application<br>Redirect URI: <code>http://localhost:3000/auth/callback</code></div></div>
        <div class="step"><span>4</span><div>Paste Client ID + Secret below</div></div>
      </div>

      ${error ? `<div class="error-box">${error}</div>` : ''}

      <form method="POST" action="/dev-setup">
        <label>Client ID</label>
        <input name="clientId" placeholder="1234...apps.googleusercontent.com" required>
        <label>Client Secret</label>
        <input name="clientSecret" placeholder="GOCSPX-..." required>
        <button class="btn btn-primary" type="submit">Save Credentials</button>
      </form>
      <p class="dev-note">Saved to <code>local/app-credentials.json</code> (gitignored). Contractors only see the Sign in with Google button.</p>
    </div>
  `);
}

// ============================================================
// ONBOARDING PAGE — what the contractor sees
// ============================================================

function onboardingPage(error) {
  return page('Set Up Your Proposal System', `
    <div class="hero">
      <div class="hero-icon">🏠</div>
      <h1>Set up your roofing proposal system</h1>
      <p class="hero-sub">Enter your details, sign in with Google, and we'll create your private workspace in seconds — proposals, invoices, and templates all in your own Google Drive.</p>
    </div>

    ${error ? `<div class="error-box">⚠️ ${error}</div>` : ''}

    <div class="form-card">
      <form id="form" method="POST" action="/onboard">

        <div class="field-group">
          <label>Company Name <span class="req">*</span></label>
          <input type="text" name="companyName" placeholder="e.g. Acme Roofing Inc." required autocomplete="organization">
        </div>

        <div class="field-row">
          <div class="field-group">
            <label>Phone</label>
            <input type="text" name="companyPhone" placeholder="(555) 000-0000" autocomplete="tel">
          </div>
          <div class="field-group">
            <label>Email</label>
            <input type="text" name="companyEmail" placeholder="you@company.com" autocomplete="email">
          </div>
        </div>

        <div class="field-row">
          <div class="field-group">
            <label>License Number</label>
            <input type="text" name="companyLicense" placeholder="Lic # 1234567">
          </div>
          <div class="field-group">
            <label>Tagline</label>
            <input type="text" name="companyTagline" placeholder="Licensed &amp; Insured">
          </div>
        </div>

        <div class="field-group">
          <label>Company Logo <span class="opt">optional — appears on proposals</span></label>
          <div class="logo-drop" id="drop">
            <input type="file" id="fileInput" accept="image/*">
            <div class="drop-inner">
              <div class="drop-icon">🖼️</div>
              <div class="drop-text"><strong>Click to upload</strong> or drag &amp; drop</div>
              <div class="drop-hint">PNG · JPG · SVG · Auto-resized to 400×120 px</div>
            </div>
          </div>
          <div class="logo-preview" id="logoPreview">
            <img id="previewImg" src="" alt="">
            <div class="preview-meta">
              <span id="previewName"></span>
              <button type="button" id="removeBtn" class="remove-btn">✕ Remove</button>
            </div>
          </div>
          <input type="hidden" name="logoBase64" id="logoInput">
        </div>

        <div class="divider"></div>

        <button class="btn btn-google" type="submit" id="submitBtn">
          <svg width="20" height="20" viewBox="0 0 18 18" style="flex-shrink:0">
            <path fill="#fff" d="M17.64 9.2c0-.637-.057-1.251-.164-1.84H9v3.481h4.844c-.209 1.125-.843 2.078-1.796 2.717v2.258h2.908C16.658 14.215 17.64 11.907 17.64 9.2z" opacity=".9"/>
            <path fill="#fff" d="M9 18c2.43 0 4.467-.806 5.956-2.184l-2.908-2.258c-.806.54-1.837.86-3.048.86-2.344 0-4.328-1.584-5.036-3.711H.957v2.332C2.438 15.983 5.482 18 9 18z" opacity=".9"/>
            <path fill="#fff" d="M3.964 10.707c-.18-.54-.282-1.117-.282-1.707s.102-1.167.282-1.707V4.961H.957C.347 6.175 0 7.55 0 9s.348 2.826.957 4.039l3.007-2.332z" opacity=".9"/>
            <path fill="#fff" d="M9 3.58c1.321 0 2.508.454 3.44 1.345l2.582-2.58C13.463.891 11.426 0 9 0 5.482 0 2.438 2.017.957 4.961L3.964 7.293C4.672 5.166 6.656 3.58 9 3.58z" opacity=".9"/>
          </svg>
          Sign in with Google to continue
        </button>
        <p class="privacy-note">🔒 Your data stays in your Google account. We never store or access it.</p>
      </form>
    </div>

    <script>
    // ── Logo resize client-side ───────────────────────────────
    const drop = document.getElementById('drop');
    const fileInput = document.getElementById('fileInput');
    const logoPreview = document.getElementById('logoPreview');
    const previewImg = document.getElementById('previewImg');
    const previewName = document.getElementById('previewName');
    const logoInput = document.getElementById('logoInput');
    const removeBtn = document.getElementById('removeBtn');

    function processFile(file) {
      if (!file || !file.type.startsWith('image/')) return;
      const reader = new FileReader();
      reader.onload = e => {
        const img = new Image();
        img.onload = () => {
          const MAX_W = 400, MAX_H = 120;
          let w = img.width, h = img.height;
          if (w > MAX_W) { h = Math.round(h * MAX_W / w); w = MAX_W; }
          if (h > MAX_H) { w = Math.round(w * MAX_H / h); h = MAX_H; }
          const c = document.createElement('canvas');
          c.width = w; c.height = h;
          c.getContext('2d').drawImage(img, 0, 0, w, h);
          const url = c.toDataURL('image/png');
          logoInput.value = url;
          previewImg.src = url;
          previewName.textContent = file.name + ' · ' + w + '×' + h + 'px';
          logoPreview.style.display = 'flex';
          drop.style.display = 'none';
        };
        img.src = e.target.result;
      };
      reader.readAsDataURL(file);
    }

    fileInput.addEventListener('change', e => processFile(e.target.files[0]));
    removeBtn.addEventListener('click', () => {
      logoInput.value = ''; previewImg.src = '';
      logoPreview.style.display = 'none';
      drop.style.display = 'block';
      fileInput.value = '';
    });
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
    drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('drag'); processFile(e.dataTransfer.files[0]); });

    // ── Disable button on submit ──────────────────────────────
    document.getElementById('form').addEventListener('submit', () => {
      const b = document.getElementById('submitBtn');
      b.disabled = true;
      b.innerHTML = '<span style="font-size:18px;animation:spin 1s linear infinite;display:inline-block">⟳</span> Redirecting to Google…';
    });
    </script>
  `);
}

// ============================================================
// WORKING PAGE — shown while provisioning runs
// ============================================================

function workingPage(email) {
  return page('Setting up your workspace…', `
    <div class="working-wrap">
      <div class="spinner"></div>
      <h1>Setting up your workspace</h1>
      <p class="sub">Signed in as <strong>${email}</strong><br>
      Creating your proposals folder, templates, and web app…<br>
      This takes about 30 seconds.</p>
      <div class="progress-steps" id="steps">
        <div class="ps">📁 Creating proposals folder</div>
        <div class="ps">📄 Building document templates</div>
        <div class="ps">⚙️  Deploying your private web app</div>
        <div class="ps">✅ Almost done…</div>
      </div>
    </div>
    <script>
    // Animate steps
    const steps = document.querySelectorAll('.ps');
    let i = 0;
    steps[0].classList.add('active');
    const t = setInterval(() => {
      if (i < steps.length - 1) {
        steps[i].classList.remove('active');
        steps[i].classList.add('done');
        i++;
        steps[i].classList.add('active');
      } else { clearInterval(t); }
    }, 6000);
    </script>
  `);
}

// ============================================================
// DONE PAGE — shown after successful provisioning
// ============================================================

function donePage(result) {
  const deployedBlock = result.deployed ? `
    <div class="url-card">
      <div class="url-label">Your Web App URL</div>
      <div class="url-value" id="urlVal">${result.webAppUrl}</div>
      <button class="copy-btn" onclick="copyUrl()">Copy</button>
    </div>
    <p class="url-hint">Bookmark this on your phone — it's your proposal app. Share it with anyone on your team.</p>
  ` : `
    <div class="manual-deploy">
      <div class="manual-icon">⚡</div>
      <div>
        <strong>One final step — deploy in 30 seconds</strong>
        <p>Your script is ready. Open it and click Deploy → New Deployment → Web App.<br>
        Set <em>Execute as: Me</em> and <em>Who has access: Anyone</em>, then copy the URL.</p>
        <a href="${result.scriptUrl}" target="_blank" class="btn btn-primary" style="margin-top:10px;display:inline-flex">Open Apps Script Editor →</a>
      </div>
    </div>
  `;

  const logoHtml = result.logoUrl
    ? `<div class="logo-row"><img src="${result.logoUrl}" alt="logo" onerror="this.parentElement.style.display='none'"></div>` : '';

  return page('You\'re all set! 🎉', `
    <div class="done-wrap">
      <div class="done-head">
        <div class="done-tick">🎉</div>
        <h1>You're all set!</h1>
        <p class="sub">Everything has been created in <strong>${result.email}</strong>'s Google account.</p>
      </div>

      ${logoHtml}
      ${deployedBlock}

      <div class="resources">
        <a class="res-item" href="${result.spreadsheetUrl}" target="_blank">
          <span>📊</span>
          <div><strong>Google Sheet</strong><small>Catalog + Leads — your data lives here</small></div>
          <span class="arr">→</span>
        </a>
        <a class="res-item" href="${result.folderUrl}" target="_blank">
          <span>📁</span>
          <div><strong>Proposals Folder</strong><small>Generated PDFs land here automatically</small></div>
          <span class="arr">→</span>
        </a>
        <a class="res-item" href="${result.proposalTemplateUrl}" target="_blank">
          <span>📄</span>
          <div><strong>Proposal Template</strong><small>Open to add logo, branding, formatting</small></div>
          <span class="arr">→</span>
        </a>
        <a class="res-item" href="${result.invoiceTemplateUrl}" target="_blank">
          <span>🧾</span>
          <div><strong>Invoice Template</strong><small>Open to style — used for invoice generation</small></div>
          <span class="arr">→</span>
        </a>
      </div>

      <div class="privacy-confirm">
        🔒 This platform has no access to your data. Everything above belongs to your Google account.
      </div>

      <a href="/" class="btn btn-secondary" style="margin-top:20px;display:block;text-align:center">Set up another account</a>
    </div>

    <script>
    function copyUrl() {
      const url = document.getElementById('urlVal').textContent;
      navigator.clipboard.writeText(url).then(() => {
        document.querySelector('.copy-btn').textContent = 'Copied!';
        setTimeout(() => document.querySelector('.copy-btn').textContent = 'Copy', 2000);
      });
    }
    </script>
  `);
}

// ============================================================
// SHARED PAGE WRAPPER
// ============================================================

function page(title, body) {
  return `<!DOCTYPE html><html lang="en"><head>
  <meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${title}</title>
  <style>
    @keyframes spin { to { transform: rotate(360deg); } }
    @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.5} }
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
           background: #f1f5f9; min-height: 100vh; padding: 32px 16px; color: #111; }
    .wrap { max-width: 560px; margin: 0 auto; }

    /* Top bar */
    .topbar { text-align: center; margin-bottom: 28px; }
    .topbar-brand { font-size: 15px; font-weight: 800; color: #1d4ed8; }
    .topbar-brand span { color: #64748b; font-weight: 500; }

    /* Hero */
    .hero { text-align: center; margin-bottom: 24px; }
    .hero-icon { font-size: 40px; margin-bottom: 10px; }
    .hero h1 { font-size: 22px; font-weight: 800; color: #111; line-height: 1.3; }
    .hero-sub { font-size: 14px; color: #64748b; margin-top: 10px; line-height: 1.6; max-width: 420px; margin-inline: auto; }

    /* Form card */
    .form-card { background: white; border-radius: 16px; padding: 28px 24px; box-shadow: 0 4px 24px rgba(0,0,0,.08); }
    .field-group { margin-bottom: 16px; }
    .field-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    label { display: block; font-size: 12px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: .06em; margin-bottom: 5px; }
    .req { color: #ef4444; }
    .opt { font-weight: 400; text-transform: none; letter-spacing: 0; color: #94a3b8; font-size: 11px; }
    input[type=text] { width: 100%; padding: 10px 13px; border: 1.5px solid #e2e8f0; border-radius: 9px; font-size: 14px; color: #111; background: #f8fafc; transition: border .15s; }
    input[type=text]:focus { outline: none; border-color: #3b82f6; background: white; }

    /* Logo drop */
    .logo-drop { border: 2px dashed #e2e8f0; border-radius: 10px; padding: 20px; cursor: pointer; background: #f8fafc; position: relative; transition: all .2s; }
    .logo-drop:hover, .logo-drop.drag { border-color: #3b82f6; background: #eff6ff; }
    .logo-drop input[type=file] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
    .drop-inner { text-align: center; pointer-events: none; }
    .drop-icon { font-size: 24px; margin-bottom: 6px; }
    .drop-text { font-size: 13px; color: #64748b; } .drop-text strong { color: #3b82f6; }
    .drop-hint { font-size: 11px; color: #94a3b8; margin-top: 3px; }
    .logo-preview { display: none; align-items: center; gap: 12px; background: #f0fdf4; border: 1px solid #86efac; border-radius: 9px; padding: 10px 14px; margin-top: 8px; }
    .logo-preview img { max-height: 44px; max-width: 140px; object-fit: contain; border-radius: 4px; }
    .preview-meta { flex: 1; font-size: 12px; color: #166534; }
    .remove-btn { background: none; border: none; cursor: pointer; color: #dc2626; font-size: 12px; font-weight: 700; padding: 4px 8px; border-radius: 5px; }
    .remove-btn:hover { background: #fef2f2; }

    .divider { border: none; border-top: 1px solid #f1f5f9; margin: 20px 0; }

    /* Buttons */
    .btn { display: inline-flex; align-items: center; justify-content: center; gap: 10px; padding: 14px 22px; border-radius: 11px; font-size: 15px; font-weight: 700; cursor: pointer; border: none; text-decoration: none; transition: background .15s; width: 100%; }
    .btn-primary { background: #2563eb; color: white; }
    .btn-primary:hover { background: #1d4ed8; }
    .btn-google { background: #2563eb; color: white; }
    .btn-google:hover { background: #1d4ed8; }
    .btn-google:disabled { background: #93c5fd; cursor: not-allowed; }
    .btn-secondary { background: #f1f5f9; color: #475569; border: 1px solid #e2e8f0; }
    .btn-secondary:hover { background: #e2e8f0; }
    .privacy-note { font-size: 11px; color: #94a3b8; text-align: center; margin-top: 10px; }

    /* Errors */
    .error-box { background: #fef2f2; border: 1px solid #fca5a5; border-radius: 9px; padding: 11px 14px; font-size: 13px; color: #dc2626; margin-bottom: 16px; }

    /* Developer gate */
    .dev-card { background: #0f172a; border: 1px solid #1e293b; border-radius: 14px; padding: 28px 26px; color: #cbd5e1; }
    .dev-card h1 { color: #f1f5f9; font-size: 18px; margin-bottom: 8px; }
    .dev-badge { background: #7c3aed; color: white; font-size: 10px; font-weight: 800; text-transform: uppercase; letter-spacing: .1em; padding: 3px 9px; border-radius: 20px; display: inline-block; margin-bottom: 12px; }
    .dev-card .sub { font-size: 13px; color: #64748b; margin-bottom: 20px; line-height: 1.6; }
    .steps { margin-bottom: 20px; }
    .step { display: flex; gap: 12px; padding: 7px 0; font-size: 13px; color: #94a3b8; line-height: 1.5; align-items: flex-start; }
    .step span { width: 22px; height: 22px; border-radius: 50%; background: #3b82f6; color: white; font-size: 11px; font-weight: 800; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 1px; }
    .step a { color: #60a5fa; } .step strong { color: #e2e8f0; }
    .dev-card label { color: #64748b; font-size: 11px; text-transform: uppercase; letter-spacing: .07em; margin-top: 14px; margin-bottom: 5px; display: block; }
    .dev-card input { width: 100%; padding: 10px 12px; background: #0f172a; border: 1px solid #334155; border-radius: 8px; font-size: 13px; color: #f1f5f9; }
    .dev-card input:focus { outline: none; border-color: #3b82f6; }
    .dev-card .btn { margin-top: 16px; }
    .dev-note { font-size: 11px; color: #334155; margin-top: 12px; text-align: center; }
    code { background: rgba(255,255,255,.08); padding: 1px 5px; border-radius: 4px; font-family: monospace; font-size: 11px; }
    .dev-card code { background: #0f172a; border: 1px solid #1e293b; color: #7dd3fc; }

    /* Working page */
    .working-wrap { background: white; border-radius: 16px; padding: 40px 28px; text-align: center; box-shadow: 0 4px 24px rgba(0,0,0,.08); }
    .spinner { width: 48px; height: 48px; border: 4px solid #e2e8f0; border-top-color: #2563eb; border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto 20px; }
    .working-wrap h1 { font-size: 20px; font-weight: 800; margin-bottom: 8px; }
    .working-wrap .sub { font-size: 13px; color: #64748b; line-height: 1.7; }
    .progress-steps { margin-top: 24px; text-align: left; background: #f8fafc; border-radius: 10px; padding: 14px 18px; }
    .ps { font-size: 13px; color: #94a3b8; padding: 5px 0; display: flex; gap: 8px; align-items: center; }
    .ps.active { color: #2563eb; font-weight: 600; animation: pulse 1.5s infinite; }
    .ps.done { color: #16a34a; }

    /* Done page */
    .done-wrap { }
    .done-head { background: white; border-radius: 16px; padding: 28px; text-align: center; box-shadow: 0 4px 24px rgba(0,0,0,.08); margin-bottom: 16px; }
    .done-tick { font-size: 48px; margin-bottom: 10px; }
    .done-head h1 { font-size: 24px; font-weight: 800; color: #16a34a; }
    .done-head .sub { font-size: 13px; color: #64748b; margin-top: 6px; }
    .logo-row { display: flex; justify-content: center; margin-bottom: 16px; }
    .logo-row img { max-height: 52px; max-width: 200px; object-fit: contain; background: white; border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px 16px; }
    .url-card { background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 11px; padding: 16px 18px; display: flex; align-items: center; gap: 12px; margin-bottom: 8px; }
    .url-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .07em; color: #1e40af; margin-bottom: 4px; }
    .url-value { font-size: 12px; color: #1e3a8a; word-break: break-all; flex: 1; font-family: monospace; }
    .copy-btn { background: #2563eb; color: white; border: none; border-radius: 8px; padding: 8px 14px; font-size: 12px; font-weight: 700; cursor: pointer; flex-shrink: 0; }
    .copy-btn:hover { background: #1d4ed8; }
    .url-hint { font-size: 12px; color: #64748b; margin-bottom: 16px; }
    .manual-deploy { display: flex; gap: 14px; background: #fefce8; border: 1px solid #fde047; border-radius: 11px; padding: 16px 18px; margin-bottom: 16px; align-items: flex-start; }
    .manual-icon { font-size: 24px; flex-shrink: 0; }
    .manual-deploy strong { display: block; margin-bottom: 4px; }
    .manual-deploy p { font-size: 13px; color: #713f12; line-height: 1.5; }
    .resources { display: flex; flex-direction: column; gap: 8px; margin-bottom: 12px; }
    .res-item { display: flex; align-items: center; gap: 12px; background: white; border: 1px solid #e2e8f0; border-radius: 11px; padding: 13px 16px; text-decoration: none; color: #111; transition: border-color .15s; }
    .res-item:hover { border-color: #3b82f6; }
    .res-item > span:first-child { font-size: 22px; flex-shrink: 0; }
    .res-item div { flex: 1; }
    .res-item strong { display: block; font-size: 14px; }
    .res-item small { font-size: 12px; color: #64748b; }
    .arr { color: #94a3b8; font-size: 16px; }
    .privacy-confirm { background: #f0fdf4; border: 1px solid #86efac; border-radius: 10px; padding: 12px 16px; font-size: 12px; color: #166534; text-align: center; }
  </style>
</head><body>
<div class="wrap">
  <div class="topbar"><div class="topbar-brand">🏠 Roofing Proposal <span>Platform</span></div></div>
  ${body}
</div>
</body></html>`;
}

// ============================================================
// ROUTES
// ============================================================

// ── Developer gate ────────────────────────────────────────────

app.get('/dev-setup', (req, res) => res.send(devGatePage()));

app.post('/dev-setup', (req, res) => {
  const { clientId, clientSecret } = req.body;
  if (!clientId?.trim() || !clientSecret?.trim()) {
    return res.send(devGatePage('Both fields are required.'));
  }
  fs.writeFileSync(
    auth.APP_CREDS_PATH,
    JSON.stringify({ clientId: clientId.trim(), clientSecret: clientSecret.trim() }, null, 2)
  );
  console.log('\n✅ app-credentials.json saved\n');
  res.redirect('/');
});

// ── Landing ───────────────────────────────────────────────────

app.get('/', (req, res) => {
  if (!auth.hasAppCreds()) return res.redirect('/dev-setup');
  res.send(onboardingPage());
});

// ── Collect company info → store in session → go to Google ───

app.post('/onboard', (req, res) => {
  if (!auth.hasAppCreds()) return res.redirect('/dev-setup');

  const { companyName, companyPhone, companyEmail, companyLicense, companyTagline, logoBase64 } = req.body;
  if (!companyName?.trim()) {
    return res.send(onboardingPage('Company name is required.'));
  }

  // Store form data in memory under a random state token
  const state = crypto.randomBytes(16).toString('hex');
  pendingSessions.set(state, {
    companyName:    companyName.trim(),
    companyPhone:   (companyPhone   || '').trim(),
    companyEmail:   (companyEmail   || '').trim(),
    companyLicense: (companyLicense || '').trim(),
    companyTagline: (companyTagline || '').trim(),
    logoBase64:     logoBase64 || null,
  });

  // Redirect to Google OAuth, passing state so we get it back in callback
  const url = auth.getClient().generateAuthUrl({
    access_type: 'offline',
    scope: auth.SCOPES,
    prompt: 'consent',
    state,
  });
  res.redirect(url);
});

// ── OAuth callback → kick off provisioning, redirect to poll page

app.get('/auth/callback', async (req, res) => {
  const { code, state, error } = req.query;

  if (error) return res.send(onboardingPage(`Google denied access: ${error}`));
  const sessionData = pendingSessions.get(state);
  if (!sessionData) return res.send(onboardingPage('Session expired — please try again.'));
  pendingSessions.delete(state);

  let email = '';
  let client;
  try {
    client = auth.getClient();
    const { tokens } = await client.getToken(code);
    client.setCredentials(tokens);
    const { google } = require('googleapis');
    const info = await google.oauth2({ version: 'v2', auth: client }).userinfo.get();
    email = info.data.email;
  } catch (err) {
    return res.send(onboardingPage('Authentication failed: ' + err.message));
  }

  const jobId = crypto.randomBytes(8).toString('hex');
  provisionJobs.set(jobId, { status: 'running', email });

  // Fire-and-forget provisioning
  provision(client, sessionData)
    .then(result => {
      result.email = email;
      provisionJobs.set(jobId, { status: 'done', result });
      console.log(`\n✅ ${sessionData.companyName} (${email}) — provisioned`);
      if (result.webAppUrl) console.log(`   🔗 ${result.webAppUrl}\n`);
    })
    .catch(err => {
      provisionJobs.set(jobId, { status: 'error', message: err.message });
      console.error(`\n❌ Provision failed: ${err.message}\n`);
    });

  res.redirect(`/working/${jobId}?email=${encodeURIComponent(email)}`);
});

// ── Working page (polls until done) ──────────────────────────

app.get('/working/:jobId', (req, res) => {
  const email = req.query.email || '';
  res.send(page('Setting up your workspace…', `
    <div class="working-wrap">
      <div class="spinner"></div>
      <h1>Setting up your workspace</h1>
      <p class="sub">Signed in as <strong>${email}</strong><br>
      Creating your proposals folder, templates, and web app…<br>
      This takes about 30 seconds.</p>
      <div class="progress-steps">
        <div class="ps active" id="s1">📁 Creating proposals folder &amp; templates</div>
        <div class="ps" id="s2">📊 Setting up your Google Sheet</div>
        <div class="ps" id="s3">⚙️ Deploying your private web app</div>
        <div class="ps" id="s4">🎉 Almost done…</div>
      </div>
    </div>
    <script>
    const steps = ['s1','s2','s3','s4'];
    let cur = 0;
    const animate = setInterval(() => {
      if (cur < steps.length - 1) {
        document.getElementById(steps[cur]).className = 'ps done';
        cur++;
        document.getElementById(steps[cur]).className = 'ps active';
      }
    }, 7000);

    // Poll for completion
    const poll = setInterval(async () => {
      const r = await fetch('/api/job/${req.params.jobId}');
      const d = await r.json();
      if (d.status === 'done') {
        clearInterval(poll); clearInterval(animate);
        window.location = '/done/${req.params.jobId}';
      } else if (d.status === 'error') {
        clearInterval(poll); clearInterval(animate);
        document.querySelector('h1').textContent = 'One more step needed';
        document.querySelector('.spinner').style.display = 'none';

        let msg = d.message || '';
        let extra = '';
        if (msg.includes('Apps Script API') || msg.includes('script.googleapis.com')) {
          extra = '<div style="margin-top:16px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:14px 16px;font-size:13px;line-height:1.7">'
                + '⚙️ You need to enable the Apps Script API in your Google account:<br><br>'
                + '<a href="https://script.google.com/home/usersettings" target="_blank" '
                + 'style="background:#2563eb;color:white;padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:700;display:inline-block">'
                + 'Enable Apps Script API →</a><br><br>'
                + 'After enabling it, come back and <a href="/">try again</a>.</div>';
        } else {
          extra = '<br><small style="color:#94a3b8">' + msg + '</small><br><br><a href="/">Try again</a>';
        }
        document.querySelector('.sub').innerHTML = 'Almost there — just one quick setting to enable.' + extra;
      }
    }, 3000);
    </script>
  `));
});

app.get('/api/job/:jobId', (req, res) => {
  const job = provisionJobs.get(req.params.jobId) || { status: 'not_found' };
  res.json(job);
});

app.get('/done/:jobId', (req, res) => {
  const job = provisionJobs.get(req.params.jobId);
  if (!job || job.status !== 'done') return res.redirect('/');
  res.send(donePage(job.result));
  // Clean up job from memory (token was already released with the client object)
  setTimeout(() => provisionJobs.delete(req.params.jobId), 60000);
});

// ============================================================
// START
// ============================================================

app.listen(PORT, () => {
  console.log(`\n🏠 Roofing Proposal Platform`);
  console.log(`   http://localhost:${PORT}\n`);
  if (!auth.hasAppCreds()) {
    console.log(`   ⚠️  Developer setup needed`);
    console.log(`   → http://localhost:${PORT}/dev-setup\n`);
  } else {
    console.log(`   ✅ Ready — contractors can onboard at this URL\n`);
  }
});
