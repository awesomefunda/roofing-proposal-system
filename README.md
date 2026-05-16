# Roofing Proposal & E-Signature System

A free, self-contained roofing proposal and e-signature tool built entirely on Google Workspace. No server, no monthly fees, no third-party subscriptions. Everything runs inside your own Google account.

> **Built for small roofing contractors** who want professional proposals and legally-binding digital signatures without paying $25–45/month for DocuSign or PandaDoc.

---

## Why this instead of DocuSign / PandaDoc?

| | This system | DocuSign | PandaDoc |
|--|--|--|--|
| Monthly cost | **$0** | $25–45/mo | $19–49/mo |
| Your data lives | **Your Google Drive** | DocuSign servers | PandaDoc servers |
| Vendor lock-in | **None** | High | High |
| Setup time | ~30 min | Minutes | Minutes |
| Custom branding | **Full control** | Limited on free | Limited on free |
| Audit trail | **Yes (IP, hash, timestamp)** | Yes | Yes |
| Legal basis | **ESIGN Act + UETA** | ESIGN Act | ESIGN Act |

The tradeoff: this requires a one-time setup and is self-maintained. You own the code, the data, and the templates.

---

## What it does

- **Create proposals** from a reusable catalog of line items — one click, no re-typing
- **Email the client** a branded HTML email with a secure signing link (PDF attached)
- **Client signs electronically** via a canvas signature on any device — no app, no login required
- **Email KBA** (Knowledge-Based Authentication) — client must enter their email before they can see or sign anything
- **Tamper-proof PDFs** — SHA-256 hash computed at creation; signing is blocked if the document changed
- **Certificate of Completion** — appended automatically (timestamp, IP address, browser, document hash, signature image)
- **Legal compliance** — ESIGN Act (15 U.S.C. § 7001) and UETA compliant
- **Generate invoices** from signed proposals — no re-entry needed
- **Historical pricing** — loading a previous estimate uses the original prices, not today's catalog rates
- **Contractor PIN gate** — protects the form from casual access

---

## Architecture

Everything runs in your Google account. No external servers, no API keys, no webhooks.

```
Contractor's browser
    ↕  (google.script.run)
Google Apps Script  ←→  Google Sheets   (Leads + Catalog tabs)
        ↕                Google Docs     (proposal + invoice templates)
        ↕                Google Drive    (signed PDFs, folders)
        ↕                Gmail           (client emails, owner notifications)
```

| Layer | Technology | Why |
|-------|-----------|-----|
| Backend | Google Apps Script | Free, runs in your account, no server |
| Database | Google Sheets | Easy to inspect, edit, export |
| Document engine | Google Docs → PDF | Native quality, full colour, free |
| File storage | Google Drive | Automatic, searchable, shareable |
| Email | Gmail | No SMTP setup, 500 emails/day free |
| Hosting | GAS Web App (`/exec`) | Free, HTTPS, no server management |
| E-sig legal basis | ESIGN Act + UETA | Federal + state law (all 50 states) |

---

## Project structure

```
gas/
  Code.gs         — All backend logic (proposals, signing, PDF, email, Drive, settings)
  Index.html      — Contractor UI (PIN gate, proposal form, leads tab, settings tab)
  Sign.html       — Customer signing page (email KBA → signature canvas → submit)
  appsscript.json — OAuth scopes and web app configuration

Roofing_Catalog_Template.xlsx  — Import into your Google Sheet to populate the Catalog tab
```

---

## How onboarding works

### Prerequisites
- A Google account (free Gmail works fine)
- [Node.js](https://nodejs.org/) (LTS) installed
- [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`

---

### Step 1 — Create your Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) → **Blank spreadsheet**
2. Name it: `Roofing Proposals — [Your Company Name]`
3. Create two tabs: **Leads** and **Catalog**
4. In the **Catalog** tab, set row 1 headers:
   ```
   Item | Description | Price | Unit | Category | Active
   ```
5. Import `Roofing_Catalog_Template.xlsx` (File → Import → Upload → Replace current sheet), or manually add your line items. Set column F (`Active`) to `Yes` for items you want in the proposal form.
6. Copy the **Sheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/[COPY_THIS]/edit
   ```

The Leads tab is auto-managed by the script — do not add headers manually.

---

### Step 2 — Create the Apps Script project

1. In your spreadsheet: **Extensions → Apps Script**
2. Note the URL of the editor:
   ```
   https://script.google.com/home/projects/[YOUR_SCRIPT_ID]/edit
   ```
   Copy the Script ID.

---

### Step 3 — Push the code with clasp

```bash
# Log in to Google (opens browser auth)
clasp login

# Clone the project (links this repo to your Apps Script project)
cd roofing-proposal-system
clasp clone YOUR_SCRIPT_ID --rootDir ./gas

# Push all files
clasp push --force
```

> Every time you edit files locally, run `clasp push --force` to upload.

---

### Step 4 — Authorize the script

In the Apps Script editor:

1. Select `getAppState` from the function dropdown
2. Click **Run**
3. Google prompts: **Authorization Required → Review Permissions → Allow**

This grants the script access to Sheets, Docs, Drive, and Gmail.

---

### Step 5 — Deploy as a Web App

1. In Apps Script editor: **Deploy → New deployment**
2. Click the ⚙️ gear → **Web app**
3. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone *(clients sign without a Google login)*
4. Click **Deploy** → copy the URL:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```

> ⚠️ Every time you push new code, you must also create a **new deployment version**: Deploy → Manage deployments → Edit → New version → Deploy. The URL stays the same.

---

### Step 6 — Save the Web App URL as a Script Property

This is the step most people miss. Run this **once** in the Apps Script editor after each deployment:

```javascript
function saveWebAppUrl() {
  PropertiesService.getScriptProperties().setProperty(
    'WEBAPP_URL',
    'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec'
  );
  Logger.log('Saved: ' + PropertiesService.getScriptProperties().getProperty('WEBAPP_URL'));
}
```

**Why this matters:** `ScriptApp.getService().getUrl()` returns different values depending on context — in the editor it returns the `/dev` URL (owner-only), and when a logged-in owner visits the web app, Google may redirect to a user-specific `/u/1/s/...` URL that returns 404 for clients. Storing the canonical URL explicitly in Script Properties avoids both problems. All signing links in emails will use this URL.

---

### Step 7 — First-run setup in the app

1. Open the Web App URL in your browser
2. Enter **default PIN: `4766`**
3. The setup wizard appears — fill in company name, owner name, phone, email, license number
4. Click **Save & Create Templates**

The system automatically creates in your Drive:
- A **Proposals folder**
- A **Proposal template** Google Doc (branded with your info)
- An **Invoice template** Google Doc

---

### Step 8 — Add your Sheet ID to Settings

1. In the app, open the **Settings tab**
2. Paste your Sheet ID (from Step 1) into the **Google Sheet ID** field
3. Click **Save Settings**

Setup is complete.

---

## How the customer experience works

### What the client receives

After you submit a proposal, the client gets a branded HTML email:
- Your company name, phone, and email in the header
- The proposal PDF attached (their reference copy, before signing)
- A large **"Review & Sign Your Proposal"** button
- The signing URL in plain text as backup
- Footer with your website link

### The signing flow

**Step 1 — Email verification (KBA)**
The page opens with a single blank field: "Enter the email address we sent this to." No proposal data is visible yet. The client must type the exact email address on file to proceed. This is the Knowledge-Based Authentication step — it proves the right person has access to that inbox.

**Step 2 — Proposal summary**
After passing KBA, the client sees their name, estimate number, and total amount. Clean and unambiguous — no confusion about what they're agreeing to.

**Step 3 — ESIGN consent**
A checkbox: "I agree this electronic signature is legally binding under the ESIGN Act (15 U.S.C. § 7001)." They must check it before the signature canvas activates.

**Step 4 — Draw initials + full signature**
Two canvas fields. Works with a finger on mobile, mouse on desktop. No app download required.

**Step 5 — Submit**
The system:
1. Re-checks the document hash — rejects if the Doc was edited after sending
2. Inserts signature images into the Google Doc at `{{SIGNATURE_1}}`
3. Appends a Certificate of Completion page (timestamp, IP, browser, hash, both signature images)
4. Exports the Doc as PDF → saves to Drive
5. Emails the signed PDF to both the client ("Your Signed Proposal") and you ("✅ SIGNED")

### What the contractor receives

An email: `✅ SIGNED — [Estimate #] · [Client Name]` with the signed PDF attached. The Leads sheet row status changes to `Signed`.

---

## Daily workflow

### Creating a proposal

1. Open the web app URL → enter PIN
2. Fill in client details (name, street/city/state/zip, phone, email)
3. Select catalog items (scope of work)
4. Add any custom line items (description + price)
5. Review the auto-calculated total (or manually override)
6. Click **Create Proposal & Send Signing Link**

### Loading a previous estimate

Use the **Load from Previous Estimate** dropdown. Fills the form with original line items and prices — not current catalog rates. Safe against price inflation.

### Generating an invoice

In the **Leads tab**, find a signed lead → click **+ Invoice**. Created from existing data, no re-entry.

### Changing settings

Go to the **Settings tab** anytime: company info, PIN, Sheet ID, rotted wood rate, web app URL.

---

## After every code update

```bash
# 1. Push new code
clasp push --force

# 2. In Apps Script editor:
#    Deploy → Manage deployments → Edit → New version → Deploy

# 3. If the deployment ID changed, update WEBAPP_URL:
#    Run saveWebAppUrl() in the editor with the new URL
```

---

## Security model

| Threat | Mitigation |
|--------|-----------|
| Guessing signing URLs | UUID v4 tokens — 2¹²² possible values |
| Wrong person signing | Email KBA — must enter exact address on file |
| Document altered after sending | SHA-256 tamper detection — mismatch blocks signing |
| Contractor form accessed by clients | PIN gate — verified server-side |
| Double-submission race condition | LockService with 10-second wait on sheet writes |
| Double signing | Status check — `Signed` status blocks retry |
| Proposal data leaking before auth | Minimal server injection — only `valid`/`token`/`companyName` in page source; full proposal data only after KBA |
| Legal enforceability | ESIGN Act consent checkbox, Certificate of Completion page |

---

## Privacy model

| Data | Where it lives |
|------|---------------|
| Proposals, PDFs, invoices | Your Google Drive |
| Client names, emails, addresses | Your Google Sheet |
| Emails sent | Your Gmail |
| Script execution | Your Google Apps Script |
| Developer/third-party access | **Zero** — no API keys, no external calls, no telemetry |

---

## Legal compliance notes

The signed PDF contains three pages:
1. **Proposal** — scope of work, total, embedded client signature, date signed
2. **Printed name** — client's typed/printed name
3. **Certificate of Completion** — ESIGN Act + UETA citation, signer name, estimate #, amount, timestamp with timezone, IP address, User-Agent (browser + OS), SHA-256 document hash, signature image, initials image

This is sufficient for enforcing scope/price disputes in court. For **California B&P Code § 7159** (home improvement contracts > $500), you should also add approximate start/completion dates and a 3-day right-to-cancel notice to your proposal template.

---

## Known limitations

**Script owner's browser shows "Invalid Link" on signing URLs**
When you (the owner) visit the signing URL while logged into your Google account, Google may redirect to a user-specific `/u/1/s/...` path that returns 404. This only affects you — real customers get the canonical URL and sign without issue. Test signing links in incognito mode or from a different Google account.

**`clasp push` does not deploy**
Pushing code uploads it but does not activate it. You must separately create a new deployment version in the Apps Script editor.

**Gmail daily limit: 500 emails/day**
Free Google accounts. Google Workspace accounts get 1,500/day. Not a practical limit for a roofing contractor.

**Editing a proposal Doc breaks signing**
The SHA-256 hash is locked at proposal creation. If you edit the Doc later, the hash won't match and signing is rejected. Create a new proposal if scope or price changes.

---

## Troubleshooting

**"Invalid or Expired Link"**
Token not found in Leads sheet. Check that the proposal was saved (Leads tab has a row), SHEET_ID is set in Settings, and WEBAPP_URL matches the current deployment.

**Signing works in incognito, not in regular browser**
You are the script owner. Normal behavior. Real customers are unaffected. Use incognito for testing.

**"The document was modified after it was sent"**
The Google Doc was edited after sending. Create a new proposal.

**"Proposal saved, but email could not be sent"**
Gmail quota hit or invalid email. The signing link still works — find the token in the Leads sheet and share it manually.

**Changes not live after `clasp push`**
Create a new deployment version. See "After every code update" above.

**Signing email has wrong deployment URL**
WEBAPP_URL script property points to an old deployment. Run `saveWebAppUrl()` with the current URL.

---

## Catalog setup

`Roofing_Catalog_Template.xlsx` includes 39 pre-priced line items across 8 categories:

- Shingle Roofing (tear-off, 30yr, 50yr, ridge cap, starter strip)
- Flat Roofing (TPO, modified bitumen, EPDM, coating)
- Metal Roofing (standing seam, corrugated, trim)
- Underlayment & Decking (felt, synthetic, plywood, OSB)
- Ventilation & Flashing (ridge vent, soffit, step flashing, drip edge)
- Gutters & Accessories (gutter install, guard, downspout)
- Repairs & Specialty (leak repair, skylight, chimney)
- Permits & Admin (permit, HOA filing, inspection)

Prices are sample values — update to your actual rates. Set column F to `Yes`/`No` to show/hide items in the proposal form.

---

## Customizing the proposal template

Your Proposal template Google Doc lives in the Proposals folder in Drive. Style it freely — fonts, colours, logo, layout. Keep the `{{TAG}}` placeholders intact.

Supported tags:

```
{{CLIENT_NAME}}       {{CLIENT_ADDRESS}}    {{CLIENT_PHONE}}
{{CLIENT_EMAIL}}      {{DATE}}              {{ESTIMATE_NUM}}
{{SCOPE_OF_WORK}}     {{TOTAL}}             {{ROOF_TYPE}}
{{SQ_FT}}             {{WARRANTY}}          {{COLOR}}
{{NOTES}}             {{COMPANY_NAME}}      {{OWNER_NAME}}
{{COMPANY_PHONE}}     {{COMPANY_EMAIL}}     {{COMPANY_LICENSE}}
{{ROTTED_WOOD_RATE}}  {{SIGNATURE_1}}       {{SIGNED_AT}}
```

> ⚠️ `{{SIGNATURE_1}}` must be on its own paragraph line in the template. The script replaces that entire paragraph with the signature image. If it shares a line with other text, the surrounding content will be corrupted.

---

## Leads sheet columns (23 total)

| Col | Field | Notes |
|-----|-------|-------|
| A | Timestamp | Proposal creation time |
| B | Estimate # | Auto-generated, e.g. `OR-260515-02` |
| C | Client Name | |
| D | Client Email | Required for signing link |
| E | Client Phone | |
| F | Client Address | Full address |
| G | Roof Type | |
| H | Sq Ft | |
| I | Warranty | |
| J | Color | |
| K | Total ($) | |
| L | Status | `Awaiting Signature` / `Signed` / `Invoiced` |
| M | Sign Token | UUID v4 — the one-time signing key |
| N | Doc ID | Google Doc ID |
| O | Doc Hash | SHA-256 tamper baseline |
| P | Doc URL | Editable Google Doc URL |
| Q | Signed PDF URL | After signing |
| R | Signed At | Timestamp |
| S | Signer IP | Client IP at signing |
| T | Signer UA | Client browser + OS |
| U | Invoice URL | After invoice generated |
| V | Line Items (JSON) | Historical prices per item |
| W | Notes | Proposal notes |

---

## License

MIT — use it, fork it, adapt it. No attribution required.
