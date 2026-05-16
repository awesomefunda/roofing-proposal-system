# Roofing Proposal System

A free, self-contained roofing proposal and e-signature tool built entirely on Google Workspace. No server, no monthly fees, no third-party apps. Everything runs inside your own Google account.

---

## What it does

- **Create proposals** from a catalog of line items with one click
- **Send a signing link** to the client by email (no DocuSign subscription needed)
- **Client signs electronically** with a legally-binding canvas signature (ESIGN Act compliant)
- **Tamper-proof PDFs** — SHA-256 hash detects any edits between sending and signing
- **Certificate of Completion** embedded in the final signed PDF (audit trail, IP, timestamp)
- **Generate invoices** from existing proposals — no re-entry
- **Load previous estimates** with original quoted prices preserved (inflation-safe)
- **Contractor PIN** protects the form from casual access

---

## Architecture

| Layer | Technology |
|-------|-----------|
| Backend | Google Apps Script (runs in your Google account) |
| Database | Google Sheets (Leads + Catalog tabs) |
| Document engine | Google Docs → PDF (native quality, full colour) |
| File storage | Google Drive |
| Email | Gmail (sent from your address) |
| Hosting | Google Apps Script Web App (free, no server) |
| E-signature legal basis | ESIGN Act 15 U.S.C. § 7001 / UETA |

---

## Project structure

```
gas/
  Code.gs            — All backend logic (proposals, signing, PDF, email, Drive)
  Index.html         — Contractor UI (PIN gate, form, leads tab, settings)
  Sign.html          — Client signing page (email KBA, canvas signature)
  InvoiceTemplate.html  — Legacy (superseded by Google Doc templates)
  ProposalTemplate.html — Legacy (superseded by Google Doc templates)
  Sidebar.html       — Legacy sidebar (superseded by web app)
```

The active files are `Code.gs`, `Index.html`, and `Sign.html`. The three legacy files are kept for reference.

---

## One-time setup (~10 minutes)

### Prerequisites
- A Google account
- [Node.js](https://nodejs.org/) installed (for clasp)
- [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`

---

### Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a **new blank spreadsheet**
2. Rename it to something like `Roofing Proposals — [Your Company]`
3. Add two tabs: **Leads** and **Catalog**
4. In the **Catalog** tab, add these headers in row 1:
   ```
   Item | Description | Price | Unit | Category | Active
   ```
5. Add your standard scope-of-work items (set Active = `Yes` to show in the form)
6. Copy the spreadsheet ID from the URL:
   ```
   https://docs.google.com/spreadsheets/d/[THIS_IS_YOUR_ID]/edit
   ```

---

### Step 2 — Create the Apps Script project

1. In your spreadsheet, go to **Extensions → Apps Script**
2. Note the script project URL — you'll need the **Script ID** from it:
   ```
   https://script.google.com/home/projects/[THIS_IS_YOUR_SCRIPT_ID]/edit
   ```

---

### Step 3 — Push the code with clasp

In your terminal, from the `roofing-proposal-system/` folder:

```bash
# Login to Google (opens browser)
clasp login

# Link to your Apps Script project
clasp clone YOUR_SCRIPT_ID --rootDir ./gas

# Push all files
clasp push --force
```

> **Note:** `clasp push` uploads `Code.gs`, `Index.html`, `Sign.html` (and any other files in `gas/`) to your Apps Script project.

---

### Step 4 — Deploy as a Web App

This step creates the URL that both Omar (contractor) and clients (signing link) use.

1. In the Apps Script editor, click **Deploy → New deployment**
2. Click the ⚙️ gear icon next to "Select type" → choose **Web app**
3. Set the following:
   - **Description:** `Roofing Proposal System v1`
   - **Execute as:** `Me` (your Google account)
   - **Who has access:** `Anyone` ← **required** so clients can open signing links without logging in
4. Click **Deploy**
5. Copy the **Web App URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```
   This is Omar's app URL **and** the base for all client signing links.

> ⚠️ Every time you push new code with `clasp push`, you must create a **new deployment** (or use "Manage deployments → Edit" on an existing one) for the changes to take effect. Clicking "Test deployments" uses the latest code but requires a Google login — not suitable for clients.

---

### Step 5 — First-run setup in the app

1. Open the **Web App URL** in your browser
2. Enter the **default PIN: `4766`** (change it in Settings after setup)
3. You'll see the setup wizard — enter:
   - Company name, owner name, phone, email, license number
4. Click **Save & Create Templates**

The system will automatically create:
- A **Proposals folder** in your Google Drive
- A **Proposal template** Google Doc (with your branding, colours, and `{{TAG}}` placeholders)
- An **Invoice template** Google Doc

Both templates open normally in Google Docs — you can customise fonts, colours, and layout anytime. The `{{TAGS}}` are replaced automatically at proposal time.

---

### Step 6 — Add the Sheet ID to Script Properties

The script needs to know which spreadsheet to write Leads to.

1. In Apps Script, go to **Project Settings → Script Properties**
2. Add a property: `SHEET_ID` = your spreadsheet ID from Step 1
3. Click **Save**

> Alternatively, after completing the setup wizard in the app, the script auto-detects the active spreadsheet if you run it from Extensions → Apps Script while the sheet is open.

---

## Deploying updates

After editing any `.gs` or `.html` file locally:

```bash
# Push changes to Apps Script
clasp push --force

# Then in the Apps Script editor:
# Deploy → Manage deployments → Edit → Version: New version → Deploy
```

The Web App URL stays the same — only the version behind it changes.

---

## Changing the contractor PIN

1. Open the app URL, enter the current PIN
2. Go to **Settings tab**
3. Enter the new PIN in the **Contractor PIN** field
4. Click **Save Settings**

The default PIN is `4766`.

---

## Daily workflow

### Creating a proposal

1. Open the Web App URL → enter PIN
2. Fill in client details (name, address, phone, **email — required for signing link**)
3. Check scope-of-work items from the catalog
4. Add any custom line items
5. Enter the total amount
6. Click **Create Proposal & Send Signing Link**

The system will:
- Duplicate the Proposal template Google Doc into your Drive
- Replace all `{{TAGS}}` with client and project data
- Compute a SHA-256 tamper-detection hash of the document
- Generate a unique UUID signing link (not guessable)
- Email the client a professional HTML email with the signing link
- Notify you (the contractor) with the editable Doc link

### Client signing flow

1. Client receives email → clicks **Review & Sign Proposal**
2. Signing page appears — client enters their **email address** to verify identity
3. After email verification, the full proposal summary appears (total, doc link)
4. Client reads the ESIGN Act consent, checks the agreement checkbox
5. Client draws **initials** and **full signature** with finger or mouse
6. Client clicks **Sign & Accept Proposal**
7. The system:
   - Verifies the document hasn't been modified since sending (SHA-256 check)
   - Inserts the signature image into the Google Doc at the `{{SIGNATURE_1}}` placeholder
   - Appends a **Certificate of Completion** page (timestamp, IP, doc hash, signature images)
   - Exports the signed Doc as PDF → saves to Drive
   - Emails the signed PDF to both the client and you

### Editing an estimate before signing

The editable Doc link is sent to you in the owner notification email and shown in the app's success box. Open it in Google Docs and edit normally — no script needed. If you change the price or scope after sending the signing link, the client's signing attempt will be rejected (tamper detection), and you'll need to create a new proposal.

### Generating an invoice

1. Go to the **Leads tab** in the app
2. Find the signed lead → click **Generate Invoice →**
3. The invoice is created from the existing lead data — no re-entry

### Loading a previous estimate

In the **New Proposal tab**, use the **Load from Previous Estimate** dropdown. The form fills in with the exact original line items and prices — not the current catalog prices. This means old quotes won't change if catalog prices have increased.

---

## The Leads sheet (23 columns)

The script manages this sheet automatically. Columns:

| Col | Field | Notes |
|-----|-------|-------|
| A | Timestamp | Date/time proposal was created |
| B | Estimate # | Auto-generated (e.g. `OR-260515-02`) |
| C | Client Name | |
| D | Client Email | Required for signing link |
| E | Client Phone | |
| F | Client Address | |
| G | Roof Type | |
| H | Sq Ft | |
| I | Warranty | |
| J | Color | |
| K | Total ($) | |
| L | Status | `Awaiting Signature` / `Signed` / `Invoiced` |
| M | Sign Token | UUID v4 — the signing URL token |
| N | Doc ID | Google Doc ID of the editable proposal |
| O | Doc Hash | SHA-256 of doc body text (tamper baseline) |
| P | Doc URL | Editable Doc URL |
| Q | Signed PDF URL | After client signs |
| R | Signed At | ISO timestamp |
| S | Signer IP | Client IP at signing |
| T | Signer UA | Client browser/OS |
| U | Invoice URL | After invoice is generated |
| V | Line Items (JSON) | Raw line items with historical prices |
| W | Notes | Proposal notes |

---

## Customizing templates

After initial setup, both templates live in your **Proposals folder** in Google Drive. Open them in Google Docs and style them however you like. The only constraint: don't delete the `{{TAG}}` placeholders — the script replaces those at runtime.

Supported tags in the Proposal template:

```
{{CLIENT_NAME}}       {{CLIENT_ADDRESS}}    {{CLIENT_PHONE}}
{{CLIENT_EMAIL}}      {{DATE}}              {{ESTIMATE_NUM}}
{{SCOPE_OF_WORK}}     {{TOTAL}}             {{ROOF_TYPE}}
{{SQ_FT}}             {{WARRANTY}}          {{COLOR}}
{{NOTES}}             {{COMPANY_NAME}}      {{OWNER_NAME}}
{{COMPANY_PHONE}}     {{COMPANY_EMAIL}}     {{COMPANY_LICENSE}}
{{ROTTED_WOOD_RATE}}  {{SIGNATURE_1}}       {{SIGNED_AT}}
```

> ⚠️ `{{SIGNATURE_1}}` **must be on its own line** (its own paragraph) in the template. The script clears that paragraph and inserts the signature image there at signing time.

---

## Security model

| Threat | Mitigation |
|--------|-----------|
| Someone guesses signing URL | UUID v4 tokens — 2¹²² possible values, not guessable |
| Client data in page source | Minimal server injection — sensitive data withheld until email KBA |
| Wrong person signs | Email KBA — must enter exact email on file before seeing proposal |
| Document altered after sending | SHA-256 tamper detection — mismatch blocks signing |
| Contractor app accessed by client | PIN gate — verified server-side (PIN never in JS source) |
| Two proposals submitted at once | LockService with 10-second wait on sheet writes |
| Double signing | Status check before processing — `Signed` status blocks re-signing |
| Legal enforceability | ESIGN Act consent checkbox, audit trail (IP, UA, timestamp, hash) in Certificate of Completion |

---

## Privacy model

| Data | Where it lives |
|------|---------------|
| Proposals & PDFs | Your Google Drive |
| Leads & client data | Your Google Sheet |
| Email sending | Your Gmail |
| Script execution | Your Google account |
| Developer access | **Zero** — no API keys, no external server |

---

## Troubleshooting

**"This signing link is not valid"**
→ Token not found in Leads sheet. Check that the lead was created successfully (the Leads tab should show a row with a Sign Token).

**"The document was modified after it was sent"**
→ Someone edited the editable Doc between sending and signing. Create a new proposal.

**"Proposal saved, but email could not be sent"**
→ Gmail daily send quota hit (500 emails/day on free accounts). The signing link still works — copy it from the success box and share manually.

**Signature image not inserted in PDF**
→ `{{SIGNATURE_1}}` must be on its own paragraph line in the Proposal template Doc. Check the template.

**Changes not live after `clasp push`**
→ You must also create a new deployment version in the Apps Script editor after pushing.

---

## License

MIT © 2026 Harsh Kumar
