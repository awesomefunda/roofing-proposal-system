# Roofing Proposal System

A free, self-contained proposal and invoice tool for roofing contractors — built entirely on Google Workspace. No server, no subscriptions, no third-party access to your data.

---

## How it works

The system is a Google Sheet with an embedded Apps Script. When a contractor makes a copy of the template sheet, everything runs **inside their own Google account**:

- PDFs are saved to **their** Google Drive
- Emails are sent from **their** Gmail
- The developer has **zero access** to any data

---

## Setup (one time, ~3 minutes)

### Step 1 — Make a copy of the template sheet

Click this link (replace `SHEET_ID` with your published template ID):

```
https://docs.google.com/spreadsheets/d/SHEET_ID/copy
```

Google will copy the Sheet — including the embedded script — directly into your Drive.

### Step 2 — Open the sidebar

In your Google Sheet, click:

> **Extensions → Apps Script → Run → onOpen** (first time only, to trigger the menu)

Then in the sheet:

> **🏠 Roofing Tools → ✏️ New Proposal**

Google will ask you to **Allow** the script to access your Drive and Gmail. This is a one-time, 30-second step — click Allow.

### Step 3 — Enter your company details

Fill in your company name, phone, email, and license number. Click **Save & Continue**. Done.

---

## Daily use

1. Open your Google Sheet
2. Click **🏠 [Your Company] → ✏️ New Proposal**
3. Fill in client details and check off scope of work items
4. Click **Send Proposal to Client**

The client receives a professional PDF by email. A copy is saved to your Drive. The lead is logged in the Leads sheet automatically.

To generate an invoice for an existing job:
- Click the row in the Leads sheet
- Click **🏠 [Your Company] → 🧾 Invoice Selected Row**

---

## Customizing the catalog

Open the **Catalog** sheet tab. Add, edit, or remove line items. The sidebar picks them up automatically — no code changes needed.

| Column | Description |
|--------|-------------|
| Item | Short key (e.g. `shingle_30yr`) |
| Description | Text shown on the PDF |
| Price | Dollar amount (0 = Included) |
| Unit | Unit label (e.g. `per job`) |
| Category | Groups items in the sidebar |
| Active | `Yes` to show, anything else to hide |

---

## Sharing with someone (like a friend or colleague)

1. Create your Google Sheet template with the script embedded using [clasp](https://github.com/google/clasp) or manually via the Apps Script editor
2. Share the `/copy` link: `https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/copy`
3. They click the link, follow the 3-minute setup above
4. They are fully independent — you have zero access to their data

---

## Project structure

```
gas/
  Code.gs       — All server-side logic (PDF generation, email, Drive, Sheets)
  Sidebar.html  — The sidebar UI (form, catalog checkboxes, settings tab)
```

That's the entire system. No Node.js server, no Railway, no OAuth app registration.

---

## Privacy model

| What | Where it lives |
|------|---------------|
| Proposals folder | Contractor's Google Drive |
| PDF files | Contractor's Google Drive |
| Leads data | Contractor's Google Sheet |
| Email sending | Contractor's Gmail |
| Script execution | Contractor's Google account |
| Developer access | **None — zero** |

---

## License

MIT © 2026 Harsh Kumar
