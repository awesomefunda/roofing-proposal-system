# Roofing Proposal Platform

A self-service SaaS onboarding platform for roofing contractors. Each contractor signs in with Google and gets their own private proposal system — folders, templates, a Google Sheet, and a deployed web app — all inside their own Google account. The platform owner has zero access to contractor data.

---

## How It Works

### The Big Picture

```
Contractor visits platform
        ↓
Fills in company details + logo
        ↓
Clicks "Sign in with Google"
        ↓
Platform creates everything in their Google account (~30 sec)
        ↓
Contractor gets their private Web App URL
        ↓
Platform is never involved again
```

### What Gets Created (in the contractor's Google account)

| Resource | Purpose |
|---|---|
| Google Sheet | Catalog of line items + Leads tracker |
| Proposals Folder | Generated proposal PDFs land here |
| Templates Folder | Holds proposal/invoice templates + logo |
| Proposal Template Doc | Editable — contractor can add branding |
| Invoice Template Doc | Editable — used for invoice generation |
| Apps Script Project | The web app backend (runs as contractor) |
| Web App Deployment | The URL the contractor bookmarks on their phone |

---

## Privacy Model

### Platform owner has zero ongoing access

The contractor's OAuth token is used **once** during the 30-second provisioning window to create their resources. It is held only in server memory and is never written to disk, logged, or stored anywhere. After provisioning completes, the token is garbage collected.

After setup, the contractor's system is completely independent:

- The Apps Script runs **as the contractor** (`executeAs: USER_DEPLOYING`)
- All Drive/Sheet/Gmail operations use **their** Google account
- There is no callback, webhook, or API call back to this platform
- The platform owner cannot read their leads, proposals, or any data

### Why the platform needs OAuth credentials at all

Google OAuth requires a registered app (Client ID + Secret) to act as the authentication entry point. This is how every "Sign in with Google" button works — one registered app, many users, each with their own isolated data. The platform credentials are only the **door** — they do not grant any ongoing access to what's behind it.

---

## Auth Flows Explained

There are two completely separate authorization steps:

### 1. Onboarding (platform side)
When a contractor clicks "Sign in with Google" on the onboarding form:
- They authenticate through the **platform's** Google OAuth app
- This grants a temporary token used to create their Drive resources
- Token is discarded after ~30 seconds
- This uses the platform's `GOOGLE_CLIENT_ID` + `GOOGLE_CLIENT_SECRET`

### 2. First Web App Open (contractor side)
When a contractor opens their Web App URL for the first time:
- Google asks them to authorize **their own Apps Script** to access their Drive/Sheets/Gmail
- This has nothing to do with the platform credentials
- It's Google saying: "this script in your account wants permissions — do you allow it?"
- They click Allow once, and never see this screen again

---

## Developer Setup

### Prerequisites
- Node.js 18+
- A Google Cloud project with these APIs enabled:
  - Google Drive API
  - Google Sheets API
  - Google Apps Script API
- An OAuth 2.0 Client ID (Web application type)
  - Redirect URI: `http://localhost:3000/auth/callback` (local) or your deployed URL

### Local Development

```bash
# Install dependencies
npm install

# Copy and fill in your OAuth credentials
cp local/app-credentials.json.example local/app-credentials.json
# Edit local/app-credentials.json with your real Client ID + Secret

# Start the server
npm run dev
# → http://localhost:3000
```

### Environment Variables (Production)

| Variable | Description |
|---|---|
| `GOOGLE_CLIENT_ID` | OAuth Client ID from Google Cloud Console |
| `GOOGLE_CLIENT_SECRET` | OAuth Client Secret |
| `REDIRECT_URI` | Full callback URL e.g. `https://your-app.railway.app/auth/callback` |

### Deployment (Railway)

1. Push to GitHub
2. Connect repo to [railway.app](https://railway.app)
3. Add the three environment variables above in Railway → Service → Variables
4. Generate a public domain in Railway → Service → Settings → Networking
5. Add the domain's `/auth/callback` URL to Google Cloud Console → Credentials → Authorized redirect URIs

---

## Google Cloud Console Setup

### One-time steps (developer does this)

1. Create a project at [console.cloud.google.com](https://console.cloud.google.com)
2. Enable APIs: Drive, Sheets, Apps Script
3. Create OAuth 2.0 Client ID (Web application)
4. Go to Google Auth Platform → Audience:
   - User type: External
   - Add your email as a test user
5. Go to Data Access → add these scopes:
   - `https://www.googleapis.com/auth/drive`
   - `https://www.googleapis.com/auth/spreadsheets`
   - `https://www.googleapis.com/auth/script.projects`
   - `https://www.googleapis.com/auth/script.deployments`
   - `https://www.googleapis.com/auth/userinfo.email`

### For real contractors (before going live)

The Apps Script scopes (`script.projects`, `script.deployments`) are **restricted scopes** — Google requires a security review before non-test users can use them. Before onboarding real contractors:
- Submit your app for Google verification via Google Auth Platform → Verification Center
- Until verified, only emails added as test users can go through onboarding

---

## Contractor Requirements

Each contractor needs to do **one thing** before their system is fully ready:

**Enable the Apps Script API in their own Google account:**
→ [https://script.google.com/home/usersettings](https://script.google.com/home/usersettings)

The platform error page shows a clickable button for this automatically if it's not enabled.

---

## Project Structure

```
roofing-proposal-system/
├── local/
│   ├── server.js              # Express onboarding platform
│   ├── auth.js                # Google OAuth2 client + credential loading
│   ├── provision.js           # Core provisioning — creates all Drive resources
│   ├── app-credentials.json   # Your OAuth Client ID + Secret (gitignored)
│   └── app-credentials.json.example  # Template — fill and copy
├── gas/
│   ├── Code.gs                # Apps Script backend (%%BOOTSTRAP_JSON%% injected at deploy)
│   └── Index.html             # Web app frontend
├── railpack.toml              # Railway deployment config
└── package.json
```

### How Code.gs templating works

`gas/Code.gs` is a template file. The `%%BOOTSTRAP_JSON%%` placeholder is replaced by `provision.js` at onboarding time with a JSON string containing the contractor's real resource IDs (Sheet ID, folder IDs, template IDs, company info). The script reads this on first run and stores everything in Google's Script Properties — after that the bootstrap data is never used again.

---

## npm Scripts

| Script | What it does |
|---|---|
| `npm run dev` | Start local dev server on port 3000 |
| `npm start` | Start server (used by Railway in production) |
| `npm run push` | Push gas/ to Apps Script via clasp (local testing only) |
