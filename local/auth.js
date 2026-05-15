// ============================================================
// local/auth.js — Google OAuth2 token management
// ============================================================
//
// Two separate concerns:
//
//   APP CREDENTIALS  (local/app-credentials.json)
//   ─────────────────────────────────────────────
//   Created ONCE by the developer. Contains the Google OAuth
//   Client ID + Secret for this app. Gitignored so they're
//   never committed. End users never see or touch this file.
//
//   USER SESSION  (local/.dev-config.json)
//   ──────────────────────────────────────
//   Created per user, per machine. Contains the contractor's
//   OAuth tokens, company name, logo, Drive folder/template IDs.
//   Also gitignored.
//
// ============================================================

const { google } = require('googleapis');
const fs         = require('fs');
const path       = require('path');

const APP_CREDS_PATH = path.join(__dirname, 'app-credentials.json');
const CONFIG_PATH    = path.join(__dirname, '.dev-config.json');
const REDIRECT       = process.env.REDIRECT_URI || 'http://localhost:3000/auth/callback';

// Requested once during onboarding to provision the contractor's
// private system. Token is discarded after provisioning completes.
const SCOPES = [
  'https://www.googleapis.com/auth/drive',               // create folders, docs, logo
  'https://www.googleapis.com/auth/spreadsheets',        // create + populate the Sheet
  'https://www.googleapis.com/auth/script.projects',     // create the GAS project
  'https://www.googleapis.com/auth/script.deployments',  // deploy as Web App
  'https://www.googleapis.com/auth/userinfo.email',
];

// ── App credentials (developer-side, one-time) ─────────────────

function loadAppCreds() {
  // In production, credentials come from environment variables
  if (process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET) {
    return { clientId: process.env.GOOGLE_CLIENT_ID, clientSecret: process.env.GOOGLE_CLIENT_SECRET };
  }
  // In local dev, read from gitignored file
  if (fs.existsSync(APP_CREDS_PATH)) {
    try { return JSON.parse(fs.readFileSync(APP_CREDS_PATH, 'utf8')); }
    catch { return {}; }
  }
  return {};
}

function hasAppCreds() {
  const { clientId, clientSecret } = loadAppCreds();
  return !!(clientId && clientSecret);
}

// ── User session config ────────────────────────────────────────

function loadConfig() {
  if (fs.existsSync(CONFIG_PATH)) {
    try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); }
    catch { return {}; }
  }
  return {};
}

function saveConfig(patch) {
  const existing = loadConfig();
  fs.writeFileSync(CONFIG_PATH, JSON.stringify({ ...existing, ...patch }, null, 2));
}

// ── OAuth2 client ──────────────────────────────────────────────

function getClient() {
  const { clientId, clientSecret } = loadAppCreds();
  return new google.auth.OAuth2(clientId, clientSecret, REDIRECT);
}

function getAuthenticatedClient() {
  const cfg    = loadConfig();
  const client = getClient();
  client.setCredentials(cfg.tokens);
  client.on('tokens', tokens => {
    const current = loadConfig();
    saveConfig({ tokens: { ...current.tokens, ...tokens } });
  });
  return client;
}

// ── Status checks ──────────────────────────────────────────────

function isConfigured() { return hasAppCreds(); }

function isAuthenticated() {
  const { tokens } = loadConfig();
  return !!(tokens && (tokens.access_token || tokens.refresh_token));
}

function hasFolderSetup() {
  const { folderId, folderUrl, proposalTemplateId, invoiceTemplateId } = loadConfig();
  return !!(folderId && folderUrl && proposalTemplateId && invoiceTemplateId);
}

function isReady() {
  return hasAppCreds() && isAuthenticated() && hasFolderSetup();
}

module.exports = {
  APP_CREDS_PATH,
  CONFIG_PATH,
  SCOPES,
  loadAppCreds,
  hasAppCreds,
  loadConfig,
  saveConfig,
  getClient,
  getAuthenticatedClient,
  isConfigured,
  isAuthenticated,
  hasFolderSetup,
  isReady,
};
