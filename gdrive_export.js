#!/usr/bin/env node
'use strict';

/**
 * Google Drive Export Utility
 *
 * Exports a Google Sheet to XLSX, Google Doc to PDF, or Google Slide to PPTX,
 * then uploads the result to a target Drive folder.
 *
 * Usage:
 *   node gdrive_export.js <source_url_or_id> <folder_url_or_id> [options]
 *   node gdrive_export.js --help
 *
 * First-time setup:
 *   1. Create a Google Cloud project and enable the Drive API.
 *   2. Create OAuth 2.0 credentials (Desktop app) and download the JSON file.
 *   3. Save it to ~/.config/gdrive_export/credentials.json
 *   4. Run the script once — a browser window will open for authorization.
 *      The token is cached at ~/.config/gdrive_export/token.json for subsequent runs.
 */

const fs = require('fs');
const path = require('path');
const os = require('os');
const http = require('http');
const { URL } = require('url');
const { Readable } = require('stream');
const { execSync } = require('child_process');
const { google } = require('googleapis');
const { Command } = require('commander');

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const CONFIG_DIR = path.join(os.homedir(), '.config', 'gdrive_export');
const CREDENTIALS_FILE = path.join(CONFIG_DIR, 'credentials.json');
const TOKEN_FILE = path.join(CONFIG_DIR, 'token.json');

const SCOPES = [
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/drive.file',
];

const EXPORT_MAP = {
  'application/vnd.google-apps.spreadsheet': {
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ext: 'xlsx',
  },
  'application/vnd.google-apps.document': {
    mime: 'application/pdf',
    ext: 'pdf',
  },
  'application/vnd.google-apps.presentation': {
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ext: 'pptx',
  },
};

const FRIENDLY_TYPE = {
  'application/vnd.google-apps.spreadsheet': 'Spreadsheet',
  'application/vnd.google-apps.document': 'Document',
  'application/vnd.google-apps.presentation': 'Presentation',
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function extractId(urlOrId) {
  const match = urlOrId.match(/\/(?:d|folders)\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : urlOrId;
}

function utcTimestamp() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  return (
    `${now.getUTCFullYear()}` +
    `${pad(now.getUTCMonth() + 1)}` +
    `${pad(now.getUTCDate())}` +
    `_${pad(now.getUTCHours())}` +
    `${pad(now.getUTCMinutes())}` +
    `${pad(now.getUTCSeconds())}z`
  );
}

function openBrowser(url) {
  const cmd =
    process.platform === 'darwin' ? 'open' :
    process.platform === 'win32'  ? 'start' :
    'xdg-open';
  try { execSync(`${cmd} "${url}"`); } catch { /* ignore — user can open manually */ }
}

// ---------------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------------

async function getAuth() {
  if (!fs.existsSync(CREDENTIALS_FILE)) {
    console.error(
      `OAuth credentials file not found at:\n  ${CREDENTIALS_FILE}\n\n` +
      'Steps to fix:\n' +
      '  1. Go to https://console.cloud.google.com/\n' +
      '  2. Create a project and enable the Google Drive API.\n' +
      '  3. Under APIs & Services > Credentials, create an OAuth 2.0 Client ID\n' +
      '     (Application type: Desktop app).\n' +
      '  4. Download the JSON and save it to the path above.\n' +
      '  5. Re-run this script.'
    );
    process.exit(1);
  }

  const keyFile = JSON.parse(fs.readFileSync(CREDENTIALS_FILE, 'utf8'));
  const { client_id, client_secret } = keyFile.installed ?? keyFile.web;

  // Use cached token if present and unexpired.
  if (fs.existsSync(TOKEN_FILE)) {
    const saved = JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8'));
    const client = new google.auth.OAuth2(client_id, client_secret);
    client.setCredentials(saved);

    if (!saved.expiry_date || Date.now() < saved.expiry_date - 60_000) {
      return client;
    }
    if (saved.refresh_token) {
      const { credentials } = await client.refreshAccessToken();
      fs.writeFileSync(TOKEN_FILE, JSON.stringify(credentials));
      client.setCredentials(credentials);
      return client;
    }
  }

  // Full browser-based OAuth flow with a local redirect server.
  return new Promise((resolve, reject) => {
    const server = http.createServer();
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address();
      const redirectUri = `http://127.0.0.1:${port}`;
      const client = new google.auth.OAuth2(client_id, client_secret, redirectUri);

      const authUrl = client.generateAuthUrl({ access_type: 'offline', scope: SCOPES });
      console.log('Opening browser for authorization…');
      console.log(`If the browser does not open, visit:\n  ${authUrl}\n`);
      openBrowser(authUrl);

      server.once('request', async (req, res) => {
        const code = new URL(req.url, redirectUri).searchParams.get('code');
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end('<p>Authorization complete — you may close this tab.</p>');
        server.close();

        if (!code) return reject(new Error('No authorization code received in redirect'));

        const { tokens } = await client.getToken(code);
        client.setCredentials(tokens);
        fs.mkdirSync(CONFIG_DIR, { recursive: true });
        fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens));
        console.log('Authorization successful. Token cached.');
        resolve(client);
      });
    });
  });
}

// ---------------------------------------------------------------------------
// Core logic
// ---------------------------------------------------------------------------

async function exportAndUpload(sourceId, folderId, { appendTimestamp, filename, dryRun }) {
  const auth = await getAuth();
  const drive = google.drive({ version: 'v3', auth });

  let meta;
  try {
    ({ data: meta } = await drive.files.get({
      fileId: sourceId,
      fields: 'name,mimeType',
      supportsAllDrives: true,
    }));
  } catch (err) {
    console.error(`Could not access source file (id=${sourceId}): ${err.message}`);
    process.exit(1);
  }

  const { name, mimeType } = meta;
  const mapping = EXPORT_MAP[mimeType];
  if (!mapping) {
    const supported = Object.values(FRIENDLY_TYPE).join(', ');
    console.error(`Unsupported file type: ${mimeType}\nSupported types: ${supported}`);
    process.exit(1);
  }

  const { mime: exportMime, ext } = mapping;
  const baseName = filename ?? name;
  const ts = appendTimestamp ? utcTimestamp() : null;
  const exportName = ts ? `${baseName}_${ts}.${ext}` : `${baseName}.${ext}`;

  console.log(`Source : ${FRIENDLY_TYPE[mimeType]} — '${name}' (id=${sourceId})`);
  console.log(`Export : ${exportName}  [${ext.toUpperCase()}]`);
  console.log(`Folder : ${folderId}`);

  if (dryRun) {
    console.log('[dry-run] Skipping export and upload.');
    return;
  }

  // Export file content.
  let exportData;
  try {
    const response = await drive.files.export(
      { fileId: sourceId, mimeType: exportMime },
      { responseType: 'arraybuffer' }
    );
    exportData = Buffer.from(response.data);
  } catch (err) {
    console.error(`Export failed: ${err.message}`);
    process.exit(1);
  }

  // Check for an existing file with the same name in the target folder.
  let existingId = null;
  try {
    const { data } = await drive.files.list({
      q: `name = ${JSON.stringify(exportName)} and '${folderId}' in parents and trashed = false`,
      fields: 'files(id)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    existingId = data.files[0]?.id ?? null;
  } catch (err) {
    console.error(`Failed to check for existing file: ${err.message}`);
    process.exit(1);
  }

  // Upload to target folder — overwrite if a file with the same name exists.
  let uploaded;
  try {
    if (existingId) {
      console.log(`Overwriting existing file (id=${existingId})…`);
      ({ data: uploaded } = await drive.files.update({
        fileId: existingId,
        requestBody: { name: exportName },
        media: { mimeType: exportMime, body: Readable.from(exportData) },
        fields: 'id,name,webViewLink',
        supportsAllDrives: true,
      }));
    } else {
      ({ data: uploaded } = await drive.files.create({
        requestBody: { name: exportName, parents: [folderId] },
        media: { mimeType: exportMime, body: Readable.from(exportData) },
        fields: 'id,name,webViewLink',
        supportsAllDrives: true,
      }));
    }
  } catch (err) {
    console.error(`Upload failed: ${err.message}`);
    process.exit(1);
  }

  console.log(`Uploaded: ${uploaded.webViewLink ?? uploaded.id}`);
}

// ---------------------------------------------------------------------------
// CLI entry point
// ---------------------------------------------------------------------------

const program = new Command();
program
  .name('gdrive_export')
  .description('Export a Google Sheet/Doc/Slide to XLSX/PDF/PPTX and upload it to a Drive folder.')
  .argument('<source>', 'Google Sheet/Doc/Slide URL or file ID')
  .argument('<folder>', 'Target Drive folder URL or folder ID')
  .option('--append-timestamp', 'Append a UTC timestamp suffix (YYYYMMDDz_HHMMSSz) to the filename')
  .option('--filename <name>', "Override the base filename instead of using the source file's Drive name")
  .option('--dry-run', 'Print what would happen without downloading or uploading anything')
  .addHelpText('after', `
Examples:
  node gdrive_export.js <FILE_URL> <FOLDER_URL>
  node gdrive_export.js <FILE_ID> <FOLDER_ID> --append-timestamp
  node gdrive_export.js <FILE_ID> <FOLDER_ID> --filename "Q1 Report"
  node gdrive_export.js <FILE_ID> <FOLDER_ID> --filename "Q1 Report" --append-timestamp
  node gdrive_export.js <FILE_ID> <FOLDER_ID> --dry-run`)
  .action(async (source, folder, opts) => {
    await exportAndUpload(extractId(source), extractId(folder), {
      appendTimestamp: opts.appendTimestamp ?? false,
      filename: opts.filename ?? null,
      dryRun: opts.dryRun ?? false,
    });
  });

program.parseAsync(process.argv).catch(err => {
  console.error(err.message);
  process.exit(1);
});
