/**
 * sheets-api.js — Direct Google Sheets & Drive API wrapper for SheetOps
 *
 * Uses OAuth2 tokens already stored by clasp at ~/.clasprc.json.
 * Does NOT require clasp run or Execution API or GCP project association.
 *
 * Required OAuth scopes (add via: node sheetops.js auth):
 *   https://www.googleapis.com/auth/spreadsheets
 *   https://www.googleapis.com/auth/drive
 */

import type { GridRange, Color, SpreadsheetMetadata, SheetMetadata, NamedRange, SheetSnapshot, WorkbookSnapshot, VisualOptions, Dynamic, JsonObject, CellValue, RangeTarget, WriteOptions, BackupOptions } from "../types";


import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";

// Resolve googleapis from project-local or global node_modules
let google: typeof import("googleapis").google;
try {
  google = require("googleapis").google;
} catch(_) {
  try {
    google = require(path.join(__dirname, "../..", "node_modules", "googleapis")).google;
  } catch(_2) {
    google = require("/opt/homebrew/lib/node_modules/googleapis").google;
  }
}

const CLASPRC        = path.join(os.homedir(), ".clasprc.json");
const SHEETOPS_CREDS = path.join(os.homedir(), ".sheetops-creds.json");

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/script.projects",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/script.scriptapp",
  "https://www.googleapis.com/auth/userinfo.email"
];

// ─── Credential store (multi-account v2) ─────────────────────────────────────
//
//  ~/.sheetops-creds.json format (v2):
//  {
//    "version": 2,
//    "client_id": "...",
//    "client_secret": "...",
//    "active": "user@example.com",
//    "accounts": {
//      "user@example.com": { access_token, refresh_token, scope, token_type, expiry_date }
//    }
//  }
//
//  v1 (flat) files are still readable as single-account fallback.

function _loadCreds() {
  if (!fs.existsSync(SHEETOPS_CREDS)) return null;
  return JSON.parse(fs.readFileSync(SHEETOPS_CREDS, "utf8"));
}

function _saveCreds(creds: Dynamic) {
  fs.writeFileSync(SHEETOPS_CREDS, JSON.stringify(creds, null, 2) + "\n");
  try { fs.chmodSync(SHEETOPS_CREDS, 0o600); } catch(_) {}
}

/** Normalise any creds file (v1 or v2) to v2 structure. */
function _normalise(raw: Dynamic) {
  if (!raw) return null;
  if (raw.version === 2) return raw;
  // v1 flat → v2
  const acctKey = "default";
  return {
    version:       2,
    client_id:     raw.client_id,
    client_secret: raw.client_secret,
    active:        acctKey,
    accounts: {
      [acctKey]: {
        access_token:  raw.access_token,
        refresh_token: raw.refresh_token,
        scope:         raw.scope,
        token_type:    raw.token_type,
        expiry_date:   raw.expiry_date,
        redirect_uri:  raw.redirect_uri
      }
    }
  };
}

/** Return the token bundle for a given account key (or active). */
function _getTokens(account?: string) {
  const creds = _normalise(_loadCreds());
  if (!creds) throw new Error("Not authenticated. Run: node sheetops.js auth");
  const key = account || creds.active;
  const tokens = creds.accounts && creds.accounts[key];
  if (!tokens || (!tokens.access_token && !tokens.refresh_token)) {
    throw new Error(`Account '${key}' not authenticated. Run: node sheetops.js auth`);
  }
  return { creds, key, tokens };
}

// ─── Auth ────────────────────────────────────────────────────────────────────

/**
 * Build an OAuth2 client for a specific account (or the active account).
 * Falls back to ~/.clasprc.json if no SheetOps creds exist.
 */
function getOAuth2Client(account?: string): import("google-auth-library").OAuth2Client {
  // 1 — SheetOps own credentials
  if (fs.existsSync(SHEETOPS_CREDS)) {
    try {
      const { creds, key, tokens } = _getTokens(account);
      const oauth2 = new google.auth.OAuth2(
        creds.client_id,
        creds.client_secret,
        tokens.redirect_uri || "http://localhost:3000/callback"
      );
      oauth2.setCredentials({
        access_token:  tokens.access_token,
        refresh_token: tokens.refresh_token,
        token_type:    tokens.token_type || "Bearer",
        expiry_date:   tokens.expiry_date,
        scope:         tokens.scope
      });
      // Auto-save refreshed tokens back
      oauth2.on("tokens", (newTokens: import("google-auth-library").Credentials) => {
        try {
          const current = _normalise(_loadCreds());
          if (current && current.accounts && current.accounts[key]) {
            Object.assign(current.accounts[key], newTokens);
            _saveCreds(current);
          }
        } catch(_) {}
      });
      return oauth2;
    } catch(e) {
      if (account) throw e; // explicit account not found — propagate
      // fall through to clasp fallback for default account
    }
  }

  // 2 — Clasp fallback
  if (!fs.existsSync(CLASPRC)) {
    throw new Error("Not authenticated. Run: node sheetops.js auth");
  }
  const rc     = JSON.parse(fs.readFileSync(CLASPRC, "utf8"));
  const tokens = rc.tokens && rc.tokens.default ? rc.tokens.default : rc;
  if (!tokens.access_token && !tokens.refresh_token) {
    throw new Error("No valid tokens found. Run: node sheetops.js auth");
  }
  if (!tokens.client_id || !tokens.client_secret) {
    throw new Error("Clasp credentials missing client_id/client_secret. Run: node sheetops.js auth");
  }
  const oauth2 = new google.auth.OAuth2(
    tokens.client_id,
    tokens.client_secret
  );
  oauth2.setCredentials({
    access_token:  tokens.access_token,
    refresh_token: tokens.refresh_token,
    token_type:    tokens.token_type || "Bearer",
    expiry_date:   tokens.expiry_date
  });
  oauth2.on("tokens", (newTokens: import("google-auth-library").Credentials) => {
    try {
      const current = JSON.parse(fs.readFileSync(CLASPRC, "utf8"));
      const target  = current.tokens && current.tokens.default ? current.tokens.default : current;
      Object.assign(target, newTokens);
      fs.writeFileSync(CLASPRC, JSON.stringify(current, null, 2) + "\n");
    } catch(_) {}
  });
  return oauth2;
}

/** Returns true if a valid Sheets-scoped credential is available for the given account. */
function hasSpreadsheetScope(account?: string) {
  if (fs.existsSync(SHEETOPS_CREDS)) {
    try {
      const { tokens } = _getTokens(account);
      return !!(tokens.access_token || tokens.refresh_token);
    } catch(_) {}
  }
  try {
    const rc     = JSON.parse(fs.readFileSync(CLASPRC, "utf8"));
    const tokens = rc.tokens && rc.tokens.default ? rc.tokens.default : rc;
    return (tokens.scope || "").includes("spreadsheets");
  } catch(_) { return false; }
}

/** List all authenticated accounts. Returns [{ key, active, scope }] */
function listAccounts() {
  const creds = _normalise(_loadCreds());
  if (!creds || !creds.accounts) return [];
  return Object.entries(creds.accounts as Record<string, {scope?: string}>).map(([key, t]) => ({
    key,
    active: key === creds.active,
    scope:  t.scope || ""
  }));
}

/** Switch the active account. */
function switchAccount(email: string) {
  const creds = _normalise(_loadCreds());
  if (!creds) throw new Error("No credentials found.");
  if (!creds.accounts[email]) throw new Error(`Account '${email}' not found. Run: node sheetops.js auth`);
  creds.active = email;
  _saveCreds(creds);
}

/** Remove an account from the store. */
function removeAccount(email: string) {
  const creds = _normalise(_loadCreds());
  if (!creds) throw new Error("No credentials found.");
  if (!creds.accounts[email]) throw new Error(`Account '${email}' not found.`);
  delete creds.accounts[email];
  if (creds.active === email) {
    const remaining = Object.keys(creds.accounts);
    creds.active = remaining.length ? remaining[0] : null;
  }
  _saveCreds(creds);
  return creds.active;
}

/**
 * Migrate any "default" account key to the real email address.
 * Call this on init — safe to call any time, no-ops if already resolved.
 * Returns the resolved email if migration happened, null otherwise.
 */
async function migrateDefaultAccount() {
  const creds = _normalise(_loadCreds());
  if (!creds || !creds.accounts || !creds.accounts["default"]) return null;
  try {
    const oauth2 = getOAuth2Client("default");
    const email  = await getUserEmail(oauth2);
    if (!email || email === "default") return null;
    // Rename the key
    creds.accounts[email] = creds.accounts["default"];
    delete creds.accounts["default"];
    if (creds.active === "default") creds.active = email;
    _saveCreds(creds);
    return email;
  } catch(_) { return null; }
}

/** Get the authenticated Google account email via userinfo API. */
async function getUserEmail(oauth2: import("google-auth-library").OAuth2Client) {
  try {
    const oauth2Api = google.oauth2({ version: "v2", auth: oauth2 });
    const res = await oauth2Api.userinfo.get();
    return res.data.email || null;
  } catch(_) { return null; }
}

/**
 * runAuthFlow(clientId, clientSecret)
 *
 * Opens a browser OAuth consent flow, captures the token, detects the account
 * email via userinfo, and saves to ~/.sheetops-creds.json in v2 multi-account
 * format. Safe to run multiple times — adds/updates the account without
 * touching other accounts.
 */
async function runAuthFlow(clientId: string, clientSecret: string) {
  const http      = require("http");
  const net       = require("net");
  const { exec }  = require("child_process");

  async function findFreePort(start: number) {
    for (let p = start; p < start + 20; p++) {
      const free = await new Promise((res) => {
        const s = net.createServer();
        s.once("error", () => res(false));
        s.once("listening", () => { s.close(); res(true); });
        s.listen(p, "localhost");
      });
      if (free) return p;
    }
    throw new Error("No free port found in range 3000-3019");
  }

  const PORT         = await findFreePort(3000);
  const REDIRECT_URI = `http://localhost:${PORT}/callback`;
  const oauth2       = new google.auth.OAuth2(clientId, clientSecret, REDIRECT_URI);

  const authUrl = oauth2.generateAuthUrl({
    access_type: "offline",
    scope:       SCOPES,
    prompt:      "consent"
  });

  exec(`open "${authUrl}"`);
  process.stdout.write(`\nOpening browser for Google sign-in...\nURL (if browser doesn't open):\n${authUrl}\n\nWaiting for callback on http://localhost:${PORT}/callback ...\n`);

  const code = await new Promise<string>((resolve, reject) => {
    const server = http.createServer((req: import("node:http").IncomingMessage, res: import("node:http").ServerResponse) => {
      try {
        const url  = new URL(req.url || "/", `http://localhost:${PORT}`);
        const code = url.searchParams.get("code");
        const e    = url.searchParams.get("error");
        if (code) {
          res.writeHead(200, { "Content-Type": "text/html" });
          res.end("<html><body style='font-family:sans-serif;padding:2em'><h2>✅ SheetOps: authorisation complete.</h2><p>You can close this tab and return to the terminal.</p></body></html>");
          server.close(); resolve(code);
        } else {
          res.writeHead(400, { "Content-Type": "text/html" });
          res.end(`<html><body><h2>Error: ${e || "unknown"}</h2></body></html>`);
          server.close(); reject(new Error(`OAuth error: ${e || "no code"}`));
        }
      } catch(err) { reject(err); }
    });
    server.on("error", reject);
    server.listen(PORT, "localhost");
    setTimeout(() => { server.close(); reject(new Error("Timed out (2 min). Please try again.")); }, 120000);
  });

  const { tokens } = await oauth2.getToken(code);
  oauth2.setCredentials(tokens);

  // Detect which Google account was authenticated
  const email = await getUserEmail(oauth2);
  const accountKey = email || "default";

  // Load existing creds (or create new v2 structure)
  let existing = _normalise(_loadCreds());
  if (!existing) {
    existing = { version: 2, client_id: clientId, client_secret: clientSecret, active: accountKey, accounts: {} };
  } else {
    // Always update client credentials in case they changed
    existing.client_id     = clientId;
    existing.client_secret = clientSecret;
  }
  existing.accounts[accountKey] = {
    access_token:  tokens.access_token,
    refresh_token: tokens.refresh_token,
    scope:         tokens.scope,
    token_type:    tokens.token_type || "Bearer",
    expiry_date:   tokens.expiry_date,
    redirect_uri:  REDIRECT_URI
  };
  existing.active = accountKey;

  _saveCreds(existing);
  return { email: accountKey, scope: tokens.scope };
}

// ─── Sheets helpers ──────────────────────────────────────────────────────────

function sheets(auth: import("google-auth-library").OAuth2Client) {
  return google.sheets({ version: "v4", auth });
}

function drive(auth: import("google-auth-library").OAuth2Client) {
  return google.drive({ version: "v3", auth });
}

/** SHA-256 hex hash of a JSON-serialisable value */
import * as crypto from "node:crypto";
function hash(val: unknown) {
  return crypto.createHash("sha256").update(JSON.stringify(val)).digest("hex");
}

/** Normalise a Sheets API 2-D values array (fill missing rows/cols with "") */
function normalise(values: CellValue[][] | null | undefined, numRows: number, numCols: number) {
  const result: CellValue[][] = [];
  for (let r = 0; r < numRows; r++) {
    const row = values && values[r] ? values[r] : [];
    const filled: CellValue[] = [];
    for (let c = 0; c < numCols; c++) {
      filled.push(row[c] !== undefined ? row[c] : "");
    }
    result.push(filled);
  }
  return result;
}

// ─── Public API ──────────────────────────────────────────────────────────────

/**
 * getSpreadsheetMeta(spreadsheetId)
 * Returns full spreadsheet metadata (sheets, namedRanges, properties).
 */
async function getSpreadsheetMeta(spreadsheetId: string): Promise<SpreadsheetMetadata> {
  const auth  = getOAuth2Client();
  const api   = sheets(auth);
  const res   = await api.spreadsheets.get({
    spreadsheetId,
    includeGridData: false,
    fields: "spreadsheetId,properties,sheets(properties,protectedRanges),namedRanges,developerMetadata"
  });
  return res.data as SpreadsheetMetadata;
}

/**
 * healthcheck(spreadsheetId)
 * Returns a health summary: name, sheet list, named ranges.
 */
async function healthcheck(spreadsheetId: string) {
  const meta = await getSpreadsheetMeta(spreadsheetId);
  const sheetList = (meta.sheets || []).map((s: SheetMetadata) => ({
    name:     s.properties.title,
    sheetId:  s.properties.sheetId,
    index:    s.properties.index,
    hidden:   s.properties.hidden || false,
    rowCount: s.properties.gridProperties ? s.properties.gridProperties.rowCount : 0,
    colCount: s.properties.gridProperties ? s.properties.gridProperties.columnCount : 0
  }));
  const namedRanges = (meta.namedRanges || []).map((nr: NamedRange) => ({
    name:    nr.name,
    rangeId: nr.namedRangeId,
    range:   nr.range
  }));
  return {
    spreadsheetId:   meta.spreadsheetId,
    spreadsheetName: meta.properties ? meta.properties.title : "(unknown)",
    sheetCount:      sheetList.length,
    sheets:          sheetList,
    namedRanges:     namedRanges
  };
}

/**
 * snapshotWorkbook(spreadsheetId)
 * Full snapshot: all sheet metadata + header rows + named ranges + protected ranges.
 */
async function snapshotWorkbook(spreadsheetId: string) {
  const auth  = getOAuth2Client();
  const api   = sheets(auth);

  // Full metadata
  const metaRes = await api.spreadsheets.get({
    spreadsheetId,
    includeGridData: false
  });
  const meta = metaRes.data as SpreadsheetMetadata;

  // Fetch header row (row 1) for every non-hidden sheet
  const sheetSnapshots: SheetSnapshot[] = [];
  for (const sh of (meta.sheets || [])) {
    const props   = sh.properties;
    const isHidden = props.hidden || false;
    let sampleHeader: CellValue[] = [];
    let lastRow = 0;
    let lastCol = 0;

    try {
      // Get the used range info via a small batchGet
      const valRes = await api.spreadsheets.values.get({
        spreadsheetId,
        range: `'${props.title}'!A1:Z1`,
        valueRenderOption: "FORMATTED_VALUE"
      });
      sampleHeader = (valRes.data.values && valRes.data.values[0]) ? valRes.data.values[0] : [];

      // Approximate last row/col from sheet grid properties
      lastRow = props.gridProperties ? props.gridProperties.rowCount : 0;
      lastCol = props.gridProperties ? props.gridProperties.columnCount : 0;
    } catch(_) {
      // Sheet may have no data or be inaccessible
    }

    sheetSnapshots.push({
      name:     props.title,
      sheetId:  props.sheetId,
      index:    props.index,
      hidden:   isHidden,
      lastRow,
      lastColumn: lastCol,
      frozenRows: props.gridProperties ? (props.gridProperties.frozenRowCount || 0) : 0,
      frozenCols: props.gridProperties ? (props.gridProperties.frozenColumnCount || 0) : 0,
      sampleHeader,
      tabColor: props.tabColorStyle ? props.tabColorStyle.rgbColor : null
    });
  }

  const namedRanges = (meta.namedRanges || []).map((nr: NamedRange) => ({
    name:    nr.name,
    rangeId: nr.namedRangeId,
    range:   nr.range
  }));

  const protectedRanges: Dynamic[] = [];
  for (const sh of (meta.sheets || [])) {
    for (const pr of (sh.protectedRanges || [])) {
      protectedRanges.push({
        description: pr.description || "",
        sheetId:     pr.range ? pr.range.sheetId : null,
        sheetName:   sh.properties.title,
        range:       pr.range,
        editors:     pr.editors ? (pr.editors.users || []) : []
      });
    }
  }

  const snapshotHash = hash({ sheets: sheetSnapshots.map((s: SheetSnapshot) => ({ name: s.name, lastRow: s.lastRow })) });

  return {
    spreadsheetId,
    spreadsheetName: meta.properties ? meta.properties.title : "(unknown)",
    snapshotHash,
    sheets:          sheetSnapshots,
    namedRanges,
    protectedRanges,
    // Note: triggers are only visible via Apps Script API — not available through Sheets API
    triggers:        []
  };
}

/**
 * readRange(spreadsheetId, target)
 * target: { namedRange, sheetName, a1 }
 * Returns { sheet, a1, rows, cols, values, hash }
 */
async function readRange(spreadsheetId: string, target: RangeTarget) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);

  let rangeA1: string;
  if (target.namedRange) {
    rangeA1 = target.namedRange; // Sheets API accepts named ranges directly
  } else if (target.sheetName) {
    rangeA1 = `'${target.sheetName.replace(/'/g, "''")}'${target.a1 ? "!" + target.a1 : ""}`;
  } else {
    throw new Error("target must have namedRange, or sheetName + a1");
  }

  const res = await api.spreadsheets.values.get({
    spreadsheetId,
    range: rangeA1,
    valueRenderOption: "FORMATTED_VALUE",
    dateTimeRenderOption: "FORMATTED_STRING"
  });

  const values: CellValue[][] = res.data.values || [[]];
  const rows   = values.length;
  const cols   = values.reduce((m: number, r: CellValue[]) => Math.max(m, r.length), 0);
  const normalised = normalise(values, rows, cols);

  return {
    sheet:  res.data.range ? res.data.range.split("!")[0].replace(/'/g, "") : target.sheetName || "",
    a1:     res.data.range || rangeA1,
    rows,
    cols,
    values: normalised,
    hash:   hash(normalised)
  };
}

/**
 * writeRangeDryRun(spreadsheetId, target, newValues, opts)
 * Returns a preview without writing.
 */
async function writeRangeDryRun(spreadsheetId: string, target: RangeTarget, newValues: CellValue[][], opts: WriteOptions = {}) {
  const current = await readRange(spreadsheetId, target);
  const warnings: string[] = [];

  if (opts.expectedHash && opts.expectedHash !== current.hash) {
    return {
      dryRun: true, ok: false,
      error: "expectedHash mismatch — sheet has changed since last read",
      expected: opts.expectedHash, actual: current.hash
    };
  }

  const totalCells = (newValues || []).reduce((s: number, r: CellValue[]) => s + (r || []).length, 0);
  if (totalCells > 100) warnings.push(`LARGE_WRITE: ${totalCells} cells (>100). Set confirmLarge:true.`);

  return {
    dryRun:            true,
    sheet:             current.sheet,
    a1:                current.a1,
    rows:              current.rows,
    cols:              current.cols,
    currentHash:       current.hash,
    currentValues:     current.values,
    proposedValues:    newValues || [],
    totalCellsAffected: totalCells,
    warnings,
    requiresApproval:  warnings.length > 0 || totalCells > 100
  };
}

async function appendRowsDryRun(spreadsheetId: string, target: RangeTarget, rows: CellValue[][]) {
  const preview = await writeRangeDryRun(spreadsheetId, target, rows);
  if ("error" in preview) return preview;
  const appendRange = `'${target.sheetName!.replace(/'/g, "''")}'${target.a1 ? "!" + target.a1 : ""}`;
  return { ...preview, appendRange, warnings: [...preview.warnings, `APPEND: Google selects the next row of the table in ${appendRange}; the exact row is known only after apply.`], requiresApproval: true };
}

async function clearRangeDryRun(spreadsheetId: string, target: RangeTarget) {
  const current = await readRange(spreadsheetId, target);
  const a1 = String(current.a1).split("!").pop()!;
  const match = a1.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  const cells = match ? (_colToIndex(match[3] || match[1]) - _colToIndex(match[1]) + 1) * (Number(match[4] || match[2]) - Number(match[2]) + 1) : null;
  return { dryRun: true, sheet: current.sheet, a1: current.a1, currentHash: current.hash, totalCellsAffected: cells,
    warnings: [cells === null ? "DESTRUCTIVE_CLEAR: the entire target range will be cleared; its extent is unknown." : `DESTRUCTIVE_CLEAR: ${cells} target cells will be cleared.`], requiresApproval: true };
}

/**
 * writeRange(spreadsheetId, target, newValues, opts)
 * Writes values. Guards: expectedHash, confirmLarge.
 */
async function writeRange(spreadsheetId: string, target: RangeTarget, newValues: CellValue[][], opts: WriteOptions = {}) {
  // Hash check
  if (opts.expectedHash) {
    const current = await readRange(spreadsheetId, target);
    if (opts.expectedHash !== current.hash) {
      throw new Error(`expectedHash mismatch — aborting write. Expected: ${opts.expectedHash} Got: ${current.hash}`);
    }
  }

  const totalCells = (newValues || []).reduce((s: number, r: CellValue[]) => s + (r || []).length, 0);
  if (totalCells > 100 && !opts.confirmLarge) {
    throw new Error(`Write affects ${totalCells} cells (>100). Set confirmLarge:true.`);
  }

  const auth = getOAuth2Client();
  const api  = sheets(auth);

  let rangeA1: string;
  if (target.namedRange) {
    rangeA1 = target.namedRange;
  } else {
    rangeA1 = `'${target.sheetName!.replace(/'/g, "''")}'!${target.a1}`;
  }

  await api.spreadsheets.values.update({
    spreadsheetId,
    range: rangeA1,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: newValues }
  });

  const after = await readRange(spreadsheetId, target);
  return { sheet: after.sheet, a1: after.a1, cellsWritten: totalCells, newHash: after.hash };
}

/**
 * appendRows(spreadsheetId, sheetName, rows)
 */
async function appendRows(spreadsheetId: string, sheetName: string, rows: CellValue[][], a1?: string) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);
  const res  = await api.spreadsheets.values.append({
    spreadsheetId,
    range:           `'${sheetName.replace(/'/g, "''")}'${a1 ? "!" + a1 : ""}`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody:     { values: rows }
  });
  const updates = res.data.updates || {};
  return {
    sheet:          sheetName,
    rowsAppended:   rows.length,
    updatedRange:   updates.updatedRange || "",
    updatedRows:    updates.updatedRows || 0
  };
}

/**
 * clearRange(spreadsheetId, target)
 * Clears cell values only (not formatting).
 */
async function clearRange(spreadsheetId: string, target: RangeTarget) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);
  let rangeA1 = target.namedRange || `'${target.sheetName!.replace(/'/g, "''")}'!${target.a1}`;
  await api.spreadsheets.values.clear({ spreadsheetId, range: rangeA1, requestBody: {} });
  return { sheet: target.sheetName || target.namedRange, a1: rangeA1, cleared: true };
}

/**
 * backupSpreadsheet(spreadsheetId, spreadsheetName, reason, opts)
 * Copies the spreadsheet in Drive.
 *
 * opts.folderId — place the copy in this Drive folder (else it inherits the
 *                 source's parent, the legacy behaviour).
 * opts.retain   — rolling retention: after copying, keep only the newest
 *                 `retain` "<name>__backup__*" files in that folder and trash
 *                 the rest (reversible — Drive trash, ~30 days). Requires
 *                 folderId; ignored otherwise.
 */
async function backupSpreadsheet(spreadsheetId: string, spreadsheetName: string, reason?: string, opts: BackupOptions = {}) {
  const auth = getOAuth2Client();
  const d    = drive(auth);
  const ts   = new Date().toISOString().replace(/[:.]/g, "-");
  const name = `${spreadsheetName}__backup__${ts}`;
  const requestBody: JsonObject = { name };
  if (opts.folderId) requestBody.parents = [opts.folderId];
  const res  = await d.files.copy({
    fileId:      spreadsheetId,
    requestBody,
    fields:      "id,name,parents"
  });
  if (!res.data.id) throw new Error("Drive copy returned no backup ID");
  const out: JsonObject = { backupName: name, backupId: res.data.id, reason: reason || "" };
  if (opts.folderId && Number(opts.retain) > 0) {
    try { out.pruned = await pruneBackups(d, opts.folderId, spreadsheetName, Number(opts.retain), res.data.id); }
    catch(e) { out.retentionWarning = "Backup created, but retention cleanup failed: " + e.message; }
  }
  return out;
}

/**
 * pruneBackups(d, folderId, spreadsheetName, retain)
 * Rolling retention: trash all but the newest `retain` "<name>__backup__*"
 * files in `folderId`, preserving the just-created copy by ID. Only exact
 * generated timestamp names qualify; renamed pinned anchors are excluded. Trashing is
 * reversible (Drive trash). Returns the list of trashed names.
 */
async function pruneBackups(d: import("googleapis").drive_v3.Drive, folderId: string, spreadsheetName: string, retain: number, keepId: string) {
  const prefix = `${spreadsheetName}__backup__`;
  const escapeQuery = (value: string) => value.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
  const files: Array<{id: string; name: string}> = [];
  let pageToken: string | undefined;
  do {
    const r = await d.files.list({
      q:        `'${escapeQuery(folderId)}' in parents and trashed = false and name contains '${escapeQuery(prefix)}'`,
      fields:   "nextPageToken, files(id,name)",
      pageSize: 1000,
      pageToken
    });
    files.push(...(r.data.files || []).flatMap(f => typeof f.id === "string" && typeof f.name === "string" ? [{id: f.id, name: f.name}] : []));
    pageToken = r.data.nextPageToken || undefined;
  } while (pageToken);

  // Drive contains is token-fuzzy. Keep only generated names; reserve one retention slot for the new copy.
  const rolling = files
    .filter((f) => f.name.startsWith(prefix) && /^\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}(?:-\d{3}Z)?$/.test(f.name.slice(prefix.length)))
    .filter((f) => f.id !== keepId)
    .sort((a, b) => (a.name < b.name ? 1 : a.name > b.name ? -1 : a.id.localeCompare(b.id)));

  const trashed: string[] = [];
  for (const f of rolling.slice(Math.max(0, Math.floor(retain) - 1))) {
    await d.files.update({ fileId: f.id, requestBody: { trashed: true } });
    trashed.push(f.name);
  }
  return trashed;
}

/**
 * logOperation(spreadsheetId, entry)
 * Appends a row to __AGENT_OPS_LOG (creates the sheet if missing).
 */
async function logOperation(spreadsheetId: string, entry: {operationId?: string; type?: string; target?: unknown; rowCount?: number; colCount?: number; status?: string; message?: string}) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);
  const LOG_SHEET = "__AGENT_OPS_LOG";

  // Ensure the log sheet exists
  const meta = await getSpreadsheetMeta(spreadsheetId);
  const exists = (meta.sheets || []).some((s) => s.properties.title === LOG_SHEET);

  if (!exists) {
    // Create the log sheet and add header
    await api.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          addSheet: {
            properties: { title: LOG_SHEET, hidden: true }
          }
        }]
      }
    });
    await api.spreadsheets.values.append({
      spreadsheetId,
      range:           LOG_SHEET,
      valueInputOption: "RAW",
      requestBody: {
        values: [["timestamp","operationId","type","target","rowCount","colCount","status","message","actor"]]
      }
    });
  }

  const row = [
    new Date().toISOString(),
    entry.operationId   || "",
    entry.type          || "",
    JSON.stringify(entry.target || {}),
    entry.rowCount      || 0,
    entry.colCount      || 0,
    entry.status        || "ok",
    entry.message       || "",
    "SheetOps"
  ];

  await api.spreadsheets.values.append({
    spreadsheetId,
    range:            LOG_SHEET,
    valueInputOption: "RAW",
    requestBody:      { values: [row] }
  });

  return { logged: true };
}

/**
 * runValidation(spreadsheetId, checks)
 * Runs validation checks against the live spreadsheet.
 */
async function runValidation(spreadsheetId: string, checks: Dynamic) {
  const meta = await getSpreadsheetMeta(spreadsheetId);
  const sheetNames = (meta.sheets || []).map((s) => s.properties.title);
  const namedRanges = (meta.namedRanges || []).map((nr) => nr.name);

  const results: Dynamic[] = [];
  for (const check of (checks || [])) {
    let pass = false, note = "";
    try {
      switch (check.type) {
        case "sheetExists":
          pass = sheetNames.includes(check.name);
          note = pass ? "found" : "not found";
          break;
        case "namedRangeExists":
          pass = namedRanges.includes(check.name);
          note = pass ? "found" : "not found";
          break;
        case "logSheetExists":
          pass = sheetNames.includes("__AGENT_OPS_LOG");
          note = pass ? "found" : "not found";
          break;
        case "rowCountMin":
        case "rowCountMax": {
          const auth = getOAuth2Client();
          const api  = sheets(auth);
          const res  = await api.spreadsheets.values.get({
            spreadsheetId,
            range: `'${check.sheetName}'!A:A`,
            valueRenderOption: "FORMATTED_VALUE"
          });
          const rows = (res.data.values || []).length;
          pass = check.type === "rowCountMin" ? rows >= check.min : rows <= check.max;
          note = `rows=${rows}`;
          break;
        }
        case "cellNotEmpty":
        case "cellEquals": {
          const val = await readRange(spreadsheetId, check.target);
          const cell = val.values[0] ? val.values[0][0] : "";
          pass = check.type === "cellNotEmpty"
            ? (cell !== "" && cell !== null && cell !== undefined)
            : String(cell) === String(check.expected);
          note = `value=${cell}` + (check.type === "cellEquals" ? ` expected=${check.expected}` : "");
          break;
        }
        default:
          pass = false; note = `unknown check type: ${check.type}`;
      }
    } catch(e) {
      pass = false; note = `error: ${e.message}`;
    }
    results.push({ check, pass, note });
  }

  return { allPassed: results.every((r) => r.pass), results };
}

/**
 * exportRange(spreadsheetId, target, format)
 * Exports a range as { csv: "...", json: [[...]], headers: [...], rows: N }
 * format: "csv" | "json" | "both" (default "both")
 */
async function exportRange(spreadsheetId: string, target: RangeTarget, format?: string) {
  // Allow sheetName-only target (whole sheet) — use named range style without A1
  if (target.sheetName && !target.a1 && !target.namedRange) {
    target = { ...target, a1: "A1:ZZZ" }; // Sheets API clips to actual data
  }
  const data = await readRange(spreadsheetId, target);
  const rows = data.values;
  const headers = rows.length ? rows[0] : [];
  const dataRows = rows.slice(1);

  let csv = "";
  if (!format || format === "csv" || format === "both") {
    csv = rows.map((r) => r.map((c) => {
      const s = String(c === null || c === undefined ? "" : c);
      return s.includes(",") || s.includes('"') || s.includes("\n") ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(",")).join("\n");
  }

  const jsonRows = dataRows.map((r) => {
    const obj: JsonObject = {};
    headers.forEach((h, i) => { obj[String(h || `col${i}`)] = r[i] !== undefined ? r[i] : ""; });
    return obj;
  });

  return {
    sheet:   data.sheet,
    a1:      data.a1,
    rows:    rows.length,
    cols:    data.cols,
    headers,
    csv:     csv || null,
    json:    (!format || format === "json" || format === "both") ? jsonRows : null,
    hash:    data.hash
  };
}

/**
 * findInSheet(spreadsheetId, sheetName, query, a1Range)
 * Searches for cells containing query (case-insensitive substring).
 * Returns [{ row, col, a1, value }]
 */
async function findInSheet(spreadsheetId: string, sheetName: string, query: string, a1Range?: string) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);

  const range = a1Range ? `'${sheetName}'!${a1Range}` : `'${sheetName}'`;
  const res   = await api.spreadsheets.values.get({
    spreadsheetId,
    range,
    valueRenderOption: "FORMATTED_VALUE"
  });

  const values: CellValue[][] = res.data.values || [];
  const q      = query.toLowerCase();
  const matches: Array<{row: number; col: number; a1: string; value: CellValue}> = [];

  // Determine start row/col from range if specified
  let startRow = 1;
  let startCol = 1;
  if (a1Range) {
    const colMatch = a1Range.match(/^([A-Z]+)/i);
    const rowMatch = a1Range.match(/\d+/);
    if (colMatch) startCol = colMatch[1].toUpperCase().split("").reduce((n: number, c: string) => n * 26 + c.charCodeAt(0) - 64, 0);
    if (rowMatch) startRow = parseInt(rowMatch[0]);
  }

  values.forEach((row: CellValue[], ri: number) => {
    row.forEach((cell: CellValue, ci: number) => {
      if (String(cell).toLowerCase().includes(q)) {
        const colLetter = (startCol + ci - 1);
        const colA1 = (() => {
          let n = colLetter, s = "";
          while (n > 0) { s = String.fromCharCode(((n - 1) % 26) + 65) + s; n = Math.floor((n - 1) / 26); }
          return s;
        })();
        matches.push({
          row:   startRow + ri,
          col:   startCol + ci,
          a1:    `${colA1}${startRow + ri}`,
          value: String(cell)
        });
      }
    });
  });

  return { sheet: sheetName, query, matchCount: matches.length, matches };
}

// ─── Colour / A1 helpers ─────────────────────────────────────────────────────

/** Convert "#RRGGBB" to Sheets API color object { red, green, blue } */
function _hexToColor(hex: string) {
  const h = (hex || "").replace("#", "").padEnd(6, "0");
  return {
    red:   parseInt(h.substring(0, 2), 16) / 255,
    green: parseInt(h.substring(2, 4), 16) / 255,
    blue:  parseInt(h.substring(4, 6), 16) / 255
  };
}

/** Convert a colour value (hex string or {red,green,blue} object) */
function _toColor(c: string | Color) {
  return typeof c === "string" ? _hexToColor(c) : c;
}

/** Column letter(s) → 0-based index.  "A" → 0, "B" → 1, "Z" → 25, "AA" → 26 */
function _colToIndex(col: string) {
  return col.toUpperCase().split("").reduce((n: number, c: string) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
}

/**
 * Parse an A1 range string → Sheets API GridRange (0-based, endRow/Col exclusive).
 * Supports "A1:D10", "A:D", "A1:D", "B2".
 * Omits endRowIndex if open-ended (whole-column range).
 */
function _parseA1ToGridRange(a1: string, sheetId: number) {
  const match = a1.match(/^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/i);
  if (!match) return { sheetId };
  const startCol = _colToIndex(match[1]);
  const startRow = match[2] ? parseInt(match[2]) - 1 : 0;
  const gr: GridRange = { sheetId, startColumnIndex: startCol, startRowIndex: startRow };
  if (match[3]) gr.endColumnIndex = _colToIndex(match[3]) + 1;
  if (match[4]) gr.endRowIndex    = parseInt(match[4]);
  return gr;
}

// ─── Sheet creation ───────────────────────────────────────────────────────────

/**
 * createSpreadsheet(schema, folderId?)
 *
 * Creates a fully-formatted Google Spreadsheet from a schema object.
 *
 * schema = {
 *   title:    "My Sheet",
 *   locale:   "en_US",          // optional
 *   timeZone: "Europe/London",  // optional
 *   sheets: [
 *     {
 *       name:       "Revenue",
 *       tabColor:   "#0f9d58",
 *       frozenRows: 1,
 *       frozenCols: 0,
 *       headers:    ["Month", "Revenue", "Expenses"],
 *       data:       [["Jan", 50000, 30000], ...],
 *       headerFormat: {
 *         bold: true,
 *         backgroundColor: "#1565C0",   // hex or {r,g,b}
 *         foregroundColor: "#FFFFFF",
 *         fontSize: 11,
 *         horizontalAlignment: "CENTER"  // LEFT | CENTER | RIGHT
 *       },
 *       columnWidths: { "A": 160, "B": 120 },   // col letter → pixels
 *       numberFormats: {
 *         "B:C": { type: "CURRENCY", pattern: "£#,##0.00" }
 *       },
 *       conditionalFormats: [  // optional
 *         {
 *           range: "B2:B100",
 *           condition: { type: "NUMBER_LESS", values: [{ userEnteredValue: "0" }] },
 *           format: { backgroundColor: "#FFCCCC" }
 *         }
 *       ],
 *       namedRanges: [{ name: "RevenueData", a1: "A2:C100" }]
 *     }
 *   ]
 * }
 *
 * Returns { spreadsheetId, url, title, sheets[] }
 */
async function createSpreadsheet(schema: Dynamic, folderId?: string) {
  const auth   = getOAuth2Client();
  const sheetsApi = google.sheets({ version: "v4", auth });
  const driveApi  = google.drive({ version: "v3", auth });

  // ── Step 1: Create spreadsheet ────────────────────────────────────────────
  const createRes = await sheetsApi.spreadsheets.create({
    requestBody: {
      properties: {
        title:    schema.title    || "New Spreadsheet",
        locale:   schema.locale,
        timeZone: schema.timeZone
      },
      sheets: (schema.sheets || []).map((s: JsonObject, i: number) => ({
        properties: {
          title:    s.name || `Sheet${i + 1}`,
          index:    i,
          tabColor: s.tabColor ? _hexToColor(s.tabColor) : undefined,
          gridProperties: {
            frozenRowCount:    s.frozenRows  || 0,
            frozenColumnCount: s.frozenCols  || 0
          }
        }
      }))
    }
  });

  const spreadsheetId = createRes.data.spreadsheetId;
  if (!spreadsheetId) throw new Error("Spreadsheet creation returned no ID");

  // Build name → sheetId map from the response
  const sheetMap: JsonObject = {};
  (createRes.data.sheets || []).forEach((s) => {
    if (s.properties?.title && typeof s.properties.sheetId === "number") sheetMap[s.properties.title] = s.properties.sheetId;
  });

  // ── Step 2: Move to folder if specified ───────────────────────────────────
  if (folderId) {
    const fileRes   = await driveApi.files.get({ fileId: spreadsheetId, fields: "parents" });
    const prevParents = (fileRes.data.parents || []).join(",");
    await driveApi.files.update({
      fileId:        spreadsheetId,
      addParents:    folderId,
      removeParents: prevParents,
      fields:        "id, parents"
    });
  }

  // ── Step 3: Write data (headers + rows) ───────────────────────────────────
  const valueUpdates: Dynamic[] = [];
  for (const sheet of (schema.sheets || [])) {
    const rows: Dynamic[] = [];
    if (sheet.headers) rows.push(sheet.headers);
    if (sheet.data)    rows.push(...sheet.data);
    if (rows.length) {
      valueUpdates.push({ range: `'${sheet.name}'!A1`, values: rows });
    }
  }
  if (valueUpdates.length) {
    await sheetsApi.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: "USER_ENTERED", data: valueUpdates }
    });
  }

  // ── Step 4: Formatting via batchUpdate ────────────────────────────────────
  const requests: Dynamic[] = [];

  for (const sheet of (schema.sheets || [])) {
    const sheetId = sheetMap[sheet.name];
    if (sheetId === undefined) continue;

    // Header row formatting
    if (sheet.headerFormat && sheet.headers) {
      const fmt = sheet.headerFormat;
      const cellFmt: JsonObject = { textFormat: {} };
      if (fmt.bold        !== undefined) cellFmt.textFormat.bold      = fmt.bold;
      if (fmt.italic      !== undefined) cellFmt.textFormat.italic    = fmt.italic;
      if (fmt.fontSize)                  cellFmt.textFormat.fontSize  = fmt.fontSize;
      if (fmt.foregroundColor)           cellFmt.textFormat.foregroundColor = _toColor(fmt.foregroundColor);
      if (fmt.backgroundColor)           cellFmt.backgroundColor     = _toColor(fmt.backgroundColor);
      if (fmt.horizontalAlignment)       cellFmt.horizontalAlignment  = fmt.horizontalAlignment;
      if (fmt.verticalAlignment)         cellFmt.verticalAlignment    = fmt.verticalAlignment;
      if (fmt.wrapStrategy)              cellFmt.wrapStrategy         = fmt.wrapStrategy;

      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: 0, endRowIndex: 1,
                   startColumnIndex: 0, endColumnIndex: sheet.headers.length },
          cell:   { userEnteredFormat: cellFmt },
          fields: "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy)"
        }
      });
    }

    // Column widths
    if (sheet.columnWidths) {
      Object.entries(sheet.columnWidths).forEach(([col, width]) => {
        const ci = typeof col === "number" ? col : _colToIndex(col);
        requests.push({
          updateDimensionProperties: {
            range: { sheetId, dimension: "COLUMNS", startIndex: ci, endIndex: ci + 1 },
            properties: { pixelSize: width },
            fields: "pixelSize"
          }
        });
      });
    }

    // Number formats for columns/ranges
    if (sheet.numberFormats) {
      Object.entries(sheet.numberFormats as Record<string, {type: string; pattern: string}>).forEach(([colRange, fmt]) => {
        const gr = _parseA1ToGridRange(colRange, sheetId);
        if (!gr.startRowIndex) gr.startRowIndex = sheet.headers ? 1 : 0; // skip header
        requests.push({
          repeatCell: {
            range: gr,
            cell:   { userEnteredFormat: { numberFormat: { type: fmt.type || "NUMBER", pattern: fmt.pattern || "" } } },
            fields: "userEnteredFormat.numberFormat"
          }
        });
      });
    }

    // Conditional formats
    if (sheet.conditionalFormats) {
      sheet.conditionalFormats.forEach((cf: JsonObject) => {
        const rule: JsonObject = {
          ranges:    [_parseA1ToGridRange(cf.range, sheetId)],
          booleanRule: {
            condition: cf.condition,
            format:    {}
          }
        };
        if (cf.format.backgroundColor) rule.booleanRule.format.backgroundColor = _toColor(cf.format.backgroundColor);
        if (cf.format.textFormat)       rule.booleanRule.format.textFormat       = cf.format.textFormat;
        requests.push({ addConditionalFormatRule: { rule, index: 0 } });
      });
    }

    // Row height for header
    if (sheet.headerRowHeight) {
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: 1 },
          properties: { pixelSize: sheet.headerRowHeight },
          fields: "pixelSize"
        }
      });
    }
  }

  if (requests.length) {
    await sheetsApi.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  }

  // ── Step 5: Named ranges ─────────────────────────────────────────────────
  const nrRequests: Dynamic[] = [];
  for (const sheet of (schema.sheets || [])) {
    const sheetId = sheetMap[sheet.name];
    if (!sheet.namedRanges) continue;
    sheet.namedRanges.forEach((nr: {name: string; a1: string}) => {
      nrRequests.push({
        addNamedRange: {
          namedRange: { name: nr.name, range: _parseA1ToGridRange(nr.a1, sheetId) }
        }
      });
    });
  }
  if (nrRequests.length) {
    await sheetsApi.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: nrRequests } });
  }

  return {
    spreadsheetId,
    url:    `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
    title:  schema.title,
    sheets: Object.keys(sheetMap)
  };
}

/**
 * formatRange(spreadsheetId, requests)
 *
 * Apply raw Sheets API batchUpdate requests (formatting, merges, etc.).
 * Useful for post-creation refinements or standalone formatting tasks.
 */
async function formatRange(spreadsheetId: string, requests: Dynamic) {
  const auth = getOAuth2Client();
  await google.sheets({ version: "v4", auth }).spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests }
  });
  return { requestCount: requests.length };
}

// ─── Presentation & visualization ───────────────────────────────────────────────
// Theme-driven styling / auto-sizing / charts. Callers pass a `theme` object (design
// tokens from theme.json); this layer converts hex→Color and emits batchUpdate requests.
// Ranges are 0-based, end-exclusive. Nothing here hardcodes a colour.

/** Resolve a sheet's numeric sheetId from its title (throws if absent). */
async function getSheetIdByTitle(spreadsheetId: string, sheetTitle: string) {
  const meta = await getSpreadsheetMeta(spreadsheetId);
  const s = (meta.sheets || []).find((x) => x.properties.title === sheetTitle);
  if (!s) throw new Error(`Sheet not found: ${sheetTitle}`);
  return s.properties.sheetId;
}

/**
 * getSheetVisuals(spreadsheetId, sheetTitle!)
 * One read returning everything the presentation layer inspects or verifies:
 * frozen counts, per-column & per-row pixel sizes, banded ranges, charts, conditional formats.
 */
async function getSheetVisuals(spreadsheetId: string, sheetTitle: string) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);
  const res  = await api.spreadsheets.get({
    spreadsheetId,
    ranges: [sheetTitle],
    fields: "sheets(properties(sheetId,title,gridProperties),data(columnMetadata(pixelSize),rowMetadata(pixelSize)),bandedRanges,charts(chartId,spec(title,basicChart(chartType),pieChart)),conditionalFormats)"
  });
  const sheet = (res.data.sheets || []).find((s) => s.properties?.title === sheetTitle) || (res.data.sheets || [])[0];
  if (!sheet?.properties || typeof sheet.properties.sheetId !== "number") throw new Error(`Sheet not found or missing ID: ${sheetTitle}`);
  const gp    = sheet.properties.gridProperties || {};
  const data0 = (sheet.data && sheet.data[0]) || {};
  return {
    sheetId:      sheet.properties.sheetId,
    title:        sheet.properties.title,
    frozenRows:   gp.frozenRowCount || 0,
    frozenCols:   gp.frozenColumnCount || 0,
    columnWidths: (data0.columnMetadata || []).map((c) => c.pixelSize ?? undefined),
    rowHeights:   (data0.rowMetadata   || []).map((r) => r.pixelSize ?? undefined),
    bandedRanges: sheet.bandedRanges || [],
    charts:       sheet.charts || [],
    conditionalFormats: sheet.conditionalFormats || []
  };
}

/**
 * autofit(spreadsheetId, opts)
 * opts: { sheetTitle, startColumn=0, endColumn=null, maxWidth, minWidth, wrap=true, resizeRows=true }
 * Auto-sizes columns to content, caps overly-wide columns (wrapping their text),
 * floors too-narrow columns, then grows row heights to fit wrapped text.
 * autoResizeDimensions fits to content server-side; WRAP auto-grows row height (a capped
 * column would otherwise clip), so we only pin width, never row height.
 * Returns { sheetId, before:[widths], after:[widths], capped:[colIndexes], rowsResized }.
 */
async function autofit(spreadsheetId: string, opts: VisualOptions = {}) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);
  const { sheetTitle } = opts;
  const wrap       = opts.wrap !== false;
  const resizeRows = opts.resizeRows !== false;
  const maxWidth   = opts.maxWidth || null;
  const minWidth   = opts.minWidth || null;

  const v0       = await getSheetVisuals(spreadsheetId, sheetTitle!);
  const sheetId  = v0.sheetId;
  const colCount = v0.columnWidths.length || 26;
  const rowCount = v0.rowHeights.length   || 1000;
  const startColumn = opts.startColumn || 0;
  const endColumn   = opts.endColumn   || colCount;

  // Pass 1: auto-resize columns to fit content.
  await api.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{
      autoResizeDimensions: { dimensions: { sheetId, dimension: "COLUMNS", startIndex: startColumn, endIndex: endColumn } }
    }] }
  });

  const v1     = await getSheetVisuals(spreadsheetId, sheetTitle!);
  const before = v0.columnWidths;
  const fitted = v1.columnWidths;

  // Pass 2: cap wide cols (+ wrap), floor narrow cols, then re-fit row heights.
  const reqs: Dynamic[] = [];
  const capped: number[] = [];
  for (let c = startColumn; c < endColumn; c++) {
    const w = fitted[c];
    if (w === undefined) continue;
    let target = null;
    if (maxWidth && w > maxWidth)      { target = maxWidth; capped.push(c); }
    else if (minWidth && w < minWidth) { target = minWidth; }
    if (target !== null) {
      reqs.push({ updateDimensionProperties: {
        range: { sheetId, dimension: "COLUMNS", startIndex: c, endIndex: c + 1 },
        properties: { pixelSize: target }, fields: "pixelSize" } });
    }
    if (wrap && maxWidth && w > maxWidth) {
      reqs.push({ repeatCell: {
        range: { sheetId, startColumnIndex: c, endColumnIndex: c + 1 },
        cell:  { userEnteredFormat: { wrapStrategy: "WRAP" } },
        fields: "userEnteredFormat.wrapStrategy" } });
    }
  }
  if (reqs.length) {
    await api.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: reqs } });
  }
  // Re-fit row heights in a SEPARATE batch, so auto-resize sees the applied width cap + WRAP
  // (auto-resize is computed against pre-batch state, so it must run after the wrap commits).
  // Run unconditionally: it fits every row to its content — grows wrapped rows, leaves single-line
  // rows at their minimum — and stays correct even when a column was wrapped on a prior run.
  if (resizeRows) {
    await api.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ autoResizeDimensions: { dimensions: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: rowCount } } }] }
    });
  }

  const v2 = await getSheetVisuals(spreadsheetId, sheetTitle!);
  return { sheetId, before, after: v2.columnWidths, capped, rowsResized: !!resizeRows };
}

/**
 * Infer a column's semantic type from sampled (unformatted) values + header name.
 * Returns: currency | percent | date | integer | decimal | text.
 * Heuristic by design — see Part 9 in AGENTS.md; callers can override per column.
 */
function _inferColumnType(header: CellValue, values: CellValue[]) {
  const h = String(header || "").toLowerCase();
  const nonEmpty = values.filter((v: CellValue) => v !== null && v !== undefined && v !== "");
  if (!nonEmpty.length) return "text";

  const dateRe  = /^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$|^\d{1,2}\/\d{1,2}\/\d{2,4}$/;
  const dateish = nonEmpty.filter((v: CellValue) => typeof v === "string" && dateRe.test(String(v).trim()));
  if (dateish.length >= Math.ceil(nonEmpty.length * 0.6)) return "date";

  const nums    = nonEmpty.filter((v: CellValue) => typeof v === "number");
  const numeric = nums.length >= Math.ceil(nonEmpty.length * 0.6);
  if (numeric) {
    if (/date|day|month|year|dob/.test(h)) return "date";
    const pctHeader = /%|percent|margin|rate|growth|share|yield|roi|apr|apy/.test(h);
    if (pctHeader && nums.every((n) => n >= -1.5 && n <= 1.5)) return "percent";
    const curHeader = /£|\$|€|price|cost|revenue|expense|amount|balance|total|value|salary|income|spend|worth|cash|\bfee/.test(h);
    if (curHeader) return "currency";
    return nums.every((n) => Number.isInteger(n)) ? "integer" : "decimal";
  }
  return "text";
}

/**
 * styleTable(spreadsheetId, opts)
 * One-shot beautify of a tabular range from the theme: header treatment + freeze +
 * per-column alignment + number-format inference + header border + banding + autofit.
 * opts: { sheetTitle, range?, theme, headerRows=1, freeze, banding, borders, autofit, inferNumberFormats }
 * With no `range`, styles the A1-anchored used region of the sheet.
 * Returns a summary incl. the inferred column types.
 */
async function styleTable(spreadsheetId: string, opts: VisualOptions = {}) {
  const auth  = getOAuth2Client();
  const api   = sheets(auth);
  const theme = opts.theme || {};
  const sheetTitle = opts.sheetTitle;
  const headerRows = opts.headerRows == null ? 1 : parseInt(String(opts.headerRows));

  const sheetId = await getSheetIdByTitle(spreadsheetId, sheetTitle!);

  // Read used values UNFORMATTED so numbers stay numbers for type inference.
  const rangeA1 = opts.range ? `'${sheetTitle}'!${opts.range}` : `'${sheetTitle}'`;
  const res = await api.spreadsheets.values.get({
    spreadsheetId, range: rangeA1,
    valueRenderOption: "UNFORMATTED_VALUE", dateTimeRenderOption: "SERIAL_NUMBER"
  });
  const values: CellValue[][] = res.data.values || [];
  if (!values.length) throw new Error("No data found to style in range: " + rangeA1);

  // Table extent + offset (respect an explicit range's top-left; else anchor at A1).
  let top = 0, left = 0;
  if (opts.range) {
    const gr = _parseA1ToGridRange(opts.range, sheetId);
    top  = gr.startRowIndex    || 0;
    left = gr.startColumnIndex || 0;
  }
  const numRows = values.length;
  const numCols = values.reduce((m: number, r: CellValue[]) => Math.max(m, r.length), 0);
  const bottom  = top + numRows;   // end-exclusive
  const right   = left + numCols;

  const headers  = values[0] || [];
  const bodyRows = values.slice(headerRows);
  const columnType: string[] = [];
  for (let c = 0; c < numCols; c++) columnType[c] = _inferColumnType(headers[c], bodyRows.map((r: CellValue[]) => r[c]));

  const reqs: Dynamic[] = [];

  // 1. Header treatment
  if (theme.header && headerRows > 0) {
    const H = theme.header, cellFmt: JsonObject = { textFormat: {} };
    if (H.bold !== undefined)  cellFmt.textFormat.bold      = H.bold;
    if (H.fontSize)            cellFmt.textFormat.fontSize   = H.fontSize;
    if (H.foregroundColor)     cellFmt.textFormat.foregroundColor = _toColor(H.foregroundColor);
    if (H.backgroundColor)     cellFmt.backgroundColor       = _toColor(H.backgroundColor);
    if (H.horizontalAlignment) cellFmt.horizontalAlignment   = H.horizontalAlignment;
    if (H.verticalAlignment)   cellFmt.verticalAlignment     = H.verticalAlignment;
    if (H.wrapStrategy)        cellFmt.wrapStrategy           = H.wrapStrategy;
    reqs.push({ repeatCell: {
      range: { sheetId, startRowIndex: top, endRowIndex: top + headerRows, startColumnIndex: left, endColumnIndex: right },
      cell:  { userEnteredFormat: cellFmt },
      fields: "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy)" } });
  }

  // 2. Freeze header row(s)
  if (opts.freeze !== false && headerRows > 0) {
    reqs.push({ updateSheetProperties: {
      properties: { sheetId, gridProperties: { frozenRowCount: top + headerRows } },
      fields: "gridProperties.frozenRowCount" } });
  }

  // 3. Per-column alignment + inferred number format (body rows only — never clobbers banding)
  const align = theme.alignment || {}, nf = theme.numberFormats || {};
  for (let c = 0; c < numCols; c++) {
    const t = columnType[c];
    const fmt: JsonObject = {}, fields: string[] = [];
    const a = align[t] || align.text;
    if (a) { fmt.horizontalAlignment = a; fields.push("horizontalAlignment"); }
    if (opts.inferNumberFormats !== false && nf[t]) { fmt.numberFormat = { type: nf[t].type, pattern: nf[t].pattern }; fields.push("numberFormat"); }
    if (fields.length) {
      reqs.push({ repeatCell: {
        range: { sheetId, startRowIndex: top + headerRows, endRowIndex: bottom, startColumnIndex: left + c, endColumnIndex: left + c + 1 },
        cell:  { userEnteredFormat: fmt },
        fields: "userEnteredFormat(" + fields.join(",") + ")" } });
    }
  }

  // 4. Header bottom border
  if (opts.borders !== false && theme.border && theme.border.headerBottom) {
    const b = theme.border.headerBottom;
    reqs.push({ updateBorders: {
      range: { sheetId, startRowIndex: top, endRowIndex: top + headerRows, startColumnIndex: left, endColumnIndex: right },
      bottom: { style: b.style || "SOLID", color: _toColor(b.color || "#000000") } } });
  }

  // 5. Banding (delete overlapping first — banded ranges can't overlap)
  let banded = false;
  if (opts.banding !== false && theme.banding) {
    const existing = (await getSheetVisuals(spreadsheetId, sheetTitle!)).bandedRanges;
    existing.forEach((br: JsonObject) => reqs.push({ deleteBanding: { bandedRangeId: br.bandedRangeId } }));
    const B = theme.banding, bp: JsonObject = {};
    if (B.headerColor)     bp.headerColor     = _toColor(B.headerColor);
    if (B.firstBandColor)  bp.firstBandColor  = _toColor(B.firstBandColor);
    if (B.secondBandColor) bp.secondBandColor = _toColor(B.secondBandColor);
    reqs.push({ addBanding: { bandedRange: {
      range: { sheetId, startRowIndex: top, endRowIndex: bottom, startColumnIndex: left, endColumnIndex: right },
      rowProperties: bp } } });
    banded = true;
  }

  if (reqs.length) await api.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: reqs } });

  // 6. Autofit last (reads back widths to cap/wrap)
  let fit = null;
  if (opts.autofit !== false) {
    const af = theme.autofit || {};
    fit = await autofit(spreadsheetId, { sheetTitle, maxWidth: af.maxColWidth, minWidth: af.minColWidth, wrap: af.wrapOverCap !== false });
  }

  return {
    sheetId, range: rangeA1, numRows, numCols,
    columnTypes: headers.map((h: CellValue, i: number) => ({ column: h, type: columnType[i] })),
    frozen: opts.freeze !== false && headerRows > 0,
    banded, autofit: fit ? { capped: fit.capped, after: fit.after } : null
  };
}

/**
 * addChart(spreadsheetId, opts)
 * Embed a chart from a data range. First column = domain (labels), remaining = series.
 * opts: { sheetTitle, dataRange, type=line, title, anchor='H2', targetSheetTitle?, newSheet?, theme }
 * type: line|column|bar|area|scatter|pie. Returns { chartId, type, title, dataRange }.
 */
async function addChart(spreadsheetId: string, opts: VisualOptions = {}) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);
  const { sheetTitle, dataRange } = opts;
  const type  = String(opts.type || "line").toUpperCase();
  const title = opts.title || "";
  const chartTheme = (opts.theme && opts.theme.chart) || {};
  const palette = chartTheme.palette || [];

  const srcSheetId = await getSheetIdByTitle(spreadsheetId, sheetTitle!);
  const gr = _parseA1ToGridRange(dataRange!, srcSheetId);
  if (gr.startColumnIndex == null || gr.endColumnIndex == null || gr.startRowIndex == null || gr.endRowIndex == null) {
    throw new Error("add-chart needs a bounded data range like A1:C13");
  }
  const src = (sc: number, ec: number) => ({ sources: [{ sheetId: srcSheetId, startRowIndex: gr.startRowIndex, endRowIndex: gr.endRowIndex, startColumnIndex: sc, endColumnIndex: ec }] });
  const domainCol = gr.startColumnIndex;

  // Position (overlay on a sheet, or a new sheet).
  let position: Dynamic;
  if (opts.newSheet) {
    position = { newSheet: true };
  } else {
    const targetSheetId = opts.targetSheetTitle ? await getSheetIdByTitle(spreadsheetId, opts.targetSheetTitle) : srcSheetId;
    const anchor = _parseA1ToGridRange(opts.anchor || "H2", targetSheetId);
    position = { overlayPosition: {
      anchorCell:  { sheetId: targetSheetId, rowIndex: anchor.startRowIndex || 0, columnIndex: anchor.startColumnIndex || 0 },
      widthPixels: chartTheme.widthPixels || 600, heightPixels: chartTheme.heightPixels || 371 } };
  }

  let spec: Dynamic;
  if (type === "PIE") {
    spec = { title, pieChart: {
      legendPosition: chartTheme.legendPosition && chartTheme.legendPosition.startsWith("BOTTOM") ? "RIGHT_LEGEND" : (chartTheme.legendPosition || "RIGHT_LEGEND"),
      domain: { sourceRange: src(domainCol, domainCol + 1) },
      series: { sourceRange: src(domainCol + 1, domainCol + 2) },
      threeDimensional: false } };
  } else {
    const series: Dynamic[] = [];
    let si = 0;
    for (let c = domainCol + 1; c < gr.endColumnIndex; c++) {
      const s: JsonObject = { series: { sourceRange: src(c, c + 1) }, targetAxis: "LEFT_AXIS" };
      if (palette[si]) s.color = _toColor(palette[si]);
      series.push(s); si++;
    }
    spec = { title, basicChart: {
      chartType: type, legendPosition: chartTheme.legendPosition || "BOTTOM_LEGEND", headerCount: 1,
      domains: [{ domain: { sourceRange: src(domainCol, domainCol + 1) } }], series } };
  }

  const res = await api.spreadsheets.batchUpdate({
    spreadsheetId, requestBody: { requests: [{ addChart: { chart: { spec, position } } }] }
  });
  const chartId = res.data.replies?.[0]?.addChart?.chart?.chartId;
  if (chartId == null) throw new Error("Chart creation returned no chart ID");
  return { chartId, type, title, dataRange };
}

/**
 * conditionalFormat(spreadsheetId, opts)
 * opts: { sheetTitle, range, rule, value?, color?, theme }
 * rule: negative-red | less-than | greater-than | heatmap. Returns { sheetId, rule, range }.
 */
async function conditionalFormat(spreadsheetId: string, opts: VisualOptions = {}) {
  const auth = getOAuth2Client();
  const api  = sheets(auth);
  const { sheetTitle, range } = opts;
  const rule = String(opts.rule || "negative-red");
  const cond = (opts.theme && opts.theme.conditional) || {};
  const sheetId = await getSheetIdByTitle(spreadsheetId, sheetTitle!);
  const gr = _parseA1ToGridRange(range!, sheetId);

  let ruleObj: Dynamic;
  if (rule === "negative-red") {
    ruleObj = { ranges: [gr], booleanRule: {
      condition: { type: "NUMBER_LESS", values: [{ userEnteredValue: "0" }] },
      format: { textFormat: { foregroundColor: _toColor(opts.color || cond.negativeColor || "#c00000") } } } };
  } else if (rule === "less-than" || rule === "greater-than") {
    if (opts.value === undefined) throw new Error(`--rule ${rule} needs --value N`);
    ruleObj = { ranges: [gr], booleanRule: {
      condition: { type: rule === "less-than" ? "NUMBER_LESS" : "NUMBER_GREATER", values: [{ userEnteredValue: String(opts.value) }] },
      format: { backgroundColor: _toColor(opts.color || "#f4cccc") } } };
  } else if (rule === "heatmap") {
    const hm = cond.heatmap || {};
    ruleObj = { ranges: [gr], gradientRule: {
      minpoint: { color: _toColor(hm.min || "#f4cccc"), type: "MIN" },
      midpoint: { color: _toColor(hm.mid || "#ffffff"), type: "PERCENTILE", value: "50" },
      maxpoint: { color: _toColor(hm.max || "#b6d7a8"), type: "MAX" } } };
  } else {
    throw new Error("Unknown --rule: " + rule + " (negative-red|less-than|greater-than|heatmap)");
  }

  await api.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addConditionalFormatRule: { rule: ruleObj, index: 0 } }] } });
  return { sheetId, rule, range };
}

/**
 * addTitle(spreadsheetId, opts)
 * Prepend a title band (and optional subtitle strip) above the existing table:
 * inserts row(s) at the top, writes + merges + styles them from the theme, and bumps
 * the frozen-row count so title + subtitle + any already-frozen header stay pinned.
 * Run AFTER style-table (the insert shifts banding/charts/conditional ranges, which
 * Sheets auto-adjusts; only the frozen count needs fixing, which this does).
 * opts: { sheetTitle, title, subtitle?, width?, theme }. Returns { bandRows, width, frozenRowCount }.
 */
async function addTitle(spreadsheetId: string, opts: VisualOptions = {}) {
  const api   = sheets(getOAuth2Client());
  const theme = opts.theme || {};
  const sheetTitle = opts.sheetTitle;
  const bandRows = opts.subtitle ? 2 : 1;

  const v = await getSheetVisuals(spreadsheetId, sheetTitle!);
  const sheetId = v.sheetId;

  // Span width: explicit, else the current header-row width.
  let width = opts.width ? parseInt(String(opts.width)) : null;
  if (!width) {
    const res = await api.spreadsheets.values.get({ spreadsheetId, range: `'${sheetTitle}'!1:1` });
    width = ((res.data.values && res.data.values[0]) || []).length || 1;
  }

  // 1. Insert band rows at the very top.
  await api.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ insertDimension: { range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: bandRows }, inheritFromBefore: false } }] }
  });

  // 2. Write the text.
  const data = [{ range: `'${sheetTitle}'!A1`, values: [[opts.title || ""]] }];
  if (opts.subtitle) data.push({ range: `'${sheetTitle}'!A2`, values: [[opts.subtitle]] });
  await api.spreadsheets.values.batchUpdate({ spreadsheetId, requestBody: { valueInputOption: "RAW", data } });

  // 3. Merge + style each band row, then bump the freeze.
  const reqs: Dynamic[] = [];
  const bands = [{ tok: theme.title, row: 0 }];
  if (opts.subtitle) bands.push({ tok: theme.subtitle, row: 1 });
  bands.forEach(({ tok, row }) => {
    reqs.push({ mergeCells: { range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: 0, endColumnIndex: width }, mergeType: "MERGE_ALL" } });
    if (!tok) return;
    const cf: JsonObject = { textFormat: {} };
    if (tok.bold !== undefined)  cf.textFormat.bold     = tok.bold;
    if (tok.fontSize)            cf.textFormat.fontSize  = tok.fontSize;
    if (tok.foregroundColor)     cf.textFormat.foregroundColor = _toColor(tok.foregroundColor);
    if (tok.backgroundColor)     cf.backgroundColor      = _toColor(tok.backgroundColor);
    if (tok.horizontalAlignment) cf.horizontalAlignment  = tok.horizontalAlignment;
    if (tok.verticalAlignment)   cf.verticalAlignment    = tok.verticalAlignment;
    reqs.push({ repeatCell: {
      range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: 0, endColumnIndex: width },
      cell:  { userEnteredFormat: cf },
      fields: "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment)" } });
    if (tok.rowHeight) reqs.push({ updateDimensionProperties: {
      range: { sheetId, dimension: "ROWS", startIndex: row, endIndex: row + 1 }, properties: { pixelSize: tok.rowHeight }, fields: "pixelSize" } });
  });
  const frozenRowCount = (v.frozenRows || 0) + bandRows;
  reqs.push({ updateSheetProperties: { properties: { sheetId, gridProperties: { frozenRowCount } }, fields: "gridProperties.frozenRowCount" } });
  await api.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: reqs } });

  return { sheetId, bandRows, width, frozenRowCount };
}

/**
 * addSparklines(spreadsheetId, opts)
 * Write an inline =SPARKLINE per row: for each row of the source block, a word-sized
 * chart of that row lands in the target column at the same row.
 * opts: { sheetTitle, sourceRange (e.g. "B2:M13"), targetCol (e.g. "N"), type, color, negColor, theme }
 * type: line | column | bar | winloss. Returns { count, targetRange, type }.
 */
async function addSparklines(spreadsheetId: string, opts: VisualOptions = {}) {
  const api = sheets(getOAuth2Client());
  const sp  = (opts.theme && opts.theme.sparkline) || {};
  const { sheetTitle } = opts;
  const type = String(opts.type || sp.type || "line").toLowerCase();
  const color      = opts.color      || sp.color      || null;
  const negColor   = opts.negColor   || sp.negColor   || null;
  const highColor  = opts.highColor  || sp.highColor  || null;
  const lowColor   = opts.lowColor   || sp.lowColor   || null;
  const firstColor = opts.firstColor || sp.firstColor || null;
  const lastColor  = opts.lastColor  || sp.lastColor  || null;

  const m = String(opts.sourceRange || "").match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!m) throw new Error("sparkline --source needs a bounded block like B2:M13");
  const c1 = m[1], startRow = parseInt(m[2]), c2 = m[3], endRow = parseInt(m[4]);
  const targetCol = String(opts.targetCol || "").toUpperCase().replace(/[^A-Z]/g, "");
  if (!targetCol) throw new Error("sparkline --target needs a column letter, e.g. N");

  const pairs = [`"charttype","${type}"`];
  if (color) pairs.push(`"color","${color}"`);
  // Column/bar/winloss sparklines can colour the peak, trough, endpoints and negatives
  // individually — line sparklines only take a single colour.
  if (type === "column" || type === "bar" || type === "winloss") {
    if (highColor)  pairs.push(`"highcolor","${highColor}"`);
    if (lowColor)   pairs.push(`"lowcolor","${lowColor}"`);
    if (firstColor) pairs.push(`"firstcolor","${firstColor}"`);
    if (lastColor)  pairs.push(`"lastcolor","${lastColor}"`);
    if (negColor)   pairs.push(`"negcolor","${negColor}"`);
  } else if (negColor) {
    pairs.push(`"negcolor","${negColor}"`);
  }
  const optStr = `{${pairs.join(";")}}`;

  const values: string[][] = [];
  for (let r = startRow; r <= endRow; r++) values.push([`=SPARKLINE(${c1}${r}:${c2}${r}, ${optStr})`]);
  const targetRange = `'${sheetTitle}'!${targetCol}${startRow}:${targetCol}${endRow}`;
  await api.spreadsheets.values.update({ spreadsheetId, range: targetRange, valueInputOption: "USER_ENTERED", requestBody: { values } });
  return { count: values.length, targetRange, type };
}

/**
 * markInputVsComputed(spreadsheetId, opts)
 * Style each non-empty cell in a range by kind: formula cells (computed / derived) vs
 * literal cells (input / editable), from theme.marking. Makes it obvious which cells are
 * safe to type into and which are driven by formulas.
 * opts: { sheetTitle, range, theme }. Returns { input, computed, cells }.
 */
async function markInputVsComputed(spreadsheetId: string, opts: VisualOptions = {}) {
  const api  = sheets(getOAuth2Client());
  const mark = (opts.theme && opts.theme.marking) || {};
  const { sheetTitle, range } = opts;
  const sheetId = await getSheetIdByTitle(spreadsheetId, sheetTitle!);
  const gr = _parseA1ToGridRange(range!, sheetId);
  const startRow = gr.startRowIndex || 0, startCol = gr.startColumnIndex || 0;

  const res  = await api.spreadsheets.values.get({ spreadsheetId, range: `'${sheetTitle}'!${range}`, valueRenderOption: "FORMULA" });
  const rows: CellValue[][] = res.data.values || [];

  const buildFmt = (kind: "input" | "computed") => {
    const tok = mark[kind];
    if (!tok) return null;
    const cf: JsonObject = { textFormat: {} }, fields: string[] = [];
    if (tok.italic !== undefined) { cf.textFormat.italic = tok.italic; fields.push("textFormat.italic"); }
    if (tok.bold   !== undefined) { cf.textFormat.bold   = tok.bold;   fields.push("textFormat.bold"); }
    if (tok.foregroundColor)      { cf.textFormat.foregroundColor = _toColor(tok.foregroundColor); fields.push("textFormat.foregroundColor"); }
    if (tok.backgroundColor)      { cf.backgroundColor = _toColor(tok.backgroundColor); fields.push("backgroundColor"); }
    return fields.length ? { cf, fields } : null;
  };
  const inputFmt = buildFmt("input"), computedFmt = buildFmt("computed");

  // Modeling convention: colour hardcoded NUMBER inputs, mark formulas as derived, and
  // leave text labels untouched (colouring every label just floods the table). Set
  // opts.includeText / theme.marking.includeText to also mark text literals as input.
  const includeText = opts.includeText || mark.includeText || false;
  const reqs: Dynamic[] = []; let nInput = 0, nComputed = 0;
  rows.forEach((row: CellValue[], ri: number) => {
    (row || []).forEach((cell: CellValue, ci: number) => {
      if (cell === "" || cell === null || cell === undefined) return;
      const isFormula = typeof cell === "string" && cell.startsWith("=");
      const isNumber  = typeof cell === "number";
      let f = null;
      if (isFormula)                     { f = computedFmt; nComputed++; }
      else if (isNumber || includeText)  { f = inputFmt;    nInput++; }
      else return; // text label — leave as-is
      if (!f) return;
      reqs.push({ repeatCell: {
        range: { sheetId, startRowIndex: startRow + ri, endRowIndex: startRow + ri + 1, startColumnIndex: startCol + ci, endColumnIndex: startCol + ci + 1 },
        cell:  { userEnteredFormat: f.cf },
        fields: "userEnteredFormat(" + f.fields.join(",") + ")" } });
    });
  });
  if (reqs.length) await api.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: reqs } });
  return { input: nInput, computed: nComputed, cells: reqs.length };
}

/**
 * addSheetTab(spreadsheetId, name, options?)
 * options: { tabColor, index, hidden }
 */
async function addSheetTab(spreadsheetId: string, name: string, options?: {tabColor?: string; index?: number; hidden?: boolean}) {
  const auth = getOAuth2Client();
  const api  = google.sheets({ version: "v4", auth });
  const props: JsonObject = { title: name };
  if (options && options.tabColor) props.tabColor  = _hexToColor(options.tabColor);
  if (options && options.index !== undefined) props.index = options.index;
  if (options && options.hidden) props.hidden = true;

  const res = await api.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: props } }] }
  });
  const added = res.data.replies?.[0]?.addSheet?.properties;
  if (!added || typeof added.sheetId !== "number") throw new Error("Tab creation returned no sheet ID");
  return { sheetId: added.sheetId, name: added.title };
}

/**
 * deleteSheetTab(spreadsheetId, sheetNameOrId)
 * sheetNameOrId: string (name) or number (sheetId)
 */
async function deleteSheetTab(spreadsheetId: string, sheetNameOrId: string | number) {
  const auth = getOAuth2Client();
  const api  = google.sheets({ version: "v4", auth });
  let sheetId = typeof sheetNameOrId === "number" ? sheetNameOrId : null;
  if (sheetId === null) {
    const meta = await api.spreadsheets.get({ spreadsheetId, fields: "sheets.properties" });
    const found = (meta.data.sheets || []).find((s) => s.properties?.title === sheetNameOrId);
    if (!found?.properties || typeof found.properties.sheetId !== "number") throw new Error(`Sheet tab not found: ${sheetNameOrId}`);
    sheetId = found.properties.sheetId;
  }
  await api.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ deleteSheet: { sheetId } }] }
  });
  return { deleted: sheetNameOrId, sheetId };
}

/**
 * renameSheetTab(spreadsheetId, currentName, newName)
 */
async function renameSheetTab(spreadsheetId: string, currentName: string, newName: string) {
  const auth = getOAuth2Client();
  const api  = google.sheets({ version: "v4", auth });
  const meta = await api.spreadsheets.get({ spreadsheetId, fields: "sheets.properties" });
  const found = (meta.data.sheets || []).find((s) => s.properties?.title === currentName);
  if (!found?.properties) throw new Error(`Sheet tab not found: ${currentName}`);
  await api.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        updateSheetProperties: {
          properties: { sheetId: found.properties.sheetId, title: newName },
          fields: "title"
        }
      }]
    }
  });
  return { oldName: currentName, newName };
}

/**
 * findDriveFolder(name, parentId?)
 * Search Drive for a folder by name. Returns { id, name, parents } or null.
 */
async function findDriveFolder(name: string, parentId?: string) {
  const auth = getOAuth2Client();
  const api  = google.drive({ version: "v3", auth });
  let q = `mimeType='application/vnd.google-apps.folder' and name='${name.replace(/'/g, "\\'")}' and trashed=false`;
  if (parentId) q += ` and '${parentId}' in parents`;
  const res = await api.files.list({ q, fields: "files(id,name,parents)", pageSize: 10 });
  return res.data.files?.[0] || null;
}

/**
 * createDriveFolder(name, parentId?)
 * Create a new folder in Drive. Returns { id, name }.
 */
async function createDriveFolder(name: string, parentId?: string) {
  const auth = getOAuth2Client();
  const api  = google.drive({ version: "v3", auth });
  const meta: JsonObject = { name, mimeType: "application/vnd.google-apps.folder" };
  if (parentId) meta.parents = [parentId];
  const res = await api.files.create({ requestBody: meta, fields: "id, name" });
  return { id: res.data.id, name: res.data.name };
}

/**
 * compareSnapshots(snap1, snap2)
 * Pure function — no API calls.
 * snap1, snap2: objects returned by snapshotWorkbook (the .data field).
 * Returns { added, removed, changed, unchanged, hashChanged }
 */
function compareSnapshots(snap1: WorkbookSnapshot, snap2: WorkbookSnapshot) {
  const sheets1 = snap1.sheets || [];
  const sheets2 = snap2.sheets || [];
  const map1: Record<string, SheetSnapshot> = {};
  const map2: Record<string, SheetSnapshot> = {};
  sheets1.forEach((s: SheetSnapshot) => { map1[s.name] = s; });
  sheets2.forEach((s: SheetSnapshot) => { map2[s.name] = s; });

  const added   = sheets2.filter((s: SheetSnapshot) => !map1[s.name]).map((s: SheetSnapshot) => s.name);
  const removed = sheets1.filter((s: SheetSnapshot) => !map2[s.name]).map((s: SheetSnapshot) => s.name);
  const changed: Array<{name: string; diffs: string[]}> = [];
  let unchanged = 0;

  sheets1.forEach((s: SheetSnapshot) => {
    const s2 = map2[s.name];
    if (!s2) return;
    const diffs: string[] = [];
    if (s.lastRow    !== s2.lastRow)    diffs.push(`rows: ${s.lastRow} → ${s2.lastRow}`);
    if (s.lastColumn !== s2.lastColumn) diffs.push(`cols: ${s.lastColumn} → ${s2.lastColumn}`);
    if (s.hidden     !== s2.hidden)     diffs.push(`hidden: ${s.hidden} → ${s2.hidden}`);
    if (diffs.length) changed.push({ name: s.name, diffs });
    else unchanged++;
  });

  const nr1Names = new Set((snap1.namedRanges || []).map((n: {name: string}) => n.name));
  const nr2Names = new Set((snap2.namedRanges || []).map((n: {name: string}) => n.name));
  const addedNR   = [...nr2Names].filter((n: string) => !nr1Names.has(n));
  const removedNR = [...nr1Names].filter((n: string) => !nr2Names.has(n));

  return {
    snapshotHash1: snap1.snapshotHash,
    snapshotHash2: snap2.snapshotHash,
    hashChanged:   snap1.snapshotHash !== snap2.snapshotHash,
    sheets:  { added, removed, changed, unchanged },
    namedRanges: { added: addedNR, removed: removedNR }
  };
}

/**
 * runScript(scriptId, functionName, params, devMode)
 *
 * Calls an Apps Script function via the Execution API (REST, not clasp run).
 * Requires:
 *  - Apps Script API enabled on the associated Google Cloud project
 *  - The Apps Script project linked to that GCP project
 *  - OAuth token has script.projects scope (run: node sheetops.js auth)
 *
 * devMode=true  → runs the latest saved version (head deployment) — use during dev
 * devMode=false → runs the published/deployed version — use in production
 */
async function runScript(scriptId: string, functionName: string, params: unknown[], devMode: boolean) {
  const auth   = getOAuth2Client();
  const script = google.script({ version: "v1", auth });

  const res = await script.scripts.run({
    scriptId,
    requestBody: {
      function:   functionName,
      parameters: params   || [],
      devMode:    devMode !== false  // default true
    }
  });

  const data = res.data;

  // Apps Script Execution API embeds errors inside the 200 response body
  if (data.error) {
    const e       = data.error;
    const details = (e.details || [])
      .map((d: {errorMessage?: string; errorType?: string}) => d.errorMessage || d.errorType || JSON.stringify(d))
      .filter(Boolean).join("; ");
    throw new Error(`Script error [${e.code}]: ${e.message}${details ? " — " + details : ""}`);
  }

  return data.response ? data.response.result : null;
}

export {
  // auth
  getOAuth2Client,
  hasSpreadsheetScope,
  runAuthFlow,
  getUserEmail,
  migrateDefaultAccount,
  listAccounts,
  switchAccount,
  removeAccount,
  // sheets
  getSpreadsheetMeta,
  healthcheck,
  snapshotWorkbook,
  readRange,
  writeRangeDryRun,
  appendRowsDryRun,
  clearRangeDryRun,
  writeRange,
  appendRows,
  clearRange,
  backupSpreadsheet,
  logOperation,
  runValidation,
  exportRange,
  findInSheet,
  // creation & formatting
  createSpreadsheet,
  formatRange,
  // presentation & visualization
  getSheetIdByTitle,
  getSheetVisuals,
  autofit,
  styleTable,
  addChart,
  conditionalFormat,
  addTitle,
  addSparklines,
  markInputVsComputed,
  addSheetTab,
  deleteSheetTab,
  renameSheetTab,
  findDriveFolder,
  createDriveFolder,
  compareSnapshots,
  // script
  runScript,
  // utils
  hash,
  SHEETOPS_CREDS
};
