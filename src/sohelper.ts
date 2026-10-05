#!/usr/bin/env node
/**
 * sohelper.js — low-level Sheets API helper for agent sessions.
 *
 * Reads formula text and submits raw requests for existing agent integrations.
 * Uses the same OAuth client + active account as the SheetOps CLI.
 *
 * Usage:
 *   node sohelper.js get    <project> "<Sheet>!A1:C10" [--formulas]
 *   node sohelper.js values <project> <payload.json>     # values.batchUpdate (USER_ENTERED)
 *   node sohelper.js batch  <project> <requests.json>    # spreadsheets.batchUpdate
 *
 * values payload.json: [{ "range": "Sheet!A1", "values": [["x", 1]] }, ...]
 * requests.json:       [{ "repeatCell": {...} }, ...]  (raw batchUpdate requests)
 */
import * as fs from "node:fs";
import * as path from "node:path";
const { google } = require("googleapis");
const { getOAuth2Client } = require("./lib/sheets-api.js");

function spreadsheetIdFor(project: string) {
  const cfg = JSON.parse(
    fs.readFileSync(path.join(__dirname, "..", "projects", project, "project.config.json"), "utf8")
  );
  return cfg.spreadsheetId;
}

async function main() {
  const [cmd, project, arg, flag] = process.argv.slice(2);
  if (!cmd || !project || !arg) {
    console.error("usage: sohelper.js get|values|batch <project> <range|file.json> [--formulas]");
    process.exit(1);
  }
  const auth = getOAuth2Client();
  const sheets = google.sheets({ version: "v4", auth });
  const spreadsheetId = spreadsheetIdFor(project);

  if (cmd === "get") {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: arg,
      valueRenderOption: flag === "--formulas" ? "FORMULA" : "UNFORMATTED_VALUE",
      dateTimeRenderOption: "FORMATTED_STRING",
    });
    console.log(JSON.stringify(res.data.values || [], null, 1));
  } else if (cmd === "values") {
    const data = JSON.parse(fs.readFileSync(arg, "utf8"));
    const res = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: "USER_ENTERED", data },
    });
    console.log(`updated ${res.data.totalUpdatedCells} cells in ${res.data.totalUpdatedSheets} sheet(s)`);
  } else if (cmd === "batch") {
    const requests = JSON.parse(fs.readFileSync(arg, "utf8"));
    const res = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });
    console.log(`batchUpdate ok: ${(res.data.replies || []).length} replies`);
  } else {
    console.error(`unknown command: ${cmd}`);
    process.exit(1);
  }
}

main().catch((e: {response?: {data?: {error?: {message?: string}}}; message: string}) => {
  console.error(e.response?.data?.error?.message || e.message);
  process.exit(1);
});
