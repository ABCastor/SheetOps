<h1><img src="docs/readme-header.svg" width="112" height="72" align="absmiddle" alt=""> <img src="docs/readme-title.svg" width="160" height="40" align="absmiddle" alt="SheetOps*"></h1>

[![Checks](https://github.com/ABCastor/SheetOps/actions/workflows/checks.yml/badge.svg?branch=main)](https://github.com/ABCastor/SheetOps/actions/workflows/checks.yml)

SheetOps is a local CLI for Google Sheets and Apps Script. It reads ranges, previews JSON patches, copies Drive backups, and applies confirmed changes through Google's APIs.

## Install

Use [Node.js](https://nodejs.org/) 18 or later. Script pull and push also need [clasp](https://github.com/google/clasp).

```sh
git clone https://github.com/ABCastor/SheetOps.git
cd SheetOps
npm ci
node bin/sheetops.js --help
```

The runtime source is TypeScript. `npm ci` builds it into `dist/`; the original `bin/sheetops.js`, `lib/sheets-api.js`, and `sohelper.js` paths remain available for existing integrations. Requiring the package root returns the API exports; invoke the CLI through its bin entry or `bin/sheetops.js`.

For authenticated commands, supply your own Google OAuth client. Follow [SETUP.md](SETUP.md), then run `node bin/sheetops.js auth` and `node bin/sheetops.js init`. Credentials stay in your home directory; workbook configuration, snapshots, and logs stay local and are excluded from Git.

## Read and change a range

```sh
node bin/sheetops.js add-sheet --url "YOUR_SHEET_URL" --name demo
node bin/sheetops.js read --project demo --sheet Data --a1 A1:B2
node bin/sheetops.js dry-run-patch --project demo --patch patch.json
node bin/sheetops.js apply-patch --project demo --patch patch.json --confirmed
```

A minimal `patch.json`:

```json
{
  "operationId": "20261005-120000-update-input",
  "project": "demo",
  "reason": "Update the agreed input value",
  "backupRequired": true,
  "operations": [
    {
      "type": "setValues",
      "target": { "sheetName": "Data", "a1": "A1" },
      "values": [[42]]
    }
  ]
}
```

`apply-patch` always requires `--confirmed`. It validates every operation before writing, creates a backup by default, and aborts when a required backup copy fails. A retention cleanup failure is reported separately after the copy succeeds. Value payloads above 100 cells need `confirmLarge: true` on that operation; clearing a range needs `confirmDestructive: true`. For `setValues`, add `expectedHash` from a previous range read to reject changes to the displayed values. The hash does not detect changes hidden by formatting or a formula with the same displayed result; append and clear operations reject this unsupported field. Configure `defaultBackupFolderId` in `sheetops.config.json` to put copies in a separate Drive folder; optional retention moves older copies with generated backup names to Drive's trash while preserving the newly created copy.

The REST patch path supports `setValues`, `appendRows`, and `clearRange`. It rejects unknown fields, other operation types, and `expectedWorkbookHash`. [The patch schema](templates/patch-schema.json) describes this REST format. `appendRows` requires a sheet name and accepts an optional A1 table anchor; preview and apply use the same anchor. Google chooses the final append row during apply. Clear previews count a bounded rectangle, including empty cells, or report that its extent is unknown. Patches are sequential, so an API failure can leave earlier operations applied. A backup gives you a recovery copy; it does not make the patch atomic. A local application record blocks reuse of the same operation ID after an attempt. If a command fails, inspect the workbook and `ops/patches/attempts/` before planning recovery; do not rerun the patch blindly. Review the current patch before `--confirmed`: dry-run does not save approval or bind it to the file contents.

`read` returns formatted values and a hash, and `setValues` reads back the result. Formula, hidden-sheet, and protected-range checks are not enforced by this REST path; their `allowFormulaOverwrite`, `allowHiddenSheet`, and `allowProtected` flags are rejected as unsupported. Inspect the workbook before confirming changes. The optional [AgentOps.gs](templates/AgentOps.gs) bridge has separate server-side protections. `sohelper.js` and formatting commands submit direct API requests and do not use the patch approval gates.

## More commands

Run `--help` for workbook snapshots, CSV/JSON export, range search, table styling, charts, and tab management. Apps Script workflows use `pull-script`, `push-script --confirmed`, and `run-script`; execution needs the Google Cloud association described in [SETUP.md](SETUP.md). Google Docs documents are not supported.

```sh
npm run check
```

The checks compile the source, run fake-API contract tests without Google credentials, and scan for personal paths and email addresses. They do not replace testing with your own workbook and permissions.

## License

[Apache License 2.0](LICENSE) for this release's owned code. [NOTICE](NOTICE) retains the earlier MIT notice. Dependencies keep their own terms.

<p><a href="https://abcastor.com"><img src="docs/castor-footer.svg" width="350" alt="Chip, the Castor beaver, by Castor, we give a dam"></a></p>
