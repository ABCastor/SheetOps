# CLAUDE.md: SheetOps Project {{PROJECT_NAME}}

## Identity

- Project name: `{{PROJECT_NAME}}`
- Spreadsheet ID: `{{SPREADSHEET_ID}}`
- Script ID: `{{SCRIPT_ID}}`
- Root: `<SHEETOPS_ROOT>/projects/{{PROJECT_NAME}}/`

## Current status

<!-- Update after each session -->
- AgentOps installed: NO
- Last snapshot: none
- Last backup: none
- Acceptance test: NOT RUN

## Rules

1. No writes to live data before acceptance test passes.
2. No Apps Script push without showing git diff and getting explicit approval.
3. No formula overwrite without `allowFormulaOverwrite: true`.
4. No row deletion without backup and `confirmDestructive: true`.
5. No writes to protected or hidden sheets without the corresponding flags.
6. All writes go through patch files. Never call `agent_writeRange` directly.
7. Always run `dry-run-patch` before `apply-patch`.
8. Always take a backup before any destructive operation.
9. Writes over 100 cells require explicit approval.

## Named ranges

<!-- e.g. Q1Targets: Sheet1!B2:D2 -->
Run `sheetops snapshot --project {{PROJECT_NAME}}` to populate.

## Known sensitive areas

<!-- List sheets, ranges, or functions that need extra care -->
None documented yet.

## Session start

```bash
node bin/sheetops.js pull-script --project {{PROJECT_NAME}}
node bin/sheetops.js snapshot   --project {{PROJECT_NAME}}
node bin/sheetops.js health     --project {{PROJECT_NAME}}
```

## Apply a change

```bash
# Create patch file in ops/patches/, then:
node bin/sheetops.js dry-run-patch --project {{PROJECT_NAME}} --patch ops/patches/my-patch.json
node bin/sheetops.js backup        --project {{PROJECT_NAME}} --reason "pre-patch"
node bin/sheetops.js apply-patch   --project {{PROJECT_NAME}} --patch ops/patches/my-patch.json --confirmed
node bin/sheetops.js validate      --project {{PROJECT_NAME}}
```

## Rollback

1. Open Google Drive and find the backup copy named `{{PROJECT_NAME}}__backup__YYYYMMDD-HHMMSS`.
2. Restore via Drive or use the Sheet's File > Version history.
3. For Apps Script: `cd appsscript && git checkout <previous-commit>`, then `push-script --confirmed`.
