# SheetOps Operating Protocol

This document governs every session where an AI agent interacts with a Google Sheet via SheetOps. It applies across all projects.

---

## Part 1: Onboarding a New Sheet

Run these steps in order before doing anything else with a new Sheet.

### Step 1: Create the project folder

```bash
node bin/sheetops.js add-sheet --url "SHEET_URL" [--name "my-name"]
# or with explicit IDs:
node bin/sheetops.js add-sheet --spreadsheet-id "ID" --script-id "SCRIPT_ID"
```

### Step 2: Pull the bound Apps Script project

```bash
node bin/sheetops.js pull-script --project PROJECT_NAME
```

If the clone fails due to auth or permissions, note the blocker and continue with sheet inspection only.

### Step 3: Confirm git initialised

`add-sheet` initialises git automatically. Verify:

```bash
cd projects/PROJECT_NAME/appsscript && git log --oneline
```

### Step 4: Snapshot the workbook

```bash
node bin/sheetops.js snapshot --project PROJECT_NAME
```

This writes `docs/workbook-map.md` automatically.

### Step 5: Review the maps

Read and annotate:

- `docs/workbook-map.md`: all sheets, named ranges, protected ranges, triggers
- `docs/apps-script-map.md`: all .gs files, functions, triggers

Add any risky areas to the project's `CLAUDE.md`.

### Step 6: Identify risks

Check for:

- Hidden sheets
- Protected ranges
- Formulas that drive other calculations
- External service calls (UrlFetchApp, Gmail, Drive, Calendar)
- Triggers (time-driven, form submit, on-edit)
- Shared or published deployments

### Step 7: Install AgentOps.gs

Only after showing the diff and receiving explicit approval:

1. Copy `templates/AgentOps.gs` into `appsscript/`
2. Show the diff: `git diff`
3. Explain what AgentOps.gs does
4. After approval: `node bin/sheetops.js push-script --project PROJECT_NAME --confirmed`

### Step 8: Run the acceptance test

```bash
node bin/sheetops.js run-script --project PROJECT_NAME --function agent_acceptanceTest
```

All 12 checks must show `pass: true`.

### Step 9: Fix failures and repeat

Fix any failing test, re-push with approval, re-run until all pass.

### Step 10: Declare ready

Update `project.config.json`: set `agentOpsInstalled: true` and `status: "ready"`. Update the project's `CLAUDE.md` status section.

---

## Part 2: Formula Awareness

Formulas are first-class data, not an afterthought. A cell showing `£42,500` might be `=SUMIF(B:B,"Jan",C:C)`. The formula tells you how the sheet works; the value tells you only what it shows right now.

`agent_readRange` returns both `values` and `formulas` for every cell. Always inspect both. The user should never have to ask "is there a formula behind this?" — you have the data, report it without being asked.

Rules:

- When reporting cell contents, lead with the formula if one exists. State what it references and what it derives.
- Before writing to any cell, read the range first and check whether a formula is already there. Decide explicitly whether the patch replaces the formula or preserves it. Flag this to the user.
- If a formula references another sheet, a named range, or an external source (`IMPORTRANGE`, `QUERY`, `GOOGLEFINANCE`, etc.), say so explicitly.
- Formulas are valid values in a patch. Write them as strings starting with `=`:

```json
{
  "type": "setValues",
  "target": { "namedRange": "ProfitCalc" },
  "values": [["=C2-D2", "=E2/C2*100"]]
}
```

When creating a Sheet with `create-sheet`, prefer formulas over static values wherever data should update automatically. A well-built Sheet uses formulas to derive totals, rates, and summaries — not hardcoded numbers the user has to maintain by hand.

## Part 3: Every Task on a Sheet

Follow this checklist for every task.

**Pre-task**

- Pull latest: `sheetops pull-script --project NAME`
- Check git: `cd appsscript && git status`
- Snapshot: `sheetops snapshot --project NAME`
- Read the ranges needed — check both values AND formulas for every target cell

**Planning**

- Propose a specific, concrete plan
- Identify exact ranges and named ranges affected
- Flag any formula, hidden sheet, or protected range risks

**Sheet content changes**

- Create a patch file at `ops/patches/YYYYMMDD-HHMMSS-description.json`
- Validate against `templates/patch-schema.json`
- Dry-run: `sheetops dry-run-patch --project NAME --patch ops/patches/FILE.json`
- Review dry-run output for warnings
- Backup: `sheetops backup --project NAME --reason "..."`
- Present plan and dry-run summary
- Wait for explicit approval
- Apply: `sheetops apply-patch --project NAME --patch FILE.json --confirmed`
- Validate: `sheetops validate --project NAME`
- Commit local changes
- Report exactly what changed

**Apps Script changes**

- Pull latest
- Check git status
- Identify the specific functions being changed
- Modify only the local `.gs` file(s)
- Show the full diff: `cd appsscript && git diff`
- Wait for explicit approval
- Push: `sheetops push-script --project NAME --confirmed`
- Validate
- Commit
- Report what changed

---

## Part 4: Safety Limits

These cannot be overridden by task context.

| Condition | Required |
|-----------|----------|
| Write over 100 cells | `confirmLarge: true` and explicit approval |
| Formula cell in target | `allowFormulaOverwrite: true` and explicit approval |
| Hidden sheet target | `allowHiddenSheet: true` and explicit approval |
| Protected range target | `allowProtected: true` and explicit approval |
| Row deletion | `confirmDestructive: true`, backup, and explicit approval |
| Range clear | `confirmDestructive: true` and explicit approval |
| Apps Script push | `--confirmed` flag and explicit approval |
| Trigger or deployment change | Explicit approval |
| Sharing or permission change | Not permitted |

---

## Part 5: Patch File Naming

```
ops/patches/YYYYMMDD-HHMMSS-short-description.json
```

Example: `ops/patches/20260101-143015-update-q1-targets.json`

After successful application, patches move to `ops/patches/applied/`.

---

## Part 6: Rollback

**Sheet data**

1. Open Google Drive and find the backup copy: `SHEET_NAME__backup__YYYYMMDD-HHMMSS`
2. Restore manually, or use the Sheet's File > Version history

**Apps Script**

```bash
cd projects/PROJECT_NAME/appsscript
git log --oneline              # find the target commit
git checkout <commit-hash> -- .
git diff HEAD                  # review
node bin/sheetops.js push-script --project PROJECT_NAME --confirmed
```

---

## Part 7: Pre-flight Checklist

Run before any real work on a new Sheet:

- Node, npm, git, clasp all respond to `--version`
- `clasp login --status` shows an authenticated account
- `add-sheet` creates the project directory
- `pull-script` clones the bound script, or logs a clear error
- Git has at least one commit in `appsscript/`
- `snapshot` writes a JSON file to `ops/snapshots/`
- `workbook-map.md` is populated with real sheet data
- `apps-script-map.md` lists real functions
- AgentOps.gs installed and pushed after approval
- `agent_healthcheck()` returns `ok: true`
- `agent_snapshotWorkbook()` returns `ok: true`
- `agent_readRange()` reads a small safe range successfully
- `agent_backupSpreadsheet()` creates a Drive copy
- `dry-run-patch` on a test patch returns no errors
- `agent_writeRange()` writes to `__AGENT_TEST` sheet successfully
- `agent_runValidation()` passes all standard checks
- `__AGENT_OPS_LOG` sheet exists and has entries
- `log-summary` shows recent log entries
- Rollback instructions are documented in the project's `CLAUDE.md`
- `project.config.json` shows `status: "ready"`
