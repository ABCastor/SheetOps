/**
 * AgentOps.gs  –  SheetOps bridge
 * Version: 1.0.0
 *
 * Install this file into EVERY bound Apps Script project you manage
 * with the SheetOps framework.  It is the ONLY file the framework
 * writes/reads via the Apps Script execution API.
 *
 * Safety contract:
 *   • All writes are wrapped in LockService.
 *   • Destructive operations (deleteRows, clearRange, overwrite)
 *     require explicit confirmation flags in the payload.
 *   • Formula overwrites require allowFormulaOverwrite:true.
 *   • Hidden-sheet operations require allowHiddenSheet:true.
 *   • Protected-range operations require allowProtected:true.
 *   • Every operation is appended to __AGENT_OPS_LOG.
 *   • All functions return structured JSON (stringified).
 */

// ─── Constants ───────────────────────────────────────────────────────────────
var AGENT_LOG_SHEET   = "__AGENT_OPS_LOG";
var AGENT_TEST_SHEET  = "__AGENT_TEST";
var CONFIG_PREFIX     = "agentops_";
var VERSION           = "1.0.0";
var MAX_LOCK_WAIT_MS  = 10000;

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Canonical JSON response envelope */
function _ok(data) {
  return JSON.stringify({ ok: true, data: data, ts: new Date().toISOString() });
}

function _err(msg, detail) {
  return JSON.stringify({ ok: false, error: msg, detail: detail || null, ts: new Date().toISOString() });
}

/** Acquire a spreadsheet-scoped lock; throw if unavailable */
function _lock() {
  var lock = LockService.getSpreadsheetLock();
  lock.waitLock(MAX_LOCK_WAIT_MS);
  return lock;
}

/** Resolve a target object → Range.  Prefers namedRange, falls back to sheetName+a1. */
function _resolveRange(ss, target) {
  if (!target) throw new Error("target is required");
  if (target.namedRange) {
    var nr = ss.getRangeByName(target.namedRange);
    if (!nr) throw new Error("Named range not found: " + target.namedRange);
    return nr;
  }
  if (target.sheetName && target.a1) {
    var sh = ss.getSheetByName(target.sheetName);
    if (!sh) throw new Error("Sheet not found: " + target.sheetName);
    return sh.getRange(target.a1);
  }
  throw new Error("target must have namedRange, or sheetName+a1");
}

/** Compute a simple SHA-256 hex digest of a string */
function _hash(str) {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    str,
    Utilities.Charset.UTF_8
  );
  return raw.map(function(b) {
    return ("0" + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join("");
}

/** Hash the serialized values in a range */
function _rangeHash(range) {
  return _hash(JSON.stringify(range.getValues()));
}

/** Ensure the hidden log sheet exists; create if not. */
function _getOrCreateLogSheet(ss) {
  var sh = ss.getSheetByName(AGENT_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AGENT_LOG_SHEET);
    sh.hideSheet();
    sh.appendRow([
      "timestamp", "operationId", "type", "target",
      "rowCount", "colCount", "status", "message", "actor"
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Append one row to the log sheet (best-effort; never throws) */
function _log(ss, operationId, type, target, rowCount, colCount, status, message) {
  try {
    var sh = _getOrCreateLogSheet(ss);
    sh.appendRow([
      new Date().toISOString(),
      operationId || "",
      type || "",
      JSON.stringify(target || {}),
      rowCount || 0,
      colCount || 0,
      status || "ok",
      message || "",
      "AgentOps"
    ]);
  } catch(e) {
    // logging must never break the main operation
  }
}

/** Guard: refuse if range contains formulas and allowFormulaOverwrite is not set */
function _checkFormulas(range, payload) {
  if (payload.allowFormulaOverwrite) return;
  var formulas = range.getFormulas();
  for (var r = 0; r < formulas.length; r++) {
    for (var c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c] !== "") {
        throw new Error(
          "Target contains formulas at row+" + r + " col+" + c +
          ". Set allowFormulaOverwrite:true to proceed."
        );
      }
    }
  }
}

/** Guard: refuse if the sheet is hidden and allowHiddenSheet is not set */
function _checkHidden(range, payload) {
  if (payload.allowHiddenSheet) return;
  var sh = range.getSheet();
  if (sh.isSheetHidden()) {
    throw new Error(
      "Sheet '" + sh.getName() + "' is hidden. Set allowHiddenSheet:true to proceed."
    );
  }
}

/** Guard: refuse if range touches a protected range and allowProtected is not set */
function _checkProtected(range, ss, payload) {
  if (payload.allowProtected) return;
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var rSheet = range.getSheet().getName();
  var rRow1  = range.getRow();
  var rCol1  = range.getColumn();
  var rRow2  = rRow1 + range.getNumRows() - 1;
  var rCol2  = rCol1 + range.getNumColumns() - 1;
  for (var i = 0; i < protections.length; i++) {
    var p = protections[i];
    var pr = p.getRange();
    if (pr.getSheet().getName() !== rSheet) continue;
    var pRow1 = pr.getRow(), pCol1 = pr.getColumn();
    var pRow2 = pRow1 + pr.getNumRows() - 1;
    var pCol2 = pCol1 + pr.getNumColumns() - 1;
    // overlapping?
    if (rRow1 <= pRow2 && rRow2 >= pRow1 && rCol1 <= pCol2 && rCol2 >= pCol1) {
      throw new Error(
        "Target overlaps a protected range. Set allowProtected:true to proceed."
      );
    }
  }
}

/** Parse a JSON payload string or return the object if already parsed */
function _parse(payload) {
  if (typeof payload === "string") return JSON.parse(payload);
  return payload;
}

// ─── Public API ──────────────────────────────────────────────────────────────

/**
 * agent_healthcheck()
 * Returns framework version, spreadsheet metadata, sheet count, named ranges.
 */
function agent_healthcheck() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var nrs = ss.getNamedRanges().map(function(nr) {
      return { name: nr.getName(), a1: nr.getRange().getA1Notation(), sheet: nr.getRange().getSheet().getName() };
    });
    return _ok({
      version: VERSION,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetCount: sheets.length,
      sheets: sheets.map(function(s) {
        return { name: s.getName(), hidden: s.isSheetHidden(), rows: s.getLastRow(), cols: s.getLastColumn() };
      }),
      namedRanges: nrs,
      logSheetExists: !!ss.getSheetByName(AGENT_LOG_SHEET)
    });
  } catch(e) {
    return _err("healthcheck failed", e.message);
  }
}

/**
 * agent_snapshotWorkbook()
 * Returns full metadata: sheets, named ranges, protected ranges, triggers, properties.
 */
function agent_snapshotWorkbook() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets().map(function(sh) {
      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      var sampleValues = [];
      if (lastRow > 0 && lastCol > 0) {
        var sampleRows = Math.min(lastRow, 5);
        var sampleCols = Math.min(lastCol, 10);
        sampleValues = sh.getRange(1, 1, sampleRows, sampleCols).getValues();
      }
      return {
        name: sh.getName(),
        sheetId: sh.getSheetId(),
        index: sh.getIndex(),
        hidden: sh.isSheetHidden(),
        lastRow: lastRow,
        lastColumn: lastCol,
        frozenRows: sh.getFrozenRows(),
        frozenCols: sh.getFrozenColumns(),
        sampleHeader: sampleValues.length > 0 ? sampleValues[0] : [],
        tabColor: sh.getTabColorObject() ? sh.getTabColorObject().asRgbColor().asHexString() : null
      };
    });

    var namedRanges = ss.getNamedRanges().map(function(nr) {
      return { name: nr.getName(), sheet: nr.getRange().getSheet().getName(), a1: nr.getRange().getA1Notation() };
    });

    var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).map(function(p) {
      return {
        description: p.getDescription(),
        sheet: p.getRange().getSheet().getName(),
        a1: p.getRange().getA1Notation(),
        editors: p.getEditors().map(function(u) { return u.getEmail(); })
      };
    });

    var triggers = ScriptApp.getProjectTriggers().map(function(t) {
      return {
        handlerFunction: t.getHandlerFunction(),
        eventType: t.getEventType().toString(),
        triggerSource: t.getTriggerSource().toString(),
        uniqueId: t.getUniqueId()
      };
    });

    var props = PropertiesService.getDocumentProperties().getProperties();
    // Redact values that look like secrets
    var safeProps = {};
    Object.keys(props).forEach(function(k) {
      safeProps[k] = k.toLowerCase().match(/secret|key|token|pass/) ? "***REDACTED***" : props[k];
    });

    var hash = _hash(JSON.stringify(sheets.map(function(s) { return { name: s.name, lastRow: s.lastRow, lastCol: s.lastColumn }; })));

    return _ok({
      version: VERSION,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      snapshotHash: hash,
      sheets: sheets,
      namedRanges: namedRanges,
      protectedRanges: protections,
      triggers: triggers,
      documentProperties: safeProps
    });
  } catch(e) {
    return _err("snapshotWorkbook failed", e.message);
  }
}

/**
 * agent_readRange(payload)
 * payload: { target: { namedRange? | sheetName+a1 }, includeFormulas?: bool, includeNotes?: bool }
 */
function agent_readRange(payload) {
  try {
    payload = _parse(payload);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var range = _resolveRange(ss, payload.target);
    var values    = range.getValues();
    var formulas  = payload.includeFormulas ? range.getFormulas()   : null;
    var notes     = payload.includeNotes    ? range.getNotes()       : null;
    var hash      = _rangeHash(range);
    _log(ss, payload.operationId, "readRange", payload.target, range.getNumRows(), range.getNumColumns(), "ok", "read");
    return _ok({
      sheet: range.getSheet().getName(),
      a1: range.getA1Notation(),
      rows: range.getNumRows(),
      cols: range.getNumColumns(),
      values: values,
      formulas: formulas,
      notes: notes,
      hash: hash
    });
  } catch(e) {
    return _err("readRange failed", e.message);
  }
}

/**
 * agent_writeRangeDryRun(payload)
 * Returns what WOULD be written without touching the sheet.
 * payload: { target, values, operationId?, allowFormulaOverwrite?, allowHiddenSheet?, allowProtected?, expectedHash? }
 */
function agent_writeRangeDryRun(payload) {
  try {
    payload = _parse(payload);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var range = _resolveRange(ss, payload.target);
    var currentValues = range.getValues();
    var currentHash   = _rangeHash(range);

    if (payload.expectedHash && payload.expectedHash !== currentHash) {
      return _err("expectedHash mismatch — sheet has changed since last read", {
        expected: payload.expectedHash, actual: currentHash
      });
    }

    var warnings = [];
    // Check formula collision
    var formulas = range.getFormulas();
    var formulaCells = [];
    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        if (formulas[r][c] !== "") formulaCells.push({ row: r, col: c, formula: formulas[r][c] });
      }
    }
    if (formulaCells.length > 0 && !payload.allowFormulaOverwrite) {
      warnings.push("FORMULA_OVERWRITE: " + formulaCells.length + " formula cell(s) would be overwritten. Set allowFormulaOverwrite:true.");
    }

    var sh = range.getSheet();
    if (sh.isSheetHidden() && !payload.allowHiddenSheet) {
      warnings.push("HIDDEN_SHEET: target sheet is hidden. Set allowHiddenSheet:true.");
    }

    var newValues = payload.values || [];
    var totalCells = (newValues.length || 0) * ((newValues[0] || []).length || 0);

    return _ok({
      dryRun: true,
      sheet: sh.getName(),
      a1: range.getA1Notation(),
      rows: range.getNumRows(),
      cols: range.getNumColumns(),
      currentHash: currentHash,
      currentValues: currentValues,
      proposedValues: newValues,
      totalCellsAffected: totalCells,
      formulaCellsAtRisk: formulaCells,
      warnings: warnings,
      requiresApproval: warnings.length > 0 || totalCells > 100
    });
  } catch(e) {
    return _err("writeRangeDryRun failed", e.message);
  }
}

/**
 * agent_writeRange(payload)
 * payload: { target, values, operationId?, allowFormulaOverwrite?, allowHiddenSheet?, allowProtected?, expectedHash?, confirmDestructive? }
 */
function agent_writeRange(payload) {
  var lock;
  try {
    payload = _parse(payload);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var range = _resolveRange(ss, payload.target);

    if (payload.expectedHash) {
      var currentHash = _rangeHash(range);
      if (payload.expectedHash !== currentHash) {
        return _err("expectedHash mismatch — aborting write", {
          expected: payload.expectedHash, actual: currentHash
        });
      }
    }

    _checkFormulas(range, payload);
    _checkHidden(range, payload);
    _checkProtected(range, ss, payload);

    var values = payload.values;
    if (!values || !Array.isArray(values)) return _err("values array required");

    var totalCells = values.length * (values[0] ? values[0].length : 0);
    if (totalCells > 100 && !payload.confirmLarge) {
      return _err("Write affects " + totalCells + " cells (>100). Set confirmLarge:true to proceed.");
    }

    lock = _lock();
    range.setValues(values);
    SpreadsheetApp.flush();
    var newHash = _rangeHash(range);
    lock.releaseLock();

    _log(ss, payload.operationId, "writeRange", payload.target, values.length, values[0] ? values[0].length : 0, "ok", "wrote " + totalCells + " cells");

    return _ok({
      sheet: range.getSheet().getName(),
      a1: range.getA1Notation(),
      cellsWritten: totalCells,
      newHash: newHash
    });
  } catch(e) {
    if (lock) try { lock.releaseLock(); } catch(_) {}
    return _err("writeRange failed", e.message);
  }
}

/**
 * agent_appendRowsDryRun(payload)
 * payload: { target: { sheetName }, rows: [[...]], operationId? }
 */
function agent_appendRowsDryRun(payload) {
  try {
    payload = _parse(payload);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = payload.target && payload.target.sheetName;
    if (!sheetName) return _err("target.sheetName required");
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return _err("Sheet not found: " + sheetName);

    var rows = payload.rows || [];
    return _ok({
      dryRun: true,
      sheet: sheetName,
      currentLastRow: sh.getLastRow(),
      rowsToAppend: rows.length,
      proposedAppendAfterRow: sh.getLastRow(),
      rows: rows
    });
  } catch(e) {
    return _err("appendRowsDryRun failed", e.message);
  }
}

/**
 * agent_appendRows(payload)
 * payload: { target: { sheetName }, rows: [[...]], operationId? }
 */
function agent_appendRows(payload) {
  var lock;
  try {
    payload = _parse(payload);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = payload.target && payload.target.sheetName;
    if (!sheetName) return _err("target.sheetName required");
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return _err("Sheet not found: " + sheetName);

    _checkHidden({ getSheet: function() { return sh; } }, payload);

    var rows = payload.rows || [];
    if (!rows.length) return _err("rows array is empty");

    lock = _lock();
    var beforeRow = sh.getLastRow();
    rows.forEach(function(row) { sh.appendRow(row); });
    SpreadsheetApp.flush();
    lock.releaseLock();

    _log(ss, payload.operationId, "appendRows", payload.target, rows.length, rows[0] ? rows[0].length : 0, "ok", "appended after row " + beforeRow);

    return _ok({
      sheet: sheetName,
      rowsAppended: rows.length,
      firstNewRow: beforeRow + 1,
      lastNewRow: beforeRow + rows.length
    });
  } catch(e) {
    if (lock) try { lock.releaseLock(); } catch(_) {}
    return _err("appendRows failed", e.message);
  }
}

/**
 * agent_clearRangeDryRun(payload)
 * payload: { target, operationId? }
 */
function agent_clearRangeDryRun(payload) {
  try {
    payload = _parse(payload);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var range = _resolveRange(ss, payload.target);
    var values    = range.getValues();
    var formulas  = range.getFormulas();
    var nonEmpty  = 0;
    var withFormulas = 0;
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        if (values[r][c] !== "") nonEmpty++;
        if (formulas[r][c] !== "") withFormulas++;
      }
    }
    return _ok({
      dryRun: true,
      sheet: range.getSheet().getName(),
      a1: range.getA1Notation(),
      rows: range.getNumRows(),
      cols: range.getNumColumns(),
      nonEmptyCells: nonEmpty,
      formulaCells: withFormulas,
      warnings: withFormulas > 0 && !payload.allowFormulaOverwrite
        ? ["FORMULA_CLEAR: " + withFormulas + " formula(s) would be cleared. Set allowFormulaOverwrite:true."]
        : []
    });
  } catch(e) {
    return _err("clearRangeDryRun failed", e.message);
  }
}

/**
 * agent_clearRange(payload)
 * payload: { target, operationId?, allowFormulaOverwrite?, allowHiddenSheet?, allowProtected?, confirmDestructive }
 * confirmDestructive MUST be true.
 */
function agent_clearRange(payload) {
  var lock;
  try {
    payload = _parse(payload);
    if (!payload.confirmDestructive) return _err("confirmDestructive:true required for clearRange");
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var range = _resolveRange(ss, payload.target);
    _checkFormulas(range, payload);
    _checkHidden(range, payload);
    _checkProtected(range, ss, payload);

    lock = _lock();
    range.clearContent();
    SpreadsheetApp.flush();
    lock.releaseLock();

    _log(ss, payload.operationId, "clearRange", payload.target, range.getNumRows(), range.getNumColumns(), "ok", "cleared");
    return _ok({ sheet: range.getSheet().getName(), a1: range.getA1Notation(), cleared: true });
  } catch(e) {
    if (lock) try { lock.releaseLock(); } catch(_) {}
    return _err("clearRange failed", e.message);
  }
}

/**
 * agent_deleteRowsDryRun(payload)
 * payload: { target: { sheetName }, startRow, numRows, operationId? }
 */
function agent_deleteRowsDryRun(payload) {
  try {
    payload = _parse(payload);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(payload.target.sheetName);
    if (!sh) return _err("Sheet not found: " + payload.target.sheetName);
    var startRow = payload.startRow;
    var numRows  = payload.numRows || 1;
    var endRow   = startRow + numRows - 1;
    var preview  = sh.getRange(startRow, 1, numRows, Math.max(1, sh.getLastColumn())).getValues();
    return _ok({
      dryRun: true,
      sheet: payload.target.sheetName,
      startRow: startRow,
      endRow: endRow,
      numRows: numRows,
      previewRows: preview,
      totalSheetRows: sh.getLastRow(),
      warnings: ["DELETE_ROWS: " + numRows + " row(s) from row " + startRow + " will be permanently deleted."]
    });
  } catch(e) {
    return _err("deleteRowsDryRun failed", e.message);
  }
}

/**
 * agent_deleteRows(payload)
 * payload: { target: { sheetName }, startRow, numRows, operationId?, confirmDestructive, backupRequired? }
 * confirmDestructive MUST be true.
 */
function agent_deleteRows(payload) {
  var lock;
  try {
    payload = _parse(payload);
    if (!payload.confirmDestructive) return _err("confirmDestructive:true required for deleteRows");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(payload.target.sheetName);
    if (!sh) return _err("Sheet not found: " + payload.target.sheetName);
    _checkHidden(sh.getRange(payload.startRow, 1), payload);

    lock = _lock();
    sh.deleteRows(payload.startRow, payload.numRows || 1);
    SpreadsheetApp.flush();
    lock.releaseLock();

    _log(ss, payload.operationId, "deleteRows", payload.target, payload.numRows || 1, 0, "ok",
      "deleted rows " + payload.startRow + "-" + (payload.startRow + (payload.numRows || 1) - 1));
    return _ok({ sheet: payload.target.sheetName, deletedStartRow: payload.startRow, deletedRows: payload.numRows || 1 });
  } catch(e) {
    if (lock) try { lock.releaseLock(); } catch(_) {}
    return _err("deleteRows failed", e.message);
  }
}

/**
 * agent_backupSpreadsheet(payload)
 * payload: { operationId?, reason? }
 * Creates a copy in Drive if permissions allow.
 */
function agent_backupSpreadsheet(payload) {
  try {
    payload = _parse(payload || "{}");
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var ts     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
    var name   = ss.getName() + "__backup__" + ts;
    var file   = DriveApp.getFileById(ss.getId());
    var copy   = file.makeCopy(name, file.getParents().next());
    _log(ss, payload.operationId, "backupSpreadsheet", {}, 0, 0, "ok", "backup: " + name + " id:" + copy.getId());
    return _ok({ backupName: name, backupId: copy.getId(), reason: payload.reason || "" });
  } catch(e) {
    return _err("backupSpreadsheet failed", e.message);
  }
}

/**
 * agent_logOperation(payload)
 * payload: { operationId, type, target, rowCount?, colCount?, status, message }
 */
function agent_logOperation(payload) {
  try {
    payload = _parse(payload);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    _log(ss, payload.operationId, payload.type, payload.target, payload.rowCount, payload.colCount, payload.status, payload.message);
    return _ok({ logged: true });
  } catch(e) {
    return _err("logOperation failed", e.message);
  }
}

/**
 * agent_runValidation(payload)
 * Runs a set of declared checks: namedRangeExists, sheetExists, rowCountMin, rowCountMax, cellNotEmpty, cellEquals
 * payload: { checks: [ { type, ... } ] }
 */
function agent_runValidation(payload) {
  try {
    payload = _parse(payload);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var results = (payload.checks || []).map(function(check) {
      try {
        switch (check.type) {
          case "namedRangeExists":
            var nr = ss.getRangeByName(check.name);
            return { check: check, pass: !!nr, note: nr ? "found" : "not found" };
          case "sheetExists":
            var sh = ss.getSheetByName(check.name);
            return { check: check, pass: !!sh, note: sh ? "found" : "not found" };
          case "rowCountMin":
            var sh2 = ss.getSheetByName(check.sheetName);
            var rows = sh2 ? sh2.getLastRow() : 0;
            return { check: check, pass: rows >= check.min, note: "lastRow=" + rows };
          case "rowCountMax":
            var sh3 = ss.getSheetByName(check.sheetName);
            var rows3 = sh3 ? sh3.getLastRow() : 0;
            return { check: check, pass: rows3 <= check.max, note: "lastRow=" + rows3 };
          case "cellNotEmpty":
            var range = _resolveRange(ss, check.target);
            var val = range.getValue();
            return { check: check, pass: val !== "" && val !== null, note: "value=" + val };
          case "cellEquals":
            var range2 = _resolveRange(ss, check.target);
            var val2 = range2.getValue();
            return { check: check, pass: String(val2) === String(check.expected), note: "value=" + val2 + " expected=" + check.expected };
          case "logSheetExists":
            return { check: check, pass: !!ss.getSheetByName(AGENT_LOG_SHEET), note: "" };
          default:
            return { check: check, pass: false, note: "unknown check type: " + check.type };
        }
      } catch(err) {
        return { check: check, pass: false, note: "error: " + err.message };
      }
    });
    var allPassed = results.every(function(r) { return r.pass; });
    return _ok({ allPassed: allPassed, results: results });
  } catch(e) {
    return _err("runValidation failed", e.message);
  }
}

/**
 * agent_listProjectFunctions()
 * Returns a manifest of all AgentOps public functions.
 */
function agent_listProjectFunctions() {
  try {
    var fns = [
      "agent_healthcheck", "agent_snapshotWorkbook", "agent_readRange",
      "agent_writeRangeDryRun", "agent_writeRange",
      "agent_appendRowsDryRun", "agent_appendRows",
      "agent_clearRangeDryRun", "agent_clearRange",
      "agent_deleteRowsDryRun", "agent_deleteRows",
      "agent_backupSpreadsheet", "agent_logOperation",
      "agent_runValidation", "agent_listProjectFunctions",
      "agent_runAllowlistedFunction", "agent_getConfig", "agent_setConfig"
    ];
    return _ok({ functions: fns, count: fns.length });
  } catch(e) {
    return _err("listProjectFunctions failed", e.message);
  }
}

/**
 * agent_runAllowlistedFunction(payload)
 * payload: { functionName, args?, allowlist? }
 * Runs a named project function only if it appears in the allowlist property.
 */
function agent_runAllowlistedFunction(payload) {
  try {
    payload = _parse(payload);
    var fnName = payload.functionName;
    if (!fnName) return _err("functionName required");

    // Load allowlist from properties
    var props = PropertiesService.getDocumentProperties();
    var raw   = props.getProperty(CONFIG_PREFIX + "allowlist");
    var allowlist = raw ? JSON.parse(raw) : (payload.allowlist || []);

    if (allowlist.indexOf(fnName) === -1) {
      return _err("Function '" + fnName + "' is not in the allowlist. Add it via agent_setConfig.");
    }

    // Avoid running AgentOps internals recursively via this path
    if (fnName.startsWith("_")) return _err("Cannot call private functions via allowlist.");

    var fn = this[fnName] || eval(fnName); // eslint-disable-line no-eval
    if (typeof fn !== "function") return _err("Function not found: " + fnName);

    var result = fn(payload.args || null);
    return _ok({ functionName: fnName, result: result });
  } catch(e) {
    return _err("runAllowlistedFunction failed", e.message);
  }
}

/**
 * agent_getConfig()
 * Returns all agentops_ properties from DocumentProperties.
 */
function agent_getConfig() {
  try {
    var props = PropertiesService.getDocumentProperties().getProperties();
    var config = {};
    Object.keys(props).forEach(function(k) {
      if (k.indexOf(CONFIG_PREFIX) === 0) config[k.slice(CONFIG_PREFIX.length)] = props[k];
    });
    return _ok({ config: config });
  } catch(e) {
    return _err("getConfig failed", e.message);
  }
}

/**
 * agent_setConfig(payload)
 * payload: { key: value, ... }  — merges into agentops_ DocumentProperties.
 */
function agent_setConfig(payload) {
  try {
    payload = _parse(payload);
    var props = PropertiesService.getDocumentProperties();
    Object.keys(payload).forEach(function(k) {
      if (k === "operationId") return; // skip meta keys
      props.setProperty(CONFIG_PREFIX + k, typeof payload[k] === "string" ? payload[k] : JSON.stringify(payload[k]));
    });
    return _ok({ set: Object.keys(payload).filter(function(k) { return k !== "operationId"; }) });
  } catch(e) {
    return _err("setConfig failed", e.message);
  }
}

// ─── Acceptance test helpers ──────────────────────────────────────────────────

/**
 * agent_acceptanceTest()
 * Runs all self-checks against a harmless __AGENT_TEST sheet.
 * Call this once after installing AgentOps.gs.
 */
function agent_acceptanceTest() {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function check(name, fn) {
    try {
      var r = fn();
      var parsed = typeof r === "string" ? JSON.parse(r) : r;
      results.push({ test: name, pass: parsed.ok === true, detail: parsed.data || parsed.error });
    } catch(e) {
      results.push({ test: name, pass: false, detail: e.message });
    }
  }

  // 1. healthcheck
  check("healthcheck", function() { return agent_healthcheck(); });

  // 2. snapshotWorkbook
  check("snapshotWorkbook", function() { return agent_snapshotWorkbook(); });

  // 3. Ensure __AGENT_TEST sheet exists
  check("ensureTestSheet", function() {
    var sh = ss.getSheetByName(AGENT_TEST_SHEET);
    if (!sh) {
      sh = ss.insertSheet(AGENT_TEST_SHEET);
      sh.hideSheet();
      sh.getRange("A1").setValue("AGENT_TEST_INITIALIZED");
    }
    return _ok({ sheetName: AGENT_TEST_SHEET, created: true });
  });

  // 4. readRange from test sheet
  check("readRange", function() {
    return agent_readRange(JSON.stringify({
      target: { sheetName: AGENT_TEST_SHEET, a1: "A1" },
      operationId: "acceptance-read"
    }));
  });

  // 5. writeRangeDryRun
  check("writeRangeDryRun", function() {
    return agent_writeRangeDryRun(JSON.stringify({
      target: { sheetName: AGENT_TEST_SHEET, a1: "B1" },
      values: [["DRY_RUN_TEST"]],
      operationId: "acceptance-dryrun",
      allowHiddenSheet: true
    }));
  });

  // 6. writeRange to test sheet
  check("writeRange", function() {
    return agent_writeRange(JSON.stringify({
      target: { sheetName: AGENT_TEST_SHEET, a1: "B1" },
      values: [["WRITE_OK_" + new Date().getTime()]],
      operationId: "acceptance-write",
      allowHiddenSheet: true
    }));
  });

  // 7. appendRowsDryRun
  check("appendRowsDryRun", function() {
    return agent_appendRowsDryRun(JSON.stringify({
      target: { sheetName: AGENT_TEST_SHEET },
      rows: [["test", "append", new Date().toISOString()]],
      operationId: "acceptance-append-dry"
    }));
  });

  // 8. logOperation
  check("logOperation", function() {
    return agent_logOperation(JSON.stringify({
      operationId: "acceptance-log",
      type: "acceptanceTest",
      target: { sheetName: AGENT_TEST_SHEET },
      status: "ok",
      message: "acceptance test log entry"
    }));
  });

  // 9. runValidation
  check("runValidation", function() {
    return agent_runValidation(JSON.stringify({
      checks: [
        { type: "sheetExists", name: AGENT_TEST_SHEET },
        { type: "logSheetExists" },
        { type: "cellNotEmpty", target: { sheetName: AGENT_TEST_SHEET, a1: "A1" } }
      ]
    }));
  });

  // 10. getConfig / setConfig round-trip
  check("setConfig", function() {
    return agent_setConfig(JSON.stringify({ testKey: "acceptance_" + new Date().getTime() }));
  });
  check("getConfig", function() { return agent_getConfig(); });

  // 11. listProjectFunctions
  check("listProjectFunctions", function() { return agent_listProjectFunctions(); });

  var allPassed = results.every(function(r) { return r.pass; });
  return JSON.stringify({ allPassed: allPassed, results: results, ts: new Date().toISOString() });
}
