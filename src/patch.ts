import type { CellValue, Patch, PatchOperation, RangeTarget } from './types';

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}
function text(value: unknown): value is string { return typeof value === 'string' && value.length > 0; }
function keys(value: Record<string, unknown>, allowed: string[], label: string) {
  for (const key of Object.keys(value)) if (!allowed.includes(key)) throw new Error(`${label} has unsupported property: ${key}`);
}
function matrix(value: unknown): value is CellValue[][] {
  return Array.isArray(value) && value.every(row => Array.isArray(row) && row.every(cell => cell === null || typeof cell === 'string' || typeof cell === 'boolean' || (typeof cell === 'number' && Number.isFinite(cell))));
}
export function parsePatch(value: unknown, project: string, applying = false): Patch {
  if (!record(value)) throw new Error('Patch must be an object');
  for (const key of ['operationId', 'project', 'reason']) if (!text(value[key])) throw new Error('Patch missing: ' + key);
  if (value.project !== project) throw new Error(`Patch project '${value.project}' does not match --project '${project}'`);
  if (!Array.isArray(value.operations) || !value.operations.length) throw new Error('Patch has no operations');
  for (const key of ['requiresApproval', 'backupRequired']) if (value[key] !== undefined && typeof value[key] !== 'boolean') throw new Error(`${key} must be boolean`);
  if (value.expectedWorkbookHash !== undefined) throw new Error('expectedWorkbookHash is unsupported by the REST patch path; use expectedHash per range');
  keys(value, ['operationId', 'project', 'reason', 'requiresApproval', 'backupRequired', 'operations'], 'Patch');
  const operations = value.operations.map((op: unknown, index: number): PatchOperation => {
    if (!record(op) || !['setValues', 'appendRows', 'clearRange'].includes(String(op.type))) throw new Error(`Unsupported operation at [${index}]: ${record(op) ? String(op.type) : 'invalid operation'}`);
    if (!record(op.target)) throw new Error(`Operation [${index}] needs a target object`);
    keys(op.target, ['namedRange', 'sheetName', 'a1'], `Operation [${index}] target`);
    if (op.target.namedRange !== undefined && (op.target.sheetName !== undefined || op.target.a1 !== undefined)) throw new Error(`Operation [${index}] target must choose namedRange or sheetName, not both`);
    if (op.type === 'appendRows') {
      if (!text(op.target.sheetName) || op.target.namedRange !== undefined) throw new Error(`Operation [${index}] appendRows requires sheetName and does not support namedRange`);
      if (op.target.a1 !== undefined && !text(op.target.a1)) throw new Error(`Operation [${index}] a1 must be a string`);
    } else if (!(text(op.target.namedRange) || (text(op.target.sheetName) && text(op.target.a1)))) throw new Error(`Operation [${index}] needs namedRange or sheetName + a1`);
    for (const key of ['confirmLarge', 'confirmDestructive']) if (op[key] !== undefined && typeof op[key] !== 'boolean') throw new Error(`Operation [${index}] ${key} must be boolean`);
    for (const key of ['allowFormulaOverwrite', 'allowHiddenSheet', 'allowProtected']) if (op[key] !== undefined) throw new Error(`Operation [${index}] ${key} is unsupported by the REST patch path`);
    keys(op, ['type', 'target', 'values', 'expectedHash', 'confirmLarge', 'confirmDestructive'], `Operation [${index}]`);
    if (op.expectedHash !== undefined && !text(op.expectedHash)) throw new Error(`Operation [${index}] expectedHash must be a string`);
    if (op.expectedHash !== undefined && op.type !== 'setValues') throw new Error(`Operation [${index}] expectedHash is supported only for setValues`);
    if (op.type === 'clearRange' && op.values !== undefined) throw new Error(`Operation [${index}] clearRange does not accept values`);
    if (op.type !== 'clearRange' && !matrix(op.values)) throw new Error(`Operation [${index}] values must be a matrix of cell values`);
    if (applying && op.type === 'clearRange' && op.confirmDestructive !== true) throw new Error(`Operation [${index}] confirmDestructive:true required for clearRange`);
    if (op.type === 'appendRows' && !text(op.target.sheetName)) throw new Error(`Operation [${index}] appendRows requires target.sheetName`);
    const cells = matrix(op.values) ? op.values.reduce((sum, row) => sum + row.length, 0) : 0;
    if (applying && cells > 100 && op.confirmLarge !== true) throw new Error(`Operation [${index}] affects ${cells} cells (>100). Set confirmLarge:true.`);
    return {
      type: op.type as PatchOperation['type'], target: op.target as RangeTarget,
      values: op.values as CellValue[][] | undefined, expectedHash: op.expectedHash as string | undefined,
      confirmLarge: op.confirmLarge === true, confirmDestructive: op.confirmDestructive === true,
    };
  });
  return {operationId: value.operationId as string, project, reason: value.reason as string,
    requiresApproval: value.requiresApproval as boolean | undefined, backupRequired: value.backupRequired as boolean | undefined, operations};
}
