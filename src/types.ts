/** Unvalidated JSON and Google transport responses in the legacy API surface. */
export type Dynamic = any;
export type JsonObject = Record<string, Dynamic>;
export type CellValue = string | number | boolean | null;
export interface RangeTarget { namedRange?: string; sheetName?: string; a1?: string; }
export interface WriteOptions { expectedHash?: string; confirmLarge?: boolean; }
export interface BackupOptions { folderId?: string; retain?: number; }
/** The parser accepts boolean switches and strings. Commands validate valued flags. */
export interface CliArgs extends JsonObject { _: string[]; }
export interface PatchOperation {
  type: 'setValues' | 'appendRows' | 'clearRange';
  target: RangeTarget;
  values?: CellValue[][];
  expectedHash?: string;
  confirmLarge?: boolean;
  confirmDestructive?: boolean;
}
export interface Patch {
  operationId: string;
  project: string;
  reason: string;
  requiresApproval?: boolean;
  backupRequired?: boolean;
  operations: PatchOperation[];
}

export interface GridRange { sheetId: number; startColumnIndex?: number; startRowIndex?: number; endColumnIndex?: number; endRowIndex?: number; }
export interface Color { red?: number; green?: number; blue?: number; }
export interface SheetMetadata {
  properties: {title: string; sheetId: number; index?: number; hidden?: boolean; gridProperties?: {rowCount: number; columnCount: number; frozenRowCount?: number; frozenColumnCount?: number}; tabColorStyle?: {rgbColor?: Color};};
  protectedRanges?: Array<{description?: string; range?: GridRange; editors?: {users?: string[]}}>;
}
export interface NamedRange { name: string; namedRangeId?: string; range?: GridRange; }
export interface SpreadsheetMetadata { spreadsheetId: string; properties?: {title: string}; sheets?: SheetMetadata[]; namedRanges?: NamedRange[]; }
export interface SheetSnapshot { name: string; sheetId: number; index?: number; hidden: boolean; lastRow: number; lastColumn: number; frozenRows: number; frozenCols: number; sampleHeader: CellValue[]; tabColor?: Color | null; }
export interface WorkbookSnapshot {snapshotHash: string; sheets: SheetSnapshot[]; namedRanges: Array<{name: string}>; }
export interface VisualOptions { sheetTitle?: string; range?: string; theme?: JsonObject; wrap?: boolean; resizeRows?: boolean; maxWidth?: number; minWidth?: number; startColumn?: number; endColumn?: number; headerRows?: string | number; freeze?: boolean; banding?: boolean; borders?: boolean; autofit?: boolean; inferNumberFormats?: boolean; type?: string; title?: string; newSheet?: boolean; targetSheetTitle?: string; anchor?: string; dataRange?: string; rule?: string; color?: string; value?: number; subtitle?: string; width?: string | number; sourceRange?: string; targetCol?: string; negColor?: string; highColor?: string; lowColor?: string; firstColor?: string; lastColor?: string; includeText?: boolean; }
