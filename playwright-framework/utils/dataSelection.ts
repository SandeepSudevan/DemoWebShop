/**
 * Data selection utilities - reads test data from Excel (or JSON fallback).
 * Equivalent to EM.Data_Selection Java class.
 */

import * as XLSX from 'xlsx';
import path from 'path';

let Parameter1: string | null = null;

export type LocatorType = 'name' | 'id' | 'xpath' | 'css' | 'linktext' | 'title' | 'value' | 'class' | 'attribute/name';

export interface TestStepRow {
  module: string;
  modFunc: string;
  objectType: string;
  propName: string;
  propValue: string;
  addPropName?: string;
  addPropValue?: string;
  parameter1?: string;
  [key: string]: string | undefined;
}

/**
 * Get cell value as string from sheet (handles numeric, string, formula).
 */
export function cellType(
  column: number,
  row: number,
  sheet: XLSX.WorkSheet,
  workbook: XLSX.WorkBook
): string | null {
  try {
    const cellRef = XLSX.utils.encode_cell({ r: row, c: column });
    const cell = sheet[cellRef];
    if (!cell) return null;

    if (cell.t === 'n') {
      let val = String(cell.v);
      if (val.endsWith('.0')) val = val.replace('.0', '');
      return val;
    }
    if (cell.t === 's') return cell.v as string;
    if (cell.t === 'f') {
      const resolved = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as unknown[][];
      const rowData = resolved[row];
      const raw = rowData?.[column];
      return raw != null ? String(raw) : null;
    }
    return cell.v != null ? String(cell.v) : null;
  } catch {
    return null;
  }
}

/**
 * Resolve parameter for a step (Data1..Data36 style or direct cell).
 */
export function getParameter(
  dataSelection: string,
  row: number,
  col: number,
  sheet: XLSX.WorkSheet,
  workbook: XLSX.WorkBook
): string | null {
  const dataColMap: Record<string, number> = {};
  for (let i = 1; i <= 36; i++) dataColMap[`Data${i}`] = 77 + i;
  const colIndex = dataColMap[dataSelection || 'Data1'] ?? 78;
  return cellType(colIndex, row, sheet, workbook);
}

/**
 * ModDataSel - get parameter based on module (default: cell at column 78).
 */
export function modDataSel(
  module: string,
  sheet: XLSX.WorkSheet,
  row: number,
  col: number,
  workbook: XLSX.WorkBook,
  dataSelection: string
): string | null {
  if (module === 'DataSelection') {
    return getParameter(dataSelection || 'Data1', row, col, sheet, workbook);
  }
  return cellType(78, row, sheet, workbook);
}

/**
 * Parse a test case row into Object1, ProName1, ProValue1, AddProname, AddProValue, Parameter1.
 */
export function parseStepRow(
  sheet: XLSX.WorkSheet,
  workbook: XLSX.WorkBook,
  rowIndex: number,
  startCol: number = 2
): {
  object1: string | null;
  propName: string | null;
  propValue: string | null;
  addPropName: string | null;
  addPropValue: string | null;
  parameter1: string | null;
} {
  const object1 = cellType(startCol, rowIndex, sheet, workbook);
  const propName = cellType(startCol + 1, rowIndex, sheet, workbook);
  let propValue = cellType(startCol + 2, rowIndex, sheet, workbook);
  if (propValue && propValue.includes('.') && propValue.includes('E')) {
    propValue = propValue.replace('.', '').substring(0, 12);
  }
  const addPropName = cellType(startCol + 3, rowIndex, sheet, workbook);
  const addPropValue = cellType(startCol + 4, rowIndex, sheet, workbook);
  const parameter1 = cellType(startCol + 5, rowIndex, sheet, workbook);
  return {
    object1,
    propName: propName?.toLowerCase() === 'attribute/name' ? 'name' : propName ?? null,
    propValue: propValue ?? null,
    addPropName: addPropName ?? null,
    addPropValue: addPropValue ?? null,
    parameter1: parameter1 ?? null,
  };
}

/**
 * Load workbook from path (supports .xls via xlsx).
 */
export function loadWorkbook(filePath: string): XLSX.WorkBook {
  const fullPath = path.isAbsolute(filePath) ? filePath : path.resolve(process.cwd(), filePath);
  const workbook = XLSX.readFile(fullPath, { type: 'file', cellDates: true });
  return workbook;
}

/**
 * Get sheet by name from workbook.
 */
export function getSheet(workbook: XLSX.WorkBook, sheetName: string): XLSX.WorkSheet | undefined {
  return workbook.Sheets[sheetName];
}
