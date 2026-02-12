/**
 * Result reporting - writes step results, checkpoints, and execution time.
 * Equivalent to EM.ResultReporting Java class.
 */

import * as XLSX from 'xlsx';
import path from 'path';
import fs from 'fs';

let rowPrint = 1;
let rowPrintChk = 1;
let rowPrintTime = 1;

export interface StepReportRow {
  testScenarioRow: number;
  moduleName: string;
  functionName: string;
  result: string;
  comments: string;
}

export interface CheckPointReportRow {
  testScenarioRow: number;
  functionName: string;
  result: string;
  comments: string;
}

export interface TimeReportRow {
  testScenarioRow: number;
  testScenario: string;
  startTime: string;
  endTime: string;
  elapsedSeconds: number;
}

let stepReportSheet: XLSX.WorkSheet;
let checkPointSheet: XLSX.WorkSheet;
let timeSheet: XLSX.WorkSheet;
let workbook: XLSX.WorkBook;
let reportFilePath: string;

/**
 * Initialize report workbook and sheets (create or load).
 */
export function initReport(filePath: string): void {
  reportFilePath = path.isAbsolute(filePath) ? filePath : path.resolve(process.cwd(), filePath);
  const dir = path.dirname(reportFilePath);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

  if (fs.existsSync(reportFilePath)) {
    workbook = XLSX.readFile(reportFilePath, { type: 'file' });
  } else {
    workbook = XLSX.utils.book_new();
  }

  stepReportSheet = workbook.Sheets['Step Report'] ?? XLSX.utils.aoa_to_sheet([
    ['TestScenarioRow', 'Module_Name', 'Function_Name', 'Result', 'Comments'],
  ]);
  checkPointSheet = workbook.Sheets['CheckPoint'] ?? XLSX.utils.aoa_to_sheet([
    ['TestScenarioRow', 'Function_Name', 'Result', 'Comments'],
  ]);
  timeSheet = workbook.Sheets['Time'] ?? XLSX.utils.aoa_to_sheet([
    ['TestScenarioRow', 'Function_Name', 'Start_Time', 'End_Time', 'Elapsed_Time_in_seconds'],
  ]);

  workbook.Sheets['Step Report'] = stepReportSheet;
  workbook.Sheets['CheckPoint'] = checkPointSheet;
  workbook.Sheets['Time'] = timeSheet;
  if (!workbook.SheetNames?.length) workbook.SheetNames = ['Step Report', 'CheckPoint', 'Time'];

  rowPrint = 1;
  rowPrintChk = 1;
  rowPrintTime = 1;
}

function writeReport(): void {
  XLSX.writeFile(workbook, reportFilePath, { bookType: 'xlsx' });
}

function getNextRow(sheet: XLSX.WorkSheet): number {
  const ref = sheet['!ref'];
  if (!ref) return 1;
  const range = XLSX.utils.decode_range(ref);
  return range.e.r + 1;
}

/**
 * Append step result.
 */
export function printResult(
  moduleName: string,
  modFunc: string,
  result: string,
  comments: string,
  testScenarioRow: number
): void {
  const row = [testScenarioRow + 1, moduleName, modFunc, result, comments];
  const newRow = getNextRow(stepReportSheet);
  row.forEach((val, c) => {
    const ref = XLSX.utils.encode_cell({ r: newRow, c });
    stepReportSheet[ref] = { t: 's', v: String(val) };
  });
  stepReportSheet['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: newRow, c: 4 },
  });
  rowPrint++;
  writeReport();
}

/**
 * Append checkpoint result.
 */
export function printResultCheckpoint(
  moduleName: string,
  result: string,
  comments: string,
  testScenarioRow: number
): void {
  const row = [testScenarioRow + 1, moduleName, result, comments];
  const newRow = getNextRow(checkPointSheet);
  row.forEach((val, c) => {
    const ref = XLSX.utils.encode_cell({ r: newRow, c });
    checkPointSheet[ref] = { t: 's', v: String(val) };
  });
  checkPointSheet['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: newRow, c: 3 },
  });
  rowPrintChk++;
  writeReport();
}

/**
 * Append execution time row.
 */
export function printResultTime(
  testScenario: string,
  startTime: string,
  endTime: string,
  elapsedSeconds: number,
  testScenarioRow: number
): void {
  const row = [testScenarioRow + 1, testScenario, startTime, endTime, elapsedSeconds];
  const newRow = getNextRow(timeSheet);
  row.forEach((val, c) => {
    const ref = XLSX.utils.encode_cell({ r: newRow, c });
    timeSheet[ref] = c === 4 ? { t: 'n', v: Number(val) } : { t: 's', v: String(val) };
  });
  timeSheet['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: newRow, c: 4 },
  });
  rowPrintTime++;
  writeReport();
}

export function getReportFilePath(): string {
  return reportFilePath;
}
