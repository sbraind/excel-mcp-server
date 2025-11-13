import ExcelJS from 'exceljs';
import { promises as fs } from 'fs';
import { ERROR_MESSAGES } from '../constants.js';
import type { CellValue } from 'exceljs';

export async function loadWorkbook(filePath: string): Promise<ExcelJS.Workbook> {
  try {
    await fs.access(filePath);
  } catch {
    throw new Error(`${ERROR_MESSAGES.FILE_NOT_FOUND}: ${filePath}`);
  }

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filePath);
    return workbook;
  } catch (error) {
    throw new Error(`${ERROR_MESSAGES.READ_ERROR}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

export function getSheet(workbook: ExcelJS.Workbook, sheetName: string): ExcelJS.Worksheet {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`${ERROR_MESSAGES.SHEET_NOT_FOUND}: ${sheetName}`);
  }
  return sheet;
}

export async function saveWorkbook(workbook: ExcelJS.Workbook, filePath: string, createBackup: boolean = false): Promise<void> {
  try {
    if (createBackup) {
      try {
        await fs.access(filePath);
        const backupPath = `${filePath}.backup`;
        await fs.copyFile(filePath, backupPath);
      } catch {
        // File doesn't exist, no backup needed
      }
    }

    await workbook.xlsx.writeFile(filePath);
  } catch (error) {
    throw new Error(`${ERROR_MESSAGES.WRITE_ERROR}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

export function columnLetterToNumber(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}

export function columnNumberToLetter(num: number): string {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

export function cellValueToString(value: CellValue): string {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'object') {
    if ('formula' in value && value.formula) {
      return `=${value.formula}`;
    }
    if ('result' in value) {
      return String(value.result);
    }
    if ('text' in value) {
      return String(value.text);
    }
    return JSON.stringify(value);
  }
  return String(value);
}

export function formatDataAsTable(data: any[][], headers?: string[]): string {
  if (data.length === 0) {
    return 'No data';
  }

  const tableData = headers ? [headers, ...data] : data;
  const colWidths: number[] = [];

  // Calculate column widths
  for (const row of tableData) {
    for (let i = 0; i < row.length; i++) {
      const cellText = cellValueToString(row[i]);
      colWidths[i] = Math.max(colWidths[i] || 0, cellText.length);
    }
  }

  // Build table
  let table = '';
  for (let i = 0; i < tableData.length; i++) {
    const row = tableData[i];
    const cells = row.map((cell, j) => {
      const text = cellValueToString(cell);
      return text.padEnd(colWidths[j] || 0);
    });
    table += '| ' + cells.join(' | ') + ' |\n';

    // Add separator after header
    if (headers && i === 0) {
      table += '| ' + colWidths.map(w => '-'.repeat(w)).join(' | ') + ' |\n';
    }
  }

  return table;
}

export function parseRange(range: string): { startCol: number; startRow: number; endCol: number; endRow: number } {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(ERROR_MESSAGES.INVALID_RANGE);
  }

  return {
    startCol: columnLetterToNumber(match[1]),
    startRow: parseInt(match[2]),
    endCol: columnLetterToNumber(match[3]),
    endRow: parseInt(match[4]),
  };
}
