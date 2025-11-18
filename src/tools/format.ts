import type { CellFormat } from '../types.js';
import { loadWorkbook, getSheet, saveWorkbook, columnLetterToNumber } from './helpers.js';

export async function formatCell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  format: CellFormat,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const cell = sheet.getCell(cellAddress);

  // Apply font formatting
  if (format.font) {
    cell.font = {
      ...cell.font,
      name: format.font.name,
      size: format.font.size,
      bold: format.font.bold,
      italic: format.font.italic,
      underline: format.font.underline,
      color: format.font.color ? { argb: format.font.color } : undefined,
    };
  }

  // Apply fill formatting
  if (format.fill) {
    cell.fill = {
      type: 'pattern',
      pattern: format.fill.pattern,
      fgColor: format.fill.fgColor ? { argb: format.fill.fgColor } : undefined,
      bgColor: format.fill.bgColor ? { argb: format.fill.bgColor } : undefined,
    };
  }

  // Apply alignment
  if (format.alignment) {
    cell.alignment = {
      ...cell.alignment,
      ...format.alignment,
    };
  }

  // Apply borders
  if (format.border) {
    const border: any = {};
    if (format.border.top) {
      border.top = {
        style: format.border.top.style,
        color: format.border.top.color ? { argb: format.border.top.color } : undefined,
      };
    }
    if (format.border.left) {
      border.left = {
        style: format.border.left.style,
        color: format.border.left.color ? { argb: format.border.left.color } : undefined,
      };
    }
    if (format.border.bottom) {
      border.bottom = {
        style: format.border.bottom.style,
        color: format.border.bottom.color ? { argb: format.border.bottom.color } : undefined,
      };
    }
    if (format.border.right) {
      border.right = {
        style: format.border.right.style,
        color: format.border.right.color ? { argb: format.border.right.color } : undefined,
      };
    }
    cell.border = border;
  }

  // Apply number format
  if (format.numFmt) {
    cell.numFmt = format.numFmt;
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Cell ${cellAddress} formatted`,
    cellAddress,
    appliedFormats: Object.keys(format),
  }, null, 2);
}

export async function setColumnWidth(
  filePath: string,
  sheetName: string,
  column: string | number,
  width: number,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const colNumber = typeof column === 'string' ? columnLetterToNumber(column) : column;
  const col = sheet.getColumn(colNumber);
  col.width = width;

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Column ${column} width set to ${width}`,
    column,
    width,
  }, null, 2);
}

export async function setRowHeight(
  filePath: string,
  sheetName: string,
  row: number,
  height: number,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const excelRow = sheet.getRow(row);
  excelRow.height = height;
  excelRow.commit();

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Row ${row} height set to ${height}`,
    row,
    height,
  }, null, 2);
}

export async function mergeCells(
  filePath: string,
  sheetName: string,
  range: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  sheet.mergeCells(range);

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Cells merged in range ${range}`,
    range,
  }, null, 2);
}
