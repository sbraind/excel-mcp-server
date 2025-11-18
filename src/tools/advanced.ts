import { loadWorkbook, getSheet, saveWorkbook, columnLetterToNumber, formatDataAsTable } from './helpers.js';
import type { ResponseFormat } from '../types.js';

export async function insertRows(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Insert rows using ExcelJS
  sheet.spliceRows(startRow, 0, ...Array(count).fill([]));

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Inserted ${count} row(s) at row ${startRow}`,
    startRow,
    count,
  }, null, 2);
}

export async function insertColumns(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const colNumber = typeof startColumn === 'string' ? columnLetterToNumber(startColumn) : startColumn;

  // Insert columns using ExcelJS
  sheet.spliceColumns(colNumber, 0, ...Array(count).fill([]));

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Inserted ${count} column(s) at column ${startColumn}`,
    startColumn,
    count,
  }, null, 2);
}

export async function unmergeCells(
  filePath: string,
  sheetName: string,
  range: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Unmerge cells
  sheet.unMergeCells(range);

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Cells unmerged in range ${range}`,
    range,
  }, null, 2);
}

export async function getMergedCells(
  filePath: string,
  sheetName: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Get all merged cells from the sheet model
  const mergedCells: string[] = [];

  if (sheet.model.merges) {
    sheet.model.merges.forEach((merge) => {
      mergedCells.push(merge);
    });
  }

  if (responseFormat === 'markdown') {
    let md = `# Merged Cells in ${sheetName}\n\n`;
    md += `**Total merged ranges**: ${mergedCells.length}\n\n`;

    if (mergedCells.length > 0) {
      md += '## Merged Ranges\n\n';
      const tableData = mergedCells.map((range, index) => [index + 1, range]);
      md += formatDataAsTable(tableData, ['#', 'Range']);
    } else {
      md += '*No merged cells found*\n';
    }

    return md;
  }

  return JSON.stringify({
    sheetName,
    mergedCellsCount: mergedCells.length,
    mergedRanges: mergedCells,
  }, null, 2);
}
