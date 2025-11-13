import { loadWorkbook, getSheet, saveWorkbook, parseRange, columnLetterToNumber } from './helpers.js';

export async function deleteRows(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  sheet.spliceRows(startRow, count);

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Deleted ${count} row(s) starting from row ${startRow}`,
    startRow,
    count,
  }, null, 2);
}

export async function deleteColumns(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const colNumber = typeof startColumn === 'string' ? columnLetterToNumber(startColumn) : startColumn;
  sheet.spliceColumns(colNumber, count);

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Deleted ${count} column(s) starting from column ${startColumn}`,
    startColumn,
    count,
  }, null, 2);
}

export async function copyRange(
  filePath: string,
  sourceSheetName: string,
  sourceRange: string,
  targetSheetName: string,
  targetCell: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sourceSheet = getSheet(workbook, sourceSheetName);
  const targetSheet = getSheet(workbook, targetSheetName);

  const { startRow, startCol, endRow, endCol } = parseRange(sourceRange);
  const targetCellMatch = targetCell.match(/^([A-Z]+)(\d+)$/);

  if (!targetCellMatch) {
    throw new Error(`Invalid target cell address: ${targetCell}`);
  }

  const targetStartCol = columnLetterToNumber(targetCellMatch[1]);
  const targetStartRow = parseInt(targetCellMatch[2]);

  // Copy data and formatting
  for (let row = startRow; row <= endRow; row++) {
    const rowOffset = row - startRow;
    const targetRowNum = targetStartRow + rowOffset;
    const targetRow = targetSheet.getRow(targetRowNum);

    for (let col = startCol; col <= endCol; col++) {
      const colOffset = col - startCol;
      const targetColNum = targetStartCol + colOffset;

      const sourceCell = sourceSheet.getRow(row).getCell(col);
      const targetCellObj = targetRow.getCell(targetColNum);

      // Copy value
      targetCellObj.value = sourceCell.value;

      // Copy formatting
      targetCellObj.style = { ...sourceCell.style };
    }

    targetRow.commit();
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Range copied from ${sourceSheetName}!${sourceRange} to ${targetSheetName}!${targetCell}`,
    sourceSheet: sourceSheetName,
    sourceRange,
    targetSheet: targetSheetName,
    targetCell,
    rowsCopied: endRow - startRow + 1,
    columnsCopied: endCol - startCol + 1,
  }, null, 2);
}
