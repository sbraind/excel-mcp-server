import { loadWorkbook, getSheet, saveWorkbook } from './helpers.js';

export async function createSheet(
  filePath: string,
  sheetName: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);

  // Check if sheet already exists
  if (workbook.getWorksheet(sheetName)) {
    throw new Error(`Sheet "${sheetName}" already exists`);
  }

  workbook.addWorksheet(sheetName);
  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Sheet "${sheetName}" created`,
    sheetName,
  }, null, 2);
}

export async function deleteSheet(
  filePath: string,
  sheetName: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  workbook.removeWorksheet(sheet.id);
  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Sheet "${sheetName}" deleted`,
    sheetName,
  }, null, 2);
}

export async function renameSheet(
  filePath: string,
  oldName: string,
  newName: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, oldName);

  // Check if new name already exists
  if (workbook.getWorksheet(newName)) {
    throw new Error(`Sheet "${newName}" already exists`);
  }

  sheet.name = newName;
  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Sheet renamed from "${oldName}" to "${newName}"`,
    oldName,
    newName,
  }, null, 2);
}

export async function duplicateSheet(
  filePath: string,
  sourceSheetName: string,
  newSheetName: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sourceSheet = getSheet(workbook, sourceSheetName);

  // Check if new name already exists
  if (workbook.getWorksheet(newSheetName)) {
    throw new Error(`Sheet "${newSheetName}" already exists`);
  }

  // Create new sheet
  const newSheet = workbook.addWorksheet(newSheetName);

  // Copy all data and formatting
  sourceSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const newRow = newSheet.getRow(rowNumber);
    newRow.height = row.height;

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);

      // Copy value
      newCell.value = cell.value;

      // Copy formatting
      newCell.style = { ...cell.style };
    });

    newRow.commit();
  });

  // Copy column widths
  sourceSheet.columns.forEach((column, index) => {
    if (column && column.width) {
      const newColumn = newSheet.getColumn(index + 1);
      newColumn.width = column.width;
    }
  });

  // Copy merged cells
  if (sourceSheet.model.merges) {
    sourceSheet.model.merges.forEach((merge) => {
      newSheet.mergeCells(merge);
    });
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Sheet "${sourceSheetName}" duplicated as "${newSheetName}"`,
    sourceSheetName,
    newSheetName,
    rowsCopied: sourceSheet.rowCount,
  }, null, 2);
}
