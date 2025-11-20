import ExcelJS from 'exceljs';
import { loadWorkbook, getSheet, saveWorkbook, parseRange, columnNumberToLetter } from './helpers.js';
import {
  isExcelRunning,
  isFileOpenInExcel,
  updateCellViaAppleScript,
  saveFileViaAppleScript,
  addRowViaAppleScript,
  writeRangeViaAppleScript,
  setFormulaViaAppleScript,
} from './excel-applescript.js';

export async function writeWorkbook(
  filePath: string,
  sheetName: string,
  data: any[][],
  createBackup: boolean = false
): Promise<string> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName);

  // Write data
  data.forEach((row, rowIndex) => {
    const excelRow = sheet.getRow(rowIndex + 1);
    row.forEach((value, colIndex) => {
      excelRow.getCell(colIndex + 1).value = value;
    });
    excelRow.commit();
  });

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Workbook created at ${filePath}`,
    sheetName,
    rowsWritten: data.length,
    columnsWritten: data[0]?.length || 0,
  }, null, 2);
}

export async function updateCell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  value: any,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and has this file open
  console.error(`[updateCell] Checking execution method for ${filePath}`);
  const excelRunning = await isExcelRunning();
  const fileOpen = excelRunning ? await isFileOpenInExcel(filePath) : false;

  console.error(`[updateCell] Method selection: excelRunning=${excelRunning}, fileOpen=${fileOpen}`);

  if (fileOpen) {
    console.error(`[updateCell] Using AppleScript method for real-time collaboration`);
    // Use AppleScript for real-time collaboration
    await updateCellViaAppleScript(filePath, sheetName, cellAddress, value);
    await saveFileViaAppleScript(filePath);

    return JSON.stringify({
      success: true,
      message: `Cell ${cellAddress} updated (via Excel)`,
      cellAddress,
      newValue: value,
      method: 'applescript',
      note: 'Changes are visible immediately in Excel'
    }, null, 2);
  } else {
    console.error(`[updateCell] Using ExcelJS method (file-based editing)`);
    // Fallback to ExcelJS for file-based editing
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const cell = sheet.getCell(cellAddress);
    cell.value = value;

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Cell ${cellAddress} updated`,
      cellAddress,
      newValue: value,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.'
    }, null, 2);
  }
}

export async function writeRange(
  filePath: string,
  sheetName: string,
  range: string,
  data: any[][],
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and has this file open
  console.error(`[writeRange] Checking execution method for ${filePath}`);
  const excelRunning = await isExcelRunning();
  const fileOpen = excelRunning ? await isFileOpenInExcel(filePath) : false;

  console.error(`[writeRange] Method selection: excelRunning=${excelRunning}, fileOpen=${fileOpen}`);

  if (fileOpen) {
    console.error(`[writeRange] Using AppleScript method for real-time collaboration`);
    // Use AppleScript for real-time collaboration
    const { startRow, startCol } = parseRange(range);
    const startCell = `${columnNumberToLetter(startCol)}${startRow}`;
    await writeRangeViaAppleScript(filePath, sheetName, startCell, data);
    await saveFileViaAppleScript(filePath);

    return JSON.stringify({
      success: true,
      message: `Range ${range} updated (via Excel)`,
      range,
      rowsWritten: data.length,
      columnsWritten: data[0]?.length || 0,
      method: 'applescript',
      note: 'Changes are visible immediately in Excel'
    }, null, 2);
  } else {
    console.error(`[writeRange] Using ExcelJS method (file-based editing)`);
    // Fallback to ExcelJS for file-based editing
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const { startRow, startCol } = parseRange(range);

    data.forEach((row, rowIndex) => {
      const excelRow = sheet.getRow(startRow + rowIndex);
      row.forEach((value, colIndex) => {
        excelRow.getCell(startCol + colIndex).value = value;
      });
      excelRow.commit();
    });

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Range ${range} updated`,
      range,
      rowsWritten: data.length,
      columnsWritten: data[0]?.length || 0,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.'
    }, null, 2);
  }
}

export async function addRow(
  filePath: string,
  sheetName: string,
  data: any[],
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and has this file open
  console.error(`[addRow] Checking execution method for ${filePath}`);
  const excelRunning = await isExcelRunning();
  const fileOpen = excelRunning ? await isFileOpenInExcel(filePath) : false;

  console.error(`[addRow] Method selection: excelRunning=${excelRunning}, fileOpen=${fileOpen}`);

  if (fileOpen) {
    console.error(`[addRow] Using AppleScript method for real-time collaboration`);
    // Use AppleScript for real-time collaboration
    await addRowViaAppleScript(filePath, sheetName, data);
    await saveFileViaAppleScript(filePath);

    return JSON.stringify({
      success: true,
      message: `Row added (via Excel)`,
      cellsWritten: data.length,
      method: 'applescript',
      note: 'Changes are visible immediately in Excel'
    }, null, 2);
  } else {
    console.error(`[addRow] Using ExcelJS method (file-based editing)`);
    // Fallback to ExcelJS for file-based editing
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const newRow = sheet.addRow(data);
    newRow.commit();

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Row added at position ${newRow.number}`,
      rowNumber: newRow.number,
      cellsWritten: data.length,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.'
    }, null, 2);
  }
}

export async function setFormula(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  formula: string,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and has this file open
  console.error(`[setFormula] Checking execution method for ${filePath}`);
  const excelRunning = await isExcelRunning();
  const fileOpen = excelRunning ? await isFileOpenInExcel(filePath) : false;

  console.error(`[setFormula] Method selection: excelRunning=${excelRunning}, fileOpen=${fileOpen}`);

  // Remove leading = if present
  const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;

  if (fileOpen) {
    console.error(`[setFormula] Using AppleScript method for real-time collaboration`);
    // Use AppleScript for real-time collaboration
    await setFormulaViaAppleScript(filePath, sheetName, cellAddress, cleanFormula);
    await saveFileViaAppleScript(filePath);

    return JSON.stringify({
      success: true,
      message: `Formula set in cell ${cellAddress} (via Excel)`,
      cellAddress,
      formula: `=${cleanFormula}`,
      method: 'applescript',
      note: 'Changes are visible immediately in Excel'
    }, null, 2);
  } else {
    console.error(`[setFormula] Using ExcelJS method (file-based editing)`);
    // Fallback to ExcelJS for file-based editing
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const cell = sheet.getCell(cellAddress);
    cell.value = { formula: cleanFormula };

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Formula set in cell ${cellAddress}`,
      cellAddress,
      formula: `=${cleanFormula}`,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.'
    }, null, 2);
  }
}
