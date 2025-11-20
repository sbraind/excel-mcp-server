import { exec } from 'child_process';
import { promisify } from 'util';
import { basename } from 'path';

const execAsync = promisify(exec);

// Configuration
const APPLESCRIPT_TIMEOUT = 10000; // 10 seconds
const MAX_RETRIES = 3;
const RETRY_DELAY = 500; // 500ms

/**
 * Escape a string for safe use in AppleScript
 * Escapes backslashes, double quotes, and single quotes
 */
function escapeAppleScriptString(str: string): string {
  return str
    .replace(/\\/g, '\\\\')   // Backslashes first
    .replace(/"/g, '\\"')      // Double quotes
    .replace(/'/g, "\\'");     // Single quotes/apostrophes
}

/**
 * Convert column number to Excel letter (1=A, 27=AA, etc.)
 */
function numberToColumnLetter(num: number): string {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

/**
 * Convert Excel column letter to number (A=1, AA=27, etc.)
 */
function columnLetterToNumber(letter: string): number {
  let num = 0;
  for (let i = 0; i < letter.length; i++) {
    num = num * 26 + (letter.charCodeAt(i) - 64);
  }
  return num;
}

/**
 * Validate cell address format (e.g., "A1", "Z100")
 * Security: Prevents injection attacks by ensuring cell addresses match expected format
 */
function validateCellAddress(address: string): void {
  if (!/^[A-Z]+\d+$/.test(address)) {
    throw new Error(`Invalid cell address format: ${address}. Expected format like "A1" or "AA100"`);
  }
  // Validate column is within Excel limits (A-XFD = 1-16384)
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (match) {
    const colNum = columnLetterToNumber(match[1]);
    if (colNum > 16384) {
      throw new Error(`Column ${match[1]} exceeds Excel's maximum column (XFD)`);
    }
    const rowNum = parseInt(match[2]);
    if (rowNum < 1 || rowNum > 1048576) {
      throw new Error(`Row ${rowNum} must be between 1 and 1048576`);
    }
  }
}

/**
 * Validate range format (e.g., "A1:B10")
 * Security: Prevents injection attacks by ensuring ranges match expected format
 */
function validateRange(range: string): void {
  if (!/^[A-Z]+\d+:[A-Z]+\d+$/.test(range)) {
    throw new Error(`Invalid range format: ${range}. Expected format like "A1:B10"`);
  }
  const [start, end] = range.split(':');
  validateCellAddress(start);
  validateCellAddress(end);
}

/**
 * Format value for AppleScript based on type
 * Numbers are passed without quotes, strings are escaped and quoted
 */
function formatValueForAppleScript(value: string | number): string {
  if (typeof value === 'number') {
    return String(value); // No quotes for numbers
  }
  // String: escape and wrap in quotes
  return `"${escapeAppleScriptString(value)}"`;
}

/**
 * Execute AppleScript with timeout and retry logic
 * Security: Escapes single quotes to prevent command injection
 */
async function execAppleScriptWithRetry(
  script: string,
  retries: number = MAX_RETRIES,
  timeout: number = APPLESCRIPT_TIMEOUT
): Promise<string> {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      // Security fix: Escape single quotes in the script to prevent command injection
      // Replace ' with '\'' (close quote, literal quote, open quote)
      const escapedScript = script.replace(/'/g, "'\\''");
      const { stdout } = await execAsync(`osascript -e '${escapedScript}'`, {
        timeout,
      });
      return stdout.trim();
    } catch (error: any) {
      console.error(`[AppleScript] Attempt ${attempt}/${retries} failed:`, {
        error: error.message,
        code: error.code,
        killed: error.killed,
      });

      if (attempt === retries) {
        throw error;
      }

      // Wait before retry with exponential backoff
      await new Promise(resolve => setTimeout(resolve, RETRY_DELAY * attempt));
    }
  }
  throw new Error('Max retries exceeded');
}

/**
 * Check if Microsoft Excel is running
 */
export async function isExcelRunning(): Promise<boolean> {
  try {
    const script = `
      tell application "System Events"
        return (name of processes) contains "Microsoft Excel"
      end tell
    `;
    const result = await execAppleScriptWithRetry(script, 2, 5000);
    const isRunning = result === 'true';
    console.error(`[AppleScript] Excel running: ${isRunning}`);
    return isRunning;
  } catch (error: any) {
    console.error('[AppleScript] Failed to check if Excel is running:', error.message);
    return false;
  }
}

/**
 * Check if a specific Excel file is open
 */
export async function isFileOpenInExcel(filePath: string): Promise<boolean> {
  try {
    const fileName = basename(filePath);
    const escapedFileName = escapeAppleScriptString(fileName);
    console.error(`[AppleScript] Checking if file is open: ${fileName}`);

    const script = `
      tell application "Microsoft Excel"
        set openWorkbooks to name of every workbook
        return openWorkbooks contains "${escapedFileName}"
      end tell
    `;
    const result = await execAppleScriptWithRetry(script, 2, 5000);
    const isOpen = result === 'true';
    console.error(`[AppleScript] File "${fileName}" open: ${isOpen}`);
    return isOpen;
  } catch (error: any) {
    console.error(`[AppleScript] Failed to check if file is open:`, {
      file: basename(filePath),
      error: error.message,
    });
    return false;
  }
}

/**
 * Update a cell value in an open Excel file
 */
export async function updateCellViaAppleScript(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  value: string | number
): Promise<void> {
  // Security: Validate cell address format before processing
  validateCellAddress(cellAddress);

  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);
  const formattedValue = formatValueForAppleScript(value);

  console.error(`[AppleScript] Updating cell ${cellAddress} in "${fileName}"/"${sheetName}"`);


  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        tell worksheet "${escapedSheetName}"
          set value of range "${cellAddress}" to ${formattedValue}
        end tell
      end tell
    end tell
  `;

  try {
    await execAppleScriptWithRetry(script);
    console.error(`[AppleScript] Successfully updated cell ${cellAddress}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to update cell:`, {
      file: fileName,
      sheet: sheetName,
      cell: cellAddress,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Read a cell value from an open Excel file
 */
export async function readCellViaAppleScript(
  filePath: string,
  sheetName: string,
  cellAddress: string
): Promise<string> {
  // Security: Validate cell address format before processing
  validateCellAddress(cellAddress);

  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        tell worksheet "${escapedSheetName}"
          return value of range "${cellAddress}" as string
        end tell
      end tell
    end tell
  `;

  const result = await execAppleScriptWithRetry(script);
  return result;
}

/**
 * Get list of sheets in an open Excel file
 */
export async function getSheetsViaAppleScript(filePath: string): Promise<string[]> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        return name of every worksheet
      end tell
    end tell
  `;

  const result = await execAppleScriptWithRetry(script);
  // Parse the AppleScript list format: "Sheet1, Sheet2, Sheet3"
  return result.split(', ').filter(s => s.length > 0);
}

/**
 * Add a new row to a sheet in an open Excel file
 */
export async function addRowViaAppleScript(
  filePath: string,
  sheetName: string,
  rowData: (string | number)[]
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  console.error(`[AppleScript] Adding row to "${fileName}"/"${sheetName}" with ${rowData.length} cells`);

  try {
    // First, find the last row
    const lastRowScript = `
      tell application "Microsoft Excel"
        tell workbook "${escapedFileName}"
          tell worksheet "${escapedSheetName}"
            return count of (get used range)
          end tell
        end tell
      end tell
    `;

    const lastRowResult = await execAppleScriptWithRetry(lastRowScript);
    const lastRow = parseInt(lastRowResult) || 0;
    const newRow = lastRow + 1;

    console.error(`[AppleScript] Adding row at position ${newRow}`);

    // Add each cell value
    for (let i = 0; i < rowData.length; i++) {
      const col = numberToColumnLetter(i + 1); // 1=A, 2=B, ..., 27=AA, etc.
      const cellAddress = `${col}${newRow}`;
      const value = rowData[i];
      const formattedValue = formatValueForAppleScript(value);

      const script = `
        tell application "Microsoft Excel"
          tell workbook "${escapedFileName}"
            tell worksheet "${escapedSheetName}"
              set value of range "${cellAddress}" to ${formattedValue}
            end tell
          end tell
        end tell
      `;

      await execAppleScriptWithRetry(script);
    }

    console.error(`[AppleScript] Successfully added row at position ${newRow}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to add row:`, {
      file: fileName,
      sheet: sheetName,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Save the open Excel file
 */
export async function saveFileViaAppleScript(filePath: string): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);

  console.error(`[AppleScript] Saving file "${fileName}"`);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        save
      end tell
    end tell
  `;

  try {
    await execAppleScriptWithRetry(script);
    console.error(`[AppleScript] Successfully saved file "${fileName}"`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to save file:`, {
      file: fileName,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Create a new sheet in an open Excel file
 */
export async function createSheetViaAppleScript(
  filePath: string,
  sheetName: string
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        make new worksheet with properties {name:"${escapedSheetName}"}
      end tell
    end tell
  `;

  await execAppleScriptWithRetry(script);
}

/**
 * Delete a sheet in an open Excel file
 */
export async function deleteSheetViaAppleScript(
  filePath: string,
  sheetName: string
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        delete worksheet "${escapedSheetName}"
      end tell
    end tell
  `;

  await execAppleScriptWithRetry(script);
}

/**
 * Write a 2D array to a range starting at startCell in an open Excel file
 */
export async function writeRangeViaAppleScript(
  filePath: string,
  sheetName: string,
  startCell: string,
  data: (string | number)[][]
): Promise<void> {
  // Security: Validate cell address format before processing
  validateCellAddress(startCell);

  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  console.error(`[AppleScript] Writing range starting at ${startCell} in "${fileName}"/"${sheetName}" with ${data.length} rows`);

  try {
    // Parse the start cell address to get row and column
    const match = startCell.match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      throw new Error(`Invalid cell address: ${startCell}`);
    }
    const startColLetter = match[1];
    const startRow = parseInt(match[2]);
    const startCol = columnLetterToNumber(startColLetter);

    // Write each cell in the data array
    for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
      const row = data[rowIdx];
      for (let colIdx = 0; colIdx < row.length; colIdx++) {
        const targetRow = startRow + rowIdx;
        const targetCol = startCol + colIdx;
        const targetColLetter = numberToColumnLetter(targetCol);
        const cellAddress = `${targetColLetter}${targetRow}`;
        const value = row[colIdx];
        const formattedValue = formatValueForAppleScript(value);

        const script = `
          tell application "Microsoft Excel"
            tell workbook "${escapedFileName}"
              tell worksheet "${escapedSheetName}"
                set value of range "${cellAddress}" to ${formattedValue}
              end tell
            end tell
          end tell
        `;

        await execAppleScriptWithRetry(script);
      }
    }

    console.error(`[AppleScript] Successfully wrote range starting at ${startCell}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to write range:`, {
      file: fileName,
      sheet: sheetName,
      startCell,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Set a formula in a cell in an open Excel file
 */
export async function setFormulaViaAppleScript(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  formula: string
): Promise<void> {
  // Security: Validate cell address format before processing
  validateCellAddress(cellAddress);

  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);
  const escapedFormula = escapeAppleScriptString(formula);

  console.error(`[AppleScript] Setting formula in cell ${cellAddress} in "${fileName}"/"${sheetName}"`);


  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        tell worksheet "${escapedSheetName}"
          set formula of range "${cellAddress}" to "${escapedFormula}"
        end tell
      end tell
    end tell
  `;

  try {
    await execAppleScriptWithRetry(script);
    console.error(`[AppleScript] Successfully set formula in cell ${cellAddress}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to set formula:`, {
      file: fileName,
      sheet: sheetName,
      cell: cellAddress,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Format a cell in an open Excel file
 * Supported format properties: fontName, fontSize, fontBold, fontItalic, fontColor, fillColor, horizontalAlignment, verticalAlignment
 */
export async function formatCellViaAppleScript(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  format: {
    fontName?: string;
    fontSize?: number;
    fontBold?: boolean;
    fontItalic?: boolean;
    fontColor?: string;
    fillColor?: string;
    horizontalAlignment?: string;
    verticalAlignment?: string;
  }
): Promise<void> {
  // Security: Validate cell address format before processing
  validateCellAddress(cellAddress);

  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  console.error(`[AppleScript] Formatting cell ${cellAddress} in "${fileName}"/"${sheetName}"`);


  try {
    // Build the AppleScript commands for each format property
    const formatCommands: string[] = [];

    if (format.fontName) {
      const escapedFontName = escapeAppleScriptString(format.fontName);
      formatCommands.push(`set font name of range "${cellAddress}" to "${escapedFontName}"`);
    }

    if (format.fontSize !== undefined) {
      formatCommands.push(`set font size of range "${cellAddress}" to ${format.fontSize}`);
    }

    if (format.fontBold !== undefined) {
      formatCommands.push(`set font bold of range "${cellAddress}" to ${format.fontBold}`);
    }

    if (format.fontItalic !== undefined) {
      formatCommands.push(`set font italic of range "${cellAddress}" to ${format.fontItalic}`);
    }

    if (format.fontColor) {
      // Color format in AppleScript: {red, green, blue} where each is 0-65535
      const escapedColor = escapeAppleScriptString(format.fontColor);
      formatCommands.push(`set font color of range "${cellAddress}" to "${escapedColor}"`);
    }

    if (format.fillColor) {
      const escapedColor = escapeAppleScriptString(format.fillColor);
      formatCommands.push(`set interior color of range "${cellAddress}" to "${escapedColor}"`);
    }

    if (format.horizontalAlignment) {
      const alignment = format.horizontalAlignment.toLowerCase();
      let alignmentConstant = 'left';
      if (alignment === 'center') alignmentConstant = 'center';
      else if (alignment === 'right') alignmentConstant = 'right';
      formatCommands.push(`set horizontal alignment of range "${cellAddress}" to ${alignmentConstant}`);
    }

    if (format.verticalAlignment) {
      const alignment = format.verticalAlignment.toLowerCase();
      let alignmentConstant = 'top';
      if (alignment === 'center' || alignment === 'middle') alignmentConstant = 'center';
      else if (alignment === 'bottom') alignmentConstant = 'bottom';
      formatCommands.push(`set vertical alignment of range "${cellAddress}" to ${alignmentConstant}`);
    }

    // Execute each format command
    for (const command of formatCommands) {
      const script = `
        tell application "Microsoft Excel"
          tell workbook "${escapedFileName}"
            tell worksheet "${escapedSheetName}"
              ${command}
            end tell
          end tell
        end tell
      `;
      await execAppleScriptWithRetry(script);
    }

    console.error(`[AppleScript] Successfully formatted cell ${cellAddress}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to format cell:`, {
      file: fileName,
      sheet: sheetName,
      cell: cellAddress,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Set column width in an open Excel file
 */
export async function setColumnWidthViaAppleScript(
  filePath: string,
  sheetName: string,
  column: string | number,
  width: number
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  // Convert column to letter if it's a number
  const columnLetter = typeof column === 'number' ? numberToColumnLetter(column) : column;

  console.error(`[AppleScript] Setting column ${columnLetter} width to ${width} in "${fileName}"/"${sheetName}"`);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        tell worksheet "${escapedSheetName}"
          set column width of column "${columnLetter}" to ${width}
        end tell
      end tell
    end tell
  `;

  try {
    await execAppleScriptWithRetry(script);
    console.error(`[AppleScript] Successfully set column ${columnLetter} width to ${width}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to set column width:`, {
      file: fileName,
      sheet: sheetName,
      column: columnLetter,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Set row height in an open Excel file
 */
export async function setRowHeightViaAppleScript(
  filePath: string,
  sheetName: string,
  row: number,
  height: number
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  console.error(`[AppleScript] Setting row ${row} height to ${height} in "${fileName}"/"${sheetName}"`);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        tell worksheet "${escapedSheetName}"
          set row height of row ${row} to ${height}
        end tell
      end tell
    end tell
  `;

  try {
    await execAppleScriptWithRetry(script);
    console.error(`[AppleScript] Successfully set row ${row} height to ${height}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to set row height:`, {
      file: fileName,
      sheet: sheetName,
      row,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Merge cells in an open Excel file
 */
export async function mergeCellsViaAppleScript(
  filePath: string,
  sheetName: string,
  range: string
): Promise<void> {
  // Security: Validate range format before processing
  validateRange(range);

  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  console.error(`[AppleScript] Merging cells ${range} in "${fileName}"/"${sheetName}"`);


  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        tell worksheet "${escapedSheetName}"
          merge range "${range}"
        end tell
      end tell
    end tell
  `;

  try {
    await execAppleScriptWithRetry(script);
    console.error(`[AppleScript] Successfully merged cells ${range}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to merge cells:`, {
      file: fileName,
      sheet: sheetName,
      range,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Unmerge cells in an open Excel file
 */
export async function unmergeCellsViaAppleScript(
  filePath: string,
  sheetName: string,
  range: string
): Promise<void> {
  // Security: Validate range format before processing
  validateRange(range);

  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  console.error(`[AppleScript] Unmerging cells ${range} in "${fileName}"/"${sheetName}"`);


  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        tell worksheet "${escapedSheetName}"
          unmerge range "${range}"
        end tell
      end tell
    end tell
  `;

  try {
    await execAppleScriptWithRetry(script);
    console.error(`[AppleScript] Successfully unmerged cells ${range}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to unmerge cells:`, {
      file: fileName,
      sheet: sheetName,
      range,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Rename a sheet in an open Excel file
 */
export async function renameSheetViaAppleScript(
  filePath: string,
  oldName: string,
  newName: string
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedOldName = escapeAppleScriptString(oldName);
  const escapedNewName = escapeAppleScriptString(newName);

  console.error(`[AppleScript] Renaming sheet from "${oldName}" to "${newName}" in "${fileName}"`);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${escapedFileName}"
        set name of worksheet "${escapedOldName}" to "${escapedNewName}"
      end tell
    end tell
  `;

  try {
    await execAppleScriptWithRetry(script);
    console.error(`[AppleScript] Successfully renamed sheet from "${oldName}" to "${newName}"`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to rename sheet:`, {
      file: fileName,
      oldName,
      newName,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Delete rows in an open Excel file
 */
export async function deleteRowsViaAppleScript(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  console.error(`[AppleScript] Deleting ${count} row(s) starting at row ${startRow} in "${fileName}"/"${sheetName}"`);

  try {
    // Delete rows one by one (delete from the same start position each time)
    for (let i = 0; i < count; i++) {
      const script = `
        tell application "Microsoft Excel"
          tell workbook "${escapedFileName}"
            tell worksheet "${escapedSheetName}"
              delete row ${startRow}
            end tell
          end tell
        end tell
      `;
      await execAppleScriptWithRetry(script);
    }

    console.error(`[AppleScript] Successfully deleted ${count} row(s) starting at row ${startRow}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to delete rows:`, {
      file: fileName,
      sheet: sheetName,
      startRow,
      count,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Delete columns in an open Excel file
 */
export async function deleteColumnsViaAppleScript(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  // Convert column to letter if it's a number
  const startColumnLetter = typeof startColumn === 'number' ? numberToColumnLetter(startColumn) : startColumn;

  console.error(`[AppleScript] Deleting ${count} column(s) starting at column ${startColumnLetter} in "${fileName}"/"${sheetName}"`);

  try {
    // Delete columns one by one (delete from the same start position each time)
    for (let i = 0; i < count; i++) {
      const script = `
        tell application "Microsoft Excel"
          tell workbook "${escapedFileName}"
            tell worksheet "${escapedSheetName}"
              delete column "${startColumnLetter}"
            end tell
          end tell
        end tell
      `;
      await execAppleScriptWithRetry(script);
    }

    console.error(`[AppleScript] Successfully deleted ${count} column(s) starting at column ${startColumnLetter}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to delete columns:`, {
      file: fileName,
      sheet: sheetName,
      startColumn: startColumnLetter,
      count,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Insert rows in an open Excel file
 */
export async function insertRowsViaAppleScript(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  console.error(`[AppleScript] Inserting ${count} row(s) at row ${startRow} in "${fileName}"/"${sheetName}"`);

  try {
    // Insert rows one by one
    for (let i = 0; i < count; i++) {
      const script = `
        tell application "Microsoft Excel"
          tell workbook "${escapedFileName}"
            tell worksheet "${escapedSheetName}"
              make new row at before row ${startRow}
            end tell
          end tell
        end tell
      `;
      await execAppleScriptWithRetry(script);
    }

    console.error(`[AppleScript] Successfully inserted ${count} row(s) at row ${startRow}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to insert rows:`, {
      file: fileName,
      sheet: sheetName,
      startRow,
      count,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Insert columns in an open Excel file
 */
export async function insertColumnsViaAppleScript(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number
): Promise<void> {
  const fileName = basename(filePath);
  const escapedFileName = escapeAppleScriptString(fileName);
  const escapedSheetName = escapeAppleScriptString(sheetName);

  // Convert column to letter if it's a number
  const startColumnLetter = typeof startColumn === 'number' ? numberToColumnLetter(startColumn) : startColumn;

  console.error(`[AppleScript] Inserting ${count} column(s) at column ${startColumnLetter} in "${fileName}"/"${sheetName}"`);

  try {
    // Insert columns one by one
    for (let i = 0; i < count; i++) {
      const script = `
        tell application "Microsoft Excel"
          tell workbook "${escapedFileName}"
            tell worksheet "${escapedSheetName}"
              make new column at before column "${startColumnLetter}"
            end tell
          end tell
        end tell
      `;
      await execAppleScriptWithRetry(script);
    }

    console.error(`[AppleScript] Successfully inserted ${count} column(s) at column ${startColumnLetter}`);
  } catch (error: any) {
    console.error(`[AppleScript] Failed to insert columns:`, {
      file: fileName,
      sheet: sheetName,
      startColumn: startColumnLetter,
      count,
      error: error.message,
    });
    throw error;
  }
}
