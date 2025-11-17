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
 */
async function execAppleScriptWithRetry(
  script: string,
  retries: number = MAX_RETRIES,
  timeout: number = APPLESCRIPT_TIMEOUT
): Promise<string> {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const { stdout } = await execAsync(`osascript -e '${script}'`, {
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
