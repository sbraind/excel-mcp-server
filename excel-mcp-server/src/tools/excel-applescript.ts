import { exec } from 'child_process';
import { promisify } from 'util';
import { basename } from 'path';

const execAsync = promisify(exec);

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
    const { stdout } = await execAsync(`osascript -e '${script}'`);
    return stdout.trim() === 'true';
  } catch {
    return false;
  }
}

/**
 * Check if a specific Excel file is open
 */
export async function isFileOpenInExcel(filePath: string): Promise<boolean> {
  try {
    const fileName = basename(filePath);
    const script = `
      tell application "Microsoft Excel"
        set openWorkbooks to name of every workbook
        return openWorkbooks contains "${fileName}"
      end tell
    `;
    const { stdout } = await execAsync(`osascript -e '${script}'`);
    return stdout.trim() === 'true';
  } catch {
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
  const escapedValue = typeof value === 'string' ? value.replace(/"/g, '\\"') : value;

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${fileName}"
        tell worksheet "${sheetName}"
          set value of range "${cellAddress}" to "${escapedValue}"
        end tell
      end tell
    end tell
  `;

  await execAsync(`osascript -e '${script}'`);
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

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${fileName}"
        tell worksheet "${sheetName}"
          return value of range "${cellAddress}" as string
        end tell
      end tell
    end tell
  `;

  const { stdout } = await execAsync(`osascript -e '${script}'`);
  return stdout.trim();
}

/**
 * Get list of sheets in an open Excel file
 */
export async function getSheetsViaAppleScript(filePath: string): Promise<string[]> {
  const fileName = basename(filePath);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${fileName}"
        return name of every worksheet
      end tell
    end tell
  `;

  const { stdout } = await execAsync(`osascript -e '${script}'`);
  // Parse the AppleScript list format: "Sheet1, Sheet2, Sheet3"
  return stdout.trim().split(', ').filter(s => s.length > 0);
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

  // First, find the last row
  const lastRowScript = `
    tell application "Microsoft Excel"
      tell workbook "${fileName}"
        tell worksheet "${sheetName}"
          return count of (get used range)
        end tell
      end tell
    end tell
  `;

  const { stdout } = await execAsync(`osascript -e '${lastRowScript}'`);
  const lastRow = parseInt(stdout.trim()) || 0;
  const newRow = lastRow + 1;

  // Add each cell value
  for (let i = 0; i < rowData.length; i++) {
    const col = String.fromCharCode(65 + i); // A, B, C, ...
    const cellAddress = `${col}${newRow}`;
    const value = rowData[i];
    const escapedValue = typeof value === 'string' ? value.replace(/"/g, '\\"') : value;

    const script = `
      tell application "Microsoft Excel"
        tell workbook "${fileName}"
          tell worksheet "${sheetName}"
            set value of range "${cellAddress}" to "${escapedValue}"
          end tell
        end tell
      end tell
    `;

    await execAsync(`osascript -e '${script}'`);
  }
}

/**
 * Save the open Excel file
 */
export async function saveFileViaAppleScript(filePath: string): Promise<void> {
  const fileName = basename(filePath);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${fileName}"
        save
      end tell
    end tell
  `;

  await execAsync(`osascript -e '${script}'`);
}

/**
 * Create a new sheet in an open Excel file
 */
export async function createSheetViaAppleScript(
  filePath: string,
  sheetName: string
): Promise<void> {
  const fileName = basename(filePath);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${fileName}"
        make new worksheet with properties {name:"${sheetName}"}
      end tell
    end tell
  `;

  await execAsync(`osascript -e '${script}'`);
}

/**
 * Delete a sheet in an open Excel file
 */
export async function deleteSheetViaAppleScript(
  filePath: string,
  sheetName: string
): Promise<void> {
  const fileName = basename(filePath);

  const script = `
    tell application "Microsoft Excel"
      tell workbook "${fileName}"
        delete worksheet "${sheetName}"
      end tell
    end tell
  `;

  await execAsync(`osascript -e '${script}'`);
}
