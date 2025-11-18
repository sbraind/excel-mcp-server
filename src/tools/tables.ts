import { loadWorkbook, getSheet, saveWorkbook, parseRange } from './helpers.js';

export async function createTable(
  filePath: string,
  sheetName: string,
  range: string,
  tableName: string,
  tableStyle: string = 'TableStyleMedium2',
  showFirstColumn: boolean = false,
  showLastColumn: boolean = false,
  showRowStripes: boolean = true,
  showColumnStripes: boolean = false,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Parse range
  const { startRow, startCol, endRow, endCol } = parseRange(range);

  // Read the headers from the first row
  const headerRow = sheet.getRow(startRow);
  const headers: any[] = [];
  for (let col = startCol; col <= endCol; col++) {
    const cell = headerRow.getCell(col);
    headers.push({ name: cell.value?.toString() || `Column${col}` });
  }

  // ExcelJS supports tables
  sheet.addTable({
    name: tableName,
    ref: range,
    headerRow: true,
    totalsRow: false,
    style: {
      theme: tableStyle as any,
      showFirstColumn,
      showLastColumn,
      showRowStripes,
      showColumnStripes,
    },
    columns: headers,
    rows: [],
  });

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Table "${tableName}" created`,
    range,
    tableName,
    style: tableStyle,
    columns: headers.length,
    rows: endRow - startRow,
  }, null, 2);
}
