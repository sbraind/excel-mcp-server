import { loadWorkbook, getSheet, cellValueToString, formatDataAsTable, parseRange } from './helpers.js';
import type { WorkbookInfo, SheetInfo, ResponseFormat } from '../types.js';

export async function readWorkbook(filePath: string, responseFormat: ResponseFormat = 'json'): Promise<string> {
  const workbook = await loadWorkbook(filePath);

  const sheets: SheetInfo[] = [];
  workbook.eachSheet((worksheet) => {
    sheets.push({
      name: worksheet.name,
      rowCount: worksheet.rowCount,
      columnCount: worksheet.columnCount,
      state: worksheet.state,
    });
  });

  const info: WorkbookInfo = {
    sheets,
    creator: workbook.creator,
    created: workbook.created,
    modified: workbook.modified,
  };

  if (responseFormat === 'markdown') {
    let md = `# Workbook: ${filePath}\n\n`;
    md += `**Created**: ${info.created ? new Date(info.created).toLocaleString() : 'N/A'}\n`;
    md += `**Modified**: ${info.modified ? new Date(info.modified).toLocaleString() : 'N/A'}\n`;
    md += `**Creator**: ${info.creator || 'N/A'}\n\n`;
    md += `## Sheets (${sheets.length})\n\n`;
    for (const sheet of sheets) {
      md += `- **${sheet.name}**: ${sheet.rowCount} rows × ${sheet.columnCount} columns`;
      if (sheet.state && sheet.state !== 'visible') {
        md += ` (${sheet.state})`;
      }
      md += '\n';
    }
    return md;
  }

  return JSON.stringify(info, null, 2);
}

export async function readSheet(
  filePath: string,
  sheetName: string,
  range?: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  let data: any[][] = [];

  if (range) {
    const { startRow, startCol, endRow, endCol } = parseRange(range);
    for (let row = startRow; row <= endRow; row++) {
      const rowData: any[] = [];
      for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getRow(row).getCell(col);
        rowData.push(cell.value);
      }
      data.push(rowData);
    }
  } else {
    sheet.eachRow((row) => {
      const rowData: any[] = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        rowData.push(cell.value);
      });
      data.push(rowData);
    });
  }

  if (responseFormat === 'markdown') {
    let md = `# Sheet: ${sheetName}\n\n`;
    if (range) {
      md += `**Range**: ${range}\n\n`;
    }
    md += `**Rows**: ${data.length}\n`;
    md += `**Columns**: ${data[0]?.length || 0}\n\n`;
    md += '## Data Preview\n\n';
    md += formatDataAsTable(data.slice(0, 100));
    if (data.length > 100) {
      md += `\n\n*Showing first 100 of ${data.length} rows*`;
    }
    return md;
  }

  return JSON.stringify({ sheetName, range, rowCount: data.length, columnCount: data[0]?.length || 0, data }, null, 2);
}

export async function readRange(
  filePath: string,
  sheetName: string,
  range: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  return readSheet(filePath, sheetName, range, responseFormat);
}

export async function getCell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const cell = sheet.getCell(cellAddress);
  const result = {
    address: cellAddress,
    value: cell.value,
    type: cell.type,
    formula: cell.formula,
    numFmt: cell.numFmt,
  };

  if (responseFormat === 'markdown') {
    let md = `# Cell: ${cellAddress}\n\n`;
    md += `**Sheet**: ${sheetName}\n`;
    md += `**Value**: ${cellValueToString(cell.value)}\n`;
    md += `**Type**: ${cell.type}\n`;
    if (cell.formula) {
      md += `**Formula**: =${cell.formula}\n`;
    }
    if (cell.numFmt) {
      md += `**Format**: ${cell.numFmt}\n`;
    }
    return md;
  }

  return JSON.stringify(result, null, 2);
}

export async function getFormula(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const cell = sheet.getCell(cellAddress);
  const result = {
    address: cellAddress,
    formula: cell.formula || null,
    value: cell.value,
  };

  if (responseFormat === 'markdown') {
    let md = `# Formula: ${cellAddress}\n\n`;
    md += `**Sheet**: ${sheetName}\n`;
    if (cell.formula) {
      md += `**Formula**: =${cell.formula}\n`;
      md += `**Result**: ${cellValueToString(cell.value)}\n`;
    } else {
      md += `**No formula** (direct value: ${cellValueToString(cell.value)})\n`;
    }
    return md;
  }

  return JSON.stringify(result, null, 2);
}
