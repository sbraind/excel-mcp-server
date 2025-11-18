import { loadWorkbook, getSheet, parseRange, cellValueToString, formatDataAsTable } from './helpers.js';
import type { ResponseFormat } from '../types.js';

interface SearchResult {
  cellAddress: string;
  row: number;
  column: number;
  value: any;
}

export async function searchValue(
  filePath: string,
  sheetName: string,
  searchValue: any,
  range?: string,
  caseSensitive: boolean = false,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const results: SearchResult[] = [];
  let startRow = 1;
  let endRow = sheet.rowCount;
  let startCol = 1;
  let endCol = sheet.columnCount;

  if (range) {
    const parsed = parseRange(range);
    startRow = parsed.startRow;
    endRow = parsed.endRow;
    startCol = parsed.startCol;
    endCol = parsed.endCol;
  }

  const searchStr = String(searchValue);
  const compareValue = caseSensitive ? searchStr : searchStr.toLowerCase();

  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      const cell = sheet.getRow(row).getCell(col);
      const cellStr = cellValueToString(cell.value);
      const cellCompare = caseSensitive ? cellStr : cellStr.toLowerCase();

      if (cellCompare.includes(compareValue)) {
        results.push({
          cellAddress: cell.address,
          row,
          column: col,
          value: cell.value,
        });
      }
    }
  }

  if (responseFormat === 'markdown') {
    let md = `# Search Results\n\n`;
    md += `**Sheet**: ${sheetName}\n`;
    md += `**Search value**: ${searchValue}\n`;
    if (range) {
      md += `**Range**: ${range}\n`;
    }
    md += `**Case sensitive**: ${caseSensitive}\n`;
    md += `**Matches found**: ${results.length}\n\n`;

    if (results.length > 0) {
      md += '## Results\n\n';
      const tableData = results.map(r => [r.cellAddress, r.row, r.column, cellValueToString(r.value)]);
      md += formatDataAsTable(tableData, ['Cell', 'Row', 'Column', 'Value']);
    }

    return md;
  }

  return JSON.stringify({
    sheetName,
    searchValue,
    range,
    caseSensitive,
    matchesFound: results.length,
    results,
  }, null, 2);
}

export async function filterRows(
  filePath: string,
  sheetName: string,
  column: string | number,
  condition: 'equals' | 'contains' | 'greater_than' | 'less_than' | 'not_empty',
  value?: any,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const colNumber = typeof column === 'string'
    ? column.charCodeAt(0) - 64
    : column;

  const filteredRows: any[][] = [];
  const rowNumbers: number[] = [];

  sheet.eachRow((row, rowNumber) => {
    const cell = row.getCell(colNumber);
    const cellValue = cell.value;
    const cellStr = cellValueToString(cellValue);

    let match = false;

    switch (condition) {
      case 'equals':
        match = cellStr === String(value);
        break;
      case 'contains':
        match = cellStr.toLowerCase().includes(String(value).toLowerCase());
        break;
      case 'greater_than':
        const numValue = typeof cellValue === 'number' ? cellValue : parseFloat(cellStr);
        match = !isNaN(numValue) && numValue > Number(value);
        break;
      case 'less_than':
        const numVal = typeof cellValue === 'number' ? cellValue : parseFloat(cellStr);
        match = !isNaN(numVal) && numVal < Number(value);
        break;
      case 'not_empty':
        match = cellValue !== null && cellValue !== undefined && cellStr.trim() !== '';
        break;
    }

    if (match) {
      const rowData: any[] = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        rowData.push(cell.value);
      });
      filteredRows.push(rowData);
      rowNumbers.push(rowNumber);
    }
  });

  if (responseFormat === 'markdown') {
    let md = `# Filtered Rows\n\n`;
    md += `**Sheet**: ${sheetName}\n`;
    md += `**Column**: ${column}\n`;
    md += `**Condition**: ${condition}\n`;
    if (value !== undefined) {
      md += `**Value**: ${value}\n`;
    }
    md += `**Matches found**: ${filteredRows.length}\n\n`;

    if (filteredRows.length > 0) {
      md += '## Results\n\n';
      md += formatDataAsTable(filteredRows.slice(0, 100));
      if (filteredRows.length > 100) {
        md += `\n\n*Showing first 100 of ${filteredRows.length} rows*`;
      }
    }

    return md;
  }

  return JSON.stringify({
    sheetName,
    column,
    condition,
    value,
    matchesFound: filteredRows.length,
    rowNumbers,
    data: filteredRows,
  }, null, 2);
}
