import { loadWorkbook, getSheet, saveWorkbook, parseRange } from './helpers.js';

export async function createPivotTable(
  filePath: string,
  sourceSheetName: string,
  sourceRange: string,
  targetSheetName: string,
  targetCell: string,
  rows: string[],
  _columns: string[] = [],
  values: Array<{ field: string; aggregation: 'sum' | 'count' | 'average' | 'min' | 'max' }>,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sourceSheet = getSheet(workbook, sourceSheetName);
  const targetSheet = getSheet(workbook, targetSheetName);

  // Parse source range
  const { startRow, startCol, endRow, endCol } = parseRange(sourceRange);

  // Read source data
  const sourceData: any[][] = [];
  for (let row = startRow; row <= endRow; row++) {
    const rowData: any[] = [];
    for (let col = startCol; col <= endCol; col++) {
      const cell = sourceSheet.getRow(row).getCell(col);
      rowData.push(cell.value);
    }
    sourceData.push(rowData);
  }

  // First row should be headers
  const headers = sourceData[0];
  const dataRows = sourceData.slice(1);

  // Create a simplified pivot table structure
  // This is a basic implementation - full pivot tables require more complex logic

  // Build aggregation
  const pivotData: Map<string, Map<string, number>> = new Map();

  dataRows.forEach((row) => {
    // Get row key
    const rowKey = rows.map(field => {
      const fieldIndex = headers.indexOf(field);
      return fieldIndex >= 0 ? String(row[fieldIndex] || '') : '';
    }).join('|');

    if (!pivotData.has(rowKey)) {
      pivotData.set(rowKey, new Map());
    }

    // Aggregate values
    values.forEach(({ field, aggregation }) => {
      const fieldIndex = headers.indexOf(field);
      if (fieldIndex >= 0) {
        const value = Number(row[fieldIndex]) || 0;
        const key = `${field}_${aggregation}`;
        const current = pivotData.get(rowKey)!.get(key) || 0;

        let newValue: number;
        switch (aggregation) {
          case 'sum':
          case 'average':
            newValue = current + value;
            break;
          case 'count':
            newValue = current + 1;
            break;
          case 'min':
            newValue = current === 0 ? value : Math.min(current, value);
            break;
          case 'max':
            newValue = Math.max(current, value);
            break;
          default:
            newValue = current;
        }

        pivotData.get(rowKey)!.set(key, newValue);
      }
    });
  });

  // Write pivot table to target
  const targetCellMatch = targetCell.match(/^([A-Z]+)(\d+)$/);
  if (!targetCellMatch) {
    throw new Error(`Invalid target cell: ${targetCell}`);
  }

  let currentRow = parseInt(targetCellMatch[2]);

  // Write headers
  const headerRow = targetSheet.getRow(currentRow);
  let col = 1;
  rows.forEach(rowField => {
    headerRow.getCell(col++).value = rowField;
  });
  values.forEach(({ field, aggregation }) => {
    headerRow.getCell(col++).value = `${field} (${aggregation})`;
  });
  headerRow.font = { bold: true };
  headerRow.commit();
  currentRow++;

  // Write data
  pivotData.forEach((valueMap, rowKey) => {
    const dataRow = targetSheet.getRow(currentRow++);
    let colIndex = 1;

    // Write row labels
    const rowParts = rowKey.split('|');
    rowParts.forEach(part => {
      dataRow.getCell(colIndex++).value = part;
    });

    // Write aggregated values
    values.forEach(({ field, aggregation }) => {
      const key = `${field}_${aggregation}`;
      let value = valueMap.get(key) || 0;

      // For average, divide by count
      if (aggregation === 'average') {
        const countKey = `${field}_count`;
        const count = valueMap.get(countKey) || 1;
        value = value / count;
      }

      dataRow.getCell(colIndex++).value = value;
    });

    dataRow.commit();
  });

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Pivot table created at ${targetSheetName}!${targetCell}`,
    rowsProcessed: dataRows.length,
    pivotRows: pivotData.size,
    note: 'This is a simplified pivot table. For full Excel pivot table features, use native Excel automation.',
  }, null, 2);
}
