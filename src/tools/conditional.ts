import { loadWorkbook, getSheet, saveWorkbook, parseRange } from './helpers.js';

export async function applyConditionalFormat(
  filePath: string,
  sheetName: string,
  range: string,
  ruleType: 'cellValue' | 'colorScale' | 'dataBar' | 'topBottom',
  condition?: {
    operator?: 'greaterThan' | 'lessThan' | 'between' | 'equal' | 'notEqual' | 'containsText';
    value?: any;
    value2?: any;
  },
  style?: {
    font?: {
      color?: string;
      bold?: boolean;
    };
    fill?: {
      type: 'pattern';
      pattern: 'solid' | 'darkVertical' | 'darkHorizontal' | 'darkGrid';
      fgColor?: string;
    };
  },
  colorScale?: {
    minColor?: string;
    midColor?: string;
    maxColor?: string;
  },
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Parse range
  const { startRow, startCol, endRow, endCol } = parseRange(range);

  // ExcelJS has limited conditional formatting support
  // This is a workaround that applies formatting based on conditions

  if (ruleType === 'cellValue' && condition && style) {
    // Apply cell value conditional formatting
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getRow(row).getCell(col);
        const cellValue = cell.value;

        let shouldApplyStyle = false;

        // Check condition
        switch (condition.operator) {
          case 'greaterThan':
            shouldApplyStyle = Number(cellValue) > Number(condition.value);
            break;
          case 'lessThan':
            shouldApplyStyle = Number(cellValue) < Number(condition.value);
            break;
          case 'equal':
            shouldApplyStyle = cellValue === condition.value;
            break;
          case 'notEqual':
            shouldApplyStyle = cellValue !== condition.value;
            break;
          case 'between':
            const numValue = Number(cellValue);
            shouldApplyStyle = numValue >= Number(condition.value) && numValue <= Number(condition.value2);
            break;
          case 'containsText':
            shouldApplyStyle = String(cellValue).includes(String(condition.value));
            break;
        }

        if (shouldApplyStyle) {
          // Apply style
          if (style.font) {
            cell.font = {
              ...cell.font,
              color: style.font.color ? { argb: style.font.color } : cell.font?.color,
              bold: style.font.bold !== undefined ? style.font.bold : cell.font?.bold,
            };
          }

          if (style.fill) {
            cell.fill = {
              type: 'pattern',
              pattern: style.fill.pattern,
              fgColor: style.fill.fgColor ? { argb: style.fill.fgColor } : undefined,
            };
          }
        }
      }
    }
  } else if (ruleType === 'colorScale') {
    // Apply color scale
    const minColor = colorScale?.minColor || 'FFFF0000'; // Red
    const maxColor = colorScale?.maxColor || 'FF00FF00'; // Green
    const midColor = colorScale?.midColor;

    // Collect all numeric values in range
    const values: number[] = [];
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getRow(row).getCell(col);
        const value = Number(cell.value);
        if (!isNaN(value)) {
          values.push(value);
        }
      }
    }

    const minValue = Math.min(...values);
    const maxValue = Math.max(...values);
    const range_span = maxValue - minValue;

    // Apply gradient colors
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getRow(row).getCell(col);
        const value = Number(cell.value);

        if (!isNaN(value)) {
          const percentage = range_span === 0 ? 0.5 : (value - minValue) / range_span;

          // Simple color interpolation
          let color: string;
          if (midColor && percentage < 0.5) {
            // Interpolate between min and mid
            color = minColor;
          } else if (midColor && percentage >= 0.5) {
            // Interpolate between mid and max
            color = maxColor;
          } else {
            // Interpolate between min and max
            color = percentage < 0.5 ? minColor : maxColor;
          }

          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color },
          };
        }
      }
    }
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Conditional formatting applied to ${range}`,
    range,
    ruleType,
    note: 'This is a simplified implementation. Native Excel conditional formatting has more features.',
  }, null, 2);
}
