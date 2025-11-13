import { z } from 'zod';

// Common schemas
export const responseFormatSchema = z.enum(['json', 'markdown']).default('json');
export const filePathSchema = z.string().describe('Path to the Excel file');
export const sheetNameSchema = z.string().describe('Name of the sheet');
export const cellAddressSchema = z.string().regex(/^[A-Z]+\d+$/, 'Invalid cell address (e.g., A1, B2)');
export const rangeSchema = z.string().regex(/^[A-Z]+\d+:[A-Z]+\d+$/, 'Invalid range (e.g., A1:D10)');

// Read operations
export const readWorkbookSchema = z.object({
  filePath: filePathSchema,
  responseFormat: responseFormatSchema,
});

export const readSheetSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.optional().describe('Optional range to read (e.g., A1:D10)'),
  responseFormat: responseFormatSchema,
});

export const readRangeSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema,
  responseFormat: responseFormatSchema,
});

export const getCellSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  responseFormat: responseFormatSchema,
});

export const getFormulaSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  responseFormat: responseFormatSchema,
});

// Write operations
export const writeWorkbookSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema.default('Sheet1'),
  data: z.array(z.array(z.any())).describe('2D array of data to write'),
  createBackup: z.boolean().default(false),
});

export const updateCellSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  value: z.any().describe('Value to write to the cell'),
  createBackup: z.boolean().default(false),
});

export const writeRangeSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema,
  data: z.array(z.array(z.any())).describe('2D array of data to write'),
  createBackup: z.boolean().default(false),
});

export const addRowSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  data: z.array(z.any()).describe('Array of values for the new row'),
  createBackup: z.boolean().default(false),
});

export const setFormulaSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  formula: z.string().describe('Excel formula (without = sign)'),
  createBackup: z.boolean().default(false),
});

// Format operations
export const formatCellSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  format: z.object({
    font: z.object({
      name: z.string().optional(),
      size: z.number().optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      color: z.string().optional().describe('Hex color code (e.g., FF0000 for red)'),
    }).optional(),
    fill: z.object({
      type: z.literal('pattern'),
      pattern: z.enum(['solid', 'darkVertical', 'darkHorizontal', 'darkGrid']),
      fgColor: z.string().optional().describe('Foreground hex color'),
      bgColor: z.string().optional().describe('Background hex color'),
    }).optional(),
    alignment: z.object({
      horizontal: z.enum(['left', 'center', 'right', 'fill', 'justify']).optional(),
      vertical: z.enum(['top', 'middle', 'bottom']).optional(),
      wrapText: z.boolean().optional(),
    }).optional(),
    border: z.object({
      top: z.object({ style: z.string(), color: z.string().optional() }).optional(),
      left: z.object({ style: z.string(), color: z.string().optional() }).optional(),
      bottom: z.object({ style: z.string(), color: z.string().optional() }).optional(),
      right: z.object({ style: z.string(), color: z.string().optional() }).optional(),
    }).optional(),
    numFmt: z.string().optional().describe('Number format (e.g., "0.00", "$#,##0.00")'),
  }),
  createBackup: z.boolean().default(false),
});

export const setColumnWidthSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  column: z.union([z.string(), z.number()]).describe('Column letter (A, B, C) or number (1, 2, 3)'),
  width: z.number().describe('Width in Excel units (approximately characters)'),
  createBackup: z.boolean().default(false),
});

export const setRowHeightSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  row: z.number().describe('Row number (1-based)'),
  height: z.number().describe('Height in points'),
  createBackup: z.boolean().default(false),
});

export const mergeCellsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema,
  createBackup: z.boolean().default(false),
});

// Sheet management
export const createSheetSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  createBackup: z.boolean().default(false),
});

export const deleteSheetSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  createBackup: z.boolean().default(false),
});

export const renameSheetSchema = z.object({
  filePath: filePathSchema,
  oldName: z.string().describe('Current sheet name'),
  newName: z.string().describe('New sheet name'),
  createBackup: z.boolean().default(false),
});

export const duplicateSheetSchema = z.object({
  filePath: filePathSchema,
  sourceSheetName: z.string().describe('Name of sheet to duplicate'),
  newSheetName: z.string().describe('Name for the duplicated sheet'),
  createBackup: z.boolean().default(false),
});

// Operations
export const deleteRowsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  startRow: z.number().describe('Starting row number (1-based)'),
  count: z.number().describe('Number of rows to delete'),
  createBackup: z.boolean().default(false),
});

export const deleteColumnsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  startColumn: z.union([z.string(), z.number()]).describe('Starting column (letter or number)'),
  count: z.number().describe('Number of columns to delete'),
  createBackup: z.boolean().default(false),
});

export const copyRangeSchema = z.object({
  filePath: filePathSchema,
  sourceSheetName: sheetNameSchema,
  sourceRange: rangeSchema,
  targetSheetName: sheetNameSchema,
  targetCell: cellAddressSchema.describe('Top-left cell of destination'),
  createBackup: z.boolean().default(false),
});

// Analysis
export const searchValueSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  searchValue: z.any().describe('Value to search for'),
  range: rangeSchema.optional().describe('Optional range to search within'),
  caseSensitive: z.boolean().default(false),
  responseFormat: responseFormatSchema,
});

export const filterRowsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  column: z.union([z.string(), z.number()]).describe('Column to filter by'),
  condition: z.enum(['equals', 'contains', 'greater_than', 'less_than', 'not_empty']),
  value: z.any().optional().describe('Value to compare against (not needed for not_empty)'),
  responseFormat: responseFormatSchema,
});

// Charts
export const createChartSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  chartType: z.enum(['line', 'bar', 'column', 'pie', 'scatter', 'area']),
  dataRange: rangeSchema.describe('Range of data for the chart'),
  position: cellAddressSchema.describe('Top-left cell where chart will be placed'),
  title: z.string().optional().describe('Chart title'),
  showLegend: z.boolean().default(true),
  createBackup: z.boolean().default(false),
});

// Pivot Tables
export const createPivotTableSchema = z.object({
  filePath: filePathSchema,
  sourceSheetName: sheetNameSchema.describe('Sheet containing source data'),
  sourceRange: rangeSchema.describe('Range of source data'),
  targetSheetName: sheetNameSchema.describe('Sheet for pivot table'),
  targetCell: cellAddressSchema.describe('Top-left cell for pivot table'),
  rows: z.array(z.string()).describe('Fields for row labels'),
  columns: z.array(z.string()).optional().describe('Fields for column labels'),
  values: z.array(z.object({
    field: z.string(),
    aggregation: z.enum(['sum', 'count', 'average', 'min', 'max']),
  })).describe('Fields to aggregate'),
  createBackup: z.boolean().default(false),
});

// Tables
export const createTableSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.describe('Range to convert to table'),
  tableName: z.string().describe('Name for the table'),
  tableStyle: z.string().optional().default('TableStyleMedium2').describe('Excel table style'),
  showFirstColumn: z.boolean().default(false),
  showLastColumn: z.boolean().default(false),
  showRowStripes: z.boolean().default(true),
  showColumnStripes: z.boolean().default(false),
  createBackup: z.boolean().default(false),
});

// Validation operations
export const validateFormulaSyntaxSchema = z.object({
  formula: z.string().describe('Formula to validate (without = sign)'),
});

export const validateExcelRangeSchema = z.object({
  range: z.string().describe('Range to validate (e.g., A1:D10)'),
});

export const getDataValidationInfoSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  responseFormat: responseFormatSchema,
});

// Advanced operations
export const insertRowsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  startRow: z.number().describe('Row number where to insert (1-based)'),
  count: z.number().describe('Number of rows to insert'),
  createBackup: z.boolean().default(false),
});

export const insertColumnsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  startColumn: z.union([z.string(), z.number()]).describe('Column where to insert (letter or number)'),
  count: z.number().describe('Number of columns to insert'),
  createBackup: z.boolean().default(false),
});

export const unmergeCellsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.describe('Range to unmerge'),
  createBackup: z.boolean().default(false),
});

export const getMergedCellsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  responseFormat: responseFormatSchema,
});

// Conditional formatting
export const applyConditionalFormatSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema,
  ruleType: z.enum(['cellValue', 'colorScale', 'dataBar', 'topBottom']),
  condition: z.object({
    operator: z.enum(['greaterThan', 'lessThan', 'between', 'equal', 'notEqual', 'containsText']).optional(),
    value: z.any().optional(),
    value2: z.any().optional().describe('Second value for "between" operator'),
  }).optional(),
  style: z.object({
    font: z.object({
      color: z.string().optional(),
      bold: z.boolean().optional(),
    }).optional(),
    fill: z.object({
      type: z.literal('pattern'),
      pattern: z.enum(['solid', 'darkVertical', 'darkHorizontal', 'darkGrid']),
      fgColor: z.string().optional(),
    }).optional(),
  }).optional(),
  colorScale: z.object({
    minColor: z.string().optional().default('FFFF0000'),
    midColor: z.string().optional(),
    maxColor: z.string().optional().default('FF00FF00'),
  }).optional().describe('For colorScale type'),
  createBackup: z.boolean().default(false),
});
