#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  InitializeRequestSchema,
  ErrorCode,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { ZodError } from 'zod';

// Import tool implementations
import { readWorkbook, readSheet, readRange, getCell, getFormula } from './tools/read.js';
import { writeWorkbook, updateCell, writeRange, addRow, setFormula } from './tools/write.js';
import { formatCell, setColumnWidth, setRowHeight, mergeCells } from './tools/format.js';
import { createSheet, deleteSheet, renameSheet, duplicateSheet } from './tools/sheets.js';
import { deleteRows, deleteColumns, copyRange } from './tools/operations.js';
import { searchValue, filterRows } from './tools/analysis.js';
import { createChart } from './tools/charts.js';
import { createPivotTable } from './tools/pivots.js';
import { createTable } from './tools/tables.js';
import { validateFormulaSyntax, validateExcelRange, getDataValidationInfo } from './tools/validation.js';
import { insertRows, insertColumns, unmergeCells, getMergedCells } from './tools/advanced.js';
import { applyConditionalFormat } from './tools/conditional.js';

import { TOOL_ANNOTATIONS } from './constants.js';
import * as schemas from './schemas/index.js';
import { setAllowedDirectories } from './tools/helpers.js';

// User configuration storage
interface UserConfig {
  createBackupByDefault?: boolean;
  defaultResponseFormat?: 'json' | 'markdown';
  allowedDirectories?: string[];
}

let userConfig: UserConfig = {
  createBackupByDefault: false,
  defaultResponseFormat: 'json',
  allowedDirectories: [],
};

// Schema mapping for validation
const toolSchemas: Record<string, any> = {
  excel_read_workbook: schemas.readWorkbookSchema,
  excel_read_sheet: schemas.readSheetSchema,
  excel_read_range: schemas.readRangeSchema,
  excel_get_cell: schemas.getCellSchema,
  excel_get_formula: schemas.getFormulaSchema,
  excel_write_workbook: schemas.writeWorkbookSchema,
  excel_update_cell: schemas.updateCellSchema,
  excel_write_range: schemas.writeRangeSchema,
  excel_add_row: schemas.addRowSchema,
  excel_set_formula: schemas.setFormulaSchema,
  excel_format_cell: schemas.formatCellSchema,
  excel_set_column_width: schemas.setColumnWidthSchema,
  excel_set_row_height: schemas.setRowHeightSchema,
  excel_merge_cells: schemas.mergeCellsSchema,
  excel_create_sheet: schemas.createSheetSchema,
  excel_delete_sheet: schemas.deleteSheetSchema,
  excel_rename_sheet: schemas.renameSheetSchema,
  excel_duplicate_sheet: schemas.duplicateSheetSchema,
  excel_delete_rows: schemas.deleteRowsSchema,
  excel_delete_columns: schemas.deleteColumnsSchema,
  excel_copy_range: schemas.copyRangeSchema,
  excel_search_value: schemas.searchValueSchema,
  excel_filter_rows: schemas.filterRowsSchema,
  excel_create_chart: schemas.createChartSchema,
  excel_create_pivot_table: schemas.createPivotTableSchema,
  excel_create_table: schemas.createTableSchema,
  excel_validate_formula_syntax: schemas.validateFormulaSyntaxSchema,
  excel_validate_range: schemas.validateExcelRangeSchema,
  excel_get_data_validation_info: schemas.getDataValidationInfoSchema,
  excel_insert_rows: schemas.insertRowsSchema,
  excel_insert_columns: schemas.insertColumnsSchema,
  excel_unmerge_cells: schemas.unmergeCellsSchema,
  excel_get_merged_cells: schemas.getMergedCellsSchema,
  excel_apply_conditional_format: schemas.applyConditionalFormatSchema,
};

// Create server instance
const server = new Server(
  {
    name: 'excel-mcp-server',
    version: '2.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Handle initialization
server.setRequestHandler(InitializeRequestSchema, async () => {
  return {
    protocolVersion: '2024-11-05',
    capabilities: {
      tools: {},
    },
    serverInfo: {
      name: 'excel-mcp-server',
      version: '2.0.0',
    },
  };
});

// List all available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      // READ OPERATIONS
      {
        name: 'excel_read_workbook',
        description: 'List all sheets and metadata of an Excel workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_read_sheet',
        description: 'Read complete data from a sheet (with optional range)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Optional range (e.g., A1:D10)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_read_range',
        description: 'Read a specific range of cells',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to read (e.g., A1:D10)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'range'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_get_cell',
        description: 'Read value from a specific cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'cellAddress'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_get_formula',
        description: 'Read the formula from a specific cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'cellAddress'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // WRITE OPERATIONS
      {
        name: 'excel_write_workbook',
        description: 'Create a new Excel file with data',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path for the new Excel file' },
            sheetName: { type: 'string', description: 'Name for the sheet', default: 'Sheet1' },
            data: { type: 'array', description: '2D array of data to write' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'data'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_update_cell',
        description: 'Update value of a specific cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            value: { description: 'Value to write' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'cellAddress', 'value'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_write_range',
        description: 'Write multiple cells simultaneously',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to write (e.g., A1:D10)' },
            data: { type: 'array', description: '2D array of data to write' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range', 'data'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_add_row',
        description: 'Add a row at the end of the sheet',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            data: { type: 'array', description: 'Array of values for the new row' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'data'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_set_formula',
        description: 'Set or modify a formula in a cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            formula: { type: 'string', description: 'Excel formula (without = sign)' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'cellAddress', 'formula'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // FORMAT OPERATIONS
      {
        name: 'excel_format_cell',
        description: 'Change cell formatting (color, font, borders, alignment)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            format: { type: 'object', description: 'Format options' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'cellAddress', 'format'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_set_column_width',
        description: 'Adjust width of a column',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            column: { description: 'Column letter (A) or number (1)' },
            width: { type: 'number', description: 'Width in Excel units' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'column', 'width'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_set_row_height',
        description: 'Adjust height of a row',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            row: { type: 'number', description: 'Row number (1-based)' },
            height: { type: 'number', description: 'Height in points' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'row', 'height'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_merge_cells',
        description: 'Merge cells in a range',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to merge (e.g., A1:D1)' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // SHEET MANAGEMENT
      {
        name: 'excel_create_sheet',
        description: 'Create a new sheet in the workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name for the new sheet' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_delete_sheet',
        description: 'Delete a sheet from the workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet to delete' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: { ...TOOL_ANNOTATIONS.DESTRUCTIVE, destructiveHint: 'true' },
      },
      {
        name: 'excel_rename_sheet',
        description: 'Rename a sheet',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            oldName: { type: 'string', description: 'Current sheet name' },
            newName: { type: 'string', description: 'New sheet name' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'oldName', 'newName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_duplicate_sheet',
        description: 'Duplicate a complete sheet',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sourceSheetName: { type: 'string', description: 'Name of sheet to duplicate' },
            newSheetName: { type: 'string', description: 'Name for duplicated sheet' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sourceSheetName', 'newSheetName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // OPERATIONS
      {
        name: 'excel_delete_rows',
        description: 'Delete specific rows',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            startRow: { type: 'number', description: 'Starting row number (1-based)' },
            count: { type: 'number', description: 'Number of rows to delete' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'startRow', 'count'],
        },
        annotations: { ...TOOL_ANNOTATIONS.DESTRUCTIVE, destructiveHint: 'true' },
      },
      {
        name: 'excel_delete_columns',
        description: 'Delete specific columns',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            startColumn: { description: 'Starting column (letter or number)' },
            count: { type: 'number', description: 'Number of columns to delete' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'startColumn', 'count'],
        },
        annotations: { ...TOOL_ANNOTATIONS.DESTRUCTIVE, destructiveHint: 'true' },
      },
      {
        name: 'excel_copy_range',
        description: 'Copy range to another location',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sourceSheetName: { type: 'string', description: 'Source sheet name' },
            sourceRange: { type: 'string', description: 'Source range (e.g., A1:D10)' },
            targetSheetName: { type: 'string', description: 'Target sheet name' },
            targetCell: { type: 'string', description: 'Top-left cell of destination' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sourceSheetName', 'sourceRange', 'targetSheetName', 'targetCell'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // ANALYSIS
      {
        name: 'excel_search_value',
        description: 'Search for a value in sheet/range',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            searchValue: { description: 'Value to search for' },
            range: { type: 'string', description: 'Optional range to search within' },
            caseSensitive: { type: 'boolean', default: false },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'searchValue'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_filter_rows',
        description: 'Filter rows by condition',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            column: { description: 'Column to filter by' },
            condition: { type: 'string', enum: ['equals', 'contains', 'greater_than', 'less_than', 'not_empty'] },
            value: { description: 'Value to compare against' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'column', 'condition'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // CHARTS
      {
        name: 'excel_create_chart',
        description: 'Create a chart (line, bar, column, pie, scatter, area)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            chartType: { type: 'string', enum: ['line', 'bar', 'column', 'pie', 'scatter', 'area'] },
            dataRange: { type: 'string', description: 'Range of data (e.g., A1:D10)' },
            position: { type: 'string', description: 'Position for chart (e.g., F2)' },
            title: { type: 'string', description: 'Chart title' },
            showLegend: { type: 'boolean', default: true },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'chartType', 'dataRange', 'position'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // PIVOT TABLES
      {
        name: 'excel_create_pivot_table',
        description: 'Create a pivot table for data analysis',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sourceSheetName: { type: 'string', description: 'Source sheet name' },
            sourceRange: { type: 'string', description: 'Source data range' },
            targetSheetName: { type: 'string', description: 'Target sheet for pivot table' },
            targetCell: { type: 'string', description: 'Target cell (e.g., A1)' },
            rows: { type: 'array', items: { type: 'string' }, description: 'Row fields' },
            columns: { type: 'array', items: { type: 'string' }, description: 'Column fields' },
            values: { type: 'array', description: 'Value fields with aggregation' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sourceSheetName', 'sourceRange', 'targetSheetName', 'targetCell', 'rows', 'values'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // TABLES
      {
        name: 'excel_create_table',
        description: 'Convert a range to an Excel table with formatting',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to convert (e.g., A1:D10)' },
            tableName: { type: 'string', description: 'Name for the table' },
            tableStyle: { type: 'string', default: 'TableStyleMedium2' },
            showFirstColumn: { type: 'boolean', default: false },
            showLastColumn: { type: 'boolean', default: false },
            showRowStripes: { type: 'boolean', default: true },
            showColumnStripes: { type: 'boolean', default: false },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range', 'tableName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // VALIDATION
      {
        name: 'excel_validate_formula_syntax',
        description: 'Validate Excel formula syntax without applying it',
        inputSchema: {
          type: 'object',
          properties: {
            formula: { type: 'string', description: 'Formula to validate (without = sign)' },
          },
          required: ['formula'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_validate_range',
        description: 'Validate if a range string is valid',
        inputSchema: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Range to validate (e.g., A1:D10)' },
          },
          required: ['range'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_get_data_validation_info',
        description: 'Get data validation rules for a cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'cellAddress'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // ADVANCED OPERATIONS
      {
        name: 'excel_insert_rows',
        description: 'Insert rows at a specific position',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            startRow: { type: 'number', description: 'Row number to insert at (1-based)' },
            count: { type: 'number', description: 'Number of rows to insert' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'startRow', 'count'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_insert_columns',
        description: 'Insert columns at a specific position',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            startColumn: { description: 'Column to insert at (letter or number)' },
            count: { type: 'number', description: 'Number of columns to insert' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'startColumn', 'count'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_unmerge_cells',
        description: 'Unmerge previously merged cells',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to unmerge (e.g., A1:D1)' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_get_merged_cells',
        description: 'List all merged cell ranges in a sheet',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // CONDITIONAL FORMATTING
      {
        name: 'excel_apply_conditional_format',
        description: 'Apply conditional formatting to a range',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to format (e.g., A1:D10)' },
            ruleType: { type: 'string', enum: ['cellValue', 'colorScale', 'dataBar', 'topBottom'] },
            condition: { type: 'object', description: 'Condition for cellValue type' },
            style: { type: 'object', description: 'Style to apply' },
            colorScale: { type: 'object', description: 'Color scale settings' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range', 'ruleType'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
    ],
  };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const { name, arguments: args } = request.params;

    if (!args) {
      throw new McpError(ErrorCode.InvalidParams, 'Missing arguments');
    }

    // Validate arguments with Zod schema
    const schema = toolSchemas[name];
    if (!schema) {
      throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }

    let validatedArgs: any;
    try {
      validatedArgs = schema.parse(args);
    } catch (error) {
      if (error instanceof ZodError) {
        const issues = error.errors.map(e => `${e.path.join('.')}: ${e.message}`).join(', ');
        throw new McpError(ErrorCode.InvalidParams, `Invalid arguments: ${issues}`);
      }
      throw error;
    }

    // Apply user config defaults
    if (validatedArgs.createBackup === undefined && userConfig.createBackupByDefault !== undefined) {
      validatedArgs.createBackup = userConfig.createBackupByDefault;
    }
    if (validatedArgs.responseFormat === undefined && userConfig.defaultResponseFormat !== undefined) {
      validatedArgs.responseFormat = userConfig.defaultResponseFormat;
    }

    let result: string;

    switch (name) {
      // Read operations
      case 'excel_read_workbook':
        result = await readWorkbook(validatedArgs.filePath, validatedArgs.responseFormat);
        break;
      case 'excel_read_sheet':
        result = await readSheet(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.range, validatedArgs.responseFormat);
        break;
      case 'excel_read_range':
        result = await readRange(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.range, validatedArgs.responseFormat);
        break;
      case 'excel_get_cell':
        result = await getCell(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.responseFormat);
        break;
      case 'excel_get_formula':
        result = await getFormula(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.responseFormat);
        break;

      // Write operations
      case 'excel_write_workbook':
        result = await writeWorkbook(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.data, validatedArgs.createBackup);
        break;
      case 'excel_update_cell':
        result = await updateCell(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.value, validatedArgs.createBackup);
        break;
      case 'excel_write_range':
        result = await writeRange(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.range, validatedArgs.data, validatedArgs.createBackup);
        break;
      case 'excel_add_row':
        result = await addRow(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.data, validatedArgs.createBackup);
        break;
      case 'excel_set_formula':
        result = await setFormula(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.formula, validatedArgs.createBackup);
        break;

      // Format operations
      case 'excel_format_cell':
        result = await formatCell(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.format, validatedArgs.createBackup);
        break;
      case 'excel_set_column_width':
        result = await setColumnWidth(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.column, validatedArgs.width, validatedArgs.createBackup);
        break;
      case 'excel_set_row_height':
        result = await setRowHeight(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.row, validatedArgs.height, validatedArgs.createBackup);
        break;
      case 'excel_merge_cells':
        result = await mergeCells(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.range, validatedArgs.createBackup);
        break;

      // Sheet management
      case 'excel_create_sheet':
        result = await createSheet(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.createBackup);
        break;
      case 'excel_delete_sheet':
        result = await deleteSheet(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.createBackup);
        break;
      case 'excel_rename_sheet':
        result = await renameSheet(validatedArgs.filePath, validatedArgs.oldName, validatedArgs.newName, validatedArgs.createBackup);
        break;
      case 'excel_duplicate_sheet':
        result = await duplicateSheet(validatedArgs.filePath, validatedArgs.sourceSheetName, validatedArgs.newSheetName, validatedArgs.createBackup);
        break;

      // Operations
      case 'excel_delete_rows':
        result = await deleteRows(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.startRow, validatedArgs.count, validatedArgs.createBackup);
        break;
      case 'excel_delete_columns':
        result = await deleteColumns(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.startColumn, validatedArgs.count, validatedArgs.createBackup);
        break;
      case 'excel_copy_range':
        result = await copyRange(
          validatedArgs.filePath,
          validatedArgs.sourceSheetName,
          validatedArgs.sourceRange,
          validatedArgs.targetSheetName,
          validatedArgs.targetCell,
          validatedArgs.createBackup
        );
        break;

      // Analysis
      case 'excel_search_value':
        result = await searchValue(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.searchValue,
          validatedArgs.range,
          validatedArgs.caseSensitive,
          validatedArgs.responseFormat
        );
        break;
      case 'excel_filter_rows':
        result = await filterRows(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.column,
          validatedArgs.condition,
          validatedArgs.value,
          validatedArgs.responseFormat
        );
        break;

      // Charts
      case 'excel_create_chart':
        result = await createChart(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.chartType,
          validatedArgs.dataRange,
          validatedArgs.position,
          validatedArgs.title,
          validatedArgs.showLegend,
          validatedArgs.createBackup
        );
        break;

      // Pivot tables
      case 'excel_create_pivot_table':
        result = await createPivotTable(
          validatedArgs.filePath,
          validatedArgs.sourceSheetName,
          validatedArgs.sourceRange,
          validatedArgs.targetSheetName,
          validatedArgs.targetCell,
          validatedArgs.rows,
          validatedArgs.columns,
          validatedArgs.values,
          validatedArgs.createBackup
        );
        break;

      // Tables
      case 'excel_create_table':
        result = await createTable(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.tableName,
          validatedArgs.tableStyle,
          validatedArgs.showFirstColumn,
          validatedArgs.showLastColumn,
          validatedArgs.showRowStripes,
          validatedArgs.showColumnStripes,
          validatedArgs.createBackup
        );
        break;

      // Validation
      case 'excel_validate_formula_syntax':
        result = await validateFormulaSyntax(validatedArgs.formula);
        break;

      case 'excel_validate_range':
        result = await validateExcelRange(validatedArgs.range);
        break;

      case 'excel_get_data_validation_info':
        result = await getDataValidationInfo(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.cellAddress,
          validatedArgs.responseFormat
        );
        break;

      // Advanced operations
      case 'excel_insert_rows':
        result = await insertRows(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.startRow,
          validatedArgs.count,
          validatedArgs.createBackup
        );
        break;

      case 'excel_insert_columns':
        result = await insertColumns(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.startColumn,
          validatedArgs.count,
          validatedArgs.createBackup
        );
        break;

      case 'excel_unmerge_cells':
        result = await unmergeCells(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.createBackup
        );
        break;

      case 'excel_get_merged_cells':
        result = await getMergedCells(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.responseFormat
        );
        break;

      // Conditional formatting
      case 'excel_apply_conditional_format':
        result = await applyConditionalFormat(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.ruleType,
          validatedArgs.condition,
          validatedArgs.style,
          validatedArgs.colorScale,
          validatedArgs.createBackup
        );
        break;

      default:
        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }

    return {
      content: [
        {
          type: 'text',
          text: result,
        },
      ],
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ error: errorMessage }, null, 2),
        },
      ],
      isError: true,
    };
  }
});

// Handle configuration updates via notifications
// Note: This will be called by Claude Desktop when user updates config
async function handleConfigUpdate(config: any) {
  try {
    if (config.createBackupByDefault !== undefined) {
      userConfig.createBackupByDefault = config.createBackupByDefault;
    }

    if (config.defaultResponseFormat !== undefined) {
      userConfig.defaultResponseFormat = config.defaultResponseFormat;
    }

    if (config.allowedDirectories !== undefined) {
      userConfig.allowedDirectories = Array.isArray(config.allowedDirectories)
        ? config.allowedDirectories
        : [];

      // Update allowed directories in helpers
      setAllowedDirectories(userConfig.allowedDirectories || []);
    }

    console.error('Configuration updated:', userConfig);
  } catch (error) {
    console.error('Error handling configuration:', error);
  }
}

// Set up notification handler
server.notification = async (notification: any) => {
  if (notification.method === 'notifications/configure') {
    await handleConfigUpdate(notification.params?.config || notification.params);
  }
};

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Excel MCP Server running on stdio');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
