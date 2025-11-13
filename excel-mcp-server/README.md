# Excel MCP Server

A powerful Model Context Protocol (MCP) server for working with Excel files using TypeScript and ExcelJS.

## Features

- **34 comprehensive tools** for Excel manipulation
- Full support for reading, writing, formatting, and analyzing Excel files
- **NEW**: Charts, Pivot Tables, Excel Tables, and Conditional Formatting
- Built with the official MCP SDK
- Type-safe with TypeScript and Zod validation
- Preserves formatting when modifying files
- Optional backup creation before modifications
- Supports both JSON and Markdown response formats

## Installation

### 🚀 Quick Installation (Recommended) - One Click!

The easiest way to install this server is using the pre-built MCPB bundle:

1. **Download** the latest `excel-mcp-server.mcpb` file from the [releases page](https://github.com/sbraind/Experimentos/releases)
2. **Double-click** the `.mcpb` file, or:
   - Open Claude Desktop
   - Go to **Settings** → **Extensions** → **Advanced Settings**
   - Click **"Install Extension..."**
   - Select the downloaded `.mcpb` file
3. **Restart** Claude Desktop
4. **Done!** No Node.js installation, no config files to edit

> **Note:** One-click installation works on Claude Desktop for macOS and Windows. All dependencies are bundled - no additional setup required!

For more details, see [BUNDLE.md](BUNDLE.md).

---

### 🛠️ Manual Installation (Advanced)

If you prefer to build from source:

#### Step 1: Build the project

```bash
cd /home/user/Experimentos/excel-mcp-server
npm install
npm run build
```

#### Step 2: Configure Claude Desktop

Add this configuration to your Claude Desktop config file:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["${__dirname}/dist/index.js"]
    }
  }
}
```

**Note**: When using the MCPB bundle or manual installation, use `${__dirname}` which automatically resolves to the server's directory. For manual installations without MCPB, you can also use absolute paths like `/home/user/Experimentos/excel-mcp-server/dist/index.js`.

#### Step 3: Restart Claude Desktop

Close and reopen Claude Desktop completely.

#### Step 4: Verify

The server should now be available in Claude. Try:
```
List the sheets in /home/user/Experimentos/excel-mcp-server/test.xlsx
```

For detailed installation instructions and troubleshooting, see [INSTALLATION.md](INSTALLATION.md).

---

## Configuration Options

The server supports several configuration options that can be set through Claude Desktop's MCP configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["${__dirname}/dist/index.js"],
      "config": {
        "createBackupByDefault": false,
        "defaultResponseFormat": "json",
        "allowedDirectories": []
      }
    }
  }
}
```

### Available Options:

- **`createBackupByDefault`** (boolean, default: `false`)
  Automatically create backup files (`.backup` extension) before modifying Excel files. When enabled, every destructive operation will create a backup unless explicitly disabled in the tool call.

- **`defaultResponseFormat`** (string: `"json"` or `"markdown"`, default: `"json"`)
  Default format for tool responses. Can be overridden per tool call with the `responseFormat` parameter.

- **`allowedDirectories`** (array of strings, default: `[]`)
  List of directories where the server is allowed to read/write Excel files. When empty, all directories are accessible. Use this to restrict file access for security:
  ```json
  "allowedDirectories": [
    "/home/user/Documents/Excel",
    "/home/user/Projects/data"
  ]
  ```
  The server will reject any file operations outside these directories.

### Input Validation

All tool inputs are validated using Zod schemas. Invalid parameters will return clear error messages indicating what's wrong:

- Cell addresses must match format `A1`, `B2`, etc.
- Ranges must match format `A1:D10`
- File paths are checked against `allowedDirectories` if configured
- Missing required parameters are reported immediately

---

## Available Tools

### 📖 Reading (5 tools)

#### 1. `excel_read_workbook`
List all sheets and metadata of an Excel workbook.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "responseFormat": "json"
}
```

#### 2. `excel_read_sheet`
Read complete data from a sheet with optional range.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "A1:D10",
  "responseFormat": "markdown"
}
```

#### 3. `excel_read_range`
Read a specific range of cells.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "B2:E20",
  "responseFormat": "json"
}
```

#### 4. `excel_get_cell`
Read value from a specific cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "A1",
  "responseFormat": "json"
}
```

#### 5. `excel_get_formula`
Read the formula from a specific cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "D5",
  "responseFormat": "json"
}
```

### ✏️ Writing (5 tools)

#### 6. `excel_write_workbook`
Create a new Excel file with data.

**Example:**
```json
{
  "filePath": "./output.xlsx",
  "sheetName": "MyData",
  "data": [
    ["Name", "Age", "City"],
    ["Alice", 30, "New York"],
    ["Bob", 25, "Los Angeles"]
  ],
  "createBackup": false
}
```

#### 7. `excel_update_cell`
Update value of a specific cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "B2",
  "value": 1500,
  "createBackup": true
}
```

#### 8. `excel_write_range`
Write multiple cells simultaneously.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "A1:C2",
  "data": [
    ["Header1", "Header2", "Header3"],
    [100, 200, 300]
  ],
  "createBackup": false
}
```

#### 9. `excel_add_row`
Add a row at the end of the sheet.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "data": ["Product X", 150, "2024-01-15"],
  "createBackup": false
}
```

#### 10. `excel_set_formula`
Set or modify a formula in a cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "D2",
  "formula": "SUM(B2:C2)",
  "createBackup": false
}
```

### 🎨 Formatting (4 tools)

#### 11. `excel_format_cell`
Change cell formatting (color, font, borders, alignment).

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "A1",
  "format": {
    "font": {
      "bold": true,
      "size": 14,
      "color": "FF0000"
    },
    "fill": {
      "type": "pattern",
      "pattern": "solid",
      "fgColor": "FFFF00"
    },
    "alignment": {
      "horizontal": "center",
      "vertical": "middle"
    }
  },
  "createBackup": false
}
```

#### 12. `excel_set_column_width`
Adjust width of a column.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "column": "A",
  "width": 20,
  "createBackup": false
}
```

#### 13. `excel_set_row_height`
Adjust height of a row.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "row": 1,
  "height": 30,
  "createBackup": false
}
```

#### 14. `excel_merge_cells`
Merge cells in a range.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "A1:D1",
  "createBackup": false
}
```

### 📑 Sheet Management (4 tools)

#### 15. `excel_create_sheet`
Create a new sheet in the workbook.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "NewSheet",
  "createBackup": false
}
```

#### 16. `excel_delete_sheet`
Delete a sheet from the workbook.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "OldSheet",
  "createBackup": true
}
```

#### 17. `excel_rename_sheet`
Rename a sheet.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "oldName": "Sheet1",
  "newName": "Sales2024",
  "createBackup": false
}
```

#### 18. `excel_duplicate_sheet`
Duplicate a complete sheet.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sourceSheetName": "Template",
  "newSheetName": "January",
  "createBackup": false
}
```

### 🔧 Operations (3 tools)

#### 19. `excel_delete_rows`
Delete specific rows.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "startRow": 5,
  "count": 3,
  "createBackup": true
}
```

#### 20. `excel_delete_columns`
Delete specific columns.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "startColumn": "C",
  "count": 2,
  "createBackup": true
}
```

#### 21. `excel_copy_range`
Copy range to another location.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sourceSheetName": "Sales",
  "sourceRange": "A1:D10",
  "targetSheetName": "Backup",
  "targetCell": "A1",
  "createBackup": false
}
```

### 📊 Analysis (2 tools)

#### 22. `excel_search_value`
Search for a value in sheet/range.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "searchValue": "Apple",
  "range": "A1:Z100",
  "caseSensitive": false,
  "responseFormat": "markdown"
}
```

#### 23. `excel_filter_rows`
Filter rows by condition.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "column": "B",
  "condition": "greater_than",
  "value": 1000,
  "responseFormat": "json"
}
```

### 📈 Charts (1 tool)

#### 24. `excel_create_chart`
Create charts from data ranges.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "chartType": "column",
  "dataRange": "A1:B10",
  "position": "D2",
  "title": "Monthly Sales",
  "showLegend": true,
  "createBackup": false
}
```

**Note**: ExcelJS has limited native chart support. This creates a chart placeholder with metadata.

### 🔄 Pivot Tables (1 tool)

#### 25. `excel_create_pivot_table`
Create pivot tables for data analysis.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sourceSheetName": "Sales",
  "sourceRange": "A1:D100",
  "targetSheetName": "Pivot",
  "targetCell": "A1",
  "rows": ["Product"],
  "columns": ["Month"],
  "values": [
    { "field": "Amount", "aggregation": "sum" }
  ],
  "createBackup": false
}
```

### 📋 Excel Tables (1 tool)

#### 26. `excel_create_table`
Convert ranges to formatted Excel tables.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Data",
  "range": "A1:D50",
  "tableName": "SalesTable",
  "tableStyle": "TableStyleMedium2",
  "showRowStripes": true,
  "createBackup": false
}
```

### ✅ Validation (3 tools)

#### 27. `excel_validate_formula_syntax`
Validate formula syntax without applying it.

**Example:**
```json
{
  "formula": "SUM(A1:A10) / COUNT(B1:B10)"
}
```

#### 28. `excel_validate_range`
Validate if a range string is valid.

**Example:**
```json
{
  "range": "A1:Z100"
}
```

#### 29. `excel_get_data_validation_info`
Get data validation rules for a cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Input",
  "cellAddress": "A1",
  "responseFormat": "json"
}
```

### 🔧 Advanced Operations (4 tools)

#### 30. `excel_insert_rows`
Insert rows at a specific position.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "startRow": 5,
  "count": 3,
  "createBackup": false
}
```

#### 31. `excel_insert_columns`
Insert columns at a specific position.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "startColumn": "C",
  "count": 2,
  "createBackup": false
}
```

#### 32. `excel_unmerge_cells`
Unmerge previously merged cells.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Report",
  "range": "A1:D1",
  "createBackup": false
}
```

#### 33. `excel_get_merged_cells`
List all merged cell ranges in a sheet.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Report",
  "responseFormat": "markdown"
}
```

### 🎨 Conditional Formatting (1 tool)

#### 34. `excel_apply_conditional_format`
Apply conditional formatting to ranges.

**Example (Cell Value):**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "B2:B100",
  "ruleType": "cellValue",
  "condition": {
    "operator": "greaterThan",
    "value": 1000
  },
  "style": {
    "fill": {
      "type": "pattern",
      "pattern": "solid",
      "fgColor": "FF00FF00"
    },
    "font": {
      "bold": true
    }
  },
  "createBackup": false
}
```

**Example (Color Scale):**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "C2:C100",
  "ruleType": "colorScale",
  "colorScale": {
    "minColor": "FFFF0000",
    "maxColor": "FF00FF00"
  },
  "createBackup": false
}
```

## Development

### Build
```bash
npm run build
```

### Watch mode
```bash
npm run watch
```

### Run
```bash
npm start
```

## Error Handling

All tools include robust error handling and will return descriptive error messages for:
- File not found
- Sheet not found
- Invalid cell addresses or ranges
- Invalid formatting options
- Write errors

## Features

### Backup Support
Most write operations support an optional `createBackup` parameter. When set to `true`, a backup of the original file will be created with a `.backup` extension before modifications.

### Response Formats
Read operations support both `json` and `markdown` response formats:
- **JSON**: Structured data, ideal for programmatic processing
- **Markdown**: Human-readable tables and formatted output

### Data Preview
When reading large datasets, the markdown format automatically shows a preview of the first 100 rows.

## Dependencies

- `@modelcontextprotocol/sdk` - Official MCP SDK
- `exceljs` - Excel file manipulation
- `zod` - Schema validation
- `typescript` - Type safety

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
