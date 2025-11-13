# MCP Bundle Distribution Guide

This document explains how to create and distribute the Excel MCP Server as an MCP Bundle (`.mcpb` file) for easy installation in Claude Desktop and other MCP-compatible applications.

## What is an MCP Bundle?

An MCP Bundle is a standardized format for packaging Model Context Protocol servers, similar to Chrome extensions (`.crx`) or VS Code extensions (`.vsix`). It allows for single-click installation without manual configuration.

## Prerequisites

Install the MCP Bundle CLI tool globally:

```bash
npm install -g @anthropic-ai/mcpb
```

Verify installation:
```bash
mcpb --version
```

## Creating a Bundle

### Step 1: Ensure the project is built

```bash
cd /home/user/Experimentos/excel-mcp-server
npm install
npm run build
```

### Step 2: Verify manifest.json

The `manifest.json` file is already configured with:
- All 34 tools declared
- Server configuration for Node.js
- User-configurable options
- Rich metadata (author, repository, keywords, etc.)

### Step 3: Create the bundle

Using npm script (recommended):
```bash
npm run pack:mcpb
```

Or using mcpb CLI directly:
```bash
mcpb pack .
```

This will create a file named `excel-mcp-server-2.0.0.mcpb` in the current directory.

## Bundle Structure

The `.mcpb` file is a ZIP archive containing:

```
excel-mcp-server-2.0.0.mcpb
├── manifest.json          # MCP Bundle manifest
├── package.json           # NPM package configuration
├── dist/                  # Compiled TypeScript code
│   ├── index.js
│   ├── schemas/
│   └── tools/
├── node_modules/          # Bundled dependencies
└── README.md              # Documentation
```

## Installing the Bundle

### Method 1: Single-Click Installation (Future)

Once MCP-compatible applications support bundle installation, users will be able to:
1. Download the `.mcpb` file
2. Double-click or drag-and-drop into the application
3. The server will be automatically configured and ready to use

### Method 2: Manual Installation via Claude Desktop

For now, users can extract and configure manually:

1. **Extract the bundle:**
   ```bash
   unzip excel-mcp-server-2.0.0.mcpb -d excel-mcp-server
   ```

2. **Configure Claude Desktop:**

   Add to `claude_desktop_config.json`:
   ```json
   {
     "mcpServers": {
       "excel": {
         "command": "node",
         "args": ["/path/to/excel-mcp-server/dist/index.js"]
       }
     }
   }
   ```

3. **Restart Claude Desktop**

## Distribution

### Option 1: GitHub Releases

1. Create a new release on GitHub
2. Upload the `.mcpb` file as a release asset
3. Users can download directly from the releases page

Example:
```bash
gh release create v2.0.0 \
  excel-mcp-server-2.0.0.mcpb \
  --title "Excel MCP Server v2.0.0" \
  --notes "Release notes here"
```

### Option 2: NPM Package

The server can also be installed via npm:
```bash
npm install -g excel-mcp-server
```

Then configure Claude Desktop to use the global installation.

### Option 3: Direct Download

Host the `.mcpb` file on a web server or CDN for direct download.

## Manifest Configuration

The `manifest.json` includes:

- **34 declared tools** for Excel operations
- **User configuration options:**
  - `createBackupByDefault`: Auto-backup before modifications
  - `defaultResponseFormat`: JSON or Markdown responses
- **Compatibility requirements:** Node.js >=18.0.0
- **Rich metadata:** Author, repository, keywords, license

## Updating the Bundle

When releasing a new version:

1. Update version in both `package.json` and `manifest.json`
2. Rebuild the project: `npm run build`
3. Create new bundle: `npm run pack:mcpb`
4. Test the bundle
5. Create GitHub release with the new `.mcpb` file

## Verification

After creating a bundle, verify its contents:

```bash
unzip -l excel-mcp-server-2.0.0.mcpb
```

Check that it includes:
- ✅ manifest.json
- ✅ dist/ directory with compiled code
- ✅ node_modules/ with dependencies
- ✅ package.json
- ✅ README.md

## Why Use MCP Bundles?

### For Users:
- **Easy installation:** One file, no manual configuration
- **Self-contained:** All dependencies included
- **Version management:** Clear version tracking
- **Consistency:** Same installation experience across applications

### For Developers:
- **Standardized distribution:** Follow ecosystem best practices
- **Better discoverability:** Rich metadata for searching
- **Professional packaging:** Production-ready format
- **Future-proof:** Compatible with emerging MCP ecosystem

## Resources

- [MCP Bundle Specification](https://github.com/modelcontextprotocol/mcpb)
- [MCP Documentation](https://modelcontextprotocol.io)
- [Claude Desktop](https://claude.ai/download)

## Troubleshooting

### Bundle creation fails

- Ensure `manifest.json` is valid JSON
- Check that `dist/` directory exists and contains compiled code
- Verify all required dependencies are installed

### Bundle is too large

- Remove unnecessary files before packing
- Use `npm install --production` to exclude dev dependencies
- Consider using `.mcpbignore` file (similar to `.gitignore`)

### Server doesn't start after installation

- Verify Node.js version: `node --version` (must be >=18.0.0)
- Check manifest.json server configuration
- Review Claude Desktop logs for errors

## Support

For issues or questions:
- GitHub Issues: https://github.com/sbraind/Experimentos/issues
- Documentation: https://github.com/sbraind/Experimentos/tree/main/excel-mcp-server
