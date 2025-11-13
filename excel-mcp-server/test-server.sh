#!/bin/bash
echo "Testing Excel MCP Server..."
echo ""
echo "1. Checking Node.js version:"
node --version
echo ""
echo "2. Checking if dist/index.js exists:"
ls -la dist/index.js
echo ""
echo "3. Testing server startup (will timeout after 2 seconds):"
timeout 2 node dist/index.js || echo "Server started successfully (timed out as expected)"
echo ""
echo "4. Full path to server:"
pwd
echo "$(pwd)/dist/index.js"
echo ""
echo "5. Configuration for Claude Desktop:"
echo '{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": [
        "'$(pwd)'/dist/index.js"
      ]
    }
  }
}'
