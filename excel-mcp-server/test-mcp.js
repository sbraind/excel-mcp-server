import { spawn } from 'child_process';

console.log('🧪 Testing MCP Server...\n');

const server = spawn('node', ['dist/index.js'], {
  stdio: ['pipe', 'pipe', 'pipe']
});

let output = '';
let hasResponse = false;

server.stdout.on('data', (data) => {
  output += data.toString();
  hasResponse = true;

  try {
    const lines = output.split('\n').filter(line => line.trim());
    lines.forEach(line => {
      const response = JSON.parse(line);
      console.log('✅ Server Response:', JSON.stringify(response, null, 2));
    });
  } catch (e) {
    // Still accumulating data
  }
});

server.stderr.on('data', (data) => {
  console.error('❌ Error:', data.toString());
});

// Send initialize request
const initRequest = {
  jsonrpc: "2.0",
  id: 1,
  method: "initialize",
  params: {
    protocolVersion: "2024-11-05",
    capabilities: {},
    clientInfo: {
      name: "test-client",
      version: "1.0.0"
    }
  }
};

console.log('📤 Sending initialize request...\n');
server.stdin.write(JSON.stringify(initRequest) + '\n');

// Wait for response
setTimeout(() => {
  if (hasResponse) {
    console.log('\n✅ MCP Server is working correctly!');
  } else {
    console.log('\n⚠️  No response received from server');
  }
  server.kill();
  process.exit(hasResponse ? 0 : 1);
}, 2000);
