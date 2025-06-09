#!/usr/bin/env node
import { spawn } from 'child_process';
import { existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { platform } from 'os';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const rootDir = join(__dirname, '..');

const isWindows = platform() === 'win32';
const venvPath = join(rootDir, 'venv');
const pythonPath = isWindows ? join(venvPath, 'Scripts', 'python.exe') : join(venvPath, 'bin', 'python');

// Check if venv exists, if not run setup
if (!existsSync(venvPath)) {
  console.log('🔧 First time setup required. Installing dependencies...\n');
  const setup = spawn('node', [join(rootDir, 'setup.js')], { stdio: 'inherit' });
  setup.on('close', (code) => {
    if (code === 0) {
      startServer();
    } else {
      process.exit(code);
    }
  });
} else {
  startServer();
}

function startServer() {
  // Set environment variable to use venv Python
  process.env.EXCEL_MCP_PYTHON = pythonPath;
  
  // Check if running from dist (bundled) or src
  const distPath = join(rootDir, 'excel-mcp.js');
  const srcPath = join(rootDir, 'src', 'index.js');
  
  const serverPath = existsSync(distPath) ? distPath : srcPath;
  
  if (!existsSync(serverPath)) {
    console.error('❌ Server file not found. Please run npm run build first.');
    process.exit(1);
  }
  
  // Start the server
  const server = spawn('node', [serverPath], { 
    stdio: 'inherit',
    env: { ...process.env, EXCEL_MCP_PYTHON: pythonPath }
  });
  
  server.on('error', (err) => {
    console.error('Failed to start server:', err);
    process.exit(1);
  });
  
  process.on('SIGINT', () => {
    server.kill('SIGINT');
    process.exit(0);
  });
}