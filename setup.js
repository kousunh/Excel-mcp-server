#!/usr/bin/env node
import { spawn } from 'child_process';
import { existsSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { platform } from 'os';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const isWindows = platform() === 'win32';
const venvPath = join(__dirname, 'venv');
const pythonPath = isWindows ? join(venvPath, 'Scripts', 'python.exe') : join(venvPath, 'bin', 'python');
const pipPath = isWindows ? join(venvPath, 'Scripts', 'pip.exe') : join(venvPath, 'bin', 'pip');

async function runCommand(command, args = [], options = {}) {
  return new Promise((resolve, reject) => {
    const proc = spawn(command, args, { stdio: 'inherit', ...options });
    proc.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(`Command failed with exit code ${code}`));
      } else {
        resolve();
      }
    });
    proc.on('error', reject);
  });
}

async function setup() {
  console.log('🚀 Setting up Excel MCP Server...\n');

  try {
    // Check if Python is installed
    try {
      await runCommand('python', ['--version']);
    } catch {
      try {
        await runCommand('python3', ['--version']);
      } catch {
        console.error('❌ Python is not installed. Please install Python 3.x first.');
        process.exit(1);
      }
    }

    // Create venv if it doesn't exist
    if (!existsSync(venvPath)) {
      console.log('📦 Creating Python virtual environment...');
      try {
        await runCommand('python', ['-m', 'venv', 'venv']);
      } catch {
        await runCommand('python3', ['-m', 'venv', 'venv']);
      }
    }

    // Install Python dependencies
    console.log('📚 Installing Python dependencies...');
    await runCommand(pipPath, ['install', '-r', 'requirements.txt']);

    console.log('\n✅ Setup completed successfully!');
    console.log('\n📖 How to use:');
    console.log('1. Start Excel and open a workbook');
    console.log('2. Run: npx excel-mcp');
    console.log('\n🔧 For Claude Desktop, add to your config:');
    console.log(JSON.stringify({
      mcpServers: {
        "excel-mcp": {
          command: "npx",
          args: ["excel-mcp"],
          env: {}
        }
      }
    }, null, 2));

  } catch (error) {
    console.error('❌ Setup failed:', error.message);
    process.exit(1);
  }
}

// Run setup if called directly
if (process.argv[1] === fileURLToPath(import.meta.url)) {
  setup();
}