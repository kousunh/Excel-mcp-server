@echo off
echo 🚀 Setting up Excel MCP Server...

REM Check if we're in the right directory
if not exist "package.json" (
    echo ❌ Error: package.json not found. Please run this script from the excel-mcp-server directory.
    pause
    exit /b 1
)

REM Create virtual environment
echo 📦 Creating Python virtual environment...
python -m venv venv

REM Activate virtual environment
echo ⚡ Activating virtual environment...
call venv\Scripts\activate

REM Install Python dependencies
echo 📚 Installing Python dependencies...
pip install -r requirements.txt

REM Install Node dependencies
echo 🔧 Installing Node dependencies...
npm install

echo.
echo ✅ Setup complete! Excel MCP Server is ready to use.
echo.
echo 📖 Next steps:
echo 1. Start Excel and open a workbook
echo 2. Configure your AI assistant with the server
echo.
echo 🔧 For Claude Desktop, add to %%APPDATA%%\Claude\claude_desktop_config.json:
echo.
echo {
echo   "mcpServers": {
echo     "excel-mcp": {
echo       "command": "node",
echo       "args": [
echo         "%CD%\excel-mcp.js"
echo       ],
echo       "env": {},
echo       "cwd": "%CD%"
echo     }
echo   }
echo }
echo.
echo 🔧 For Cursor IDE, add to %%USERPROFILE%%\.cursor\mcp.json:
echo.
echo {
echo   "mcpServers": {
echo     "excel-mcp": {
echo       "command": "node",
echo       "args": [
echo         "%CD%\excel-mcp.js"
echo       ],
echo       "env": {},
echo       "cwd": "%CD%"
echo     }
echo   }
echo }
echo.
pause