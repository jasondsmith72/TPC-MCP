# Total PC Commander - Setup Script
# This script installs requirements and configures the TPC-MCP server

# Ensure script is running with administrator privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "This script requires administrator privileges. Please run as administrator."
    exit
}

$ErrorActionPreference = "Stop"

Write-Host "====== Total PC Commander (TPC-MCP) Setup ======" -ForegroundColor Cyan
Write-Host "This script will install and configure the TPC-MCP server."

# Check for Python installation
try {
    $pythonVersion = python --version
    Write-Host "Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Python not found. Please install Python 3.7 or newer and try again." -ForegroundColor Red
    Write-Host "You can download Python from https://www.python.org/downloads/" -ForegroundColor Yellow
    exit
}

# Install Python dependencies
Write-Host "`nInstalling Python dependencies..." -ForegroundColor Cyan
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if ($LASTEXITCODE -ne 0) {
    Write-Host "Error installing Python dependencies. Please check the error message above." -ForegroundColor Red
    exit
}

Write-Host "Python dependencies installed successfully." -ForegroundColor Green

# Create Claude Desktop configuration
$claudeConfigDir = "$env:APPDATA\Claude"
$claudeConfigFile = "$claudeConfigDir\claude_desktop_config.json"

# Create the Claude config directory if it doesn't exist
if (-not (Test-Path $claudeConfigDir)) {
    New-Item -ItemType Directory -Path $claudeConfigDir | Out-Null
    Write-Host "Created Claude configuration directory: $claudeConfigDir" -ForegroundColor Green
}

# Get script directory
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDir = Split-Path $scriptPath -Parent

# Generate the Claude Desktop configuration file
$claudeConfig = @{
    mcpServers = @{
        "tpc-commander" = @{
            command = "python"
            args = @($scriptDir + "\tpc_server.py")
        }
    }
}

# Convert to JSON and save
$claudeConfigJson = $claudeConfig | ConvertTo-Json -Depth 10
Set-Content -Path $claudeConfigFile -Value $claudeConfigJson

Write-Host "Created Claude Desktop configuration at: $claudeConfigFile" -ForegroundColor Green
Write-Host "Configuration contents:"
Write-Host $claudeConfigJson

# Complete
Write-Host "`n====== Setup Complete ======" -ForegroundColor Green
Write-Host "To use Total PC Commander:"
Write-Host "1. Start or restart Claude Desktop"
Write-Host "2. Check for the MCP icon in Claude Desktop (hammer icon)"
Write-Host "3. Start giving commands to control your PC"
Write-Host "`nExample commands:"
Write-Host "- 'Show me what's on my screen right now'"
Write-Host "- 'Open Chrome and navigate to github.com'"
Write-Host "- 'Show me all running processes'"
Write-Host "- 'Show me system information'"
Write-Host "`nThanks for using Total PC Commander!"
