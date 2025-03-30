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

# Configure Windows Remote Assistance
Write-Host "`nConfiguring Windows Remote Assistance..." -ForegroundColor Cyan
try {
    # Enable Remote Assistance
    $raPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Remote Assistance"
    if (-not (Test-Path $raPath)) {
        New-Item -Path $raPath -Force | Out-Null
    }
    Set-ItemProperty -Path $raPath -Name "fAllowToGetHelp" -Value 1 -Type DWord
    Write-Host "Windows Remote Assistance enabled." -ForegroundColor Green
    
    # Configure firewall rules
    $firewallRuleName = "Windows Remote Assistance"
    $existingRule = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue
    
    if (-not $existingRule) {
        New-NetFirewallRule -DisplayName $firewallRuleName -Direction Inbound -Program "%SystemRoot%\System32\msra.exe" -Action Allow | Out-Null
        Write-Host "Firewall rule for Windows Remote Assistance created." -ForegroundColor Green
    } else {
        Write-Host "Firewall rule for Windows Remote Assistance already exists." -ForegroundColor Green
    }
} catch {
    Write-Host "Warning: Could not configure Windows Remote Assistance. Some features may not work." -ForegroundColor Yellow
    Write-Host "Error: $_" -ForegroundColor Red
}

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

# Create shortcut for easy launch
$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "Start TPC-MCP.lnk"

try {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = "powershell.exe"
    $Shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$scriptDir\tpc_server.py`""
    $Shortcut.Description = "Start Total PC Commander MCP Server"
    $Shortcut.WorkingDirectory = $scriptDir
    $Shortcut.Save()
    Write-Host "Created desktop shortcut: $shortcutPath" -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not create desktop shortcut." -ForegroundColor Yellow
    Write-Host "Error: $_" -ForegroundColor Red
}

# Complete
Write-Host "`n====== Setup Complete ======" -ForegroundColor Green
Write-Host "To use Total PC Commander:"
Write-Host "1. Start or restart Claude Desktop"
Write-Host "2. Check for the MCP icon in Claude Desktop (hammer icon)"
Write-Host "3. Start giving commands to control your PC"
Write-Host "`nNew Features Available:"
Write-Host "- Auto-refresh screen captures (start_auto_refresh)"
Write-Host "- Window-specific captures (capture_window)"
Write-Host "- Advanced mouse control (right_click, double_click, drag_mouse)"
Write-Host "- Windows Remote Assistance integration (start_remote_assistance)"
Write-Host "- PowerShell script execution (execute_powershell)"
Write-Host "- Screen recording (start_screen_recording_ps)"
Write-Host "`nExample commands:"
Write-Host "- 'Show me what's on my screen right now'"
Write-Host "- 'Start auto-refreshing my screen every 5 seconds'"
Write-Host "- 'Open Chrome and navigate to github.com'"
Write-Host "- 'Show me all running processes'"
Write-Host "- 'Start Windows Remote Assistance so I can control this computer'"
Write-Host "`nThanks for using Total PC Commander!"
