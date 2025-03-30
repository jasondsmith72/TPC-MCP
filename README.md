# Total PC Commander (TPC-MCP)

A comprehensive remote PC control solution built on the Model Context Protocol (MCP) that allows for seamless remote management of Windows systems without requiring additional software installation on the target machine.

## Features

### Real-time Screen Viewing and Control

- **Auto-refresh Screen Capture**: View the screen with automatic periodic updates
- **Window-specific Capture**: Capture only specific application windows
- **Quality Control**: Adjust image quality to balance detail and size
- **Windows Remote Assistance Integration**: Leverage built-in Windows remote control

### System Management

- **Remote Command Execution**: Execute commands and run programs remotely
- **PowerShell Script Execution**: Run complex PowerShell scripts remotely
- **Process Management**: View, start and stop processes
- **File System Operations**: Browse, read, and modify files

### Advanced Control Options

- **Enhanced Mouse Control**: Click, right-click, double-click, and drag operations
- **Keyboard Input**: Send keystrokes to applications
- **Screen Recording**: Record screen activity using built-in Windows features

### Zero Client Installation

- Leverages built-in Windows capabilities
- No third-party software required on target machine
- Works with Claude Desktop and other MCP-compatible clients

## Architecture

TPC-MCP uses the Model Context Protocol to provide a standardized interface for remote PC control:

- **MCP Server**: Runs on the target Windows PC and exposes tools for remote control
- **MCP Client**: Any MCP-compatible client (like Claude Desktop) can connect and control the PC

## Security

- All actions require explicit user confirmation
- Commands are executed with the permissions of the logged-in user
- Security logs track all remote activities

## Requirements

- Windows 10/11
- PowerShell 5.1+
- Python 3.7+
- Required Python packages: mcp 1.2+, pywin32, pillow, psutil
- [Claude Desktop](https://claude.ai/download) or other MCP-compatible client

## Getting Started

1. Clone this repository to the target computer
2. Run the setup script as administrator: `.\setup.ps1`
3. Start or restart Claude Desktop
4. Start giving commands to control your PC

## Example Commands

### Screen Viewing
- "Show me what's on my screen right now"
- "Start auto-refreshing my screen every 5 seconds with 70% quality"
- "Show me just the Chrome window"
- "Start Windows Remote Assistance so I can control this PC directly"

### System Control
- "Open Chrome and navigate to github.com"
- "Show me all running processes"
- "Kill the process with ID 1234"
- "Execute this PowerShell script: Get-Process | Sort-Object CPU -Descending | Select-Object -First 5"

### Mouse and Keyboard
- "Click at position (500, 300)"
- "Right-click at position (800, 400)"
- "Drag from (200, 200) to (400, 500)"
- "Send these keystrokes to the active window: Hello, world!"

### File Operations
- "Show me what's in C:\\Users\\username\\Documents"
- "Read the file C:\\path\\to\\file.txt"

## Technical Details

For detailed technical information about how TPC-MCP works, see the [TECHSPECS.md](TECHSPECS.md) document.

## License

MIT License
