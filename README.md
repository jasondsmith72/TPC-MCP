# Total PC Commander (TPC-MCP)

A remote PC control solution built on the Model Context Protocol (MCP) that allows for seamless remote management of Windows systems without requiring additional software installation on the target machine.

## Features

- **Remote Screen Viewing**: View the target PC's screen in real-time
- **Remote Command Execution**: Execute commands and run programs remotely
- **Process Management**: View, start and stop processes
- **File System Operations**: Browse, read, and modify files
- **System Information**: Get detailed system information
- **Zero Client Installation**: Leverages built-in Windows capabilities

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
- [Claude Desktop](https://claude.ai/download) or other MCP-compatible client

## Getting Started

1. Clone this repository to the target computer
2. Run the setup script: `.\setup.ps1`
3. Configure the MCP connection in your client
4. Connect and start controlling the PC remotely

## Usage

Example commands:
- "Show me what's on the screen right now"
- "Open Chrome and navigate to github.com"
- "Show me all running processes"
- "Kill process with ID 1234"
- "Show me system information"

## License

MIT License
