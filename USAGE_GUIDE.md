# Total PC Commander - Usage Guide

This guide provides detailed instructions on how to use Total PC Commander's features effectively with Claude Desktop.

## Table of Contents

1. [Installation and Setup](#installation-and-setup)
2. [Basic Usage](#basic-usage)
3. [Screen Viewing Features](#screen-viewing-features)
4. [Remote Control Features](#remote-control-features)
5. [System Management](#system-management)
6. [File Operations](#file-operations)
7. [Advanced Features](#advanced-features)
8. [Troubleshooting](#troubleshooting)

## Installation and Setup

### Prerequisites

Before installing, ensure you have:

- Windows 10 or 11
- Python 3.7 or higher installed
- Claude Desktop application installed
- Administrator privileges for setup

### Installation Steps

1. **Clone the Repository**:
   ```
   git clone https://github.com/jasondsmith72/TPC-MCP.git
   cd TPC-MCP
   ```

2. **Run the Setup Script**:
   Right-click on `setup.ps1` and select "Run with PowerShell" or open an administrator PowerShell window and run:
   ```powershell
   .\setup.ps1
   ```

3. **Restart Claude Desktop**:
   After the setup completes, close and restart Claude Desktop.

4. **Verify Installation**:
   In Claude Desktop, you should see the hammer icon indicating MCP tools are available. Click it to view the available tools.

## Basic Usage

The TPC-MCP server communicates with Claude Desktop to execute commands on your computer. Here's how to use it:

1. **Start a Conversation**:
   Begin by asking Claude about your PC or requesting an action.

2. **Review and Approve Commands**:
   When Claude needs to execute a tool, it will ask for your permission. Review what the tool will do before approving.

3. **View Results**:
   Claude will display the results of the command, including any captured screenshots or command output.

## Screen Viewing Features

### Basic Screenshot

To take a basic screenshot:
```
Show me what's on my screen right now
```

Claude will use the `capture_screen` tool to take a screenshot and display it in the chat.

### Window-Specific Capture

To capture a specific window:
```
Show me just the Chrome window
```
or
```
Capture the window titled "Document - Word"
```

This will use the `capture_window` tool to take a screenshot of just that application window.

### Auto-Refresh Screen Capture

To start automatic screen refreshes:
```
Start auto-refreshing my screen every 5 seconds
```
or with quality settings:
```
Start auto-refreshing my screen with 3 second interval and 60% quality
```

To view the latest auto-refresh screenshot:
```
Show me the latest screenshot
```

To stop auto-refreshing:
```
Stop auto-refreshing my screen
```

### Windows Remote Assistance

To use Windows' built-in Remote Assistance feature:
```
Start Windows Remote Assistance
```

This will launch Windows Remote Assistance in "offer help" mode, allowing you to establish a direct connection.

## Remote Control Features

### Mouse Control

Basic click:
```
Click at position (500, 300)
```

Right-click:
```
Right-click at position (800, 400)
```

Double-click:
```
Double-click at position (500, 300)
```

Move the mouse without clicking:
```
Move the mouse to position (500, 300)
```

Drag and drop:
```
Drag from position (200, 200) to (400, 500)
```

### Keyboard Control

Send keystrokes to the active window:
```
Type "Hello, world!" in the active window
```
or
```
Send these keystrokes: Hello, world!
```

## System Management

### Process Management

List running processes:
```
Show me all running processes
```
or
```
What are the top memory-consuming processes right now?
```

Kill a process:
```
Kill the process with ID 1234
```
or
```
Terminate the chrome.exe process
```

### Application Control

Open an application:
```
Open Notepad
```
or
```
Launch Chrome and navigate to github.com
```

### Command Execution

Execute a command:
```
Run this command: ipconfig /all
```

Execute a PowerShell script:
```
Run this PowerShell script:
Get-Process | Sort-Object CPU -Descending | Select-Object -First 5
```

### System Information

Get system information:
```
Show me system information
```
or
```
What are my computer's specs?
```

## File Operations

### Browsing Directories

List directory contents:
```
Show me what's in C:\Users\username\Documents
```
or
```
List the files in my Downloads folder
```

### Reading Files

Read a text file:
```
Read the file C:\path\to\file.txt
```
or
```
Show me the contents of my README.md file
```

## Advanced Features

### Screen Recording

Start a screen recording:
```
Start recording my screen
```

This will use Windows' built-in screen recording capabilities.

### Remote Assistance Setup

Configure Remote Assistance settings:
```
Configure Remote Assistance for easy access
```

## Troubleshooting

### Common Issues

#### Claude doesn't see the tools
- Restart Claude Desktop
- Check if the configuration file was created properly
- Look for any errors in the Claude logs

#### Screen capture shows blank or black screens
- Make sure you're not trying to capture a DRM-protected content
- Try using Window-specific capture instead of full screen

#### Commands fail to execute
- Check if you're running as an administrator for privileged commands
- Verify the syntax of your commands
- Check Windows permissions

### Logging

View log files:
```
Show me the TPC-MCP logs
```

### Getting Help

For more assistance, submit an issue on the GitHub repository: https://github.com/jasondsmith72/TPC-MCP/issues
