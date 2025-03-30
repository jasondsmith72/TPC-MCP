# Technical Specifications - Total PC Commander (TPC-MCP)

This document provides detailed technical information about the Total PC Commander MCP implementation, its architecture, and core functionality.

## 1. Architecture Overview

TPC-MCP is built on the Model Context Protocol (MCP), which provides a standardized way for AI models to interact with external systems. The architecture consists of:

- **MCP Server**: Python application that exposes Windows system functionalities as MCP tools
- **Transport Layer**: Uses stdio transport for local communication with Claude Desktop or other MCP clients
- **Client Integration**: Configured to work with Claude Desktop by default
- **Tool Framework**: Implements various system tools using the FastMCP SDK

## 2. Core Technologies

### 2.1 Python Frameworks and Libraries

- **MCP SDK**: Version 1.2.0+ for protocol implementation and server logic
- **PyWin32**: Windows API integration for screen capture, keyboard control, mouse control
- **Pillow**: Image processing for screen captures
- **psutil**: System and process monitoring
- **subprocess**: Command execution and application control

### 2.2 MCP Implementation

TPC-MCP uses the FastMCP class from the MCP Python SDK, which provides:

- Simplified tool registration via Python decorators
- Automatic tool schema generation from Python type hints
- JSON-RPC message handling
- Protocol version negotiation
- Transport management

### 2.3 Windows Integration

The server integrates with Windows using:

- **Win32 API**: Direct access to Windows functionality through PyWin32
- **Windows System Commands**: Access to system utilities via subprocess
- **Python Standard Library**: File system operations and system information

## 3. Core Functionalities

### 3.1 Screen Capture System

The screen capture system uses a multi-layered approach:

```
┌─────────────────┐     ┌────────────────┐     ┌────────────────┐
│ ImageGrab.grab()│ ──▶ │ PNG Conversion │ ──▶ │ Base64 Encoding│
└─────────────────┘     └────────────────┘     └────────────────┘
```

**Technical Details**:
- Uses PIL's ImageGrab for reliable cross-version screen capture
- Image is converted to PNG format for optimal quality/size balance
- Base64 encoding allows embedding in Markdown responses
- Returned as Markdown image for direct display in client UIs

**Implementation Notes**:
- Captures the entire screen at native resolution
- PNG compression level optimized for response time
- Typical capture-to-display latency: ~200-500ms depending on screen size
- Memory usage: Approximately (Width × Height × 4 bytes) + overhead

### 3.2 Remote Control System

The remote control system provides keyboard and mouse control capabilities:

**Keyboard Control**:
- Uses Win32 API to simulate key presses
- Supports ASCII character set by default
- Key events include both keydown and keyup events
- Small delay between keypresses ensures reliable input

**Mouse Control**:
- Uses Win32 API's SetCursorPos and mouse_event functions
- Supports absolute positioning based on screen coordinates
- Handles left mouse button clicks
- Functions properly regardless of DPI scaling settings

**Security Considerations**:
- All inputs require explicit user approval via MCP client
- Inputs are executed with the permissions of the currently logged-in user
- No persistent access - requires active MCP connection

### 3.3 Process Management

The process management system provides comprehensive process monitoring and control:

**Process Listing**:
- Uses psutil to enumerate all running processes
- Collects PID, name, memory usage, CPU usage
- Sorts by memory usage (highest first)
- Optimized for large process lists (thousands of processes)

**Process Termination**:
- Uses psutil's kill() method
- Performs proper process termination
- Handles permission errors and missing processes gracefully
- Provides detailed feedback on termination results

### 3.4 File System Operations

File system operations are implemented with security in mind:

**Directory Browsing**:
- Uses os.listdir and os.path functions
- Shows file/directory distinction, size, and modification time
- Handles permission errors gracefully
- Sorts directories first, then files alphabetically

**File Reading**:
- Supports text file reading with proper encoding handling
- Includes file size limits (1MB) to prevent excessive memory usage
- Returns content with appropriate syntax highlighting based on file extension
- Uses error='replace' for encoding to handle non-UTF-8 files gracefully

### 3.5 Command Execution

The command execution system allows controlled access to system commands:

**Implementation Details**:
- Uses subprocess.run with shell=True for full shell support
- Captures both stdout and stderr
- Implements 30-second timeout to prevent hanging
- Returns formatted output with exit code

**Security Notes**:
- All commands are executed with the permissions of the current user
- No elevation of privileges is performed
- Requires explicit user approval via MCP client

### 3.6 Application Control

The application control system provides a high-level interface for launching applications:

**Implementation Details**:
- Maps common application names to their executable names
- Uses subprocess.Popen for non-blocking execution
- Handles path resolution automatically
- Includes error handling for missing applications

**Supported Applications**:
- Common Windows applications (Notepad, Calculator, Explorer)
- Web browsers (Chrome, Edge, Firefox)
- Office applications (Word, Excel)
- Command shells (PowerShell, cmd.exe)
- Custom applications via direct executable name

## 4. Protocol Implementation

### 4.1 MCP Tool Definitions

Tools are defined using the FastMCP decorator pattern:

```python
@mcp.tool()
async def example_tool(param1: str, param2: int) -> str:
    """Tool description
    
    Args:
        param1: Parameter description
        param2: Parameter description
    """
    # Implementation
    return "Result"
```

**Key Features**:
- Automatic schema generation from type hints
- Parameter validation
- Descriptive help text from docstrings
- Async execution for non-blocking operation
- Structured return values for client display

### 4.2 Transport Configuration

TPC-MCP uses stdio transport for local machine operation:

**Technical Details**:
- JSON-RPC over stdin/stdout
- Non-blocking message processing
- Automatic serialization/deserialization
- Protocol error handling
- Connection lifecycle management

**Configuration**:
```python
mcp.run(transport='stdio')
```

### 4.3 Protocol Compatibility

TPC-MCP is compatible with:

- MCP Specification version 0.5.0+
- Claude Desktop and other MCP-compliant clients
- JSON-RPC 2.0

## 5. Security Model

### 5.1 Permission Model

TPC-MCP operates entirely within the permission boundary of the current user:

- No elevation of privileges
- No access to protected system resources
- Standard Windows security model applies
- All operations subject to Windows ACLs and permissions

### 5.2 Client Authorization

All operations require explicit client-side authorization:

- MCP clients (like Claude Desktop) enforce human-in-the-loop approval
- Each tool execution requires separate approval
- No persistent authorization - authorization is per-action

### 5.3 Input Validation

Comprehensive input validation is implemented:

- Path traversal protection for file operations
- File size limits for read operations
- Command timeout limits
- Type checking and validation via MCP schema

### 5.4 Error Handling

Robust error handling prevents information leakage:

- Exceptions are caught and formatted appropriately
- Error messages provide useful information without exposing internals
- Resource cleanup occurs regardless of operation success/failure

## 6. Deployment and Configuration

### 6.1 Installation Process

The setup.ps1 script performs:

- Python dependency installation
- Claude Desktop configuration
- Environment validation
- User feedback

### 6.2 Claude Desktop Integration

Integration with Claude Desktop is achieved through:

- Configuration file at %APPDATA%\Claude\claude_desktop_config.json
- Server process management via Claude Desktop
- Tool discovery via MCP protocol
- User interaction and approval workflow

### 6.3 System Requirements

Minimum system requirements:

- Windows 10 or Windows 11
- Python 3.7 or higher
- 4GB RAM (8GB recommended)
- Claude Desktop (latest version)

## 7. Performance Characteristics

### 7.1 Resource Usage

Typical resource utilization:

- Memory: 60-120MB depending on active operations
- CPU: <5% during idle, 10-30% during active operations
- Disk: Minimal (read-only access unless explicitly writing files)
- Network: None (local operation only)

### 7.2 Response Times

Typical response times:

- Simple tools (process listing, system info): <100ms
- Screen capture: 200-500ms
- Command execution: Varies by command
- File operations: Varies by file size

### 7.3 Scalability

The system is designed for single-machine control:

- One user/session at a time
- Not designed for multi-user access
- Not designed for remote network access (local machine only)

## 8. Limitations

### 8.1 Technical Limitations

- No support for DirectX/OpenGL screen capture (desktop apps only)
- Limited to ASCII character set for keyboard input
- No support for specialized keyboard keys (media keys, etc.)
- No audio capture capabilities
- Screen capture at native resolution only (no scaling options)

### 8.2 Security Limitations

- No support for UAC elevation
- No support for privileged operations
- No support for secure credential handling
- No encryption beyond standard Windows security

## 9. Future Development

### 9.1 Planned Enhancements

- Enhanced keyboard support (special keys, keyboard shortcuts)
- Window-specific screen capture
- Audio capture and control
- Advanced mouse operations (right-click, drag operations)
- Remote network operation support
- Session recording and playback

### 9.2 API Extension

The tool framework can be extended with:

- New tools via the @mcp.tool() decorator
- Enhanced functionality via Python modules
- Custom resource types
- Specialized prompt templates

## 10. Troubleshooting

### 10.1 Common Issues

- **Tool not appearing in Claude Desktop**: Check claude_desktop_config.json and Claude logs
- **Screen capture not working**: Verify Python PIL installation and permissions
- **Commands failing**: Check user permissions and command syntax
- **High latency**: Check system resources and reduce screen resolution

### 10.2 Logging

Logs are available at:
- Claude Desktop logs: %APPDATA%\Claude\logs\mcp-*.log
- Standard error output (when run manually)

## 11. Technical Support

For technical support:
- GitHub Issues: https://github.com/jasondsmith72/TPC-MCP/issues
- Documentation: GitHub repository README and this TECHSPECS document
