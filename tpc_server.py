#!/usr/bin/env python3
"""
Total PC Commander - MCP Server
Remote PC control solution using Model Context Protocol
"""

import os
import base64
import io
import sys
import psutil
import time
import subprocess
from datetime import datetime
from typing import List, Dict, Any, Optional
from pathlib import Path

import win32gui
import win32ui
import win32con
import win32api
from PIL import Image, ImageGrab

from mcp.server.fastmcp import FastMCP

# Initialize FastMCP server
mcp = FastMCP("tpc-server")

# ----- Screen Capture Tools ----- #

@mcp.tool()
async def capture_screen() -> str:
    """Capture the current screen and return as base64 encoded image.
    
    Returns a base64-encoded PNG image of the current screen.
    """
    try:
        # Capture entire screen using PIL
        screenshot = ImageGrab.grab()
        
        # Convert to base64
        buffer = io.BytesIO()
        screenshot.save(buffer, format="PNG")
        img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # Return as Markdown image format with base64 data
        return f"![Screenshot](data:image/png;base64,{img_str})"
    except Exception as e:
        return f"Error capturing screen: {str(e)}"

@mcp.tool()
async def get_active_window_info() -> str:
    """Get information about the currently active window.
    
    Returns details about the foreground window, including title and position.
    """
    try:
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect
        width = right - left
        height = bottom - top
        
        return (
            f"Active Window Information:\n"
            f"- Title: {title}\n"
            f"- Handle: {hwnd}\n"
            f"- Position: Left={left}, Top={top}, Right={right}, Bottom={bottom}\n"
            f"- Size: {width}x{height}"
        )
    except Exception as e:
        return f"Error getting active window info: {str(e)}"

# ----- Command Execution Tools ----- #

@mcp.tool()
async def execute_command(command: str) -> str:
    """Execute a Windows command.
    
    Args:
        command: Command to execute (e.g., 'dir', 'ipconfig', etc.)
    
    Returns the output of the command.
    """
    try:
        # Execute the command and capture output
        result = subprocess.run(
            command, 
            shell=True, 
            capture_output=True, 
            text=True,
            timeout=30
        )
        
        # Format the response
        output = result.stdout or result.stderr
        exit_code = result.returncode
        
        return f"Command executed with exit code {exit_code}:\n\n```\n{output}\n```"
    except subprocess.TimeoutExpired:
        return "Command timed out after 30 seconds"
    except Exception as e:
        return f"Error executing command: {str(e)}"

@mcp.tool()
async def open_application(app_name: str) -> str:
    """Open an application.
    
    Args:
        app_name: Name of the application to open (e.g., 'notepad', 'chrome')
    
    Returns confirmation that the application was opened.
    """
    try:
        # Common application mappings
        app_mappings = {
            "notepad": "notepad.exe",
            "chrome": "chrome.exe",
            "edge": "msedge.exe",
            "firefox": "firefox.exe",
            "explorer": "explorer.exe",
            "calc": "calc.exe",
            "calculator": "calc.exe",
            "word": "winword.exe",
            "excel": "excel.exe",
            "powershell": "powershell.exe",
            "cmd": "cmd.exe"
        }
        
        # Check if we have a mapping, otherwise use the name directly
        executable = app_mappings.get(app_name.lower(), app_name)
        
        # Run the application
        subprocess.Popen(executable)
        
        return f"Application '{app_name}' ({executable}) has been opened."
    except Exception as e:
        return f"Error opening application '{app_name}': {str(e)}"

# ----- Process Management Tools ----- #

@mcp.tool()
async def list_processes() -> str:
    """List all running processes.
    
    Returns a list of running processes with their PIDs, names, and memory usage.
    """
    try:
        processes = []
        for proc in psutil.process_iter(['pid', 'name', 'memory_info', 'cpu_percent']):
            try:
                proc_info = proc.info
                memory_mb = proc_info['memory_info'].rss / (1024 * 1024) if proc_info['memory_info'] else 0
                processes.append({
                    'pid': proc_info['pid'],
                    'name': proc_info['name'],
                    'memory_mb': round(memory_mb, 2),
                    'cpu_percent': proc_info['cpu_percent']
                })
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        # Sort by memory usage (highest first)
        processes.sort(key=lambda x: x['memory_mb'], reverse=True)
        
        # Format as text table (top 20 by memory)
        result = "PID\tName\tMemory (MB)\tCPU %\n"
        result += "-" * 50 + "\n"
        
        for proc in processes[:20]:
            result += f"{proc['pid']}\t{proc['name']}\t{proc['memory_mb']}\t{proc['cpu_percent']}\n"
        
        return f"Top 20 processes by memory usage:\n\n```\n{result}\n```"
    except Exception as e:
        return f"Error listing processes: {str(e)}"

@mcp.tool()
async def kill_process(pid: int) -> str:
    """Kill a process by its PID.
    
    Args:
        pid: Process ID to kill
    
    Returns confirmation that the process was killed.
    """
    try:
        # Get process info before killing
        process = psutil.Process(pid)
        process_name = process.name()
        
        # Kill the process
        process.kill()
        
        return f"Process '{process_name}' (PID: {pid}) has been terminated."
    except psutil.NoSuchProcess:
        return f"No process found with PID {pid}."
    except psutil.AccessDenied:
        return f"Access denied when trying to kill process with PID {pid}. Try running with administrator privileges."
    except Exception as e:
        return f"Error killing process: {str(e)}"

# ----- System Information Tools ----- #

@mcp.tool()
async def get_system_info() -> str:
    """Get detailed system information.
    
    Returns information about the system, including OS, CPU, memory, and disk usage.
    """
    try:
        # OS information
        uname = os.uname() if hasattr(os, 'uname') else platform.uname()
        system_info = f"System: {uname.system}\n"
        system_info += f"Node Name: {uname.node}\n"
        system_info += f"Release: {uname.release}\n"
        system_info += f"Version: {uname.version}\n"
        system_info += f"Machine: {uname.machine}\n\n"
        
        # CPU information
        system_info += f"CPU Count (Logical): {psutil.cpu_count(logical=True)}\n"
        system_info += f"CPU Count (Physical): {psutil.cpu_count(logical=False)}\n"
        system_info += f"CPU Usage: {psutil.cpu_percent(interval=1)}%\n\n"
        
        # Memory information
        memory = psutil.virtual_memory()
        system_info += f"Total Memory: {memory.total / (1024**3):.2f} GB\n"
        system_info += f"Available Memory: {memory.available / (1024**3):.2f} GB\n"
        system_info += f"Used Memory: {memory.used / (1024**3):.2f} GB ({memory.percent}%)\n\n"
        
        # Disk information
        disk = psutil.disk_usage('/')
        system_info += f"Total Disk Space: {disk.total / (1024**3):.2f} GB\n"
        system_info += f"Used Disk Space: {disk.used / (1024**3):.2f} GB ({disk.percent}%)\n"
        system_info += f"Free Disk Space: {disk.free / (1024**3):.2f} GB\n"
        
        return f"System Information:\n\n```\n{system_info}\n```"
    except Exception as e:
        return f"Error retrieving system information: {str(e)}"

# ----- File System Tools ----- #

@mcp.tool()
async def list_directory(path: str = ".") -> str:
    """List files and directories at the specified path.
    
    Args:
        path: Directory path to list (default: current directory)
    
    Returns a list of files and directories at the specified path.
    """
    try:
        # Expand user directory if needed
        full_path = os.path.expanduser(path)
        
        # Check if path exists
        if not os.path.exists(full_path):
            return f"Path does not exist: {full_path}"
        
        # Check if path is a directory
        if not os.path.isdir(full_path):
            return f"Path is not a directory: {full_path}"
        
        # List contents
        contents = []
        for item in os.listdir(full_path):
            item_path = os.path.join(full_path, item)
            is_dir = os.path.isdir(item_path)
            size = os.path.getsize(item_path) if os.path.isfile(item_path) else 0
            modified = datetime.fromtimestamp(os.path.getmtime(item_path)).strftime('%Y-%m-%d %H:%M:%S')
            
            contents.append({
                'name': item,
                'type': 'Directory' if is_dir else 'File',
                'size': size,
                'modified': modified
            })
        
        # Sort by type (directories first) then by name
        contents.sort(key=lambda x: (0 if x['type'] == 'Directory' else 1, x['name']))
        
        # Format as text table
        result = f"Contents of {full_path}:\n\n"
        result += "Name\tType\tSize\tModified\n"
        result += "-" * 80 + "\n"
        
        for item in contents:
            size_str = f"{item['size'] / 1024:.2f} KB" if item['type'] == 'File' else ""
            result += f"{item['name']}\t{item['type']}\t{size_str}\t{item['modified']}\n"
        
        return f"```\n{result}\n```"
    except Exception as e:
        return f"Error listing directory: {str(e)}"

@mcp.tool()
async def read_file(path: str) -> str:
    """Read the contents of a text file.
    
    Args:
        path: Path to the file to read
    
    Returns the contents of the file.
    """
    try:
        # Expand user directory if needed
        full_path = os.path.expanduser(path)
        
        # Check if file exists
        if not os.path.exists(full_path):
            return f"File does not exist: {full_path}"
        
        # Check if file is too large
        if os.path.getsize(full_path) > 1024 * 1024:  # 1 MB limit
            return f"File is too large to read directly: {full_path} ({os.path.getsize(full_path) / (1024*1024):.2f} MB)"
        
        # Attempt to read the file
        with open(full_path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        
        # Get file extension for syntax highlighting
        _, ext = os.path.splitext(full_path)
        ext = ext.lstrip('.')
        
        return f"Contents of {full_path}:\n\n```{ext}\n{content}\n```"
    except Exception as e:
        return f"Error reading file: {str(e)}"

# ----- Windows Control Tools ----- #

@mcp.tool()
async def send_keystrokes(keys: str) -> str:
    """Send keystrokes to the active window.
    
    Args:
        keys: Keystrokes to send (e.g., 'Hello, world!')
    
    Returns confirmation that the keystrokes were sent.
    """
    try:
        # This is a simplified version - in a real implementation, 
        # you would use a library like pyautogui or win32com
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0:
            return "No window is currently active."
        
        # Use the Windows API to send keystrokes
        for c in keys:
            win32api.keybd_event(ord(c), 0, 0, 0)
            win32api.keybd_event(ord(c), 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.01)
        
        return f"Sent keystrokes: '{keys}' to the active window."
    except Exception as e:
        return f"Error sending keystrokes: {str(e)}"

@mcp.tool()
async def click_at_position(x: int, y: int) -> str:
    """Perform a mouse click at the specified screen coordinates.
    
    Args:
        x: X coordinate
        y: Y coordinate
    
    Returns confirmation that the click was performed.
    """
    try:
        # Move mouse to position
        win32api.SetCursorPos((x, y))
        
        # Perform click
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
        
        return f"Clicked at position ({x}, {y})."
    except Exception as e:
        return f"Error clicking at position: {str(e)}"

# ----- Main Function ----- #

if __name__ == "__main__":
    # Initialize and run the server
    print("Starting Total PC Commander MCP Server...", file=sys.stderr)
    mcp.run(transport='stdio')
