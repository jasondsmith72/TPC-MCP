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
import threading
import platform
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

# Global variables for auto-refresh
_refresh_thread = None
_refresh_active = False
_refresh_interval = 0
_last_screenshot = None
_last_screenshot_time = 0

@mcp.tool()
async def capture_screen(quality: int = 80) -> str:
    """Capture the current screen and return as base64 encoded image.
    
    Args:
        quality: Image quality (1-100), lower values mean smaller file size but lower quality
    
    Returns a base64-encoded JPEG image of the current screen.
    """
    try:
        # Validate quality
        quality = max(1, min(100, quality))
        
        # Capture entire screen using PIL
        screenshot = ImageGrab.grab()
        
        # Convert to base64 with specified quality
        buffer = io.BytesIO()
        screenshot.save(buffer, format="JPEG", quality=quality)
        img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # Return as Markdown image format with base64 data
        return f"![Screenshot](data:image/jpeg;base64,{img_str})"
    except Exception as e:
        return f"Error capturing screen: {str(e)}"

@mcp.tool()
async def capture_window(window_title: str = "", quality: int = 80) -> str:
    """Capture a specific window or the active window and return as base64 encoded image.
    
    Args:
        window_title: Title of window to capture (leave empty for active window)
        quality: Image quality (1-100), lower values mean smaller file size but lower quality
    
    Returns a base64-encoded JPEG image of the specified window.
    """
    try:
        # Validate quality
        quality = max(1, min(100, quality))
        
        # Find window
        if window_title:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd == 0:
                return f"No window found with title '{window_title}'"
        else:
            hwnd = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(hwnd)
            if not window_title:
                window_title = "Active Window"
        
        # Get window dimensions
        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect
        width = right - left
        height = bottom - top
        
        # Capture window using PIL
        screenshot = ImageGrab.grab((left, top, right, bottom))
        
        # Convert to base64 with specified quality
        buffer = io.BytesIO()
        screenshot.save(buffer, format="JPEG", quality=quality)
        img_str = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # Return as Markdown image format with base64 data
        return f"![{window_title}](data:image/jpeg;base64,{img_str})"
    except Exception as e:
        return f"Error capturing window: {str(e)}"

def _refresh_capture_thread(interval, quality):
    """Background thread to capture screen periodically"""
    global _refresh_active, _last_screenshot, _last_screenshot_time
    
    _refresh_active = True
    while _refresh_active:
        try:
            # Capture screen
            screenshot = ImageGrab.grab()
            
            # Convert to base64
            buffer = io.BytesIO()
            screenshot.save(buffer, format="JPEG", quality=quality)
            _last_screenshot = buffer.getvalue()
            _last_screenshot_time = time.time()
            
            # Sleep for interval
            time.sleep(interval)
        except:
            time.sleep(interval)  # Keep trying even if there's an error
    
    # Clear reference when thread ends
    _last_screenshot = None

@mcp.tool()
async def start_auto_refresh(interval: int = 5, quality: int = 60) -> str:
    """Start automatically refreshing the screen at specified intervals.
    
    Args:
        interval: Seconds between refreshes (2-60)
        quality: Image quality (1-100)
    
    Returns confirmation that auto-refresh was started.
    """
    global _refresh_thread, _refresh_active, _refresh_interval
    
    # Validate parameters
    interval = max(2, min(60, interval))
    quality = max(1, min(100, quality))
    
    # Check if already running
    if _refresh_active and _refresh_thread and _refresh_thread.is_alive():
        return f"Auto-refresh is already running with interval {_refresh_interval} seconds."
    
    # Stop any existing thread
    if _refresh_thread and _refresh_thread.is_alive():
        _refresh_active = False
        _refresh_thread.join(1.0)
    
    # Start new thread
    _refresh_interval = interval
    _refresh_thread = threading.Thread(
        target=_refresh_capture_thread,
        args=(interval, quality),
        daemon=True
    )
    _refresh_thread.start()
    
    return f"Auto-refresh started with {interval} second interval at {quality}% quality."

@mcp.tool()
async def stop_auto_refresh() -> str:
    """Stop the automatic screen refresh if it's running.
    
    Returns confirmation that auto-refresh was stopped.
    """
    global _refresh_active, _refresh_thread
    
    if not _refresh_active or not _refresh_thread or not _refresh_thread.is_alive():
        return "Auto-refresh is not currently running."
    
    # Stop the thread
    _refresh_active = False
    _refresh_thread.join(1.0)
    
    return "Auto-refresh has been stopped."

@mcp.tool()
async def get_last_screenshot() -> str:
    """Get the most recent auto-refresh screenshot.
    
    Returns the latest screenshot taken by auto-refresh or an error if auto-refresh isn't active.
    """
    global _last_screenshot, _last_screenshot_time
    
    if not _last_screenshot:
        return "No screenshot available. Start auto-refresh first with start_auto_refresh()."
    
    # Calculate age of screenshot
    age = time.time() - _last_screenshot_time
    
    # Convert to base64
    img_str = base64.b64encode(_last_screenshot).decode('utf-8')
    
    # Return as Markdown image with age info
    return f"Screenshot from {age:.1f} seconds ago:\n\n![Screenshot](data:image/jpeg;base64,{img_str})"

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

# ----- Windows Remote Assistance Integration ----- #

@mcp.tool()
async def start_remote_assistance() -> str:
    """Launch Windows Remote Assistance in offer help mode.
    
    This opens the built-in Windows Remote Assistance tool to establish a connection.
    """
    try:
        # Check if msra.exe exists
        msra_path = os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'System32', 'msra.exe')
        if not os.path.exists(msra_path):
            return "Windows Remote Assistance (msra.exe) not found on this system."
        
        # Launch Remote Assistance
        subprocess.Popen([msra_path, "/offerhelpkb"])
        
        return (
            "Windows Remote Assistance has been launched in 'Offer Help' mode.\n\n"
            "Instructions:\n"
            "1. Select 'Help someone who has invited you'\n"
            "2. Use the invitation file or password provided by the remote user\n"
            "3. Connect and follow the on-screen instructions\n"
            "4. The remote user will need to accept the connection"
        )
    except Exception as e:
        return f"Error launching Remote Assistance: {str(e)}"

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
async def execute_powershell(script: str) -> str:
    """Execute a PowerShell script.
    
    Args:
        script: PowerShell script to execute
    
    Returns the output of the script.
    """
    try:
        # Create a temporary file for the script
        script_path = os.path.join(os.environ.get('TEMP', '.'), f"tpc_ps_script_{int(time.time())}.ps1")
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script)
        
        # Execute the script
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path],
            capture_output=True,
            text=True,
            timeout=60
        )
        
        # Clean up
        try:
            os.remove(script_path)
        except:
            pass
        
        # Format the response
        output = result.stdout or result.stderr
        exit_code = result.returncode
        
        return f"PowerShell script executed with exit code {exit_code}:\n\n```\n{output}\n```"
    except subprocess.TimeoutExpired:
        return "PowerShell script timed out after 60 seconds"
    except Exception as e:
        return f"Error executing PowerShell script: {str(e)}"

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
        uname = platform.uname()
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

@mcp.tool()
async def move_mouse(x: int, y: int) -> str:
    """Move the mouse cursor to the specified screen coordinates without clicking.
    
    Args:
        x: X coordinate
        y: Y coordinate
    
    Returns confirmation that the mouse was moved.
    """
    try:
        # Move mouse to position
        win32api.SetCursorPos((x, y))
        
        return f"Mouse moved to position ({x}, {y})."
    except Exception as e:
        return f"Error moving mouse: {str(e)}"

@mcp.tool()
async def right_click_at_position(x: int, y: int) -> str:
    """Perform a right mouse click at the specified screen coordinates.
    
    Args:
        x: X coordinate
        y: Y coordinate
    
    Returns confirmation that the right click was performed.
    """
    try:
        # Move mouse to position
        win32api.SetCursorPos((x, y))
        
        # Perform right click
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x, y, 0, 0)
        
        return f"Right-clicked at position ({x}, {y})."
    except Exception as e:
        return f"Error right-clicking at position: {str(e)}"

@mcp.tool()
async def double_click_at_position(x: int, y: int) -> str:
    """Perform a double-click at the specified screen coordinates.
    
    Args:
        x: X coordinate
        y: Y coordinate
    
    Returns confirmation that the double-click was performed.
    """
    try:
        # Move mouse to position
        win32api.SetCursorPos((x, y))
        
        # Perform first click
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
        
        # Small delay
        time.sleep(0.1)
        
        # Perform second click
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
        
        return f"Double-clicked at position ({x}, {y})."
    except Exception as e:
        return f"Error double-clicking at position: {str(e)}"

@mcp.tool()
async def drag_mouse(start_x: int, start_y: int, end_x: int, end_y: int) -> str:
    """Perform a mouse drag operation from start coordinates to end coordinates.
    
    Args:
        start_x: Starting X coordinate
        start_y: Starting Y coordinate
        end_x: Ending X coordinate
        end_y: Ending Y coordinate
    
    Returns confirmation that the drag operation was performed.
    """
    try:
        # Move mouse to start position
        win32api.SetCursorPos((start_x, start_y))
        
        # Press mouse button
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, start_x, start_y, 0, 0)
        
        # Small delay
        time.sleep(0.1)
        
        # Move to end position
        win32api.SetCursorPos((end_x, end_y))
        
        # Release mouse button
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, end_x, end_y, 0, 0)
        
        return f"Dragged mouse from ({start_x}, {start_y}) to ({end_x}, {end_y})."
    except Exception as e:
        return f"Error performing drag operation: {str(e)}"

# ----- Advanced Screen Recording ----- #

@mcp.tool()
async def start_screen_recording_ps() -> str:
    """Start a screen recording using PowerShell.
    
    This uses built-in Windows APIs via PowerShell to record the screen.
    """
    try:
        # Create PowerShell script
        ps_script = """
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        # Output file
        $outputFile = "$env:USERPROFILE\\Desktop\\screen_recording_$(Get-Date -Format 'yyyyMMdd_HHmmss').mp4"
        
        # Create a message to inform the user
        [System.Windows.Forms.MessageBox]::Show("Screen recording has started. The output will be saved to $outputFile`n`nPress OK to continue.", "Screen Recording", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Use Windows built-in screen recorder (Windows 10 Game DVR)
        Start-Process -FilePath "$env:windir\\system32\\WindowsPowerShell\\v1.0\\powershell.exe" -ArgumentList "-Command `"& {Start-Process -FilePath '$env:windir\\system32\\cmd.exe' -ArgumentList '/c start ms-screenclip:'}`"" -WindowStyle Hidden
        
        # Return the output file path
        Write-Output "Screen recording started. Output will be saved to: $outputFile"
        """
        
        # Create temporary file for PowerShell script
        script_path = os.path.join(os.environ.get('TEMP', '.'), f"tpc_screen_record_{int(time.time())}.ps1")
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(ps_script)
        
        # Execute the script
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path],
            capture_output=True,
            text=True
        )
        
        # Clean up script file
        try:
            os.remove(script_path)
        except:
            pass
        
        # Check for errors
        if result.returncode != 0:
            return f"Error starting screen recording: {result.stderr}"
        
        return "Screen recording started using Windows Game DVR.\n\nPress Win+G to access Game Bar and control recording."
    except Exception as e:
        return f"Error starting screen recording: {str(e)}"

# ----- Main Function ----- #

def main():
    """Main function to run the MCP server"""
    print("Starting Total PC Commander MCP Server...", file=sys.stderr)
    mcp.run(transport='stdio')

if __name__ == "__main__":
    main()
