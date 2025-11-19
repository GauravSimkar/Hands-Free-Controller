import platform
import sys
import os

def get_active_window():
    """Get current focused window title"""
    try:
        if platform.system() == 'Windows':
            import win32gui
            window = win32gui.GetForegroundWindow()
            return win32gui.GetWindowText(window)
        elif platform.system() == 'Linux':
            from subprocess import check_output
            return check_output(['xdotool', 'getwindowfocus', 'getwindowname']).decode().strip()
        else:
            return "Unknown OS"
    except Exception as e:
        print(f"Window detection error: {e}")
        return "Unknown"
    
import winreg


def find_installed_app(app_name):
    """Search Windows registry for application paths"""
    locations = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths")
    ]
    
    for root, base_key in locations:
        try:
            with winreg.OpenKey(base_key, root) as key:
                for subkey in [f"{app_name}.exe", app_name]:
                    try:
                        path = winreg.QueryValue(key, subkey)
                        if os.path.exists(path):
                            return f'start "" "{path}"'
                    except FileNotFoundError:
                        continue
        except Exception:
            continue
            
    return None

import win32gui

def get_all_open_windows():
    """Fast window detection using Win32 API"""
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            windows.append(win32gui.GetWindowText(hwnd))
        return True
    windows = []
    win32gui.EnumWindows(callback, windows)
    return set(windows)