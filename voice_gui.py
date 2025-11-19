import tkinter as tk
from tkinter import ttk, scrolledtext
from threading import Thread
import platform
import ctypes
from ctypes import wintypes
from voice_assistant import VoiceAssistant

class VoiceTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.assistant = None
        self.log_buffer = []
        
        # Create UI components
        self.create_widgets()
        self.create_loading_window()
        self.start_assistant()

    def create_widgets(self):
        # Main UI elements
        header = ttk.Label(self, text="Voice Control Assistant", style='Header.TLabel')
        header.pack(pady=(10, 20))

        self.status_indicator = ttk.Label(
            self,
            text="Initializing...",
            foreground="#3498db",
            font=('Helvetica', 10, 'bold')
        )
        self.status_indicator.pack(pady=5)

        log_frame = ttk.Frame(self)
        log_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        self.log_area = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            state='disabled',
            bg='#ecf0f1',
            font=('Consolas', 9),
            padx=10,
            pady=10
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def create_loading_window(self):
        # Create the floating status window
        self.loading_window = tk.Toplevel(self)
        self.loading_window.overrideredirect(True)
        self.loading_window.attributes("-topmost", True)
        self.loading_window.attributes("-alpha", 0.7)
        self.loading_window.configure(bg='#ffffff')

        self.loading_label = ttk.Label(
            self.loading_window,
            text="Initializing...",
            font=('Helvetica', 10, 'bold'),
            background='#ffffff',
            padding=10
        )
        self.loading_label.pack()

        # Initial positioning
        self.update_loading_position()
        
        # Set click-through properties (Windows only)
        if platform.system() == "Windows":
            self.set_click_through()

    def set_click_through(self):
        """Windows-specific click-through implementation"""
        hwnd = ctypes.windll.user32.GetParent(self.loading_window.winfo_id())
        
        # Set extended window style
        ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
        ex_style |= 0x80000  # WS_EX_LAYERED
        ex_style |= 0x20      # WS_EX_TRANSPARENT
        ctypes.windll.user32.SetWindowLongW(hwnd, -20, ex_style)
        
        # Set layered window attributes
        ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, 178, 0x2)

    def get_work_area(self):
        """Get available screen space excluding taskbar (Windows)"""
        if platform.system() == "Windows":
            from ctypes import wintypes
            rect = wintypes.RECT()
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)
            return rect.right, rect.bottom
        else:
            return (
                self.winfo_screenwidth(),
                self.winfo_screenheight()
            )

    def update_loading_position(self):
        """Position window at bottom-right of screen, just above taskbar"""
        if not self.loading_window.winfo_exists():
            return

        # Get updated window dimensions
        self.loading_window.update_idletasks()
        window_width = self.loading_window.winfo_width()
        window_height = self.loading_window.winfo_height()

        # Get screen dimensions (work area excludes taskbar)
        work_right, work_bottom = self.get_work_area()
        screen_right = self.loading_window.winfo_screenwidth()
        
        # Calculate position
        x = screen_right - window_width - 5  # 5px from right edge
        y = work_bottom - window_height - 2  # 2px above taskbar

        # Only update if position changed significantly
        if not hasattr(self, '_last_pos') or abs(x - self._last_pos[0]) > 2 or abs(y - self._last_pos[1]) > 2:
            self.loading_window.geometry(f"+{x}+{y}")
            self._last_pos = (x, y)

    def start_assistant(self):
        try:
            self.log("Starting voice assistant...")
            self.update_status("waiting")
            self.assistant = VoiceAssistant(gui=self)
            Thread(target=self.assistant.run, daemon=True).start()
        except Exception as e:
            self.log(f"Failed to start assistant: {e}")
            self.update_status("error")

    def update_status(self, status):
        status_config = {
            "waiting": ("Waiting for wake word...", "#f1c40f"),
            "active": ("Active - Listening", "#2ecc71"),
            "error": ("Error Occurred", "#e74c3c")
        }
        text, color = status_config.get(status, ("Unknown Status", "#95a5a6"))
        self.status_indicator.config(text=text, foreground=color)

    def log(self, message):
        def _append():
            self.log_buffer.append(message)
            if len(self.log_buffer) > 10:
                self.log_buffer.pop(0)
            self.log_area.config(state='normal')
            self.log_area.delete(1.0, tk.END)
            self.log_area.insert(tk.END, "\n\n".join(self.log_buffer))
            self.log_area.see(tk.END)
            self.log_area.config(state='disabled')
        self.after(0, _append)

    def start_loading(self, message=None):
        if message:
            text = f"Processing: {message[:35]}..." if len(message) > 40 else f"Processing: {message}"
        else:
            text = "Processing commands..."
        self.loading_label.config(text=text)
        self.update_loading_position()

    def stop_loading(self):
        self.loading_label.config(text="Ready for new commands")
        self.update_loading_position()

    def stop(self):
        if self.assistant:
            VoiceAssistant.is_active = False
        if self.loading_window:
            self.loading_window.destroy()