from dotenv import load_dotenv
load_dotenv()
import tkinter as tk
from tkinter import ttk
from voice_gui import VoiceTab
from mouse_gui import MouseTab
from instruction_tab import InstructionsTab  # New import
from css import Style
from config import Config
import os
import sys

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Integrated Assistants")
        self.geometry("800x800")
        
        # Initialize styling once for entire application
        self.style = Style()
        
        # Create widgets after style initialization
        self.create_widgets()
        
    def create_widgets(self):
        # Create notebook with styled tabs
        self.notebook = ttk.Notebook(self)
        
        # Add voice tab
        self.voice_tab = VoiceTab(self.notebook)
        self.notebook.add(self.voice_tab, text=" Voice Control ")
        
        # Add mouse tab
        self.mouse_tab = MouseTab(self.notebook)
        self.notebook.add(self.mouse_tab, text=" Mouse Control ")
        
        # Add instructions tab
        self.instructions_tab = InstructionsTab(self.notebook)
        self.notebook.add(self.instructions_tab, text=" Instructions ")
        
        self.notebook.pack(expand=1, fill="both")
        self.notebook.select(0)  # Set voice tab as default

    def on_closing(self):
        self.voice_tab.stop()
        self.mouse_tab.stop()
        self.destroy()

if __name__ == "__main__":
    # Use resource_path for credentials
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = resource_path('voiceKey.json')
    app = MainApplication()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()