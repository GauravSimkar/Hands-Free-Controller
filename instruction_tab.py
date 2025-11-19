import tkinter as tk
from tkinter import ttk

class InstructionsTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.create_widgets()

    def create_widgets(self):
        container = ttk.Frame(self, style='Card.TFrame')
        container.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        # Face Gestures Section
        ttk.Label(container, text="Face Gestures (Mouse Control)", style='Header.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        gestures = [
            ("Cursor Movement", "Move your head to control the cursor"),
            ("Single Click", "Blink normally once"),
            ("Double Click", "Blink twice quickly"),
            ("Right Click", "Raise eyebrows"),
            ("Scroll", "Open mouth and move head:\n↑ Scroll Up | ↓ Scroll Down\n← Scroll Left | → Scroll Right"),
            ("Drag & Drop", "Open mouth to start dragging,\nclose to release")
        ]
        
        for text, desc in gestures:
            frame = ttk.Frame(container)
            frame.pack(fill=tk.X, pady=5)
            ttk.Label(frame, text=text+":", style='InputLabel.TLabel', width=15).pack(side=tk.LEFT)
            ttk.Label(frame, text=desc, style='InputLabel.TLabel').pack(side=tk.LEFT, padx=5)

        # Voice Control Section
        ttk.Label(container, text="Voice Commands", style='Header.TLabel').pack(anchor=tk.W, pady=(20, 10))
        
        voice_commands = [
            ("System Control:", "Open apps: 'Open [app name]'\nPress keys: 'Press [keys]'\nRun commands: 'Run [command]'"),
            ("Navigation:", "Open URL: 'Go to [website]'\nType text: 'Type [text]'\nClick: 'Click here'"),
            ("Special:", "Stop listening: 'Abrogate'\nRead text: 'Read from cursor'\nStop speech: 'Stop speaking'")
        ]

        for text, desc in voice_commands:
            frame = ttk.Frame(container)
            frame.pack(fill=tk.X, pady=5)
            ttk.Label(frame, text=text, style='InputLabel.TLabel', width=15).pack(side=tk.LEFT)
            ttk.Label(frame, text=desc, style='InputLabel.TLabel').pack(side=tk.LEFT, padx=5)