# mouse_gui.py
import tkinter as tk
from tkinter import ttk, messagebox
from face_controller import FaceController
import keyboard
from css import Style  # Renamed from config.py to css.py

DEFAULT_PARAMS = {
    'EMA_ALPHA': 0.15,
    'CURSOR_SENSITIVITY_X': 8,
    'CURSOR_SENSITIVITY_Y': 8,
    'BLINK_THRESHOLD_RATIO': 0.75,
    'CALIBRATION_FRAMES': 30,
    'MOVEMENT_THRESHOLD': 0.02,
    'DOUBLE_CLICK_THRESHOLD': 0.8,
    'EYEBROW_RAISE_THRESHOLD_RATIO': 1.4,
    'RIGHT_CLICK_COOLDOWN': 1.0,
    'MOUTH_OPEN_THRESHOLD_RATIO': 1.5,
    'SCROLL_THRESHOLD': 0.03,
    'SCROLL_STEP': 50,
    'SCROLL_INTERVAL': 0.05
}

class MouseTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.params = DEFAULT_PARAMS.copy()
        self.face_controller = FaceController(self.params)
        self.face_controller.start()
        self.entries = {}
        self.create_widgets()
        self.setup_hotkeys()
        
    def create_widgets(self):
        # Header
        header = ttk.Label(self, text="Facial Mouse Control", style='Header.TLabel')
        header.pack(pady=(10, 20), anchor=tk.CENTER)

        # Parameters Frame
        params_frame = ttk.Frame(self, style='Card.TFrame')
        params_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        parameters = [
            ('EMA Alpha', 'EMA_ALPHA'),
            ('Cursor Sensitivity X', 'CURSOR_SENSITIVITY_X'),
            ('Cursor Sensitivity Y', 'CURSOR_SENSITIVITY_Y'),
            ('Blink Threshold', 'BLINK_THRESHOLD_RATIO'),
            ('Calibration Frames', 'CALIBRATION_FRAMES'),
            ('Movement Threshold', 'MOVEMENT_THRESHOLD'),
            ('Double Click Threshold', 'DOUBLE_CLICK_THRESHOLD'),
            ('Eyebrow Raise Ratio', 'EYEBROW_RAISE_THRESHOLD_RATIO'),
            ('Right Click Cooldown', 'RIGHT_CLICK_COOLDOWN'),
            ('Mouth Open Ratio', 'MOUTH_OPEN_THRESHOLD_RATIO'),
            ('Scroll Threshold', 'SCROLL_THRESHOLD'),
            ('Scroll Step', 'SCROLL_STEP'),
            ('Scroll Interval', 'SCROLL_INTERVAL')
        ]

        for row, (label_text, param) in enumerate(parameters):
            self._create_input(params_frame, label_text, param, row)

        # Control Buttons
        btn_frame = ttk.Frame(self, style='Card.TFrame')
        btn_frame.pack(pady=10, padx=20, fill=tk.X)

        self.save_btn = ttk.Button(
            btn_frame, 
            text="Save Settings", 
            command=self.save_parameters,
            style='Primary.TButton'
        )
        self.save_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.toggle_btn = ttk.Button(
            btn_frame, 
            text="Pause (Ctrl+Caps Lock)", 
            command=self.toggle_control,
            style='Secondary.TButton'
        )
        self.toggle_btn.pack(side=tk.RIGHT, padx=5, pady=5)

    def _create_input(self, parent, label_text, param, row):
        frame = ttk.Frame(parent, style='Input.TFrame')
        frame.grid(row=row, column=0, sticky="ew", pady=3)

        ttk.Label(frame, text=label_text, style='InputLabel.TLabel').pack(side=tk.LEFT, padx=5)
        
        default_value = DEFAULT_PARAMS[param]
        validate_cmd = (self.register(self.validate_number), '%P', param)
        
        entry = ttk.Entry(
            frame, 
            validate="key", 
            validatecommand=validate_cmd,
            style='Input.TEntry'
        )
        entry.insert(0, str(self.params[param]))
        entry.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)
        
        self.entries[param] = entry

    def validate_number(self, new_value, param):
        if new_value == "":
            return True
            
        default_value = DEFAULT_PARAMS[param]
        try:
            if isinstance(default_value, int):
                int(new_value)
            else:
                float(new_value)
            return True
        except ValueError:
            parts = new_value.split('.')
            if len(parts) == 2 and parts[0] == '' and parts[1].isdigit():
                return True
            return False

    def save_parameters(self):
        errors = []
        updated_params = {}
        
        for param, entry in self.entries.items():
            value_str = entry.get()
            default_value = DEFAULT_PARAMS[param]
            
            try:
                if isinstance(default_value, int):
                    converted = int(float(value_str))
                else:
                    converted = float(value_str)
                updated_params[param] = converted
            except ValueError:
                errors.append(f"Invalid value for {param}: {value_str}")
        
        if errors:
            messagebox.showerror("Validation Error", "\n".join(errors))
            return
        
        for param, value in updated_params.items():
            self.params[param] = value
            self.face_controller.update_param(param, value)
        
        messagebox.showinfo("Success", "Parameters saved successfully")

    def toggle_control(self):
        self.face_controller.toggle_pause()
        btn_text = "Resume (Ctrl+Caps Lock)" if self.face_controller.paused else "Pause (Ctrl+Caps Lock)"
        self.toggle_btn.config(text=btn_text)
        
    def setup_hotkeys(self):
        keyboard.add_hotkey('ctrl+caps lock', self.toggle_control)
        
    def stop(self):
        self.face_controller.stop()
        keyboard.unhook_all_hotkeys()