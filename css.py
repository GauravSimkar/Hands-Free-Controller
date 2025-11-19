# css.py
import tkinter as tk
from tkinter import ttk

class Style:
    _instance = None
    _theme_created = False

    def __new__(cls):
        if not cls._instance:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if not self._initialized:
            self.style = ttk.Style()
            self._configure_styles()
            self._initialized = True
            
    def _configure_styles(self):
        if not Style._theme_created:
            self.style.theme_create('custom', parent='clam', settings={
                'TFrame': {
                    'configure': {'background': '#f0f2f5'}
                },
                'Main.TFrame': {
                    'configure': {'background': '#ffffff'}
                },
                'Card.TFrame': {
                    'configure': {
                        'background': '#ffffff',
                        'relief': 'groove',
                        'borderwidth': 2
                    }
                },
                'Header.TLabel': {
                    'configure': {
                        'font': ('Helvetica', 14, 'bold'),
                        'foreground': '#2d3436',
                        'background': '#ffffff'
                    }
                },
                'Input.TFrame': {
                    'configure': {
                        'background': '#ffffff',
                        'padding': 5
                    }
                },
                'InputLabel.TLabel': {
                    'configure': {
                        'font': ('Helvetica', 9),
                        'foreground': '#4a4a4a',
                        'background': '#ffffff'
                    }
                },
                'Input.TEntry': {
                    'configure': {
                        'fieldbackground': '#f8f9fa',
                        'foreground': '#2d3436',
                        'padding': 5,
                        'relief': 'flat'
                    },
                    'map': {
                        'fieldbackground': [('focus', '#e3f2fd')]
                    }
                },
                'Primary.TButton': {
                    'configure': {
                        'font': ('Helvetica', 10, 'bold'),
                        'foreground': '#ffffff',
                        'background': '#3498db',
                        'padding': 8,
                        'relief': 'flat'
                    },
                    'map': {
                        'background': [
                            ('pressed', '#2980b9'),
                            ('active', '#5dade2')
                        ]
                    }
                },
                'Secondary.TButton': {
                    'configure': {
                        'font': ('Helvetica', 10),
                        'foreground': '#4a4a4a',
                        'background': '#ecf0f1',
                        'padding': 6,
                        'relief': 'flat'
                    },
                    'map': {
                        'background': [
                            ('pressed', '#bdc3c7'),
                            ('active', '#d0d3d6')
                        ]
                    }
                }
            })
            self.style.theme_use('custom')
            Style._theme_created = True