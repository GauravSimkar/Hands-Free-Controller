# build.spec
# -*- mode: python -*-
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

block_cipher = None

a = Analysis(
    ['main_app.py'],
    pathex=[],
    binaries=collect_dynamic_libs('mediapipe'),
    datas=[
        *collect_data_files('mediapipe', subdir='modules'),
        ('voiceKey.json', '.'),
        ('*.json', '.'),
        ('.env', '.'),
        ('*.py', '.'),  # Include all project files
    ],
    hiddenimports=[
        'cv2', 'keyboard', 'pyautogui',
        'mediapipe', 'numpy', 'tkinter',
        'mediapipe.python.solutions.face_mesh',
        'mediapipe.tasks.python.vision.face_detector',
        'pythoncom', 'win32timezone', 'pywintypes',
        'yaml', 'julep', 'dotenv', 'pyttsx3', 'pyperclip',
        'win32gui', 'win32con'  # New dependencies added
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='VisoVox',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    onefile=True
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='VisoVox',
)