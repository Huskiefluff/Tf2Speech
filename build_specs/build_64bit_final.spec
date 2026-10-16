# -*- mode: python ; coding: utf-8 -*-
# 64-bit Build - SINGLE EXE FILE - Final Version without Acapela
# Self-contained executable that extracts to exe directory

import os
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# Base path - this spec lives in build_specs\, so the project is one folder up (works wherever it's cloned)
project_root = Path(SPECPATH).parent

# git doesn't keep empty folders, so make sure the bundled temp folder exists on a fresh clone
(project_root / 'voice_data' / 'temp').mkdir(parents=True, exist_ok=True)

# The Python running PyInstaller (its Tcl/Tk files get bundled) - C:\Python313 on the original build PC
python64_path = Path(sys.base_prefix)

# Collect all tkinter submodules
tkinter_modules = collect_submodules('tkinter')

# Build data list - Include only DECtalk voices (Acapela removed)
datas = [
    # DECtalk vs6 files - ALL languages (us, uk, fr, gr, sp, la)
    (str(project_root / 'voice_data' / 'dectalk' / 'vs6'), 'dectalk/vs6'),
    
    # DECtalk extended voices from GitHub
    (str(project_root / 'voice_data' / 'dectalk' / 'extended'), 'dectalk/extended'),
    
    # Create empty temp directory
    (str(project_root / 'voice_data' / 'temp'), 'temp'),
    
    # TCL/TK runtime files for tkinter GUI  
    # PyInstaller expects these in specific directories
    (str(python64_path / 'tcl' / 'tcl8.6'), '_tcl_data/tcl/tcl8.6'),
    (str(python64_path / 'tcl' / 'tk8.6'), '_tk_data/tk/tk8.6'),
]

# Main analysis
a = Analysis(
    [str(project_root / 'src' / 'main_32bit_simple.py')],  # Still using same entry point
    pathex=[str(project_root / 'src')],
    binaries=[
        # Tkinter DLLs for GUI - 64-bit versions
        (str(python64_path / 'DLLs' / '_tkinter.pyd'), '.'),
        (str(python64_path / 'DLLs' / 'tcl86t.dll'), '.'),
        (str(python64_path / 'DLLs' / 'tk86t.dll'), '.'),
    ],
    datas=datas,
    hiddenimports=[
        # Core TTS
        'pyttsx3',
        'pyttsx3.drivers',
        'pyttsx3.drivers.sapi5',
        
        # Windows COM
        'win32com',
        'win32com.client',
        'pythoncom',
        'comtypes',
        'comtypes.client',
        
        # Audio processing
        'numpy',
        'scipy',
        'scipy.signal',
        'scipy.io',
        'scipy.io.wavfile',
        'sounddevice',
        'pyaudio',
        'wave',
        
        # System
        'ctypes',
        'winreg',
        'subprocess',
        'socket',
        
        # GUI - Force include tkinter and all submodules
        'tkinter',
        'tkinter.ttk',
        'tkinter.scrolledtext',
        'tkinter.messagebox',
        'tkinter.filedialog',
        '_tkinter',
    ] + tkinter_modules + [
        
        # Utilities
        'tempfile',
        'threading',
        'queue',
        'logging',
        'json',
        'pathlib',
        'typing',
        'time',
        'struct',
        'io',
        'os',
        'sys',
        're',
        
        # Our modules
        'temp_utils',
        'main_32bit_full',
        'dectalk_native',
        'sapi5_direct',
        'drg_monitor',  # New DRG monitor for Deep Rock Galactic support
        'dj_player',  # DJ window (music + YouTube via ffmpeg/yt-dlp)
        'soundboard',  # chat soundboard (Sven ChatSounds style)
        'textaloud_voice',  # TextAloud voices (Will From Afar) via TextAloud's automation interface
        'ui_theme',  # Windows 11 look (Sun Valley theme), dark mode, scaling, scrolling tabs
        'sv_ttk',
        'voice_installer',  # Voices tab: Windows' extra voices, free Piper voices
        'urllib.request',
        'array',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'PIL',
        'pandas',
        'notebook',
        'jupyter',
        'IPython',
        'tornado',
        'jedi',
        'cv2',
        'opencv',
        'pytest',
        'setuptools',
        'pip',
        'wheel',
        'cryptography',
        'h5py',
        'numba',
        'sqlalchemy',
        'flask',
        'django',
        'requests',
        'urllib3',
        'certifi',
    ],
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
    [],
    name='TTS_TF2_64BIT',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=['*.dll', 'tcl86t.dll', 'tk86t.dll'],  # Don't compress DLLs or Tk files
    runtime_tmpdir=None,  # Use default extraction (will use _MEIPASS)
    console=False,  # no cmd window; everything is still logged to tts_tf2.log next to the exe
    disable_windowed_traceback=False,
    argv_emulation=False,
    # Note: No target_arch specified - defaults to 64-bit on 64-bit Python
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
    version_file=None,
)