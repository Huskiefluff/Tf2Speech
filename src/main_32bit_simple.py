#!/usr/bin/env python3
"""
TTS for Team Fortress 2 - Simple 32-bit Build Entry Point
Minimal entry point that includes tkinter
"""

import sys
import os
import logging
import logging.handlers
from pathlib import Path
import traceback


def wait_for_previous_copy():
    """After an update the new copy is started with --wait-pid <old copy>: wait (at most 20 s) until the old copy
    has closed, so the two never run at the same time or write the settings and log together"""
    if '--wait-pid' not in sys.argv:
        return
    try:
        import ctypes
        pid = int(sys.argv[sys.argv.index('--wait-pid') + 1])
        handle = ctypes.windll.kernel32.OpenProcess(0x00100000, False, pid)  # SYNCHRONIZE
        if handle:
            ctypes.windll.kernel32.WaitForSingleObject(handle, 20000)
            ctypes.windll.kernel32.CloseHandle(handle)
    except (ValueError, IndexError, OSError, AttributeError):
        pass


wait_for_previous_copy()

# Set up logging first
if getattr(sys, 'frozen', False):
    log_dir = Path(sys.executable).parent
else:
    log_dir = Path.cwd()

# Chat is never written to the log (only lengths, voices and names), and the log is capped at about 10 MB:
# tts_tf2.log plus the previous one as tts_tf2.log.1
log_file = log_dir / 'tts_tf2.log'
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.handlers.RotatingFileHandler(log_file, maxBytes=5_000_000, backupCount=1, encoding='utf-8')]
    # Console output only when there is a console (the windowed exe has none, so sys.stderr is None there)
    + ([logging.StreamHandler()] if sys.stderr else [])
)


def show_fatal_error(error_msg):
    """With no console window, a startup crash would otherwise vanish silently - show it in a pop-up"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("TTS TF2 - failed to start", error_msg[-1500:] + f"\n\nFull details are in {log_file}")
        root.destroy()
    except Exception:
        pass

logger = logging.getLogger(__name__)

def main():
    """Main entry point"""
    # Simple test mode
    if len(sys.argv) > 1 and sys.argv[1] in ['--test', '-test', '/test']:
        print("Running in test mode...")
        print("\n1. Testing SAPI5...")
        try:
            import pyttsx3
            engine = pyttsx3.init('sapi5')
            print("   [OK] SAPI5 initialized")
            engine.stop()
        except Exception as e:
            print(f"   [FAIL] SAPI5 failed: {e}")
        
        print("\n2. Testing DECtalk...")
        try:
            from dectalk_native import DECtalkNative
            dectalk = DECtalkNative()
            if dectalk.is_available():
                print("   [OK] DECtalk loaded")
        except Exception as e:
            print(f"   [FAIL] DECtalk failed: {e}")
        
        print("\nTest complete!")
        return
    
    # Normal GUI mode
    try:
        logger.info("Starting TTS TF2 GUI...")
        
        # Import everything needed for GUI
        import tkinter as tk
        from tkinter import ttk, scrolledtext, messagebox
        
        # Import the main app
        from main_32bit_full import TTSReplicaWASAPI
        
        # Create and run app
        app = TTSReplicaWASAPI()
        app.run()
        
    except ImportError as e:
        error_msg = f"Failed to load GUI components: {e}\n\nTraceback:\n{traceback.format_exc()}"
        logger.error(error_msg)
        print(error_msg)
        print("\nThe GUI requires tkinter. You can run in test mode with: --test")
        show_fatal_error(error_msg)
    except Exception as e:
        error_msg = f"Fatal error: {e}\n\nTraceback:\n{traceback.format_exc()}"
        logger.error(error_msg)
        print(error_msg)
        show_fatal_error(error_msg)

if __name__ == "__main__":
    main()