#!/usr/bin/env python3
"""
TTS for Team Fortress 2 - Simple 32-bit Build Entry Point
Minimal entry point that includes tkinter
"""

import sys
import os
import logging
from pathlib import Path
import traceback

# Set up logging first
if getattr(sys, 'frozen', False):
    log_dir = Path(sys.executable).parent
else:
    log_dir = Path.cwd()

log_file = log_dir / 'tts_tf2.log'
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

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
    except Exception as e:
        error_msg = f"Fatal error: {e}\n\nTraceback:\n{traceback.format_exc()}"
        logger.error(error_msg)
        print(error_msg)

if __name__ == "__main__":
    main()