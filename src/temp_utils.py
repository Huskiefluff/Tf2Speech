"""
Temporary file utilities for TTS system
Ensures temp files are created within the application bundle, not Windows temp
"""

import os
import sys
import tempfile
from pathlib import Path
import logging
import atexit
import shutil

logger = logging.getLogger(__name__)

# Track temp files for cleanup
_temp_files = set()
_temp_dir = None

def get_temp_dir():
    """
    Get the temporary directory for the application.
    For frozen builds, uses _MEIPASS or exe directory.
    For development, uses project temp directory.
    """
    global _temp_dir
    
    if _temp_dir is not None:
        return _temp_dir
    
    if getattr(sys, 'frozen', False):
        # Running in PyInstaller bundle
        if hasattr(sys, '_MEIPASS'):
            # Use the extracted bundle directory for temp files
            temp_dir = Path(sys._MEIPASS) / "temp"
        else:
            # Fallback to exe directory if _MEIPASS not available
            exe_dir = Path(sys.executable).parent
            temp_dir = exe_dir / "temp"
    else:
        # Development mode - use project temp directory
        temp_dir = Path(__file__).parent.parent / "voice_data" / "temp"
    
    # Ensure temp directory exists
    temp_dir.mkdir(parents=True, exist_ok=True)
    logger.debug(f"Using temp directory: {temp_dir}")
    
    _temp_dir = temp_dir
    return temp_dir

def get_temp_file(suffix='', prefix='tts_'):
    """
    Create a temporary file in the application's temp directory.
    Returns the full path to the temp file.
    """
    temp_dir = get_temp_dir()
    
    # Create temp file in our directory
    fd, path = tempfile.mkstemp(suffix=suffix, prefix=prefix, dir=str(temp_dir))
    os.close(fd)  # Close the file descriptor
    
    # Track for cleanup
    _temp_files.add(path)
    
    logger.debug(f"Created temp file: {path}")
    return path

def cleanup_old_temp_files(max_age_hours=24):
    """
    Clean up old temporary files from the temp directory.
    """
    import time
    
    temp_dir = get_temp_dir()
    current_time = time.time()
    max_age_seconds = max_age_hours * 3600
    
    try:
        for file_path in temp_dir.glob("tts_*"):
            if file_path.is_file():
                file_age = current_time - file_path.stat().st_mtime
                if file_age > max_age_seconds:
                    try:
                        file_path.unlink()
                        logger.debug(f"Deleted old temp file: {file_path}")
                    except Exception as e:
                        logger.warning(f"Could not delete temp file {file_path}: {e}")
    except Exception as e:
        logger.warning(f"Error cleaning temp files: {e}")

def cleanup_on_exit():
    """
    Clean up all temp files created during this session.
    Called automatically on exit.
    """
    for path in _temp_files:
        try:
            if os.path.exists(path):
                os.unlink(path)
                logger.debug(f"Cleaned up temp file: {path}")
        except Exception as e:
            logger.warning(f"Could not clean up temp file {path}: {e}")
    
    # Clear the set
    _temp_files.clear()
    
    # If running from _MEIPASS, don't delete the temp dir itself
    # as PyInstaller manages it
    if not (getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')):
        # In development or portable mode, clean empty temp dir
        temp_dir = get_temp_dir()
        if temp_dir and temp_dir.exists():
            try:
                # Only remove if empty
                if not any(temp_dir.iterdir()):
                    temp_dir.rmdir()
                    logger.debug(f"Removed empty temp directory: {temp_dir}")
            except Exception as e:
                logger.debug(f"Could not remove temp directory: {e}")

# Register cleanup function to run on exit
atexit.register(cleanup_on_exit)