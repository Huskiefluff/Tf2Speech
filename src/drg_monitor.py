#!/usr/bin/env python3
"""
Deep Rock Galactic Chat Log Monitor
Monitors CSV format chat logs from DRG
"""

import csv
import time
import threading
import logging
from pathlib import Path
from typing import Optional, Dict, List, Callable
from datetime import datetime

logger = logging.getLogger(__name__)

class DRGLogMonitor:
    """Monitor Deep Rock Galactic CSV chat log file"""
    
    def __init__(self, log_path: str):
        self.log_path = Path(log_path) if log_path else None
        self.last_line_index = -1  # Track last processed line by index
        self.running = False
        self.callbacks = []
        self.thread = None
        
    def add_callback(self, callback: Callable):
        """Add a callback function to be called when new messages arrive"""
        self.callbacks.append(callback)
        
    def parse_csv_line(self, row: List[str]) -> Optional[Dict]:
        """Parse CSV row format: index,timestamp,steamid,username,message"""
        try:
            if len(row) < 5:
                return None
                
            index = int(row[0]) if row[0].isdigit() else -1
            timestamp = row[1]
            steamid = row[2]
            username = row[3]
            message = row[4]
            
            # Skip if we've already processed this line
            if index <= self.last_line_index:
                return None
                
            # Update last processed index
            self.last_line_index = max(self.last_line_index, index)
            
            # Check if message is a TTS command
            # DRG mod does NOT strip the prefix - it logs the full message
            # So "!tts hello" in game becomes "!tts hello" in CSV
            
            # Pass the full message to main app to handle with configurable prefix
            actual_message = message
            
            # Mark as TTS command if it starts with !tts
            if message.lower().startswith('!tts '):
                is_tts_command = True
            else:
                is_tts_command = False
            
            return {
                'username': username,
                'message': actual_message,
                'steamid': steamid,
                'timestamp': timestamp,
                'index': index,
                'is_tts_command': is_tts_command,
                'is_dead': False,  # DRG doesn't have dead state in chat
                'is_team': False,  # DRG doesn't have team chat in this format
                'game': 'drg',
                'raw': ','.join(row)
            }
            
        except Exception as e:
            logger.error(f"Failed to parse DRG CSV line: {e}, row: {row}")
            return None
            
    def monitor_loop(self):
        """Monitor loop for CSV file"""
        while self.running:
            try:
                if self.log_path and self.log_path.exists():
                    with open(self.log_path, 'r', encoding='utf-8', errors='ignore') as f:
                        reader = csv.reader(f)
                        
                        for row in reader:
                            if not self.running:
                                break
                                
                            parsed = self.parse_csv_line(row)
                            if parsed:
                                # Only process TTS commands for DRG
                                if parsed['is_tts_command']:
                                    for callback in self.callbacks:
                                        try:
                                            callback(parsed)
                                        except Exception as e:
                                            logger.error(f"DRG callback error: {e}")
                                            
            except Exception as e:
                logger.error(f"DRG monitor error: {e}")
                
            time.sleep(0.5)  # Check every 500ms
            
    def start(self):
        """Start monitoring the DRG log file"""
        if self.running:
            return
            
        self.running = True
        
        # Read the file to find the highest index
        if self.log_path and self.log_path.exists():
            try:
                with open(self.log_path, 'r', encoding='utf-8', errors='ignore') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if row and row[0].isdigit():
                            self.last_line_index = max(self.last_line_index, int(row[0]))
            except Exception as e:
                logger.error(f"Failed to read initial DRG log state: {e}")
                
        logger.info(f"Starting DRG monitor from index {self.last_line_index}")
        self.thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.thread.start()
        
    def stop(self):
        """Stop monitoring"""
        self.running = False
        if self.thread:
            self.thread.join(timeout=1)
            
    def get_last_position(self) -> int:
        """Get the last processed line index"""
        return self.last_line_index
        
    def set_last_position(self, position: int):
        """Set the last processed line index"""
        self.last_line_index = position