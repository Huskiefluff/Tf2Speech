#!/usr/bin/env python3
"""
TTS for Team Fortress 2 - Replica with WASAPI Audio
Enhanced version with direct device audio routing
"""

import sys
import os
import time
import logging
import json
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from pathlib import Path
import threading
import queue
from typing import Dict, List, Optional
import traceback
from tkinter import filedialog

# Import DRG monitor
try:
    from drg_monitor import DRGLogMonitor
    DRG_MONITOR_AVAILABLE = True
except Exception as e:
    logger.warning(f"DRG monitor not available: {e}")
    DRG_MONITOR_AVAILABLE = False
    class DRGLogMonitor:
        def __init__(self, log_path):
            pass
        def start(self):
            pass
        def stop(self):
            pass

# Get exe directory for log file
if getattr(sys, 'frozen', False):
    log_dir = Path(sys.executable).parent
else:
    log_dir = Path.cwd()

# Configure logging FIRST - in exe directory
log_file = log_dir / 'tts_tf2.log'
try:
    logging.basicConfig(
        level=logging.DEBUG,  # More verbose for debugging
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
except Exception as e:
    # Fallback if can't create log in exe dir
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('tts_tf2.log'),
            logging.StreamHandler()
        ]
    )

logger = logging.getLogger(__name__)

# Log startup info immediately
logger.info("="*60)
logger.info("TTS TF2 Starting...")
logger.info(f"Python: {sys.version}")
logger.info(f"Frozen: {getattr(sys, 'frozen', False)}")
logger.info(f"Executable: {sys.executable}")
logger.info(f"Working directory: {os.getcwd()}")
if getattr(sys, 'frozen', False):
    logger.info(f"Exe directory: {Path(sys.executable).parent}")
    if hasattr(sys, '_MEIPASS'):
        logger.info(f"Bundle directory: {sys._MEIPASS}")
logger.info("="*60)

# Import SAPI5 
from sapi5_direct import DirectOutputTTS

# Try to import DECtalk native with fallback
try:
    from dectalk_native import DECtalkNative, DECtalkManager
    DECTALK_NATIVE_AVAILABLE = True
    print(f"SUCCESS: DECtalk native module imported successfully")
    logger.info("DECtalk native module imported successfully")
except Exception as e:
    print(f"Warning: DECtalk native not available: {e}")
    logger.warning(f"Failed to import DECtalk native: {e}")
    DECTALK_NATIVE_AVAILABLE = False
    # Create dummy classes
    class DECtalkNative:
        def __init__(self):
            self.available = False
        def is_available(self):
            return False

    
    class DECtalkManager:
        def __init__(self, audio_manager=None):
            self.use_dectalk = False
        def is_dectalk_voice(self, voice_name):
            return False
        def speak(self, text, voice_name=None):
            return False



class GameLogMonitor:
    """Base class for game log monitoring"""
    def __init__(self, log_path: str):
        self.log_path = Path(log_path) if log_path else None
        self.running = False
        self.callbacks = []
        self.thread = None
        
    def add_callback(self, callback):
        self.callbacks.append(callback)
        
    def start(self):
        pass
        
    def stop(self):
        self.running = False
        if self.thread:
            self.thread.join(timeout=1)

class TF2LogMonitor(GameLogMonitor):
    """Monitor TF2 log file - exact replica of old system behavior"""
    
    def __init__(self, log_path: str):
        super().__init__(log_path)
        self.last_position = 0
        
    def add_callback(self, callback):
        self.callbacks.append(callback)
        
    def parse_line(self, line: str) -> Optional[Dict]:
        """Parse exactly like old system"""
        line = line.strip()
        if not line:
            return None
        
        # Debug logging for !block commands
        if "!block" in line.lower():
            logger.info(f"LOG PARSER: Found !block in line: '{line}'")
            
        # Match patterns from old system
        is_dead = "*DEAD*" in line
        is_team = "(TEAM)" in line
        
        # Clean line
        clean = line.replace("*DEAD*", "").replace("(TEAM)", "").strip()
        
        # Extract username and message
        if " : " in clean:
            parts = clean.split(" : ", 1)
            if len(parts) == 2:
                username = parts[0].strip().strip('"')
                message = parts[1].strip()
                
                # Debug logging for parsed !block commands
                if "!block" in message.lower():
                    logger.info(f"LOG PARSER: Parsed !block command - user: '{username}', message: '{message}'")
                
                return {
                    'username': username,
                    'message': message,
                    'is_dead': is_dead,
                    'is_team': is_team,
                    'raw': line
                }
        
        # Debug logging for unparsed !block lines
        if "!block" in line.lower():
            logger.info(f"LOG PARSER: Failed to parse !block line: '{line}'")
            
        return None
        
    def monitor_loop(self):
        """Monitor loop matching old system"""
        while self.running:
            try:
                if self.log_path and self.log_path.exists():
                    with open(self.log_path, 'r', encoding='utf-8', errors='ignore') as f:
                        f.seek(self.last_position)
                        
                        for line in f:
                            parsed = self.parse_line(line)
                            if parsed:
                                for callback in self.callbacks:
                                    try:
                                        callback(parsed)
                                    except Exception as e:
                                        logger.error(f"Callback error: {e}")
                                        
                        self.last_position = f.tell()
                        
            except Exception as e:
                logger.error(f"Monitor error: {e}")
                
            time.sleep(0.1)
            
    def start(self):
        if self.running:
            return
            
        self.running = True
        
        # Start from end of file
        if self.log_path and self.log_path.exists():
            with open(self.log_path, 'r') as f:
                f.seek(0, 2)
                self.last_position = f.tell()
                
        self.thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.thread.start()
        
    def stop(self):
        self.running = False
        if self.thread:
            self.thread.join(timeout=1)


class TTSReplicaWASAPI:
    """TTS GUI with WASAPI audio routing"""
    
    def __init__(self):
        self.config = self.load_config()
        self.monitors = {}  # Dictionary to hold monitors for each game
        self.current_game = self.config.get('current_game', 'tf2')
        self.audio_manager = None
        self.is_running = False
        
        # Initialize hidden voices list
        self.hidden_voices = self.config.get('hidden_voices', [])
        
        # Game configurations
        self.game_configs = self.init_game_configs()
        
        # Track last speaker for admin commands (per game)
        self.last_speakers = {'tf2': None, 'drg': None}
        
        # Message queue for sequential playback
        self.message_queue = queue.Queue()
        self.queue_thread = None
        self.queue_running = False
        self.currently_speaking_user = None
        self.current_voice_id = None  # Cache current voice to avoid unnecessary changes
        
        # Voice command mappings - will be populated with actual system voices
        # MUST be initialized BEFORE init_audio() is called!
        self.voice_commands = {}
        
        # Load saved voice commands if available
        saved_commands = self.config.get('voice_commands', {})
        if saved_commands:
            self.voice_commands = saved_commands
        else:
            # Default mapping - will be updated after voices are loaded if these don't exist
            self.voice_commands = {
                'v 0': 'Microsoft Sam',
                'v 1': 'Microsoft Lili - Chinese (China)',
                'v 2': 'Microsoft Zira Desktop - English (United States)',
                'v 3': 'Microsoft Mary',
                'v 4': 'Microsoft Anna - English (United States)',
                'v 5': '',
                'v 6': '',
                'v 7': 'Microsoft Irina Desktop - Russian',
                'v 8': '',
                'v 9': 'Microsoft David Desktop - English (United States)',
            }
        
        
        # User voice preferences - username -> voice name/id
        self.user_voice_preferences = self.config.get('user_voice_preferences', {})
        
        # Random voice for new users
        self.random_voice_enabled = self.config.get('random_voice_enabled', False)
        self.random_voice_exclusions = self.config.get('random_voice_exclusions', [])
        
        # Voice toggle command (default /vt, configurable)
        self.voice_toggle_command = self.config.get('voice_toggle_command', '/vt')
        
        # DECtalk voice profiles
        self.dectalk_profiles = self.config.get('dectalk_profiles', self.get_default_dectalk_profiles())
        # Ensure we always have profiles even if config is empty
        if not self.dectalk_profiles:
            self.dectalk_profiles = self.get_default_dectalk_profiles()
        self.dectalk_enabled = self.config.get('dectalk_enabled', True)  # Default to True so DECtalk is available
        
        # Now initialize audio AFTER all required attributes are set
        self.init_audio()
        
        # Initialize DECtalk native
        logger.info(f"DECTALK_NATIVE_AVAILABLE flag: {DECTALK_NATIVE_AVAILABLE}")
        print(f"DECTALK_NATIVE_AVAILABLE flag: {DECTALK_NATIVE_AVAILABLE}")
        self.dectalk_native = DECtalkNative()
        logger.info(f"DECtalk native instance created: {self.dectalk_native}")
        self.dectalk_manager = None  # Will be initialized after audio_manager
        
        # Initialize announcement vars to prevent errors
        self.announcement_vars = {}
        
        self.setup_gui()
        self.load_settings()
        
        # Auto-generate voice commands if empty
        if not self.voice_commands or all(not v for v in self.voice_commands.values()):
            self.auto_generate_voice_commands()
        
    def init_audio(self):
        """Initialize WASAPI audio manager"""
        try:
            self.audio_manager = DirectOutputTTS()
            self.voices = self.audio_manager.get_voices() if self.audio_manager else []
            self.audio_devices = self.audio_manager.get_devices() if self.audio_manager else []
            
            # Log device detection
            logger.info(f"Detected {len(self.audio_devices)} audio devices")
            for device in self.audio_devices:
                logger.info(f"  Audio device: {device}")
            
            # Initialize DECtalk manager with audio manager
            # Re-initialize DECtalk native to ensure it finds the extracted files in PyInstaller bundle
            logger.info("Initializing DECtalk for audio manager...")
            self.dectalk_native = DECtalkNative()
            logger.info(f"DECtalk native available after re-init: {self.dectalk_native.is_available()}")
            logger.info(f"DECtalk native path: {self.dectalk_native.dectalk_path}")
            
            self.dectalk_manager = DECtalkManager(self.audio_manager)
            # Share the same DECtalk instance to ensure consistency
            self.dectalk_manager.dectalk = self.dectalk_native
            self.dectalk_manager.use_dectalk = self.dectalk_native.is_available()
            logger.info(f"DECtalk manager configured with use_dectalk: {self.dectalk_manager.use_dectalk}")
            
            
            # Log all discovered voices
            logger.info(f"Found {len(self.voices)} Windows TTS voices:")
            for i, voice in enumerate(self.voices):
                logger.info(f"  Voice {i}: {voice.name} (ID: {voice.id})")
                
            # Auto-populate voice commands with actual voices if not saved
            if not self.config.get('voice_commands'):
                for i in range(min(10, len(self.voices))):
                    self.voice_commands[f'v {i}'] = self.voices[i].name
                logger.info("Auto-populated voice commands with system voices")
            
            # Don't populate longform_voice_combo here - GUI not ready yet
            # It will be populated in load_settings() after GUI is created
            
            # Apply default voice from config if set
            default_voice = self.config.get('default_voice', '')
            if default_voice:
                for voice in self.voices:
                    if voice.name == default_voice:
                        self.audio_manager.set_voice(voice.id)
                        self.current_voice_id = voice.id
                        logger.info(f"Set default voice to: {default_voice}")
                        break
            
            logger.info(f"Audio initialized with {len(self.voices)} voices and {len(self.audio_devices)} devices")
            
        except Exception as e:
            logger.error(f"Failed to init audio: {e}")
            self.audio_manager = None
            self.voices = []
            self.audio_devices = []
            
    def get_config_path(self) -> Path:
        """Get config path - portable for frozen builds"""
        if getattr(sys, 'frozen', False):
            # Save config next to exe for portability
            return Path(sys.executable).parent / 'config.json'
        else:
            # Development mode - use user profile
            import os
            user_profile = Path(os.environ.get('USERPROFILE', Path.home()))
            config_dir = user_profile / '.tts_tf2'
            config_dir.mkdir(parents=True, exist_ok=True)
            return config_dir / 'config.json'
    
    def load_config(self) -> Dict:
        """Load config with default fallback"""
        # Load default config first
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # In frozen build, default config is in bundle
            default_config_path = Path(sys._MEIPASS) / 'default_config.json'
        else:
            default_config_path = Path(__file__).parent / 'default_config.json'
        
        default_config = {}
        if default_config_path.exists():
            try:
                with open(default_config_path, 'r') as f:
                    default_config = json.load(f)
            except Exception as e:
                logger.error(f"Error loading default config: {e}")
        
        # Get config path
        config_path = self.get_config_path()
        
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    user_config = json.load(f)
                    
                    # Merge with defaults (user config takes precedence)
                    config = self.merge_configs(default_config, user_config)
                    
                    # Migrate legacy config to new multi-game format
                    config = self.migrate_config(config)
                        
                    # Handle auto_block as dict or bool
                    if 'auto_block' in config:
                        if isinstance(config['auto_block'], dict):
                            config['auto_block_enabled'] = config['auto_block'].get('enabled', False)
                            config['auto_block_keywords'] = config['auto_block'].get('keywords', [])
                        elif isinstance(config['auto_block'], bool):
                            config['auto_block_enabled'] = config['auto_block']
                    
                    # Ensure announcements is a dict
                    if 'announcements' in config and not isinstance(config['announcements'], dict):
                        config['announcements'] = default_config.get('announcements', {})
                    
                    return config
            except Exception as e:
                logger.error(f"Error loading config: {e}")
        
        # Return default config for new users
        logger.info("No config found, using defaults")
        return default_config
    
    def merge_configs(self, default: Dict, user: Dict) -> Dict:
        """Recursively merge user config with defaults"""
        result = default.copy()
        for key, value in user.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self.merge_configs(result[key], value)
            else:
                result[key] = value
        return result
    
    def init_game_configs(self):
        """Initialize game configurations"""
        games = self.config.get('games', {})
        
        # Ensure we have configs for both games
        if 'tf2' not in games:
            games['tf2'] = {
                'log_path': self.config.get('log_path', ''),
                'admins': self.config.get('admins', []),
                'blocked': self.config.get('blocked', []),
                'last_position': 0
            }
        
        if 'drg' not in games:
            games['drg'] = {
                'log_path': '',
                'admins': [],
                'blocked': [],
                'last_line_index': -1
            }
        
        return games
    
    def migrate_config(self, config):
        """Migrate old config format to new multi-game format"""
        if 'games' not in config:
            # Create new games structure
            config['games'] = {
                'tf2': {
                    'log_path': config.get('log_path', config.get('tf2', {}).get('log_path', '')),
                    'admins': config.get('admins', []),
                    'blocked': config.get('blocked', []),
                    'last_position': 0
                },
                'drg': {
                    'log_path': '',
                    'admins': [],
                    'blocked': [],
                    'last_line_index': -1
                }
            }
            
            # Set current game
            config['current_game'] = 'tf2'
            
            # Remove old format keys
            for key in ['log_path', 'tf2', 'tf2_log_path']:
                config.pop(key, None)
        
        return config
    
    def get_current_game_config(self):
        """Get config for currently selected game"""
        game_key = 'tf2' if self.current_game == 'Team Fortress 2' else 'drg'
        return self.game_configs.get(game_key, {})
    
    def on_game_changed(self, event=None):
        """Handle game selection change"""
        # Save current game's log path before switching
        old_game_key = 'tf2' if self.current_game == 'Team Fortress 2' else 'drg'
        if hasattr(self, 'log_path_var'):
            current_log_path = self.log_path_var.get()
            if current_log_path:
                self.game_configs[old_game_key]['log_path'] = current_log_path
                self.config['games'][old_game_key]['log_path'] = current_log_path
        
        selected_game = self.game_combo.get()
        game_key = 'tf2' if selected_game == 'Team Fortress 2' else 'drg'
        
        # Stop current monitor
        if self.is_running:
            self.stop_tts()
        
        # Update current game
        self.current_game = selected_game
        self.config['current_game'] = selected_game
        
        # Update UI with game-specific settings
        game_config = self.game_configs[game_key]
        self.log_path_var.set(game_config.get('log_path', ''))
        
        # Update admin and blocked lists
        self.admin_listbox.delete(0, tk.END)
        for admin in game_config.get('admins', []):
            self.admin_listbox.insert(tk.END, admin)
        
        self.blocked_listbox.delete(0, tk.END)
        for blocked in game_config.get('blocked', []):
            self.blocked_listbox.insert(tk.END, blocked)
            
        # Update TTS command prefix
        if hasattr(self, 'tts_prefix_entry'):
            self.tts_prefix_entry.delete(0, tk.END)
            self.tts_prefix_entry.insert(0, game_config.get('tts_command_prefix', '!tts'))
        
        # Save config after switching
        self.save_config()
        
        # Clear chat display
        self.chat_display.delete('1.0', tk.END)
        self.chat_display.insert('1.0', f"Switched to {selected_game}\n")
        
        # Save config
        self.save_config()
        
        logger.info(f"Switched to game: {selected_game}")
    
    def auto_generate_voice_commands(self):
        """Auto-generate voice commands from available voices"""
        logger.info("Auto-generating voice commands...")
        
        # Clear existing commands
        self.voice_commands = {}
        
        # Add numbered shortcuts for first 10 SAPI5 voices
        if self.voices:
            for i in range(min(10, len(self.voices))):
                self.voice_commands[f'v {i}'] = self.voices[i].name
        
        # Add DECtalk shortcuts if available
        if self.dectalk_native and self.dectalk_native.is_available():
            dectalk_shortcuts = {
                'paul': '[DECtalk] Perfect Paul',
                'betty': '[DECtalk] Betty',
                'harry': '[DECtalk] Harry',
                'frank': '[DECtalk] Frank',
                'dennis': '[DECtalk] Dennis',
                'kit': '[DECtalk] Kit',
                'ursula': '[DECtalk] Ursula',
                'rita': '[DECtalk] Rita',
                'wendy': '[DECtalk] Wendy',
                'sings': '[DECtalk] DECtalk Sings'
            }
            for cmd, voice in dectalk_shortcuts.items():
                self.voice_commands[cmd] = voice
        
        # Save to config
        self.config['voice_commands'] = self.voice_commands
        self.save_config()
        
        # Update GUI if voice commands tab exists
        if hasattr(self, 'voice_commands_text'):
            # Refresh the voice commands display
            self.voice_commands_text.delete("1.0", tk.END)
            for cmd, voice in sorted(self.voice_commands.items()):
                self.voice_commands_text.insert(tk.END, f"{cmd}: {voice}\n")
        
        logger.info(f"Generated {len(self.voice_commands)} voice commands")
    
    def export_config(self):
        """Export configuration to file with selective export options"""
        # Create export dialog
        export_dialog = tk.Toplevel(self.root)
        export_dialog.title("Export Configuration")
        export_dialog.geometry("400x300")
        
        ttk.Label(export_dialog, text="Select sections to export:", font=('Arial', 10, 'bold')).pack(pady=10)
        
        # Checkboxes for sections
        export_options = {}
        sections = [
            ('games', 'Game Settings (log paths, admins, blocked)'),
            ('voices', 'Voice Settings'),
            ('voice_commands', 'Voice Commands'),
            ('user_voices', 'User Voice Preferences'),
            ('announcements', 'Announcements'),
            ('auto_block', 'Auto-block Settings'),
            ('ui', 'UI Settings (collapsed states)'),
            ('audio_device', 'Audio Device Settings')
        ]
        
        for key, label in sections:
            var = tk.BooleanVar(value=True)
            export_options[key] = var
            ttk.Checkbutton(export_dialog, text=label, variable=var).pack(anchor=tk.W, padx=20)
        
        def do_export():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                initialfile=f"tts_config_{time.strftime('%Y%m%d_%H%M%S')}.json"
            )
            
            if file_path:
                try:
                    # Build export config with selected sections
                    export_config = {'version': '2.0'}
                    for key, var in export_options.items():
                        if var.get() and key in self.config:
                            export_config[key] = self.config[key]
                    
                    with open(file_path, 'w') as f:
                        json.dump(export_config, f, indent=2)
                    
                    messagebox.showinfo("Export Successful", f"Config exported to:\n{file_path}")
                    logger.info(f"Config exported to {file_path}")
                    export_dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Export Failed", f"Failed to export config:\n{e}")
                    logger.error(f"Failed to export config: {e}")
        
        ttk.Button(export_dialog, text="Export", command=do_export).pack(pady=20)
        ttk.Button(export_dialog, text="Cancel", command=export_dialog.destroy).pack()
    
    def import_config(self):
        """Import configuration from file with backup"""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Create backup of current config
                backup_path = self.get_config_path().parent / f"config_backup_{time.strftime('%Y%m%d_%H%M%S')}.json"
                with open(backup_path, 'w') as f:
                    json.dump(self.config, f, indent=2)
                logger.info(f"Created config backup at {backup_path}")
                
                with open(file_path, 'r') as f:
                    imported_config = json.load(f)
                
                # Merge with defaults to ensure all required fields exist
                if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                    default_config_path = Path(sys._MEIPASS) / 'default_config.json'
                else:
                    default_config_path = Path(__file__).parent / 'default_config.json'
                
                if default_config_path.exists():
                    with open(default_config_path, 'r') as f:
                        default_config = json.load(f)
                    self.config = self.merge_configs(default_config, imported_config)
                else:
                    self.config = imported_config
                
                # Save and reload
                self.save_config()
                self.load_settings()
                
                messagebox.showinfo("Import Successful", 
                                  "Config imported successfully.\nSome settings may require restart.")
                logger.info(f"Config imported from {file_path}")
            except Exception as e:
                messagebox.showerror("Import Failed", f"Failed to import config:\n{e}")
                logger.error(f"Failed to import config: {e}")
    
    def reset_config(self):
        """Reset configuration to defaults"""
        result = messagebox.askyesno("Reset Configuration", 
                                    "This will reset all settings to defaults.\nContinue?")
        
        if result:
            # Load default config
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                default_config_path = Path(sys._MEIPASS) / 'default_config.json'
            else:
                default_config_path = Path(__file__).parent / 'default_config.json'
            
            if default_config_path.exists():
                try:
                    with open(default_config_path, 'r') as f:
                        self.config = json.load(f)
                    
                    # Preserve TF2 log path if it exists
                    current_log = self.log_path_var.get()
                    if current_log and Path(current_log).exists():
                        self.config['log_path'] = current_log
                    
                    self.save_config()
                    self.load_settings()
                    
                    messagebox.showinfo("Reset Complete", "Settings reset to defaults.")
                    logger.info("Config reset to defaults")
                except Exception as e:
                    messagebox.showerror("Reset Failed", f"Failed to reset config:\n{e}")
                    logger.error(f"Failed to reset config: {e}")
    
        
    def save_config(self):
        """Save config"""
        config_path = self.get_config_path()
        
        try:
            with open(config_path, 'w') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            logger.error(f"Failed to save config: {e}")
    
    def on_random_voice_toggle(self):
        """Handle random voice toggle change"""
        self.random_voice_enabled = self.random_voice_var.get()
        self.config['random_voice_enabled'] = self.random_voice_enabled
        self.save_config()
        logger.info(f"Random voice for new users: {'enabled' if self.random_voice_enabled else 'disabled'}")
    
    def open_random_voice_exclusions(self):
        """Open dialog to configure random voice exclusions"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Random Voice Exclusions")
        dialog.geometry("400x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Select voices to EXCLUDE from random assignment:", 
                 font=('Arial', 10, 'bold')).pack(pady=10)
        ttk.Label(dialog, text="(Checked voices will NOT be assigned to new users)").pack()
        
        # Buttons FIRST (at bottom, but pack before canvas so they get space)
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        
        def save_exclusions():
            self.random_voice_exclusions = [
                voice for voice, var in self.exclusion_vars.items() if var.get()
            ]
            self.config['random_voice_exclusions'] = self.random_voice_exclusions
            self.save_config()
            logger.info(f"Random voice exclusions updated: {len(self.random_voice_exclusions)} voices excluded")
            dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_exclusions).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # Create scrollable frame for checkboxes
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Get all available voices from voice_commands
        self.exclusion_vars = {}
        all_voices = set()
        
        # Add SAPI voices from voice_commands
        for cmd, voice_name in self.voice_commands.items():
            if voice_name and voice_name.strip():
                all_voices.add(voice_name)
        
        # Add DECtalk voices
        if hasattr(self, 'dectalk_profiles'):
            for profile_name in self.dectalk_profiles.keys():
                all_voices.add(f"[DECtalk] {profile_name}")
        
        # Create checkboxes for each voice
        for voice_name in sorted(all_voices):
            var = tk.BooleanVar(value=voice_name in self.random_voice_exclusions)
            self.exclusion_vars[voice_name] = var
            ttk.Checkbutton(scrollable_frame, text=voice_name, variable=var).pack(anchor=tk.W, padx=10, pady=2)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def get_random_voice_for_user(self, username):
        """Assign a random voice to a new user"""
        import random
        
        # Get all available voices
        available_voices = []
        
        # Add SAPI voices from voice_commands
        for cmd, voice_name in self.voice_commands.items():
            if voice_name and voice_name.strip():
                if voice_name not in self.random_voice_exclusions:
                    available_voices.append(voice_name)
        
        # Add DECtalk voices (if not excluded)
        if hasattr(self, 'dectalk_profiles') and self.dectalk_profiles:
            for profile_name in self.dectalk_profiles.keys():
                dectalk_voice = f"[DECtalk] {profile_name}"
                if dectalk_voice not in self.random_voice_exclusions:
                    available_voices.append(dectalk_voice)
        
        if not available_voices:
            logger.warning("No voices available for random assignment (all excluded?)")
            return None
        
        # Pick a random voice
        chosen_voice = random.choice(available_voices)
        
        # Save it as the user's preference
        self.user_voice_preferences[username] = chosen_voice
        self.config['user_voice_preferences'] = self.user_voice_preferences
        self.save_config()
        
        # Update the UI listbox if it exists
        if hasattr(self, 'user_voices_listbox'):
            self.user_voices_listbox.insert(tk.END, f"{username}: {chosen_voice}")
        
        logger.info(f"Assigned random voice to new user {username}: {chosen_voice}")
        return chosen_voice
            
    def setup_gui(self):
        """Setup GUI with enhanced audio device selection"""
        self.root = tk.Tk()
        self.root.title("TTS_GUI - WASAPI Enhanced")
        self.root.geometry("1100x1250")
        self.root.minsize(1100, 1100)  # Set minimum size to prevent resize issues
        
        # Style
        style = ttk.Style()
        style.theme_use('default')
        
        # Top frame with chat display and controls
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left side - chat display
        left_frame = ttk.Frame(top_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Control buttons
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        self.start_btn = ttk.Button(button_frame, text="Start TTS", command=self.start_tts)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        
        self.stop_btn = ttk.Button(button_frame, text="Stop TTS", command=self.stop_tts, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        
        self.reload_btn = ttk.Button(button_frame, text="Reload TTS", command=self.reload_tts)
        self.reload_btn.pack(side=tk.LEFT, padx=2)
        
        self.test_audio_btn = ttk.Button(button_frame, text="Test Audio", command=self.test_audio_device)
        self.test_audio_btn.pack(side=tk.LEFT, padx=2)
        
        # Chat display
        self.chat_display = scrolledtext.ScrolledText(left_frame, height=15, width=50)
        self.chat_display.pack(fill=tk.BOTH, expand=True)
        
        # Right side - Admin and Blocked lists
        right_frame = ttk.Frame(top_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)
        
        # ADMINS section
        admin_label = ttk.Label(right_frame, text="ADMINS", font=('Arial', 10, 'bold'))
        admin_label.pack()
        
        self.admin_listbox = tk.Listbox(right_frame, height=8, width=30)
        self.admin_listbox.pack(pady=5)
        
        admin_btn_frame = ttk.Frame(right_frame)
        admin_btn_frame.pack()
        
        ttk.Button(admin_btn_frame, text="✓ Save", command=self.save_admins).pack(side=tk.LEFT, padx=2)
        ttk.Button(admin_btn_frame, text="✖ Remove", command=self.remove_admin).pack(side=tk.LEFT, padx=2)
        ttk.Button(admin_btn_frame, text="+ Add", command=self.add_admin).pack(side=tk.LEFT, padx=2)
        
        # BLOCKED section
        blocked_label = ttk.Label(right_frame, text="BLOCKED", font=('Arial', 10, 'bold'))
        blocked_label.pack(pady=(10, 0))
        
        self.blocked_listbox = tk.Listbox(right_frame, height=8, width=30)
        self.blocked_listbox.pack(pady=5)
        
        blocked_btn_frame = ttk.Frame(right_frame)
        blocked_btn_frame.pack()
        
        ttk.Button(blocked_btn_frame, text="✓ Save", command=self.save_blocked).pack(side=tk.LEFT, padx=2)
        ttk.Button(blocked_btn_frame, text="✖ Remove", command=self.remove_blocked).pack(side=tk.LEFT, padx=2)
        ttk.Button(blocked_btn_frame, text="+ Add", command=self.add_blocked).pack(side=tk.LEFT, padx=2)
        
        # USER VOICE PREFERENCES section
        user_voices_label = ttk.Label(right_frame, text="USER VOICES", font=('Arial', 10, 'bold'))
        user_voices_label.pack(pady=(10, 0))
        
        self.user_voices_listbox = tk.Listbox(right_frame, height=8, width=30)
        self.user_voices_listbox.pack(pady=5)
        
        user_voices_btn_frame = ttk.Frame(right_frame)
        user_voices_btn_frame.pack()
        
        ttk.Button(user_voices_btn_frame, text="✓ Save", command=self.save_user_voices).pack(side=tk.LEFT, padx=2)
        ttk.Button(user_voices_btn_frame, text="✖ Remove", command=self.remove_user_voice).pack(side=tk.LEFT, padx=2)
        ttk.Button(user_voices_btn_frame, text="+ Add", command=self.add_user_voice).pack(side=tk.LEFT, padx=2)
        ttk.Button(user_voices_btn_frame, text="✎ Edit", command=self.edit_user_voice).pack(side=tk.LEFT, padx=2)
        
        # Long-form text input section (for DECtalk singing and long messages)
        # Use PanedWindow for resizable long-form text
        longform_paned = tk.PanedWindow(self.root, orient=tk.VERTICAL, sashwidth=5, sashrelief=tk.RAISED)
        longform_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        longform_frame = ttk.LabelFrame(longform_paned, text="Long-form Text Input (DECtalk Singing/Long Messages)")
        
        # Text input area with resizable height
        text_input_frame = ttk.Frame(longform_frame)
        text_input_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.longform_text = scrolledtext.ScrolledText(text_input_frame, height=4, width=80, wrap=tk.WORD)
        self.longform_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add character counter
        char_count_frame = ttk.Frame(text_input_frame)
        char_count_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)
        
        self.char_count_label = ttk.Label(char_count_frame, text="0 chars")
        self.char_count_label.pack()
        
        # Update character count on text change
        def update_char_count(event=None):
            text = self.longform_text.get("1.0", tk.END).strip()
            self.char_count_label.config(text=f"{len(text)} chars")
        
        self.longform_text.bind("<KeyRelease>", update_char_count)
        
        longform_paned.add(longform_frame, minsize=100)
        
        # Controls for long-form text
        longform_controls = ttk.Frame(longform_frame)
        longform_controls.pack(fill=tk.X, padx=5, pady=2)
        
        # Voice selector for long-form text
        ttk.Label(longform_controls, text="Voice:").pack(side=tk.LEFT, padx=2)
        self.longform_voice_var = tk.StringVar(value="Default")
        self.longform_voice_combo = ttk.Combobox(longform_controls, textvariable=self.longform_voice_var, 
                                                 width=30, state='readonly')
        self.longform_voice_combo.pack(side=tk.LEFT, padx=2)
        
        ttk.Button(longform_controls, text="Speak Long Text", command=self.speak_longform).pack(side=tk.LEFT, padx=5)
        ttk.Button(longform_controls, text="Clear", command=lambda: self.longform_text.delete("1.0", tk.END)).pack(side=tk.LEFT, padx=2)
        
        # TF2 output controls
        self.tf2_output_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(longform_controls, text="Output to TF2 Chat (110 char chunks)", 
                       variable=self.tf2_output_var).pack(side=tk.LEFT, padx=10)
        
        ttk.Label(longform_controls, text="Note: TF2 has 127 char limit. Text will be split into chunks.").pack(side=tk.LEFT, padx=5)
        
        # Quick Speak/Apply controls - placed between longform and tabs
        quick_speak_frame = ttk.Frame(longform_paned)
        
        test_frame = ttk.Frame(quick_speak_frame)
        test_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(test_frame, text="Quick Speak:").pack(side=tk.LEFT, padx=5)
        self.test_entry = ttk.Entry(test_frame, width=60)
        self.test_entry.pack(side=tk.LEFT, padx=5)
        self.test_entry.insert(0, "Enter your text here")
        
        ttk.Button(test_frame, text="Speak", command=self.speak_test).pack(side=tk.LEFT, padx=2)
        ttk.Button(test_frame, text="✓ Apply", command=self.apply_settings).pack(side=tk.LEFT, padx=2)
        
        longform_paned.add(quick_speak_frame, minsize=50)
        
        # Bottom section - Tabs
        self.notebook = ttk.Notebook(longform_paned)
        longform_paned.add(self.notebook, minsize=250)
        
        # Create tabs
        self.create_settings_tab()
        self.create_audio_devices_tab()
        self.create_voice_commands_tab()
        self.create_dectalk_tab()
        self.create_announcements_tab()
        self.create_auto_block_tab()
        self.create_help_tab()
        self.create_testing_tab()
        
        
    def create_settings_tab(self):
        """Settings tab with config import/export"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Settings")
        settings_frame = ttk.Frame(tab)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Game selection
        ttk.Label(settings_frame, text="Game").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        self.game_combo = ttk.Combobox(settings_frame, values=["Team Fortress 2", "Deep Rock Galactic"], width=50)
        self.game_combo.grid(row=0, column=1, padx=10, pady=5)
        self.game_combo.set(self.config.get('current_game', 'Team Fortress 2'))
        self.game_combo.bind('<<ComboboxSelected>>', self.on_game_changed)
        
        # TTS Command Prefix - moved from Voice Commands to Settings
        ttk.Label(settings_frame, text="TTS Command Prefix").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        self.tts_prefix_entry = ttk.Entry(settings_frame, width=50)
        self.tts_prefix_entry.grid(row=1, column=1, padx=10, pady=5)
        self.tts_prefix_entry.insert(0, self.get_current_game_config().get('tts_command_prefix', '!tts'))
        ttk.Label(settings_frame, text="Command users type in game chat (e.g., !tts, !speak). DRG strips first !tts.").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        
        # Game Log File
        ttk.Label(settings_frame, text="Game Log File").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        self.log_path_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.log_path_var, width=50).grid(row=2, column=1, padx=10, pady=5)
        
        # Admin Username
        ttk.Label(settings_frame, text="Admin Username").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        self.admin_username_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.admin_username_var, width=50).grid(row=3, column=1, padx=10, pady=5)
        
        # Private Mode
        ttk.Label(settings_frame, text="Private Mode").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        self.private_mode_var = tk.BooleanVar()
        ttk.Checkbutton(settings_frame, text="Active", variable=self.private_mode_var).grid(row=4, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Auto Block
        ttk.Label(settings_frame, text="Auto Block").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        self.auto_block_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(settings_frame, text="Active", variable=self.auto_block_var).grid(row=5, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Random Voice for New Users
        ttk.Label(settings_frame, text="Random Voice (New Users)").grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
        random_voice_frame = ttk.Frame(settings_frame)
        random_voice_frame.grid(row=6, column=1, sticky=tk.W, padx=10, pady=5)
        self.random_voice_var = tk.BooleanVar(value=self.random_voice_enabled)
        ttk.Checkbutton(random_voice_frame, text="Active", variable=self.random_voice_var, 
                       command=self.on_random_voice_toggle).pack(side=tk.LEFT)
        ttk.Button(random_voice_frame, text="Configure Exclusions...", 
                  command=self.open_random_voice_exclusions).pack(side=tk.LEFT, padx=10)
        
        # Default Voice
        ttk.Label(settings_frame, text="Default Voice").grid(row=7, column=0, sticky=tk.W, padx=10, pady=5)
        self.default_voice_combo = ttk.Combobox(settings_frame, width=50)
        self.default_voice_combo.grid(row=7, column=1, padx=10, pady=5)
        
        # DECtalk Voice Volume
        ttk.Label(settings_frame, text="DECtalk Voice Volume").grid(row=8, column=0, sticky=tk.W, padx=10, pady=5)
        self.dectalk_volume_var = tk.DoubleVar(value=self.config.get('dectalk_volume', 0.7))  # Default to 70% for DECtalk
        dectalk_volume_frame = ttk.Frame(settings_frame)
        dectalk_volume_frame.grid(row=8, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Config Import/Export buttons
        config_frame = ttk.LabelFrame(settings_frame, text="Configuration Management")
        config_frame.grid(row=9, column=0, columnspan=3, sticky=tk.W+tk.E, padx=10, pady=10)
        
        ttk.Button(config_frame, text="Export Config", command=self.export_config).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(config_frame, text="Import Config", command=self.import_config).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(config_frame, text="Reset to Defaults", command=self.reset_config).pack(side=tk.LEFT, padx=5, pady=5)
        self.dectalk_volume_slider = ttk.Scale(dectalk_volume_frame, from_=0.0, to=1.0, 
                                               variable=self.dectalk_volume_var, 
                                               orient=tk.HORIZONTAL, length=300)
        self.dectalk_volume_slider.pack(side=tk.LEFT)
        self.dectalk_volume_label = ttk.Label(dectalk_volume_frame, text=f"{int(self.dectalk_volume_var.get() * 100)}%")
        self.dectalk_volume_label.pack(side=tk.LEFT, padx=10)
        self.dectalk_volume_var.trace('w', self.update_dectalk_volume_label)
        
        # Populate voices
        if self.voices:
            self.refresh_voice_combos()
            
    def create_audio_devices_tab(self):
        """Audio Devices tab with real Windows devices"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Audio Devices")
        
        # Description
        desc_label = ttk.Label(tab, text="Select audio output device for TTS:", font=('Arial', 10))
        desc_label.pack(pady=10)
        
        # Device list frame
        list_frame = ttk.Frame(tab)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Treeview for devices
        columns = ('Device', 'Channels', 'Sample Rate', 'API')
        self.device_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=10)
        
        # Define columns
        self.device_tree.heading('Device', text='Device Name')
        self.device_tree.heading('Channels', text='Channels')
        self.device_tree.heading('Sample Rate', text='Sample Rate')
        self.device_tree.heading('API', text='API')
        
        self.device_tree.column('Device', width=400)
        self.device_tree.column('Channels', width=80)
        self.device_tree.column('Sample Rate', width=100)
        self.device_tree.column('API', width=150)
        
        # Populate devices
        if not self.audio_devices:
            # Add a message about no devices
            self.device_tree.insert('', tk.END, values=(
                "No audio devices detected - using default",
                "2",
                "48000 Hz",
                "Default"
            ))
            logger.warning("No audio devices found in UI, showing default message")
        else:
            for device in self.audio_devices:
                default = " [DEFAULT]" if device.get('is_default', False) else ""
                self.device_tree.insert('', tk.END, values=(
                    device['name'] + default,
                    device.get('channels', 2),
                    f"{int(device.get('sample_rate', 48000))} Hz",
                    device.get('api', 'Unknown')
                ), tags=('default',) if device.get('is_default', False) else ())
            
        # Highlight default device
        self.device_tree.tag_configure('default', background='lightgreen')
        
        self.device_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.device_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.device_tree.configure(yscrollcommand=scrollbar.set)
        
        # Control buttons
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Select Device", command=self.select_audio_device).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Test Selected", command=self.test_selected_device).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Refresh Devices", command=self.refresh_audio_devices).pack(side=tk.LEFT, padx=5)
        
        # Current device label
        self.current_device_label = ttk.Label(tab, text="Current Device: Default", font=('Arial', 10, 'bold'))
        self.current_device_label.pack(pady=10)
        
    def create_voice_commands_tab(self):
        """Voice Commands tab - editable"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Voice Commands")
        
        # Create two paned sections
        paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left side - Voice Commands
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=2)
        
        # Info label
        info_label = ttk.Label(left_frame, text="Custom voice commands: Use any alphanumeric trigger (e.g., /c, /sam, /v0)")
        info_label.pack(pady=5)
        
        # Voice Toggle Command configuration
        toggle_frame = ttk.Frame(left_frame)
        toggle_frame.pack(pady=5)
        
        ttk.Label(toggle_frame, text="User Voice Toggle Command:").pack(side=tk.LEFT, padx=5)
        self.voice_toggle_entry = ttk.Entry(toggle_frame, width=10)
        self.voice_toggle_entry.pack(side=tk.LEFT, padx=5)
        self.voice_toggle_entry.insert(0, self.voice_toggle_command)
        
        ttk.Label(toggle_frame, text="(e.g. /vt to let users set their default voice)").pack(side=tk.LEFT, padx=5)
        
        # Create table
        columns = ('COMMAND', 'VOICE')
        self.voice_tree = ttk.Treeview(left_frame, columns=columns, show='headings', height=10)
        
        self.voice_tree.heading('COMMAND', text='COMMAND')
        self.voice_tree.heading('VOICE', text='VOICE')
        
        self.voice_tree.column('COMMAND', width=80)
        self.voice_tree.column('VOICE', width=300)
        
        # Add voice commands
        for cmd, voice in self.voice_commands.items():
            self.voice_tree.insert('', tk.END, values=(cmd, voice), tags=(cmd,))
            
        self.voice_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Bind double-click to edit and Delete key to remove
        self.voice_tree.bind('<Double-Button-1>', self.edit_voice_command)
        self.voice_tree.bind('<Delete>', lambda e: self.remove_voice_command())
        
        # Button frame with enhanced controls
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(pady=5)
        
        # Command management buttons
        ttk.Button(btn_frame, text="Add Command", command=self.add_voice_command).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Edit Command", command=self.edit_voice_command_full).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Remove Command", command=self.remove_voice_command).pack(side=tk.LEFT, padx=2)
        ttk.Separator(btn_frame, orient='vertical').pack(side=tk.LEFT, padx=5, fill='y')
        ttk.Button(btn_frame, text="Save Commands", command=self.save_voice_commands).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Reset to Defaults", command=self.reset_voice_commands).pack(side=tk.LEFT, padx=2)
        
        # Right side - Voice Management
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)
        
        ttk.Label(right_frame, text="Available Voices:").pack(pady=5)
        
        # Voice list with checkboxes
        list_frame = ttk.Frame(right_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Voice listbox
        self.available_voices_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, selectmode=tk.MULTIPLE)
        self.available_voices_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.available_voices_listbox.yview)
        
        # Hidden voices tracking
        self.hidden_voices = self.config.get('hidden_voices', [])
        
        # Populate voice list
        self.refresh_available_voices()
        
        # Management buttons
        mgmt_frame = ttk.Frame(right_frame)
        mgmt_frame.pack(pady=5)
        
        ttk.Button(mgmt_frame, text="Hide Selected", command=self.hide_selected_voices).pack(side=tk.LEFT, padx=2)
        ttk.Button(mgmt_frame, text="Show All", command=self.show_all_voices).pack(side=tk.LEFT, padx=2)
        ttk.Button(mgmt_frame, text="Check Duplicates", command=self.check_duplicate_voices).pack(side=tk.LEFT, padx=2)
        
    def get_default_dectalk_profiles(self):
        """Get default DECtalk voice profiles"""
        return {
            "Perfect Paul": "[:np]",
            "Betty": "[:nb]",
            "Harry": "[:nh]",
            "Frank": "[:nf]",
            "Dennis": "[:nd]",
            "Kit": "[:nk]",
            "Ursula": "[:nu]",
            "Rita": "[:nr]",
            "Wendy": "[:nw]",
            "Doctor Dennis": "[:nd][:dv gv 85]",
            "Huge Harry": "[:nh][:dv gv 100]",
            "Beautiful Betty": "[:nb][:dv gv 90]",
            "Frail Frank": "[:nf][:dv gv 80]",
            "Kit the Kid": "[:nk][:rate 250]",
            "Uppity Ursula": "[:nu][:dv gv 95]",
            "Rough Rita": "[:nr][:dv gv 85]",
            "Whispering Wendy": "[:nw][:dv gv 75]",
            "Variable Paul": "[:np][:rate 200]",
            "DECtalk Sings": "[:np][:rate 120][:pitch 200]"
        }
    
    def create_dectalk_tab(self):
        """DECtalk extended voices tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="DECtalk")
        
        # Enable/Disable DECtalk
        enable_frame = ttk.Frame(tab)
        enable_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.dectalk_enabled_var = tk.BooleanVar(value=self.dectalk_enabled)
        ttk.Checkbutton(enable_frame, text="Enable DECtalk Extended Voices", 
                       variable=self.dectalk_enabled_var,
                       command=self.toggle_dectalk).pack(side=tk.LEFT)
        
        ttk.Label(enable_frame, text="(Built-in emulation - no installation required)").pack(side=tk.LEFT, padx=(20, 0))
        
        # DECtalk profiles section
        profiles_frame = ttk.LabelFrame(tab, text="DECtalk Voice Profiles")
        profiles_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Profiles list
        list_frame = ttk.Frame(profiles_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left side - Profile list
        left_frame = ttk.Frame(list_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ttk.Label(left_frame, text="Voice Profiles:").pack()
        
        scrollbar = ttk.Scrollbar(left_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.dectalk_profiles_listbox = tk.Listbox(left_frame, yscrollcommand=scrollbar.set, height=15)
        self.dectalk_profiles_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.dectalk_profiles_listbox.yview)
        
        # Bind selection event
        self.dectalk_profiles_listbox.bind('<<ListboxSelect>>', self.on_dectalk_profile_select)
        
        # Right side - Profile editor
        right_frame = ttk.Frame(list_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(20, 0))
        
        ttk.Label(right_frame, text="Profile Name:").pack(anchor=tk.W)
        self.dectalk_profile_name_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.dectalk_profile_name_var, width=30).pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(right_frame, text="DECtalk Code:").pack(anchor=tk.W)
        self.dectalk_profile_code_var = tk.StringVar()
        code_entry = ttk.Entry(right_frame, textvariable=self.dectalk_profile_code_var, width=30)
        code_entry.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(right_frame, text="Example: [:np] for Paul, [:nb] for Betty", font=('Arial', 8)).pack(anchor=tk.W)
        ttk.Label(right_frame, text="Advanced: [:np][:rate 200][:pitch 150]", font=('Arial', 8)).pack(anchor=tk.W)
        
        # Profile buttons
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Add Profile", command=self.add_dectalk_profile).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Update Profile", command=self.update_dectalk_profile).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Remove Profile", command=self.remove_dectalk_profile).pack(side=tk.LEFT, padx=2)
        
        # Test button
        ttk.Button(right_frame, text="Test Selected Profile", command=self.test_dectalk_profile, width=25).pack(pady=10)
        
        # Management buttons
        mgmt_frame = ttk.Frame(profiles_frame)
        mgmt_frame.pack(pady=10)
        
        ttk.Button(mgmt_frame, text="Save Profiles", command=self.save_dectalk_profiles).pack(side=tk.LEFT, padx=5)
        ttk.Button(mgmt_frame, text="Reset to Defaults", command=self.reset_dectalk_profiles).pack(side=tk.LEFT, padx=5)
        ttk.Button(mgmt_frame, text="Export Profiles", command=self.export_dectalk_profiles).pack(side=tk.LEFT, padx=5)
        ttk.Button(mgmt_frame, text="Import Profiles", command=self.import_dectalk_profiles).pack(side=tk.LEFT, padx=5)
        
        # Info section
        info_frame = ttk.LabelFrame(tab, text="Information")
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        
        info_text = ("DECtalk emulation provides classic DECtalk voices using your existing SAPI5 voices.\n"
                    "Profiles will appear as additional voices in the Voice Commands tab.\n"
                    "You can assign them to any voice command slot (e.g., /v 10, /v 11, etc.)")
        ttk.Label(info_frame, text=info_text, wraplength=500).pack(padx=10, pady=10)
        
        # Populate profiles list
        self.refresh_dectalk_profiles()
    
    def create_auto_block_tab(self):
        """Auto Block tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Auto Block")
        
        ttk.Label(tab, text="Auto-block keywords (one per line):").pack(pady=10)
        
        self.auto_block_text = scrolledtext.ScrolledText(tab, height=10, width=60)
        self.auto_block_text.pack(padx=10, pady=10)
        
        ttk.Button(tab, text="Save Keywords", command=self.save_auto_block).pack(pady=5)
    
    def create_help_tab(self):
        """Help tab with instructions and information"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Help")
        
        # Create scrolled text for help content
        help_text = scrolledtext.ScrolledText(tab, wrap=tk.WORD, width=80, height=20)
        help_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        help_content = """TTS FOR TEAM FORTRESS 2 - HELP GUIDE
=====================================

QUICK START:
1. Set your TF2 log path in Settings tab
2. Select your audio output device in Audio Devices tab  
3. Start monitoring to begin

VOICE COMMANDS:
• /v 0-###: Quick voice selection
• Custom commands in Voice Commands tab

USER VOICE PREFERENCES:
• Users set default voice with: /vt [voice name]
• Example: "/vt 0 (Microsoft Sam)"
• Preference saved automatically, and speaks message if one was sent

ADMIN COMMANDS:
• !block [username]: Block a user
• !unblock [username]: Unblock a user
• !admin [username]: Add admin
• !unadmin [username]: Remove admin
• !stop : stop all voice, clear que

SPECIAL FEATURES:
• Longform Text: Type longer messages and Send
• Auto-block: Add keywords to block
• DECtalk: Classic voices with profiles

AUDIO ROUTING:
• Select "Voicemeeter Input" for virtual cable
• Use "Default" for standard speakers

Version: TTS TF2 64-bit v2.0
Build: Full Featured"""
        
        help_text.insert("1.0", help_content)
        help_text.config(state='disabled')  # Make read-only
        
    def select_audio_device(self):
        """Select the highlighted audio device"""
        selection = self.device_tree.selection()
        if selection:
            item = self.device_tree.item(selection[0])
            device_name = item['values'][0].replace(" [DEFAULT]", "")
            
            if self.audio_manager and self.audio_manager.set_device(device_name):
                self.current_device_label.config(text=f"Current Device: {device_name}")
                self.config['audio_device'] = device_name
                messagebox.showinfo("Success", f"Audio device set to: {device_name}")
            else:
                messagebox.showerror("Error", f"Failed to set device: {device_name}")
                
    def test_selected_device(self):
        """Test the selected audio device"""
        selection = self.device_tree.selection()
        if selection:
            item = self.device_tree.item(selection[0])
            device_name = item['values'][0].replace(" [DEFAULT]", "")
            
            if self.audio_manager:
                self.audio_manager.speak(f"Testing audio on {device_name}", device_override=device_name)
        else:
            messagebox.showwarning("No Selection", "Please select a device first")
            
    def test_audio_device(self):
        """Test current audio device"""
        if self.audio_manager:
            self.audio_manager.test_current_device()
            
    def refresh_audio_devices(self):
        """Refresh the list of audio devices"""
        if self.audio_manager:
            self.audio_devices = self.audio_manager.get_devices()
            
            # Clear tree
            for item in self.device_tree.get_children():
                self.device_tree.delete(item)
                
            # Check if we have devices
            if not self.audio_devices:
                self.device_tree.insert('', tk.END, values=(
                    "No audio devices detected - using default",
                    "2",
                    "48000 Hz",
                    "Default"
                ))
                logger.warning("No audio devices found after refresh")
            else:
                # Repopulate
                for device in self.audio_devices:
                    default = " [DEFAULT]" if device.get('is_default', False) else ""
                    self.device_tree.insert('', tk.END, values=(
                        device['name'] + default,
                        device.get('channels', 2),
                        f"{int(device.get('sample_rate', 48000))} Hz",
                        device.get('api', 'Unknown')
                    ), tags=('default',) if device.get('is_default', False) else ())
                
    def load_settings(self):
        """Load all settings"""
        self.log_path_var.set(self.config.get('log_path', ''))
        self.admin_username_var.set(self.config.get('admin_username', 'Huskiefluffs'))
        self.private_mode_var.set(self.config.get('private_mode', False))
        
        # Populate long-form voice combo now that GUI is ready
        if hasattr(self, 'longform_voice_combo') and self.audio_manager:
            self.populate_longform_voice_combo()
        
        # Set default voice from config
        default_voice = self.config.get('default_voice', '')
        if default_voice and self.audio_manager:
            # Apply the default voice to the audio manager
            for voice in self.voices:
                if voice.name == default_voice:
                    if self.current_voice_id != voice.id:
                        self.audio_manager.set_voice(voice.id)
                        self.current_voice_id = voice.id
                    break
        
        # Handle auto_block
        auto_block = self.config.get('auto_block_enabled', self.config.get('auto_block', True))
        if isinstance(auto_block, dict):
            auto_block = auto_block.get('enabled', True)
        self.auto_block_var.set(auto_block)
        
        # Clear and load lists
        self.admin_listbox.delete(0, tk.END)
        self.blocked_listbox.delete(0, tk.END)
        
        for admin in self.config.get('admins', []):
            self.admin_listbox.insert(tk.END, admin)
            
        for blocked in self.config.get('blocked', []):
            self.blocked_listbox.insert(tk.END, blocked)
            
        # Load user voice preferences
        if hasattr(self, 'user_voices_listbox'):
            self.user_voices_listbox.delete(0, tk.END)
            for username, voice in self.user_voice_preferences.items():
                self.user_voices_listbox.insert(tk.END, f"{username}: {voice}")
            
        # Load auto-block keywords
        keywords = self.config.get('auto_block_keywords', [])
        if keywords and hasattr(self, 'auto_block_text'):
            self.auto_block_text.delete(1.0, tk.END)
            self.auto_block_text.insert(1.0, '\n'.join(keywords))
            
        # Set audio device
        device_name = self.config.get('audio_device', 'Default')
        if self.audio_manager:
            self.audio_manager.set_device(device_name)
            self.current_device_label.config(text=f"Current Device: {device_name}")
            
    def on_chat_message(self, msg):
        """Handle chat message"""
        
        # Debug logging for all !block messages
        if "!block" in msg.get('message', '').lower():
            logger.info(f"ON_CHAT_MESSAGE: Received !block message from {msg.get('username', 'unknown')}: '{msg.get('message', '')}'")
            logger.info(f"ON_CHAT_MESSAGE: Full message dict: {msg}")
        
        # Format and display
        timestamp = time.strftime("%H:%M:%S")
        prefix = ""
        
        if msg['is_dead']:
            prefix += "*DEAD*"
        if msg['is_team']:
            prefix += "(TEAM) "
            
        display_text = f"{prefix}{msg['username']} : {msg['message']}\n"
        
        # Add to chat display
        self.root.after(0, lambda: self.chat_display.insert(tk.END, display_text))
        self.root.after(0, lambda: self.chat_display.see(tk.END))
        
        # Don't track last_speaker here - only track when actually speaking
        
        # FIRST: Check if blocked - must check exact username match
        blocked_users = [self.blocked_listbox.get(i) for i in range(self.blocked_listbox.size())]
        if msg['username'] in blocked_users:
            logger.info(f"Blocked user {msg['username']} tried to speak")
            return
            
        # SECOND: Check auto-block keywords BEFORE any other processing
        if self.auto_block_var.get():
            keywords = self.config.get('auto_block_keywords', [])
            message_lower = msg['message'].lower()
            for keyword in keywords:
                if keyword and keyword.lower() in message_lower:
                    logger.info(f"Auto-blocking {msg['username']} for keyword: {keyword}")
                    # Block the user and clear their messages
                    self.clear_user_from_queue(msg['username'], add_to_blocked=True)
                    self.save_blocked()
                    
                    # Queue the announcement to play next (use config text)
                    announcement_text = self.get_announcement_text('AUTOBLOCK', username=msg['username'])
                    if announcement_text:
                        # Insert announcement at front of queue
                        temp_messages = []
                        while not self.message_queue.empty():
                            try:
                                temp_messages.append(self.message_queue.get_nowait())
                            except:
                                break
                        
                        # Add the announcement
                        announcement_voice = self.config.get('announcement_voice', '')
                        self.message_queue.put({'text': announcement_text, 'voice': announcement_voice, 'username': '__announcement__'})
                        
                        # Re-add other messages
                        for temp_msg in temp_messages:
                            self.message_queue.put(temp_msg)
                    
                    # Update UI on main thread
                    self.root.after(0, lambda u=msg['username']: self.update_ui_after_autoblock(u))
                    return
        
        # THIRD: Check for admin commands FIRST (before !tts processing)
        message = msg['message'].strip()
        admins = [self.admin_listbox.get(i) for i in range(self.admin_listbox.size())]
        
        # Check for TTS stop command
        current_tts_prefix = self.get_current_game_config().get('tts_command_prefix', '!tts')
        tts_stop_command = f"{current_tts_prefix} stop"
        if message.lower().strip() == tts_stop_command.lower():
            if msg['username'] in admins:
                self.force_stop_tts()
                return
            else:
                logger.info(f"Non-admin {msg['username']} tried to stop TTS")
                return
                
        # Check for !stop command (stops all current speech)
        if message.lower().strip() == '!stop':
            if msg['username'] in admins:
                logger.info(f"Admin {msg['username']} issued !stop command")
                self.stop_all_speech()
                # Show in chat that speech was stopped
                self.chat_display.insert(tk.END, f"[SYSTEM] Speech stopped by {msg['username']}\n")
                self.chat_display.see(tk.END)
                return
            else:
                logger.info(f"Non-admin {msg['username']} tried to use !stop")
                return
                
        # Handle !block and !admin commands  
        # Support both "!block" and "!block add" for compatibility (with trailing spaces)
        if message.lower().strip() == '!block add' or message.lower().strip() == '!block':
            logger.info(f"!block command received from {msg['username']}")
            logger.info(f"Current admins list: {admins}")
            logger.info(f"Is {msg['username']} in admins? {msg['username'] in admins}")
            if msg['username'] in admins:
                logger.info(f"Admin confirmed. Last speaker: {self.last_speaker}, Currently speaking: {self.currently_speaking_user}")
                # Determine who to block - prioritize currently speaking user, then last speaker
                user_to_block = self.currently_speaking_user if self.currently_speaking_user else self.last_speaker
                
                if user_to_block and user_to_block != msg['username']:
                    logger.info(f"Admin block command: blocking {user_to_block}")
                    # Clear the blocked user's messages from queue and add to blocked list
                    self.clear_user_from_queue(user_to_block, add_to_blocked=True)
                    # Save after blocking
                    self.save_blocked()
                    # Queue the announcement to play next (before other users' messages)
                    # Insert at front of queue by temporarily storing and re-adding messages
                    temp_messages = []
                    while not self.message_queue.empty():
                        try:
                            temp_messages.append(self.message_queue.get_nowait())
                        except:
                            break
                    
                    # Add the announcement as the next message (use config text)
                    announcement_text = self.get_announcement_text('BLOCK ADD', username=user_to_block)
                    if announcement_text:
                        # Get announcement voice
                        announcement_voice = self.config.get('announcement_voice', '')
                        self.message_queue.put({'text': announcement_text, 'voice': announcement_voice, 'username': '__announcement__'})
                    
                    # Re-add the other messages
                    for msg in temp_messages:
                        self.message_queue.put(msg)
                else:
                    logger.info("No valid user to block")
            else:
                logger.info(f"User {msg['username']} is not admin")
            return
                
        if message.lower().strip() == '!admin add':
            if msg['username'] in admins:
                # Add last speaker as admin
                if self.last_speaker and self.last_speaker != msg['username']:
                    logger.info(f"Admin command: adding {self.last_speaker} as admin")
                    self.admin_listbox.insert(tk.END, self.last_speaker)
                    self.save_admins()
                    
                    # Queue the announcement to play next (use config text)
                    announcement_text = self.get_announcement_text('ADMIN ADD', username=self.last_speaker)
                    if announcement_text:
                        # Insert announcement at front of queue
                        temp_messages = []
                        while not self.message_queue.empty():
                            try:
                                temp_messages.append(self.message_queue.get_nowait())
                            except:
                                break
                        
                        # Add the announcement
                        announcement_voice = self.config.get('announcement_voice', '')
                        self.message_queue.put({'text': announcement_text, 'voice': announcement_voice, 'username': '__announcement__'})
                        
                        # Re-add other messages
                        for msg in temp_messages:
                            self.message_queue.put(msg)
                else:
                    logger.info("No valid last speaker to add as admin")
                return
                
        if message.lower().strip() == '!block clear':
            if msg['username'] in admins:
                # Remove only the last entry from blocked list
                size = self.blocked_listbox.size()
                if size > 0:
                    last_user = self.blocked_listbox.get(size - 1)
                    logger.info(f"Admin command: removing last blocked user: {last_user}")
                    self.blocked_listbox.delete(size - 1)
                    self.save_blocked()
                    
                    # Queue announcement for unblocking (use config text, not hardcoded)
                    announcement_text = self.get_announcement_text('BLOCK REMOVE', username=last_user)
                    if announcement_text:
                        # Insert announcement at front of queue
                        temp_messages = []
                        while not self.message_queue.empty():
                            try:
                                temp_messages.append(self.message_queue.get_nowait())
                            except:
                                break
                        
                        # Add the announcement
                        announcement_voice = self.config.get('announcement_voice', '')
                        self.message_queue.put({'text': announcement_text, 'voice': announcement_voice, 'username': '__announcement__'})
                        
                        # Re-add other messages
                        for msg in temp_messages:
                            self.message_queue.put(msg)
                else:
                    logger.info("Block list is already empty")
                return
        
        # FOURTH: Check for TTS commands - these should NOT be spoken  
        current_tts_prefix = self.get_current_game_config().get('tts_command_prefix', '!tts')
        tts_command_trigger = f"{current_tts_prefix} "
        
        # Check if message starts with the TTS command prefix
        if message.startswith(tts_command_trigger):
            # This is a TTS command, extract just the text part
            text_to_speak = message[len(tts_command_trigger):].strip()  # Remove prefix
            
            # Debug logging for DRG
            game_type = msg.get('game', 'unknown')
            logger.info(f"Processing !tts command from {game_type}: '{text_to_speak}' by {msg['username']}")
            
            # Check private mode - only admins can use !tts in private mode
            if self.private_mode_var.get():
                admins = [self.admin_listbox.get(i) for i in range(self.admin_listbox.size())]
                if msg['username'] not in admins:
                    logger.info(f"Non-admin {msg['username']} tried to use TTS in private mode")
                    return
                    
            # Check for voice commands in the !tts text
            # Check for any slash command (including /vt, /v, etc.)
            if text_to_speak.startswith('/'):
                self.process_voice_command(text_to_speak, username=msg['username'])
            elif text_to_speak.startswith('v '):
                # Handle old format without slash
                self.process_voice_command(text_to_speak, username=msg['username'])
            else:
                # No voice command - check for user's preferred voice first
                username = msg['username']
                if username and username in self.user_voice_preferences:
                    # Apply user's preferred voice
                    user_voice = self.user_voice_preferences[username]
                    logger.info(f"Applying {username}'s preferred voice: {user_voice}")
                    # Pass the voice to speak so DECtalk voices work correctly
                    self.speak(text_to_speak, voice_name=user_voice, username=username)
                elif username and self.random_voice_enabled:
                    # New user and random voice is enabled - assign them a random voice
                    random_voice = self.get_random_voice_for_user(username)
                    if random_voice:
                        logger.info(f"Assigned random voice to new user {username}: {random_voice}")
                        self.speak(text_to_speak, voice_name=random_voice, username=username)
                    else:
                        # Fallback to default if random assignment failed
                        self.apply_default_voice()
                        self.speak(text_to_speak, username=username)
                else:
                    # Apply system default voice
                    self.apply_default_voice()
                    # Just speak the text after !tts, not the command itself
                    self.speak(text_to_speak, username=username)
            # Track the last speaker when they actually use TTS
            logger.info(f"Setting last_speaker to: {msg['username']}")
            self.last_speaker = msg['username']
            return
            
                
        # Check private mode for regular messages
        if self.private_mode_var.get():
            if msg['username'] not in admins:
                logger.info(f"Private mode: ignoring non-admin {msg['username']}")
                return
                
        # Check for voice commands - support flexible patterns
        import re
        
        # Check for /v [number] format specifically (most common)
        if re.match(r'^/v\s+\d+', message):
            self.process_voice_command(message, username=msg['username'])
            return
            
        # Check for other slash commands: /[trigger]
        if message.startswith('/'):
            # Check if it could be a voice command
            if re.match(r'^/[a-zA-Z0-9_]+', message):
                # Check if we have this command or if it's a v[number] pattern
                match = re.match(r'^/([a-zA-Z0-9_]+)', message)
                if match:
                    trigger = match.group(1)
                    # Check direct trigger or v[number] format
                    if trigger in self.voice_commands or (trigger.startswith('v') and trigger[1:].isdigit()):
                        self.process_voice_command(message, username=msg['username'])
                        return
                        
        # Check for legacy v [number] format without slash
        if message.startswith('v ') and len(message) > 2 and message[2].isdigit():
            self.process_voice_command(message, username=msg['username'])
            return
            
        # Check for direct custom commands (without slash)
        first_word = message.split(' ')[0]
        if first_word in self.voice_commands:
            self.process_voice_command(message, username=msg['username'])
            return
            
        # Regular message - DO NOT speak unless it's a !tts command
        # All non-admin messages require !tts to be spoken
        # (Admin commands like !stop, !block are already handled above)
            
    def process_voice_command(self, message, username=None):
        """Process voice command with flexible pattern support"""
        import re
        original_message = message
        
        # Check if message starts with a slash (voice command indicator)
        if message.startswith('/'):
            # First check for user voice toggle command (e.g., /vt)
            toggle_cmd = self.voice_toggle_command.lstrip('/')  # Remove leading slash for matching
            pattern = rf'^/{re.escape(toggle_cmd)}\s+(\d+)(?:\s+(.*))?$'
            match = re.match(pattern, message)
            if match:
                voice_num = match.group(1)
                text = match.group(2) if match.group(2) else ""
                cmd = f"v {voice_num}"
                
                logger.info(f"User voice toggle command: user={username}, voice={voice_num}")
                
                # Set user's default voice preference
                if username and cmd in self.voice_commands:
                    voice_name = self.voice_commands.get(cmd, '')
                    
                    # Check if voice is actually configured (not empty)
                    if not voice_name:
                        logger.warning(f"Voice command 'v {voice_num}' has no voice mapped")
                        # Still speak the text with default voice if provided
                        if text:
                            self.speak(text, username=username)
                        return
                    
                    self.user_voice_preferences[username] = voice_name
                    
                    # Update UI if exists
                    if hasattr(self, 'user_voices_listbox'):
                        # Find and update or add entry
                        found = False
                        for i in range(self.user_voices_listbox.size()):
                            entry = self.user_voices_listbox.get(i)
                            if entry.startswith(f"{username}:"):
                                self.user_voices_listbox.delete(i)
                                self.user_voices_listbox.insert(i, f"{username}: {voice_name}")
                                found = True
                                break
                        if not found:
                            self.user_voices_listbox.insert(tk.END, f"{username}: {voice_name}")
                    
                    # Save the preference
                    self.config['user_voice_preferences'] = self.user_voice_preferences
                    self.save_config()
                    
                    logger.info(f"Set {username}'s default voice to: {voice_name}")
                    
                    # Speak the text with the new voice if provided
                    if text:
                        self.speak_with_voice(text, voice_name, username)
                else:
                    logger.warning(f"Voice command 'v {voice_num}' is not configured or no username")
                    # Still speak the text with default voice
                    if text:
                        self.speak(text, username=username)
                return
            
            # Then check for /v [number] format (with space)
            match = re.match(r'^/v\s+(\d+)(?:\s+(.*))?$', message)
            if match:
                voice_num = match.group(1)
                text = match.group(2) if match.group(2) else ""
                cmd = f"v {voice_num}"
                
                logger.info(f"Voice command (with space): '{cmd}', Text: '{text}'")
                logger.info(f"Text repr for debugging: {repr(text)}")  # Show exact string representation
                
                if cmd in self.voice_commands:
                    voice_name = self.voice_commands[cmd]
                    logger.info(f"Voice command matched - Command: '{cmd}' -> Voice: '{voice_name}'")
                    
                    if text:
                        # Use temporary voice for this message only
                        self.speak_with_voice(text, voice_name, username)
                    # Note: If no text, we don't speak the command itself
                else:
                    # Command not mapped - don't fall back to direct index
                    logger.warning(f"Voice command 'v {voice_num}' is not configured")
                    # Speak the message with current voice
                    if text:
                        self.speak(text, username=username)
                    else:
                        # Don't speak the command itself
                        logger.info("Undefined voice command, skipping")
                return
            
            # Then check for other patterns: /[command] [text]
            match = re.match(r'^/([a-zA-Z0-9_]+)(?:\s+(.*))?$', message)
            
            if match:
                trigger = match.group(1)
                text = match.group(2) if match.group(2) else ""
                
                logger.info(f"Voice command trigger: '{trigger}', Text: '{text}'")
                
                # Check if we have this command mapped
                if trigger in self.voice_commands:
                    voice_name = self.voice_commands[trigger]
                    logger.info(f"Command '{trigger}' mapped to voice: {voice_name}")
                    
                    if text:
                        # Use temporary voice for this message only
                        self.speak_with_voice(text, voice_name, username)
                    else:
                        # No text provided, don't speak the command itself
                        logger.info("No text after command, skipping speech")
                            
                # Check for legacy v[number] format (v0-v9)
                elif trigger.startswith('v') and len(trigger) > 1:
                    # Try to extract number from v0, v1, etc.
                    try:
                        voice_num = trigger[1:]
                        legacy_cmd = f"v {voice_num}"
                        
                        if legacy_cmd in self.voice_commands:
                            voice_name = self.voice_commands[legacy_cmd]
                            logger.info(f"Legacy command '{legacy_cmd}' mapped to voice: {voice_name}")
                            
                            if text:
                                # Use temporary voice for this message only
                                self.speak_with_voice(text, voice_name, username)
                        else:
                            # Command not mapped - don't use direct index
                            logger.warning(f"Voice command '{legacy_cmd}' is not configured")
                            if text:
                                self.speak(text, username=username)
                            else:
                                logger.info("Undefined voice command, skipping")
                    except (ValueError, IndexError):
                        logger.warning(f"Unknown command trigger: {trigger}")
                        self.speak(original_message)
                else:
                    logger.warning(f"Unknown command trigger: {trigger}")
                    self.speak(original_message)
            else:
                # Slash but no valid command pattern
                self.speak(original_message)
                
        # Support for commands without slash (legacy format)
        elif message.startswith('v '):
            # Legacy format: v [number] [text]
            parts = message.split(' ', 2)
            if len(parts) >= 2:
                voice_num = parts[1]
                cmd = f"v {voice_num}"
                text = parts[2] if len(parts) > 2 else ""
                
                if cmd in self.voice_commands:
                    voice_name = self.voice_commands[cmd]
                    logger.info(f"Legacy format - Mapped to voice: {voice_name}")
                    
                    if text:
                        # Use temporary voice for this message only
                        self.speak_with_voice(text, voice_name, username)
                else:
                    # Command not mapped - don't use direct index
                    logger.warning(f"Voice command 'v {voice_num}' is not configured")
                    if text:
                        self.speak(text)
                    else:
                        logger.info("Undefined voice command, skipping")
            else:
                self.speak(original_message)
        else:
            # Check if it matches any custom command without slash
            parts = message.split(' ', 1)
            if parts[0] in self.voice_commands:
                trigger = parts[0]
                text = parts[1] if len(parts) > 1 else ""
                voice_name = self.voice_commands[trigger]
                
                logger.info(f"Direct command '{trigger}' mapped to voice: {voice_name}")
                
                if text:
                    # Use temporary voice for this message only
                    self.speak_with_voice(text, voice_name, username)
            else:
                # Not a voice command, speak normally
                self.speak(message, username=username)
            
    def apply_default_voice(self):
        """Apply the default voice from config"""
        default_voice = self.config.get('default_voice', '')
        if default_voice and self.audio_manager:
            for voice in self.voices:
                if voice.name == default_voice:
                    # Skip if already using this voice
                    if self.current_voice_id == voice.id:
                        logger.debug(f"Already using default voice: {default_voice}")
                        return True
                    self.audio_manager.set_voice(voice.id)
                    self.current_voice_id = voice.id
                    logger.debug(f"Applied default voice: {default_voice}")
                    return True
        return False
    
    def apply_voice(self, voice_name):
        """Apply specific voice"""
        logger.info(f"apply_voice called with: {voice_name}")
        if self.voices and self.audio_manager:
            # Check if it's a DECtalk profile
            if voice_name.startswith("[DECtalk] "):
                # DECtalk will be handled in speech processing
                logger.info(f"DECtalk profile selected: {voice_name}")
                return True
            
            
            # Try exact match first
            for voice in self.voices:
                if voice.name == voice_name:
                    # Skip if already using this voice
                    if self.current_voice_id == voice.id:
                        logger.debug(f"Already using voice: {voice_name}")
                        return True
                    success = self.audio_manager.set_voice(voice.id)
                    if success:
                        self.current_voice_id = voice.id
                    logger.info(f"Changed voice to: {voice_name} (ID: {voice.id}) - Success: {success}")
                    return success
            # Try partial match
            for voice in self.voices:
                if voice_name.lower() in voice.name.lower() or voice.name.lower() in voice_name.lower():
                    # Skip if already using this voice
                    if self.current_voice_id == voice.id:
                        logger.debug(f"Already using voice: {voice.name}")
                        return True
                    success = self.audio_manager.set_voice(voice.id)
                    if success:
                        self.current_voice_id = voice.id
                    logger.info(f"Changed voice to: {voice.name} (ID: {voice.id}, matched {voice_name}) - Success: {success}")
                    return success
            logger.warning(f"Voice not found: {voice_name}")
        return False
                
    def speak(self, text, voice_name=None, username=None):
        """Queue text for speaking"""
        if self.audio_manager:
            # Add to queue with voice info and username
            logger.info(f"Adding message to speech queue: '{text}' with voice '{voice_name}' from {username}")
            self.message_queue.put({'text': text, 'voice': voice_name, 'username': username})
    
    def populate_longform_voice_combo(self):
        """Populate the long-form voice combo with all available voices"""
        voice_list = ["Default"]
        
        # Add Windows TTS voices
        for voice in self.voices:
            if voice.name not in self.hidden_voices:
                voice_list.append(voice.name)
        
        # Add DECtalk profiles if available
        if self.dectalk_enabled and self.dectalk_profiles:
            for profile_name in sorted(self.dectalk_profiles.keys()):
                voice_list.append(f"[DECtalk] {profile_name}")
        
        
        # Update combo box
        self.longform_voice_combo['values'] = voice_list
        logger.info(f"Populated long-form voice combo with {len(voice_list)} voices")
    
    def speak_longform(self):
        """Handle long-form text input with optional TF2 chat output"""
        text = self.longform_text.get("1.0", tk.END).strip()
        if not text:
            logger.warning("No text entered in long-form input")
            return
        
        # Get selected voice
        selected_voice = self.longform_voice_var.get()
        if selected_voice == "Default":
            selected_voice = None
        
        # Check if TF2 output is enabled
        output_to_tf2 = self.tf2_output_var.get()
        
        if output_to_tf2 and self.log_path_var.get():
            # Split text into 110 character chunks for TF2 chat
            chunks = []
            current_chunk = ""
            words = text.split()
            
            for word in words:
                # Check if adding this word would exceed the limit
                test_chunk = current_chunk + " " + word if current_chunk else word
                if len(test_chunk) <= 110:
                    current_chunk = test_chunk
                else:
                    # Save current chunk and start new one
                    if current_chunk:
                        chunks.append(current_chunk)
                    current_chunk = word
            
            # Add any remaining text
            if current_chunk:
                chunks.append(current_chunk)
            
            # Output each chunk to TF2 chat
            import time
            for i, chunk in enumerate(chunks):
                tf2_command = f'say "{chunk}"'
                with open(self.log_path_var.get(), 'a', encoding='utf-8') as f:
                    f.write(tf2_command + '\n')
                logger.info(f"TF2 chat output ({i+1}/{len(chunks)}): {chunk}")
                # Small delay between chunks to avoid flooding
                if i < len(chunks) - 1:
                    time.sleep(0.1)
        
        # Speak the full text with selected voice
        logger.info(f"Speaking long-form text ({len(text)} chars) with voice: {selected_voice}")
        self.speak(text, voice_name=selected_voice, username="[Long-form Input]")
        
        # Clear the text area after speaking
        # Don't clear longform text - keep it for reuse
        # self.longform_text.delete("1.0", tk.END)
            
    def process_speech_queue(self):
        """Process messages from the queue sequentially"""
        import time
        
        while self.queue_running:
            try:
                # Wait for a message (timeout to check running status)
                msg = self.message_queue.get(timeout=0.5)
                
                # Check if this user is already blocked before speaking
                blocked_users = [self.blocked_listbox.get(i) for i in range(self.blocked_listbox.size())]
                if msg.get('username') in blocked_users:
                    logger.info(f"Skipping message from blocked user: {msg.get('username')}")
                    continue
                
                # Track who is currently speaking
                self.currently_speaking_user = msg.get('username')
                
                if msg['voice']:
                    # Apply specified voice
                    logger.info(f"Applying voice from message: {msg['voice']}")
                    self.apply_voice(msg['voice'])
                else:
                    # Check if user has a voice preference
                    username = msg.get('username')
                    if username and username in self.user_voice_preferences:
                        # Apply user's preferred voice
                        user_voice = self.user_voice_preferences[username]
                        logger.info(f"Applying {username}'s preferred voice: {user_voice}")
                        self.apply_voice(user_voice)
                    elif username and self.random_voice_enabled:
                        # New user and random voice is enabled - assign them a random voice
                        random_voice = self.get_random_voice_for_user(username)
                        if random_voice:
                            logger.info(f"Assigned random voice to new user {username}: {random_voice}")
                            self.apply_voice(random_voice)
                        else:
                            # Fallback to default if random assignment failed
                            self.apply_default_voice()
                    else:
                        # Apply default voice
                        self.apply_default_voice()
                
                # Speak using async speech so it can be interrupted
                if self.audio_manager:
                    text_to_speak = msg['text']
                    current_voice = msg.get('voice')
                    
                    # If no voice specified in message, check if user has a DECtalk preference
                    if not current_voice:
                        username = msg.get('username')
                        if username and username in self.user_voice_preferences:
                            user_voice = self.user_voice_preferences[username]
                            if user_voice.startswith("[DECtalk] "):
                                current_voice = user_voice
                                logger.info(f"Using {username}'s DECtalk preference: {current_voice}")
                    
                    # Check if using DECtalk
                    used_special_voice = False
                    if current_voice and current_voice.startswith("[DECtalk] "):
                        # Use native DECtalk
                        profile_name = current_voice.replace("[DECtalk] ", "")
                        logger.info(f"Using native DECtalk - Profile: {profile_name}")
                        
                        # Debug logging to understand why DECtalk isn't working
                        logger.info(f"DECtalk manager exists: {self.dectalk_manager is not None}")
                        logger.info(f"DECtalk native exists: {self.dectalk_native is not None}")
                        if self.dectalk_native:
                            logger.info(f"DECtalk native available: {self.dectalk_native.is_available()}")
                        
                        # Always use dectalk_native directly for reliability
                        if self.dectalk_native and self.dectalk_native.is_available():
                            logger.info("Using dectalk_native directly")
                            # Get current audio device from config
                            current_device = self.config.get('audio_device', 'Default')
                            logger.info(f"Current audio device from config: {current_device}")
                            
                            # Set audio device for DECtalk if it contains VoiceMeeter
                            if 'voicemeeter' in current_device.lower():
                                self.dectalk_native.set_audio_device('VoiceMeeter Input')
                                logger.info("Set DECtalk audio to VoiceMeeter Input")
                            
                            # Speak with native DECtalk directly with volume control
                            logger.info(f"Text being sent to DECtalk: '{text_to_speak}'")
                            dectalk_volume = self.dectalk_volume_var.get() if hasattr(self, 'dectalk_volume_var') else 0.7
                            success = self.dectalk_native.speak(text_to_speak, profile_name, use_wav=True, device_override=current_device, volume=dectalk_volume)
                            if success:
                                used_special_voice = True
                                logger.info("DECtalk speech completed successfully")
                            else:
                                # DECtalk didn't complete (stopped or failed) - don't fall back to SAPI5
                                used_special_voice = True
                                logger.info("DECtalk speech did not complete (stopped or failed)")
                        elif self.dectalk_manager and self.dectalk_manager.use_dectalk:
                            # Get current audio device from config
                            current_device = self.config.get('audio_device', 'Default')
                            logger.info(f"Current audio device from config (manager path): {current_device}")
                            
                            # Speak with DECtalk (this is synchronous)
                            success = self.dectalk_manager.speak(text_to_speak, current_voice, current_device)
                            if success:
                                used_special_voice = True
                                logger.info("DECtalk speech completed")
                            else:
                                # DECtalk didn't complete - don't fall back to SAPI5
                                used_special_voice = True
                                logger.info("DECtalk speech did not complete (stopped or failed)")
                        else:
                            logger.warning("DECtalk not available, falling back to SAPI5")
                            self.audio_manager.speak(text_to_speak, use_sync=False)
                    else:
                        # Normal SAPI5 speech
                        logger.info(f"Using SAPI5 speech - Voice: {current_voice}")
                        self.audio_manager.speak(text_to_speak, use_sync=False)
                    
                    # Wait for speech to complete or be interrupted (only for SAPI5)
                    if not used_special_voice:
                        max_wait = 60  # Maximum 60 seconds per message
                        wait_time = 0
                        
                        while wait_time < max_wait:
                            # Check if still speaking
                            if not self.audio_manager.is_speaking():
                                # Speech completed naturally
                                break
                                
                            # Check if user was blocked (every 50ms for quick response)
                            time.sleep(0.05)
                            wait_time += 0.05
                            
                            # Check if user is now blocked
                            blocked_users = [self.blocked_listbox.get(i) for i in range(self.blocked_listbox.size())]
                            if msg.get('username') in blocked_users:
                                logger.info(f"User {msg.get('username')} was blocked during speech")
                                self.audio_manager.stop_all_speech()
                                # Clear remaining messages from this user (already in blocked list)
                                self.clear_user_from_queue(msg.get('username'), add_to_blocked=False)
                                break
                                
                            # Check if speech was interrupted by clearing currently_speaking_user
                            if self.currently_speaking_user != msg.get('username'):
                                logger.info(f"Speech interrupted for user {msg.get('username')}")
                                self.audio_manager.stop_all_speech()
                                break
                
                # Clear current speaker when done
                if self.currently_speaking_user == msg.get('username'):
                    self.currently_speaking_user = None
                    
            except queue.Empty:
                continue
            except Exception as e:
                logger.error(f"Error processing speech queue: {e}")
            
    def speak_with_voice(self, text, voice_name, username=None):
        """Queue text with a specific voice"""
        if not self.audio_manager:
            return
            
        # Queue message with voice and username
        self.speak(text, voice_name, username)
        
    def clear_user_from_queue(self, username, add_to_blocked=True):
        """Remove all messages from a specific user from the queue"""
        if not username:
            return
            
        # Add to blocked list immediately if requested
        if add_to_blocked:
            blocked_users = [self.blocked_listbox.get(i) for i in range(self.blocked_listbox.size())]
            if username not in blocked_users:
                self.blocked_listbox.insert(tk.END, username)
                # Don't save yet - let the caller do that after announcement
                logger.info(f"Added {username} to blocked list")
            
        # Stop current speech if it's from this user
        if self.currently_speaking_user == username:
            if self.audio_manager:
                # Try stopping multiple times for stubborn voices
                for _ in range(3):
                    self.audio_manager.stop_all_speech()
                    import time
                    time.sleep(0.05)  # Small delay between attempts
                # Clear the flag to signal the queue processor to stop waiting
                self.currently_speaking_user = None
                logger.info(f"Stopped current speech from blocked user: {username}")
        
        # Remove all queued messages from this user
        temp_queue = []
        removed_count = 0
        
        # Empty the queue and filter out the user's messages
        while not self.message_queue.empty():
            try:
                msg = self.message_queue.get_nowait()
                if msg.get('username') != username:
                    temp_queue.append(msg)
                else:
                    removed_count += 1
            except:
                break
        
        # Put back the filtered messages
        for msg in temp_queue:
            self.message_queue.put(msg)
            
        if removed_count > 0:
            logger.info(f"Removed {removed_count} queued messages from {username}")
            
    def get_announcement_text(self, announcement_type, text="", username=None):
        """Get announcement text with username if provided"""
        # Check if this specific announcement is enabled
        # Handle key mapping for AUTOBLOCK
        toggle_key = announcement_type
        if announcement_type == "AUTOBLOCK":
            toggle_key = "AUTOBLOCK TRIGGERED"  # UI shows "AUTOBLOCK TRIGGERED" but code uses "AUTOBLOCK"
        
        if not self.config.get('announcement_enabled', {}).get(toggle_key, True):
            logger.info(f"Announcement '{announcement_type}' is disabled, skipping")
            return None
            
        # Get announcement template - ensure it's a dict
        announcements = self.config.get('announcements', {})
        if not isinstance(announcements, dict):
            announcements = {}
            self.config['announcements'] = {}
        
        # Use provided text or get from config
        if not text:
            text = announcements.get(announcement_type, "")
        
        if not text:
            return None
            
        # Include username if provided
        if username:
            # Format: "Username blocked, you can't say that"
            text = f"{username} {text}"
            
        return text
    
    def speak_announcement(self, announcement_type, text="", username=None):
        """Speak announcement with automated voice, optionally including username"""
        if not self.audio_manager:
            return
            
        # Get the announcement text
        announcement_text = self.get_announcement_text(announcement_type, text, username)
        if not announcement_text:
            return
            
        # Use a specific voice for announcements (Microsoft Sam or Zira)
        announcement_voice = self.config.get('announcement_voice', 'Microsoft Sam')
        
        # Use speak_with_voice to handle voice switching automatically
        self.speak_with_voice(announcement_text, announcement_voice)
            
    def reset_to_default_voice(self):
        """Reset to the default voice"""
        default_voice = self.config.get('default_voice', '')
        if default_voice and self.audio_manager:
            for voice in self.voices:
                if voice.name == default_voice:
                    if self.current_voice_id != voice.id:
                        self.audio_manager.set_voice(voice.id)
                        self.current_voice_id = voice.id
                    break
                
    def speak_test(self):
        """Speak the test text using current default voice"""
        text = self.test_entry.get()
        if text and self.audio_manager:
            # Use default voice from combo box (most current selection)
            default_voice = self.default_voice_combo.get()
            if default_voice:
                for voice in self.voices:
                    if voice.name == default_voice:
                        if self.current_voice_id != voice.id:
                            success = self.audio_manager.set_voice(voice.id)
                            if success:
                                self.current_voice_id = voice.id
                            logger.info(f"Speak test - Set voice to: {default_voice} (ID: {voice.id}) - Success: {success}")
                        else:
                            logger.debug(f"Speak test - Already using voice: {default_voice}")
                        break
            self.speak(text)
            
    
    def update_ui_after_autoblock(self, username):
        """Update UI after auto-blocking a user (called on main thread)"""
        # Show in chat
        self.chat_display.insert(tk.END, f"[SYSTEM] Auto-blocked user: {username}\n")
        self.chat_display.see(tk.END)
        logger.info(f"UI updated for auto-blocked user: {username}")
            
    def force_stop_tts(self):
        """Force stop all TTS"""
        if self.audio_manager:
            self.audio_manager.stop()
            
    def stop_all_speech(self):
        """Stop all current speech immediately"""
        if self.audio_manager:
            try:
                self.audio_manager.stop_all_speech()
                
                # Also try to stop DECtalk if it's playing
                if hasattr(self, 'dectalk_native') and self.dectalk_native:
                    self.dectalk_native.stop_speech()
                
                
                # Clear the message queue
                while not self.message_queue.empty():
                    try:
                        self.message_queue.get_nowait()
                    except:
                        break
                        
                logger.info("All speech stopped and queue cleared by admin command")
            except Exception as e:
                logger.error(f"Failed to stop all speech: {e}")
                
    def start_tts(self):
        """Start TTS monitoring for selected game"""
        log_path = self.log_path_var.get()
        if not log_path:
            messagebox.showerror("Error", "Please set log file path")
            return
        
        game_key = 'tf2' if self.current_game == 'Team Fortress 2' else 'drg'
        
        # Stop existing monitor for this game if any
        if game_key in self.monitors:
            self.monitors[game_key].stop()
        
        # Create appropriate monitor
        if game_key == 'tf2':
            self.monitors[game_key] = TF2LogMonitor(log_path)
        elif game_key == 'drg' and DRG_MONITOR_AVAILABLE:
            self.monitors[game_key] = DRGLogMonitor(log_path)
        else:
            messagebox.showerror("Error", f"Monitor not available for {self.current_game}")
            return
        
        # Add callback and start
        self.monitors[game_key].add_callback(self.on_chat_message)
        self.monitors[game_key].start()
        
        # Start speech queue processing thread
        self.queue_running = True
        self.queue_thread = threading.Thread(target=self.process_speech_queue, daemon=True)
        self.queue_thread.start()
        
        self.is_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        self.chat_display.insert(tk.END, "TTS started with WASAPI audio\n")
        self.chat_display.see(tk.END)
        
    def stop_tts(self):
        """Stop TTS for all games"""
        # Stop all monitors
        for monitor in self.monitors.values():
            monitor.stop()
        self.monitors.clear()
            
        # Stop queue processing
        self.queue_running = False
        if self.queue_thread:
            self.queue_thread.join(timeout=1.0)
            
        # Clear any remaining messages in queue
        while not self.message_queue.empty():
            try:
                self.message_queue.get_nowait()
            except:
                break
                
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        self.chat_display.insert(tk.END, "TTS stopped\n")
        self.chat_display.see(tk.END)
        
    def reload_tts(self):
        """Reload TTS"""
        was_running = self.is_running
        if was_running:
            self.stop_tts()
            
        # Reinitialize audio
        self.init_audio()
        
        # Reload config
        self.config = self.load_config()
        self.load_settings()
        
        if was_running:
            self.start_tts()
            
    def add_admin(self):
        """Add admin"""
        from tkinter import simpledialog
        name = simpledialog.askstring("Add Admin", "Enter username:")
        if name:
            self.admin_listbox.insert(tk.END, name)
            
    def remove_admin(self):
        """Remove selected admin"""
        sel = self.admin_listbox.curselection()
        if sel:
            self.admin_listbox.delete(sel[0])
            
    def save_admins(self):
        """Save admin list for current game"""
        admins = [self.admin_listbox.get(i) for i in range(self.admin_listbox.size())]
        game_key = 'tf2' if self.current_game == 'Team Fortress 2' else 'drg'
        
        # Save to game-specific config
        if 'games' not in self.config:
            self.config['games'] = {}
        if game_key not in self.config['games']:
            self.config['games'][game_key] = {}
        self.config['games'][game_key]['admins'] = admins
        
        # Also save to legacy location for backwards compatibility
        if game_key == 'tf2':
            self.config['admins'] = admins
        
        self.save_config()
        
    def add_blocked(self):
        """Add blocked user"""
        from tkinter import simpledialog
        name = simpledialog.askstring("Block User", "Enter username:")
        if name:
            self.blocked_listbox.insert(tk.END, name)
            
    def remove_blocked(self):
        """Remove selected blocked"""
        sel = self.blocked_listbox.curselection()
        if sel:
            self.blocked_listbox.delete(sel[0])
            
    def save_blocked(self):
        """Save blocked list for current game"""
        blocked = [self.blocked_listbox.get(i) for i in range(self.blocked_listbox.size())]
        game_key = 'tf2' if self.current_game == 'Team Fortress 2' else 'drg'
        
        # Save to game-specific config
        if 'games' not in self.config:
            self.config['games'] = {}
        if game_key not in self.config['games']:
            self.config['games'][game_key] = {}
        self.config['games'][game_key]['blocked'] = blocked
        
        # Also save to legacy location for backwards compatibility
        if game_key == 'tf2':
            self.config['blocked'] = blocked
        
        self.save_config()
        
    def add_user_voice(self):
        """Add user voice preference"""
        from tkinter import simpledialog
        username = simpledialog.askstring("User Voice", "Enter username:")
        if username:
            # Show voice selection dialog
            voice_names = [v.name for v in self.voices if v.name]
            
            # Add DECtalk voices if enabled
            if self.dectalk_enabled and self.dectalk_profiles:
                for profile_name in self.dectalk_profiles.keys():
                    voice_names.append(f"[DECtalk] {profile_name}")
            
            from tkinter import messagebox
            # Create custom dialog for voice selection
            dialog = tk.Toplevel(self.root)
            dialog.title("Select Voice")
            dialog.geometry("400x300")
            
            ttk.Label(dialog, text=f"Select voice for {username}:").pack(pady=10)
            
            voice_listbox = tk.Listbox(dialog, height=10)
            for voice in voice_names:
                voice_listbox.insert(tk.END, voice)
            voice_listbox.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
            
            def confirm_selection():
                sel = voice_listbox.curselection()
                if sel:
                    selected_voice = voice_names[sel[0]]
                    # Add to listbox in format "username: voice"
                    self.user_voices_listbox.insert(tk.END, f"{username}: {selected_voice}")
                    self.user_voice_preferences[username] = selected_voice
                dialog.destroy()
            
            ttk.Button(dialog, text="Select", command=confirm_selection).pack(pady=10)
            
    def remove_user_voice(self):
        """Remove selected user voice preference"""
        sel = self.user_voices_listbox.curselection()
        if sel:
            entry = self.user_voices_listbox.get(sel[0])
            username = entry.split(':')[0].strip()
            self.user_voices_listbox.delete(sel[0])
            if username in self.user_voice_preferences:
                del self.user_voice_preferences[username]
                
    def edit_user_voice(self):
        """Edit selected user voice preference"""
        sel = self.user_voices_listbox.curselection()
        if sel:
            entry = self.user_voices_listbox.get(sel[0])
            username = entry.split(':')[0].strip()
            
            # Show voice selection dialog
            voice_names = [v.name for v in self.voices if v.name]
            
            # Add DECtalk voices if enabled
            if self.dectalk_enabled and self.dectalk_profiles:
                for profile_name in self.dectalk_profiles.keys():
                    voice_names.append(f"[DECtalk] {profile_name}")
            
            dialog = tk.Toplevel(self.root)
            dialog.title("Edit Voice")
            dialog.geometry("400x300")
            
            ttk.Label(dialog, text=f"Select new voice for {username}:").pack(pady=10)
            
            voice_listbox = tk.Listbox(dialog, height=10)
            for voice in voice_names:
                voice_listbox.insert(tk.END, voice)
            voice_listbox.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
            
            # Select current voice if exists
            current_voice = self.user_voice_preferences.get(username)
            if current_voice in voice_names:
                idx = voice_names.index(current_voice)
                voice_listbox.selection_set(idx)
            
            def confirm_selection():
                sel_voice = voice_listbox.curselection()
                if sel_voice:
                    selected_voice = voice_names[sel_voice[0]]
                    # Update listbox
                    self.user_voices_listbox.delete(sel[0])
                    self.user_voices_listbox.insert(sel[0], f"{username}: {selected_voice}")
                    self.user_voice_preferences[username] = selected_voice
                dialog.destroy()
            
            ttk.Button(dialog, text="Update", command=confirm_selection).pack(pady=10)
            
    def save_user_voices(self):
        """Save user voice preferences"""
        from tkinter import messagebox
        # Parse listbox and update preferences
        self.user_voice_preferences = {}
        for i in range(self.user_voices_listbox.size()):
            entry = self.user_voices_listbox.get(i)
            if ':' in entry:
                username, voice = entry.split(':', 1)
                self.user_voice_preferences[username.strip()] = voice.strip()
        
        self.config['user_voice_preferences'] = self.user_voice_preferences
        self.save_config()
        messagebox.showinfo("Saved", "User voice preferences saved")
        
    def save_auto_block(self):
        """Save auto-block keywords"""
        text = self.auto_block_text.get(1.0, tk.END)
        keywords = [line.strip() for line in text.split('\n') if line.strip()]
        self.config['auto_block_keywords'] = keywords
        self.save_config()
        messagebox.showinfo("Saved", "Auto-block keywords saved")
    
    
    def update_dectalk_volume_label(self, *args):
        """Update DECtalk volume label and save to config"""
        volume = self.dectalk_volume_var.get()
        self.dectalk_volume_label.config(text=f"{int(volume * 100)}%")
        # Save to config
        self.config['dectalk_volume'] = volume
        self.save_config()
    
    def refresh_voice_combos(self):
        """Refresh voice combo boxes to exclude hidden voices and add special voices"""
        voice_names = []
        
        # Add Windows SAPI voices
        if self.voices:
            # Filter out hidden voices
            voice_names.extend([v.name for v in self.voices if v.name not in self.hidden_voices])
        
        # Add DECtalk voices if available
        if hasattr(self, 'dectalk_native') and self.dectalk_native.is_available():
            dectalk_profiles = self.dectalk_native.get_available_profiles()
            for profile in dectalk_profiles:
                voice_names.append(f"[DECtalk] {profile}")
        
        
        # Update default voice combo
        if hasattr(self, 'default_voice_combo'):
            # Get current selection
            current = self.default_voice_combo.get()
            
            # Update combo values
            self.default_voice_combo['values'] = voice_names
            
            # Restore selection if still available
            if current in voice_names:
                self.default_voice_combo.set(current)
            elif voice_names:
                self.default_voice_combo.set(voice_names[0])
    
    # DECtalk Methods
    def toggle_dectalk(self):
        """Toggle DECtalk extended voices on/off"""
        # Check if native DECtalk is available
        if not self.dectalk_native.is_available():
            messagebox.showwarning("DECtalk Not Available", 
                                 "Native DECtalk (say.exe) not found.\n"
                                 "Please place DECtalk files in the dectalk_bin folder:\n"
                                 "- say.exe\n"
                                 "- dectalk.dll\n"
                                 "- dtalk_us.dic")
            self.dectalk_enabled_var.set(False)
            return
            
        self.dectalk_enabled = self.dectalk_enabled_var.get()
        self.config['dectalk_enabled'] = self.dectalk_enabled
        self.save_config()
        
        if self.dectalk_enabled:
            messagebox.showinfo("DECtalk Native", "Native DECtalk voices enabled.\nProfiles will appear in Voice Commands.")
        else:
            messagebox.showinfo("DECtalk Native", "Native DECtalk voices disabled.")
    
    def refresh_dectalk_profiles(self):
        """Refresh the DECtalk profiles listbox"""
        self.dectalk_profiles_listbox.delete(0, tk.END)
        for name in sorted(self.dectalk_profiles.keys()):
            self.dectalk_profiles_listbox.insert(tk.END, name)
    
    def on_dectalk_profile_select(self, event):
        """Handle DECtalk profile selection"""
        selection = self.dectalk_profiles_listbox.curselection()
        if selection:
            index = selection[0]
            name = self.dectalk_profiles_listbox.get(index)
            code = self.dectalk_profiles.get(name, "")
            self.dectalk_profile_name_var.set(name)
            self.dectalk_profile_code_var.set(code)
    
    def add_dectalk_profile(self):
        """Add a new DECtalk profile"""
        name = self.dectalk_profile_name_var.get().strip()
        code = self.dectalk_profile_code_var.get().strip()
        
        if not name or not code:
            messagebox.showwarning("Invalid Input", "Please enter both profile name and DECtalk code")
            return
        
        if name in self.dectalk_profiles:
            if not messagebox.askyesno("Profile Exists", f"Profile '{name}' already exists. Overwrite?"):
                return
        
        self.dectalk_profiles[name] = code
        self.refresh_dectalk_profiles()
        self.refresh_voice_commands_display()
        messagebox.showinfo("Success", f"Profile '{name}' added successfully")
    
    def update_dectalk_profile(self):
        """Update selected DECtalk profile"""
        selection = self.dectalk_profiles_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a profile to update")
            return
        
        old_name = self.dectalk_profiles_listbox.get(selection[0])
        new_name = self.dectalk_profile_name_var.get().strip()
        code = self.dectalk_profile_code_var.get().strip()
        
        if not new_name or not code:
            messagebox.showwarning("Invalid Input", "Please enter both profile name and DECtalk code")
            return
        
        # Remove old entry if name changed
        if old_name != new_name:
            del self.dectalk_profiles[old_name]
        
        self.dectalk_profiles[new_name] = code
        self.refresh_dectalk_profiles()
        self.refresh_voice_commands_display()
        messagebox.showinfo("Success", f"Profile updated successfully")
    
    def remove_dectalk_profile(self):
        """Remove selected DECtalk profile"""
        selection = self.dectalk_profiles_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a profile to remove")
            return
        
        name = self.dectalk_profiles_listbox.get(selection[0])
        if messagebox.askyesno("Confirm Delete", f"Remove profile '{name}'?"):
            del self.dectalk_profiles[name]
            self.refresh_dectalk_profiles()
            self.refresh_voice_commands_display()
            self.dectalk_profile_name_var.set("")
            self.dectalk_profile_code_var.set("")
    
    def test_dectalk_profile(self):
        """Test selected DECtalk profile using native DECtalk"""
        selection = self.dectalk_profiles_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a profile to test")
            return
        
        name = self.dectalk_profiles_listbox.get(selection[0])
        
        # Test using native DECtalk
        test_text = f"Testing {name} voice profile."
        
        if self.dectalk_native and self.dectalk_native.is_available():
            # Test with native DECtalk
            success = self.dectalk_native.speak(test_text, name)
            if success:
                logger.info(f"Tested DECtalk profile: {name} with native DECtalk")
            else:
                messagebox.showwarning("Test Failed", f"Failed to test {name} profile")
        else:
            messagebox.showwarning("DECtalk Not Available", 
                                 "Native DECtalk is not available.\n"
                                 "Please ensure say.exe is in the dectalk_bin folder.")
    
    def save_dectalk_profiles(self):
        """Save DECtalk profiles to config"""
        self.config['dectalk_profiles'] = self.dectalk_profiles
        self.save_config()
        messagebox.showinfo("Success", "DECtalk profiles saved")
    
    def reset_dectalk_profiles(self):
        """Reset DECtalk profiles to defaults"""
        if messagebox.askyesno("Confirm Reset", "Reset all DECtalk profiles to defaults?"):
            self.dectalk_profiles = self.get_default_dectalk_profiles()
            self.config['dectalk_profiles'] = self.dectalk_profiles
            self.save_config()
            self.refresh_dectalk_profiles()
            self.refresh_voice_commands_display()
            messagebox.showinfo("Success", "DECtalk profiles reset to defaults")
    
    def export_dectalk_profiles(self):
        """Export DECtalk profiles to file"""
        from tkinter import filedialog
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(self.dectalk_profiles, f, indent=2)
                messagebox.showinfo("Success", f"Profiles exported to {filename}")
            except Exception as e:
                messagebox.showerror("Export Failed", str(e))
    
    def import_dectalk_profiles(self):
        """Import DECtalk profiles from file"""
        from tkinter import filedialog
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r') as f:
                    imported = json.load(f)
                if not isinstance(imported, dict):
                    raise ValueError("Invalid profile format")
                self.dectalk_profiles.update(imported)
                self.config['dectalk_profiles'] = self.dectalk_profiles
                self.save_config()
                self.refresh_dectalk_profiles()
                self.refresh_voice_commands_display()
                messagebox.showinfo("Success", f"Profiles imported from {filename}")
            except Exception as e:
                messagebox.showerror("Import Failed", str(e))
    
    def refresh_available_voices(self):
        """Refresh the available voices list"""
        if not hasattr(self, 'available_voices_listbox'):
            return
            
        self.available_voices_listbox.delete(0, tk.END)
        
        # Track seen voices for duplicate detection
        seen_voices = {}
        
        # Add SAPI5 voices
        for i, voice in enumerate(self.voices):
            # Check if voice is hidden
            if voice.name not in self.hidden_voices:
                # Check for duplicates
                display_name = voice.name
                if voice.name in seen_voices:
                    display_name = f"{voice.name} (duplicate #{seen_voices[voice.name] + 1})"
                    seen_voices[voice.name] += 1
                else:
                    seen_voices[voice.name] = 1
                    
                self.available_voices_listbox.insert(tk.END, f"[{i}] {display_name}")
        
        # Add DECtalk voices if enabled
        if self.dectalk_enabled and self.dectalk_profiles:
            for profile_name in self.dectalk_profiles.keys():
                self.available_voices_listbox.insert(tk.END, f"[DECtalk] {profile_name}")
        
            
    def hide_selected_voices(self):
        """Hide selected voices from the list"""
        selected = self.available_voices_listbox.curselection()
        if selected:
            for idx in selected:
                voice_text = self.available_voices_listbox.get(idx)
                # Extract voice name (remove index prefix)
                voice_name = voice_text.split('] ', 1)[1] if '] ' in voice_text else voice_text
                # Remove duplicate marker if present
                if ' (duplicate #' in voice_name:
                    voice_name = voice_name.split(' (duplicate #')[0]
                    
                if voice_name not in self.hidden_voices:
                    self.hidden_voices.append(voice_name)
            
            # Save to config
            self.config['hidden_voices'] = self.hidden_voices
            self.save_config()
            
            # Refresh lists
            self.refresh_available_voices()
            self.refresh_voice_combos()
            messagebox.showinfo("Hidden", f"Hidden {len(selected)} voice(s)")
            
    def show_all_voices(self):
        """Show all voices"""
        self.hidden_voices = []
        self.config['hidden_voices'] = []
        self.save_config()
        self.refresh_available_voices()
        self.refresh_voice_combos()
        messagebox.showinfo("Shown", "All voices are now visible")
        
    def check_duplicate_voices(self):
        """Check for duplicate voices and show report"""
        duplicates = {}
        
        # Find duplicates
        for i, voice in enumerate(self.voices):
            if voice.name in duplicates:
                duplicates[voice.name].append(i)
            else:
                duplicates[voice.name] = [i]
        
        # Filter to only actual duplicates
        actual_duplicates = {name: indices for name, indices in duplicates.items() if len(indices) > 1}
        
        if actual_duplicates:
            report = "Duplicate voices found:\n\n"
            for name, indices in actual_duplicates.items():
                report += f"{name}:\n"
                report += f"  Indices: {', '.join(str(i) for i in indices)}\n"
                report += f"  Commands: {', '.join(f'v{i}' for i in indices)}\n\n"
            
            # Show in dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Duplicate Voices Report")
            dialog.geometry("500x400")
            
            text = scrolledtext.ScrolledText(dialog, wrap=tk.WORD)
            text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            text.insert(1.0, report)
            text.config(state=tk.DISABLED)
            
            ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=5)
        else:
            messagebox.showinfo("No Duplicates", "No duplicate voices found")
    
    def validate_command_trigger(self, trigger):
        """Validate command trigger format"""
        import re
        # Allow alphanumeric, underscores, and spaces for legacy "v 0" format
        if not trigger:
            return False, "Command trigger cannot be empty"
        # Special case for legacy "v [number]" format
        if re.match(r'^v \d+$', trigger):
            return True, ""
        # For other commands, allow alphanumeric and underscores (no spaces)
        if not re.match(r'^[a-zA-Z0-9_]+$', trigger):
            return False, "Command trigger must be alphanumeric (or use 'v [number]' format)"
        if len(trigger) > 20:
            return False, "Command trigger must be 20 characters or less"
        return True, ""
    
    def add_voice_command(self):
        """Add a new voice command"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Voice Command")
        dialog.geometry("450x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Command trigger input
        ttk.Label(dialog, text="Command Trigger (e.g., 'c', 'sam', 'v10'):").pack(pady=10)
        trigger_var = tk.StringVar()
        trigger_entry = ttk.Entry(dialog, textvariable=trigger_var, width=30)
        trigger_entry.pack(pady=5)
        
        # Preview label
        preview_label = ttk.Label(dialog, text="Usage: !tts /[trigger] message", foreground='gray')
        preview_label.pack(pady=5)
        
        def update_preview(*args):
            trigger = trigger_var.get()
            if trigger:
                preview_label.config(text=f"Usage: !tts /{trigger} message")
            else:
                preview_label.config(text="Usage: !tts /[trigger] message")
        
        trigger_var.trace('w', update_preview)
        
        # Voice selection
        ttk.Label(dialog, text="Select Voice:").pack(pady=10)
        
        # Voice listbox with scrollbar
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        voice_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        voice_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=voice_listbox.yview)
        
        # Populate with available voices (excluding hidden ones)
        voice_names = []
        for voice in self.voices:
            if voice.name not in self.hidden_voices:
                voice_names.append(voice.name)
                voice_listbox.insert(tk.END, voice.name)
        
        # Add DECtalk profiles if enabled
        if self.dectalk_enabled and self.dectalk_profiles:
            voice_listbox.insert(tk.END, "")  # Separator
            voice_listbox.insert(tk.END, "--- DECtalk Profiles ---")
            for profile_name in sorted(self.dectalk_profiles.keys()):
                dectalk_label = f"[DECtalk] {profile_name}"
                voice_names.append(dectalk_label)
                voice_listbox.insert(tk.END, dectalk_label)
        
        
        # Select first voice by default
        if voice_names:
            voice_listbox.selection_set(0)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def add_command():
            trigger = trigger_var.get().strip()
            
            # Validate trigger
            valid, error_msg = self.validate_command_trigger(trigger)
            if not valid:
                messagebox.showerror("Invalid Trigger", error_msg, parent=dialog)
                return
            
            # Check for duplicates
            if trigger in self.voice_commands or f"v {trigger}" in self.voice_commands:
                messagebox.showerror("Duplicate Command", 
                                   f"Command '{trigger}' already exists!", parent=dialog)
                return
            
            # Get selected voice
            sel = voice_listbox.curselection()
            if not sel:
                messagebox.showerror("No Voice Selected", 
                                   "Please select a voice for the command", parent=dialog)
                return
            
            # Get the actual text from the listbox to handle separators
            selected_text = voice_listbox.get(sel[0])
            
            # Skip separator lines
            if selected_text == "" or selected_text == "--- DECtalk Profiles ---":
                messagebox.showerror("Invalid Selection", 
                                   "Please select a valid voice", parent=dialog)
                return
            
            # Use the selected text as the voice name
            selected_voice = selected_text
            
            # Add to commands
            self.voice_commands[trigger] = selected_voice
            
            # Add to tree
            self.voice_tree.insert('', tk.END, values=(trigger, selected_voice))
            
            # Close dialog
            dialog.destroy()
            
            messagebox.showinfo("Command Added", 
                              f"Command '/{trigger}' added successfully!\nDon't forget to save.")
        
        def cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="Add", command=add_command).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=cancel).pack(side=tk.LEFT, padx=5)
        
        # Focus on trigger entry
        trigger_entry.focus()
    
    def remove_voice_command(self):
        """Remove selected voice command"""
        selection = self.voice_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a command to remove")
            return
        
        # Get command details
        item = self.voice_tree.item(selection[0])
        cmd = item['values'][0]
        
        # Confirm deletion
        if messagebox.askyesno("Confirm Removal", 
                               f"Remove command '/{cmd}'?"):
            # Remove from tree
            self.voice_tree.delete(selection[0])
            
            # Remove from dictionary
            if cmd in self.voice_commands:
                del self.voice_commands[cmd]
            
            messagebox.showinfo("Command Removed", 
                              f"Command '/{cmd}' removed.\nDon't forget to save.")
    
    def edit_voice_command_full(self):
        """Enhanced edit allowing both trigger and voice changes"""
        selection = self.voice_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a command to edit")
            return
        
        item = self.voice_tree.item(selection[0])
        old_cmd = item['values'][0]
        old_voice = item['values'][1]
        
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Command: {old_cmd}")
        dialog.geometry("450x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Command trigger input
        ttk.Label(dialog, text="Command Trigger:").pack(pady=10)
        trigger_var = tk.StringVar(value=old_cmd)
        trigger_entry = ttk.Entry(dialog, textvariable=trigger_var, width=30)
        trigger_entry.pack(pady=5)
        
        # Preview label
        preview_label = ttk.Label(dialog, text=f"Usage: !tts /{old_cmd} message", foreground='gray')
        preview_label.pack(pady=5)
        
        def update_preview(*args):
            trigger = trigger_var.get()
            if trigger:
                preview_label.config(text=f"Usage: !tts /{trigger} message")
            else:
                preview_label.config(text="Usage: !tts /[trigger] message")
        
        trigger_var.trace('w', update_preview)
        
        # Voice selection
        ttk.Label(dialog, text="Select Voice:").pack(pady=10)
        
        # Voice listbox with scrollbar
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        voice_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        voice_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=voice_listbox.yview)
        
        # Populate with available voices
        voice_names = []
        for voice in self.voices:
            if voice.name not in self.hidden_voices:
                voice_names.append(voice.name)
                voice_listbox.insert(tk.END, voice.name)
        
        # Add DECtalk profiles if enabled
        if self.dectalk_enabled and self.dectalk_profiles:
            voice_listbox.insert(tk.END, "")  # Separator
            voice_listbox.insert(tk.END, "--- DECtalk Profiles ---")
            for profile_name in sorted(self.dectalk_profiles.keys()):
                dectalk_label = f"[DECtalk] {profile_name}"
                voice_names.append(dectalk_label)
                voice_listbox.insert(tk.END, dectalk_label)
        
        
        # Select current voice
        try:
            idx = voice_names.index(old_voice)
            voice_listbox.selection_set(idx)
            voice_listbox.see(idx)
        except ValueError:
            if voice_names:
                voice_listbox.selection_set(0)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def save_changes():
            new_trigger = trigger_var.get().strip()
            
            # Validate trigger
            valid, error_msg = self.validate_command_trigger(new_trigger)
            if not valid:
                messagebox.showerror("Invalid Trigger", error_msg, parent=dialog)
                return
            
            # Check for duplicates (if trigger changed)
            if new_trigger != old_cmd:
                if new_trigger in self.voice_commands:
                    messagebox.showerror("Duplicate Command", 
                                       f"Command '{new_trigger}' already exists!", parent=dialog)
                    return
            
            # Get selected voice
            sel = voice_listbox.curselection()
            if not sel:
                messagebox.showerror("No Voice Selected", 
                                   "Please select a voice for the command", parent=dialog)
                return
            
            # Get the actual text from the listbox to handle separators
            selected_text = voice_listbox.get(sel[0])
            
            # Skip separator lines
            if selected_text == "" or selected_text == "--- DECtalk Profiles ---":
                messagebox.showerror("Invalid Selection", 
                                   "Please select a valid voice", parent=dialog)
                return
            
            # Use the selected text as the voice name
            new_voice = selected_text
            
            # Update dictionary
            if old_cmd != new_trigger:
                # Remove old command
                if old_cmd in self.voice_commands:
                    del self.voice_commands[old_cmd]
            
            # Add new/updated command
            self.voice_commands[new_trigger] = new_voice
            
            # Update tree
            self.voice_tree.item(selection[0], values=(new_trigger, new_voice))
            
            # Close dialog
            dialog.destroy()
            
            messagebox.showinfo("Command Updated", 
                              f"Command updated successfully!\nDon't forget to save.")
        
        def cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="Save", command=save_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=cancel).pack(side=tk.LEFT, padx=5)
        
        # Focus on trigger entry
        trigger_entry.focus()
        trigger_entry.selection_range(0, tk.END)
    
    def edit_voice_command(self, event):
        """Edit voice command on double-click"""
        selection = self.voice_tree.selection()
        if selection:
            item = self.voice_tree.item(selection[0])
            cmd = item['values'][0]
            current_voice = item['values'][1]
            
            # Create dialog to select new voice
            from tkinter import simpledialog
            dialog = tk.Toplevel(self.root)
            dialog.title(f"Edit Voice for {cmd}")
            dialog.geometry("400x300")
            
            ttk.Label(dialog, text=f"Select voice for command {cmd}:").pack(pady=10)
            
            # Voice listbox
            voice_listbox = tk.Listbox(dialog, height=10)
            voice_listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
            
            # Populate with available voices (excluding hidden ones)
            voice_names = []
            for voice in self.voices:
                if voice.name not in self.hidden_voices:
                    voice_names.append(voice.name)
                    voice_listbox.insert(tk.END, voice.name)
            
            # Add DECtalk profiles if enabled
            if self.dectalk_enabled and self.dectalk_profiles:
                voice_listbox.insert(tk.END, "")  # Separator
                voice_listbox.insert(tk.END, "--- DECtalk Profiles ---")
                for profile_name in sorted(self.dectalk_profiles.keys()):
                    dectalk_label = f"[DECtalk] {profile_name}"
                    voice_names.append(dectalk_label)
                    voice_listbox.insert(tk.END, dectalk_label)
                
            # Select current voice
            try:
                idx = voice_names.index(current_voice)
                voice_listbox.selection_set(idx)
            except ValueError:
                pass
                
            def apply_voice():
                sel = voice_listbox.curselection()
                if sel:
                    new_voice = voice_names[sel[0]]
                    # Update the tree
                    self.voice_tree.item(selection[0], values=(cmd, new_voice))
                    # Update the mapping
                    self.voice_commands[cmd] = new_voice
                    dialog.destroy()
                    
            ttk.Button(dialog, text="Apply", command=apply_voice).pack(pady=10)
            
    def save_voice_commands(self):
        """Save voice command mappings"""
        # Update voice_commands from tree
        for item in self.voice_tree.get_children():
            values = self.voice_tree.item(item)['values']
            cmd = values[0]
            voice = values[1]
            self.voice_commands[cmd] = voice
            
        # Save voice toggle command
        if hasattr(self, 'voice_toggle_entry'):
            self.voice_toggle_command = self.voice_toggle_entry.get()
            self.config['voice_toggle_command'] = self.voice_toggle_command
            
        # Save TTS command prefix per game
        if hasattr(self, 'tts_prefix_entry'):
            tts_prefix = self.tts_prefix_entry.get().strip()
            if tts_prefix:
                current_game_config = self.get_current_game_config()
                current_game_config['tts_command_prefix'] = tts_prefix
            
        # Save to config
        self.config['voice_commands'] = self.voice_commands
        self.save_config()
        messagebox.showinfo("Saved", "Voice commands saved")
        
    def reset_voice_commands(self):
        """Reset voice commands to defaults"""
        # Default mappings
        self.voice_commands = {
            'v 1': 'Microsoft Lili - Chinese (China)',
            'v 7': 'Microsoft Irina Desktop - Russian',
            'v 3': 'Microsoft Mary',
            'v 9': 'Microsoft David Desktop - English (United States)',
            'v 0': 'Microsoft Sam',
            'v 4': 'Microsoft Anna - English (United States)',
            'v 2': 'Microsoft Zira Desktop - English (United States)',
        }
        
        # Clear tree and repopulate
        for item in self.voice_tree.get_children():
            self.voice_tree.delete(item)
            
        for cmd, voice in self.voice_commands.items():
            self.voice_tree.insert('', tk.END, values=(cmd, voice), tags=(cmd,))
            
        messagebox.showinfo("Reset", "Voice commands reset to defaults")
    
    
    def create_announcements_tab(self):
        """Announcements tab with automated voice settings"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Announcements")
        
        # Container
        container = ttk.Frame(tab)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title and description
        ttk.Label(container, text="AUTOMATED ANNOUNCEMENTS", font=('Arial', 12, 'bold')).pack(pady=5)
        ttk.Label(container, text="Configure announcement messages and voices").pack(pady=5)
        
        # Announcement voice selector
        voice_frame = ttk.Frame(container)
        voice_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(voice_frame, text="Announcement Voice:").pack(side=tk.LEFT, padx=5)
        self.announcement_voice_combo = ttk.Combobox(voice_frame, width=40)
        self.announcement_voice_combo.pack(side=tk.LEFT, padx=5)
        
        # Populate with voices
        if self.voices:
            voice_names = [v.name for v in self.voices]
            self.announcement_voice_combo['values'] = voice_names
            # Set to Microsoft Sam or first voice
            for v in voice_names:
                if 'Sam' in v:
                    self.announcement_voice_combo.set(v)
                    break
            else:
                self.announcement_voice_combo.set(voice_names[0])
                
        # Announcements grid
        announcements_frame = ttk.LabelFrame(container, text="Announcement Messages")
        announcements_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Announcement types and defaults - only include actually used announcements
        announcements = [
            ("ADMIN ADD", "is now an admin"),
            ("BLOCK ADD", "blocked, you abused it so you losed it"),
            ("BLOCK REMOVE", "has been unblocked"),
            ("AUTOBLOCK TRIGGERED", "blocked, you can't say that"),
            ("TTS STOPPED", "TTS stopped by admin")
        ]
        
        self.announcement_vars = {}
        self.announcement_enabled_vars = {}
        
        # Headers
        ttk.Label(announcements_frame, text="On", font=('Arial', 9, 'bold')).grid(row=0, column=0, padx=5, pady=3)
        ttk.Label(announcements_frame, text="Type", font=('Arial', 9, 'bold')).grid(row=0, column=1, sticky=tk.W, padx=10, pady=3)
        ttk.Label(announcements_frame, text="Message", font=('Arial', 9, 'bold')).grid(row=0, column=2, padx=10, pady=3)
        
        for i, (label, default) in enumerate(announcements):
            row = i + 1  # Start from row 1 due to headers
            
            # Enable/disable checkbox
            enabled_var = tk.BooleanVar(value=self.config.get('announcement_enabled', {}).get(label, True))
            self.announcement_enabled_vars[label] = enabled_var
            ttk.Checkbutton(announcements_frame, variable=enabled_var).grid(row=row, column=0, padx=5, pady=2)
            
            # Label
            ttk.Label(announcements_frame, text=label).grid(row=row, column=1, sticky=tk.W, padx=10, pady=2)
            
            # Get saved value or use default (fix AUTOBLOCK name)
            key = label
            if label == "AUTOBLOCK TRIGGERED":
                key = "AUTOBLOCK"  # Use consistent key
            saved_value = self.config.get('announcements', {}).get(key, default)
            var = tk.StringVar(value=saved_value)
            ttk.Entry(announcements_frame, textvariable=var, width=40).grid(row=row, column=2, padx=10, pady=2)
            self.announcement_vars[key] = var
            
            # Test button for each announcement
            ttk.Button(announcements_frame, text="Test", width=6,
                      command=lambda l=key: self.test_announcement(l)).grid(row=row, column=3, padx=5, pady=2)
                      
        # Save button
        ttk.Button(container, text="Save Announcements", command=self.save_announcements).pack(pady=10)
        
    def test_announcement(self, announcement_type):
        """Test a specific announcement"""
        self.speak_announcement(announcement_type)
        
    def save_announcements(self):
        """Save announcement settings"""
        # Save announcement messages
        announcements = {}
        if hasattr(self, 'announcement_vars'):
            for label, var in self.announcement_vars.items():
                announcements[label] = var.get()
        self.config['announcements'] = announcements
        
        # Save enabled states with proper key mapping
        announcement_enabled = {}
        if hasattr(self, 'announcement_enabled_vars'):
            for label, var in self.announcement_enabled_vars.items():
                # Map keys properly for save
                key = label
                if label == "AUTOBLOCK TRIGGERED":
                    key = "AUTOBLOCK TRIGGERED"  # Keep the UI key for consistency
                announcement_enabled[key] = var.get()
        self.config['announcement_enabled'] = announcement_enabled
        
        # Save announcement voice
        if hasattr(self, 'announcement_voice_combo'):
            self.config['announcement_voice'] = self.announcement_voice_combo.get()
        
        self.save_config()
        messagebox.showinfo("Saved", "Announcements saved")
        
    def create_testing_tab(self):
        """Testing tab for simulating messages"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Testing")
        
        # Main container
        main_frame = ttk.Frame(tab)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(main_frame, text="Test Message Simulator", font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(main_frame, text="Simulate chat messages to test blocking, commands, and TTS behavior")
        desc_label.pack(pady=(0, 20))
        
        # Input fields frame
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=10)
        
        # Username field
        ttk.Label(input_frame, text="Username:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.test_username_var = tk.StringVar(value="TestUser")
        self.test_username_entry = ttk.Entry(input_frame, textvariable=self.test_username_var, width=30)
        self.test_username_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Message field
        ttk.Label(input_frame, text="Message:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.test_message_var = tk.StringVar(value="Hello, this is a test message")
        self.test_message_entry = ttk.Entry(input_frame, textvariable=self.test_message_var, width=50)
        self.test_message_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Message Options")
        options_frame.pack(fill=tk.X, pady=10)
        
        # Checkboxes for message flags
        self.test_is_dead_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="*DEAD* flag", variable=self.test_is_dead_var).pack(side=tk.LEFT, padx=10, pady=5)
        
        self.test_is_team_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="(TEAM) flag", variable=self.test_is_team_var).pack(side=tk.LEFT, padx=10, pady=5)
        
        # Test buttons frame
        buttons_frame = ttk.LabelFrame(main_frame, text="Test Actions")
        buttons_frame.pack(fill=tk.X, pady=10)
        
        # Button grid
        button_row1 = ttk.Frame(buttons_frame)
        button_row1.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_row1, text="Send Test Message", command=self.send_test_message, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_row1, text="Test as Admin", command=self.test_as_admin, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_row1, text="Test as Blocked", command=self.test_as_blocked, width=20).pack(side=tk.LEFT, padx=5)
        
        button_row2 = ttk.Frame(buttons_frame)
        button_row2.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_row2, text="Test Voice Command", command=self.test_voice_command, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_row2, text="Test Auto-Block", command=self.test_auto_block, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_row2, text="Test Private Mode", command=self.test_private_mode, width=20).pack(side=tk.LEFT, padx=5)
        
        button_row3 = ttk.Frame(buttons_frame)
        button_row3.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_row3, text="Test Voice Toggle", command=self.test_voice_toggle, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_row3, text="Test User Preference", command=self.test_user_preference, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_row3, text="Clear Voice Preference", command=self.clear_test_voice_preference, width=20).pack(side=tk.LEFT, padx=5)
        
        # Quick test presets
        presets_frame = ttk.LabelFrame(main_frame, text="Quick Test Presets")
        presets_frame.pack(fill=tk.X, pady=10)
        
        preset_buttons = ttk.Frame(presets_frame)
        preset_buttons.pack(pady=10)
        
        ttk.Button(preset_buttons, text="Normal Message", command=lambda: self.set_test_preset("TestUser", "Hello everyone!")).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_buttons, text="Voice Command", command=lambda: self.set_test_preset("TestUser", "/v 2 Testing voice two")).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_buttons, text="Voice Toggle", command=lambda: self.set_test_preset("TestUser", f"{self.voice_toggle_command} 3 Setting my voice")).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_buttons, text="Admin Command", command=self.set_admin_stop_preset).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_buttons, text="Blocked Word", command=lambda: self.set_test_preset("BadUser", "This contains a blocked word")).pack(side=tk.LEFT, padx=5)
        
        # Test output
        output_frame = ttk.LabelFrame(main_frame, text="Test Output")
        output_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.test_output = scrolledtext.ScrolledText(output_frame, height=8, width=60)
        self.test_output.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
    def send_test_message(self):
        """Send a test message through the system"""
        username = self.test_username_var.get()
        message = self.test_message_var.get()
        is_dead = self.test_is_dead_var.get()
        is_team = self.test_is_team_var.get()
        
        if not username or not message:
            messagebox.showwarning("Invalid Input", "Please enter both username and message")
            return
            
        # If message doesn't start with TTS prefix, add it (unless it's a special command)
        current_tts_prefix = self.get_current_game_config().get('tts_command_prefix', '!tts')
        if not message.startswith('!'):
            message = f"{current_tts_prefix} {message}"
            
        # Create test message
        test_msg = {
            'username': username,
            'message': message,
            'is_dead': is_dead,
            'is_team': is_team,
            'raw': f"{username} : {message}"
        }
        
        # Log to test output
        self.log_test(f"Sending: {test_msg}")
        
        # Process through normal message handler
        self.on_chat_message(test_msg)
        
    def test_as_admin(self):
        """Test message as an admin user"""
        # Get first admin from list or use default
        admins = [self.admin_listbox.get(i) for i in range(self.admin_listbox.size())]
        admin_name = admins[0] if admins else "AdminUser"
        
        self.test_username_var.set(admin_name)
        self.log_test(f"Testing as admin: {admin_name}")
        self.send_test_message()
        
    def test_as_blocked(self):
        """Test message as a blocked user"""
        # Get first blocked user or create test one
        blocked = [self.blocked_listbox.get(i) for i in range(self.blocked_listbox.size())]
        blocked_name = blocked[0] if blocked else "BlockedUser"
        
        # Add to blocked list if not there
        if blocked_name == "BlockedUser" and blocked_name not in blocked:
            self.blocked_listbox.insert(tk.END, blocked_name)
            
        self.test_username_var.set(blocked_name)
        self.log_test(f"Testing as blocked user: {blocked_name}")
        self.send_test_message()
        
    def test_voice_command(self):
        """Test a voice command"""
        # Set a voice command message
        self.test_message_var.set("/v 2 This is voice command test")
        self.log_test("Testing voice command: /v 2")
        self.send_test_message()
        
    def test_auto_block(self):
        """Test auto-block functionality"""
        # Get first auto-block keyword
        keywords = self.config.get('auto_block_keywords', [])
        if keywords:
            keyword = keywords[0]
            self.test_message_var.set(f"This message contains {keyword}")
            self.log_test(f"Testing auto-block with keyword: {keyword}")
        else:
            self.test_message_var.set("Add auto-block keywords first")
            self.log_test("No auto-block keywords configured")
            
        self.send_test_message()
        
    def test_private_mode(self):
        """Test private mode (only admins can use TTS)"""
        # Toggle private mode on
        old_private = self.private_mode_var.get()
        self.private_mode_var.set(True)
        
        # Test as non-admin
        self.test_username_var.set("RegularUser")
        self.test_message_var.set("Testing private mode - should not speak")
        self.log_test("Testing private mode with non-admin user")
        self.send_test_message()
        
        # Restore private mode setting
        self.private_mode_var.set(old_private)
    
    def test_voice_toggle(self):
        """Test voice toggle command to set user's default voice"""
        username = self.test_username_var.get()
        if not username:
            username = "TestUser"
            self.test_username_var.set(username)
        
        # Test setting voice preference with voice toggle command
        import random
        voice_num = random.randint(0, min(9, len(self.voices) - 1))
        self.test_message_var.set(f"{self.voice_toggle_command} {voice_num} Testing voice toggle")
        self.log_test(f"Testing voice toggle for {username} with voice {voice_num}")
        self.send_test_message()
        
        # Log the result
        if username in self.user_voice_preferences:
            self.log_test(f"✓ Voice preference set for {username}: {self.user_voice_preferences[username]}")
        else:
            self.log_test(f"✗ Voice preference not set for {username}")
    
    def test_user_preference(self):
        """Test that a user's voice preference is applied to regular messages"""
        username = self.test_username_var.get()
        if not username:
            username = "TestUser"
            self.test_username_var.set(username)
        
        if username in self.user_voice_preferences:
            voice_pref = self.user_voice_preferences[username]
            self.test_message_var.set("This should use my preferred voice")
            self.log_test(f"Testing message with {username}'s preferred voice: {voice_pref}")
            self.send_test_message()
        else:
            self.log_test(f"No voice preference set for {username}. Set one first with 'Test Voice Toggle'")
    
    def clear_test_voice_preference(self):
        """Clear the voice preference for the test user"""
        username = self.test_username_var.get()
        if not username:
            username = "TestUser"
            self.test_username_var.set(username)
        
        if username in self.user_voice_preferences:
            del self.user_voice_preferences[username]
            self.save_config()
            self.update_user_voices_list()
            self.log_test(f"Cleared voice preference for {username}")
        else:
            self.log_test(f"No voice preference found for {username}")
        
    def set_test_preset(self, username, message):
        """Set test fields to preset values"""
        self.test_username_var.set(username)
        self.test_message_var.set(message)
        self.log_test(f"Preset loaded: {username} - {message}")
        
    def set_admin_stop_preset(self):
        """Set admin stop command preset using current TTS prefix"""
        current_tts_prefix = self.get_current_game_config().get('tts_command_prefix', '!tts')
        self.set_test_preset("AdminUser", f"{current_tts_prefix} stop")
        
    def log_test(self, message):
        """Log message to test output"""
        timestamp = time.strftime("%H:%M:%S")
        self.test_output.insert(tk.END, f"[{timestamp}] {message}\n")
        self.test_output.see(tk.END)
        
    def apply_settings(self):
        """Apply all settings"""
        game_key = 'tf2' if self.current_game == 'Team Fortress 2' else 'drg'
        
        # Save game-specific settings
        if 'games' not in self.config:
            self.config['games'] = {}
        if game_key not in self.config['games']:
            self.config['games'][game_key] = {}
        
        self.config['games'][game_key]['log_path'] = self.log_path_var.get()
        
        # Save global settings
        self.config['admin_username'] = self.admin_username_var.get()
        self.config['private_mode'] = self.private_mode_var.get()
        self.config['auto_block'] = self.auto_block_var.get()
        self.config['default_voice'] = self.default_voice_combo.get()
        self.config['current_game'] = self.current_game
        
        # Legacy support for TF2
        if game_key == 'tf2':
            self.config['log_path'] = self.log_path_var.get()
        
        # Apply the default voice immediately to the audio manager
        default_voice = self.default_voice_combo.get()
        if default_voice and self.audio_manager:
            for voice in self.voices:
                if voice.name == default_voice:
                    if self.current_voice_id != voice.id:
                        success = self.audio_manager.set_voice(voice.id)
                        if success:
                            self.current_voice_id = voice.id
                        logger.info(f"Applied voice: {default_voice} (ID: {voice.id}) - Success: {success}")
                    else:
                        logger.debug(f"Already using voice: {default_voice}")
                    break
        
        self.save_config()
        messagebox.showinfo("Settings", "Settings applied!")
        
    def run(self):
        """Run the application"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()
        
    def on_closing(self):
        """Handle window closing"""
        if self.is_running:
            self.stop_tts()
        if self.audio_manager:
            self.audio_manager.stop()
        self.save_config()
        self.root.destroy()


def main():
    """Main entry point"""
    # Check for test mode
    
    try:
        logger.info("Creating main application...")
        app = TTSReplicaWASAPI()
        logger.info("Starting application...")
        app.run()
    except Exception as e:
        error_msg = f"Fatal error: {e}\n\nTraceback:\n{traceback.format_exc()}"
        logger.error(error_msg)
        try:
            messagebox.showerror("Fatal Error", str(e))
        except:
            pass  # GUI might not be available
        sys.exit(1)


if __name__ == '__main__':
    main()