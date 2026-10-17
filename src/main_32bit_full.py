#!/usr/bin/env python3
"""
TTS for Team Fortress 2 - Replica with WASAPI Audio
Enhanced version with direct device audio routing
"""

import sys
import os
import time
import logging
import logging.handlers
import json
import re
import shutil
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import ui_theme  # Windows 11 look, dark mode, scaling, scrolling tabs
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
    logging.getLogger(__name__).warning(f"DRG monitor not available: {e}")  # module logger isn't set up yet
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
# Chat is never written to the log (only lengths, voices and names), and the log is capped at about 10 MB:
# tts_tf2.log plus the previous one as tts_tf2.log.1
log_file = log_dir / 'tts_tf2.log'
# Console output only when there is a console (the windowed exe has none, so sys.stderr is None there)
console_handlers = [logging.StreamHandler()] if sys.stderr else []
try:
    logging.basicConfig(
        level=logging.DEBUG,  # More verbose for debugging
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.handlers.RotatingFileHandler(log_file, maxBytes=5_000_000, backupCount=1, encoding='utf-8')] + console_handlers
    )
except Exception as e:
    # Fallback if can't create log in exe dir
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.handlers.RotatingFileHandler('tts_tf2.log', maxBytes=5_000_000, backupCount=1, encoding='utf-8')] + console_handlers
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

# DJ music player and chat soundboard (optional - the TTS works without them)
try:
    from dj_player import DJEngine, DJWindow, clean_title
    DJ_AVAILABLE = True
except Exception as e:
    logger.warning(f"DJ player not available: {e}")
    DJ_AVAILABLE = False
try:
    from soundboard import SoundboardEngine, SoundboardWindow
    SOUNDBOARD_AVAILABLE = True
except Exception as e:
    logger.warning(f"Soundboard not available: {e}")
    SOUNDBOARD_AVAILABLE = False
try:
    from textaloud_voice import TextAloudVoices, PREFIX as TEXTALOUD_PREFIX, is_installed as textaloud_installed
    TEXTALOUD_AVAILABLE = True
except Exception as e:
    logger.warning(f"TextAloud voices not available: {e}")
    TEXTALOUD_AVAILABLE = False
    TEXTALOUD_PREFIX = "[TextAloud] "


try:
    import voice_installer  # more Windows voices, free Piper voices
    from voice_installer import PIPER_PREFIX
    VOICES_AVAILABLE = True
except Exception as e:
    logger.warning(f"Voice downloads not available: {e}")
    VOICES_AVAILABLE = False
    PIPER_PREFIX = "[Piper] "


try:
    import sapi4_voice  # old SAPI 4 voices: Microsoft Sam 4.0, TruVoice (BonziBuddy)...
    from sapi4_voice import PREFIX as SAPI4_PREFIX
    SAPI4_AVAILABLE = True
except Exception as e:
    logger.warning(f"SAPI 4 voices not available: {e}")
    SAPI4_AVAILABLE = False
    SAPI4_PREFIX = "[SAPI4] "


def is_special_voice(name):
    """DECtalk, TextAloud and Piper voices are played by the speech thread itself, not by the Windows voice worker"""
    return bool(name) and (name.startswith("[DECtalk] ") or name.startswith(TEXTALOUD_PREFIX)
                           or name.startswith(PIPER_PREFIX) or name.startswith(SAPI4_PREFIX))

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

# Console lines that contain " : " but aren't chat, so they'd only clutter the chat window:
# - the "status" command's server-info header, e.g. "version : 11117211/24 ..." or "udp/ip  : 1.2.3.4:27015"
#   (TF2's own chat has two spaces after the colon, so a player actually named "map" still gets through)
# - material errors, e.g. 'Error: Material "models/..." : proxy "..." unable to initialize!'
NOT_CHAT = re.compile(r'^(?:(?:hostname|version|udp/ip|account|steamid|map|tags|players|edicts|sourcetv|replay)\s*: (?! )|Error: Material ")')

# What TF2 puts in front of a chatter's name: "*DEAD* ", "*SPEC* ", "(TEAM) ", "*DEAD*(TEAM) ", "(Spectator) ".
# Only stripped from the start of the line, so they can't hide inside a name (e.g. "(TEAM)AdminName")
CHAT_PREFIX = re.compile(r'^(?:\*DEAD\*|\*SPEC\*)?(?:\(TEAM\)|\(Spectator\))?(?<=[*)]) ')

# DECtalk markup: inline commands like [:np] [:phoneme on] [:t 430,500] [:dial 555], or phonemes with
# timing like [dah<600,20>] / Moonbase-style [<500,20>]space. SAPI voices read these out as gibberish.
# Command names are 2+ letters ([:np], [:rate 200]) or one letter plus numbers ([:t 430,500]), so smileys
# like [:D] or [:P] don't count.
DECTALK_MARKUP = re.compile(r'\[:\s*(?:[a-zA-Z]{2,}|[a-zA-Z]\s*\d)[^\]]*\]|\[[^\]]*<\d+(?:,\d+)?>')

# Announcement texts for types added after people already had a config.json (used until they press Save
# on the Announcements tab). Spoken as "<name> <text>".
DEFAULT_ANNOUNCEMENTS = {
    'MUTED': "has been muted",
    'SPAM MUTED': "muted for spamming",
    'SONG REQUESTED': "requested {title}",                     # {title} = the song
    'REQUEST DENIED': "your request was denied, {reason}",     # {reason} = e.g. "it's too long..."
    'REQUEST PENDING': "requested {title}, waiting for approval",
    'REQUEST APPROVED': "your request was approved, {title}",
}

# Games every config starts with; any others are added by the user in Settings
BUILTIN_GAME_NAMES = {'tf2': 'Team Fortress 2', 'drg': 'Deep Rock Galactic'}


def spam_signature(text):
    """Loose fingerprint for the spam guard: case and spacing ignored and long character runs shortened, so
    "7777777777" and "77777777777777", or "aaaaaaa" and "aaaaaaaaaaaa", count as the same message"""
    text = re.sub(r'\s+', ' ', text.lower()).strip()
    return re.sub(r'(.)\1{2,}', r'\1\1\1', text)

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

        # Skip console lines that aren't chat (status server-info header, material errors)
        if NOT_CHAT.match(line):
            return None

        # Debug logging for !block commands
        if "!block" in line.lower():
            logger.info("LOG PARSER: Found a !block line")
            
        # Chat prefix in front of the name (spectators included, so blocks/admins/voices still match them)
        prefix_match = CHAT_PREFIX.match(line)
        prefix = prefix_match.group(0) if prefix_match else ""
        is_dead = "*DEAD*" in prefix
        is_team = "(TEAM)" in prefix
        is_spec = "*SPEC*" in prefix or "(Spectator)" in prefix

        # Clean line
        clean = line[len(prefix):].strip()
        
        # Extract username and message
        if " : " in clean:
            parts = clean.split(" : ", 1)
            if len(parts) == 2:
                username = parts[0].strip().strip('"')
                message = parts[1].strip()
                
                # Debug logging for parsed !block commands
                if "!block" in message.lower():
                    logger.info(f"LOG PARSER: Parsed !block command from '{username}'")
                
                return {
                    'username': username,
                    'message': message,
                    'is_dead': is_dead,
                    'is_team': is_team,
                    'is_spec': is_spec,
                    'raw': line
                }
        
        # Debug logging for unparsed !block lines
        if "!block" in line.lower():
            logger.info("LOG PARSER: Failed to parse a !block line")
            
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
        self._config_lock = threading.RLock()  # save_config is called from the GUI, chat and speech threads
        # Clear temp leftovers from earlier runs (crash/force-close) in the background so startup isn't slowed
        threading.Thread(target=self.cleanup_leftover_temp, daemon=True).start()
        self.config = self.load_config()
        self.monitors = {}  # Dictionary to hold monitors for each game
        self.current_game = self.normalize_game_name(self.config.get('current_game', 'Team Fortress 2'))
        self.audio_manager = None
        self.is_running = False
        
        # Initialize hidden voices list
        self.hidden_voices = self.config.get('hidden_voices', [])
        
        # Game configurations
        self.game_configs = self.init_game_configs()
        # If the saved game was removed (or is an old 'tf2' style name), land on a game that exists
        self.current_game = self.game_configs[self.current_game_key()]['name']

        # Track last speaker for admin commands (per game)
        self.last_speakers = {'tf2': None, 'drg': None}
        
        # GUI updates requested from the chat/speech/DJ threads; pump_ui_calls runs them on the GUI thread
        # (Tk isn't thread-safe, so nothing outside the GUI thread touches widgets directly)
        self._ui_calls = queue.Queue()
        self.dj = None          # DJ player, started after the window is built
        self.dj_window = None
        self.soundboard = None  # chat soundboard, same
        self.soundboard_window = None

        # Message queue for sequential playback
        self.message_queue = queue.Queue()
        self.queue_thread = None
        self.queue_running = False
        self.currently_speaking_user = None
        self.last_speaker = None  # last player to use TTS (for !block / !admin add / !mute)
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

        # Temporary mutes: username -> time the mute ends (kept across restarts, oldest mute first)
        self._mute_lock = threading.Lock()
        self.muted_users = {u: t for u, t in self.config.get('muted_users', {}).items() if t > time.time()}
        self.spam_history = {}  # username -> recent [(time, spam_signature)] of their TTS messages
        self._last_denied_announce = {}  # username -> when "request denied" was last spoken for them

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
        # Plain float copy of the DECtalk volume slider for the speech thread (Tk variables aren't thread-safe)
        self.dectalk_volume = float(self.config.get('dectalk_volume', 0.7))
        # TextAloud voices (Will From Afar etc.). Their names are remembered in config "textaloud_voices", so
        # TextAloud only has to start when one of them actually speaks.
        # It closes after config "textaloud_idle_minutes" (default 10) without use and starts again when needed.
        self.textaloud = (TextAloudVoices(idle_minutes=float(self.config.get('textaloud_idle_minutes', 10)))
                          if TEXTALOUD_AVAILABLE else None)
        self.textaloud_enabled = bool(self.config.get('textaloud_enabled', True))
        self.textaloud_volume = float(self.config.get('textaloud_volume', 1.0))
        # Free Piper voices and the Windows-voices undo file live in the "voices" folder next to the TTS
        self.voices_dir = Path(self.config.get('voices_dir') or self.get_config_path().parent / 'voices')
        self.piper = voice_installer.PiperVoices(self.voices_dir) if VOICES_AVAILABLE else None
        self.piper_volume = float(self.config.get('piper_volume', 1.0))
        # Old SAPI 4 voices (Microsoft Sam 4.0, TruVoice/BonziBuddy...) if SAPI 4 is installed on this PC
        self.sapi4 = sapi4_voice.Sapi4Voices() if SAPI4_AVAILABLE else None
        self.sapi4_volume = float(self.config.get('sapi4_volume', 1.0))

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

        # Keep the MUTED list's "min left" current and drop mutes that have run out
        self.refresh_muted_list()
        self.root.after(30000, self.tick_mutes)
        self.root.after(100, self.pump_ui_calls)

        # Old SAPI 4 voices: ask the helper which are installed, then add them to the voice lists
        if self.sapi4 and self.sapi4.available():
            def find_sapi4():
                self.sapi4.refresh()
                self.run_on_gui(self.on_sapi4_voices)
            threading.Thread(target=find_sapi4, daemon=True).start()

        # First run with TextAloud installed: find its voices once (later runs use the saved list)
        if self.textaloud_enabled and 'textaloud_voices' not in self.config:
            self.refresh_textaloud_voices()

        # DJ music player: the window opens from the 🎵 DJ button; music keeps playing when it's closed
        self.dj = None
        self.dj_window = None
        if DJ_AVAILABLE:
            try:
                self.dj = DJEngine(self.config.setdefault('dj', {}), save_config=self.save_config,
                                   tts_active=self.is_tts_audio_active, announce=self.announce_now_playing,
                                   tts_device=lambda: self.config.get('audio_device', 'Default'),
                                   app_dir=self.get_config_path().parent)
            except Exception as e:
                logger.error(f"DJ player failed to start: {e}")
        self.root.after(1000, self.tick_dj_badge)

        # Chat soundboard: sounds in the "sounds" folder next to the TTS, triggered from chat
        if SOUNDBOARD_AVAILABLE:
            try:
                # same ffmpeg/ffprobe lookup as the DJ (settings path, next to the TTS, then PATH)
                finder = (lambda name: (getattr(self.dj, f'{name}_path')() if self.dj else shutil.which(name)))
                self.soundboard = SoundboardEngine(self.config.setdefault('soundboard', {}), save_config=self.save_config,
                                                   sounds_dir=self.get_config_path().parent / 'sounds',
                                                   tts_device=lambda: self.config.get('audio_device', 'Default'),
                                                   ffmpeg_path=lambda: finder('ffmpeg'),
                                                   ffprobe_path=lambda: finder('ffprobe'))
            except Exception as e:
                logger.error(f"Soundboard failed to start: {e}")

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
            # Ensure the freshly-created engine uses the user's profiles, not its hardcoded defaults
            if hasattr(self, 'dectalk_profiles') and self.dectalk_profiles:
                self.dectalk_native.voice_profiles = self.dectalk_profiles
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
        backup_path = config_path.with_name(config_path.name + '.bak')

        # Try config.json, then the copy save_config keeps of the previous version
        for path in (config_path, backup_path):
            if not path.exists():
                continue
            try:
                with open(path, 'r') as f:
                    user_config = json.load(f)

                    if path == backup_path:
                        logger.warning(f"Loaded settings from {backup_path.name}")

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
                logger.error(f"Error loading config {path.name}: {e}")
                if path == config_path:
                    # Move the unreadable file aside so the next save can't overwrite your settings with defaults
                    corrupt_path = config_path.with_name(f"config_unreadable_{time.strftime('%Y%m%d_%H%M%S')}.json")
                    try:
                        os.replace(config_path, corrupt_path)
                        logger.error(f"Moved unreadable config to {corrupt_path.name}")
                    except Exception as move_error:
                        logger.error(f"Could not move unreadable config aside: {move_error}")

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
        """Initialize game configurations. Each game: name (shown in Settings), type ('source' = Source engine
        console log, 'drg' = Deep Rock Galactic chat mod CSV), log_path, admins, blocked, tts_command_prefix."""
        # Same dict object as config['games'], so edits here are what gets saved
        games = self.config.setdefault('games', {})

        # A fresh config starts with the two built-in games (they can be removed later in Settings)
        if not games:
            games['tf2'] = {
                'log_path': self.config.get('log_path', ''),
                'admins': self.config.get('admins', []),
                'blocked': self.config.get('blocked', []),
                'last_position': 0
            }
            games['drg'] = {
                'log_path': '',
                'admins': [],
                'blocked': [],
                'last_line_index': -1
            }

        for key, game in games.items():
            game.setdefault('name', BUILTIN_GAME_NAMES.get(key, key))
            game.setdefault('type', 'drg' if key == 'drg' else 'source')

        return games

    def game_key_for(self, name):
        """Config key of the game with this display name (falls back to the first game)"""
        for wanted in (name, self.normalize_game_name(name)):
            for key, game in self.game_configs.items():
                if game.get('name') == wanted:
                    return key
        return next(iter(self.game_configs), 'tf2')

    def current_game_key(self):
        return self.game_key_for(self.current_game)

    def game_names(self):
        return [game.get('name', key) for key, game in self.game_configs.items()]
    
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
            
            # Set current game (the display name - everything else compares against 'Team Fortress 2')
            config['current_game'] = 'Team Fortress 2'
            
            # Remove old format keys
            for key in ['log_path', 'tf2', 'tf2_log_path']:
                config.pop(key, None)
        
        return config
    
    def run_on_gui(self, func):
        """Have func run on the GUI thread shortly. Safe from any thread."""
        self._ui_calls.put(func)

    def pump_ui_calls(self):
        """GUI thread, every 100 ms: run the updates other threads asked for, in order"""
        while True:
            try:
                func = self._ui_calls.get_nowait()
            except queue.Empty:
                break
            try:
                func()
            except Exception as e:
                logger.error(f"GUI update failed: {e}")
        self.root.after(100, self.pump_ui_calls)

    def open_dj_window(self):
        """🎵 DJ button"""
        if not self.dj:
            messagebox.showerror("DJ", "The DJ player couldn't start - see tts_tf2.log for why")
            return
        if self.dj_window:
            self.dj_window.lift()
            return
        voices = lambda: list(self.longform_voice_combo['values']) if hasattr(self, 'longform_voice_combo') else ['Default']
        self.dj_window = DJWindow(self.root, self.dj, voice_names=voices,
                                  on_close=lambda: setattr(self, 'dj_window', None))

    def tick_dj_badge(self):
        """Every second: "🎵 DJ (2)" on the main window while song requests wait for approval"""
        count = len(self.dj.pending) if self.dj else 0
        text = f"🎵 DJ ({count})" if count else "🎵 DJ"
        if self.dj_btn.cget('text') != text:
            self.dj_btn.config(text=text)
        self.root.after(1000, self.tick_dj_badge)

    def open_soundboard_window(self):
        """🔊 Sounds button"""
        if not self.soundboard:
            messagebox.showerror("Soundboard", "The soundboard couldn't start - see tts_tf2.log for why")
            return
        if self.soundboard_window:
            self.soundboard_window.lift()
            return
        self.soundboard_window = SoundboardWindow(self.root, self.soundboard, tools=self.dj,
                                                  on_close=lambda: setattr(self, 'soundboard_window', None))

    def is_tts_audio_active(self):
        """For DJ ducking: is a TTS voice talking right now - or a soundboard sound, if the soundboard is set to
        lower the music? (called from the DJ thread ~20 times a second)"""
        sb = getattr(self, 'soundboard', None)
        if sb and sb.get('duck_music') and sb.is_playing():
            return True
        try:
            if self.audio_manager and self.audio_manager.is_speaking():
                return True
        except Exception:
            pass
        dectalk = getattr(self, 'dectalk_native', None)
        return bool(dectalk and hasattr(dectalk, 'is_playing') and dectalk.is_playing())

    def announce_now_playing(self, text, track):
        """DJ "now playing" through the TTS - only while the TTS is started (nothing else empties the queue)"""
        if not self.is_running:
            return
        self.message_queue.put({'text': text, 'voice': self.dj_voice(), 'username': '__announcement__'})

    def cleanup_leftover_temp(self):
        """Old unpack folders and stray DECtalk WAVs from earlier runs"""
        try:
            from temp_utils import cleanup_orphaned_bundle_dirs, cleanup_old_temp_files
            cleanup_orphaned_bundle_dirs()
            cleanup_old_temp_files(max_age_hours=1)
        except Exception as e:
            logger.warning(f"Temp cleanup skipped: {e}")

    def normalize_game_name(self, name):
        """Older configs stored 'tf2'/'drg', but the code compares against the display names"""
        return {'tf2': 'Team Fortress 2', 'drg': 'Deep Rock Galactic', '': 'Team Fortress 2'}.get(name or '', name)

    def get_current_game_config(self):
        """Get config for currently selected game"""
        game_key = self.current_game_key()
        return self.game_configs.get(game_key, {})
    
    def on_game_changed(self, event=None):
        """Handle game selection change"""
        # Save current game's log path before switching (unless that game was just removed)
        old_game_key = self.current_game_key()
        if hasattr(self, 'log_path_var') and self.game_configs.get(old_game_key, {}).get('name') == self.current_game:
            current_log_path = self.log_path_var.get()
            if current_log_path:
                self.game_configs[old_game_key]['log_path'] = current_log_path
        
        selected_game = self.game_combo.get()
        game_key = self.game_key_for(selected_game)
        
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

    def add_game(self):
        """+ Add Game button"""
        self.open_game_dialog(None)

    def edit_game(self):
        """✎ Edit button: the game picked in the Game box"""
        self.open_game_dialog(self.game_key_for(self.game_combo.get()))

    def open_game_dialog(self, game_key):
        """Add a game (game_key None) or edit one: name, log type, log file, TTS command"""
        editing = game_key is not None
        game = self.game_configs.get(game_key, {}) if editing else {}
        type_labels = {
            'source': "Source game console log (TF2, L4D2, GMod, CS:S, HL2DM...)",
            'drg': "Deep Rock Galactic chat log mod",
        }

        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Game" if editing else "Add Game")
        dialog.transient(self.root)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Name").grid(row=0, column=0, sticky=tk.W, pady=4)
        name_var = tk.StringVar(value=game.get('name', ''))
        ttk.Entry(frame, textvariable=name_var, width=50).grid(row=0, column=1, sticky=tk.W, pady=4)

        ttk.Label(frame, text="Log type").grid(row=1, column=0, sticky=tk.W, pady=4)
        type_combo = ttk.Combobox(frame, values=list(type_labels.values()), state='readonly', width=48)
        type_combo.set(type_labels.get(game.get('type', 'source'), type_labels['source']))
        type_combo.grid(row=1, column=1, sticky=tk.W, pady=4)

        ttk.Label(frame, text="Log file").grid(row=2, column=0, sticky=tk.W, pady=4)
        path_var = tk.StringVar(value=game.get('log_path', ''))
        ttk.Entry(frame, textvariable=path_var, width=50).grid(row=2, column=1, sticky=tk.W, pady=4)

        def browse():
            path = filedialog.askopenfilename(parent=dialog, title="Pick the game's log file",
                                              filetypes=[("Log files", "*.txt *.log *.csv"), ("All files", "*.*")])
            if path:
                path_var.set(path)
        ttk.Button(frame, text="Browse...", command=browse).grid(row=2, column=2, padx=5, pady=4)

        ttk.Label(frame, text="TTS command").grid(row=3, column=0, sticky=tk.W, pady=4)
        prefix_var = tk.StringVar(value=game.get('tts_command_prefix', '!tts'))
        ttk.Entry(frame, textvariable=prefix_var, width=15).grid(row=3, column=1, sticky=tk.W, pady=4)

        ttk.Label(frame, style='Hint.TLabel', justify=tk.LEFT, text=(
            "Source games: add  con_logfile log.txt  to that game's autoexec.cfg (in its cfg folder),\n"
            "then pick the log.txt it creates in the game's mod folder, e.g. ...\\left4dead2\\log.txt.\n"
            + ("" if editing else "A new game starts with this game's admins and an empty blocked list.")
        )).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(8, 4))

        def save():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Game", "Give the game a name", parent=dialog)
                return
            if any(g.get('name') == name for k, g in self.game_configs.items() if k != game_key):
                messagebox.showwarning("Game", f"There's already a game called {name}", parent=dialog)
                return
            game_type = next(k for k, label in type_labels.items() if label == type_combo.get())
            is_current = editing and game.get('name') == self.current_game

            if editing:
                target = self.game_configs[game_key]
                key = game_key
            else:
                base = re.sub(r'[^a-z0-9]+', '_', name.lower()).strip('_') or 'game'
                key, n = base, 2
                while key in self.game_configs:
                    key, n = f"{base}_{n}", n + 1
                admins = [self.admin_listbox.get(i) for i in range(self.admin_listbox.size())]
                target = self.game_configs[key] = {'admins': admins, 'blocked': [], 'last_position': 0}
            target.update({
                'name': name,
                'type': game_type,
                'log_path': path_var.get().strip(),
                'tts_command_prefix': prefix_var.get().strip() or '!tts',
            })
            if key == 'tf2':
                self.config['log_path'] = target['log_path']  # legacy location

            self.game_combo['values'] = self.game_names()
            if is_current:
                # Renamed or re-pointed the game in use: refresh the Settings fields (and restart if running)
                was_running = self.is_running
                if was_running:
                    self.stop_tts()
                self.current_game = name
                self.config['current_game'] = name
                self.game_combo.set(name)
                self.log_path_var.set(target['log_path'])
                self.tts_prefix_entry.delete(0, tk.END)
                self.tts_prefix_entry.insert(0, target['tts_command_prefix'])
                if was_running:
                    self.start_tts()
            self.save_config()
            logger.info(f"{'Edited' if editing else 'Added'} game {name} ({key}, {game_type})")
            dialog.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=5, column=0, columnspan=3, pady=(8, 0))
        ttk.Button(btns, text="Save", command=save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def remove_game(self):
        """✖ Remove button: delete the game picked in the Game box, with its lists and settings"""
        name = self.game_combo.get()
        key = self.game_key_for(name)
        if len(self.game_configs) <= 1:
            messagebox.showwarning("Remove Game", "You need at least one game")
            return
        if not messagebox.askyesno("Remove Game", f"Remove {name}?\n\nIts log path, admins, blocked list and TTS command go with it."):
            return
        if self.is_running and name == self.current_game:
            self.stop_tts()
        del self.game_configs[key]
        self.game_combo['values'] = self.game_names()
        logger.info(f"Removed game {name} ({key})")
        if name == self.current_game:
            # Switch to the first remaining game; current_game is cleared so its log path isn't saved
            # into the removed game
            self.current_game = ''
            self.game_combo.set(self.game_names()[0])
            self.on_game_changed()
        else:
            self.game_combo.set(self.current_game)
            self.save_config()

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
        ui_theme.size_window(export_dialog, 400, 300, self.root)
        
        ttk.Label(export_dialog, text="Select sections to export:", style='Section.TLabel').pack(pady=10)
        
        # Checkboxes for sections
        export_options = {}
        sections = [
            ('games', 'Games (log paths, admins, blocked lists, TTS commands)'),
            ('voices', 'Voice settings (default voice, hidden voices, DECtalk profiles, random voices)'),
            ('voice_commands', 'Voice commands'),
            ('user_voices', 'User voices (each player\'s voice)'),
            ('announcements', 'Announcements'),
            ('auto_block', 'Auto-block words'),
            ('ui', 'Everything else (private mode, mutes, DJ, soundboard...)'),
            ('audio_device', 'Audio device')
        ]
        # The config keys behind each checkbox (the old export looked for keys named after the checkboxes,
        # which mostly don't exist, so user voices, auto-block words, DECtalk profiles... were never exported)
        section_keys = {
            'games': ['games', 'current_game', 'admins', 'blocked', 'log_path'],
            'voices': ['default_voice', 'hidden_voices', 'dectalk_profiles', 'dectalk_volume', 'dectalk_enabled',
                       'random_voice_enabled', 'random_voice_exclusions', 'voice_toggle_command',
                       'textaloud_enabled', 'textaloud_volume', 'textaloud_voices', 'textaloud_idle_minutes',
                       'piper_volume', 'sapi4_volume'],
            'voice_commands': ['voice_commands'],
            'user_voices': ['user_voice_preferences'],
            'announcements': ['announcements', 'announcement_enabled', 'announcement_voice'],
            'auto_block': ['auto_block', 'auto_block_enabled', 'auto_block_keywords'],
            'audio_device': ['audio_device'],
        }
        named = {k for keys in section_keys.values() for k in keys}
        section_keys['ui'] = [k for k in self.config if k not in named]
        
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
                    for section, var in export_options.items():
                        if var.get():
                            for key in section_keys[section]:
                                if key in self.config:
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
            old_config = self.config
            try:
                # Create backup of current config
                backup_path = self.get_config_path().parent / f"config_backup_{time.strftime('%Y%m%d_%H%M%S')}.json"
                with open(backup_path, 'w') as f:
                    json.dump(self.config, f, indent=2)
                logger.info(f"Created config backup at {backup_path}")
                
                with open(file_path, 'r') as f:
                    imported_config = json.load(f)
                if not isinstance(imported_config, dict):
                    raise ValueError("that file isn't a TTS config")
                imported_config.pop('version', None)

                # Merge INTO the current settings: an export of only some sections changes only those, and
                # nothing the file doesn't mention gets wiped (replacing the whole config used to do that)
                current = json.loads(json.dumps(self.config))  # deep copy
                self.config = self.merge_configs(current, imported_config)

                # Save and reload
                self.apply_config()
                self.save_config()
                self.load_settings()
                
                messagebox.showinfo("Import Successful", 
                                  "Config imported successfully.\nSome settings may require restart.")
                logger.info(f"Config imported from {file_path}")
            except Exception as e:
                # Don't leave a half-applied import live (the next save would write it out)
                if self.config is not old_config:
                    self.config = old_config
                    try:
                        self.apply_config()
                        self.load_settings()
                    except Exception as restore_error:
                        logger.error(f"Failed to restore settings after import error: {restore_error}")
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
            
            try:
                # Back up the current settings first, same as Import does
                backup_path = self.get_config_path().parent / f"config_backup_{time.strftime('%Y%m%d_%H%M%S')}.json"
                with open(backup_path, 'w') as f:
                    json.dump(self.config, f, indent=2)
                logger.info(f"Created config backup at {backup_path}")

                # default_config.json isn't shipped, so without it "defaults" means a fresh, empty config
                if default_config_path.exists():
                    with open(default_config_path, 'r') as f:
                        self.config = json.load(f)
                else:
                    self.config = {}

                # Preserve TF2 log path if it exists
                current_log = self.log_path_var.get()
                if current_log and Path(current_log).exists():
                    self.config['log_path'] = current_log

                self.apply_config()
                self.save_config()
                self.load_settings()

                messagebox.showinfo("Reset Complete", f"Settings reset to defaults.\nYour old settings were saved to {backup_path.name}")
                logger.info("Config reset to defaults")
            except Exception as e:
                messagebox.showerror("Reset Failed", f"Failed to reset config:\n{e}")
                logger.error(f"Failed to reset config: {e}")
    
        
    def save_config(self):
        """Save config: one save at a time, written to a temp file and swapped in (a crash mid-save can't
        leave a half-written config.json), with the previous version kept as config.json.bak"""
        config_path = self.get_config_path()

        with self._config_lock:
            try:
                # Chat/speech threads can change settings while we serialize; retry if a dict changes mid-way
                data = None
                for attempt in range(5):
                    try:
                        data = json.dumps(self.config, indent=2)
                        break
                    except RuntimeError:
                        time.sleep(0.02)
                if data is None:
                    logger.error("Failed to save config: settings kept changing during save")
                    return

                tmp_path = config_path.with_name(config_path.name + '.tmp')
                with open(tmp_path, 'w') as f:
                    f.write(data)
                    f.flush()
                    os.fsync(f.fileno())
                if config_path.exists():
                    shutil.copy2(config_path, config_path.with_name(config_path.name + '.bak'))
                os.replace(tmp_path, config_path)
            except Exception as e:
                logger.error(f"Failed to save config: {e}")

    def apply_config(self):
        """Re-read the in-memory settings from self.config. Import/Reset/Reload replace self.config, and
        without this the next save would write the old voices, commands and lists straight back."""
        self.current_game = self.normalize_game_name(self.config.get('current_game', 'Team Fortress 2'))
        self.hidden_voices = self.config.get('hidden_voices', [])
        self.game_configs = self.init_game_configs()
        self.current_game = self.game_configs[self.current_game_key()]['name']
        if getattr(self, 'dj', None):
            self.dj.settings = self.config.setdefault('dj', {})  # the queue in memory is written into it on the next change
        if getattr(self, 'soundboard', None):
            self.soundboard.settings = self.config.setdefault('soundboard', {})
        with self._mute_lock:
            self.muted_users = {u: t for u, t in self.config.get('muted_users', {}).items() if t > time.time()}
        self.voice_commands = self.config.get('voice_commands', {})
        self.user_voice_preferences = self.config.get('user_voice_preferences', {})
        self.random_voice_enabled = self.config.get('random_voice_enabled', False)
        self.random_voice_exclusions = self.config.get('random_voice_exclusions', [])
        self.voice_toggle_command = self.config.get('voice_toggle_command', '/vt')
        self.dectalk_profiles = self.config.get('dectalk_profiles') or self.get_default_dectalk_profiles()
        self.hidden_voices = self.load_hidden_voices()
        self.dectalk_enabled = self.config.get('dectalk_enabled', True)

        # Refresh the widgets that show these (load_settings handles the rest)
        if hasattr(self, 'game_combo'):
            self.game_combo['values'] = self.game_names()
            self.game_combo.set(self.current_game)
        if hasattr(self, 'spam_mute_var'):
            self.spam_mute_var.set(self.config.get('spam_mute_enabled', True))
            self.mute_minutes_var.set(str(self.config.get('mute_minutes', 10)))
            self.spam_repeat_var.set(str(self.config.get('spam_repeat_count', 3)))
            self.refresh_muted_list()
        # Auto-block words and announcement fields always show the new config (load_settings skips an
        # empty keyword list), so a later Save can't write the old values back
        if hasattr(self, 'auto_block_text'):
            self.auto_block_text.delete(1.0, tk.END)
            self.auto_block_text.insert(1.0, '\n'.join(self.config.get('auto_block_keywords', [])))
        announcements = self.config.get('announcements') or {}
        for key, var in getattr(self, 'announcement_vars', {}).items():
            var.set(announcements.get(key, getattr(self, 'announcement_defaults', {}).get(key, '')))
        enabled = self.config.get('announcement_enabled', {})
        for label, var in getattr(self, 'announcement_enabled_vars', {}).items():
            var.set(enabled.get(label, True))
        if hasattr(self, 'announcement_voice_combo') and self.config.get('announcement_voice'):
            self.announcement_voice_combo.set(self.config['announcement_voice'])
        if hasattr(self, 'random_voice_var'):
            self.random_voice_var.set(self.random_voice_enabled)
        if hasattr(self, 'voice_toggle_entry'):
            self.voice_toggle_entry.delete(0, tk.END)
            self.voice_toggle_entry.insert(0, self.voice_toggle_command)
        self.dectalk_volume = float(self.config.get('dectalk_volume', 0.7))
        if hasattr(self, 'dectalk_volume_var'):
            self.dectalk_volume_var.set(self.dectalk_volume)
        self.textaloud_enabled = bool(self.config.get('textaloud_enabled', True))
        self.textaloud_volume = float(self.config.get('textaloud_volume', 1.0))
        if getattr(self, 'textaloud', None):
            self.textaloud.idle_minutes = float(self.config.get('textaloud_idle_minutes', 10))
        if hasattr(self, 'textaloud_enabled_var'):
            self.textaloud_enabled_var.set(self.textaloud_enabled)
            self.textaloud_volume_var.set(self.textaloud_volume)
            self.textaloud_idle_var.set(f"{float(self.config.get('textaloud_idle_minutes', 10)):g}")
            self.refresh_textaloud_lists()
        if hasattr(self, 'tts_prefix_entry'):
            self.tts_prefix_entry.delete(0, tk.END)
            self.tts_prefix_entry.insert(0, self.get_current_game_config().get('tts_command_prefix', '!tts'))
        if hasattr(self, 'dectalk_profiles_listbox'):
            self.refresh_dectalk_profiles()
        if not self.voice_commands or all(not v for v in self.voice_commands.values()):
            self.auto_generate_voice_commands()
        if hasattr(self, 'voice_tree'):
            for item in self.voice_tree.get_children():
                self.voice_tree.delete(item)
            for cmd, voice in self.voice_commands.items():
                self.voice_tree.insert('', tk.END, values=(cmd, voice), tags=(cmd,))
    
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
        ui_theme.size_window(dialog, 400, 500, self.root)
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Select voices to EXCLUDE from random assignment:", 
                 style='Section.TLabel').pack(pady=10)
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
        all_voices.update(self.extra_voice_names())

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
                if voice_name not in self.random_voice_exclusions and voice_name not in self.hidden_voices:
                    available_voices.append(voice_name)
        
        # Add DECtalk voices (if not excluded)
        if hasattr(self, 'dectalk_profiles') and self.dectalk_profiles:
            for profile_name in self.dectalk_profiles.keys():
                dectalk_voice = f"[DECtalk] {profile_name}"
                if dectalk_voice not in self.random_voice_exclusions and dectalk_voice not in self.hidden_voices:
                    available_voices.append(dectalk_voice)
        available_voices += [v for v in self.extra_voice_names()
                             if v not in self.random_voice_exclusions and v not in self.hidden_voices]

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
            
    def px(self, n):
        """Pixels at 100% Windows scaling -> pixels on this screen"""
        return int(round(n * getattr(self, 'ui_scale', 1.0)))

    def setup_gui(self):
        """Main window. Every part can shrink: the chat log, the speak box and the tabs share the height (drag the
        dividers to change the split), the lists on the right share theirs, and a tab that doesn't fit scrolls
        instead of being cut off. Size, position and divider positions are remembered."""
        ui_theme.enable_dpi_awareness()
        self.hidden_voices = self.load_hidden_voices()
        self.root = tk.Tk()
        self.root.title("TTS_GUI - WASAPI Enhanced")
        ui_theme.apply_forced_scale(self.root)
        self.ui_scale = ui_theme.scale_of(self.root)
        self.ui_mode = ui_theme.apply_theme(self.root, self.config.get('ui_theme', 'system'))
        ui_theme.fast_resize(self.root)
        self.restore_window_size()
        px = self.px

        # Everything sits in one frame that takes the window's new size once you pause dragging its edge.
        # Re-laying out and redrawing every control on each pixel of the drag is what made resizing stutter.
        self.content = ttk.Frame(self.root)
        self.content.place(x=0, y=0, relwidth=1, relheight=1)
        self._resize_job = None
        self._content_size = None
        self.root.bind_class('Tk', '<Configure>', self._on_window_configure)  # the main window only

        # Left: the main area in resizable sections | right: the lists (drag the divider to widen them)
        self.side_paned = tk.PanedWindow(self.content, orient=tk.HORIZONTAL, sashwidth=px(6), bd=0,
                                         sashrelief=tk.FLAT, opaqueresize=True)
        self.side_paned.pack(fill=tk.BOTH, expand=True, padx=px(6), pady=px(6))
        self.main_paned = tk.PanedWindow(self.side_paned, orient=tk.VERTICAL, sashwidth=px(6), bd=0,
                                         sashrelief=tk.FLAT, opaqueresize=True)
        right_frame = self.right_frame = ttk.Frame(self.side_paned)
        self.side_paned.add(self.main_paned, stretch='always', minsize=px(540))
        self.side_paned.add(right_frame, stretch='never', minsize=px(240))

        # Chat log, with the control buttons above it
        left_frame = ttk.Frame(self.main_paned)
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=(0, px(4)))

        self.start_btn = ttk.Button(button_frame, text="Start TTS", command=self.start_tts, style='Accent.TButton')
        self.start_btn.pack(side=tk.LEFT, padx=(0, px(4)))

        self.stop_btn = ttk.Button(button_frame, text="Stop TTS", command=self.stop_tts, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, px(4)))

        self.reload_btn = ttk.Button(button_frame, text="Reload TTS", command=self.reload_tts)
        self.reload_btn.pack(side=tk.LEFT, padx=(0, px(4)))

        self.dj_btn = ttk.Button(button_frame, text="🎵 DJ", command=self.open_dj_window)
        self.dj_btn.pack(side=tk.LEFT, padx=(0, px(4)))

        self.soundboard_btn = ttk.Button(button_frame, text="🔊 Sounds", command=self.open_soundboard_window)
        self.soundboard_btn.pack(side=tk.LEFT, padx=(0, px(4)))

        self.chat_display = ui_theme.ScrolledText(left_frame, height=6, width=50)
        self.chat_display.pack(fill=tk.BOTH, expand=True)
        self.main_paned.add(left_frame, stretch='always', minsize=px(110))

        # Right side - the lists share the height and scroll
        right_frame.columnconfigure(0, weight=1)
        self.admin_listbox, row = self.side_list(right_frame, 0, "ADMINS", [
            ("✓ Save", self.save_admins), ("✖ Remove", self.remove_admin), ("+ Add", self.add_admin)], weight=1)
        self.blocked_listbox, row = self.side_list(right_frame, row, "BLOCKED", [
            ("✓ Save", self.save_blocked), ("✖ Remove", self.remove_blocked), ("+ Add", self.add_blocked)], weight=2)

        # MUTED section - temporary mutes from !mute, the spam guard, or + Mute here
        self.muted_listbox, row = self.side_list(right_frame, row, "MUTED", [
            ("✖ Unmute", self.unmute_selected), ("+ Mute", self.add_mute)], weight=1)
        self.muted_listbox_names = []  # usernames in the same order as the listbox rows

        mute_opts_frame = ttk.Frame(right_frame)
        mute_opts_frame.grid(row=row, column=0, sticky='w', pady=(px(4), 0))
        self.spam_mute_var = tk.BooleanVar(value=self.config.get('spam_mute_enabled', True))
        ttk.Checkbutton(mute_opts_frame, text="Auto-mute spam", variable=self.spam_mute_var,
                        command=self.save_mute_settings).pack(side=tk.LEFT)
        mute_for_frame = ttk.Frame(right_frame)
        mute_for_frame.grid(row=row + 1, column=0, sticky='w', pady=(px(2), 0))
        ttk.Label(mute_for_frame, text="Mute for").pack(side=tk.LEFT, padx=(0, px(4)))
        self.mute_minutes_var = tk.StringVar(value=str(self.config.get('mute_minutes', 10)))
        mute_minutes_box = ttk.Spinbox(mute_for_frame, from_=1, to=1440, width=4,
                                       textvariable=self.mute_minutes_var, command=self.save_mute_settings)
        mute_minutes_box.pack(side=tk.LEFT)
        mute_minutes_box.bind('<FocusOut>', lambda e: self.save_mute_settings())
        ttk.Label(mute_for_frame, text="min").pack(side=tk.LEFT, padx=px(4))

        spam_opts_frame = ttk.Frame(right_frame)
        spam_opts_frame.grid(row=row + 3, column=0, sticky='w', pady=(px(2), 0))
        ttk.Label(right_frame, text="Spam = the same message").grid(row=row + 2, column=0, sticky='w', pady=(px(4), 0))
        self.spam_repeat_var = tk.StringVar(value=str(self.config.get('spam_repeat_count', 3)))
        spam_repeat_box = ttk.Spinbox(spam_opts_frame, from_=2, to=10, width=3,
                                      textvariable=self.spam_repeat_var, command=self.save_mute_settings)
        spam_repeat_box.pack(side=tk.LEFT, padx=(0, px(4)))
        spam_repeat_box.bind('<FocusOut>', lambda e: self.save_mute_settings())
        ttk.Label(spam_opts_frame, text="times in 30 s").pack(side=tk.LEFT)

        # USER VOICE PREFERENCES section
        self.user_voices_listbox, row = self.side_list(right_frame, row + 4, "USER VOICES", [
            ("✓ Save", self.save_user_voices), ("✖ Remove", self.remove_user_voice),
            ("+ Add", self.add_user_voice), ("✎ Edit", self.edit_user_voice)], weight=2)

        # Speak section: long-form text (DECtalk singing and long messages) and Quick Speak
        speak_frame = self.speak_frame = ttk.Frame(self.main_paned)
        longform_frame = ttk.LabelFrame(speak_frame, text="Long-form Text Input (DECtalk Singing/Long Messages)",
                                        padding=(px(6), px(2), px(6), px(6)))
        longform_frame.pack(fill=tk.BOTH, expand=True)

        text_input_frame = ttk.Frame(longform_frame)
        text_input_frame.pack(fill=tk.BOTH, expand=True)

        self.longform_text = ui_theme.ScrolledText(text_input_frame, height=3, width=60, wrap=tk.WORD)
        self.longform_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add character counter
        char_count_frame = ttk.Frame(text_input_frame)
        char_count_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(px(6), 0))

        self.char_count_label = ttk.Label(char_count_frame, text="0 chars", style='Hint.TLabel')
        self.char_count_label.pack()

        # Update character count on text change
        def update_char_count(event=None):
            text = self.longform_text.get("1.0", tk.END).strip()
            self.char_count_label.config(text=f"{len(text)} chars")

        self.longform_text.bind("<KeyRelease>", update_char_count)

        # Controls for long-form text
        longform_controls = ttk.Frame(longform_frame)
        longform_controls.pack(fill=tk.X, pady=(px(6), 0))

        # Voice selector for long-form text
        ttk.Label(longform_controls, text="Voice:").pack(side=tk.LEFT, padx=(0, px(4)))
        self.longform_voice_var = tk.StringVar(value="Default")
        self.longform_voice_combo = ttk.Combobox(longform_controls, textvariable=self.longform_voice_var,
                                                 width=30, state='readonly')
        self.longform_voice_combo.pack(side=tk.LEFT, padx=(0, px(6)))

        ttk.Button(longform_controls, text="Speak Long Text", command=self.speak_longform).pack(side=tk.LEFT, padx=(0, px(4)))
        ttk.Button(longform_controls, text="Clear", command=lambda: self.longform_text.delete("1.0", tk.END)).pack(side=tk.LEFT)

        # Quick Speak/Apply controls - between long-form and the tabs
        test_frame = ttk.Frame(speak_frame)
        test_frame.pack(fill=tk.X, pady=(px(6), 0))

        ttk.Button(test_frame, text="✓ Apply", command=self.apply_settings).pack(side=tk.RIGHT)
        ttk.Button(test_frame, text="Speak", command=self.speak_test).pack(side=tk.RIGHT, padx=(0, px(4)))
        ttk.Label(test_frame, text="Quick Speak:").pack(side=tk.LEFT, padx=(0, px(4)))
        self.test_entry = ttk.Entry(test_frame, width=20)
        self.test_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, px(6)))
        self.test_entry.insert(0, "Enter your text here")

        self.main_paned.add(speak_frame, stretch='never', minsize=px(150))

        # Bottom section - Tabs (each scrolls if the window is too small for it)
        self.notebook = ttk.Notebook(self.main_paned)
        self.main_paned.add(self.notebook, stretch='always', minsize=px(170))

        # Create tabs
        self.create_settings_tab()
        self.create_audio_devices_tab()
        self.create_voice_commands_tab()
        self.create_dectalk_tab()
        if self.textaloud and textaloud_installed():  # only on PCs that have TextAloud - nobody else needs it
            self.create_textaloud_tab()
        if VOICES_AVAILABLE:
            self.create_voices_tab()
        self.create_announcements_tab()
        self.create_auto_block_tab()
        self.create_help_tab()
        self.create_testing_tab()

        ui_theme.apply_theme(self.root, self.config.get('ui_theme', 'system'))  # colour everything just built
        self.root.after(300, self.restore_dividers)  # once the window has its final size

    def _on_window_configure(self, event):
        if event.widget is not self.root or (event.width, event.height) == self._content_size:
            return
        if self._resize_job:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(150, self._apply_content_size)

    def _apply_content_size(self):
        """The window stopped changing size: lay everything out once for the new size"""
        self._resize_job = None
        width, height = self.root.winfo_width(), self.root.winfo_height()
        self._content_size = (width, height)
        self.content.place_configure(relwidth=0, relheight=0, width=width, height=height)

    def side_list(self, parent, row, title, buttons, weight=1):
        """One list on the right: title, the list (shares the height with the others, scrolls) and its buttons.
        Returns (listbox, next free grid row)."""
        px = self.px
        ttk.Label(parent, text=title, style='Section.TLabel').grid(row=row, column=0, sticky='w',
                                                                   pady=(px(10) if row else 0, px(3)))
        box = ttk.Frame(parent)
        box.grid(row=row + 1, column=0, sticky='nsew')
        parent.rowconfigure(row + 1, weight=weight, minsize=px(44))
        listbox = tk.Listbox(box, height=3, width=26)
        bar = ttk.Scrollbar(box, orient=tk.VERTICAL, command=listbox.yview)
        listbox.configure(yscrollcommand=bar.set)
        bar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        btns = ttk.Frame(parent)
        btns.grid(row=row + 2, column=0, sticky='w', pady=(px(4), 0))
        for text, command in buttons:
            ttk.Button(btns, text=text, command=command, style='Compact.TButton').pack(side=tk.LEFT, padx=(0, px(3)))
        return listbox, row + 3

    def add_tab(self, title):
        """A notebook tab whose contents scroll when the window is too small for them"""
        holder = ui_theme.ScrollFrame(self.notebook)
        self.notebook.add(holder, text=title)
        return holder.inner

    def restore_window_size(self):
        """Last size and position (if it's still on a screen), else a size that fits this screen"""
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.minsize(min(self.px(1100), sw - 40), min(self.px(600), sh - 80))
        saved = self.config.get('window') or {}
        geometry = saved.get('geometry', '')
        match = re.match(r'^(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$', geometry)
        on_screen = False
        if match:
            w, h, x, y = map(int, match.groups())
            try:  # the whole desktop, across all monitors
                import ctypes
                gm = ctypes.windll.user32.GetSystemMetrics
                vx, vy, vw, vh = gm(76), gm(77), gm(78), gm(79)
            except Exception:
                vx, vy, vw, vh = 0, 0, sw, sh
            on_screen = (w >= 400 and h >= 300 and vx - 50 <= x <= vx + vw - 200 and vy - 10 <= y <= vy + vh - 100)
        if on_screen:
            self.root.geometry(geometry)
        else:
            ui_theme.size_window(self.root, 1260, 1100)
        if saved.get('zoomed'):
            self.root.after_idle(lambda: self.root.state('zoomed'))

    def fit_speak_pane(self):
        """Keep the long-form / Quick Speak section tall enough for its controls - the looks differ in height, so
        after switching look (or with dividers saved under another look) it would otherwise be cut off"""
        try:
            self.root.update_idletasks()
            need = self.speak_frame.winfo_reqheight()
            self.main_paned.paneconfigure(self.speak_frame, minsize=need)
            short = need - self.speak_frame.winfo_height()
            if short > 0:
                self.main_paned.sash_place(1, 0, self.main_paned.sash_coord(1)[1] + short)
            # the list panel on the right: at least as wide as its widest row of buttons
            need = self.right_frame.winfo_reqwidth()
            self.side_paned.paneconfigure(self.right_frame, minsize=need)
            short = need - self.right_frame.winfo_width()
            if short > 0:
                self.side_paned.sash_place(0, self.side_paned.sash_coord(0)[0] - short, 0)
        except (tk.TclError, AttributeError):
            pass

    def restore_dividers(self):
        saved = self.config.get('window') or {}
        try:
            self.root.update_idletasks()
            for i, y in enumerate(saved.get('dividers', [])[:2]):
                self.main_paned.sash_place(i, 0, int(y))
            if saved.get('side_divider'):
                self.side_paned.sash_place(0, int(saved['side_divider']), 0)
        except (tk.TclError, ValueError, TypeError):
            pass
        self.fit_speak_pane()

    def remember_window_size(self):
        """on_closing: keep size, position, maximized and the divider positions for next time"""
        try:
            saved = dict(self.config.get('window') or {})
            zoomed = self.root.state() == 'zoomed'
            saved['zoomed'] = zoomed
            if not zoomed:
                saved['geometry'] = self.root.geometry()
            saved['dividers'] = [self.main_paned.sash_coord(i)[1] for i in range(2)]
            saved['side_divider'] = self.side_paned.sash_coord(0)[0]
            self.config['window'] = saved
        except Exception as e:
            logger.debug(f"Could not remember the window size: {e}")

    def on_theme_changed(self, event=None):
        choice = {'Match Windows': 'system', 'Light': 'light', 'Dark': 'dark',
                  'Classic (original look)': 'classic'}.get(self.theme_var.get(), 'system')
        self.config['ui_theme'] = choice
        self.save_config()
        self.ui_mode = ui_theme.apply_theme(self.root, choice)
        self.root.after(150, self.fit_speak_pane)  # after the new look has settled
        if hasattr(self, 'device_tree'):
            self.device_tree.tag_configure('default', background=ui_theme.colors()['good'])

    def create_settings_tab(self):
        """Settings tab with config import/export"""
        tab = self.add_tab("Settings")
        settings_frame = ttk.Frame(tab)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        settings_frame.columnconfigure(1, weight=1)
        
        # Game selection (games are added/edited/removed with the buttons next to it)
        ttk.Label(settings_frame, text="Game").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        self.game_combo = ttk.Combobox(settings_frame, values=self.game_names(), width=28, state='readonly')
        self.game_combo.grid(row=0, column=1, sticky='ew', padx=10, pady=5)
        self.game_combo.set(self.current_game)
        self.game_combo.bind('<<ComboboxSelected>>', self.on_game_changed)
        game_btn_frame = ttk.Frame(settings_frame)
        game_btn_frame.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Button(game_btn_frame, text="+ Add Game", command=self.add_game).pack(side=tk.LEFT, padx=2)
        ttk.Button(game_btn_frame, text="✎ Edit", command=self.edit_game).pack(side=tk.LEFT, padx=2)
        ttk.Button(game_btn_frame, text="✖ Remove", command=self.remove_game).pack(side=tk.LEFT, padx=2)
        
        # TTS Command Prefix - moved from Voice Commands to Settings
        ttk.Label(settings_frame, text="TTS Command Prefix").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        self.tts_prefix_entry = ttk.Entry(settings_frame, width=14)
        self.tts_prefix_entry.grid(row=1, column=1, sticky='w', padx=10, pady=5)
        self.tts_prefix_entry.insert(0, self.get_current_game_config().get('tts_command_prefix', '!tts'))
        ttk.Label(settings_frame, text="Command users type in game chat (e.g., !tts, !speak). DRG strips first !tts.", wraplength=280, style='Hint.TLabel').grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        
        # Game Log File
        ttk.Label(settings_frame, text="Game Log File").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
        self.log_path_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.log_path_var, width=28).grid(row=2, column=1, columnspan=2, sticky='ew', padx=10, pady=5)
        
        # Admin Username
        ttk.Label(settings_frame, text="Admin Username").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        self.admin_username_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.admin_username_var, width=28).grid(row=3, column=1, sticky='w', padx=10, pady=5)
        
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
        self.default_voice_combo = ttk.Combobox(settings_frame, width=28)
        self.default_voice_combo.grid(row=7, column=1, sticky='ew', padx=10, pady=5)
        
        # DECtalk Voice Volume
        ttk.Label(settings_frame, text="DECtalk Voice Volume").grid(row=8, column=0, sticky=tk.W, padx=10, pady=5)
        self.dectalk_volume_var = tk.DoubleVar(value=self.config.get('dectalk_volume', 0.7))  # Default to 70% for DECtalk
        dectalk_volume_frame = ttk.Frame(settings_frame)
        dectalk_volume_frame.grid(row=8, column=1, sticky=tk.W, padx=10, pady=5)
        
        # Config Import/Export buttons
        config_frame = ttk.LabelFrame(settings_frame, text="Configuration Management")
        config_frame.grid(row=10, column=0, columnspan=3, sticky=tk.W+tk.E, padx=10, pady=10)

        # Appearance: Windows 11 style, light / dark / follow Windows
        ttk.Label(settings_frame, text="Appearance").grid(row=9, column=0, sticky=tk.W, padx=10, pady=5)
        self.theme_var = tk.StringVar(value={'system': 'Match Windows', 'light': 'Light', 'dark': 'Dark',
                                             'classic': 'Classic (original look)'}.get(
            self.config.get('ui_theme', 'system'), 'Match Windows'))
        theme_combo = ttk.Combobox(settings_frame, textvariable=self.theme_var, state='readonly', width=22,
                                   values=['Match Windows', 'Light', 'Dark', 'Classic (original look)'])
        theme_combo.grid(row=9, column=1, sticky=tk.W, padx=10, pady=5)
        theme_combo.bind('<<ComboboxSelected>>', self.on_theme_changed)
        
        ttk.Button(config_frame, text="Export Config", command=self.export_config).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(config_frame, text="Import Config", command=self.import_config).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(config_frame, text="Reset to Defaults", command=self.reset_config).pack(side=tk.LEFT, padx=5, pady=5)
        # Up to 150%: DECtalk's own output peaks around 77% of full scale, so the very loudest bits may clip past ~130%
        self.dectalk_volume_slider = ttk.Scale(dectalk_volume_frame, from_=0.0, to=1.5,
                                               variable=self.dectalk_volume_var, 
                                               orient=tk.HORIZONTAL, length=self.px(240))
        self.dectalk_volume_slider.pack(side=tk.LEFT)
        self.dectalk_volume_label = ttk.Label(dectalk_volume_frame, text=f"{int(self.dectalk_volume_var.get() * 100)}%")
        self.dectalk_volume_label.pack(side=tk.LEFT, padx=10)
        self.dectalk_volume_var.trace('w', self.update_dectalk_volume_label)
        
        # Populate voices
        if self.voices:
            self.refresh_voice_combos()
            
    def create_audio_devices_tab(self):
        """Audio Devices tab with real Windows devices"""
        tab = self.add_tab("Audio Devices")
        
        # Description
        desc_label = ttk.Label(tab, text="Select audio output device for TTS:")
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
        self.device_tree.tag_configure('default', background=ui_theme.colors()['good'])
        
        self.device_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.device_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.device_tree.configure(yscrollcommand=scrollbar.set)
        
        # Control buttons
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Select Device", command=self.select_audio_device).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Refresh Devices", command=self.refresh_audio_devices).pack(side=tk.LEFT, padx=5)
        
        # Current device label
        self.current_device_label = ttk.Label(tab, text="Current Device: Default", style='Section.TLabel')
        self.current_device_label.pack(pady=10)
        
    def create_voice_commands_tab(self):
        """Voice Commands tab - editable"""
        tab = self.add_tab("Voice Commands")
        
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
        # second row, so the buttons fit beside the voice list
        btn_frame2 = ttk.Frame(left_frame)
        btn_frame2.pack(pady=(0, 5))
        ttk.Button(btn_frame2, text="Save Commands", command=self.save_voice_commands).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame2, text="Reset to Defaults", command=self.reset_voice_commands).pack(side=tk.LEFT, padx=2)
        
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
        self.hidden_voices = self.load_hidden_voices()
        
        # Populate voice list
        self.refresh_available_voices()
        
        # Management buttons
        mgmt_frame = ttk.Frame(right_frame)
        mgmt_frame.pack(pady=5)
        
        ttk.Button(mgmt_frame, text="Hide / Unhide Selected", command=self.hide_selected_voices).pack(side=tk.LEFT, padx=2)
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
        tab = self.add_tab("DECtalk")
        
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
        
        ttk.Label(right_frame, text="Example: [:np] for Paul, [:nb] for Betty", style='Hint.TLabel').pack(anchor=tk.W)
        ttk.Label(right_frame, text="Advanced: [:np][:rate 200][:pitch 150]", style='Hint.TLabel').pack(anchor=tk.W)
        
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
        tab = self.add_tab("Auto Block")
        
        ttk.Label(tab, text="Auto-block keywords (one per line):").pack(pady=10)
        
        self.auto_block_text = ui_theme.ScrolledText(tab, height=10, width=60)
        self.auto_block_text.pack(padx=10, pady=10)
        
        ttk.Button(tab, text="Save Keywords", command=self.save_auto_block).pack(pady=5)
    
    def create_textaloud_tab(self):
        """TextAloud tab: voices that only run inside TextAloud (Acapela Will From Afar etc.)"""
        tab = self.add_tab("TextAloud")

        ttk.Label(tab, wraplength=900, justify=tk.LEFT, text=(
            "Voices bought for TextAloud (like Acapela Will From Afar) only work inside TextAloud, so the TTS has "
            "TextAloud say each line into a sound file and plays it like DECtalk: same output device, !stop, mutes "
            "and music ducking. TextAloud runs hidden - no window or tray icon - starting when one of its voices is "
            "needed and closing when the TTS closes. If TextAloud crashes, its error box is closed for you and the "
            "line is tried again, or said in the Default Voice.")).pack(anchor=tk.W, padx=10, pady=10)

        row = ttk.Frame(tab)
        row.pack(fill=tk.X, padx=10, pady=5)
        self.textaloud_enabled_var = tk.BooleanVar(value=self.textaloud_enabled)
        ttk.Checkbutton(row, text="Use TextAloud voices", variable=self.textaloud_enabled_var,
                        command=self.on_textaloud_toggle).pack(side=tk.LEFT)
        ttk.Button(row, text="Find Voices",
                   command=lambda: self.refresh_textaloud_voices(from_button=True)).pack(side=tk.LEFT, padx=10)

        row = ttk.Frame(tab)
        row.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(row, text="Voices:").pack(side=tk.LEFT)
        self.textaloud_status_label = ttk.Label(row, text=self.textaloud_status_text())
        self.textaloud_status_label.pack(side=tk.LEFT, padx=5)

        row = ttk.Frame(tab)
        row.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(row, text="Volume:").pack(side=tk.LEFT)
        self.textaloud_volume_var = tk.DoubleVar(value=self.textaloud_volume)
        ttk.Scale(row, from_=0.0, to=1.5, variable=self.textaloud_volume_var,
                  orient=tk.HORIZONTAL, length=300).pack(side=tk.LEFT, padx=5)
        self.textaloud_volume_label = ttk.Label(row, text=f"{int(self.textaloud_volume * 100)}%")
        self.textaloud_volume_label.pack(side=tk.LEFT)
        self.textaloud_volume_var.trace('w', self.update_textaloud_volume_label)

        row = ttk.Frame(tab)
        row.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(row, text="Close TextAloud after").pack(side=tk.LEFT)
        self.textaloud_idle_var = tk.StringVar(value=f"{float(self.config.get('textaloud_idle_minutes', 10)):g}")
        idle_box = ttk.Spinbox(row, from_=0, to=120, width=5, textvariable=self.textaloud_idle_var,
                               command=self.save_textaloud_idle)
        idle_box.pack(side=tk.LEFT, padx=5)
        idle_box.bind('<FocusOut>', lambda e: self.save_textaloud_idle())
        idle_box.bind('<Return>', lambda e: self.save_textaloud_idle())
        ttk.Label(row, text="minutes without use (0 = keep it open; it starts again in ~2 s when needed)").pack(side=tk.LEFT)

        ttk.Label(tab, wraplength=900, justify=tk.LEFT, text=(
            "Give players these voices like any other: add a voice command for it (Voice Commands tab), set it as "
            "someone's User Voice, the Default Voice, or the DJ voice.")).pack(anchor=tk.W, padx=10, pady=(15, 5))

    def save_textaloud_idle(self):
        try:
            minutes = max(0.0, min(120.0, float(self.textaloud_idle_var.get())))
        except ValueError:
            minutes = float(self.config.get('textaloud_idle_minutes', 10))
        self.textaloud_idle_var.set(f"{minutes:g}")
        if self.config.get('textaloud_idle_minutes') != minutes:
            self.config['textaloud_idle_minutes'] = minutes
            self.save_config()
        if self.textaloud:
            self.textaloud.idle_minutes = minutes

    # ------------------------------------------------------------------ Voices tab (more Windows voices, Piper)
    def create_voices_tab(self):
        """Voices tab: Windows' extra voices, and free Piper voices to download"""
        tab = self.add_tab("Voices")
        px = self.px

        win = ttk.LabelFrame(tab, text="Windows voices", padding=px(8))
        win.pack(fill=tk.X, padx=10, pady=(10, 5))
        ttk.Label(win, wraplength=px(760), justify=tk.LEFT, style='Hint.TLabel', text=(
            "Windows can add more voices and languages (Settings > Time & language > Speech > Add voices). "
            "Windows keeps those apart from the voices programs like the TTS can use, so after adding some, press "
            "\"Make them usable here\" (one Windows admin prompt) and restart the TTS.")).pack(anchor='w')
        row = ttk.Frame(win)
        row.pack(fill=tk.X, pady=(px(6), 0))
        ttk.Button(row, text="Add voices in Windows Settings...",
                   command=voice_installer.open_windows_voice_settings).pack(side=tk.LEFT)
        ttk.Button(row, text="Make them usable here", command=self.make_windows_voices_usable).pack(side=tk.LEFT, padx=px(6))
        ttk.Button(row, text="Undo", command=self.undo_windows_voices).pack(side=tk.LEFT)
        self.windows_voice_label = ttk.Label(win, wraplength=px(760), justify=tk.LEFT)
        self.windows_voice_label.pack(anchor='w', pady=(px(6), 0))

        sf = ttk.LabelFrame(tab, text="Classic SAPI 4 voices (Microsoft Sam 4.0, TruVoice / BonziBuddy...)", padding=px(8))
        sf.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(sf, wraplength=px(760), justify=tk.LEFT, style='Hint.TLabel', text=(
            "Old SAPI 4 voices installed on this PC show up in every voice list as [SAPI4] ..., with a ready-made "
            "[SAPI4] Bonzi (TruVoice Adult Male #2 at BonziBuddy's pitch and speed). They need the SAPI 4 runtime "
            "(spchapi.exe) and the voice packages installed - none are included.")).pack(anchor='w')
        row = ttk.Frame(sf)
        row.pack(fill=tk.X, pady=(px(6), 0))
        self.sapi4_status = ttk.Label(row, text="Looking for SAPI 4 voices..." if self.sapi4 and self.sapi4.available()
                                      else "SAPI 4 isn't installed on this PC.")
        self.sapi4_status.pack(side=tk.LEFT)
        row = ttk.Frame(sf)
        row.pack(fill=tk.X, pady=(px(6), 0))
        self.sapi4_test_var = tk.StringVar()
        self.sapi4_test_combo = ttk.Combobox(row, textvariable=self.sapi4_test_var, state='readonly', width=34)
        self.sapi4_test_combo.pack(side=tk.LEFT)
        ttk.Button(row, text="Test", command=self.test_sapi4_voice).pack(side=tk.LEFT, padx=px(6))
        ttk.Label(row, text="Volume:").pack(side=tk.LEFT, padx=(px(16), px(4)))
        self.sapi4_volume_var = tk.DoubleVar(value=self.sapi4_volume)
        ttk.Scale(row, from_=0.0, to=1.5, variable=self.sapi4_volume_var, orient=tk.HORIZONTAL,
                  length=px(160)).pack(side=tk.LEFT)
        self.sapi4_volume_label = ttk.Label(row, text=f"{int(self.sapi4_volume * 100)}%")
        self.sapi4_volume_label.pack(side=tk.LEFT, padx=px(4))
        self.sapi4_volume_var.trace('w', self.update_sapi4_volume_label)

        pf = ttk.LabelFrame(tab, text="Free voices (Piper)", padding=px(8))
        pf.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        ttk.Label(pf, wraplength=px(760), justify=tk.LEFT, style='Hint.TLabel', text=(
            "Natural-sounding voices that run on this PC, free (Piper, MIT licence; each voice has its own licence "
            "in its MODEL_CARD). Each voice is a one-time download of about 20-130 MB. They play like DECtalk: same "
            "output device, !stop and mutes.")).pack(anchor='w')
        row = ttk.Frame(pf)
        row.pack(fill=tk.X, pady=(px(6), px(4)))
        self.piper_engine_label = ttk.Label(row, text="")
        self.piper_engine_label.pack(side=tk.LEFT)
        self.piper_engine_btn = ttk.Button(row, text="Download Piper (22 MB)", command=self.download_piper_engine)
        self.piper_engine_btn.pack(side=tk.LEFT, padx=px(6))
        row = ttk.Frame(pf)  # language filter on its own line so the row above stays narrow
        row.pack(fill=tk.X, pady=(0, px(4)))
        ttk.Label(row, text="Language:").pack(side=tk.LEFT)
        self.piper_lang_var = tk.StringVar(value=self.config.get('piper_language', 'English'))
        self.piper_lang_combo = ttk.Combobox(row, textvariable=self.piper_lang_var, state='readonly', width=24)
        self.piper_lang_combo.pack(side=tk.LEFT, padx=px(6))
        self.piper_lang_combo.bind('<<ComboboxSelected>>', lambda e: self.fill_piper_list())
        ttk.Button(row, text="Refresh list", command=lambda: self.load_piper_catalog(refresh=True)).pack(side=tk.LEFT)

        tree_frame = ttk.Frame(pf)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        columns = ('voice', 'language', 'quality', 'size', 'status')
        self.piper_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=6, selectmode='browse')
        for col, title, width in (('voice', 'Voice', 170), ('language', 'Language', 220), ('quality', 'Quality', 80),
                                  ('size', 'Size', 80), ('status', '', 110)):
            self.piper_tree.heading(col, text=title)
            self.piper_tree.column(col, width=px(width), stretch=(col in ('voice', 'language')))
        bar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.piper_tree.yview)
        self.piper_tree.configure(yscrollcommand=bar.set)
        bar.pack(side=tk.RIGHT, fill=tk.Y)
        self.piper_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.piper_tree.bind('<Double-1>', lambda e: self.download_piper_voice())

        row = ttk.Frame(pf)
        row.pack(fill=tk.X, pady=(px(6), 0))
        self.piper_download_btn = ttk.Button(row, text="Download", command=self.download_piper_voice)
        self.piper_download_btn.pack(side=tk.LEFT)
        ttk.Button(row, text="Remove", command=self.remove_piper_voice).pack(side=tk.LEFT, padx=px(6))
        ttk.Button(row, text="Test", command=self.test_piper_voice).pack(side=tk.LEFT)
        ttk.Label(row, text="Volume:").pack(side=tk.LEFT, padx=(px(16), px(4)))
        self.piper_volume_var = tk.DoubleVar(value=self.piper_volume)
        ttk.Scale(row, from_=0.0, to=1.5, variable=self.piper_volume_var, orient=tk.HORIZONTAL,
                  length=px(160)).pack(side=tk.LEFT)
        self.piper_volume_label = ttk.Label(row, text=f"{int(self.piper_volume * 100)}%")
        self.piper_volume_label.pack(side=tk.LEFT, padx=px(4))
        self.piper_volume_var.trace('w', self.update_piper_volume_label)
        self.piper_progress = ttk.Label(pf, text="", style='Hint.TLabel')
        self.piper_progress.pack(anchor='w', pady=(px(4), 0))

        self.piper_catalog = {}
        self.piper_busy = False
        self.refresh_windows_voice_label()
        self.refresh_piper_engine_label()
        self.load_piper_catalog(refresh=False)

    # SAPI 4 voices
    def on_sapi4_voices(self):
        names = list(self.sapi4.voices) if self.sapi4 else []
        if hasattr(self, 'sapi4_status'):
            truvoice = sum(1 for n in names if n.endswith('(TruVoice)'))
            self.sapi4_status.config(text=(f"{len(names)} SAPI 4 voices found ({truvoice} TruVoice)"
                                           + (" - [SAPI4] Bonzi is ready" if 'Bonzi' in names else "")) if names
                                      else "SAPI 4 is installed, but no SAPI 4 voices were found.")
            self.sapi4_test_combo['values'] = names
            if names and not self.sapi4_test_var.get():
                self.sapi4_test_var.set(names[0])
        self.refresh_all_voice_lists()

    def test_sapi4_voice(self):
        name = self.sapi4_test_var.get()
        if not (self.sapi4 and name in self.sapi4.voices):
            return

        def work():
            from temp_utils import get_temp_file, release_temp_file
            wav = get_temp_file(suffix='.wav', prefix='tts_sapi4_')
            try:
                line = "Hi there! I'm Bonzi Buddy." if name == 'Bonzi' else f"Hello, this is {name.split(' (')[0]}."
                if self.sapi4.render(line, name, wav):
                    self.dectalk_native._play_wav(wav, self.config.get('audio_device', 'Default'), self.sapi4_volume)
            finally:
                release_temp_file(wav)
        threading.Thread(target=work, daemon=True).start()

    def update_sapi4_volume_label(self, *args):
        volume = self.sapi4_volume_var.get()
        self.sapi4_volume_label.config(text=f"{int(volume * 100)}%")
        self.sapi4_volume = volume
        self.config['sapi4_volume'] = volume
        if getattr(self, '_volume_save_job', None):
            self.root.after_cancel(self._volume_save_job)
        self._volume_save_job = self.root.after(500, self.save_config)

    # Windows voices
    def refresh_windows_voice_label(self):
        status = voice_installer.windows_voice_status()
        usable = [d.split(' - ')[0].replace('Microsoft ', '') for t, d, st in status if st == 'usable']
        new = [d.split(' - ')[0].replace('Microsoft ', '') for t, d, st in status if st == 'available']
        text = f"Added here: {', '.join(usable) or 'none'}     Ready to add: {', '.join(new) or 'none'}"
        self.windows_voice_label.config(text=text)

    def make_windows_voices_usable(self):
        new = [(t, d) for t, d, st in voice_installer.windows_voice_status() if st == 'available']
        if not new:
            messagebox.showinfo("Windows voices", "There are no new Windows voices to add.\n\nAdd some first with "
                                "\"Add voices in Windows Settings\", then come back here.", parent=self.root)
            return
        names = "\n".join(d for t, d in new)
        if not messagebox.askyesno("Windows voices", f"Make these usable in the TTS?\n\n{names}\n\nWindows will ask "
                                   "for admin permission. An undo file is saved first.", parent=self.root):
            return

        def work():
            ok, message = voice_installer.make_windows_voices_usable([t for t, d in new], self.voices_dir)
            self.run_on_gui(lambda: (self.refresh_windows_voice_label(),
                                     (messagebox.showinfo if ok else messagebox.showwarning)(
                                         "Windows voices", message, parent=self.root)))
        threading.Thread(target=work, daemon=True).start()

    def undo_windows_voices(self):
        if not messagebox.askyesno("Windows voices", "Remove the Windows voices this TTS made usable?\n\n"
                                   "Windows will ask for admin permission.", parent=self.root):
            return

        def work():
            ok, message = voice_installer.undo_windows_voices(self.voices_dir)
            self.run_on_gui(lambda: (self.refresh_windows_voice_label(),
                                     (messagebox.showinfo if ok else messagebox.showwarning)(
                                         "Windows voices", message, parent=self.root)))
        threading.Thread(target=work, daemon=True).start()

    # Piper voices
    def refresh_piper_engine_label(self):
        if self.piper.available():
            self.piper_engine_label.config(text="Piper is installed.")
            self.piper_engine_btn.pack_forget()
        else:
            self.piper_engine_label.config(text="Piper (the program that speaks these voices) isn't downloaded yet.")

    def load_piper_catalog(self, refresh=False):
        """The voice list: from the saved copy, or (refresh / first time) from Piper's site in the background"""
        def work():
            try:
                catalog = voice_installer.load_catalog(self.voices_dir, refresh=refresh)
                self.run_on_gui(lambda: self.on_piper_catalog(catalog))
            except Exception as e:
                logger.warning(f"Could not get Piper's voice list: {e}")
                self.run_on_gui(lambda: self.piper_progress.config(
                    text="Couldn't get the voice list - check the internet connection and press Refresh list"))
        cached = (Path(self.voices_dir) / 'piper_catalog.json').exists()
        if refresh or not cached:
            self.piper_progress.config(text="Getting the voice list...")
            threading.Thread(target=work, daemon=True).start()
        else:
            work()

    def on_piper_catalog(self, catalog):
        self.piper_catalog = catalog
        families = sorted({v['language']['name_english'] for v in catalog.values()})
        self.piper_lang_combo['values'] = ['All languages', 'Downloaded'] + families
        if self.piper_lang_var.get() not in self.piper_lang_combo['values']:
            self.piper_lang_var.set('English' if 'English' in families else 'All languages')
        self.piper_progress.config(text=f"{len(catalog)} voices in {len(families)} languages")
        self.fill_piper_list()

    def fill_piper_list(self):
        self.config['piper_language'] = self.piper_lang_var.get()
        installed = set(voice_installer.installed_piper_voices(self.voices_dir))
        want = self.piper_lang_var.get()
        self.piper_tree.delete(*self.piper_tree.get_children())
        for key, v in sorted(self.piper_catalog.items(),
                             key=lambda kv: (kv[0] not in installed, kv[1]['language']['name_english'], kv[0])):
            lang = v['language']
            if want == 'Downloaded' and key not in installed:
                continue
            if want not in ('All languages', 'Downloaded') and lang['name_english'] != want:
                continue
            size = voice_installer.voice_files(v)[2] / 1e6
            self.piper_tree.insert('', tk.END, iid=key, values=(
                key.split('-')[1] if key.count('-') >= 2 else key,
                f"{lang['name_english']} ({lang['country_english']})", v['quality'], f"{size:.0f} MB",
                "downloaded" if key in installed else ""))

    def selected_piper_key(self):
        sel = self.piper_tree.selection()
        if not sel:
            messagebox.showinfo("Voices", "Pick a voice in the list first.", parent=self.root)
            return None
        return sel[0]

    def piper_progress_cb(self, label):
        last = {'t': 0}

        def progress(done, total):
            if time.time() - last['t'] < 0.2 and done < total:
                return
            last['t'] = time.time()
            pct = f" {done * 100 // total}%" if total else ""
            self.run_on_gui(lambda: self.piper_progress.config(text=f"{label}{pct} ({done / 1e6:.0f} MB)"))
        return progress

    def download_piper_engine(self, then=None):
        if self.piper_busy:
            return
        self.piper_busy = True

        def work():
            try:
                voice_installer.install_piper_engine(self.voices_dir, progress=self.piper_progress_cb("Downloading Piper"))
                self.run_on_gui(lambda: (self.piper_progress.config(text="Piper downloaded."),
                                         self.refresh_piper_engine_label()))
                if then:
                    self.run_on_gui(then)
            except Exception as e:
                logger.error(f"Piper download failed: {e}")
                self.run_on_gui(lambda: self.piper_progress.config(text=f"Piper download failed: {e}"))
            finally:
                self.piper_busy = False
        threading.Thread(target=work, daemon=True).start()

    def download_piper_voice(self):
        key = self.selected_piper_key()
        if not key or self.piper_busy:
            return
        if key in voice_installer.installed_piper_voices(self.voices_dir):
            self.piper_progress.config(text=f"{key} is already downloaded - press Test to hear it")
            return
        if not self.piper.available():
            self.download_piper_engine(then=self.download_piper_voice)  # the engine first, then this voice
            return
        entry = self.piper_catalog[key]
        self.piper_busy = True

        def work():
            try:
                voice_installer.install_piper_voice(self.voices_dir, entry,
                                                    progress=self.piper_progress_cb(f"Downloading {key}"))
                self.run_on_gui(lambda: (self.piper_progress.config(
                    text=f"{key} downloaded - it's in the voice lists as {PIPER_PREFIX}{key}"),
                    self.fill_piper_list(), self.refresh_textaloud_lists()))
            except Exception as e:
                logger.error(f"Voice download failed: {e}")
                self.run_on_gui(lambda: self.piper_progress.config(text=f"Download failed: {e}"))
            finally:
                self.piper_busy = False
        threading.Thread(target=work, daemon=True).start()

    def remove_piper_voice(self):
        key = self.selected_piper_key()
        if not key or key not in voice_installer.installed_piper_voices(self.voices_dir):
            return
        if not messagebox.askyesno("Voices", f"Delete the voice {key}?\n\nPlayers who use it get the default "
                                   "voice until you download it again.", parent=self.root):
            return
        self.piper._kill(key)
        voice_installer.remove_piper_voice(self.voices_dir, key)
        self.piper_progress.config(text=f"{key} removed")
        self.fill_piper_list()
        self.refresh_textaloud_lists()

    def test_piper_voice(self):
        key = self.selected_piper_key()
        if not key:
            return
        if key not in voice_installer.installed_piper_voices(self.voices_dir) or not self.piper.available():
            self.piper_progress.config(text="Download the voice first")
            return

        def work():
            from temp_utils import get_temp_file, release_temp_file
            wav = get_temp_file(suffix='.wav', prefix='tts_piper_')
            try:
                name = key.split('-')[1] if key.count('-') >= 2 else key
                if self.piper.render(f"Hello, this is {name}.", key, wav):
                    self.dectalk_native._play_wav(wav, self.config.get('audio_device', 'Default'), self.piper_volume)
            finally:
                release_temp_file(wav)
        threading.Thread(target=work, daemon=True).start()

    def update_piper_volume_label(self, *args):
        volume = self.piper_volume_var.get()
        self.piper_volume_label.config(text=f"{int(volume * 100)}%")
        self.piper_volume = volume
        self.config['piper_volume'] = volume
        if getattr(self, '_volume_save_job', None):
            self.root.after_cancel(self._volume_save_job)
        self._volume_save_job = self.root.after(500, self.save_config)

    def create_help_tab(self):
        """Help tab with instructions and information"""
        tab = self.add_tab("Help")
        
        # Create scrolled text for help content
        help_text = ui_theme.ScrolledText(tab, wrap=tk.WORD, width=80, height=20)
        help_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        prefix = self.get_current_game_config().get('tts_command_prefix', '!tts')
        toggle = self.voice_toggle_command or '/vt'
        help_content = f"""TTS FOR TEAM FORTRESS 2 - HELP
==============================

GETTING STARTED
1. Settings tab: pick the game (or + Add Game) and check its log file path.
2. Audio Devices tab: pick where the voices play (e.g. "VoiceMeeter Input" to send them into voice chat).
3. Add yourself under ADMINS on the right and press Save.
4. Press Start TTS.
In the game's console the log must be on: con_logfile log.txt (the TTS reads it; nothing is sent to the game).

WHAT PLAYERS TYPE
  {prefix} hello there	speak "hello there" in their voice
  {prefix} /v 3 hello	say it in voice 3 this time (the numbers are on the Voice Commands tab)
  {prefix} /sam hello	any custom trigger you made on the Voice Commands tab
  {prefix} {toggle} 3 hello	make voice 3 their voice from now on (the text is optional)
  {prefix} [:np] hello	DECtalk codes always use DECtalk for that line, even if their voice is another one
New players get a random voice if Random Voice is on (Settings), unless you set one under USER VOICES.

ADMIN COMMANDS (only people in the ADMINS list)
  !stop	stop all speech and clear the queue
  !block	block whoever is talking now (or whoever used {prefix} last)
  !block clear	unblock the most recently blocked player
  !admin add	make whoever used {prefix} last an admin
  !mute	mute whoever is talking for the time set under MUTED (!mute 30 = 30 minutes)
  !unmute	lift the most recent mute
Lists on the right: select a name and use the buttons; press Save after adding or removing.
Spam guard: the same message repeated (or a burst of messages) mutes that player automatically.

MUSIC - the DJ button
  !sr song name or YouTube link	request a song (off / admins / everyone - DJ window, Options)
  !skipsong	skip the current song (admin)
  !approve / !deny	handle the oldest request waiting for your OK (admin, if approval is on)
DJ window keys: Space play/pause, Ctrl+Right next, Ctrl+Left previous, Ctrl+S stop, Delete remove,
Ctrl+Up / Ctrl+Down move a song, Ctrl+O add files, Ctrl+F search.
The music gets quieter while a voice talks. "Download ffmpeg" in the DJ window if it says ffmpeg is missing.

SOUNDS - the Sounds button
Put sound files in the "sounds" folder next to the TTS; the file name is the word (yee.wav -> yee).
Depending on the mode players type the word, or the command and the word (!s yee). Adding a number
changes the pitch (yee 150). Cooldowns and who may use them are in the Sounds window.

SILENT KEYS (tts.cfg in TF2's cfg folder, exec it from autoexec.cfg)
The keys write a line into the log as if you had typed it, without anything appearing in game chat:
  Numpad /	!stop
  Numpad -	!block
  Numpad +	!admin add
  Numpad *	!block clear
  Numpad .	!mute
  Numpad 7	!approve
  Numpad 8	!skipsong
  Numpad 9	!deny
  Numpad Enter	voice chat on/off
Change the name in tts.cfg (echo YourName : !stop) to your own Steam name.

OTHER
  Voices tab: Windows' extra voices, free Piper voices, and old SAPI 4 voices (Sam 4.0, TruVoice -
    [SAPI4] Bonzi) if SAPI 4 is installed.
  Long-form box: longer text or DECtalk songs, spoken in the voice you pick next to it.
  Quick Speak: try a line out loud. Apply saves the Settings tab.
  Announcements tab: what the TTS says when someone is blocked, muted, requests a song...
  Auto Block tab: words that block a player automatically.
  Appearance (Settings): Match Windows, Light or Dark. Drag the dividers to resize the sections."""

        help_text.configure(tabs=(self.px(260),))  # the columns below are tab-separated
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
        # Log path, admins and blocked list belong to the current game (top-level keys are TF2's legacy copy)
        game_config = self.get_current_game_config()
        log_path = game_config.get('log_path', '')
        if not log_path and self.current_game_key() == 'tf2':
            log_path = self.config.get('log_path', '')  # TF2's legacy top-level copy, never another game's
        self.log_path_var.set(log_path)
        self.admin_username_var.set(self.config.get('admin_username', ''))
        self.private_mode_var.set(self.config.get('private_mode', False))
        
        # Populate long-form voice combo now that GUI is ready
        if hasattr(self, 'longform_voice_combo') and self.audio_manager:
            self.populate_longform_voice_combo()
        
        # Set default voice from config - the dropdown too, or the next Apply saves whatever it happens to show
        default_voice = self.config.get('default_voice', '')
        if default_voice and hasattr(self, 'default_voice_combo'):
            self.default_voice_combo.set(default_voice)
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
        
        for admin in game_config.get('admins', self.config.get('admins', [])):
            self.admin_listbox.insert(tk.END, admin)

        for blocked in game_config.get('blocked', self.config.get('blocked', [])):
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
            logger.info(f"ON_CHAT_MESSAGE: Received !block message from {msg.get('username', 'unknown')}")
        
        # Format and display
        timestamp = time.strftime("%H:%M:%S")
        prefix = ""
        
        if msg['is_dead']:
            prefix += "*DEAD*"
        if msg.get('is_spec'):
            prefix += "*SPEC* "
        if msg['is_team']:
            prefix += "(TEAM) "
            
        display_text = f"{prefix}{msg['username']} : {msg['message']}\n"
        
        # Add to chat display
        self.run_on_gui(lambda: self.chat_display.insert(tk.END, display_text))
        self.run_on_gui(lambda: self.chat_display.see(tk.END))
        
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
                    self.run_on_gui(lambda u=msg['username']: self.update_ui_after_autoblock(u))
                    return
        
        # THIRD: Check for admin commands FIRST (before !tts processing)
        message = msg['message'].strip()
        admins = [self.admin_listbox.get(i) for i in range(self.admin_listbox.size())]
        
        # NOTE: "!tts stop" is intentionally NOT a command. It falls through to
        # normal TTS so the word "stop" can actually be spoken. Use !stop to stop.

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
                # Determine who to block - prioritize currently speaking user, then last speaker. Never an
                # internal speaker: during an announcement or long-form text that would silence those for good.
                user_to_block = self.currently_speaking_user
                if not user_to_block or user_to_block.startswith(('__', '[')):
                    user_to_block = self.last_speaker

                if user_to_block and user_to_block != msg['username'] and not user_to_block.startswith(('__', '[')):
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

        # !mute [minutes]: mute whoever is talking (or last used TTS) for a while - default length from the
        # MUTED section. !unmute: lift the most recent mute.
        mute_match = re.match(r'^!mute(?:\s+(\d+))?$', message.lower().strip())
        if mute_match:
            if msg['username'] in admins:
                # Whoever is talking - unless that's an announcement/long-form, then the last TTS user
                user_to_mute = self.currently_speaking_user
                if not user_to_mute or user_to_mute.startswith(('__', '[')):
                    user_to_mute = getattr(self, 'last_speaker', None)
                if user_to_mute and user_to_mute != msg['username'] and not user_to_mute.startswith(('__', '[')):
                    minutes = int(mute_match.group(1)) if mute_match.group(1) else None
                    self.mute_user(user_to_mute, minutes)
                else:
                    logger.info("No valid user to mute")
            else:
                logger.info(f"Non-admin {msg['username']} tried to use !mute")
            return

        if message.lower().strip() == '!unmute':
            if msg['username'] in admins:
                with self._mute_lock:
                    latest = next(reversed(self.muted_users), None) if self.muted_users else None
                if latest:
                    self.unmute_user(latest)
                else:
                    logger.info("Nobody is muted")
            return

        # DJ admin commands: !skipsong skips the current song; !approve / !deny handle the oldest request
        # waiting for the DJ's OK (also from a silent key). Above the muted check, like the other admin commands.
        lowered = message.lower()
        if self.dj and lowered == '!skipsong':
            if msg['username'] in admins:
                self.dj.skip()
            return
        if self.dj and lowered in ('!approve', '!deny'):
            if msg['username'] in admins:
                if not self.dj.pending:
                    logger.info("No song requests are waiting")
                elif lowered == '!approve':
                    self.dj.approve_request(0)
                else:
                    self.dj.deny_request(0)
            return

        # Temporarily muted (!mute, spam guard, or the MUTED list): nothing from here down - TTS, song
        # requests, soundboard. It comes after the admin commands so a muted admin can still !stop / !unmute.
        if self.is_muted(msg['username']):
            logger.info(f"Muted user {msg['username']} tried to speak")
            return

        # DJ: "!sr <YouTube link or song name>" requests a song (when requests are on in the DJ window).
        # Blocked and muted players never get here.
        if self.dj and (lowered == '!sr' or lowered.startswith('!sr ')):
            username = msg['username']
            is_admin = username in admins
            if self.private_mode_var.get() and not is_admin:
                return
            if not self.dj.request_allowed(is_admin):
                logger.info(f"Song request from {username} ignored (requests are off for them)")
                return
            # !sr spam counts toward the spam guard like !tts does
            if not is_admin and self.check_spam(username, message):
                logger.info(f"Spam guard: muting {username} (song requests)")
                self.mute_user(username, announcement='SPAM MUTED')
                return
            def reply(text, outcome=None, title=None, reason=None):
                """The DJ's answer: a line in the TTS window, and spoken so the player knows (if the TTS is on)"""
                self.run_on_gui(lambda: (self.chat_display.insert(tk.END, f"[DJ] {text}\n"),
                                         self.chat_display.see(tk.END)))
                if not self.is_running:
                    return
                if outcome == 'denied':
                    # automatic refusals: at most one spoken per player every 30 s, so it can't be used to spam
                    now = time.time()
                    last = self._last_denied_announce.get(username, 0)
                    if now - last < 30:
                        return
                    self._last_denied_announce[username] = now
                    self.queue_announcement_next('REQUEST DENIED', username=username, fill={'reason': reason},
                                                 voice=self.dj_voice())
                elif outcome == 'dj_denied':  # the DJ said no - always told
                    self.queue_announcement_next('REQUEST DENIED', username=username, fill={'reason': reason},
                                                 voice=self.dj_voice())
                elif outcome in ('accepted', 'pending', 'approved'):
                    kind = {'accepted': 'SONG REQUESTED', 'pending': 'REQUEST PENDING',
                            'approved': 'REQUEST APPROVED'}[outcome]
                    self.queue_announcement_next(kind, username=username, fill={'title': clean_title(title or '')},
                                                 voice=self.dj_voice())
            self.dj.request(username, message[3:], is_admin, reply)
            return

        # Soundboard: "yee" / "yee 150" (open mode) or "!s yee" (command mode) plays right away - no queue.
        # Cooldowns and who-can-use are handled inside; a trigger is never spoken as TTS.
        if self.soundboard and self.soundboard.handle_chat(msg['username'], message, msg['username'] in admins,
                                                           self.private_mode_var.get()):
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
            logger.info(f"Processing !tts command from {game_type} by {msg['username']} ({len(text_to_speak)} chars)")
            
            # Check private mode - only admins can use !tts in private mode
            if self.private_mode_var.get():
                admins = [self.admin_listbox.get(i) for i in range(self.admin_listbox.size())]
                if msg['username'] not in admins:
                    logger.info(f"Non-admin {msg['username']} tried to use TTS in private mode")
                    return

            # Spam guard: the same message over and over, or a burst of pastes -> temporary mute
            if msg['username'] not in admins and self.check_spam(msg['username'], text_to_speak):
                logger.info(f"Spam guard: muting {msg['username']}")
                self.mute_user(msg['username'], announcement='SPAM MUTED')
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
        # (uses the module-level re; a local "import re" here would break the earlier re.match for !mute)

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
                
                logger.info(f"Voice command (with space): '{cmd}' ({len(text)} chars)")
                
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
                
                logger.info(f"Voice command trigger: '{trigger}' ({len(text)} chars)")
                
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
            # DECtalk profiles and TextAloud voices are handled in speech processing
            if is_special_voice(voice_name):
                logger.info(f"Special voice selected: {voice_name}")
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
            logger.info(f"Adding message to speech queue ({len(text)} chars) with voice '{voice_name}' from {username}")
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
                if f"[DECtalk] {profile_name}" not in self.hidden_voices:
                    voice_list.append(f"[DECtalk] {profile_name}")
        voice_list += [v for v in self.extra_voice_names() if v not in self.hidden_voices]

        # Update combo box
        self.longform_voice_combo['values'] = voice_list
        logger.info(f"Populated long-form voice combo with {len(voice_list)} voices")
    
    def speak_longform(self):
        """Speak the long-form text box with the chosen voice.
        (The old "Output to TF2 Chat" option was removed: it appended say lines to log.txt, which TF2 only
        writes and never reads, so nothing ever reached the game.)"""
        text = self.longform_text.get("1.0", tk.END).strip()
        if not text:
            logger.warning("No text entered in long-form input")
            return

        # Get selected voice
        selected_voice = self.longform_voice_var.get()
        if selected_voice == "Default":
            selected_voice = None

        # Speak the full text with selected voice
        logger.info(f"Speaking long-form text ({len(text)} chars) with voice: {selected_voice}")
        self.speak(text, voice_name=selected_voice, username="[Long-form Input]")
        
        # Clear the text area after speaking
        # Don't clear longform text - keep it for reuse
        # self.longform_text.delete("1.0", tk.END)
            
    def process_speech_queue(self):
        """Process messages from the queue sequentially"""
        import time
        
        # A restart (Reload, Edit Game, Stop+Start) can begin a new speech thread while this one is still in a
        # long DECtalk message; when it finishes it must exit rather than keep speaking alongside the new one
        while self.queue_running and self.queue_thread is threading.current_thread():
            try:
                # Wait for a message (timeout to check running status)
                msg = self.message_queue.get(timeout=0.5)
                
                # Check if this user is already blocked before speaking
                blocked_users = [self.blocked_listbox.get(i) for i in range(self.blocked_listbox.size())]
                if msg.get('username') in blocked_users:
                    logger.info(f"Skipping message from blocked user: {msg.get('username')}")
                    continue
                if msg.get('username') and self.is_muted(msg.get('username')):
                    logger.info(f"Skipping message from muted user: {msg.get('username')}")
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
                    
                    # If no voice specified in message, check if user has a DECtalk/TextAloud preference
                    if not current_voice:
                        username = msg.get('username')
                        if username and username in self.user_voice_preferences:
                            user_voice = self.user_voice_preferences[username]
                            if is_special_voice(user_voice):
                                current_voice = user_voice
                                logger.info(f"Using {username}'s preference: {current_voice}")
                        elif is_special_voice(str(self.config.get('default_voice', ''))):
                            # No voice of their own and the Default Voice is DECtalk/TextAloud (apply_default_voice
                            # only knows Windows voices, so without this they'd get whichever Windows voice was last)
                            current_voice = self.config['default_voice']

                    # DECtalk commands sent with a non-DECtalk voice: use DECtalk for this one message only
                    # (their saved voice is untouched) so SAPI doesn't read "[:phoneme on]..." as gibberish
                    if not (current_voice and current_voice.startswith("[DECtalk] ")) and DECTALK_MARKUP.search(text_to_speak):
                        auto_voice = self.get_dectalk_auto_voice(msg.get('username'))
                        if auto_voice:
                            logger.info(f"DECtalk markup from {msg.get('username')} on a non-DECtalk voice ({current_voice}) - using {auto_voice} for this message")
                            current_voice = auto_voice

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
                            logger.info(f"Sending {len(text_to_speak)} chars to DECtalk")
                            dectalk_volume = self.dectalk_volume
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
                            self.audio_manager.speak(text_to_speak, use_sync=False, tag=msg.get('username'))
                    elif current_voice and current_voice.startswith(SAPI4_PREFIX):
                        logger.info(f"Using SAPI 4 voice: {current_voice}")
                        if self.speak_sapi4(text_to_speak, current_voice, msg.get('username')):
                            used_special_voice = True
                        else:
                            logger.warning("SAPI 4 voice unavailable, using the default voice")
                            self.apply_default_voice()
                            self.audio_manager.speak(text_to_speak, use_sync=False, tag=msg.get('username'))
                    elif current_voice and current_voice.startswith(PIPER_PREFIX):
                        logger.info(f"Using Piper voice: {current_voice}")
                        if self.speak_piper(text_to_speak, current_voice, msg.get('username')):
                            used_special_voice = True
                        else:
                            logger.warning("Piper voice unavailable, using the default voice")
                            self.apply_default_voice()
                            self.audio_manager.speak(text_to_speak, use_sync=False, tag=msg.get('username'))
                    elif current_voice and current_voice.startswith(TEXTALOUD_PREFIX):
                        logger.info(f"Using TextAloud voice: {current_voice}")
                        if self.speak_textaloud(text_to_speak, current_voice, msg.get('username')):
                            used_special_voice = True
                        else:
                            # TextAloud missing or broken: say it in the Default Voice rather than lose the line
                            logger.warning("TextAloud voice unavailable, using the default voice")
                            self.apply_default_voice()
                            self.audio_manager.speak(text_to_speak, use_sync=False, tag=msg.get('username'))
                    else:
                        # Normal SAPI5 speech
                        logger.info(f"Using SAPI5 speech - Voice: {current_voice}")
                        self.audio_manager.speak(text_to_speak, use_sync=False, tag=msg.get('username'))
                    
                    # Wait for speech to complete or be interrupted (only for SAPI5)
                    # Windows voices normally don't hold up the queue, so DECtalk and Windows voices can talk over
                    # each other (the way it has always behaved). is_speaking() used to be broken and always
                    # said False; now that it works, waiting is opt-in: config "sapi_wait_for_turn": true.
                    if not used_special_voice and self.config.get('sapi_wait_for_turn', False):
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
            
    def get_dectalk_auto_voice(self, username=None):
        """DECtalk voice for one-off messages containing DECtalk markup, or None if switched off/unavailable.
        Uses the user's own saved DECtalk voice if they have one (e.g. they picked a SAPI voice with /v for
        this message), otherwise config "dectalk_auto_voice" (profile name, default Perfect Paul).
        Config "dectalk_auto_switch": false turns this off."""
        if not self.config.get('dectalk_auto_switch', True) or not self.dectalk_enabled:
            return None
        if not (self.dectalk_native and self.dectalk_native.is_available()):
            return None
        saved = self.user_voice_preferences.get(username, '') if username else ''
        if saved.startswith("[DECtalk] "):
            return saved
        profiles = self.dectalk_profiles or {}
        wanted = self.config.get('dectalk_auto_voice', 'Perfect Paul').replace("[DECtalk] ", "")
        if wanted in profiles:
            return f"[DECtalk] {wanted}"
        if profiles:
            return f"[DECtalk] {next(iter(profiles))}"
        return None

    # TextAloud voices (Will From Afar etc.)
    def extra_voice_names(self):
        """TextAloud and Piper voices for the voice lists"""
        names = self.textaloud_voice_names()
        if getattr(self, 'piper', None):
            names += self.piper.voice_names()
        if getattr(self, 'sapi4', None):
            names += self.sapi4.voice_names()
        return names

    def textaloud_voice_names(self):
        """'[TextAloud] ...' entries for the voice lists - empty when switched off or TextAloud isn't installed"""
        if not (self.textaloud and self.textaloud_enabled and textaloud_installed()):
            return []
        return [TEXTALOUD_PREFIX + v for v in self.config.get('textaloud_voices', [])]

    def textaloud_status_text(self):
        if not (self.textaloud and textaloud_installed()):
            return "TextAloud 3 isn't installed"
        if not self.textaloud_enabled:
            return "Off"
        voices = self.config.get('textaloud_voices', [])
        return ", ".join(voices) if voices else "No extra voices found yet - click Find Voices"

    def refresh_textaloud_voices(self, from_button=False):
        """Ask TextAloud which voices it has - in the background, since starting TextAloud takes a few seconds"""
        if not (self.textaloud and textaloud_installed()):
            if from_button:
                messagebox.showinfo("TextAloud", "TextAloud 3 isn't installed on this PC.")
            return
        if hasattr(self, 'textaloud_status_label'):
            self.textaloud_status_label.config(text="Asking TextAloud...")
        windows = [v.name for v in (self.voices or [])]

        def work():
            try:
                voices = self.textaloud.list_voices(windows)
                # Don't leave it running just for the list - it starts again when one of its voices speaks
                self.textaloud.shutdown()
            except Exception as e:
                logger.error(f"Could not get TextAloud's voices: {e}")
                voices = None
            self.run_on_gui(lambda: self.on_textaloud_voices(voices))
        threading.Thread(target=work, daemon=True).start()

    def on_textaloud_voices(self, voices):
        if voices is None:
            if hasattr(self, 'textaloud_status_label'):
                self.textaloud_status_label.config(text="Couldn't reach TextAloud - see tts_tf2.log")
            return
        self.config['textaloud_voices'] = voices
        self.save_config()
        self.refresh_textaloud_lists()

    def refresh_textaloud_lists(self):
        if hasattr(self, 'textaloud_status_label'):
            self.textaloud_status_label.config(text=self.textaloud_status_text())
        self.refresh_voice_combos()
        if hasattr(self, 'longform_voice_combo'):
            self.populate_longform_voice_combo()
        self.refresh_available_voices()

    def on_textaloud_toggle(self):
        self.textaloud_enabled = self.textaloud_enabled_var.get()
        self.config['textaloud_enabled'] = self.textaloud_enabled
        self.save_config()
        if self.textaloud_enabled and not self.config.get('textaloud_voices'):
            self.refresh_textaloud_voices(from_button=True)
        else:
            self.refresh_textaloud_lists()

    def update_textaloud_volume_label(self, *args):
        volume = self.textaloud_volume_var.get()
        self.textaloud_volume_label.config(text=f"{int(volume * 100)}%")
        self.textaloud_volume = volume
        self.config['textaloud_volume'] = volume
        if getattr(self, '_volume_save_job', None):
            self.root.after_cancel(self._volume_save_job)
        self._volume_save_job = self.root.after(500, self.save_config)

    def speak_sapi4(self, text, voice_name, username):
        """Speak one message in an old SAPI 4 voice: the 32-bit helper writes a WAV, which plays like DECtalk"""
        from temp_utils import get_temp_file, release_temp_file
        player = getattr(self, 'dectalk_native', None)
        if not (self.sapi4 and hasattr(player, '_play_wav')):
            return False
        wav_path = get_temp_file(suffix='.wav', prefix='tts_sapi4_')
        player.stop_requested = False

        def stopped():
            return player.stop_requested or self.currently_speaking_user != username
        try:
            if not self.sapi4.render(text, voice_name[len(SAPI4_PREFIX):], wav_path, should_stop=stopped):
                return stopped()  # cut off by !stop or a mute: not a failure
            if not stopped():
                player._play_wav(wav_path, self.config.get('audio_device', 'Default'), self.sapi4_volume)
            return True
        except Exception as e:
            logger.error(f"SAPI 4 voice '{voice_name}' failed: {e}")
            return False
        finally:
            release_temp_file(wav_path)

    def speak_piper(self, text, voice_name, username):
        """Speak one message in a Piper voice: Piper writes a WAV, which plays like DECtalk (same output device,
        !stop, mutes and music ducking). False if Piper couldn't do it."""
        from temp_utils import get_temp_file, release_temp_file
        player = getattr(self, 'dectalk_native', None)
        if not (self.piper and self.piper.available() and hasattr(player, '_play_wav')):
            return False
        wav_path = get_temp_file(suffix='.wav', prefix='tts_piper_')
        player.stop_requested = False

        def stopped():
            return player.stop_requested or self.currently_speaking_user != username
        try:
            if not self.piper.render(text, voice_name[len(PIPER_PREFIX):], wav_path, should_stop=stopped):
                return stopped()  # cut off by !stop or a mute: not a failure
            if not stopped():
                player._play_wav(wav_path, self.config.get('audio_device', 'Default'), self.piper_volume)
            return True
        except Exception as e:
            logger.error(f"Piper voice '{voice_name}' failed: {e}")
            return False
        finally:
            release_temp_file(wav_path)

    def speak_textaloud(self, text, voice_name, username):
        """Speak one message in a TextAloud voice: TextAloud writes it to a WAV, which plays like DECtalk (same
        output device, !stop, mutes and music ducking). False if TextAloud couldn't do it."""
        from temp_utils import get_temp_file, release_temp_file
        player = getattr(self, 'dectalk_native', None)
        if not (self.textaloud and hasattr(player, '_play_wav')):
            return False
        voice = voice_name[len(TEXTALOUD_PREFIX):]
        wav_path = get_temp_file(suffix='.wav', prefix='tts_textaloud_')
        player.stop_requested = False  # set again by !stop (stop_all_speech) or a mute while this one is going

        def stopped():
            return player.stop_requested or self.currently_speaking_user != username
        try:
            if not self.textaloud.render(text, voice, wav_path, should_stop=stopped):
                return stopped()  # cut off during the ~1 s render: not a failure, so no fallback voice
            if not stopped():
                player._play_wav(wav_path, self.config.get('audio_device', 'Default'), self.textaloud_volume)
            return True
        except Exception as e:
            logger.error(f"TextAloud voice '{voice}' failed: {e}")
            return False
        finally:
            release_temp_file(wav_path)

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
            
        # Cut them off right away - only them, everyone else keeps talking. Windows-voice messages are handed
        # to the speech worker straight away, so theirs may be playing or waiting there:
        if self.audio_manager and hasattr(self.audio_manager, 'stop_user'):
            try:
                self.audio_manager.stop_user(username)
            except Exception as e:
                logger.error(f"Could not stop {username}'s speech: {e}")
        # DECtalk plays inside the speech thread, so while it's theirs currently_speaking_user is them
        if self.currently_speaking_user == username:
            dectalk = getattr(self, 'dectalk_native', None)
            if dectalk and hasattr(dectalk, 'is_playing') and dectalk.is_playing():
                dectalk.stop_speech()
            # Clear the flag to signal the queue processor to stop waiting
            self.currently_speaking_user = None
            logger.info(f"Stopped current speech from {username}")
        
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

    def dj_voice(self):
        """Voice for everything the DJ says (now playing, song-request messages): the DJ window's choice,
        or the Announcements tab voice when that's set to Default"""
        voice = self.config.get('dj', {}).get('announce_voice', 'Default')
        if not voice or voice == 'Default':
            voice = self.config.get('announcement_voice', '')
        return voice

    def queue_announcement_next(self, announcement_type, username=None, fill=None, voice=None):
        """Put an announcement at the front of the speech queue, like the block announcements.
        fill: values for {placeholders} in the announcement text, e.g. {'title': ..., 'reason': ...}
        voice: who says it (default: the Announcements tab voice)"""
        announcement_text = self.get_announcement_text(announcement_type, username=username)
        if not announcement_text:
            return
        for key, value in (fill or {}).items():
            # brackets out: a song title can't smuggle DECtalk commands into the announcement
            safe = str(value or '').replace('[', ' ').replace(']', ' ').replace('<', ' ').replace('>', ' ')
            announcement_text = announcement_text.replace('{' + key + '}', safe)
        waiting = []
        while not self.message_queue.empty():
            try:
                waiting.append(self.message_queue.get_nowait())
            except queue.Empty:
                break
        announcement_voice = voice or self.config.get('announcement_voice', '')
        self.message_queue.put({'text': announcement_text, 'voice': announcement_voice, 'username': '__announcement__'})
        for item in waiting:
            self.message_queue.put(item)

    def is_muted(self, username):
        """True while a temporary mute is running (a mute that has run out is dropped)"""
        with self._mute_lock:
            until = self.muted_users.get(username)
            if until is None:
                return False
            if until > time.time():
                return True
            del self.muted_users[username]
            self.config['muted_users'] = dict(self.muted_users)
        self.save_config()
        return False

    def mute_user(self, username, minutes=None, announcement='MUTED'):
        """Mute a player for a while: drop their queued messages, cut them off if they're talking, announce it.
        announcement is 'MUTED' (admin/manual) or 'SPAM MUTED' (spam guard)."""
        if not username or username.startswith('__'):
            return
        if minutes is None:
            minutes = self.config.get('mute_minutes', 10)
        with self._mute_lock:
            self.muted_users.pop(username, None)  # re-insert so the newest mute is last (for !unmute)
            self.muted_users[username] = time.time() + minutes * 60
            self.spam_history.pop(username, None)
            self.config['muted_users'] = dict(self.muted_users)
        self.save_config()
        self.clear_user_from_queue(username, add_to_blocked=False)
        logger.info(f"Muted {username} for {minutes} min ({announcement})")
        if self.is_running:  # otherwise it would sit in the queue and play whenever TTS is next started
            self.queue_announcement_next(announcement, username=username)
        reason = " (spam)" if announcement == 'SPAM MUTED' else ""
        self.run_on_gui(lambda: (self.chat_display.insert(tk.END, f"[SYSTEM] Muted {username} for {minutes} min{reason}\n"),
                                 self.chat_display.see(tk.END)))
        self.run_on_gui(self.refresh_muted_list)

    def unmute_user(self, username):
        """Lift a mute early. Returns True if they were muted."""
        with self._mute_lock:
            removed = self.muted_users.pop(username, None) is not None
            self.config['muted_users'] = dict(self.muted_users)
        if removed:
            self.save_config()
            logger.info(f"Unmuted {username}")
            self.run_on_gui(self.refresh_muted_list)
        return removed

    def check_spam(self, username, text):
        """Record a TTS message and return True if the player is spamming: the same message (loosely
        matched) spam_repeat_count times within spam_repeat_window seconds, or spam_flood_count messages
        of any kind within spam_flood_window seconds. Defaults: 3 in 30s, 5 in 10s."""
        if not self.config.get('spam_mute_enabled', True):
            return False
        repeat_count = self.config.get('spam_repeat_count', 3)
        repeat_window = self.config.get('spam_repeat_window', 30)
        flood_count = self.config.get('spam_flood_count', 5)
        flood_window = self.config.get('spam_flood_window', 10)
        now = time.time()
        signature = spam_signature(text)
        with self._mute_lock:
            keep = max(repeat_window, flood_window)
            history = [(t, s) for t, s in self.spam_history.get(username, []) if now - t <= keep]
            history.append((now, signature))
            self.spam_history[username] = history
            repeats = sum(1 for t, s in history if s == signature and now - t <= repeat_window)
            burst = sum(1 for t, s in history if now - t <= flood_window)
        return repeats >= repeat_count or burst >= flood_count

    def refresh_muted_list(self):
        """Redraw the MUTED list with minutes left (GUI thread only)"""
        if not hasattr(self, 'muted_listbox'):
            return
        now = time.time()
        with self._mute_lock:
            active = [(u, t) for u, t in self.muted_users.items() if t > now]
        self.muted_listbox.delete(0, tk.END)
        self.muted_listbox_names = []
        for username, until in active:
            minutes_left = max(1, int((until - now + 59) // 60))
            self.muted_listbox.insert(tk.END, f"{username} ({minutes_left} min left)")
            self.muted_listbox_names.append(username)

    def tick_mutes(self):
        """Every 30s: let expired mutes go and update the minutes shown"""
        for username in list(self.muted_users):
            self.is_muted(username)  # drops it if it has run out
        self.refresh_muted_list()
        self.root.after(30000, self.tick_mutes)

    def unmute_selected(self):
        """✖ Unmute button"""
        sel = self.muted_listbox.curselection()
        if sel and sel[0] < len(self.muted_listbox_names):
            self.unmute_user(self.muted_listbox_names[sel[0]])

    def add_mute(self):
        """+ Mute button: mute a player by name"""
        from tkinter import simpledialog
        name = simpledialog.askstring("Mute User", "Enter username:")
        if not name:
            return
        minutes = simpledialog.askinteger("Mute User", "Mute for how many minutes?",
                                          initialvalue=self.config.get('mute_minutes', 10), minvalue=1, maxvalue=1440)
        if minutes:
            self.mute_user(name, minutes)

    def save_mute_settings(self):
        """Auto-mute spam checkbox and the default mute length"""
        self.config['spam_mute_enabled'] = self.spam_mute_var.get()
        try:
            minutes = int(self.mute_minutes_var.get())
            if minutes >= 1:
                self.config['mute_minutes'] = minutes
        except ValueError:
            self.mute_minutes_var.set(str(self.config.get('mute_minutes', 10)))
        if hasattr(self, 'spam_repeat_var'):
            try:
                self.config['spam_repeat_count'] = max(2, min(10, int(self.spam_repeat_var.get())))
            except ValueError:
                self.spam_repeat_var.set(str(self.config.get('spam_repeat_count', 3)))
        self.save_config()

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
            text = announcements.get(announcement_type, DEFAULT_ANNOUNCEMENTS.get(announcement_type, ""))
        
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
            # A DECtalk default voice isn't a Windows voice, so it has to ride along with the message
            self.speak(text, voice_name=default_voice if is_special_voice(default_voice) else None)

    
    def update_ui_after_autoblock(self, username):
        """Update UI after auto-blocking a user (called on main thread)"""
        # Show in chat
        self.chat_display.insert(tk.END, f"[SYSTEM] Auto-blocked user: {username}\n")
        self.chat_display.see(tk.END)
        logger.info(f"UI updated for auto-blocked user: {username}")
            
    def force_stop_tts(self):
        """Force stop all TTS.

        Delegates to stop_all_speech() so "!tts stop" behaves identically to
        "!stop": it skips the current utterance, stops the native DECtalk
        engine, AND clears the pending message queue. Previously this only
        called audio_manager.stop(), leaving the queue and DECtalk in an
        inconsistent state.
        """
        self.stop_all_speech()
            
    def stop_all_speech(self):
        """Stop all current speech immediately (and any soundboard sounds)"""
        if getattr(self, 'soundboard', None):
            self.soundboard.stop_all()
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
        
        game_key = self.current_game_key()
        
        # Stop existing monitor for this game if any
        if game_key in self.monitors:
            self.monitors[game_key].stop()
        
        # Create appropriate monitor: any Source game's console log reads like TF2's
        game_type = self.game_configs.get(game_key, {}).get('type', 'source')
        if game_type == 'source':
            self.monitors[game_key] = TF2LogMonitor(log_path)
        elif game_type == 'drg' and DRG_MONITOR_AVAILABLE:
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
        self.apply_config()
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
        game_key = self.current_game_key()
        
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
        game_key = self.current_game_key()
        
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
            voice_names = [v.name for v in self.voices if v.name and v.name not in self.hidden_voices]
            
            # Add DECtalk voices if enabled
            if self.dectalk_enabled and self.dectalk_profiles:
                for profile_name in self.dectalk_profiles.keys():
                    if f"[DECtalk] {profile_name}" not in self.hidden_voices:
                        voice_names.append(f"[DECtalk] {profile_name}")
            voice_names += [v for v in self.extra_voice_names() if v not in self.hidden_voices]
            
            from tkinter import messagebox
            # Create custom dialog for voice selection
            dialog = tk.Toplevel(self.root)
            dialog.title("Select Voice")
            ui_theme.size_window(dialog, 400, 300, self.root)
            
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
            voice_names = [v.name for v in self.voices if v.name and v.name not in self.hidden_voices]
            
            # Add DECtalk voices if enabled
            if self.dectalk_enabled and self.dectalk_profiles:
                for profile_name in self.dectalk_profiles.keys():
                    if f"[DECtalk] {profile_name}" not in self.hidden_voices:
                        voice_names.append(f"[DECtalk] {profile_name}")
            voice_names += [v for v in self.extra_voice_names() if v not in self.hidden_voices]
            
            dialog = tk.Toplevel(self.root)
            dialog.title("Edit Voice")
            ui_theme.size_window(dialog, 400, 300, self.root)
            
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
        # Save to config once the slider stops moving, not on every tick
        self.dectalk_volume = volume
        self.config['dectalk_volume'] = volume
        if getattr(self, '_volume_save_job', None):
            self.root.after_cancel(self._volume_save_job)
        self._volume_save_job = self.root.after(500, self.save_config)
    
    def refresh_voice_commands_display(self):
        """After DECtalk profiles change: refresh the voice pickers and the Voice Commands list.
        (The profile Add/Update/Remove/Reset/Import buttons called this, but it was never written, so they
        stopped with an error before their success message.)"""
        self.refresh_voice_combos()
        if hasattr(self, 'voice_tree'):
            for item in self.voice_tree.get_children():
                self.voice_tree.delete(item)
            for cmd, voice in self.voice_commands.items():
                self.voice_tree.insert('', tk.END, values=(cmd, voice), tags=(cmd,))

    def update_user_voices_list(self):
        """Redraw the USER VOICES list from user_voice_preferences (was called but missing)"""
        if hasattr(self, 'user_voices_listbox'):
            self.user_voices_listbox.delete(0, tk.END)
            for username, voice in self.user_voice_preferences.items():
                self.user_voices_listbox.insert(tk.END, f"{username}: {voice}")

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
                if f"[DECtalk] {profile}" not in self.hidden_voices:
                    voice_names.append(f"[DECtalk] {profile}")
        voice_names += [v for v in self.extra_voice_names() if v not in self.hidden_voices]
        
        
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
        # Keep the native DECtalk engine's profile table in sync with the
        # user-edited profiles. Without this, the engine uses its own
        # hardcoded table and ignores any profiles added/edited in the GUI.
        if getattr(self, 'dectalk_native', None) is not None:
            self.dectalk_native.voice_profiles = self.dectalk_profiles
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
            # Test with native DECtalk - same volume slider and output device as real chat messages
            success = self.dectalk_native.speak(test_text, name, use_wav=True,
                                                device_override=self.config.get('audio_device', 'Default'),
                                                volume=self.dectalk_volume)
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
    
    HIDDEN_MARK = " (hidden)"

    def load_hidden_voices(self):
        """The hidden-voices list from config. DECtalk voices used to be saved without their "[DECtalk] " prefix,
        so hiding them never worked - those entries are fixed up here."""
        hidden = list(self.config.get('hidden_voices', []))
        profiles = getattr(self, 'dectalk_profiles', None) or self.config.get('dectalk_profiles') or {}
        windows = {v.name for v in (getattr(self, 'voices', None) or [])}
        fixed = []
        for name in hidden:
            if name in profiles and name not in windows and not name.startswith('['):
                name = f"[DECtalk] {name}"
            if name not in fixed:
                fixed.append(name)
        if fixed != hidden:
            self.config['hidden_voices'] = fixed
        return fixed

    def refresh_all_voice_lists(self):
        self.refresh_available_voices()
        self.refresh_voice_combos()
        if hasattr(self, 'longform_voice_combo'):
            self.populate_longform_voice_combo()

    def refresh_available_voices(self):
        """Every voice, hidden ones marked "(hidden)" - hidden voices are left out of all the voice pickers and
        are never given to new players at random. Voice commands that use one still work."""
        if not hasattr(self, 'available_voices_listbox'):
            return

        self.available_voices_listbox.delete(0, tk.END)
        hidden_colour = ui_theme.colors()['hint']

        def add(text, name):
            hidden = name in self.hidden_voices
            self.available_voices_listbox.insert(tk.END, text + (self.HIDDEN_MARK if hidden else ""))
            if hidden:
                self.available_voices_listbox.itemconfigure(tk.END, foreground=hidden_colour)

        # Track seen voices for duplicate detection
        seen_voices = {}
        for i, voice in enumerate(self.voices):
            display_name = voice.name
            if voice.name in seen_voices:
                display_name = f"{voice.name} (duplicate #{seen_voices[voice.name] + 1})"
                seen_voices[voice.name] += 1
            else:
                seen_voices[voice.name] = 1
            add(f"[{i}] {display_name}", voice.name)
        if self.dectalk_enabled and self.dectalk_profiles:
            for profile_name in self.dectalk_profiles.keys():
                add(f"[DECtalk] {profile_name}", f"[DECtalk] {profile_name}")
        for name in self.extra_voice_names():
            add(name, name)

    def voice_name_from_row(self, text):
        """'[3] Microsoft Sam (hidden)' -> 'Microsoft Sam'; '[DECtalk] Betty' stays '[DECtalk] Betty'"""
        if text.endswith(self.HIDDEN_MARK):
            text = text[:-len(self.HIDDEN_MARK)]
        if ' (duplicate #' in text:
            text = text.split(' (duplicate #')[0]
        return re.sub(r'^\[\d+\] ', '', text)

    def hide_selected_voices(self):
        """Hide the selected voices - or, if they're all hidden already, show them again"""
        selected = self.available_voices_listbox.curselection()
        if not selected:
            messagebox.showinfo("Voices", "Select one or more voices in the list first.", parent=self.root)
            return
        names = [self.voice_name_from_row(self.available_voices_listbox.get(i)) for i in selected]
        unhide = all(n in self.hidden_voices for n in names)
        for name in names:
            if unhide:
                self.hidden_voices.remove(name)
            elif name not in self.hidden_voices:
                self.hidden_voices.append(name)
        self.config['hidden_voices'] = self.hidden_voices
        self.save_config()
        self.refresh_all_voice_lists()
        messagebox.showinfo("Voices", f"{'Showing' if unhide else 'Hid'} {len(names)} voice(s).", parent=self.root)

    def show_all_voices(self):
        """Show all voices"""
        self.hidden_voices = []
        self.config['hidden_voices'] = []
        self.save_config()
        self.refresh_all_voice_lists()
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
            ui_theme.size_window(dialog, 500, 400, self.root)
            
            text = ui_theme.ScrolledText(dialog, wrap=tk.WORD)
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
        ui_theme.size_window(dialog, 450, 400, self.root)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Command trigger input
        ttk.Label(dialog, text="Command Trigger (e.g., 'c', 'sam', 'v10'):").pack(pady=10)
        trigger_var = tk.StringVar()
        trigger_entry = ttk.Entry(dialog, textvariable=trigger_var, width=30)
        trigger_entry.pack(pady=5)
        
        # Preview label
        preview_label = ttk.Label(dialog, text="Usage: !tts /[trigger] message", style='Hint.TLabel')
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
                if dectalk_label in self.hidden_voices:
                    continue
                voice_names.append(dectalk_label)
                voice_listbox.insert(tk.END, dectalk_label)
        textaloud_names = [v for v in self.extra_voice_names() if v not in self.hidden_voices]
        if textaloud_names:
            voice_listbox.insert(tk.END, "")  # Separator
            voice_listbox.insert(tk.END, "--- More Voices ---")
            for name in textaloud_names:
                voice_names.append(name)
                voice_listbox.insert(tk.END, name)
        
        
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
            if selected_text == "" or selected_text.startswith("---"):
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
        ui_theme.size_window(dialog, 450, 400, self.root)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Command trigger input
        ttk.Label(dialog, text="Command Trigger:").pack(pady=10)
        trigger_var = tk.StringVar(value=old_cmd)
        trigger_entry = ttk.Entry(dialog, textvariable=trigger_var, width=30)
        trigger_entry.pack(pady=5)
        
        # Preview label
        preview_label = ttk.Label(dialog, text=f"Usage: !tts /{old_cmd} message", style='Hint.TLabel')
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
                if dectalk_label in self.hidden_voices:
                    continue
                voice_names.append(dectalk_label)
                voice_listbox.insert(tk.END, dectalk_label)
        textaloud_names = [v for v in self.extra_voice_names() if v not in self.hidden_voices]
        if textaloud_names:
            voice_listbox.insert(tk.END, "")  # Separator
            voice_listbox.insert(tk.END, "--- More Voices ---")
            for name in textaloud_names:
                voice_names.append(name)
                voice_listbox.insert(tk.END, name)
        
        
        # Select current voice - by listbox row (the list has divider rows that voice_names doesn't, so an
        # index into voice_names would land two rows off for DECtalk voices)
        rows = list(voice_listbox.get(0, tk.END))
        if old_voice in rows:
            voice_listbox.selection_set(rows.index(old_voice))
            voice_listbox.see(rows.index(old_voice))
        elif voice_names:
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
            if selected_text == "" or selected_text.startswith("---"):
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
            ui_theme.size_window(dialog, 400, 300, self.root)
            
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
                    if dectalk_label in self.hidden_voices:
                        continue
                    voice_names.append(dectalk_label)
                    voice_listbox.insert(tk.END, dectalk_label)
            textaloud_names = [v for v in self.extra_voice_names() if v not in self.hidden_voices]
            if textaloud_names:
                voice_listbox.insert(tk.END, "")  # Separator
                voice_listbox.insert(tk.END, "--- More Voices ---")
                for name in textaloud_names:
                    voice_names.append(name)
                    voice_listbox.insert(tk.END, name)
                
            # Select current voice - by listbox row (the divider rows mean voice_names indexes don't line up)
            rows = list(voice_listbox.get(0, tk.END))
            if current_voice in rows:
                voice_listbox.selection_set(rows.index(current_voice))
                voice_listbox.see(rows.index(current_voice))

            def apply_voice():
                sel = voice_listbox.curselection()
                if sel:
                    # Read the row's text: indexing voice_names picked the voice two rows further down
                    new_voice = voice_listbox.get(sel[0])
                    if not new_voice or new_voice.startswith('---'):
                        return  # a divider row, not a voice
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
        tab = self.add_tab("Announcements")
        
        # Container
        container = ttk.Frame(tab)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title and description
        ttk.Label(container, text="AUTOMATED ANNOUNCEMENTS", style='Title.TLabel').pack(pady=5)
        ttk.Label(container, text="Configure announcement messages and voices").pack(pady=5)
        
        # Announcement voice selector
        voice_frame = ttk.Frame(container)
        voice_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(voice_frame, text="Announcement Voice:").pack(side=tk.LEFT, padx=5)
        self.announcement_voice_combo = ttk.Combobox(voice_frame, width=40)
        self.announcement_voice_combo.pack(side=tk.LEFT, padx=5)
        
        # Populate with voices
        if self.voices:
            voice_names = [v.name for v in self.voices if v.name not in self.hidden_voices]
            self.announcement_voice_combo['values'] = voice_names
            # The saved choice, else Microsoft Sam, else the first voice (showing anything else here would get
            # saved over your choice by the next Save Announcements)
            saved_voice = self.config.get('announcement_voice', '')
            if saved_voice:
                self.announcement_voice_combo.set(saved_voice)
            else:
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
            ("TTS STOPPED", "TTS stopped by admin"),
            ("MUTED", DEFAULT_ANNOUNCEMENTS['MUTED']),
            ("SPAM MUTED", DEFAULT_ANNOUNCEMENTS['SPAM MUTED']),
            ("SONG REQUESTED", DEFAULT_ANNOUNCEMENTS['SONG REQUESTED']),
            ("REQUEST DENIED", DEFAULT_ANNOUNCEMENTS['REQUEST DENIED']),
            ("REQUEST PENDING", DEFAULT_ANNOUNCEMENTS['REQUEST PENDING']),
            ("REQUEST APPROVED", DEFAULT_ANNOUNCEMENTS['REQUEST APPROVED'])
        ]
        
        self.announcement_vars = {}
        self.announcement_enabled_vars = {}
        self.announcement_defaults = {}  # for apply_config to refill the fields after Import/Reset/Reload

        # Headers
        ttk.Label(announcements_frame, text="On", style='Section.TLabel').grid(row=0, column=0, padx=5, pady=3)
        ttk.Label(announcements_frame, text="Type", style='Section.TLabel').grid(row=0, column=1, sticky=tk.W, padx=10, pady=3)
        ttk.Label(announcements_frame, text="Message", style='Section.TLabel').grid(row=0, column=2, padx=10, pady=3)
        
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
            self.announcement_defaults[key] = default
            
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
        tab = self.add_tab("Testing")
        
        # Main container
        main_frame = ttk.Frame(tab)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(main_frame, text="Test Message Simulator", style='Title.TLabel')
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
        
        self.test_output = ui_theme.ScrolledText(output_frame, height=8, width=60)
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
        game_key = self.current_game_key()
        
        # Save game-specific settings
        if 'games' not in self.config:
            self.config['games'] = {}
        if game_key not in self.config['games']:
            self.config['games'][game_key] = {}
        
        self.config['games'][game_key]['log_path'] = self.log_path_var.get()

        # TTS command prefix (per game) - Apply saves everything on this tab
        if hasattr(self, 'tts_prefix_entry'):
            tts_prefix = self.tts_prefix_entry.get().strip()
            if tts_prefix:
                self.config['games'][game_key]['tts_command_prefix'] = tts_prefix

        # Save global settings
        self.config['admin_username'] = self.admin_username_var.get()
        self.config['private_mode'] = self.private_mode_var.get()
        self.config['auto_block'] = self.auto_block_var.get()
        self.config['auto_block_enabled'] = self.auto_block_var.get()  # what load_settings reads first
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
        if getattr(self, 'dj', None):
            self.dj.shutdown()  # stops the music and saves the queue
        if getattr(self, 'soundboard', None):
            self.soundboard.shutdown()
        if self.audio_manager:
            self.audio_manager.stop()
        if getattr(self, 'textaloud', None):
            self.textaloud.shutdown()  # only if the TTS started it
        if getattr(self, 'piper', None):
            self.piper.shutdown()
        if getattr(self, 'sapi4', None):
            self.sapi4.shutdown()
        self.remember_window_size()
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