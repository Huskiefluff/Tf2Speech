"""
Native DECtalk Integration Module
Uses actual DECtalk say.exe for authentic voice synthesis
"""

import os
import sys
import subprocess
from temp_utils import get_temp_file, cleanup_old_temp_files
import threading
import queue
import logging
import time
from pathlib import Path
import wave

try:
    import pyaudio
    PYAUDIO_AVAILABLE = True
except ImportError:
    PYAUDIO_AVAILABLE = False
    print("Warning: PyAudio not available for DECtalk audio playback")

logger = logging.getLogger(__name__)

# CENTRALIZED VOICE DATA PATH
VOICE_DATA_BASE = Path(r"\\HUSKIEFILES\Programs\Dev\tts\tts_tf2\tts voice data dectalk-acapela")

class DECtalkNative:
    """Native DECtalk integration using say.exe"""
    
    def __init__(self):
        # Find DECtalk binaries
        self.dectalk_path = self._find_dectalk_path()
        self.available = self.dectalk_path is not None
        
        if self.available:
            logger.info(f"DECtalk found at: {self.dectalk_path}")
        else:
            logger.warning("DECtalk binaries not found")
        
        # Audio output
        self.pyaudio = None
        self.audio_device_index = None
        self.pyaudio_available = PYAUDIO_AVAILABLE
        
        # Track current processes for stopping
        self.current_process = None
        self.current_stream = None
        self.stop_requested = False  # Flag to stop playback gracefully
        
        # Voice profiles - DECtalk codes
        self.voice_profiles = {
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
            "DECtalk Sings": "[:np][:phone on]"
        }
        
    def _find_dectalk_path(self):
        """Find DECtalk say.exe executable - using vs6 version"""
        # Handle PyInstaller bundle
        if getattr(sys, 'frozen', False):
            # Running in a PyInstaller bundle
            if hasattr(sys, '_MEIPASS'):
                # Onefile mode - use _MEIPASS directory
                bundle_dir = Path(sys._MEIPASS)
                possible_paths = [
                    bundle_dir / "dectalk" / "vs6" / "say.exe",
                    # Legacy paths for compatibility
                    bundle_dir / "dectalk_bin" / "say.exe",
                    bundle_dir / "say.exe",
                ]
            else:
                # Onedir mode
                exe_dir = Path(sys.executable).parent
                possible_paths = [
                    exe_dir / "_internal" / "dectalk" / "vs6" / "say.exe",
                    exe_dir / "dectalk" / "vs6" / "say.exe",
                ]
        else:
            # Running in normal Python environment
            possible_paths = [
                # Centralized vs6 location (CORRECT PATH)
                Path(__file__).parent.parent / "voice_data" / "dectalk" / "vs6" / "say.exe",
                # Legacy paths
                Path(__file__).parent / "dectalk_bin" / "say.exe",
                Path(__file__).parent / "say.exe",
            ]
        
        # Add common paths
        possible_paths.extend([
            # In system PATH
            "say.exe",
            # Common installation paths
            Path("C:/Program Files/DECtalk/say.exe"),
            Path("C:/Program Files (x86)/DECtalk/say.exe"),
            Path("C:/DECtalk/say.exe"),
        ])
        
        # Log all paths being checked
        logger.info("Searching for DECtalk say.exe in the following locations:")
        for path in possible_paths:
            logger.info(f"  Checking: {path}")
        
        for path in possible_paths:
            if isinstance(path, Path):
                if path.exists():
                    return str(path)
            else:
                # Try to run it to see if it's in PATH
                try:
                    result = subprocess.run([path, "/?"], 
                                          capture_output=True, 
                                          timeout=1)
                    if result.returncode == 0:
                        return path
                except:
                    continue
        
        return None
    
    def is_available(self):
        """Check if DECtalk is available"""
        return self.available
    
    def set_audio_device(self, device_name):
        """Set the audio output device for WAV playback"""
        if not self.pyaudio_available:
            logger.warning("PyAudio not available, cannot set audio device")
            return False
        if not self.pyaudio:
            self.pyaudio = pyaudio.PyAudio()
        
        # Find device by name
        for i in range(self.pyaudio.get_device_count()):
            info = self.pyaudio.get_device_info_by_index(i)
            if device_name.lower() in info['name'].lower():
                self.audio_device_index = i
                logger.info(f"DECtalk audio device set to: {info['name']}")
                return True
        
        logger.warning(f"Audio device not found: {device_name}")
        return False
    
    def speak(self, text, voice_profile=None, use_wav=True, device_override=None, volume=1.0):
        """
        Speak text using DECtalk
        
        Args:
            text: Text to speak (can contain DECtalk commands like [:phoneme on])
            voice_profile: DECtalk voice profile name or code
            use_wav: If True, generate WAV and play it (better routing control)
                    If False, use direct output (simpler but less control)
            device_override: Override audio device for this speech
            volume: Volume multiplier (0.0 to 1.0, default 1.0)
        """
        if not self.available:
            logger.error("DECtalk not available")
            return False
        
        # Reset stop flag for new speech
        self.stop_requested = False
        
        # Get voice code if profile name provided
        voice_code = ""
        if voice_profile:
            if voice_profile in self.voice_profiles:
                voice_code = self.voice_profiles[voice_profile]
            elif voice_profile.startswith("[:"):
                voice_code = voice_profile
            else:
                logger.warning(f"Unknown DECtalk profile: {voice_profile}")
        
        # Check if text contains Moonbase Alpha-style singing commands or phoneme commands
        import re
        
        # Check for Moonbase Alpha style [<duration,pitch>]text patterns
        moonbase_pattern = r'\[<(\d+),(\d+)>\](\w+)'
        
        # Check for various phoneme patterns - be more flexible
        phoneme_patterns = [
            r'\[[^\]]*<\d+(?:,\d+)?>\]',    # Match any phoneme pattern with <duration> or <duration,pitch>
            r'\[:t\d+,\d+\]',                # [:t timing commands]
            r'\[:dial\d+\]',                 # [:dial phone number]
            r'\[:phone\s+on\]',              # [:phone on] already present
        ]
        
        has_phonemes = any(re.search(pattern, text) for pattern in phoneme_patterns)
        has_moonbase = re.search(moonbase_pattern, text)
        
        if has_moonbase:
            logger.info("Detected Moonbase Alpha singing syntax - translating to DECtalk phonemes")
            
            # Convert Moonbase Alpha syntax to DECtalk phoneme syntax
            def convert_moonbase_to_dectalk(match):
                duration = match.group(1)
                pitch = match.group(2)
                word = match.group(3)
                
                # Extended phoneme mapping for common Moonbase Alpha memes
                phoneme_map = {
                    'spayyyyyyyyyyyace': f's<100,{pitch}>p<100,{pitch}>ey<{duration},{pitch}>s',
                    'spayyyyyyyyyy': f's<100,{pitch}>p<100,{pitch}>ey<{duration},{pitch}>',
                    'space': f's<100,{pitch}>p<100,{pitch}>ey<{duration},{pitch}>s',
                    'john': f'jh<{duration},{pitch}>aa<{duration},{pitch}>n',
                    'madden': f'm<100,{pitch}>ae<{duration},{pitch}>d<100,{pitch}>ih<{duration},{pitch}>n',
                    'aeiou': f'ey<200,{pitch}>iy<200,{pitch}>ay<200,{pitch}>ow<200,{pitch}>uw<200,{pitch}>',
                    'uuuuuuuuuuuuuuuu': f'uw<{duration},{pitch}>',
                }
                
                # Check if we have a phoneme mapping
                word_lower = word.lower()
                
                # First check exact matches
                if word_lower in phoneme_map:
                    return f'[{phoneme_map[word_lower]}]'
                
                # Then check prefix matches
                for key in phoneme_map:
                    if word_lower.startswith(key[:3]):  # Match first 3 chars for variations
                        return f'[{phoneme_map[key]}]'
                
                # For single letters or very short words, try to convert to phonemes
                if len(word_lower) <= 2:
                    letter_phonemes = {
                        'a': 'ey', 'e': 'iy', 'i': 'ay', 'o': 'ow', 'u': 'uw',
                        's': 's', 'p': 'p', 't': 't', 'k': 'k', 'b': 'b',
                        'd': 'd', 'f': 'f', 'g': 'g', 'h': 'hh', 'j': 'jh',
                        'l': 'l', 'm': 'm', 'n': 'n', 'r': 'r', 'v': 'v',
                        'w': 'w', 'y': 'y', 'z': 'z'
                    }
                    phonemes = []
                    for letter in word_lower:
                        if letter in letter_phonemes:
                            phonemes.append(f'{letter_phonemes[letter]}<{duration},{pitch}>')
                    if phonemes:
                        return f'[{"".join(phonemes)}]'
                
                # Default: just pass through, DECtalk might handle it
                logger.warning(f"No phoneme mapping for '{word}', passing through")
                return f'[{word}<{duration},{pitch}>]'
            
            # Replace all Moonbase patterns
            converted_text = re.sub(moonbase_pattern, convert_moonbase_to_dectalk, text)
            
            # Enable phoneme mode
            if voice_code:
                full_text = voice_code + "[:phone on] " + converted_text
            else:
                full_text = "[:phone on] " + converted_text
                
        elif has_phonemes:
            logger.info("Detected phoneme syntax - enabling phoneme mode")
            
            # Clean up any problematic commands that might cause "command error"
            # Remove or fix problematic dial commands at the end
            cleaned_text = text
            
            # If text ends with [:np] or other voice commands, ensure proper spacing
            if re.search(r'\[:n[pbhfkdurw]\]$', cleaned_text):
                # Voice command at end is OK, but make sure there's no trailing issues
                pass
            
            # Already in phoneme format, just enable phoneme mode
            if voice_code:
                full_text = voice_code + "[:phone on] " + cleaned_text
            else:
                full_text = "[:phone on] " + cleaned_text
        else:
            # Combine voice code with text normally
            if voice_code:
                # Prepend the voice selection code to the text
                full_text = voice_code + " " + text
            else:
                full_text = text
        
        # Debug logging to see what text we're sending to DECtalk
        logger.info(f"DECtalk full_text being sent: '{full_text}'")
        
        try:
            if use_wav:
                # Generate WAV file and play it
                return self._speak_via_wav(full_text, None, device_override, volume)
            else:
                # Direct output (less control over routing)
                return self._speak_direct(full_text, None)
        except Exception as e:
            logger.error(f"DECtalk speak failed: {e}")
            return False
    
    def _speak_direct(self, text, voice_code):
        """Speak directly using say.exe"""
        # Kill any existing process first
        self.stop_speech()
        
        cmd = [self.dectalk_path]
        
        # Note: voice_code is now embedded in text, so we don't use -pre
        # Just pass the text which may contain DECtalk commands
        cmd.append(text)
        
        logger.debug(f"DECtalk command: {' '.join(cmd)}")
        
        # Run say.exe from its directory so it can find the dictionary
        dectalk_dir = Path(self.dectalk_path).parent
        # Use CREATE_NO_WINDOW flag to prevent console window popup on Windows
        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
        
        try:
            # Use Popen so we can track and kill the process
            self.current_process = subprocess.Popen(
                cmd, 
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True, 
                cwd=str(dectalk_dir), 
                creationflags=creationflags, 
                shell=False
            )
            
            # Wait for completion (but process can be killed)
            stdout, stderr = self.current_process.communicate()
            
            if self.current_process.returncode != 0 and self.current_process.returncode is not None:
                logger.error(f"DECtalk error: {stderr}")
                return False
                
        except Exception as e:
            logger.error(f"DECtalk execution error: {e}")
            return False
        finally:
            self.current_process = None
        
        return True
    
    def _speak_via_wav(self, text, voice_code, device_override=None, volume=1.0):
        """Generate WAV file and play it"""
        # Kill any existing process/stream first
        self.stop_speech()
        
        # Create temporary WAV file in local temp directory
        wav_path = get_temp_file(suffix='.wav', prefix='dectalk_')
        
        try:
            # Generate WAV with DECtalk
            cmd = [self.dectalk_path, "-w", wav_path]
            
            # Note: voice_code is now embedded in text, so we don't use -pre
            # Just pass the text which may contain DECtalk commands
            cmd.append(text)
            
            logger.debug(f"DECtalk WAV command: {' '.join(cmd)}")
            logger.info(f"DECtalk text argument: {repr(text)}")  # Show exact string being passed
            logger.info(f"DECtalk cmd list: {cmd}")  # Show command as list
            
            # Run say.exe to generate WAV from its directory so it can find the dictionary
            dectalk_dir = Path(self.dectalk_path).parent
            # Use CREATE_NO_WINDOW flag to prevent console window popup on Windows
            creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            
            # Use Popen for the WAV generation so we can kill it if needed
            self.current_process = subprocess.Popen(
                cmd, 
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True, 
                cwd=str(dectalk_dir), 
                creationflags=creationflags, 
                shell=False
            )
            
            # Wait for WAV generation to complete
            stdout, stderr = self.current_process.communicate(timeout=10)
            
            if self.current_process.returncode != 0:
                logger.error(f"DECtalk WAV generation error: {stderr}")
                return False
            
            self.current_process = None
            
            # Play the WAV file with volume adjustment
            return self._play_wav(wav_path, device_override, volume)
            
        except subprocess.TimeoutExpired:
            logger.error("DECtalk WAV generation timed out")
            if self.current_process:
                self.current_process.kill()
                self.current_process = None
            return False
        except Exception as e:
            logger.error(f"DECtalk WAV generation error: {e}")
            return False
        finally:
            # Clean up temporary file
            try:
                os.unlink(wav_path)
            except:
                pass
    
    def _play_wav(self, wav_path, device_override=None, volume=1.0):
        """Play a WAV file to the specified audio device with volume control"""
        if not os.path.exists(wav_path):
            logger.error(f"WAV file not found: {wav_path}")
            return False
        
        if not self.pyaudio_available:
            logger.warning("PyAudio not available, cannot play WAV file")
            return False
            
        try:
            # Initialize PyAudio if needed
            if not self.pyaudio:
                self.pyaudio = pyaudio.PyAudio()
            
            # Open WAV file
            wf = wave.open(wav_path, 'rb')
            
            # Get audio parameters
            channels = wf.getnchannels()
            sample_width = wf.getsampwidth()
            framerate = wf.getframerate()
            
            # Determine output device
            output_device = self.audio_device_index
            if device_override:
                # Special handling for VoiceMeeter
                if 'voicemeeter' in device_override.lower():
                    # Look specifically for VoiceMeeter Input device
                    for i in range(self.pyaudio.get_device_count()):
                        info = self.pyaudio.get_device_info_by_index(i)
                        if 'voicemeeter' in info['name'].lower() and 'input' in info['name'].lower():
                            output_device = i
                            logger.info(f"DECtalk routing to VoiceMeeter device: {info['name']} (index {i})")
                            break
                else:
                    # Find device by name
                    for i in range(self.pyaudio.get_device_count()):
                        info = self.pyaudio.get_device_info_by_index(i)
                        if device_override.lower() in info['name'].lower():
                            output_device = i
                            logger.info(f"DECtalk routing to device: {info['name']} (index {i})")
                            break
            
            # Open audio stream
            self.current_stream = self.pyaudio.open(
                format=self.pyaudio.get_format_from_width(sample_width),
                channels=channels,
                rate=framerate,
                output=True,
                output_device_index=output_device
            )
            
            # Play audio with volume adjustment
            chunk_size = 1024
            data = wf.readframes(chunk_size)
            
            # Import struct for audio data manipulation
            import struct
            
            # Track if we were stopped during THIS playback
            was_stopped = False
            
            # Reset stop flag right before playback starts
            self.stop_requested = False
            
            while data and self.current_stream and not self.stop_requested:
                # Apply volume adjustment if not 1.0
                if volume != 1.0 and sample_width == 2:  # 16-bit audio
                    # Convert bytes to list of 16-bit samples
                    samples = list(struct.unpack(f'<{len(data)//2}h', data))
                    # Apply volume
                    samples = [int(min(max(s * volume, -32768), 32767)) for s in samples]
                    # Convert back to bytes
                    data = struct.pack(f'<{len(samples)}h', *samples)
                
                try:
                    if self.current_stream:
                        self.current_stream.write(data)
                except:
                    break  # Stream was stopped
                data = wf.readframes(chunk_size)
            
            # Check if we exited due to stop request
            if self.stop_requested:
                was_stopped = True
            
            # Clean up - wrapped in try/except in case stop_speech already closed it
            try:
                if self.current_stream:
                    self.current_stream.stop_stream()
                    self.current_stream.close()
                    self.current_stream = None
            except:
                pass  # Stream may have been closed by stop_speech
            
            try:
                wf.close()
            except:
                pass
            
            # Return True if completed normally, False if stopped
            return not was_stopped
            
        except Exception as e:
            logger.error(f"WAV playback error: {e}")
            return False
    
    def speak_async(self, text, voice_profile=None, use_wav=True, device_override=None, volume=1.0):
        """Speak text asynchronously"""
        thread = threading.Thread(
            target=self.speak,
            args=(text, voice_profile, use_wav, device_override, volume),
            daemon=True
        )
        thread.start()
        return thread
    
    def test_voice(self, voice_profile):
        """Test a specific DECtalk voice"""
        if voice_profile in self.voice_profiles:
            test_text = f"Hello, this is {voice_profile} speaking."
            return self.speak(test_text, voice_profile)
        else:
            logger.error(f"Unknown voice profile: {voice_profile}")
            return False
    
    def get_available_profiles(self):
        """Get list of available DECtalk profiles"""
        return list(self.voice_profiles.keys())
    
    def stop_speech(self):
        """Stop any current DECtalk speech"""
        stopped = False
        
        # Set stop flag FIRST - this will cause playback loop to exit gracefully
        self.stop_requested = True
        
        # Kill the say.exe process if it's running
        if self.current_process:
            try:
                logger.info("Killing DECtalk say.exe process")
                self.current_process.terminate()
                # Give it a moment to terminate gracefully
                try:
                    self.current_process.wait(timeout=0.5)
                except subprocess.TimeoutExpired:
                    # Force kill if it didn't terminate
                    self.current_process.kill()
                self.current_process = None
                stopped = True
            except Exception as e:
                logger.error(f"Error killing DECtalk process: {e}")
        
        # Give the playback loop a moment to exit before touching the stream
        time.sleep(0.05)
        
        # Stop the audio stream if it's playing
        if self.current_stream:
            try:
                logger.info("Stopping DECtalk audio stream")
                stream = self.current_stream
                self.current_stream = None  # Clear reference first
                try:
                    stream.stop_stream()
                except:
                    pass  # May already be stopped
                try:
                    stream.close()
                except:
                    pass  # May already be closed
                stopped = True
            except Exception as e:
                logger.error(f"Error stopping audio stream: {e}")
        
        # Note: Don't reset stop_requested here - let the next speak() call reset it
        # This allows callers to check if speech was stopped vs failed
        
        if stopped:
            logger.info("DECtalk speech stopped successfully")
        
        return stopped
    
    def cleanup(self):
        """Clean up resources"""
        if self.pyaudio and self.pyaudio_available:
            self.pyaudio.terminate()
            self.pyaudio = None


class DECtalkManager:
    """Manages DECtalk integration with fallback to SAPI5"""
    
    def __init__(self, audio_manager=None):
        self.dectalk = DECtalkNative()
        self.audio_manager = audio_manager  # SAPI5 manager for fallback
        self.use_dectalk = self.dectalk.is_available()
        
    def is_dectalk_voice(self, voice_name):
        """Check if a voice name is a DECtalk profile"""
        if not voice_name:
            return False
        
        # Check if it's marked as DECtalk
        if voice_name.startswith("[DECtalk] "):
            return True
        
        # Check if it's a known DECtalk profile
        profile_name = voice_name.replace("[DECtalk] ", "")
        return profile_name in self.dectalk.voice_profiles
    
    def speak(self, text, voice_name=None, device=None):
        """
        Speak text using DECtalk or SAPI5
        
        Args:
            text: Text to speak
            voice_name: Voice name (DECtalk profile or SAPI5 voice)
            device: Audio output device
        """
        # Check if it's a DECtalk voice
        if self.is_dectalk_voice(voice_name) and self.use_dectalk:
            # Extract profile name
            profile_name = voice_name.replace("[DECtalk] ", "")
            
            # Set audio device if specified
            if device and "voicemeeter" in device.lower():
                self.dectalk.set_audio_device("VoiceMeeter Input")
            elif device:
                self.dectalk.set_audio_device(device)
            
            # Speak with DECtalk
            logger.info(f"Speaking with DECtalk profile: {profile_name}")
            return self.dectalk.speak(text, profile_name, use_wav=True)
        else:
            # Fall back to SAPI5
            if self.audio_manager:
                logger.info(f"Speaking with SAPI5 voice: {voice_name}")
                return self.audio_manager.speak(text)
            else:
                logger.error("No audio manager available for SAPI5")
                return False
    
    def cleanup(self):
        """Clean up resources"""
        self.dectalk.cleanup()
    
    def stop(self):
        """Stop any current DECtalk speech"""
        return self.dectalk.stop_speech()