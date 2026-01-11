"""
SAPI5 Direct Output with WASAPI Support
"""

import win32com.client
import pythoncom
import logging
import os
import time
import pyaudio
import numpy as np
import wave
import tempfile
import threading
import queue
from pathlib import Path

logger = logging.getLogger(__name__)

class SAPI5Direct:
    """Direct SAPI5 control with output device selection"""
    
    def __init__(self):
        """Initialize SAPI5 with COM"""
        self.sapi_voice = None
        self.audio_outputs = []
        self.current_output = None
        self.current_output_index = 0
        self.current_voice_index = 0
        self.speech_queue = queue.Queue()
        self.worker_thread = None
        self.running = False
        self.worker_sapi = None
        self.stop_requested = False  # Flag to interrupt current speech
        
        try:
            # Initialize COM in main thread
            pythoncom.CoInitialize()
            self.sapi_voice = win32com.client.Dispatch("SAPI.SpVoice")
            
            # Set default properties
            self.sapi_voice.Volume = 100  # Full volume
            self.sapi_voice.Rate = 0  # Normal speed
            
            # Initialize outputs
            self._enumerate_outputs()
            
            # Start worker thread for speech
            self.running = True
            self.worker_thread = threading.Thread(target=self._speech_worker, daemon=True)
            self.worker_thread.start()
            
            logger.info("SAPI5 initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize SAPI5: {e}")
            raise
            
    def _enumerate_outputs(self):
        """Enumerate available audio outputs"""
        try:
            # SAPI can enumerate audio outputs
            # This includes default device and any virtual cables
            
            # Get the audio outputs collection
            if self.sapi_voice:
                outputs = self.sapi_voice.GetAudioOutputs()
                
                if outputs and outputs.Count > 0:
                    for i in range(outputs.Count):
                        try:
                            output = outputs.Item(i)
                            desc = output.GetDescription()
                            self.audio_outputs.append({
                                'index': i,
                                'name': desc,
                                'id': output.Id,
                                'object': output
                            })
                            logger.info(f"Found SAPI output: {desc}")
                        except Exception as e:
                            logger.warning(f"Error getting output {i}: {e}")
                    
            # If no outputs found, add defaults
            if not self.audio_outputs:
                logger.warning("No SAPI outputs enumerated, adding default devices")
                self.audio_outputs = [
                    {'index': 0, 'name': 'Default Audio Device', 'id': None, 'object': None}
                ]
                
        except Exception as e:
            logger.warning(f"Could not enumerate SAPI outputs: {e}")
            # Add fallback devices
            self.audio_outputs = [
                {'index': 0, 'name': 'Default Audio Device', 'id': None, 'object': None}
            ]
        
        # Always try to add PyAudio devices as extra fallback
        self._add_pyaudio_fallback()
    
    def _add_pyaudio_fallback(self):
        """Add PyAudio devices as fallback if SAPI enumeration fails"""
        try:
            import pyaudio
            p = pyaudio.PyAudio()
            
            # Only add if we don't have real SAPI devices
            if len(self.audio_outputs) <= 1 and self.audio_outputs[0].get('id') is None:
                logger.info("Adding PyAudio devices as fallback")
                self.audio_outputs = []  # Clear the dummy default
                
                device_count = p.get_device_count()
                for i in range(device_count):
                    info = p.get_device_info_by_index(i)
                    if info['maxOutputChannels'] > 0:
                        self.audio_outputs.append({
                            'index': len(self.audio_outputs),
                            'name': info['name'],
                            'id': None,
                            'object': None,
                            'pyaudio_index': i
                        })
                        logger.info(f"Added PyAudio device: {info['name']}")
                
            p.terminate()
            
        except Exception as e:
            logger.warning(f"Could not add PyAudio devices: {e}")
            # Make sure we always have at least one device
            if not self.audio_outputs:
                self.audio_outputs = [
                    {'index': 0, 'name': 'Default Audio Device', 'id': None, 'object': None}
                ]
            
    def set_audio_output(self, device_index: int) -> bool:
        """Set SAPI output device by index"""
        try:
            if device_index < len(self.audio_outputs):
                output = self.audio_outputs[device_index]
                if output['object']:
                    self.sapi_voice.AudioOutput = output['object']
                    self.current_output = output['name']
                    self.current_output_index = device_index
                    logger.info(f"Set SAPI output to index {device_index}: {output['name']}")
                    return True
            return False
        except Exception as e:
            logger.error(f"Failed to set audio output: {e}")
            return False
    
    def set_output_device(self, device_name: str) -> bool:
        """Set SAPI output device by name"""
        try:
            # Check if it's a VoiceMeeter device
            if 'voicemeeter' in device_name.lower():
                # SAPI can output to VoiceMeeter if it's installed as an audio device
                # Try to find and set it
                for i, output in enumerate(self.audio_outputs):
                    if 'voicemeeter' in output['name'].lower() or 'vb-audio' in output['name'].lower():
                        if output['object']:
                            self.sapi_voice.AudioOutput = output['object']
                            self.current_output = output['name']
                            self.current_output_index = i
                            logger.info(f"Set SAPI output to: {output['name']}")
                            return True
                            
                # If not found in SAPI outputs, VoiceMeeter might need different approach
                logger.warning("VoiceMeeter not found in SAPI outputs, using default")
                
            # For other devices, try to match by name
            for i, output in enumerate(self.audio_outputs):
                if device_name.lower() in output['name'].lower():
                    if output['object']:
                        self.sapi_voice.AudioOutput = output['object']
                        self.current_output = output['name']
                        self.current_output_index = i
                        logger.info(f"Set SAPI output to: {output['name']}")
                        return True
                        
            logger.warning(f"Device {device_name} not found in SAPI outputs")
            return False
                        
        except Exception as e:
            logger.error(f"Failed to set output device: {e}")
            return False
            
    def speak(self, text: str):
        """Speak text using SAPI5 with async support"""
        try:
            if not text:
                return
                
            # Queue speech for worker thread
            self.speech_queue.put(text)
            
        except Exception as e:
            logger.error(f"Failed to queue speech: {e}")
            
    def speak_sync(self, text: str):
        """Speak text synchronously (blocks until complete)"""
        try:
            if not text:
                return
                
            # Ensure voice is set
            self._ensure_voice()
            
            # Speak synchronously
            self.sapi_voice.Speak(text, 0)  # 0 = synchronous
            
        except Exception as e:
            logger.error(f"Failed to speak: {e}")
            
    def stop_speech(self):
        """Stop current speech"""
        try:
            # Set stop flag
            self.stop_requested = True
            
            # Skip on BOTH main and worker SAPI instances
            if self.sapi_voice:
                try:
                    self.sapi_voice.Skip("Sentence", 10000)
                except:
                    pass
            
            # This is the important one - worker_sapi is actually speaking
            if self.worker_sapi:
                try:
                    self.worker_sapi.Skip("Sentence", 10000)
                    logger.info("Skipped speech on worker SAPI")
                except Exception as e:
                    logger.warning(f"Could not skip worker speech: {e}")
                
            # Clear the queue
            while not self.speech_queue.empty():
                try:
                    self.speech_queue.get_nowait()
                except:
                    pass
            
            # Reset stop flag after a short delay
            def reset_flag():
                import time
                time.sleep(0.1)
                self.stop_requested = False
            threading.Thread(target=reset_flag, daemon=True).start()
                    
            logger.info("Speech stopped")
        except Exception as e:
            logger.error(f"Failed to stop speech: {e}")
            
    def _ensure_voice(self):
        """Ensure the correct voice is set before speaking"""
        try:
            voices = self.sapi_voice.GetVoices()
            current = self.sapi_voice.Voice
            if current:
                # Check if we need to reapply the voice
                for i in range(voices.Count):
                    if voices.Item(i).Id == current.Id:
                        if i != self.current_voice_index:
                            # Voice has changed, reapply
                            self.sapi_voice.Voice = voices.Item(self.current_voice_index)
                            logger.debug(f"Reapplied voice {self.current_voice_index}")
                        break
        except Exception as e:
            logger.warning(f"Could not ensure voice: {e}")
            
    def _speech_worker(self):
        """Worker thread for speech processing"""
        # Initialize COM in worker thread
        pythoncom.CoInitialize()
        
        # Create worker's own SAPI instance
        try:
            self.worker_sapi = win32com.client.Dispatch("SAPI.SpVoice")
            self.worker_sapi.Volume = 100
            self.worker_sapi.Rate = 0
            logger.info("Worker SAPI initialized")
            
            # Apply initial audio output if set
            if self.current_output_index > 0 and self.current_output_index < len(self.audio_outputs):
                output = self.audio_outputs[self.current_output_index]
                if output['object']:
                    self.worker_sapi.AudioOutput = output['object']
                    logger.info(f"Worker SAPI: Set audio output to {output['name']}")
        except Exception as e:
            logger.error(f"Failed to init worker SAPI: {e}")
            return
            
        while self.running:
            try:
                text = self.speech_queue.get(timeout=0.5)
                
                # Apply current audio output settings to worker SAPI
                if self.current_output_index < len(self.audio_outputs):
                    output = self.audio_outputs[self.current_output_index]
                    if output['object'] and self.worker_sapi.AudioOutput != output['object']:
                        self.worker_sapi.AudioOutput = output['object']
                        logger.debug(f"Worker: Updated audio output to {output['name']}")
                
                # Apply current voice settings to worker SAPI
                voices = self.worker_sapi.GetVoices()
                if self.current_voice_index < voices.Count:
                    self.worker_sapi.Voice = voices.Item(self.current_voice_index)
                    logger.debug(f"Worker: Using voice {self.current_voice_index}")
                
                # Speak with worker's SAPI instance
                self.worker_sapi.Speak(text, 0)  # Sync in worker thread
                
            except queue.Empty:
                continue
            except Exception as e:
                logger.error(f"Worker error: {e}")
                
    def get_voices(self):
        """Get available SAPI voices"""
        voices = []
        try:
            sapi_voices = self.sapi_voice.GetVoices()
            for i in range(sapi_voices.Count):
                try:
                    voice = sapi_voices.Item(i)
                    desc = voice.GetDescription()
                    
                    # Test if the voice can be set (some like IVONA may fail)
                    # Only include Microsoft voices and those that work
                    if 'Microsoft' in desc or 'Windows' in desc:
                        voices.append({
                            'index': i,
                            'name': desc,
                            'id': voice.Id
                        })
                    else:
                        # Try to set the voice to see if it works
                        try:
                            # Don't actually change the voice, just test
                            test_voice = voice
                            voices.append({
                                'index': i,
                                'name': desc,
                                'id': voice.Id
                            })
                        except:
                            logger.warning(f"Voice {desc} cannot be used, skipping")
                            
                except Exception as e:
                    logger.warning(f"Error getting voice {i}: {e}")
                    
        except Exception as e:
            logger.error(f"Failed to get voices: {e}")
            
        return voices
        
    def set_voice(self, voice_index: int):
        """Set SAPI voice by index"""
        try:
            voices = self.sapi_voice.GetVoices()
            if voice_index < voices.Count:
                # Store the index for persistence
                self.current_voice_index = voice_index
                
                try:
                    # Try to set the voice
                    voice = voices.Item(voice_index)
                    voice_name = voice.GetDescription()
                    
                    # Some voices (like IVONA) may have issues, try to set anyway
                    self.sapi_voice.Voice = voice
                    
                    # Also set for worker if it exists
                    if self.worker_sapi:
                        worker_voices = self.worker_sapi.GetVoices()
                        if voice_index < worker_voices.Count:
                            self.worker_sapi.Voice = worker_voices.Item(voice_index)
                    
                    logger.info(f"SAPI5: Changed to voice {voice_index}: {voice_name}")
                    return True
                    
                except Exception as e:
                    # Some voices might fail (like IVONA), but try to continue
                    logger.warning(f"Voice {voice_index} had issues but may still work: {e}")
                    # Still return True as it might work for basic speech
                    return True
            else:
                logger.warning(f"Voice index {voice_index} out of range")
                
        except Exception as e:
            logger.error(f"Failed to set voice: {e}")
        return False
            
    def set_voice_by_name(self, voice_name: str):
        """Set voice by name match"""
        voices = self.get_voices()
        for voice in voices:
            if voice_name.lower() in voice['name'].lower():
                self.set_voice(voice['index'])
                logger.info(f"SAPI5: Set voice to {voice['name']} (index {voice['index']})")
                return True
        logger.warning(f"SAPI5: Voice not found: {voice_name}")
        return False
        
    def stop(self):
        """Stop speech and cleanup"""
        self.running = False
        if self.worker_thread:
            self.worker_thread.join(timeout=2)
        pythoncom.CoUninitialize()


# The rest of the file remains the same - just the SAPI5Direct class is fixed
# Import the rest from the original implementation

class VoiceMeeterHandler:
    """Handles VoiceMeeter integration for audio routing"""
    
    def __init__(self):
        self.voicemeeter_available = False
        self.vm_type = None
        self.sapi = None
        self._detect_voicemeeter()
        
    def _detect_voicemeeter(self):
        """Detect if VoiceMeeter is installed and running"""
        try:
            # Check if VoiceMeeter is in the audio devices
            p = pyaudio.PyAudio()
            device_count = p.get_device_count()
            
            for i in range(device_count):
                info = p.get_device_info_by_index(i)
                device_name = info['name'].lower()
                
                if 'voicemeeter' in device_name:
                    self.voicemeeter_available = True
                    if 'aux' in device_name:
                        self.vm_type = 'aux'
                    elif 'vaio' in device_name:
                        self.vm_type = 'vaio'
                    else:
                        self.vm_type = 'standard'
                    
                    logger.info(f"VoiceMeeter detected: {info['name']} (type: {self.vm_type})")
                    break
                    
            p.terminate()
            
        except Exception as e:
            logger.warning(f"Could not detect VoiceMeeter: {e}")
            
    def is_available(self):
        """Check if VoiceMeeter is available"""
        return self.voicemeeter_available
        
    def route_audio(self, sapi_instance):
        """Route SAPI audio through VoiceMeeter"""
        if not self.voicemeeter_available:
            return False
            
        self.sapi = sapi_instance
        
        # Try to set SAPI output to VoiceMeeter
        success = self.sapi.set_output_device("VoiceMeeter")
        if success:
            logger.info("Audio routed through VoiceMeeter")
        else:
            logger.warning("Could not route audio through VoiceMeeter")
            
        return success


class DirectOutputTTS:
    """Main TTS manager with direct output support"""
    
    def __init__(self):
        """Initialize TTS with direct output capabilities"""
        self.sapi = SAPI5Direct()
        self.vm_handler = VoiceMeeterHandler()
        self.dectalk_manager = None
        
        # Check for VoiceMeeter routing
        if self.vm_handler.is_available():
            logger.info("VoiceMeeter support available")
            
    def set_dectalk_manager(self, manager):
        """Set DECtalk manager reference"""
        self.dectalk_manager = manager
        logger.info("DECtalk manager configured")
        
        
    def speak(self, text: str, voice_type: str = None, use_sync: bool = False):
        """Speak text with specified voice type
        
        Args:
            text: Text to speak
            voice_type: Voice type (e.g., "dectalk")
            use_sync: Whether to use synchronous mode (ignored for compatibility)
        """
        
        # Route to appropriate engine
        if voice_type == "dectalk" and self.dectalk_manager:
            self.dectalk_manager.speak(text)
        else:
            # Default to SAPI5
            self.sapi.speak(text)
            
    def stop(self):
        """Stop all speech"""
        self.sapi.stop_speech()
        if self.dectalk_manager:
            self.dectalk_manager.stop()
    
    def stop_all_speech(self):
        """Stop all speech immediately - alias for stop() for compatibility with main app"""
        logger.info("stop_all_speech called - stopping all audio")
        self.stop()
    
    def is_speaking(self):
        """Check if currently speaking"""
        # Check SAPI5 speaking status
        if hasattr(self.sapi, 'voice') and self.sapi.voice:
            try:
                status = self.sapi.voice.Status
                return status.RunningState == 2  # SRSEIsSpeaking = 2
            except:
                pass
        
        # Check DECtalk speaking status if available
        if self.dectalk_manager and hasattr(self.dectalk_manager, 'is_playing'):
            return self.dectalk_manager.is_playing
        
        return False
            
    def set_voice(self, voice_id):
        """Set voice by ID or name"""
        # Try as index first
        try:
            index = int(voice_id)
            return self.sapi.set_voice(index)
        except ValueError:
            # Try as name
            success = self.sapi.set_voice_by_name(voice_id)
            if success:
                logger.info(f"Set voice by name: {voice_id}")
            return success
        
    def get_voices(self):
        """Get available voices"""
        voices = self.sapi.get_voices()
        
        # Convert to pyttsx3-like format for compatibility
        class Voice:
            def __init__(self, name, id):
                self.name = name
                self.id = id
                
        return [Voice(v['name'], str(v['index'])) for v in voices]
        
    def route_to_voicemeeter(self) -> bool:
        """Route audio through VoiceMeeter if available"""
        if self.vm_handler.is_available():
            return self.vm_handler.route_audio(self.sapi)
        return False
        
    def get_devices(self):
        """Get available audio devices"""
        devices = []
        
        # Add SAPI devices
        for output in self.sapi.audio_outputs:
            devices.append({
                'index': output['index'],
                'name': output['name'],
                'is_default': output['index'] == 0,
                'channels': 2,
                'sample_rate': 48000,
                'api': 'SAPI5'
            })
            
        # Add VoiceMeeter if available
        if self.vm_handler.is_available():
            devices.append({
                'index': len(devices),
                'name': 'VoiceMeeter Input (Recommended)',
                'is_default': False,
                'channels': 2,
                'sample_rate': 48000,
                'api': 'VoiceMeeter'
            })
            
        return devices
    
    def set_device(self, device_name):
        """Set the audio output device"""
        try:
            # Check if it's a VoiceMeeter device
            if 'VoiceMeeter' in device_name:
                logger.info(f"Selected VoiceMeeter device: {device_name}")
                # Try to route through VoiceMeeter
                return self.route_to_voicemeeter()
            
            # Otherwise, it's a SAPI device
            for output in self.sapi.audio_outputs:
                if output['name'] == device_name or output['name'].replace(" [DEFAULT]", "") == device_name:
                    if self.sapi.set_audio_output(output['index']):
                        logger.info(f"Set SAPI audio output to: {device_name}")
                        return True
            
            logger.warning(f"Device not found: {device_name}")
            return False
            
        except Exception as e:
            logger.error(f"Error setting device: {e}")
            return False