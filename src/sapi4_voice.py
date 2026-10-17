"""
Old SAPI 4 voices: Microsoft Sam/Mike/Mary 4.0 (and their Hall/Space/RoboSoft variants), Lernout & Hauspie
TruVoice (BonziBuddy's "Adult Male #2"), ... - whatever SAPI 4 voices are installed on this PC.

SAPI 4 engines are 32-bit only, so a small 32-bit helper (sapi4_tts.exe, built from tools/sapi4_tts) writes each
line to a WAV file, which plays like DECtalk: same output device, volume slider, !stop, mutes and music ducking.
Nothing from Microsoft or L&H is included - the voices have to be installed on the PC (SAPI 4 runtime spchapi.exe
plus the voice packages).
"""

import os
import sys
import time
import logging
import threading
import subprocess
from pathlib import Path

logger = logging.getLogger(__name__)

PREFIX = "[SAPI4] "
NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
TRUVOICE_TAIL = ", American English (TruVoice)"
# Ready-made voices: name -> (SAPI 4 voice, pitch, speed). BonziBuddy used TruVoice Adult Male #2 like this.
PRESETS = {
    "Bonzi": ("Adult Male #2, American English (TruVoice)", 140, 157),
}


def helper_path():
    """sapi4_tts.exe: unpacked next to the program in the built exe, in src\\ when run from source"""
    base = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
    return base / 'sapi4_tts.exe'


def sapi4_installed():
    """True if the SAPI 4 runtime is installed (its text-to-speech enumerator is registered)"""
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                                       r"WOW6432Node\CLSID\{D67C0280-C743-11cd-80E5-00AA003E4B50}"))
        return True
    except OSError:
        return False


def display_name(mode):
    """'Adult Male #2, American English (TruVoice)' -> 'Adult Male #2 (TruVoice)'; 'Sam' stays 'Sam'"""
    return mode[:-len(TRUVOICE_TAIL)] + " (TruVoice)" if mode.endswith(TRUVOICE_TAIL) else mode


class Sapi4Voices:
    def __init__(self, idle_minutes=10):
        self.idle_minutes = idle_minutes
        self.proc = None
        self.lock = threading.Lock()
        self.last_used = time.time()
        self.voices = {}          # display name -> (SAPI 4 voice name, pitch or None, speed or None)
        self._watch = None

    def available(self):
        return helper_path().exists() and sapi4_installed()

    # ---------------------------------------------------------------- helper process
    def _start(self):
        if self.proc and self.proc.poll() is None:
            return self.proc
        self.proc = subprocess.Popen([str(helper_path())], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                     stderr=subprocess.DEVNULL, text=True, encoding='utf-8', bufsize=1,
                                     creationflags=NO_WINDOW)
        if self._watch is None:
            self._watch = threading.Thread(target=self._idle_watch, daemon=True, name='SAPI4 idle')
            self._watch.start()
        return self.proc

    def _kill(self):
        proc, self.proc = self.proc, None
        if proc:
            try:
                proc.kill()
                proc.wait(timeout=3)  # so it lets go of the file it was writing
            except (OSError, subprocess.TimeoutExpired):
                pass

    def _ask(self, line, timeout, should_stop=None, until=None):
        """Send one command and collect the reply lines (until a line for which until(line) is true)"""
        proc = self._start()
        proc.stdin.write(line + "\n")
        proc.stdin.flush()
        lines, result = [], {}

        def read():
            while True:
                answer = proc.stdout.readline()
                if not answer:
                    break
                answer = answer.rstrip("\r\n")
                lines.append(answer)
                if until is None or until(answer):
                    break
            result['done'] = True
        reader = threading.Thread(target=read, daemon=True)
        reader.start()
        deadline = time.time() + timeout
        while reader.is_alive() and time.time() < deadline:
            reader.join(0.05)
            if should_stop and should_stop():
                break
        if reader.is_alive():  # stopped or stuck: it can't take another command now
            self._kill()
            return None
        return lines

    # ---------------------------------------------------------------- voices
    def refresh(self):
        """Ask the helper which SAPI 4 voices are installed (plus the presets that can be made from them)"""
        voices = {}
        if self.available():
            with self.lock:
                try:
                    lines = self._ask("LIST", 15, until=lambda l: l == "END") or []
                except OSError as e:
                    logger.warning(f"SAPI 4 helper didn't start: {e}")
                    lines = []
            modes = [l.split("\t")[1] for l in lines if l.startswith("VOICE\t") and l.count("\t") >= 4]
            for name, (mode, pitch, speed) in PRESETS.items():
                if mode in modes:
                    voices[name] = (mode, pitch, speed)
            for mode in modes:
                voices.setdefault(display_name(mode), (mode, None, None))
        self.voices = voices
        return voices

    def voice_names(self):
        return [PREFIX + name for name in self.voices]

    def render(self, text, name, wav_path, should_stop=None):
        """Write text in the voice called name (without the prefix) to wav_path. True when done."""
        if name not in self.voices:
            raise FileNotFoundError(f"SAPI 4 voice {name} isn't installed")
        mode, pitch, speed = self.voices[name]
        text = ' '.join((text or '').split())  # also removes tabs and line breaks (the helper's separators)
        if not text:
            return False
        if os.path.exists(wav_path):
            os.remove(wav_path)
        command = "\t".join(["SPEAK", mode, str(pitch) if pitch else "-", str(speed) if speed else "-",
                             str(wav_path), text])
        with self.lock:
            self.last_used = time.time()
            for attempt in (1, 2):
                try:
                    lines = self._ask(command, 10 + len(text) / 10, should_stop=should_stop)
                except OSError:
                    self._kill()
                    continue  # the helper went away - start a fresh one
                self.last_used = time.time()
                if lines is None:
                    return False  # stopped (or stuck) - not retried
                if lines and lines[-1] == "OK" and os.path.exists(wav_path):
                    return True
                logger.warning(f"SAPI 4 voice {name}: {lines[-1] if lines else 'no answer'}")
                self._kill()
        return False

    def _idle_watch(self):
        while True:
            time.sleep(30)
            if self.idle_minutes and self.proc and time.time() - self.last_used > self.idle_minutes * 60 \
                    and not self.lock.locked():
                self._kill()

    def shutdown(self):
        proc = self.proc
        if proc and proc.poll() is None:
            try:
                proc.stdin.write("QUIT\n")
                proc.stdin.flush()
                proc.wait(timeout=2)
            except (OSError, subprocess.TimeoutExpired):
                pass
        self._kill()
