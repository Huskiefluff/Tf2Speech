"""
Chat soundboard, in the style of Sven Co-op's ChatSounds plugin: a player types a sound's word in chat and it
plays straight away (no queue) through the same output as the TTS.

Sounds: every audio file in the "sounds" folder next to the TTS (sub-folders too). The file name is the word:
yee.wav -> "yee". A Sven-style list file (chatsounds.txt / ChatSounds.txt, lines like "bleh twlz/bleh.wav")
in that folder adds or renames words, so ChatSounds packs can be dropped in as they are.

Modes:  open    - the whole message is the word: "yee", or with a pitch "yee 150"
        command - a command first: "!s yee" / "!s yee 150" (the command is set in the Soundboard window)
Pitch 100 is normal; 50-200 plays slower/lower .. faster/higher, like GoldSrc pitch.
"""

import logging
import os
import re
import subprocess
import sys
import threading
import time
from collections import OrderedDict
from pathlib import Path

import tkinter as tk
from tkinter import ttk
try:
    import ui_theme  # sizes that follow Windows scaling, theme hint colour
except Exception:
    ui_theme = None

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pyaudio
except ImportError:
    pyaudio = None

from dj_player import find_output_device, NO_WINDOW, AUDIO_EXTENSIONS, fmt_time

logger = logging.getLogger(__name__)

RATE = 48000
CHANNELS = 2
CHUNK_FRAMES = 1024
LIST_FILES = ('chatsounds.txt', 'ChatSounds.txt', 'soundboard.txt')

DEFAULTS = {
    'mode': 'off',                 # 'off' | 'open' | 'command'
    'command': '!s',
    'who': 'everyone',             # 'admins' | 'everyone'
    'cooldown': 10,                # seconds between one player's sounds
    'global_gap': 1.0,             # seconds between any two sounds
    'admins_skip_cooldown': True,
    'volume': 60,                  # %
    'pitch_enabled': True,
    'duck_music': False,           # lower the DJ music while a sound plays (off = sounds play over it)
    'max_at_once': 3,              # sounds playing at the same time; more are dropped, never queued
    'bass_db': 0,                  # bass boost/cut in dB (TF2's voice codec is boomy)
    'max_seconds': 30,             # clips are cut off after this long
}


class SoundboardEngine:
    """Thread-safe: handle_chat() is called from the chat thread, the window calls the rest."""

    def __init__(self, settings, save_config, sounds_dir, tts_device=None, ffmpeg_path=None, ffprobe_path=None):
        self.settings = settings
        self._save_config = save_config
        self.sounds_dir = Path(sounds_dir)
        self._tts_device = tts_device or (lambda: 'Default')
        self._ffmpeg = ffmpeg_path or (lambda: None)
        self._ffprobe = ffprobe_path or (lambda: None)
        self.sounds = {}            # word -> file path
        self.durations = {}         # word -> seconds (filled in by a background probe)
        self.version = 0            # bumps when the list changes so the window redraws
        self.status = ''
        self._lock = threading.Lock()
        self._last_by_player = {}   # username -> time of their last sound
        self._last_any = 0.0
        self._active = []           # stop-flags (threading.Event) of sounds playing now
        self._cache = OrderedDict() # (word, pitch) -> decoded PCM, small LRU
        self._pa = None
        self._pa_lock = threading.Lock()
        self.reload()

    # ------------------------------------------------------------------ settings
    def get(self, key):
        return self.settings.get(key, DEFAULTS.get(key))

    def set(self, key, value):
        self.settings[key] = value
        try:
            self._save_config()
        except Exception as e:
            logger.error(f"Soundboard: could not save settings: {e}")

    # ------------------------------------------------------------------ sound list
    def reload(self):
        """Scan the sounds folder (and any ChatSounds-style list file in it)"""
        sounds = {}
        try:
            self.sounds_dir.mkdir(parents=True, exist_ok=True)
            for path in sorted(self.sounds_dir.rglob('*'), key=lambda p: str(p).lower()):
                if path.suffix.lower() in AUDIO_EXTENSIONS and path.is_file():
                    sounds.setdefault(path.stem.lower(), str(path))
            for name in LIST_FILES:
                list_file = self.sounds_dir / name
                if list_file.is_file():
                    for line in list_file.read_text(encoding='utf-8', errors='replace').splitlines():
                        line = line.strip()
                        if not line or line.startswith(('//', '#', '[')):
                            continue
                        parts = line.split(None, 1)
                        if len(parts) == 2:
                            target = (self.sounds_dir / parts[1].strip().strip('"')).resolve()
                            if target.is_file():
                                sounds[parts[0].lower()] = str(target)
        except Exception as e:
            self.status = f"Couldn't read the sounds folder: {e}"
        with self._lock:
            self.sounds = sounds
            self._cache.clear()
            self.durations = {w: d for w, d in self.durations.items() if w in sounds}
        self.version += 1
        self.status = f"{len(sounds)} sound(s) in {self.sounds_dir}"
        threading.Thread(target=self._probe, daemon=True).start()
        return len(sounds)

    def _probe(self):
        ffprobe = self._ffprobe()
        if not ffprobe:
            return
        for word, path in list(self.sounds.items()):
            if word in self.durations:
                continue
            try:
                out = subprocess.run([ffprobe, '-v', 'error', '-show_entries', 'format=duration',
                                      '-of', 'default=noprint_wrappers=1:nokey=1', path],
                                     capture_output=True, timeout=10, creationflags=NO_WINDOW,
                                     stdin=subprocess.DEVNULL)
                self.durations[word] = float(out.stdout.decode().strip())
            except Exception:
                self.durations[word] = None
        self.version += 1

    # ------------------------------------------------------------------ chat
    def match(self, message):
        """(word, pitch) if this chat message triggers a sound in the current mode, else None"""
        mode = self.get('mode')
        if mode == 'off':
            return None
        text = message.strip().lower()
        if mode == 'command':
            command = (self.get('command') or '').strip().lower()
            if not command or not text.startswith(command + ' '):
                return None
            text = text[len(command):].strip()
        m = re.fullmatch(r'(\S+)(?:\s+(\d{1,3}))?', text)
        if not m or m.group(1) not in self.sounds:
            return None
        pitch = 100
        if m.group(2):
            if not self.get('pitch_enabled'):
                return None if mode == 'open' else (m.group(1), 100)  # "lol 5" in open chat is just chat
            pitch = max(50, min(200, int(m.group(2))))
        return m.group(1), pitch

    def handle_chat(self, username, message, is_admin, private_mode=False):
        """Called for each chat line. True if it was a sound trigger (whether or not it played)."""
        found = self.match(message)
        if not found:
            return False
        word, pitch = found
        if (self.get('who') == 'admins' or private_mode) and not is_admin:
            return True
        now = time.time()
        with self._lock:
            if now - self._last_any < float(self.get('global_gap')):
                return True
            if not (is_admin and self.get('admins_skip_cooldown')):
                if now - self._last_by_player.get(username, 0) < float(self.get('cooldown')):
                    return True
            if len(self._active) >= int(self.get('max_at_once')):
                return True
            self._last_any = now
            self._last_by_player[username] = now
        self.play(word, pitch, by=username)
        return True

    # ------------------------------------------------------------------ playback
    def play(self, word, pitch=100, by=None):
        """Start a sound right away on its own stream (overlaps whatever else is playing)"""
        path = self.sounds.get(word)
        if not path:
            return False
        stop_flag = threading.Event()
        with self._lock:
            self._active.append(stop_flag)
        threading.Thread(target=self._play_thread, args=(word, path, pitch, stop_flag, by), daemon=True).start()
        return True

    def stop_all(self):
        with self._lock:
            for flag in self._active:
                flag.set()

    def is_playing(self):
        return bool(self._active)

    def _decode(self, word, path, pitch):
        bass = float(self.get('bass_db') or 0)
        key = (word, pitch, bass)
        with self._lock:
            if key in self._cache:
                self._cache.move_to_end(key)
                return self._cache[key]
        ffmpeg = self._ffmpeg()
        if not ffmpeg:
            raise RuntimeError("ffmpeg isn't installed yet - click Download ffmpeg (Soundboard or DJ window)")
        cmd = [ffmpeg, '-hide_banner', '-loglevel', 'error', '-nostdin', '-i', path, '-vn',
               '-t', str(int(self.get('max_seconds')))]
        filters = []
        if pitch != 100:
            filters.append(f'aresample={RATE},asetrate={int(RATE * pitch / 100)},aresample={RATE}')
        if bass:
            filters.append(f'bass=g={bass:g}:f=160:width_type=o:w=1')  # low shelf below ~160 Hz
        if filters:
            cmd += ['-af', ','.join(filters)]
        cmd += ['-f', 's16le', '-acodec', 'pcm_s16le', '-ac', str(CHANNELS), '-ar', str(RATE), 'pipe:1']
        out = subprocess.run(cmd, capture_output=True, timeout=20, creationflags=NO_WINDOW,
                             stdin=subprocess.DEVNULL)
        if out.returncode != 0 or not out.stdout:
            lines = out.stderr.decode('utf-8', 'replace').strip().splitlines()
            raise RuntimeError(lines[-1] if lines else 'decode failed')
        pcm = out.stdout
        with self._lock:
            self._cache[key] = pcm
            while len(self._cache) > 40:
                self._cache.popitem(last=False)
        return pcm

    def _play_thread(self, word, path, pitch, stop_flag, by):
        stream = None
        try:
            if pyaudio is None:
                raise RuntimeError("PyAudio is missing")
            pcm = self._decode(word, path, pitch)
            with self._pa_lock:
                if self._pa is None:
                    self._pa = pyaudio.PyAudio()
                index, _ = find_output_device(self._pa, self._tts_device())
                stream = self._pa.open(format=pyaudio.paInt16, channels=CHANNELS, rate=RATE, output=True,
                                       output_device_index=index, frames_per_buffer=CHUNK_FRAMES)
            step = CHUNK_FRAMES * CHANNELS * 2
            for start in range(0, len(pcm), step):
                if stop_flag.is_set():
                    break
                stream.write(self._scale(pcm[start:start + step]))
            self.status = f"Played {word}" + (f" ({pitch})" if pitch != 100 else "") + (f" for {by}" if by else "")
        except Exception as e:
            self.status = f"Couldn't play {word}: {e}"
            logger.error(f"Soundboard: {word}: {e}")
        finally:
            if stream is not None:
                try:
                    stream.stop_stream()
                    stream.close()
                except Exception:
                    pass
            with self._lock:
                if stop_flag in self._active:
                    self._active.remove(stop_flag)

    def _scale(self, data):
        gain = self.get('volume') / 100.0
        data = data[:len(data) // 2 * 2]
        if gain == 1.0:
            return data
        if np is not None:
            samples = np.frombuffer(data, dtype=np.int16).astype(np.float32) * gain
            return np.clip(samples, -32768, 32767).astype(np.int16).tobytes()
        import array
        samples = array.array('h', data)
        for i in range(len(samples)):
            samples[i] = max(-32768, min(32767, int(samples[i] * gain)))
        return samples.tobytes()

    def shutdown(self):
        self.stop_all()
        time.sleep(0.1)
        with self._pa_lock:
            if self._pa:
                try:
                    self._pa.terminate()
                except Exception:
                    pass
                self._pa = None


class SoundboardWindow:
    """Soundboard settings + sound list. GUI thread only; reads the engine on a timer."""

    def __init__(self, root, engine, on_close=None, tools=None):
        """tools: the DJ engine, which owns the ffmpeg download (shared by the DJ and the soundboard)"""
        self.engine = engine
        self._on_close = on_close
        self._tools = tools
        self._seen = -1
        self._poll_job = None
        e = engine

        self.win = tk.Toplevel(root)
        self.win.title("Soundboard")
        if ui_theme:
            ui_theme.size_window(self.win, 900, 700, root, minsize=(640, 560))
        else:
            self.win.geometry("760x700")
            self.win.minsize(640, 560)
        self.win.protocol("WM_DELETE_WINDOW", self.close)

        how = ttk.LabelFrame(self.win, text="How players trigger sounds")
        how.pack(fill=tk.X, padx=10, pady=(10, 5))
        row = ttk.Frame(how)
        row.pack(fill=tk.X, padx=6, pady=(6, 2))
        self.mode_var = tk.StringVar(value=e.get('mode'))
        for value, text in (('off', 'Off'), ('open', 'Just the word  (yee)'), ('command', 'Command + word')):
            ttk.Radiobutton(row, text=text, value=value, variable=self.mode_var,
                            command=lambda: e.set('mode', self.mode_var.get())).pack(side=tk.LEFT, padx=(0, 10))
        self.command_var = tk.StringVar(value=e.get('command'))
        command_entry = ttk.Entry(row, textvariable=self.command_var, width=8)
        command_entry.pack(side=tk.LEFT)
        command_entry.bind('<FocusOut>', lambda ev: self._save_command())
        command_entry.bind('<Return>', lambda ev: self._save_command())
        self.example_var = tk.StringVar()
        ttk.Label(row, textvariable=self.example_var, style='Hint.TLabel').pack(side=tk.LEFT, padx=6)

        row = ttk.Frame(how)
        row.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(row, text="Who:").pack(side=tk.LEFT)
        self.who_var = tk.StringVar(value=e.get('who'))
        for value, text in (('admins', 'Admins only'), ('everyone', 'Everyone')):
            ttk.Radiobutton(row, text=text, value=value, variable=self.who_var,
                            command=lambda: e.set('who', self.who_var.get())).pack(side=tk.LEFT, padx=4)
        ttk.Label(row, text="(blocked and muted players never can; Private Mode = admins only)",
                  style='Hint.TLabel').pack(side=tk.LEFT, padx=6)

        row = ttk.Frame(how)
        row.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(row, text="Cooldown per player").pack(side=tk.LEFT)
        self.cooldown_var = tk.StringVar(value=str(e.get('cooldown')))
        self._spin(row, self.cooldown_var, 0, 600, 1, lambda: self._save_number('cooldown', self.cooldown_var, float))
        ttk.Label(row, text="s    Gap between any two sounds").pack(side=tk.LEFT)
        self.gap_var = tk.StringVar(value=str(e.get('global_gap')))
        self._spin(row, self.gap_var, 0, 60, 0.5, lambda: self._save_number('global_gap', self.gap_var, float))
        ttk.Label(row, text="s").pack(side=tk.LEFT)
        self.admin_skip_var = tk.BooleanVar(value=e.get('admins_skip_cooldown'))
        ttk.Checkbutton(row, text="Admins skip the cooldown", variable=self.admin_skip_var,
                        command=lambda: e.set('admins_skip_cooldown', self.admin_skip_var.get())).pack(side=tk.LEFT, padx=12)

        row = ttk.Frame(how)
        row.pack(fill=tk.X, padx=6, pady=(2, 6))
        self.pitch_var = tk.BooleanVar(value=e.get('pitch_enabled'))
        ttk.Checkbutton(row, text='Allow pitch ("yee 150" = faster/higher, 50-200)', variable=self.pitch_var,
                        command=lambda: e.set('pitch_enabled', self.pitch_var.get())).pack(side=tk.LEFT)
        self.duck_var = tk.BooleanVar(value=e.get('duck_music'))
        ttk.Checkbutton(row, text="Lower the DJ music while a sound plays (off = play over it)",
                        variable=self.duck_var,
                        command=lambda: e.set('duck_music', self.duck_var.get())).pack(side=tk.LEFT, padx=12)

        row = ttk.Frame(self.win)
        row.pack(fill=tk.X, padx=10, pady=2)
        ttk.Label(row, text="Sound volume").pack(side=tk.LEFT)
        self.volume_var = tk.IntVar(value=e.get('volume'))
        ttk.Scale(row, from_=0, to=100, variable=self.volume_var, orient=tk.HORIZONTAL, length=200,
                  command=lambda v: self._set_volume()).pack(side=tk.LEFT, padx=6)
        self.volume_label = ttk.Label(row, text=f"{self.volume_var.get()}%", width=5)
        self.volume_label.pack(side=tk.LEFT)
        ttk.Label(row, text="  Bass").pack(side=tk.LEFT)
        self.bass_var = tk.StringVar(value=str(int(e.get('bass_db'))))
        self._spin(row, self.bass_var, -15, 6, 1, lambda: self._save_bass())
        ttk.Label(row, text="dB").pack(side=tk.LEFT)
        ttk.Label(row, text="  At once (max)").pack(side=tk.LEFT)
        self.max_var = tk.StringVar(value=str(e.get('max_at_once')))
        self._spin(row, self.max_var, 1, 10, 1, lambda: self._save_number('max_at_once', self.max_var, int))

        sounds = ttk.LabelFrame(self.win, text="Sounds")
        sounds.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        row = ttk.Frame(sounds)
        row.pack(fill=tk.X, padx=6, pady=(6, 2))
        ttk.Label(row, text="Find:").pack(side=tk.LEFT)
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add('write', lambda *a: self._redraw())
        ttk.Entry(row, textvariable=self.filter_var, width=24).pack(side=tk.LEFT, padx=4)
        self.count_var = tk.StringVar()
        ttk.Label(row, textvariable=self.count_var).pack(side=tk.RIGHT)

        tree_frame = ttk.Frame(sounds)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=2)
        self.tree = ttk.Treeview(tree_frame, columns=('word', 'length', 'file'), show='headings', height=12)
        for col, text, width in (('word', 'Word', 160), ('length', 'Length', 70), ('file', 'File', 420)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, stretch=(col == 'file'), anchor=tk.E if col == 'length' else tk.W)
        scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind('<Double-1>', lambda ev: self._preview())

        row = ttk.Frame(sounds)
        row.pack(fill=tk.X, padx=6, pady=(2, 6))
        ttk.Button(row, text="▶ Play", command=self._preview).pack(side=tk.LEFT, padx=2)
        ttk.Button(row, text="⏹ Stop sounds", command=e.stop_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(row, text="📂 Open sounds folder", command=self._open_folder).pack(side=tk.LEFT, padx=(12, 2))
        ttk.Button(row, text="⟳ Reload", command=e.reload).pack(side=tk.LEFT, padx=2)
        ttk.Label(sounds, style='Hint.TLabel', justify=tk.LEFT, text=(
            "Put sound files in the sounds folder - the file name is the word (yee.wav -> yee). Sub-folders work too.\n"
            "A Sven ChatSounds list (chatsounds.txt, lines like: bleh twlz/bleh.wav) in that folder is read as well."
        )).pack(anchor=tk.W, padx=6, pady=(0, 6))

        self.ffmpeg_row = ttk.Frame(self.win)
        self.ffmpeg_var = tk.StringVar()
        ttk.Label(self.ffmpeg_row, text="ffmpeg is needed to play sounds:").pack(side=tk.LEFT)
        ttk.Label(self.ffmpeg_row, textvariable=self.ffmpeg_var, style='Hint.TLabel').pack(side=tk.LEFT, padx=6)
        if self._tools:
            self.ffmpeg_btn = ttk.Button(self.ffmpeg_row, text="Download ffmpeg", command=self._tools.update_ffmpeg)
            self.ffmpeg_btn.pack(side=tk.LEFT, padx=4)

        self.status_var = tk.StringVar()
        ttk.Label(self.win, textvariable=self.status_var, style='Hint.TLabel').pack(fill=tk.X, padx=10, pady=(0, 8))
        self._poll()

    def _spin(self, parent, var, low, high, step, save):
        box = ttk.Spinbox(parent, from_=low, to=high, increment=step, width=5, textvariable=var, command=save)
        box.pack(side=tk.LEFT, padx=4)
        box.bind('<FocusOut>', lambda ev: save())
        return box

    def lift(self):
        self.win.deiconify()
        self.win.lift()
        self.win.focus_force()

    def close(self):
        if self._poll_job:
            self.win.after_cancel(self._poll_job)
        self.win.destroy()
        if self._on_close:
            self._on_close()

    def _save_command(self):
        command = self.command_var.get().strip()
        if command:
            self.engine.set('command', command)
        else:
            self.command_var.set(self.engine.get('command'))

    def _save_number(self, key, var, kind):
        try:
            value = kind(var.get())
            if value < 0:
                raise ValueError
            self.engine.set(key, value)
        except ValueError:
            var.set(str(self.engine.get(key)))

    def _save_bass(self):
        try:
            self.engine.set('bass_db', max(-15, min(6, int(float(self.bass_var.get())))))
        except ValueError:
            self.bass_var.set(str(int(self.engine.get('bass_db'))))

    def _set_volume(self):
        value = int(float(self.volume_var.get()))
        self.volume_label.config(text=f"{value}%")
        self.engine.settings['volume'] = value
        if getattr(self, '_volume_job', None):
            self.win.after_cancel(self._volume_job)
        self._volume_job = self.win.after(500, lambda: self.engine.set('volume', value))

    def _preview(self):
        for item in self.tree.selection()[:1]:
            self.engine.play(self.tree.item(item, 'values')[0])

    def _open_folder(self):
        try:
            self.engine.sounds_dir.mkdir(parents=True, exist_ok=True)
            os.startfile(str(self.engine.sounds_dir))
        except Exception as ex:
            self.engine.status = f"Couldn't open the folder: {ex}"

    def _redraw(self):
        self._seen = self.engine.version
        wanted = self.filter_var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for word, path in sorted(self.engine.sounds.items()):
            if wanted and wanted not in word:
                continue
            self.tree.insert('', tk.END, values=(word, fmt_time(self.engine.durations.get(word)) if word in self.engine.durations else '',
                                                 os.path.relpath(path, self.engine.sounds_dir)))
            shown += 1
        total = len(self.engine.sounds)
        self.count_var.set(f"{shown} of {total} sound(s)" if wanted else f"{total} sound(s)")

    def _poll(self):
        try:
            if self.engine.version != self._seen:
                self._redraw()
            mode = self.engine.get('mode')
            command = self.engine.get('command')
            self.example_var.set({'off': "sounds are off", 'open': 'players type: yee   or   yee 150',
                                  'command': f'players type: {command} yee   or   {command} yee 150'}[mode])
            # ffmpeg row only while it's missing or downloading
            busy = bool(self._tools and self._tools.ffmpeg_busy)
            missing = not self.engine._ffmpeg()
            if busy or missing:
                self.ffmpeg_var.set("downloading..." if busy else "not installed (one-time ~80 MB download)")
                if not self.ffmpeg_row.winfo_ismapped():
                    self.ffmpeg_row.pack(fill=tk.X, padx=10, pady=2, before=self.win.pack_slaves()[-1])
                if self._tools:
                    self.ffmpeg_btn.configure(state=tk.DISABLED if busy else tk.NORMAL)
            elif self.ffmpeg_row.winfo_ismapped():
                self.ffmpeg_row.pack_forget()
            self.status_var.set(self._tools.status if busy and self._tools else self.engine.status)
        except tk.TclError:
            return
        self._poll_job = self.win.after(300, self._poll)
