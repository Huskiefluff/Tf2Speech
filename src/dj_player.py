"""
DJ player for the TTS: plays local music and YouTube audio into the same output as the TTS (VoiceMeeter),
with a queue, ducking under TTS speech, loudness levelling, "now playing" announcements and chat song requests.

Audio path: ffmpeg decodes anything (mp3/flac/m4a/ogg/wav, or a YouTube audio stream found by yt-dlp) to raw
48 kHz stereo PCM on a pipe -> volume and ducking are applied here -> PyAudio plays it on the output device.
Nothing is written to disk while playing (no temp files); yt-dlp runs with --no-cache-dir.

DJEngine runs the player in its own thread and never touches Tk. DJWindow is the Tk window; it only reads the
engine's state on a timer (every 250 ms), so no Tk calls ever come from a background thread.
"""

import array
import logging
import os
import random
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.request
import zipfile
from collections import deque
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import numpy as np
except ImportError:  # slower pure-Python volume fallback below
    np = None

try:
    import pyaudio
except ImportError:
    pyaudio = None

try:
    from dectalk_native import find_output_device
except Exception:  # same matching rules, kept here so the DJ still works if DECtalk can't load
    def find_output_device(pa, device_name):
        if not device_name or device_name.strip().lower() == 'default':
            return None, None
        wanted = device_name.lower().replace('(recommended)', '').strip()
        outputs = [(i, pa.get_device_info_by_index(i)['name']) for i in range(pa.get_device_count())
                   if pa.get_device_info_by_index(i).get('maxOutputChannels', 0) > 0]
        for i, name in outputs:
            if name.lower().strip() == wanted:
                return i, name
        for i, name in outputs:
            if wanted in name.lower():
                return i, name
        for i, name in outputs:
            if len(name) >= 12 and wanted.startswith(name.lower().strip()):
                return i, name
        return None, None

logger = logging.getLogger(__name__)

RATE = 48000
CHANNELS = 2
FRAMES_PER_CHUNK = 2048                       # ~43 ms of audio per write
CHUNK_BYTES = FRAMES_PER_CHUNK * CHANNELS * 2  # 16-bit samples
NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
AUDIO_EXTENSIONS = {'.mp3', '.flac', '.m4a', '.aac', '.ogg', '.opus', '.wav', '.wma', '.webm', '.mp4', '.mka', '.aiff'}
# Chat requests may only be YouTube links (or search words). Any other link would make this PC fetch a page a
# stranger chose - an easy way for someone to grab your IP address.
YOUTUBE_URL = re.compile(r'^https?://(?:www\.|m\.|music\.)?(?:youtube\.com/(?:watch\?|shorts/|live/)|youtu\.be/)\S+$', re.I)
YTDLP_DOWNLOAD_URL = "https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe"
# ffmpeg, LGPL build from BtbN (the Windows builds linked from ffmpeg.org). The "shared" build is an ~80 MB
# download; ffmpeg.exe + ffprobe.exe + their DLLs (~160 MB) are unpacked into an "ffmpeg" folder next to the
# TTS, nothing is installed, and the zip is deleted afterwards.
FFMPEG_DOWNLOAD_URL = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-lgpl-shared.zip"
FFMPEG_MISSING = "ffmpeg isn't installed yet - click Download ffmpeg in the DJ window (one-time, about 80 MB)"

DEFAULTS = {
    'volume': 40,                 # %
    'duck_enabled': True,
    'duck_level': 25,             # % of the music volume while the TTS talks
    'normalize': True,            # even out loudness between songs (ffmpeg dynaudnorm)
    'announce': False,            # "Now playing ..." through the TTS
    'announce_voice': 'Default',  # 'Default' = the announcement voice from the Announcements tab
    'requests': 'off',            # chat !sr: 'off' | 'admins' | 'everyone'
    'request_max_minutes': 8,     # 0 = no limit (admins are never limited)
    'request_limit': 1,           # songs one non-admin may have waiting in the queue
    'shuffle': False,
    'repeat': 'off',              # 'off' | 'all' | 'one'
    'output_device': '',          # '' = same device as the TTS
}


def fmt_time(seconds):
    """1:05, 12:34, 1:02:03 - or ? when unknown"""
    if seconds is None:
        return '?'
    seconds = int(max(0, seconds))
    h, rest = divmod(seconds, 3600)
    m, s = divmod(rest, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def clean_title(title):
    """Drop the usual YouTube clutter so "now playing" reads well: (Official Video), [4K], | Lyrics ..."""
    junk = r'official|video|audio|lyrics?|visuali[sz]er|hd|hq|4k|1080p|remaster(?:ed)?|mv|music video|explicit'
    title = re.sub(r'\s*[\(\[][^\)\]]*\b(?:' + junk + r')\b[^\)\]]*[\)\]]', '', title, flags=re.I)
    title = re.sub(r'\s*\|.*$', '', title)
    return re.sub(r'\s{2,}', ' ', title).strip() or title


def make_track(title, source, target, duration=None, requested_by=None):
    return {'title': title, 'source': source, 'target': target,
            'duration': duration, 'requested_by': requested_by}


class DJEngine:
    """Queue + player. All public methods are safe to call from any thread."""

    def __init__(self, settings, save_config, tts_active=None, announce=None, tts_device=None, app_dir=None):
        """
        settings:     the config['dj'] dict (read/written in place; save_config() writes the whole config)
        tts_active:   () -> bool, True while the TTS is speaking (for ducking)
        announce:     (text, track) -> None, speaks "now playing" through the TTS
        tts_device:   () -> str, the TTS's output device name (used when 'output_device' is '')
        app_dir:      folder where yt-dlp.exe is kept (next to the TTS exe)
        """
        self.settings = settings
        self._save_config = save_config
        self._tts_active = tts_active or (lambda: False)
        self._announce = announce or (lambda text, track: None)
        self._tts_device = tts_device or (lambda: 'Default')
        self.app_dir = Path(app_dir) if app_dir else Path.cwd()

        self._lock = threading.RLock()
        self.queue = [dict(make_track(**{k: t.get(k) for k in ('title', 'source', 'target', 'duration', 'requested_by')}))
                      for t in settings.get('queue', []) if isinstance(t, dict) and t.get('target')]
        self.history = []          # tracks finished this session, newest last (for Previous)
        self.current = None
        self.state = 'stopped'     # stopped | loading | playing | paused
        self.position = 0.0
        self.status = ''           # last message for the window
        self.queue_version = 0     # bumps whenever the queue changes so the window redraws it
        self.search_results = []
        self.search_version = 0
        self.search_busy = False
        self.ytdlp_busy = False
        self.ffmpeg_busy = False

        self._play_next_track = None   # "Play Now" bypasses shuffle
        self._new_track = False        # announce "now playing" once per track, not after seeks
        self._command = None           # skip | stop | seek | prev | shutdown, picked up by the player thread
        self._seek_to = 0.0
        self._gain = 0.0
        self._proc = None              # running ffmpeg or yt-dlp, killed to interrupt a blocking read
        self._pa = None
        self._wake = threading.Event()
        self._remove_partial_downloads()
        self._thread = threading.Thread(target=self._player_loop, daemon=True, name='DJ player')
        self._thread.start()

    def _remove_partial_downloads(self):
        """Leftovers from a download that was cut off (TTS closed or crashed mid-download)"""
        for name in ('ffmpeg-download.zip', 'yt-dlp.exe.part'):
            try:
                (self.app_dir / name).unlink()
            except OSError:
                pass
        for name in ('ffmpeg.new', 'ffmpeg.old'):
            if (self.app_dir / name).exists():
                if name == 'ffmpeg.old' and not (self.app_dir / 'ffmpeg').exists():
                    try:
                        os.replace(self.app_dir / name, self.app_dir / 'ffmpeg')  # the swap was interrupted
                        continue
                    except OSError:
                        pass
                shutil.rmtree(self.app_dir / name, ignore_errors=True)

    # ------------------------------------------------------------------ settings
    def get(self, key):
        return self.settings.get(key, DEFAULTS.get(key))

    def set(self, key, value):
        self.settings[key] = value
        self._save()

    def _save(self):
        try:
            self._save_config()
        except Exception as e:
            logger.error(f"DJ: could not save settings: {e}")

    def _queue_changed(self, save=True):
        self.queue_version += 1
        if save:
            self.settings['queue'] = [dict(t) for t in self.queue[:1000]]
            self._save()

    # ------------------------------------------------------------------ tools
    def ffmpeg_path(self):
        return self._tool('ffmpeg')

    def ffprobe_path(self):
        ffmpeg = self.ffmpeg_path()
        if ffmpeg:
            sibling = Path(ffmpeg).with_name('ffprobe.exe' if sys.platform == 'win32' else 'ffprobe')
            if sibling.is_file():
                return str(sibling)
        return self._tool('ffprobe')

    def ytdlp_path(self):
        return self._tool('yt-dlp')

    def _tool(self, name):
        """Configured path, then next to the TTS exe (or in its ffmpeg folder), then PATH"""
        custom = self.settings.get(f"{name.replace('-', '')}_path")
        exe = f"{name}.exe" if sys.platform == 'win32' else name
        for candidate in (custom, str(self.app_dir / exe), str(self.app_dir / 'ffmpeg' / exe)):
            if candidate and os.path.isfile(candidate):
                return candidate
        return shutil.which(name)

    def update_ffmpeg(self):
        """Download ffmpeg (LGPL build) into the "ffmpeg" folder next to the TTS - or replace it with the latest.
        Runs in the background; progress shows in self.status. Leaves nothing behind but that folder."""
        if self.ffmpeg_busy:
            return
        if self.state != 'stopped':
            self.status = "Stop the music first, then download/update ffmpeg"
            return

        def work():
            self.ffmpeg_busy = True
            zip_path = self.app_dir / 'ffmpeg-download.zip'
            staging = self.app_dir / 'ffmpeg.new'
            target = self.app_dir / 'ffmpeg'
            old = self.app_dir / 'ffmpeg.old'
            try:
                self.status = "Downloading ffmpeg..."
                with urllib.request.urlopen(FFMPEG_DOWNLOAD_URL, timeout=60) as response, open(zip_path, 'wb') as f:
                    total = int(response.headers.get('Content-Length') or 0)
                    done = 0
                    while True:
                        block = response.read(1 << 20)
                        if not block:
                            break
                        f.write(block)
                        done += len(block)
                        self.status = (f"Downloading ffmpeg... {done / 1e6:.0f} of {total / 1e6:.0f} MB" if total
                                       else f"Downloading ffmpeg... {done / 1e6:.0f} MB")
                self.status = "Unpacking ffmpeg..."
                shutil.rmtree(staging, ignore_errors=True)
                staging.mkdir(parents=True)
                with zipfile.ZipFile(zip_path) as z:
                    for info in z.infolist():
                        name = info.filename.rsplit('/', 1)[-1]
                        wanted = ('/bin/' in info.filename and (name.lower().endswith('.dll')
                                                                 or name.lower() in ('ffmpeg.exe', 'ffprobe.exe'))) \
                            or name.upper() == 'LICENSE.TXT'
                        if wanted and not info.is_dir():
                            with z.open(info) as src, open(staging / name, 'wb') as dst:
                                shutil.copyfileobj(src, dst)
                if not (staging / 'ffmpeg.exe').is_file():
                    raise RuntimeError("the download didn't contain ffmpeg.exe")
                subprocess.run([str(staging / 'ffmpeg.exe'), '-hide_banner', '-version'], capture_output=True,
                               timeout=30, check=True, creationflags=NO_WINDOW, stdin=subprocess.DEVNULL)
                shutil.rmtree(old, ignore_errors=True)
                if target.exists():
                    os.replace(target, old)  # fails if a sound is playing from it right now
                os.replace(staging, target)
                shutil.rmtree(old, ignore_errors=True)
                self.status = f"ffmpeg is ready ({target})"
            except Exception as e:
                self.status = f"ffmpeg download failed: {e}"
                if old.exists() and not target.exists():
                    try:
                        os.replace(old, target)  # put the previous copy back
                    except OSError:
                        pass
            finally:
                try:
                    zip_path.unlink()
                except OSError:
                    pass
                shutil.rmtree(staging, ignore_errors=True)
                self.ffmpeg_busy = False
        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------------ queue editing
    def add_tracks(self, tracks, front=False):
        with self._lock:
            if front:
                self.queue[0:0] = tracks
            else:
                self.queue.extend(tracks)
            self._queue_changed()

    def add_files(self, paths):
        """Add audio files now (named from the file); lengths and tags are filled in by a background probe"""
        tracks = [make_track(Path(p).stem, 'file', str(p)) for p in paths
                  if Path(p).suffix.lower() in AUDIO_EXTENSIONS]
        if not tracks:
            self.status = "No audio files there"
            return 0
        self.add_tracks(tracks)
        threading.Thread(target=self._probe_files, args=(tracks,), daemon=True).start()
        self.status = f"Added {len(tracks)} file(s)"
        return len(tracks)

    def add_folder(self, folder):
        files = sorted((str(p) for p in Path(folder).rglob('*') if p.suffix.lower() in AUDIO_EXTENSIONS),
                       key=str.lower)
        return self.add_files(files)

    def _probe_files(self, tracks):
        ffprobe = self.ffprobe_path()
        if not ffprobe:
            return
        for track in tracks:
            try:
                out = subprocess.run(
                    [ffprobe, '-v', 'error', '-show_entries', 'format=duration:format_tags=title,artist',
                     '-of', 'default=noprint_wrappers=1', track['target']],
                    capture_output=True, timeout=15, creationflags=NO_WINDOW, stdin=subprocess.DEVNULL)
                fields = {}
                for line in out.stdout.decode('utf-8', 'replace').splitlines():
                    key, _, value = line.partition('=')
                    fields[key.strip().lower().replace('tag:', '')] = value.strip()
                if fields.get('duration', 'N/A') not in ('N/A', ''):
                    track['duration'] = float(fields['duration'])
                if fields.get('title'):
                    track['title'] = f"{fields['artist']} - {fields['title']}" if fields.get('artist') else fields['title']
            except Exception as e:
                logger.debug(f"DJ: probe failed for {track['target']}: {e}")
            self.queue_version += 1  # redraw as lengths come in
        with self._lock:
            self._queue_changed()

    def remove(self, indices):
        with self._lock:
            for i in sorted(set(indices), reverse=True):
                if 0 <= i < len(self.queue):
                    del self.queue[i]
            self._queue_changed()

    def move(self, index, delta):
        with self._lock:
            new = index + delta
            if 0 <= index < len(self.queue) and 0 <= new < len(self.queue):
                self.queue[index], self.queue[new] = self.queue[new], self.queue[index]
                self._queue_changed()
                return new
        return index

    def clear(self):
        with self._lock:
            self.queue.clear()
            self._queue_changed()

    def play_now(self, index):
        with self._lock:
            if not (0 <= index < len(self.queue)):
                return
            self._play_next_track = self.queue.pop(index)
            self._queue_changed()
            if self.state == 'stopped':
                self.play()
            else:
                self._send('skip')

    # ------------------------------------------------------------------ transport
    def play(self):
        with self._lock:
            if self.state == 'paused':
                self.state = 'playing'
                return
            if self.state != 'stopped':
                return
            if not (self.queue or self.current or self._play_next_track):
                self.status = "The queue is empty - add some music first"
                return
            self.state = 'loading'
            self._command = None
            self._wake.set()

    def pause(self):
        with self._lock:
            if self.state in ('playing', 'loading'):
                self.state = 'paused'

    def toggle(self):
        if self.state in ('stopped', 'paused'):
            self.play()
        else:
            self.pause()

    def skip(self):
        if self.state != 'stopped':
            self._send('skip')

    def stop(self):
        if self.state != 'stopped':
            self._send('stop')

    def seek(self, seconds):
        if self.state != 'stopped' and self.current:
            self._seek_to = max(0.0, float(seconds))
            self._send('seek')

    def previous(self):
        """Back to the start of this song, or (in the first 3 seconds) the song before it"""
        if self.state == 'stopped':
            return
        if self.position > 3 or not self.history:
            self.seek(0)
        else:
            self._send('prev')

    def _send(self, command):
        self._command = command
        self._kill_proc()   # unblocks a read/resolve so the player thread sees the command right away
        self._wake.set()

    def _kill_proc(self):
        proc = self._proc
        if proc and proc.poll() is None:
            try:
                proc.kill()
            except Exception:
                pass

    def shutdown(self):
        self._send('shutdown')
        self._thread.join(timeout=3)
        with self._lock:
            # whatever was playing goes back to the front of the saved queue for next time
            if self.current:
                self.queue.insert(0, self.current)
                self.current = None
            self._queue_changed()
        if self._pa:
            try:
                self._pa.terminate()
            except Exception:
                pass

    # ------------------------------------------------------------------ player thread
    def _take_next(self):
        if self._play_next_track:
            track, self._play_next_track = self._play_next_track, None
            return track
        if not self.queue:
            return None
        index = random.randrange(len(self.queue)) if self.get('shuffle') else 0
        track = self.queue.pop(index)
        self._queue_changed()
        return track

    def _player_loop(self):
        failures = 0
        while True:
            self._wake.wait()
            self._wake.clear()
            while True:
                if self._command == 'shutdown':
                    return
                with self._lock:
                    if self.state == 'stopped':
                        break
                    if self.current is None:
                        self.current = self._take_next()
                        self._new_track = True
                        if self.current is None:
                            self.state, self.position = 'stopped', 0.0
                            self.status = "Queue finished"
                            break
                    track = self.current
                    start, self._seek_to = self._seek_to, 0.0
                    self._command = None
                    if self.state != 'paused':
                        self.state = 'loading'
                try:
                    outcome = self._play_track(track, start)
                except Exception as e:
                    logger.error(f"DJ: playback crashed: {e}")
                    self.status = f"Playback error: {e}"
                    outcome = 'error'
                with self._lock:
                    if outcome == 'shutdown':
                        return
                    if outcome == 'seek':
                        continue  # same song, new start position
                    if outcome == 'stop':
                        self.queue.insert(0, track)  # Play starts it again
                        self.current = None
                        self.state, self.position = 'stopped', 0.0
                        self._queue_changed()
                        break
                    if outcome == 'prev':
                        self.queue.insert(0, track)
                        if self.history:
                            self.queue.insert(0, self.history.pop())
                        self.current = None
                        self._queue_changed()
                        continue
                    failures = failures + 1 if outcome == 'error' else 0
                    if failures >= 3:
                        self.status = "Stopped after 3 songs in a row failed to play - " + self.status
                        self.current = None
                        self.state, self.position = 'stopped', 0.0
                        failures = 0
                        break
                    if outcome == 'ended' and self.get('repeat') == 'one':
                        continue  # same track again from the top
                    self.history.append(track)
                    del self.history[:-50]
                    if self.get('repeat') == 'all' and outcome in ('ended', 'skip'):
                        self.queue.append(dict(track))
                        self._queue_changed()
                    self.current = None

    def _play_track(self, track, start):
        """Play one track from start seconds. Returns ended | skip | stop | seek | prev | shutdown | error."""
        ffmpeg = self.ffmpeg_path()
        if not ffmpeg:
            self.status = FFMPEG_MISSING
            return 'stop'
        if pyaudio is None:
            self.status = "PyAudio is missing, so the DJ can't play audio"
            return 'stop'

        self.position = start
        source = track['target']
        if track['source'] == 'youtube':
            self.status = f"Loading {track['title']} from YouTube..."
            info, error = self._youtube_info(source, stream_url=True)
            if self._command:
                return self._command
            if not info:
                self.status = f"Couldn't load {track['title']}: {error}"
                return 'error'
            source = info['url']
            if not track.get('duration') and info.get('duration'):
                track['duration'] = info['duration']
        elif not os.path.isfile(source):
            self.status = f"File not found: {source}"
            return 'error'

        cmd = [ffmpeg, '-hide_banner', '-loglevel', 'error', '-nostdin']
        if track['source'] == 'youtube':
            cmd += ['-reconnect', '1', '-reconnect_streamed', '1', '-reconnect_delay_max', '5']
        if start > 0:
            cmd += ['-ss', f'{start:.2f}']
        cmd += ['-i', source, '-vn', '-sn', '-dn']
        if self.get('normalize'):
            cmd += ['-af', 'dynaudnorm=f=200:g=11:p=0.9']
        cmd += ['-f', 's16le', '-acodec', 'pcm_s16le', '-ac', str(CHANNELS), '-ar', str(RATE), 'pipe:1']

        errors = deque(maxlen=6)
        try:
            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
        except OSError as e:
            self.status = f"Couldn't start ffmpeg: {e}"
            return 'stop'
        self._proc = proc

        def drain_errors():  # keep ffmpeg's error output from filling its pipe and stalling it
            for line in proc.stderr:
                errors.append(line.decode('utf-8', 'replace').strip())
        threading.Thread(target=drain_errors, daemon=True).start()

        stream = None
        frames = 0
        try:
            try:
                stream = self._open_stream()
            except Exception as e:
                self.status = f"Couldn't open the audio output: {e}"
                return 'stop'
            self._gain = 0.0  # short fade-in
            while True:
                if self._command:
                    return self._command
                if self.state == 'paused':
                    if stream.is_active():
                        stream.stop_stream()
                    time.sleep(0.05)
                    continue
                if not stream.is_active():
                    stream.start_stream()
                data = proc.stdout.read(CHUNK_BYTES)
                if self._command:
                    return self._command
                if not data:
                    break
                if self.state == 'loading':
                    self.state = 'playing'
                    self.status = ''
                if self._new_track:  # first sound of a new song (not after a seek)
                    self._new_track = False
                    self._announce_now_playing(track)
                stream.write(self._apply_gain(data))
                frames += len(data) // (CHANNELS * 2)
                self.position = start + frames / RATE
            proc.wait(timeout=5)
            if frames == 0 and proc.returncode not in (0, None):
                self.status = f"Couldn't play {track['title']}: {errors[-1] if errors else 'ffmpeg error'}"
                return 'error'
            return 'ended'
        finally:
            self._kill_proc()
            self._proc = None
            try:
                proc.stdout.close()
            except Exception:
                pass
            if stream is not None:
                try:
                    stream.stop_stream()
                    stream.close()
                except Exception:
                    pass

    def _open_stream(self):
        if self._pa is None:
            self._pa = pyaudio.PyAudio()
        device = self.get('output_device') or self._tts_device()
        index, name = find_output_device(self._pa, device)
        logger.info(f"DJ: playing to {name or 'the default device'}")
        return self._pa.open(format=pyaudio.paInt16, channels=CHANNELS, rate=RATE, output=True,
                             output_device_index=index, frames_per_buffer=FRAMES_PER_CHUNK)

    def _apply_gain(self, data):
        """Music volume x ducking, ramped across the chunk so changes don't click.
        Ducks down in ~0.15 s when the TTS starts talking, comes back up over ~0.6 s."""
        data = data[:len(data) // 4 * 4]
        target = self.get('volume') / 100.0
        if self.get('duck_enabled'):
            try:
                if self._tts_active():
                    target *= self.get('duck_level') / 100.0
            except Exception:
                pass
        chunk_seconds = (len(data) // 4) / RATE
        start = self._gain
        if target < start:
            end = max(target, start - chunk_seconds / 0.15)
        else:
            end = min(target, start + chunk_seconds / 0.6)
        self._gain = end
        if start == end == 1.0:
            return data
        if np is not None:
            samples = np.frombuffer(data, dtype=np.int16).astype(np.float32).reshape(-1, CHANNELS)
            ramp = np.linspace(start, end, len(samples), dtype=np.float32)[:, None]
            return np.clip(samples * ramp, -32768, 32767).astype(np.int16).tobytes()
        samples = array.array('h', data)
        frames = len(samples) // CHANNELS
        for f in range(frames):
            g = start + (end - start) * f / max(1, frames - 1)
            for c in range(CHANNELS):
                i = f * CHANNELS + c
                samples[i] = max(-32768, min(32767, int(samples[i] * g)))
        return samples.tobytes()

    def _announce_now_playing(self, track):
        if not self.get('announce'):
            return
        text = f"Now playing {clean_title(track['title'])}"
        if track.get('requested_by'):
            text += f", requested by {track['requested_by']}"
        # Brackets out: a song title can't smuggle DECtalk commands ([:t 3000,99999]) into the announcement
        text = text.replace('[', ' ').replace(']', ' ').replace('<', ' ').replace('>', ' ')
        try:
            self._announce(text, track)
        except Exception as e:
            logger.error(f"DJ: announce failed: {e}")

    # ------------------------------------------------------------------ YouTube (yt-dlp)
    def _run_ytdlp(self, args, timeout=60, interruptible=False):
        """Run yt-dlp with args; returns (ok, stdout_lines, error_text).
        interruptible: the player thread's lookup - Skip/Stop kill it. (Searches and chat requests run in
        their own threads and must not take over the player's process slot.)"""
        exe = self.ytdlp_path()
        if not exe:
            return False, [], "yt-dlp isn't installed - use Download yt-dlp in the DJ window"
        base = [exe, '--no-cache-dir', '--no-warnings', '--encoding', 'utf-8']
        # YouTube now needs a JavaScript runtime; yt-dlp finds Deno by itself, Node has to be named
        runtime = [] if shutil.which('deno') or not shutil.which('node') else ['--js-runtimes', 'node']
        env = dict(os.environ, PYTHONIOENCODING='utf-8')
        for extra in (runtime, []):
            try:
                proc = subprocess.Popen(base + extra + args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                        stdin=subprocess.DEVNULL, creationflags=NO_WINDOW, env=env)
            except OSError as e:
                return False, [], f"couldn't run yt-dlp: {e}"
            if interruptible:
                self._proc = proc
            try:
                out, err = proc.communicate(timeout=timeout)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.communicate()
                return False, [], "yt-dlp took too long"
            finally:
                if interruptible:
                    self._proc = None
            err_text = err.decode('utf-8', 'replace').strip()
            if proc.returncode == 0:
                return True, [l for l in out.decode('utf-8', 'replace').splitlines() if l.strip()], ''
            if extra and 'js-runtimes' in err_text:
                continue  # older yt-dlp without that option: try again without it
            last = err_text.splitlines()[-1] if err_text else f"yt-dlp exit code {proc.returncode}"
            return False, [], last.replace('ERROR: ', '')
        return False, [], "yt-dlp failed"

    def _youtube_info(self, target, stream_url=False):
        """target: a URL or 'ytsearch1:words'. Returns ({title, duration, url, webpage_url, is_live}, error)."""
        args = ['--no-playlist', '--print', 'title', '--print', 'duration', '--print', 'webpage_url',
                '--print', 'is_live']
        if stream_url:
            args = ['-f', 'bestaudio/best'] + args + ['--print', 'urls']
        ok, lines, error = self._run_ytdlp(args + ['--', target], interruptible=stream_url)
        if not ok:
            return None, error
        if len(lines) < 4 + (1 if stream_url else 0):
            return None, "nothing found"
        try:
            duration = float(lines[1])
        except ValueError:
            duration = None
        info = {'title': lines[0], 'duration': duration, 'webpage_url': lines[2], 'is_live': lines[3] == 'True'}
        if stream_url:
            info['url'] = lines[4]
        return info, ''

    def add_youtube(self, text, play_next=False):
        """DJ window box: a link is added directly; anything else is a search (results shown in the window)"""
        text = text.strip()
        if not text:
            return
        if not text.lower().startswith(('http://', 'https://')):
            self.search(text)
            return

        def work():
            self.status = "Looking up the link..."
            info, error = self._youtube_info(text)
            if not info:
                self.status = f"Couldn't add that link: {error}"
                return
            self.add_tracks([make_track(info['title'], 'youtube', info['webpage_url'], info['duration'])],
                            front=play_next)
            self.status = f"Added {info['title']}"
        threading.Thread(target=work, daemon=True).start()

    def search(self, query, count=10):
        def work():
            self.search_busy = True
            self.status = f"Searching YouTube for {query}..."
            ok, lines, error = self._run_ytdlp(['--flat-playlist', '--print', '%(title)s\t%(duration)s\t%(id)s',
                                                '--', f'ytsearch{count}:{query}'])
            results = []
            for line in lines:
                parts = line.split('\t')
                if len(parts) == 3 and parts[2] and parts[2] != 'NA':
                    try:
                        duration = float(parts[1])
                    except ValueError:
                        duration = None
                    results.append(make_track(parts[0], 'youtube', f"https://www.youtube.com/watch?v={parts[2]}", duration))
            self.search_results = results
            self.search_version += 1
            self.search_busy = False
            self.status = f"{len(results)} result(s)" if ok else f"Search failed: {error}"
        threading.Thread(target=work, daemon=True).start()

    def update_ytdlp(self):
        """Download yt-dlp.exe next to the TTS if it's missing, otherwise let it update itself (yt-dlp -U)"""
        if self.ytdlp_busy:
            return

        def work():
            self.ytdlp_busy = True
            try:
                exe = self.ytdlp_path()
                if exe:
                    self.status = "Updating yt-dlp..."
                    proc = subprocess.run([exe, '-U'], capture_output=True, timeout=180,
                                          creationflags=NO_WINDOW, stdin=subprocess.DEVNULL)
                    lines = (proc.stdout + proc.stderr).decode('utf-8', 'replace').strip().splitlines()
                    self.status = f"yt-dlp: {lines[-1] if lines else 'done'}"
                else:
                    target = self.app_dir / 'yt-dlp.exe'
                    partial = target.with_name('yt-dlp.exe.part')
                    self.status = "Downloading yt-dlp (about 18 MB)..."
                    with urllib.request.urlopen(YTDLP_DOWNLOAD_URL, timeout=60) as response, open(partial, 'wb') as f:
                        shutil.copyfileobj(response, f)
                    os.replace(partial, target)
                    self.status = f"yt-dlp downloaded to {target}"
            except Exception as e:
                self.status = f"yt-dlp update failed: {e}"
                try:
                    (self.app_dir / 'yt-dlp.exe.part').unlink()
                except OSError:
                    pass
            finally:
                self.ytdlp_busy = False
        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------------ chat song requests
    def request_allowed(self, is_admin):
        mode = self.get('requests')
        return mode == 'everyone' or (mode == 'admins' and is_admin)

    def request(self, username, query, is_admin, reply):
        """Chat '!sr <YouTube link or search words>'. reply(text) gets the outcome (called from a worker
        thread). Returns False if requests are off for this person, so the caller can ignore the message."""
        if not self.request_allowed(is_admin):
            return False
        query = query.strip()
        if not query:
            reply(f"{username}: use !sr with a YouTube link or song name")
            return True
        if not is_admin:
            waiting = sum(1 for t in self.queue if t.get('requested_by') == username)
            if waiting >= self.get('request_limit'):
                reply(f"{username} already has a song waiting")
                return True
        if query.lower().startswith(('http://', 'https://')):
            if not YOUTUBE_URL.match(query):
                reply(f"{username}: only YouTube links can be requested")
                return True
            target = query
        else:
            target = f'ytsearch1:{query}'

        def work():
            info, error = self._youtube_info(target)
            if not info:
                reply(f"Couldn't find a song for {username}" + (f" ({error})" if error else ""))
                return
            if not is_admin:
                if info['is_live'] or info['duration'] is None:
                    reply(f"{username}: live streams can't be requested")
                    return
                limit = self.get('request_max_minutes')
                if limit and info['duration'] > limit * 60:
                    reply(f"{username}'s request is too long ({fmt_time(info['duration'])}, limit {limit} min)")
                    return
            self.add_tracks([make_track(info['title'], 'youtube', info['webpage_url'], info['duration'], username)])
            reply(f"{username} requested {info['title']} ({fmt_time(info['duration'])})")
        threading.Thread(target=work, daemon=True).start()
        return True

    # ------------------------------------------------------------------ devices
    @staticmethod
    def list_output_devices():
        """Output device names for the Output picker (MME's 31-character cut-off duplicates removed)"""
        if pyaudio is None:
            return []
        pa = pyaudio.PyAudio()
        try:
            names = []
            for i in range(pa.get_device_count()):
                info = pa.get_device_info_by_index(i)
                if info.get('maxOutputChannels', 0) > 0 and info['name'] not in names:
                    names.append(info['name'])
            return [n for n in names if not any(o != n and o.startswith(n) for o in names)]
        finally:
            pa.terminate()


class DJWindow:
    """The DJ window. Reads the engine's state on a 250 ms timer; never called from other threads."""

    def __init__(self, root, engine, voice_names=None, on_close=None):
        self.engine = engine
        self._voice_names = voice_names or (lambda: ['Default'])
        self._on_close = on_close
        self._queue_seen = -1
        self._search_seen = -1
        self._dragging = False
        self._poll_job = None

        self.win = tk.Toplevel(root)
        self.win.title("DJ")
        self.win.geometry("880x760")
        self.win.minsize(760, 640)
        self.win.protocol("WM_DELETE_WINDOW", self.close)

        self._build_now_playing()
        self._build_queue()
        self._build_youtube()
        self._build_options()
        self.status_var = tk.StringVar()
        ttk.Label(self.win, textvariable=self.status_var, foreground='gray').pack(fill=tk.X, padx=10, pady=(0, 2))
        ttk.Label(self.win, foreground='gray', text=(
            "Keys: Space play/pause   Ctrl+Right next   Ctrl+Left previous   Ctrl+S stop   "
            "Del remove   Ctrl+Up/Down move   Ctrl+O add files   Ctrl+F YouTube box")).pack(fill=tk.X, padx=10, pady=(0, 6))
        self._bind_keys()
        self._poll()

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

    # ------------------------------------------------------------------ layout
    def _build_now_playing(self):
        frame = ttk.LabelFrame(self.win, text="Now playing")
        frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        self.title_var = tk.StringVar(value="Nothing playing")
        ttk.Label(frame, textvariable=self.title_var, font=('Arial', 12, 'bold')).pack(anchor=tk.W, padx=8, pady=(6, 0))
        self.info_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.info_var).pack(anchor=tk.W, padx=8)

        seek_row = ttk.Frame(frame)
        seek_row.pack(fill=tk.X, padx=8, pady=4)
        self.pos_var = tk.DoubleVar(value=0)
        self.seek_scale = ttk.Scale(seek_row, from_=0, to=1, variable=self.pos_var, orient=tk.HORIZONTAL)
        self.seek_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.seek_scale.bind('<ButtonPress-1>', lambda e: setattr(self, '_dragging', True))
        self.seek_scale.bind('<ButtonRelease-1>', self._seek_release)
        self.time_var = tk.StringVar(value="0:00 / 0:00")
        ttk.Label(seek_row, textvariable=self.time_var, width=16, anchor=tk.E).pack(side=tk.LEFT, padx=(8, 0))

        controls = ttk.Frame(frame)
        controls.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(controls, text="⏮ Prev", width=8, command=self.engine.previous).pack(side=tk.LEFT, padx=2)
        self.play_btn = ttk.Button(controls, text="▶ Play", width=10, command=self.engine.toggle)
        self.play_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(controls, text="⏭ Next", width=8, command=self.engine.skip).pack(side=tk.LEFT, padx=2)
        ttk.Button(controls, text="⏹ Stop", width=8, command=self.engine.stop).pack(side=tk.LEFT, padx=2)

        ttk.Label(controls, text="Volume").pack(side=tk.LEFT, padx=(16, 4))
        self.volume_var = tk.IntVar(value=self.engine.get('volume'))
        ttk.Scale(controls, from_=0, to=100, variable=self.volume_var, orient=tk.HORIZONTAL, length=160,
                  command=lambda v: self._set_volume()).pack(side=tk.LEFT)
        self.volume_label = ttk.Label(controls, text=f"{self.volume_var.get()}%", width=5)
        self.volume_label.pack(side=tk.LEFT, padx=4)

        self.shuffle_var = tk.BooleanVar(value=self.engine.get('shuffle'))
        ttk.Checkbutton(controls, text="Shuffle", variable=self.shuffle_var,
                        command=lambda: self.engine.set('shuffle', self.shuffle_var.get())).pack(side=tk.LEFT, padx=(12, 4))
        ttk.Label(controls, text="Repeat").pack(side=tk.LEFT, padx=(8, 2))
        self.repeat_combo = ttk.Combobox(controls, values=["Off", "All", "One"], width=5, state='readonly')
        self.repeat_combo.set(self.engine.get('repeat').capitalize())
        self.repeat_combo.bind('<<ComboboxSelected>>',
                               lambda e: self.engine.set('repeat', self.repeat_combo.get().lower()))
        self.repeat_combo.pack(side=tk.LEFT)

    def _build_queue(self):
        frame = ttk.LabelFrame(self.win, text="Queue")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(6, 2))
        self.queue_tree = ttk.Treeview(tree_frame, columns=('n', 'title', 'length', 'by'), show='headings',
                                       height=9, selectmode='extended')
        for col, text, width, anchor in (('n', '#', 40, tk.E), ('title', 'Title', 480, tk.W),
                                         ('length', 'Length', 70, tk.E), ('by', 'Requested by', 140, tk.W)):
            self.queue_tree.heading(col, text=text)
            self.queue_tree.column(col, width=width, anchor=anchor, stretch=(col == 'title'))
        scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.queue_tree.yview)
        self.queue_tree.configure(yscrollcommand=scroll.set)
        self.queue_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.queue_tree.bind('<Double-1>', lambda e: self._play_selected())

        buttons = ttk.Frame(frame)
        buttons.pack(fill=tk.X, padx=6, pady=(2, 6))
        ttk.Button(buttons, text="+ Files...", command=self._add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons, text="+ Folder...", command=self._add_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons, text="▶ Play Now", command=self._play_selected).pack(side=tk.LEFT, padx=(12, 2))
        ttk.Button(buttons, text="▲", width=3, command=lambda: self._move(-1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons, text="▼", width=3, command=lambda: self._move(1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons, text="✖ Remove", command=self._remove_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons, text="Clear", command=self._clear).pack(side=tk.LEFT, padx=2)
        self.queue_info_var = tk.StringVar()
        ttk.Label(buttons, textvariable=self.queue_info_var).pack(side=tk.RIGHT, padx=4)

    def _build_youtube(self):
        frame = ttk.LabelFrame(self.win, text="YouTube")
        frame.pack(fill=tk.X, padx=10, pady=5)
        row = ttk.Frame(frame)
        row.pack(fill=tk.X, padx=6, pady=(6, 2))
        self.yt_var = tk.StringVar()
        self.yt_entry = ttk.Entry(row, textvariable=self.yt_var)
        self.yt_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.yt_entry.bind('<Return>', lambda e: self._youtube_go())
        ttk.Button(row, text="Add link / Search", command=self._youtube_go).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(frame, foreground='gray', text="Paste a YouTube link to add it, or type words to search").pack(anchor=tk.W, padx=6)

        self.results_tree = ttk.Treeview(frame, columns=('title', 'length'), show='headings', height=5)
        self.results_tree.heading('title', text='Search results')
        self.results_tree.heading('length', text='Length')
        self.results_tree.column('title', width=600, stretch=True)
        self.results_tree.column('length', width=70, anchor=tk.E, stretch=False)
        self.results_tree.pack(fill=tk.X, padx=6, pady=2)
        self.results_tree.bind('<Double-1>', lambda e: self._add_results(front=False))

        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, padx=6, pady=(2, 6))
        ttk.Button(row2, text="+ Add to Queue", command=lambda: self._add_results(front=False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(row2, text="Play Next", command=lambda: self._add_results(front=True)).pack(side=tk.LEFT, padx=2)
        self.ytdlp_btn = ttk.Button(row2, text="Update yt-dlp", command=self.engine.update_ytdlp)
        self.ytdlp_btn.pack(side=tk.RIGHT, padx=2)
        self.ytdlp_var = tk.StringVar()
        ttk.Label(row2, textvariable=self.ytdlp_var, foreground='gray').pack(side=tk.RIGHT, padx=6)

    def _build_options(self):
        frame = ttk.LabelFrame(self.win, text="Options")
        frame.pack(fill=tk.X, padx=10, pady=5)
        e = self.engine

        row = ttk.Frame(frame)
        row.pack(fill=tk.X, padx=6, pady=2)
        self.duck_var = tk.BooleanVar(value=e.get('duck_enabled'))
        ttk.Checkbutton(row, text="Lower the music while the TTS talks, to", variable=self.duck_var,
                        command=lambda: e.set('duck_enabled', self.duck_var.get())).pack(side=tk.LEFT)
        self.duck_level_var = tk.StringVar(value=str(e.get('duck_level')))
        duck_box = ttk.Spinbox(row, from_=0, to=100, increment=5, width=5, textvariable=self.duck_level_var,
                               command=self._save_duck_level)
        duck_box.pack(side=tk.LEFT, padx=4)
        duck_box.bind('<FocusOut>', lambda ev: self._save_duck_level())
        ttk.Label(row, text="% of its volume").pack(side=tk.LEFT)

        row = ttk.Frame(frame)
        row.pack(fill=tk.X, padx=6, pady=2)
        self.normalize_var = tk.BooleanVar(value=e.get('normalize'))
        ttk.Checkbutton(row, text="Even out loudness between songs (starts with the next song)",
                        variable=self.normalize_var,
                        command=lambda: e.set('normalize', self.normalize_var.get())).pack(side=tk.LEFT)

        row = ttk.Frame(frame)
        row.pack(fill=tk.X, padx=6, pady=2)
        self.announce_var = tk.BooleanVar(value=e.get('announce'))
        ttk.Checkbutton(row, text='Announce "now playing" with', variable=self.announce_var,
                        command=lambda: e.set('announce', self.announce_var.get())).pack(side=tk.LEFT)
        self.announce_voice = ttk.Combobox(row, values=self._voice_names(), width=40, state='readonly')
        self.announce_voice.set(e.get('announce_voice'))
        self.announce_voice.bind('<<ComboboxSelected>>', lambda ev: e.set('announce_voice', self.announce_voice.get()))
        self.announce_voice.pack(side=tk.LEFT, padx=4)
        ttk.Label(row, text="(Default = announcement voice; needs the TTS started)", foreground='gray').pack(side=tk.LEFT)

        row = ttk.Frame(frame)
        row.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(row, text="Song requests in chat (!sr link or song name):").pack(side=tk.LEFT)
        self.requests_var = tk.StringVar(value=e.get('requests'))
        for value, text in (('off', 'Off'), ('admins', 'Admins only'), ('everyone', 'Everyone')):
            ttk.Radiobutton(row, text=text, value=value, variable=self.requests_var,
                            command=lambda: e.set('requests', self.requests_var.get())).pack(side=tk.LEFT, padx=4)
        ttk.Label(row, text="  Max length").pack(side=tk.LEFT)
        self.max_len_var = tk.StringVar(value=str(e.get('request_max_minutes')))
        max_box = ttk.Spinbox(row, from_=0, to=180, width=4, textvariable=self.max_len_var, command=self._save_max_len)
        max_box.pack(side=tk.LEFT, padx=2)
        max_box.bind('<FocusOut>', lambda ev: self._save_max_len())
        ttk.Label(row, text="min (0 = any; admins have no limit)").pack(side=tk.LEFT)

        row = ttk.Frame(frame)
        row.pack(fill=tk.X, padx=6, pady=(2, 6))
        ttk.Label(row, text="Output:").pack(side=tk.LEFT)
        self.output_combo = ttk.Combobox(row, width=55, state='readonly',
                                         values=["Same as the TTS"] + e.list_output_devices())
        self.output_combo.set(e.get('output_device') or "Same as the TTS")
        self.output_combo.bind('<<ComboboxSelected>>', lambda ev: self._set_output())
        self.output_combo.pack(side=tk.LEFT, padx=4)
        ttk.Label(row, text="(pick your speakers here to test without going through the mic)",
                  foreground='gray').pack(side=tk.LEFT)

        row = ttk.Frame(frame)
        row.pack(fill=tk.X, padx=6, pady=(0, 6))
        ttk.Label(row, text="ffmpeg (plays all music and sounds):").pack(side=tk.LEFT)
        self.ffmpeg_var = tk.StringVar()
        ttk.Label(row, textvariable=self.ffmpeg_var, foreground='gray').pack(side=tk.LEFT, padx=6)
        self.ffmpeg_btn = ttk.Button(row, text="Download ffmpeg", command=e.update_ffmpeg)
        self.ffmpeg_btn.pack(side=tk.LEFT, padx=4)

    def _bind_keys(self):
        def unless_typing(action):
            def handler(event):
                if isinstance(event.widget, (tk.Entry, ttk.Entry, ttk.Spinbox, ttk.Combobox)):
                    return None
                action()
                return 'break'
            return handler
        self.win.bind('<space>', unless_typing(self.engine.toggle))
        self.win.bind('<Control-Right>', lambda ev: self.engine.skip())
        self.win.bind('<Control-Left>', lambda ev: self.engine.previous())
        self.win.bind('<Control-s>', lambda ev: self.engine.stop())
        self.win.bind('<Delete>', unless_typing(self._remove_selected))
        self.win.bind('<Control-Up>', lambda ev: self._move(-1))
        self.win.bind('<Control-Down>', lambda ev: self._move(1))
        self.win.bind('<Control-o>', lambda ev: self._add_files())
        self.win.bind('<Control-f>', lambda ev: (self.yt_entry.focus_set(), self.yt_entry.select_range(0, tk.END)))

    # ------------------------------------------------------------------ actions
    def _set_volume(self):
        value = int(float(self.volume_var.get()))
        self.volume_label.config(text=f"{value}%")
        self.engine.settings['volume'] = value  # live; saved when the slider is released (below)
        if getattr(self, '_volume_job', None):
            self.win.after_cancel(self._volume_job)
        self._volume_job = self.win.after(500, self.engine._save)

    def _save_duck_level(self):
        try:
            self.engine.set('duck_level', max(0, min(100, int(self.duck_level_var.get()))))
        except ValueError:
            self.duck_level_var.set(str(self.engine.get('duck_level')))

    def _save_max_len(self):
        try:
            self.engine.set('request_max_minutes', max(0, int(self.max_len_var.get())))
        except ValueError:
            self.max_len_var.set(str(self.engine.get('request_max_minutes')))

    def _set_output(self):
        choice = self.output_combo.get()
        self.engine.set('output_device', '' if choice == "Same as the TTS" else choice)
        if self.engine.state in ('playing', 'paused'):
            self.engine.seek(self.engine.position)  # reopen on the new device where it left off

    def _seek_release(self, event):
        self._dragging = False
        self.engine.seek(self.pos_var.get())

    def _selected_indices(self):
        return sorted(self.queue_tree.index(item) for item in self.queue_tree.selection())

    def _add_files(self):
        paths = filedialog.askopenfilenames(parent=self.win, title="Add music",
                                            filetypes=[("Audio", " ".join(f"*{x}" for x in sorted(AUDIO_EXTENSIONS))),
                                                       ("All files", "*.*")])
        if paths:
            self.engine.add_files(paths)

    def _add_folder(self):
        folder = filedialog.askdirectory(parent=self.win, title="Add a music folder (includes sub-folders)")
        if folder:
            self.engine.add_folder(folder)

    def _play_selected(self):
        sel = self._selected_indices()
        if sel:
            self.engine.play_now(sel[0])

    def _remove_selected(self):
        sel = self._selected_indices()
        if sel:
            self.engine.remove(sel)

    def _move(self, delta):
        sel = self._selected_indices()
        if len(sel) == 1:
            new = self.engine.move(sel[0], delta)
            self._redraw_queue()
            items = self.queue_tree.get_children()
            if 0 <= new < len(items):
                self.queue_tree.selection_set(items[new])
                self.queue_tree.see(items[new])

    def _clear(self):
        if self.engine.queue and messagebox.askyesno("Clear queue", "Remove everything from the queue?", parent=self.win):
            self.engine.clear()

    def _youtube_go(self):
        text = self.yt_var.get().strip()
        if not text:
            return
        self.engine.add_youtube(text)
        if text.lower().startswith(('http://', 'https://')):
            self.yt_var.set('')

    def _add_results(self, front):
        picked = [self.engine.search_results[self.results_tree.index(i)] for i in self.results_tree.selection()
                  if self.results_tree.index(i) < len(self.engine.search_results)]
        if picked:
            self.engine.add_tracks([dict(t) for t in picked], front=front)
            self.engine.status = f"Added {len(picked)} song(s)" + (" to play next" if front else "")

    # ------------------------------------------------------------------ refresh
    def _redraw_queue(self):
        self._queue_seen = self.engine.queue_version
        selected = self._selected_indices()
        self.queue_tree.delete(*self.queue_tree.get_children())
        total = 0.0
        for n, track in enumerate(list(self.engine.queue), 1):
            total += track.get('duration') or 0
            title = track['title'] + ("  [YouTube]" if track['source'] == 'youtube' else "")
            self.queue_tree.insert('', tk.END, values=(n, title, fmt_time(track.get('duration')),
                                                       track.get('requested_by') or ''))
        items = self.queue_tree.get_children()
        for i in selected:
            if i < len(items):
                self.queue_tree.selection_add(items[i])
        self.queue_info_var.set(f"{len(items)} song(s), {fmt_time(total)}")

    def _redraw_results(self):
        self._search_seen = self.engine.search_version
        self.results_tree.delete(*self.results_tree.get_children())
        for track in self.engine.search_results:
            self.results_tree.insert('', tk.END, values=(track['title'], fmt_time(track.get('duration'))))

    def _poll(self):
        e = self.engine
        try:
            track = e.current
            if track and e.state != 'stopped':
                self.title_var.set(track['title'])
                state = {'loading': "Loading...", 'paused': "Paused", 'playing': "Playing"}.get(e.state, '')
                by = f" - requested by {track['requested_by']}" if track.get('requested_by') else ''
                self.info_var.set(state + by)
                duration = track.get('duration') or max(e.position, 1)
                self.seek_scale.configure(to=duration)
                if not self._dragging:
                    self.pos_var.set(min(e.position, duration))
                self.time_var.set(f"{fmt_time(e.position)} / {fmt_time(track.get('duration'))}")
            else:
                self.title_var.set("Nothing playing")
                self.info_var.set("Add music below, then press Play")
                self.pos_var.set(0)
                self.time_var.set("0:00 / 0:00")
            self.play_btn.configure(text="⏸ Pause" if e.state in ('playing', 'loading') else "▶ Play")
            if e.queue_version != self._queue_seen:
                self._redraw_queue()
            if e.search_version != self._search_seen:
                self._redraw_results()
            self.status_var.set(e.status)
            have = e.ytdlp_path()
            self.ytdlp_var.set("working..." if e.ytdlp_busy else ("yt-dlp ready" if have else "yt-dlp not installed"))
            self.ytdlp_btn.configure(text="Update yt-dlp" if have else "Download yt-dlp",
                                     state=tk.DISABLED if e.ytdlp_busy else tk.NORMAL)
            ffmpeg = e.ffmpeg_path()
            if e.ffmpeg_busy:
                self.ffmpeg_var.set("working...")
            elif not ffmpeg:
                self.ffmpeg_var.set("not installed - one-time ~80 MB download")
            elif Path(ffmpeg).parent == e.app_dir / 'ffmpeg':
                self.ffmpeg_var.set("ready (in the TTS folder)")
            else:
                self.ffmpeg_var.set("ready (installed on this PC)")
            self.ffmpeg_btn.configure(text="Update ffmpeg" if ffmpeg else "Download ffmpeg",
                                      state=tk.DISABLED if e.ffmpeg_busy else tk.NORMAL)
        except tk.TclError:
            return  # window is closing
        self._poll_job = self.win.after(250, self._poll)
