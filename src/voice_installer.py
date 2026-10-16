"""
Getting more voices into the TTS

Windows voices: the voices Windows adds through Settings > Time & language > Speech ("OneCore" voices, e.g.
Microsoft Mark, Pavel, Herena) can't be seen by SAPI programs like the TTS. Copying their entries into SAPI's own
voice list makes them normal Windows voices. SAPI only reads that list machine-wide (a per-user copy is ignored),
so it takes one Windows admin prompt. An undo file is saved before anything changes.

Free voices: Piper (MIT) neural voices, downloaded from the official Piper releases / voice catalog. They're played
like DECtalk: Piper writes each line to a WAV file.
"""

import os
import sys
import time
import json
import ctypes
import hashlib
import logging
import threading
import subprocess
import urllib.request
from pathlib import Path

logger = logging.getLogger(__name__)

ONECORE_KEY = r"SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens"
SAPI_KEY = r"SOFTWARE\Microsoft\Speech\Voices\Tokens"
NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)


# ---------------------------------------------------------------- Windows (OneCore) voices
def open_windows_voice_settings():
    """Windows' own page for adding voices (Settings > Time & language > Speech)"""
    os.startfile('ms-settings:speech')


def windows_voices():
    """[(token name, description, usable in the TTS)] for every OneCore voice on this PC"""
    import winreg
    result = []
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, ONECORE_KEY) as root:
            for i in range(winreg.QueryInfoKey(root)[0]):
                name = winreg.EnumKey(root, i)
                with winreg.OpenKey(root, name) as key:
                    try:
                        desc = winreg.QueryValueEx(key, '')[0]
                    except OSError:
                        desc = name
                result.append((name, desc, _in_sapi(name)))
    except OSError:
        pass
    return result


def _in_sapi(token):
    import winreg
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, SAPI_KEY + '\\' + token))
        return True
    except OSError:
        return False


def _reg_value(name, value, kind):
    """One value line of a .reg file"""
    import winreg
    label = '@' if name == '' else '"' + name.replace('\\', '\\\\').replace('"', '\\"') + '"'
    if kind == winreg.REG_SZ:
        return label + '="' + str(value).replace('\\', '\\\\').replace('"', '\\"') + '"'
    if kind == winreg.REG_DWORD:
        return f'{label}=dword:{int(value) & 0xffffffff:08x}'
    if kind == winreg.REG_EXPAND_SZ:
        data = (str(value) + '\0').encode('utf-16-le')
        return f'{label}=hex(2):' + ','.join(f'{b:02x}' for b in data)
    if kind == winreg.REG_MULTI_SZ:
        data = ('\0'.join(value) + '\0\0').encode('utf-16-le')
        return f'{label}=hex(7):' + ','.join(f'{b:02x}' for b in data)
    data = value if isinstance(value, (bytes, bytearray)) else bytes(value or b'')
    return f'{label}=hex({kind:x}):' + ','.join(f'{b:02x}' for b in data)


def _copy_key_lines(src_path, dst_path):
    """.reg lines that recreate HKLM\\src_path (values and sub-keys) at HKLM\\dst_path"""
    import winreg
    lines = []
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, src_path) as key:
        n_sub, n_val, _ = winreg.QueryInfoKey(key)
        lines.append('')
        lines.append(f'[HKEY_LOCAL_MACHINE\\{dst_path}]')
        for i in range(n_val):
            name, value, kind = winreg.EnumValue(key, i)
            lines.append(_reg_value(name, value, kind))
        for i in range(n_sub):
            sub = winreg.EnumKey(key, i)
            lines += _copy_key_lines(src_path + '\\' + sub, dst_path + '\\' + sub)
    return lines


def _run_elevated(exe, args, timeout=120):
    """Run a program as administrator (Windows shows its UAC prompt). Exit code, or None if declined."""
    class SHELLEXECUTEINFO(ctypes.Structure):
        _fields_ = [('cbSize', ctypes.c_ulong), ('fMask', ctypes.c_ulong), ('hwnd', ctypes.c_void_p),
                    ('lpVerb', ctypes.c_wchar_p), ('lpFile', ctypes.c_wchar_p), ('lpParameters', ctypes.c_wchar_p),
                    ('lpDirectory', ctypes.c_wchar_p), ('nShow', ctypes.c_int), ('hInstApp', ctypes.c_void_p),
                    ('lpIDList', ctypes.c_void_p), ('lpClass', ctypes.c_wchar_p), ('hkeyClass', ctypes.c_void_p),
                    ('dwHotKey', ctypes.c_ulong), ('hIconOrMonitor', ctypes.c_void_p), ('hProcess', ctypes.c_void_p)]
    info = SHELLEXECUTEINFO()
    info.cbSize = ctypes.sizeof(info)
    info.fMask = 0x00000040  # SEE_MASK_NOCLOSEPROCESS
    info.lpVerb, info.lpFile, info.lpParameters, info.nShow = 'runas', exe, args, 0
    if not ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(info)):
        return None  # UAC declined (or failed)
    ctypes.windll.kernel32.WaitForSingleObject(ctypes.c_void_p(info.hProcess), int(timeout * 1000))
    code = ctypes.c_ulong()
    ctypes.windll.kernel32.GetExitCodeProcess(ctypes.c_void_p(info.hProcess), ctypes.byref(code))
    ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(info.hProcess))
    return code.value


def make_windows_voices_usable(tokens, work_dir):
    """Copy the given OneCore voices into SAPI's voice list (one admin prompt). Returns (ok, message).
    Saves undo_windows_voices.reg in work_dir first."""
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    tokens = [t for t in tokens if not _in_sapi(t)]
    if not tokens:
        return True, "Those voices are already usable"
    add = ['Windows Registry Editor Version 5.00']
    for t in tokens:
        add += _copy_key_lines(ONECORE_KEY + '\\' + t, SAPI_KEY + '\\' + t)
    undo_path = work_dir / 'undo_windows_voices.reg'
    old_undo = undo_path.read_text(encoding='utf-16') if undo_path.exists() else ''
    undo = ['Windows Registry Editor Version 5.00'] + [l for l in old_undo.splitlines()[1:] if l.strip()]
    undo += [f'[-HKEY_LOCAL_MACHINE\\{SAPI_KEY}\\{t}]' for t in tokens]
    undo_path.write_text('\r\n'.join(undo) + '\r\n', encoding='utf-16')
    add_path = work_dir / 'add_windows_voices.reg'
    add_path.write_text('\r\n'.join(add) + '\r\n', encoding='utf-16')
    code = _run_elevated('reg.exe', f'import "{add_path}"')
    if code is None:
        return False, "Windows' admin prompt was cancelled - nothing was changed"
    missing = [t for t in tokens if not _in_sapi(t)]
    if code != 0 or missing:
        return False, f"Windows couldn't add {len(missing)} of them (reg.exe exit {code})"
    return True, f"Added {len(tokens)} voice(s) - restart the TTS to use them"


def undo_windows_voices(work_dir):
    """Remove the voices this TTS added to SAPI's list (one admin prompt)"""
    undo_path = Path(work_dir) / 'undo_windows_voices.reg'
    if not undo_path.exists():
        return False, "Nothing to undo"
    code = _run_elevated('reg.exe', f'import "{undo_path}"')
    if code is None:
        return False, "Windows' admin prompt was cancelled - nothing was changed"
    if code == 0:
        undo_path.unlink()
        return True, "Removed - restart the TTS to update its voice lists"
    return False, f"reg.exe exit {code}"


def sapi_voice_names():
    """Names of the voices SAPI programs can already use (machine-wide list)"""
    import winreg
    names = []
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, SAPI_KEY) as root:
            for i in range(winreg.QueryInfoKey(root)[0]):
                try:
                    with winreg.OpenKey(root, winreg.EnumKey(root, i) + r'\Attributes') as attrs:
                        names.append(winreg.QueryValueEx(attrs, 'Name')[0])
                except OSError:
                    pass
    except OSError:
        pass
    return names


def windows_voice_status():
    """[(token, description, status)] - status 'usable', 'available' (can be added) or 'duplicate' (the same voice
    is already a normal Windows voice, e.g. 'Microsoft David' next to 'Microsoft David Desktop')"""
    have = [n.lower() for n in sapi_voice_names()]
    result = []
    for token, desc, usable in windows_voices():
        name = desc.split(' - ')[0].lower()
        if usable:
            status = 'usable'
        elif any(h.startswith(name) for h in have):
            status = 'duplicate'
        else:
            status = 'available'
        result.append((token, desc, status))
    return result


# ---------------------------------------------------------------- Piper (free neural voices)
PIPER_PREFIX = "[Piper] "
PIPER_ZIP_URL = "https://github.com/rhasspy/piper/releases/download/2023.11.14-2/piper_windows_amd64.zip"
PIPER_ZIP_SIZE = 22477236
CATALOG_URL = "https://huggingface.co/rhasspy/piper-voices/resolve/main/voices.json"
VOICE_BASE_URL = "https://huggingface.co/rhasspy/piper-voices/resolve/main/"


class Cancelled(Exception):
    pass


def download(url, dest, size=None, md5=None, progress=None, cancel=None):
    """Download to dest via a .part file; the size and md5 (when known) must match or nothing is kept"""
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    part = dest.with_name(dest.name + '.part')
    digest = hashlib.md5()
    done = 0
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'TF2-TTS voice downloader'})
        with urllib.request.urlopen(req, timeout=60) as r, open(part, 'wb') as f:
            total = size or int(r.headers.get('Content-Length') or 0)
            while True:
                if cancel and cancel():
                    raise Cancelled()
                block = r.read(1 << 18)
                if not block:
                    break
                f.write(block)
                digest.update(block)
                done += len(block)
                if progress:
                    progress(done, total)
        if size and done != size:
            raise IOError(f"download incomplete ({done} of {size} bytes)")
        if md5 and digest.hexdigest() != md5:
            raise IOError("download damaged (checksum mismatch)")
        part.replace(dest)
    finally:
        if part.exists():
            try:
                part.unlink()
            except OSError:
                pass


def piper_exe(voices_dir):
    return Path(voices_dir) / 'piper' / 'piper.exe'


def install_piper_engine(voices_dir, progress=None, cancel=None):
    """Download the Piper engine (22 MB, MIT, official GitHub release) into voices\\piper"""
    import zipfile
    voices_dir = Path(voices_dir)
    zip_path = voices_dir / 'piper_windows_amd64.zip'
    download(PIPER_ZIP_URL, zip_path, size=PIPER_ZIP_SIZE, progress=progress, cancel=cancel)
    try:
        with zipfile.ZipFile(zip_path) as z:
            names = z.namelist()
            if any(n.startswith('/') or '..' in n.split('/') or not n.startswith('piper/') for n in names):
                raise IOError("unexpected files in the Piper download")
            z.extractall(voices_dir)
    finally:
        zip_path.unlink(missing_ok=True)
    return piper_exe(voices_dir).exists()


def load_catalog(voices_dir, refresh=False):
    """Piper's voice list (cached in voices\\piper_catalog.json; refresh=True fetches the current one)"""
    cache = Path(voices_dir) / 'piper_catalog.json'
    if refresh or not cache.exists():
        req = urllib.request.Request(CATALOG_URL, headers={'User-Agent': 'TF2-TTS voice downloader'})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.loads(r.read().decode('utf-8'))
        cache.parent.mkdir(parents=True, exist_ok=True)
        cache.write_text(json.dumps(data), encoding='utf-8')
        return data
    return json.loads(cache.read_text(encoding='utf-8'))


def voice_files(entry):
    """(onnx path, json path, onnx size, onnx md5, json size, json md5) for a catalog entry"""
    onnx = next(p for p in entry['files'] if p.endswith('.onnx'))
    cfg = next(p for p in entry['files'] if p.endswith('.onnx.json'))
    return (onnx, cfg, entry['files'][onnx]['size_bytes'], entry['files'][onnx]['md5_digest'],
            entry['files'][cfg]['size_bytes'], entry['files'][cfg]['md5_digest'])


def voice_dir(voices_dir, key):
    return Path(voices_dir) / 'piper_voices' / key


def installed_piper_voices(voices_dir):
    root = Path(voices_dir) / 'piper_voices'
    if not root.is_dir():
        return []
    return sorted(d.name for d in root.iterdir()
                  if (d / f'{d.name}.onnx').exists() and (d / f'{d.name}.onnx.json').exists())


def install_piper_voice(voices_dir, entry, progress=None, cancel=None):
    """Download one catalog voice (checked against the catalog's sizes and checksums)"""
    onnx, cfg, onnx_size, onnx_md5, cfg_size, cfg_md5 = voice_files(entry)
    folder = voice_dir(voices_dir, entry['key'])
    download(VOICE_BASE_URL + cfg, folder / f"{entry['key']}.onnx.json", size=cfg_size, md5=cfg_md5, cancel=cancel)
    download(VOICE_BASE_URL + onnx, folder / f"{entry['key']}.onnx", size=onnx_size, md5=onnx_md5,
             progress=progress, cancel=cancel)


def remove_piper_voice(voices_dir, key):
    import shutil
    folder = voice_dir(voices_dir, key)
    if folder.is_dir() and folder.parent.name == 'piper_voices':
        shutil.rmtree(folder)


class PiperVoices:
    """Speaks with Piper: one Piper process per voice in use, kept running (a voice takes ~1 s to load, then a
    line takes ~0.2 s) and closed after idle_minutes without use."""

    def __init__(self, voices_dir, idle_minutes=10):
        self.voices_dir = Path(voices_dir)
        self.idle_minutes = idle_minutes
        self.procs = {}        # voice key -> [Popen, lock, last used]
        self.lock = threading.Lock()
        self._watch = None

    def available(self):
        return piper_exe(self.voices_dir).exists()

    def voice_names(self):
        if not self.available():
            return []
        return [PIPER_PREFIX + key for key in installed_piper_voices(self.voices_dir)]

    def _proc(self, key):
        with self.lock:
            entry = self.procs.get(key)
            if entry and entry[0].poll() is None:
                return entry
            model = voice_dir(self.voices_dir, key) / f'{key}.onnx'
            if not model.exists():
                raise FileNotFoundError(f"Piper voice {key} isn't downloaded")
            exe = piper_exe(self.voices_dir)
            proc = subprocess.Popen([str(exe), '-m', str(model), '--json-input', '-q'], cwd=str(exe.parent),
                                    stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                                    creationflags=NO_WINDOW, text=True, encoding='utf-8', bufsize=1)
            entry = [proc, threading.Lock(), time.time()]
            self.procs[key] = entry
            if self._watch is None:
                self._watch = threading.Thread(target=self._idle_watch, daemon=True, name='Piper idle')
                self._watch.start()
            return entry

    def render(self, text, key, wav_path, should_stop=None, timeout=None):
        """Write text in voice key to wav_path. True when done; False if stopped or Piper failed."""
        text = ' '.join((text or '').split())
        if not text:
            return False
        if os.path.exists(wav_path):
            os.remove(wav_path)
        for attempt in (1, 2):
            proc, lock, _ = self._proc(key)
            with lock:
                try:
                    proc.stdin.write(json.dumps({'text': text, 'output_file': str(wav_path)}) + '\n')
                    proc.stdin.flush()
                except OSError:
                    self._kill(key)
                    continue  # Piper went away - start a fresh one
                result = {}
                reader = threading.Thread(target=lambda: result.setdefault('line', proc.stdout.readline()),
                                          daemon=True)
                reader.start()
                deadline = time.time() + (timeout or 10 + len(text) / 20)
                while reader.is_alive() and time.time() < deadline:
                    reader.join(0.05)
                    if should_stop and should_stop():
                        break
                if key in self.procs:
                    self.procs[key][2] = time.time()
                if reader.is_alive():  # stopped, or Piper hung: it can't take another line now
                    self._kill(key)
                    return False
                if result.get('line') and os.path.exists(wav_path):
                    return True
                self._kill(key)
        return False

    def _kill(self, key):
        with self.lock:
            entry = self.procs.pop(key, None)
        if entry:
            try:
                entry[0].kill()
                entry[0].wait(timeout=3)  # so it lets go of the file it was writing
            except (OSError, subprocess.TimeoutExpired):
                pass

    def _idle_watch(self):
        while True:
            time.sleep(30)
            if not self.idle_minutes:
                continue
            for key, entry in list(self.procs.items()):
                if time.time() - entry[2] > self.idle_minutes * 60 and not entry[1].locked():
                    logger.info(f"Piper voice {key} unused for {self.idle_minutes} minutes - closing it")
                    self._kill(key)

    def shutdown(self):
        for key in list(self.procs):
            self._kill(key)
