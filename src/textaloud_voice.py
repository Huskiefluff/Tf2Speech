"""
TextAloud voices (e.g. Acapela "Will From Afar" bought from NextUp)

Voices sold through NextUp for TextAloud only run inside TextAloud, so they never show up as normal Windows
voices. TextAloud has an official automation interface (the one its browser/Office add-ons use), so the TTS asks
it to write each line to a WAV file and then plays that like DECtalk: same output device, volume slider, !stop,
mutes and music ducking.

TextAloud runs on a separate, invisible Windows desktop that only the TTS uses, so nothing it ever shows - its
window, a crash box - can appear on your screen or pull focus away from the game. The TTS starts it when one of
its voices is needed, closes it after a while without use, and closes it when the TTS exits. A TextAloud you
opened yourself (on your normal desktop) is used as it is and left alone.

TextAloud sometimes crashes into an "Error occurred" box (Send Error Report / Don't Send). It keeps running
behind that box but stops answering, so the box is watched for: the stuck copy is closed and the line is tried
once more with a fresh one. If it keeps failing, its voices fall back to the default voice for a few minutes.
"""

import os
import re
import time
import logging
import threading
import subprocess

try:
    import pythoncom
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

logger = logging.getLogger(__name__)

PROGID = "TextAloud_3.TextAloudIntf_3"
PREFIX = "[TextAloud] "
NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)

# TextAloud acts on tags inside the text: {{Audio=C:\any\file.mp3}} plays a file from this PC into the audio,
# SAPI XML like <silence msec="..."/> and Acapela tags like \pau=...\ add pauses that can hold the queue for
# minutes. Chat text must never reach any of those, so the characters that make them are removed.
TAG_CHARS = re.compile(r'[{}<>\\]')

# TextAloud's names for a 22 kHz mono 16-bit WAV (352.8 kbps); other values are rejected as invalid
WAV_FORMAT = ("WAV", "352.8", "22.05", "Mono")

DESKTOP_NAME = "TTS_TextAloud"  # the invisible desktop TextAloud lives on
DIALOG_CLASS = '#32770'     # TextAloud's crash box (and any other message box) - it never shows one when healthy
HANG_SECONDS = 4            # its window not answering for this long = frozen (it answers within ms even mid-line)
FAILURE_LIMIT = 3           # this many failures...
FAILURE_WINDOW = 300        # ...within 5 minutes and its voices are rested (default voice) until they age out


try:
    import ctypes
    from ctypes import wintypes
    _user32 = ctypes.WinDLL('user32', use_last_error=True)
    _user32.CreateDesktopW.restype = wintypes.HANDLE
    _user32.CreateDesktopW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p, wintypes.DWORD,
                                       wintypes.DWORD, ctypes.c_void_p]
    _user32.SetThreadDesktop.restype = wintypes.BOOL
    _user32.SetThreadDesktop.argtypes = [wintypes.HANDLE]
    _EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    _user32.EnumDesktopWindows.argtypes = [wintypes.HANDLE, _EnumWindowsProc, wintypes.LPARAM]
except Exception:
    _user32 = None

_desktop = None
_desktop_lock = threading.Lock()


def _hidden_desktop():
    """Handle of the TTS's invisible desktop for TextAloud, created on first use (None if Windows won't allow it -
    then TextAloud runs on the normal desktop with its windows hidden instead)"""
    global _desktop
    with _desktop_lock:
        if _desktop is None:
            _desktop = (_user32 and _user32.CreateDesktopW(DESKTOP_NAME, None, None, 0, 0x10000000, None)) or 0
            if not _desktop:
                logger.warning("Could not create TextAloud's hidden desktop - hiding its windows instead")
        return _desktop or None


def _hidden_desktop_windows():
    """Top-level windows on the hidden desktop (EnumWindows only sees our own desktop)"""
    if not _desktop:
        return []
    found = []

    def collect(hwnd, _):
        found.append(hwnd)
        return True
    callback = _EnumWindowsProc(collect)
    _user32.EnumDesktopWindows(_desktop, callback, 0)
    return found


def _running_on_normal_desktop():
    """TextAloud is open on your desktop (you opened it yourself) - then the TTS talks to that copy. Windows only
    connects programs on the same desktop, so asking from the hidden desktop would start a second copy, which
    TextAloud refuses."""
    pids = _process_ids()
    if not pids:
        return False
    try:
        import win32process
        return not any(win32process.GetWindowThreadProcessId(h)[1] in pids for h in _hidden_desktop_windows())
    except Exception:
        return True


class TextAloudFailed(Exception):
    pass


def clean_text(text):
    return re.sub(r'\s+', ' ', TAG_CHARS.sub(' ', text or '')).strip()


def is_installed():
    """True if TextAloud 3 is installed (its automation interface is registered)."""
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, PROGID + r"\CLSID"))
        return COM_AVAILABLE
    except OSError:
        return False


def _process_ids():
    """Process IDs of running TextAloud copies (a Windows process snapshot - about a millisecond, so the crash
    watch can look several times a second; tasklist is the fallback)"""
    try:
        import ctypes
        from ctypes import wintypes

        class PROCESSENTRY32W(ctypes.Structure):
            _fields_ = [('dwSize', wintypes.DWORD), ('cntUsage', wintypes.DWORD),
                        ('th32ProcessID', wintypes.DWORD), ('th32DefaultHeapID', ctypes.c_size_t),
                        ('th32ModuleID', wintypes.DWORD), ('cntThreads', wintypes.DWORD),
                        ('th32ParentProcessID', wintypes.DWORD), ('pcPriClassBase', ctypes.c_long),
                        ('dwFlags', wintypes.DWORD), ('szExeFile', ctypes.c_wchar * 260)]
        k32 = ctypes.WinDLL('kernel32', use_last_error=True)
        k32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
        k32.Process32FirstW.argtypes = k32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
        k32.CloseHandle.argtypes = [wintypes.HANDLE]
        snap = k32.CreateToolhelp32Snapshot(2, 0)  # TH32CS_SNAPPROCESS
        if not snap or snap == wintypes.HANDLE(-1).value:
            raise OSError("no snapshot")
        try:
            entry = PROCESSENTRY32W()
            entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
            pids = set()
            ok = k32.Process32FirstW(snap, ctypes.byref(entry))
            while ok:
                if entry.szExeFile.lower() == 'textaloudmp3.exe':
                    pids.add(entry.th32ProcessID)
                ok = k32.Process32NextW(snap, ctypes.byref(entry))
            return pids
        finally:
            k32.CloseHandle(snap)
    except Exception:
        pass
    try:
        out = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq TextAloudMP3.exe', '/FO', 'CSV', '/NH'],
                             capture_output=True, text=True, timeout=5, creationflags=NO_WINDOW).stdout
        return {int(line.split('","')[1]) for line in out.splitlines() if 'textaloudmp3.exe' in line.lower()}
    except Exception:
        return set()


def _is_running():
    return bool(_process_ids())


def _windows(pids=None):
    """(hwnd, title, class, visible) for TextAloud's top-level windows"""
    try:
        import win32gui
        import win32process
    except ImportError:
        return []
    pids = _process_ids() if pids is None else pids
    found = []
    if not pids:
        return found
    handles = []
    try:
        win32gui.EnumWindows(lambda hwnd, _: handles.append(hwnd), None)
    except Exception:
        pass
    handles += _hidden_desktop_windows()
    for hwnd in dict.fromkeys(handles):
        try:
            if win32process.GetWindowThreadProcessId(hwnd)[1] in pids:
                found.append((hwnd, win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd),
                              bool(win32gui.IsWindowVisible(hwnd))))
        except Exception:
            pass  # closed meanwhile
    return found


def _error_dialogs(pids=None):
    """Titles of message boxes TextAloud is showing - normally its 'Error occurred' crash box"""
    return [title for _, title, cls, visible in _windows(pids) if visible and cls == DIALOG_CLASS]


def _responding(pids=None):
    """False if TextAloud's main window doesn't answer a do-nothing message within a second (frozen)"""
    try:
        import ctypes
        from ctypes import wintypes
        for hwnd, _, cls, _ in _windows(pids):
            if cls == 'TApplication':
                answer = wintypes.DWORD()
                if not ctypes.windll.user32.SendMessageTimeoutW(hwnd, 0, 0, 0, 0x0002, 1000, ctypes.byref(answer)):
                    return False  # WM_NULL with SMTO_ABORTIFHUNG unanswered
    except Exception:
        pass
    return True


def _hide_windows():
    """Hide TextAloud's windows completely (minimized it still has a taskbar button). Message boxes are left for
    the crash watch to find."""
    try:
        import win32gui
        for hwnd, _, cls, visible in _windows():
            if visible and cls != DIALOG_CLASS:
                win32gui.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception as e:
        logger.debug(f"Could not hide TextAloud's window: {e}")


def _kill_all():
    for pid in _process_ids():
        subprocess.run(['taskkill', '/F', '/PID', str(pid)], capture_output=True, creationflags=NO_WINDOW)


class TextAloudVoices:
    def __init__(self, idle_minutes=10):
        self.lock = threading.Lock()      # TextAloud renders one line at a time
        self.started_by_us = False
        self.idle_minutes = idle_minutes  # close our TextAloud after this long without use (0 = keep it open)
        self.last_used = time.time()
        self.failures = []                # times of recent crashes/hangs
        self._watcher = None

    # ------------------------------------------------------------------ health
    def available(self):
        """False while its voices are rested after repeated failures"""
        cutoff = time.time() - FAILURE_WINDOW
        self.failures = [t for t in self.failures if t > cutoff]
        return len(self.failures) < FAILURE_LIMIT

    def _failed(self, why):
        self.failures.append(time.time())
        logger.error(f"TextAloud: {why}")
        if not self.available():
            logger.error(f"TextAloud failed {FAILURE_LIMIT} times in {FAILURE_WINDOW // 60} minutes - "
                         f"using the default voice instead for a while")

    def _close_crashed(self, titles):
        logger.error(f"TextAloud crashed ({', '.join(titles)}) - closing it")
        _kill_all()
        self.started_by_us = False

    def _watch(self):
        """While our hidden TextAloud runs: close it the moment its crash box appears (so the box doesn't sit on
        screen mid-game), and close it after idle_minutes without use."""
        last_answer = time.time()
        next_ping = 0
        while self.started_by_us:
            time.sleep(0.5)
            if not self.started_by_us:
                break
            titles = _error_dialogs()
            if titles:
                self._close_crashed(titles)
                self._failed("crashed while idle")
                break
            if time.time() >= next_ping and not self.lock.locked():  # mid-line, _run does this itself
                next_ping = time.time() + 2
                if _responding():
                    last_answer = time.time()
                elif time.time() - last_answer > HANG_SECONDS:
                    logger.error("TextAloud froze while idle - closing it")
                    _kill_all()
                    self.started_by_us = False
                    self._failed("froze while idle")
                    break
            idle = time.time() - self.last_used
            if self.idle_minutes and idle > self.idle_minutes * 60 and not self.lock.locked():
                logger.info(f"TextAloud unused for {self.idle_minutes} minutes - closing it")
                self.shutdown()
                break
        self._watcher = None

    # ------------------------------------------------------------------ talking to TextAloud
    def _connect(self):
        """COM object for TextAloud (starting it hidden if it isn't running). Call from the thread that uses it."""
        titles = _error_dialogs()
        if titles:  # a crashed copy (ours or not) would never answer
            self._close_crashed(titles)
            time.sleep(0.5)
        was_running = _is_running()
        ta = win32com.client.Dispatch(PROGID)
        deadline = time.time() + 20
        while not ta.ServerInitialized and time.time() < deadline:
            time.sleep(0.1)
        if not was_running:
            self.started_by_us = True
            self.last_used = time.time()
            try:
                ta.GUIMinimized = True
                ta.TrayIconVisible = False
            except Exception as e:
                logger.debug(f"Could not hide TextAloud: {e}")
            logger.info("Started TextAloud in the background")
            if self._watcher is None:
                self._watcher = threading.Thread(target=self._watch, daemon=True, name="TextAloud watch")
                self._watcher.start()
        if self.started_by_us:
            _hide_windows()  # every time, in case it popped a window up again
        return ta

    def _run(self, job, timeout):
        """Run job() against TextAloud on a helper thread, watching for TextAloud's crash box (it keeps running
        behind the box but stops answering) and for hangs, so the speech queue never freezes waiting for it.
        Raises TextAloudFailed if it crashed or hung."""
        result = {}
        # Waiting for our turn isn't "stuck": only the helper's own time counts towards giving up.
        # The helper releases the lock itself, because it can outlive this call if TextAloud hangs.
        if not self.lock.acquire(timeout=timeout):
            raise TextAloudFailed("still busy with an earlier line")
        # The helper moves to the hidden desktop before using COM; a TextAloud that COM starts for it then
        # opens there too
        desktop = None if _running_on_normal_desktop() else _hidden_desktop()

        def work():
            if desktop and not _user32.SetThreadDesktop(desktop):
                logger.debug(f"Could not use the hidden desktop ({ctypes.get_last_error()})")
            pythoncom.CoInitialize()
            try:
                result['value'] = job()
            except Exception as e:
                result['error'] = e.with_traceback(None)  # drop frames still holding TextAloud objects
            finally:
                pythoncom.CoUninitialize()
                self.lock.release()
        helper = threading.Thread(target=work, daemon=True, name="TextAloud")
        helper.start()
        deadline = time.time() + timeout
        next_check = next_ping = 0
        last_answer = time.time()
        while helper.is_alive():
            helper.join(0.1)
            if not helper.is_alive():
                break
            now = time.time()
            if now >= next_check:  # the window scan costs a few ms, so 4 times a second is plenty
                next_check = now + 0.25
                titles = _error_dialogs()
                if titles:
                    self._close_crashed(titles)
                    helper.join(5)
                    raise TextAloudFailed("crashed")
            if now >= next_ping:
                next_ping = now + 1
                if _responding():
                    last_answer = time.time()
            frozen = time.time() - last_answer > HANG_SECONDS
            if frozen or time.time() > deadline:
                if self.started_by_us:
                    _kill_all()
                    self.started_by_us = False
                    helper.join(5)
                raise TextAloudFailed("froze" if frozen else "stopped responding")
        if 'error' in result:
            raise result['error']
        return result.get('value')

    def list_voices(self, windows_voice_names=()):
        """TextAloud's voices that aren't already plain Windows voices (those work directly, without TextAloud)."""
        if not is_installed():
            return []
        captions = None
        for attempt in (1, 2):
            try:
                captions = self._run(lambda: list(self._connect().VoiceCaptionList or []), timeout=40)
                break
            except TextAloudFailed as e:
                self._failed(f"{e} while listing voices")
                if attempt == 2:
                    raise
        windows = [_words(n) for n in windows_voice_names]
        result = []
        for caption in captions:
            key = _words(caption)[:2]
            if key and any(w[:len(key)] == key for w in windows):
                continue
            result.append(caption)
        return result

    def render(self, text, voice, wav_path, should_stop=None):
        """Have TextAloud write `text` in `voice` to wav_path. True when the WAV is complete; False if it couldn't
        (crashed twice, hung, or resting after repeated failures - the caller then uses the default voice).
        should_stop() is polled while waiting, so a !stop or mute during the ~1 second render skips the line."""
        text = clean_text(text)
        if not text or not self.available():
            return False
        wait = 6 + len(text) / 20   # a normal line takes about a second; 300 characters about 4

        def job():
            ta = self._connect()
            ta.SetOutputFileFormat(*WAV_FORMAT)
            ta.PlayTextToFile(text, voice, wav_path)
            # The call returns straight away; the file appears once TextAloud has written it
            deadline = time.time() + wait
            last_size = -1
            while time.time() < deadline:
                time.sleep(0.05)
                if should_stop and should_stop():
                    try:
                        ta.Stop()
                    except Exception:
                        pass
                    return False
                size = os.path.getsize(wav_path) if os.path.exists(wav_path) else 0
                if size > 44 and size == last_size and ta.ServerIdle:
                    return True
                last_size = size
            logger.warning(f"TextAloud did not finish '{voice}' in time")
            return False

        for attempt in (1, 2):
            if os.path.exists(wav_path):
                os.remove(wav_path)  # an existing file can make TextAloud ask about overwriting
            self.last_used = time.time()
            try:
                return self._run(job, timeout=wait + 6)  # + time for TextAloud to start (2-3 s)
            except TextAloudFailed as e:
                self._failed(f"{e} on '{voice}'")
                if attempt == 2 or not self.available() or (should_stop and should_stop()):
                    return False
                logger.info("Trying that line again with a fresh TextAloud")
            finally:
                self.last_used = time.time()
        return False

    def shutdown(self):
        """Close TextAloud if the TTS started it (and make sure a crash box on the way out doesn't linger)."""
        if not self.started_by_us:
            return
        self.started_by_us = False
        if not _is_running():
            return
        try:
            self._run(lambda: win32com.client.Dispatch(PROGID).Shutdown(), timeout=10)
        except Exception as e:
            logger.debug(f"Could not close TextAloud normally: {e}")
        deadline = time.time() + 5
        while _is_running() and time.time() < deadline:
            if _error_dialogs():
                break
            time.sleep(0.2)
        if _is_running():
            _kill_all()
            logger.info("Closed TextAloud (forced)")
        else:
            logger.info("Closed TextAloud")


def _words(name):
    """'IVONA 2 Brian - British...' / 'IVONA Brian22 (UK English)' -> ['ivona', 'brian', ...] for matching."""
    return re.sub(r'[^a-z]+', ' ', name.lower()).split()
