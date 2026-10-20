"""
Updates from TF2Speech's GitHub releases.

check()    asks GitHub for the newest release (a normal web request: no login, nothing about you or your chat is
           sent) and returns it if it's newer than this copy.
download() fetches that release's zip and checks its size and SHA-256 checksum (GitHub publishes it), so a
           broken or tampered download is never installed.
install()  puts the new exe in place. Windows can't overwrite a running program but can rename it, so this copy
           becomes TTS_TF2_64BIT.exe.old (kept as a backup) and the new one takes its name. The notes files are
           updated; a changed tts.cfg is installed with the old one kept as tts.cfg.old. Settings, voices,
           sounds and downloaded tools are never touched.
restart()  starts the new copy, which waits until this one has closed (--wait-pid) before it starts up.

Only the built exe updates itself; running from the source code never does.
"""

import os
import re
import sys
import json
import shutil
import hashlib
import zipfile
import subprocess
import urllib.request
from pathlib import Path

from version import APP_VERSION

REPO = 'Huskiefluff/Tf2Speech'
LATEST_API = f'https://api.github.com/repos/{REPO}/releases/latest'
RELEASES_PAGE = f'https://github.com/{REPO}/releases'
DOWNLOAD_PREFIX = f'https://github.com/{REPO}/releases/download/'
EXE_NAME = 'TTS_TF2_64BIT.exe'
NOTES = ('READ ME FIRST.txt', 'VOICES.txt', 'CHANGELOG.txt', 'LICENSE.txt', 'THIRD-PARTY-NOTICES.txt')
ZIP_NAME = re.compile(r'TF2Speech_v\d+\.\d+\.\d+\.zip')
# For testing against a local copy of a release: TTS_UPDATE_URL points at a stand-in for GitHub's answer
TEST_URL = os.environ.get('TTS_UPDATE_URL', '')


class UpdateError(Exception):
    pass


def parse_version(text):
    """'v1.2.3' -> (1, 2, 3); None if there's no version in it"""
    m = re.search(r'(\d+)\.(\d+)\.(\d+)', text or '')
    return tuple(int(n) for n in m.groups()) if m else None


def can_update():
    """Only the built exe can replace itself"""
    return bool(getattr(sys, 'frozen', False))


def app_dir():
    return Path(sys.executable).parent if can_update() else Path.cwd()


def _get(url, timeout):
    request = urllib.request.Request(url, headers={'User-Agent': f'TF2Speech/{APP_VERSION}',
                                                   'Accept': 'application/vnd.github+json'})
    return urllib.request.urlopen(request, timeout=timeout)


def check(timeout=10):
    """The newest release if it's newer than this copy, else None. Raises OSError/ValueError if GitHub can't be
    reached or answers something unexpected."""
    with _get(TEST_URL or LATEST_API, timeout) as response:
        data = json.load(response)
    newest = parse_version(data.get('tag_name'))
    if not newest or data.get('draft') or data.get('prerelease') or newest <= parse_version(APP_VERSION):
        return None
    asset = next((a for a in data.get('assets', []) if ZIP_NAME.fullmatch(a.get('name', ''))), None)
    url = asset.get('browser_download_url') if asset else None
    if url and not (TEST_URL or url.startswith(DOWNLOAD_PREFIX)):
        url = None  # only ever download from this project's own releases
    sha256 = None
    if asset and str(asset.get('digest', '')).startswith('sha256:'):
        sha256 = asset['digest'][len('sha256:'):]
    else:  # the checksum written in the release notes
        m = re.search(r'SHA-256[^\n]*?`([0-9a-fA-F]{64})`', data.get('body') or '')
        sha256 = m.group(1) if m else None
    page = data.get('html_url') or RELEASES_PAGE
    if not page.startswith(f'https://github.com/{REPO}/'):
        page = RELEASES_PAGE
    return {'version': '.'.join(map(str, newest)), 'notes': data.get('body') or '', 'page': page,
            'url': url, 'size': asset.get('size') if asset else None, 'sha256': sha256 and sha256.lower()}


def download(info, folder, progress=None, cancelled=lambda: False):
    """Download the release zip into folder and check it. Returns the zip's path. progress(done, total) is
    called as it goes; cancelled() is polled and stops the download."""
    if not info.get('url') or not info.get('sha256'):
        raise UpdateError("this release has no download with a checksum - get it from the releases page")
    folder = Path(folder)
    part = folder / 'update-download.part'
    digest, done = hashlib.sha256(), 0
    try:
        with _get(info['url'], 30) as response, open(part, 'wb') as f:
            total = int(response.headers.get('Content-Length') or info.get('size') or 0)
            while True:
                if cancelled():
                    raise UpdateError("cancelled")
                chunk = response.read(256 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                digest.update(chunk)
                done += len(chunk)
                if progress:
                    progress(done, total)
        if info.get('size') and done != info['size']:
            raise UpdateError(f"the download is incomplete ({done:,} of {info['size']:,} bytes)")
        if digest.hexdigest() != info['sha256']:
            raise UpdateError("the download doesn't match its checksum, so it wasn't installed")
        zip_path = folder / f"update-download-v{info['version']}.zip"
        os.replace(part, zip_path)
        return zip_path
    except BaseException:
        try:
            part.unlink()
        except OSError:
            pass
        raise


def install(zip_path, folder):
    """Put the new exe and notes from the downloaded zip in place. Returns True if tts.cfg changed (the old one
    is kept as tts.cfg.old). The running program keeps running until it's restarted."""
    folder = Path(folder)
    exe = folder / EXE_NAME
    new_exe = folder / (EXE_NAME + '.new')
    with zipfile.ZipFile(zip_path) as z:
        files = {Path(n).name: n for n in z.namelist() if not n.endswith('/')}
        if EXE_NAME not in files:
            raise UpdateError(f"the download has no {EXE_NAME}")
        with z.open(files[EXE_NAME]) as src, open(new_exe, 'wb') as dst:
            shutil.copyfileobj(src, dst)
        for note in NOTES:
            if note in files:
                (folder / note).write_bytes(z.read(files[note]))
        cfg_changed = False
        if 'tts.cfg' in files:
            new_cfg, cfg = z.read(files['tts.cfg']), folder / 'tts.cfg'
            if cfg.exists() and cfg.read_bytes() != new_cfg:
                os.replace(cfg, folder / 'tts.cfg.old')
                cfg_changed = True
            cfg.write_bytes(new_cfg)
    old = folder / (EXE_NAME + '.old')
    if old.exists():
        old.unlink()                  # the backup from the update before (not running)
    if exe.exists():
        os.replace(exe, old)          # a running program can be renamed, just not overwritten
    try:
        os.replace(new_exe, exe)
    except OSError:
        if old.exists():
            os.replace(old, exe)      # put the working copy back
        raise
    try:
        Path(zip_path).unlink()
    except OSError:
        pass
    return cfg_changed


def restart(folder):
    """Start the installed copy. It waits until this one has closed, so close the program right after this."""
    env = {k: v for k, v in os.environ.items() if not k.startswith(('_PYI', '_MEIPASS'))}
    env['PYINSTALLER_RESET_ENVIRONMENT'] = '1'   # a fresh program, not a helper of this one
    flags = getattr(subprocess, 'DETACHED_PROCESS', 0) | getattr(subprocess, 'CREATE_NEW_PROCESS_GROUP', 0)
    subprocess.Popen([str(Path(folder) / EXE_NAME), '--wait-pid', str(os.getpid())], cwd=str(folder), env=env,
                     creationflags=flags, close_fds=True)


def clean_up(folder):
    """Leftovers from an update that was cut off (the backup TTS_TF2_64BIT.exe.old is kept). Only the updater's
    own files - never a zip you downloaded yourself."""
    for name in ('update-download.part', EXE_NAME + '.new'):
        try:
            (Path(folder) / name).unlink()
        except OSError:
            pass
    for zip_path in Path(folder).glob('update-download-v*.zip'):
        try:
            zip_path.unlink()
        except OSError:
            pass
