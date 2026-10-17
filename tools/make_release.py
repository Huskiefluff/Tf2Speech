r"""
Makes the release zip after build.bat: a "release" folder with the exe, tts.cfg, the notes in release_files\,
LICENSE.txt and THIRD-PARTY-NOTICES.txt - the licences of everything PyInstaller built into the exe, found by
reading the exe itself, so the list always matches what's really inside.

Run it with the same Python that built the exe:
    C:\Python313\python.exe tools\make_release.py 1.1.0
The zip lands in dist_release\TF2Speech_v1.1.0.zip. Nothing from dist\ except the exe goes in (no config.json,
logs, downloaded tools or voices).
"""

import sys
import hashlib
import zipfile
import collections
from pathlib import Path
from importlib.metadata import packages_distributions, distribution

from PyInstaller.archive.readers import CArchiveReader

ROOT = Path(__file__).resolve().parent.parent
EXE = ROOT / 'dist' / 'TTS_TF2_64BIT.exe'
NOTES = ['READ ME FIRST.txt', 'VOICES.txt', 'CHANGELOG.txt']
LICENCE_WORDS = ('licen', 'copying', 'notice')

PORTAUDIO = """\
PortAudio Portable Real-Time Audio Library (built into PyAudio)
Copyright (c) 1999-2011 Ross Bencina and Phil Burk

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and
to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions
of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO
THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
IN THE SOFTWARE.

The text above constitutes the entire PortAudio license; however, the PortAudio community also makes the
following non-binding requests: Any person wishing to distribute modifications to the Software is requested to
send the modifications to the original developer so that they can be incorporated into the canonical version.
It is also requested that these non-binding requests be included along with the license above.
"""

HEADER = """\
TF2Speech v{version} - third-party notices
{line}

TF2Speech's own code is MIT licensed (LICENSE.txt). The exe also contains the parts below, each under its own
licence. Their full licence texts follow the list.

  DECtalk         Speech engine files as published by the DECtalk project, https://github.com/dectalk/dectalk
                  (its source code was shared by its late original developer Edward Bruckert in 2015).
                  DECtalk was made by Digital Equipment Corporation; copyright FONIX Corporation and earlier
                  owners - see the LICENCE file in that repository. Included unmodified. If you hold rights to
                  DECtalk and object, please open an issue at https://github.com/Huskiefluff/Tf2Speech/issues
  sapi4_tts.exe   TF2Speech's own SAPI 4 helper (MIT, source in tools/sapi4_tts), built with Microsoft Visual
                  C++. No SAPI 4 or voice files from Microsoft or Lernout & Hauspie are included.
  PyInstaller     The exe's start-up code (the "bootloader"), GPL 2.0 with an exception that allows programs
                  built with it to be shared under any licence:
                  https://github.com/pyinstaller/pyinstaller/blob/v{pyinstaller}/COPYING.txt
  Python {python:<8} PSF License. Its licence below also covers the libraries in Python's Windows build
                  (OpenSSL, libffi, bzip2, Tcl/Tk...); the full list of software incorporated into Python is at
                  https://docs.python.org/{pyshort}/license.html
  Microsoft C++   VCRUNTIME140.dll and VCRUNTIME140_1.dll, Microsoft Distributable Code that comes with Python
  runtime         (see "Additional Conditions for this Windows binary build" in Python's licence below)
  PortAudio       MIT, built into PyAudio
{packages}

Source code for the MPL-2.0 licensed pyttsx3 is available from https://pypi.org/project/pyttsx3/{pyttsx3}/ and
https://github.com/nateshmbhat/pyttsx3 (it is included unmodified).
"""


def bundled_distributions(exe):
    """Third-party distributions (name -> import names) that PyInstaller put into the exe"""
    arch = CArchiveReader(str(exe))
    tops = set()
    for name, (_pos, _csize, _usize, _cflag, typ) in arch.toc.items():
        if typ == 'z':
            tops |= {m.split('.')[0] for m in arch.open_embedded_archive(name).toc}
        elif typ in ('b', 'x', 'n'):
            first = name.replace('/', '\\').split('\\')[0]
            tops.add(first[:-5] if first.endswith('.libs') else first)
        elif typ in ('m', 's', 'M'):
            tops.add(name.split('.')[0])
    names = packages_distributions()
    found = collections.defaultdict(set)
    for top in tops:
        for dist in names.get(top, []):
            found[dist].add(top)
    found.pop('pyinstaller', None)  # build tool: only its bootloader is in the exe (noted above)
    found.pop('pyinstaller-hooks-contrib', None)
    for dist in found:  # folders in the exe count too (pywin32's modules sit in win32\ and Pythonwin\)
        found[dist] |= tops
    return found


def licence_files(dist, tops):
    files = [f for f in (dist.files or []) if any(w in f.name.lower() for w in LICENCE_WORDS)]
    in_info = [f for f in files if '.dist-info' in str(f)]
    # numpy keeps the licences of the libraries it brings along in its main licence file; pywin32 has none in
    # .dist-info, only in its package folders - take those of the folders that are really in the exe
    bundled = {t.lower() for t in tops}
    chosen = in_info or [f for f in files if str(f).replace('\\', '/').split('/')[0].lower() in bundled
                         and '/include/' not in str(f).replace('\\', '/')]
    return [Path(dist.locate_file(f)) for f in chosen]


def notices(version):
    import PyInstaller
    dists = bundled_distributions(EXE)
    rows, texts, seen = [], [], set()
    for name in sorted(dists, key=str.lower):
        dist = distribution(name)
        meta = dist.metadata
        lic = meta.get('License-Expression') or (meta.get('License') or '').split('\n')[0][:40]
        if not lic or len(lic) > 30:
            lic = ', '.join(c.split('::')[-1].strip() for c in (meta.get_all('Classifier') or [])
                            if c.startswith('License')) or 'see below'
        rows.append(f"  {dist.metadata['Name']:<15} {dist.version:<9} {lic}")
        for path in licence_files(dist, dists[name]):
            text = path.read_text(encoding='utf-8', errors='replace').strip()
            digest = hashlib.sha256(text.encode()).hexdigest()
            if digest in seen:
                continue
            seen.add(digest)
            texts.append((f"{dist.metadata['Name']} {dist.version} - {path.name}", text))
    python_licence = Path(sys.base_prefix) / 'LICENSE.txt'
    texts.insert(0, (f"Python {sys.version.split()[0]} - LICENSE.txt", python_licence.read_text(encoding='utf-8').strip()))
    texts.append(("PortAudio", PORTAUDIO.strip()))
    head = HEADER.format(version=version, line='=' * (len(version) + 34), pyinstaller=PyInstaller.__version__,
                         python=sys.version.split()[0], pyshort='.'.join(sys.version.split('.')[:2]),
                         packages='\n'.join(rows), pyttsx3=distribution('pyttsx3').version if 'pyttsx3' in dists else '')
    if 'pyttsx3' not in dists:
        head = head[:head.index('Source code for the MPL-2.0')]
    body = '\n\n'.join(f"{'-' * 100}\n{title}\n{'-' * 100}\n\n{text}" for title, text in texts)
    return head + '\n\n' + body + '\n'


def crlf(text):
    return text.replace('\r\n', '\n').replace('\n', '\r\n').encode('utf-8-sig')


def main():
    if len(sys.argv) != 2:
        sys.exit("usage: make_release.py <version>, e.g. 1.1.0")
    version = sys.argv[1]
    if not EXE.exists():
        sys.exit(f"{EXE} not found - run build.bat first")
    out_dir = ROOT / 'dist_release'
    out_dir.mkdir(exist_ok=True)
    zip_path = out_dir / f'TF2Speech_v{version}.zip'
    files = {
        'TTS_TF2_64BIT.exe': EXE.read_bytes(),
        'tts.cfg': (ROOT / 'tf2_cfg' / 'tts.cfg').read_bytes(),
        'LICENSE.txt': crlf((ROOT / 'LICENSE').read_text(encoding='utf-8')),
        'THIRD-PARTY-NOTICES.txt': crlf(notices(version)),
    }
    for note in NOTES:
        files[note] = crlf((ROOT / 'release_files' / note).read_text(encoding='utf-8').replace('{version}', version))
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        z.writestr(zipfile.ZipInfo('release/'), b'')
        for name, data in files.items():
            z.writestr(f'release/{name}', data)
    for name, data in files.items():
        print(f"  {len(data):>12,}  release/{name}")
    digest = hashlib.sha256(zip_path.read_bytes()).hexdigest()
    print(f"{zip_path}\n  {zip_path.stat().st_size:,} bytes\n  SHA-256 {digest}")


if __name__ == '__main__':
    main()
