# TF2Speech

Text-to-Speech for Team Fortress 2 using DECtalk, Windows voices and free Piper voices.

A spiritual successor to the original TF2Speech program, allowing players to communicate via text-to-speech through in-game voice chat.

![TF2Speech](https://img.shields.io/badge/TF2-Speech-orange) ![Python](https://img.shields.io/badge/Python-3.11+-blue) ![Windows](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## Features

- **DECtalk Voices** - Classic text-to-speech from the 90s (Moonbase Alpha style!)
- **SAPI5 Voices** - Windows built-in voices (Microsoft David, Zira, etc.)
- **Multiple Voice Profiles** - 9 DECtalk voices + all installed Windows voices
- **Chat Commands** - Users can change voices with `/v` commands
- **Admin Controls** - Block, mute and stop speech, manage permissions
- **Random Voice Assignment** - New users automatically get unique voices
- **Phoneme Support** - Full DECtalk phoneme control for songs and sound effects
- **DECtalk Auto-Switch** - DECtalk commands (`[:np]`, `[dah<600,20>]`...) sent with a normal voice are spoken with DECtalk instead of read out as gibberish
- **Mutes & Spam Guard** - Temporary mutes, plus automatic muting of copy-paste spam
- **Multiple Games** - Add any Source game (L4D2, GMod, CS:S...) with its own log file, admins and block list
- **DJ** *(new)* - Play music from your PC, your Subsonic/Navidrome library or YouTube into voice chat, with a queue, ducking under TTS and chat song requests (with optional approval)
- **Soundboard** *(new)* - Sven Co-op style chat sounds: type `yee`, everyone hears it
- **More voices** *(new)* - Make Windows' extra voices usable, and download free natural-sounding [Piper](https://github.com/rhasspy/piper) voices in dozens of languages with one click
- **Classic SAPI 4 voices** *(new, optional)* - If the old SAPI 4 voices are installed on your PC (Microsoft Sam 4.0, L&H TruVoice...), they work too - including a ready-made BonziBuddy voice
- **TextAloud voices** *(new, optional)* - If you own TextAloud 3 voices (e.g. Acapela Will From Afar), TF2Speech can use them through TextAloud, completely hidden
- **New look** *(new)* - A flat Light / Dark / Match Windows look, or the Classic look; every part resizes smoothly, tabs scroll instead of getting cut off, sharp at any Windows scaling

---

## Requirements

- Windows 10/11
- [VoiceMeeter](https://vb-audio.com/Voicemeeter/) (free virtual audio cable)
- Team Fortress 2
- Python 3.11+ (only if building from source)

Nothing else needs installing. The DJ and soundboard use two free tools that TF2Speech downloads for you with one click the first time you use them (see [Third-party tools](#third-party-tools)).

---

## Installation

### Option A: Download Pre-Built Release (Recommended)

1. Go to [Releases](https://github.com/Huskiefluff/Tf2Speech/releases)
2. Download the latest `.zip` file
3. Extract to any folder
4. Run `TTS_TF2_64BIT.exe`

Your settings are saved in `config.json` next to the exe.

### Option B: Build From Source

1. Clone the repository:
   ```
   git clone https://github.com/Huskiefluff/Tf2Speech.git
   ```

2. Run `build.bat` - it installs the Python dependencies and builds the exe into `dist\`.
   Rebuilding keeps everything you already have in `dist\` (settings, block lists, downloaded tools).

   Or do it by hand:
   ```
   pip install -r requirements.txt
   pyinstaller build_specs/build_64bit_final.spec
   ```

3. Find the executable in `dist/`

The small SAPI 4 helper `src/sapi4_tts.exe` comes ready-built. To rebuild it after changing `tools/sapi4_tts/sapi4_tts.cpp`, run `tools/sapi4_tts/build_sapi4_tts.bat` (needs the Visual Studio C++ build tools).

To run without building: `python src/main_32bit_simple.py`

---

## Setup Guide

### Step 1: Install VoiceMeeter

1. Download [VoiceMeeter](https://vb-audio.com/Voicemeeter/) (the free basic version works fine)
2. Install and restart your computer
3. Open VoiceMeeter at least once to initialize the virtual audio devices

### Step 2: Configure Windows Sound Settings

1. Open **Windows Sound Settings** (right-click speaker icon → Sound settings)
2. Set your normal speakers/headphones as the **default playback device**
3. Note that "VoiceMeeter Input" now exists as an audio device

### Step 3: Configure TF2Speech

1. Launch `TTS_TF2_64BIT.exe`
2. In the **Audio Devices** tab, select `VoiceMeeter Input`
3. In the **Settings** tab, set **Game Log File** to TF2's log, e.g.
   `C:\Program Files (x86)\Steam\steamapps\common\Team Fortress 2\tf\log.txt`
4. Configure voices as desired, then press **✓ Apply**

### Step 4: Configure Team Fortress 2

Copy the `tts.cfg` file from the `tf2_cfg` folder to your TF2 cfg folder:
```
Steam\steamapps\common\Team Fortress 2\tf\cfg\
```

Add this line to your `autoexec.cfg` (create it if it doesn't exist):
```
exec tts.cfg
```

**What tts.cfg does:**
- `con_logfile log.txt` - Creates the log file that TF2Speech reads
- `voice_loopback 1` - Lets you hear your own TTS output
- Sets up keybinds for TTS controls

### Step 5: Configure VoiceMeeter Routing

1. Open VoiceMeeter
2. Under **Hardware Input 1**, select your microphone
3. Under **Hardware Out (A1)**, select your speakers/headphones
4. Enable **A1** on the "Voicemeeter VAIO" strip (so TTS plays to your ears)
5. Enable **B1** on the "Voicemeeter VAIO" strip (so TTS goes to virtual mic)

### Step 6: Set TF2 Microphone Input

1. In TF2, go to **Options → Voice**
2. Set microphone to `VoiceMeeter Output`
3. Enable voice chat

---

## Usage

### In-Game Keybinds (Default)

| Key | Function |
|-----|----------|
| `Numpad Enter` | Toggle voice (push-to-talk) |
| `Numpad /` | Stop current TTS (!stop) |
| `Numpad -` | Block whoever is talking / last TTS user (!block add) |
| `Numpad +` | Add admin (!admin add) |
| `Numpad *` | Unblock the most recently blocked player (!block clear) |
| `Numpad .` | Mute whoever is talking (!mute) |
| `Numpad 7` / `Numpad 9` | Approve / deny the oldest song request (!approve / !deny) |
| `Numpad 8` | Skip the current song (!skipsong) |

**Tip - silent keys:** the default keys `say` the command, so the whole server sees `!stop`. If you'd rather nobody sees them, use the `echo` versions in `tts.cfg`: put your exact in-game name in them and they only reach your own TF2Speech.

### Chat Commands

**Basic TTS:**
```
!tts Hello world!
```

**Change Voice (temporary, one message):**
```
!tts /v 1 This uses voice 1
!tts /david This uses a custom trigger you made on the Voice Commands tab
```

**Set Your Permanent Voice:**
```
!tts /vt 5
```
(Your voice preference is saved for future messages)

The `!tts` prefix can be changed per game in the Settings tab.

### Voice List

**Voices are fully configurable in the Settings tab!** The available SAPI5 voices depend on what's installed on your Windows system. DECtalk voices are included with TF2Speech.

**Default Windows voices (most users have these):**
- Microsoft David Desktop - English (United States)
- Microsoft Zira Desktop - English (United States)

**DECtalk voices (included):**
- Perfect Paul, Harry, Frank, Betty, Wendy, Dennis, Kit, Ursula, Rita

**More voices:** see [Voices](#voices) - Windows' extra voices, free Piper voices (`[Piper] ...`) and classic SAPI 4 voices (`[SAPI4] ...`) show up in every voice list.

**Hiding voices:** in the Voice Commands tab, select voices in the list of available voices and press **Hide / Unhide Selected**. Hidden voices are marked *(hidden)* there, left out of every other voice list and never given out at random. Select them and press it again to bring them back.

**Configuring Voice Commands:**
1. Open TF2Speech → Voice Commands tab
2. Map `/v 0`, `/v 1`, `/v 2`, etc. to any available voice
3. You can also create named shortcuts like `/david`, `/harry`, etc.

**Example configuration:**
| Command | Voice |
|---------|-------|
| `/v 0` | Microsoft David |
| `/v 1` | Microsoft Zira |
| `/v 8` | [DECtalk] Perfect Paul |
| `/v 9` | [DECtalk] Harry |
| `/v 10` | [DECtalk] Betty |

*Your setup may vary depending on installed voices and personal preference.*

### Admin Commands

```
!stop          - Stop current speech (and any soundboard sounds)
!block         - Block whoever is talking (or the last TTS user)
!block add     - Same as !block
!block clear   - Unblock the most recently blocked player
!admin add     - Add the last TTS user as admin
!mute          - Mute whoever is talking for the default time (set in the MUTED panel)
!mute 30       - Mute them for 30 minutes
!unmute        - Lift the most recent mute
!skipsong      - Skip the DJ's current song
!approve       - Approve the oldest song request waiting for your OK
!deny          - Deny it
```

### Mutes & Spam Guard

Mutes are temporary blocks - they run out on their own. Mute from chat (`!mute`), or with **+ Mute** in the MUTED panel, which shows who is muted and for how long.

With **Auto-mute spam** on, a player who sends the same message 3 times in 30 seconds (`7777777`, `aaaaaa`...), or 5 messages in 10 seconds, is muted automatically and their queued messages are dropped. Admins are never auto-muted. The announcement ("... muted for spamming") can be changed on the Announcements tab.

### Multiple Games

TF2Speech reads any Source engine game's console log. In the Settings tab use **+ Add Game** and give it:
- a name
- the log type (Source game, or Deep Rock Galactic chat mod)
- its log file - add `con_logfile log.txt` to that game's `autoexec.cfg`, then pick the `log.txt` it creates (e.g. `...\left4dead2\log.txt`)
- its TTS command

Each game keeps its own admins, block list and TTS command.

---

## DJ

Press **🎵 DJ** to open the music player. Music plays through the same output as the TTS, so the server hears it while your voice is on.

- **Your music** - add files or whole folders (mp3, flac, m4a, ogg, wav...)
- **Your library** - connect a Subsonic / Navidrome server (only a login token is saved, never the password): search, playlists, random songs
- **YouTube** - paste a link, or type words to search and pick a result
- **Queue** - reorder, shuffle, repeat, seek; the queue is saved between runs; adding to an empty queue starts playing (can be turned off)
- **Ducking** - the music drops while TTS is talking, then comes back up
- **Even loudness** - quiet and loud songs are levelled out; a **Bass** slider tames TF2's boomy voice codec
- **"Now playing"** - optional announcement through the TTS, in a DJ voice of your choice
- **Song requests** - `!sr <YouTube link or song name>` in chat. Off / admins only / everyone, from YouTube or your library, with a length limit and a limit on waiting songs per person. Players hear whether their request was accepted or denied. Optional **approval**: requests wait in the Requests tab (or `!approve` / `!deny`) until you OK them. Only YouTube links are accepted (any other link would make your PC load a page a stranger chose).
- **Keyboard** - Space play/pause, Ctrl+→ next, Ctrl+← previous, Ctrl+S stop, Del remove, Ctrl+↑/↓ move, Ctrl+O add files, Ctrl+F YouTube box

---

## Soundboard

Press **🔊 Sounds** to set it up. Inspired by Sven Co-op's ChatSounds plugin.

- Put sound files in the `sounds` folder next to the exe - the file name is the word (`yee.wav` → `yee`). Sub-folders work, and a Sven-style `chatsounds.txt` (lines like `bleh twlz/bleh.wav`) is read too.
- **Just the word** mode: a chat message that is exactly `yee` plays it. **Command** mode: `!s yee` (you pick the command).
- Pitch: `yee 150` plays faster/higher, `yee 70` slower/lower (50-200).
- Sounds play immediately - never queued - over the music (or lowering it, your choice).
- Per-player cooldown, a gap between any two sounds, and a limit on how many play at once.
- Admins only or everyone; blocked and muted players can't use it. `!stop` cuts sounds off.

---

## Voices

The **Voices** tab gets you more voices:

- **Windows voices** - Windows can add voices and languages (Settings → Time & language → Speech → Add voices). Windows keeps those apart from the voices programs like TF2Speech can use, so after adding some press **Make them usable here** - Windows asks for admin permission once - and restart TF2Speech. **Undo** removes them again. Many PCs already have a few of these (e.g. Microsoft Mark).
- **Free voices (Piper)** - natural-sounding neural voices that run on your PC, in dozens of languages. Pick one, press **Download** (20-130 MB each; the small Piper program is fetched the first time) and **Test**. They appear in every voice list as `[Piper] en_US-ryan-medium` and work like the other voices - `!stop`, mutes, the output device and the music ducking all apply. Downloads are checked against the voice list's checksums.

### Classic SAPI 4 voices (optional)

The old SAPI 4 voices from the late 90s - Microsoft Sam / Mike / Mary 4.0 and their variants, Lernout & Hauspie TruVoice (BonziBuddy's voice) and others - work too, if they are installed on your PC. They appear in every voice list as `[SAPI4] ...`, plus a ready-made **`[SAPI4] Bonzi`** (TruVoice Adult Male #2 at BonziBuddy's pitch and speed). The **Voices** tab shows what was found and has a **Test** button and a volume slider.

You need the Microsoft SAPI 4.0 runtime (`spchapi.exe`) and the voice packages (for example `tv_enua.exe` for TruVoice) installed. These used to be free Microsoft Agent downloads; Microsoft no longer hosts them, and TF2Speech doesn't include or download them. Without SAPI 4 the Voices tab just says it isn't installed.

SAPI 4 voices are 32-bit only, so TF2Speech uses a small 32-bit helper of its own (`sapi4_tts.exe`, source in [tools/sapi4_tts](tools/sapi4_tts), MIT) that says each line into a sound file, which plays like DECtalk - `!stop`, mutes, the output device and the music ducking all apply. SAPI 4 codes typed in chat (like `\Pau=...\`) are read out as text, not obeyed.

### TextAloud voices (optional)

Some voices sold for TextAloud 3 (for example Acapela "Will From Afar" bought from NextUp) only run inside TextAloud. If TextAloud 3 is installed, a **TextAloud** tab appears: TF2Speech asks TextAloud to say each line into a sound file and plays it like DECtalk. TextAloud runs on a separate hidden Windows desktop, so its window and any error boxes never appear over your game; if it crashes or freezes it is restarted, and the line falls back to your default voice if needed. Nothing from TextAloud is included - you need your own licensed copy. Without TextAloud this is all invisible.

---

## Appearance

**Settings → Appearance**: Match Windows, Light, Dark (a flat modern look) or Classic (the original grey look). Drag the dividers between the chat log, the speak box, the tabs and the lists to resize them; a tab that doesn't fit scrolls. The window remembers its size, position and dividers.

---

## Third-party tools

The DJ, soundboard and Piper voices need a few free tools. They are **not** included in the download - TF2Speech fetches them for you the first time (one click) and keeps them next to the exe. Nothing is installed in Windows.

| Tool | Used for | Downloaded from | License |
|------|----------|-----------------|---------|
| [FFmpeg](https://ffmpeg.org/) | playing music and sounds | [BtbN/FFmpeg-Builds](https://github.com/BtbN/FFmpeg-Builds) (LGPL build, ~80 MB) | LGPL |
| [yt-dlp](https://github.com/yt-dlp/yt-dlp) | finding YouTube audio | the official yt-dlp release | Unlicense |
| [Piper](https://github.com/rhasspy/piper) | free voices | the official Piper release (~22 MB) | MIT |
| [Piper voices](https://huggingface.co/rhasspy/piper-voices) | free voices | the official Piper voice list | each voice's own licence (see its MODEL_CARD) |

If FFmpeg is already installed on your PC (e.g. `winget install ffmpeg`), TF2Speech uses that instead. YouTube changes break yt-dlp now and then - the **Update yt-dlp** button fixes that.

---

## DECtalk Phoneme Fun

DECtalk supports phoneme commands for songs and sound effects!

**Change voice mid-sentence:**
```
!tts [:np]Paul says hi [:nb]Betty says hello
```

**Sing a note:**
```
!tts [ah<500,25>]
```
(500ms duration, pitch 25)

**Classic sounds:**
```
!tts soi soi soi soi          - Robot sound
!tts dfdfdfdfdfdfdfdfe        - Helicopter
!tts aeiou                    - Classic Moonbase Alpha
```

**The Gaben Song:**
```
!tts [ih<300,20>ts<100>tweh<300,22>n<100>tiy<300,25>tweh<300,25>n<100>tiy<300,27>sih<600,30>ks]
```

---

## Troubleshooting

### TTS not speaking
- Make sure VoiceMeeter is running
- Check that the Audio Device is set to "VoiceMeeter Input" in TF2Speech
- Verify the Game Log File path matches your TF2 installation (`...\tf\log.txt`)

### Others can't hear me
- In TF2 voice settings, make sure mic is set to "VoiceMeeter Output"
- Check VoiceMeeter B1 is enabled on the VAIO strip
- Make sure you're pressing the voice key (Numpad Enter by default)

### DECtalk not working
- DECtalk only works on Windows
- Make sure the `voice_data/dectalk` folder exists with the binaries
- Try SAPI5 voices as a fallback

### Voice sounds choppy
- `voice_buffer_ms` in tts.cfg only changes how *you* hear other players (it's the incoming buffer) - 100 is a good value; lowering it makes others sound choppier to you and doesn't change what they hear from you
- If others hear your TTS cut off, set TF2's voice to push-to-talk / open mic with the voice threshold off
- Check CPU usage — close background applications

### DJ or soundboard silent
- Click **Download ffmpeg** in the DJ window (one time)
- Check the DJ's **Output** - "Same as the TTS" sends it to VoiceMeeter

### YouTube songs won't load
- Click **Update yt-dlp** in the DJ window

### Something went wrong
- Everything is logged to `tts_tf2.log` next to the exe

---

## What's new in v1.1.0

- **New:** flat Light / Dark / Match Windows look or Classic; resizable sections, smooth resizing, scrolling tabs, remembers its size, sharp at any Windows scaling
- **New:** Voices tab - make Windows' extra voices usable, download free Piper voices
- **New:** optional classic SAPI 4 voices (Microsoft Sam 4.0, TruVoice...) with a ready-made `[SAPI4] Bonzi`
- **New:** optional TextAloud voices (e.g. Acapela Will From Afar), with TextAloud hidden away and crash-proofed
- **New:** DJ window (local music, Subsonic/Navidrome library, YouTube, queue, ducking, bass control, "now playing", chat song requests with spoken confirmations, approval and per-person limits)
- **New:** Chat soundboard (Sven Co-op ChatSounds style)
- **New:** Help tab rewritten to match the real commands
- **Hide / Unhide Selected** really hides voices now - from every voice list and from random voices
- Muting or blocking someone cuts their current line off immediately (Windows voices too)
- Admin commands still work for an admin who is muted
- **New:** Mutes (`!mute` / `!unmute`), MUTED panel and automatic spam muting
- **New:** Add/edit/remove games yourself
- **New:** DECtalk auto-switch for DECtalk commands sent with a normal voice
- **New:** no console window; startup errors show in a pop-up
- Spectators (`*SPEC*`) are now matched against blocks, mutes and voices
- Status-command and material-error lines no longer clutter the chat window
- Windows voices read text literally (no SAPI markup tricks like minutes of silence)
- Each queued message keeps its own voice
- Settings are saved safely (no half-written config.json; a backup is kept and used if the config is damaged)
- **✓ Apply** saves everything on the Settings tab; Import/Export keep all settings
- DECtalk volume slider now also applies to profile tests; goes up to 150%
- Old temp files from crashed runs are cleaned up automatically
- Many smaller fixes

---

## Credits

- **Original TF2Speech** - The program this is based on (RIP)
- **DECtalk** - Originally by Digital Equipment Corporation, [preserved on GitHub](https://github.com/dectalk/dectalk) thanks to developer Edward Bruckert sharing the source code
- **Moonbase Alpha** - For keeping DECtalk memes alive
- **VB-Audio** - For VoiceMeeter
- **Sven Co-op ChatSounds** - Inspiration for the soundboard
- **FFmpeg** and **yt-dlp** - Downloaded on demand for the DJ and soundboard
- **Piper** by the Rhasspy project - Free neural voices, downloaded on demand

---

## Legal

- DECtalk binaries are sourced from the [dectalk GitHub repository](https://github.com/dectalk/dectalk), originally shared by developer Edward Bruckert in 2015
- SAPI5 voices are property of Microsoft and included with Windows; TF2Speech doesn't distribute any voices
- FFmpeg (LGPL), yt-dlp (Unlicense), Piper (MIT) and Piper voices are not distributed with TF2Speech; they are downloaded from their official sources on request
- SAPI 4 (Microsoft) and SAPI 4 voices (Microsoft, Lernout & Hauspie) are not included; TF2Speech's own helper only uses voices already installed on your PC
- TextAloud and voices bought for it are not included; TF2Speech only talks to a copy you installed and licensed yourself, through TextAloud's own automation interface
- This project is not affiliated with Valve or Team Fortress 2

---

## License

MIT License - See [LICENSE](LICENSE) file

---

## Support

Having issues? Open an [Issue](https://github.com/Huskiefluff/Tf2Speech/issues) on GitHub.

Want to contribute? Pull requests welcome!

---

*"Entire team is babies."* 🎤
