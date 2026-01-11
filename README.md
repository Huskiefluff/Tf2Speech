# TF2Speech

Text-to-Speech for Team Fortress 2 using DECtalk and Windows SAPI5 voices.

A spiritual successor to the original TF2Speech program, allowing players to communicate via text-to-speech through in-game voice chat.

![TF2Speech](https://img.shields.io/badge/TF2-Speech-orange) ![Python](https://img.shields.io/badge/Python-3.10+-blue) ![Windows](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## Features

- **DECtalk Voices** - Classic text-to-speech from the 90s (Moonbase Alpha style!)
- **SAPI5 Voices** - Windows built-in voices (Microsoft David, Zira, etc.)
- **Multiple Voice Profiles** - 9 DECtalk voices + all installed Windows voices
- **Chat Commands** - Users can change voices with `/v` commands
- **Admin Controls** - Block users, stop speech, manage permissions
- **Random Voice Assignment** - New users automatically get unique voices
- **Phoneme Support** - Full DECtalk phoneme control for songs and sound effects

---

## Requirements

- Windows 10/11
- [VoiceMeeter](https://vb-audio.com/Voicemeeter/) (free virtual audio cable)
- Team Fortress 2
- Python 3.10+ (only if building from source)

---

## Installation

### Option A: Download Pre-Built Release (Recommended)

1. Go to [Releases](https://github.com/Huskiefluff/Tf2Speech/releases)
2. Download the latest `.zip` file
3. Extract to any folder
4. Run `TF2Speech.exe`

### Option B: Build From Source

1. Clone the repository:
   ```
   git clone https://github.com/Huskiefluff/Tf2Speech.git
   ```

2. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Build with PyInstaller:
   ```
   pyinstaller build_specs/build_64bit_final.spec
   ```

4. Find the executable in `dist/`

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

1. Launch `TF2Speech.exe`
2. Go to the **Settings** tab
3. Set **Audio Output Device** to `VoiceMeeter Input`
4. Set your **TF2 Directory** (e.g., `C:\Program Files (x86)\Steam\steamapps\common\Team Fortress 2`)
5. Configure voices as desired

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
| `Numpad -` | Block last speaker (!block add) |
| `Numpad +` | Add admin (!admin add) |
| `Numpad *` | Clear block list (!block clear) |

### Chat Commands

**Basic TTS:**
```
!tts Hello world!
```

**Change Voice (temporary, one message):**
```
!tts /v 1 This uses voice 1
!tts /v david This uses Microsoft David
```

**Set Your Permanent Voice:**
```
!tts /vt 5
```
(Your voice preference is saved for future messages)

### Voice List

**Voices are fully configurable in the Settings tab!** The available SAPI5 voices depend on what's installed on your Windows system. DECtalk voices are included with TF2Speech.

**Default Windows voices (most users have these):**
- Microsoft David Desktop - English (United States)
- Microsoft Zira Desktop - English (United States)

**DECtalk voices (included):**
- Perfect Paul, Harry, Frank, Betty, Wendy, Dennis, Kit, Ursula, Rita

**Configuring Voice Commands:**
1. Open TF2Speech → Settings tab
2. Find the Voice Commands section
3. Map `/v 0`, `/v 1`, `/v 2`, etc. to any available voice
4. You can also create named shortcuts like `/v david`, `/v harry`, etc.

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
!stop          - Stop current speech
!block add     - Block the last speaker
!block remove  - Unblock the last speaker  
!block clear   - Clear all blocks
!admin add     - Add last speaker as admin
!admin remove  - Remove last speaker as admin
```

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
- Check that Audio Output is set to "VoiceMeeter Input" in TF2Speech settings
- Verify the log file path matches your TF2 installation

### Others can't hear me
- In TF2 voice settings, make sure mic is set to "VoiceMeeter Output"
- Check VoiceMeeter B1 is enabled on the VAIO strip
- Make sure you're pressing the voice key (Numpad Enter by default)

### DECtalk not working
- DECtalk only works on Windows
- Make sure the `voice_data/dectalk` folder exists with the binaries
- Try SAPI5 voices as a fallback

### Voice sounds choppy
- Lower the `voice_buffer_ms` value in tts.cfg
- Check CPU usage — close background applications

---

## Credits

- **Original TF2Speech** - The program this is based on (RIP)
- **DECtalk** - Originally by Digital Equipment Corporation, [preserved on GitHub](https://github.com/dectalk/dectalk) thanks to developer Edward Bruckert sharing the source code
- **Moonbase Alpha** - For keeping DECtalk memes alive
- **VB-Audio** - For VoiceMeeter

---

## Legal

- DECtalk binaries are sourced from the [dectalk GitHub repository](https://github.com/dectalk/dectalk), originally shared by developer Edward Bruckert in 2015
- SAPI5 voices are property of Microsoft and included with Windows
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
