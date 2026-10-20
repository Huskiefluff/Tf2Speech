"""
Make chat text speakable the same way for every voice.

Emoji and text smileys: Windows' newer voices (Microsoft David, Mark) say emoji by name - "slightly smiling face" -
but DECtalk, the SAPI 4 voices, the older SAPI 5 voices (Sam, Mike, Mary), Zira, Piper and TextAloud skip them, say
"question mark", or read ":)" as punctuation. speakable() turns them into words before any voice gets the text,
using the names Windows uses (Unicode CLDR), so they sound the same whichever voice someone has.

Abbreviations: Piper pronounces short all-caps abbreviations as if they were words ("PST" -> "pst"), so
spell_abbreviations() spaces them out ("P S T") for Piper.
"""

import re
import unicodedata

# Emoji whose Unicode character name differs from the name Windows/CLDR uses - the common ones in chat.
# Everything else uses its Unicode name, which for newer emoji is the same as the CLDR name.
NAMES = {
    '\U0001F603': 'grinning face with big eyes', '\U0001F604': 'grinning face with smiling eyes',
    '\U0001F601': 'beaming face with smiling eyes', '\U0001F606': 'grinning squinting face',
    '\U0001F605': 'grinning face with sweat', '\U0001F60D': 'smiling face with heart eyes',
    '\U0001F618': 'face blowing a kiss', '\U0001F61C': 'winking face with tongue', '\U0001F61B': 'face with tongue',
    '\U0001F61D': 'squinting face with tongue', '\U0001F613': 'downcast face with sweat',
    '\U0001F97A': 'pleading face', '\U0001F92C': 'face with symbols on mouth', '\U0001F92E': 'face vomiting',
    '\U0001F975': 'hot face', '\U0001F976': 'cold face', '\U0001F928': 'face with raised eyebrow',
    '\U0001F92A': 'zany face', '\U0001F929': 'star struck', '\U0001F973': 'partying face',
    '\U0001F92F': 'exploding head', '\U0001F92D': 'face with hand over mouth', '\U0001F92B': 'shushing face',
    '\U0001F47F': 'angry face with horns', '☺': 'smiling face', '☹': 'frowning face',
    '\U0001F44D': 'thumbs up', '\U0001F44E': 'thumbs down', '\U0001F44C': 'OK hand', '\U0001F44B': 'waving hand',
    '\U0001F44F': 'clapping hands', '\U0001F64F': 'folded hands', '\U0001F595': 'middle finger',
    '\U0001F449': 'backhand index pointing right', '\U0001F448': 'backhand index pointing left',
    '\U0001F446': 'backhand index pointing up', '\U0001F447': 'backhand index pointing down',
    '☝': 'index pointing up', '\U0001F937': 'person shrugging', '\U0001F926': 'person facepalming',
    '❤': 'red heart', '♥': 'heart suit', '\U0001F4AF': 'hundred points', '⭐': 'star',
    '✅': 'check mark button', '❗': 'red exclamation mark', '❓': 'red question mark',
    '⚡': 'high voltage', '☀': 'sun', '\U0001F52B': 'water pistol', '\U0001F346': 'eggplant',
    '\U0001F3B6': 'musical notes', '\U0001F355': 'pizza', '\U0001F63A': 'grinning cat',
}

# Text smileys, said like Windows says them. Only whole ones count (a space or the line's start before, a space,
# punctuation or the end after), so "http://", "12:30" and DECtalk commands like [:np] are left alone.
SMILEYS = {
    'smiley face': [':)', ':-)', '(:', '=)', ':]'],
    'sad face': [':(', ':-(', '):', '=(', ':['],
    'grinning face': [':D', ':-D', '=D'],
    'laughing face': ['xD', 'XD'],
    'winking face': [';)', ';-)'],
    'face with tongue': [':P', ':p', ':-P', ':-p'],
    'surprised face': [':O', ':o', ':-O', ':-o'],
    'confused face': [':/', ':-/', ':\\'],
    'neutral face': [':|', ':-|'],
    'crying face': [":'(", 'T_T'],
    'happy face': ['^_^', '^^'],
    'expressionless face': ['-_-'],
    'heart': ['<3'],
    'broken heart': ['</3'],
}
_SMILEY_WORDS = {s: words for words, faces in SMILEYS.items() for s in faces}
_SMILEY = re.compile(r'(?<!\S)(' + '|'.join(re.escape(s) for s in sorted(_SMILEY_WORDS, key=len, reverse=True))
                     + r')(?=$|[\s.,!?])')

# Real emoji blocks only - not braille, box drawing or block elements, which chat "art" is made of
_EMOJI_RANGES = ((0x1F000, 0x1FAFF), (0x2600, 0x27BF), (0x2B00, 0x2BFF), (0x2300, 0x23FF), (0x2190, 0x21FF))
_JOINERS = {'‍', '︎', '️', '⃣'}           # zero-width joiner, variation selectors, keycap
_SKIN_TONES = range(0x1F3FB, 0x1F400)
_TAGS = range(0xE0020, 0xE0080)                                 # subdivision flag tags


def _is_emoji(ch):
    cp = ord(ch)
    return any(lo <= cp <= hi for lo, hi in _EMOJI_RANGES) and unicodedata.category(ch) == 'So'


def _emoji_name(ch):
    if ch in NAMES:
        return NAMES[ch]
    if 0x1F1E6 <= ord(ch) <= 0x1F1FF:                       # regional indicator: one letter of a flag
        return chr(ord(ch) - 0x1F1E6 + ord('A'))
    try:
        return unicodedata.name(ch).lower().replace(' sign', '')
    except ValueError:
        return ''


MAX_EMOJI_NAMES = 8  # per message, so emoji "art" (rows of coloured squares...) doesn't read on for a minute


def speakable(text):
    """Emoji and text smileys -> words; a run of the same emoji is said once, and at most MAX_EMOJI_NAMES emoji
    are said per message. Everything else is unchanged."""
    if not text:
        return text
    text = _SMILEY.sub(lambda m: _SMILEY_WORDS[m.group(1)], text)
    out, last, flag, said = [], None, '', 0
    for ch in text:
        cp = ord(ch)
        if ch in _JOINERS or cp in _SKIN_TONES or cp in _TAGS:
            continue
        if 0x1F1E6 <= cp <= 0x1F1FF:                        # flags are two regional-indicator letters
            flag += _emoji_name(ch)
            if len(flag) == 2:
                if said < MAX_EMOJI_NAMES:
                    out.append(f' flag {flag[0]} {flag[1]} ')
                    said += 1
                flag, last = '', None
            continue
        if _is_emoji(ch):
            if ch != last and said < MAX_EMOJI_NAMES:
                name = _emoji_name(ch)
                if name:
                    out.append(f' {name} ')
                    said += 1
            last = ch
            continue
        if not ch.isspace():
            last = None
        out.append(ch)
    return re.sub(r' {2,}', ' ', ''.join(out)).strip()


# Abbreviations Piper should spell out even though they contain a vowel (without a vowel it spells them anyway)
_SPELL_WITH_VOWEL = {'EST', 'EDT', 'UTC', 'AFK', 'IRL', 'IMO', 'IDK', 'OMG', 'USA', 'UK', 'EU', 'AI', 'FPS',
                     'MVP', 'TTS', 'PVE', 'PVP', 'IGN', 'TBH', 'ETA', 'AOE', 'URL', 'PC', 'NPC', 'XP', 'UI'}
_VOWELS = set('AEIOUY')
_ABBREVIATION = re.compile(r'\b[A-Z][A-Z0-9]{1,4}\b')


def spell_abbreviations(text):
    """'EST PST TF2' -> 'E S T P S T T F 2' for Piper. Only capitals with no vowel (PST, GG, TF2) or a known
    abbreviation (EST, AFK...) are spelled, so shouted words like 'HELLO' or 'NO' are still said as words."""
    def spell(m):
        word = m.group(0)
        if word in _SPELL_WITH_VOWEL or not (set(word) & _VOWELS):
            return ' '.join(word)
        return word
    return _ABBREVIATION.sub(spell, text)
