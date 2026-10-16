"""
Look and layout helpers for the TTS windows

- Windows 11 style (Sun Valley ttk theme, MIT) in Light, Dark or "match Windows", with dark title bars
- Sharp text at 125%/150% Windows scaling, and window/dialog sizes that scale with it
- ScrollFrame: a tab or panel that scrolls when the window is too small for it, instead of cutting it off
"""

import os
import sys
import ctypes
import logging
import tkinter as tk
from tkinter import ttk

try:
    import sv_ttk
except Exception:  # the app still works with the plain look
    sv_ttk = None

logger = logging.getLogger(__name__)

# Colours that go with the Sun Valley theme, for the classic Tk widgets it doesn't style itself
PALETTE = {
    'light': {'bg': '#fafafa', 'field': '#ffffff', 'fg': '#1c1c1c', 'border': '#d9d9d9', 'hint': '#6b6b6b',
              'select': '#2f60d8', 'selectfg': '#ffffff', 'good': '#d9f2df', 'sash': '#e5e5e5'},
    'dark': {'bg': '#1c1c1c', 'field': '#2b2b2b', 'fg': '#fafafa', 'border': '#3a3a3a', 'hint': '#a0a0a0',
             'select': '#2f60d8', 'selectfg': '#ffffff', 'good': '#1f4d2b', 'sash': '#2b2b2b'},
    # the original look: Tk's plain grey 'default' theme with Arial headings
    'classic': {'bg': '#d9d9d9', 'field': '#ffffff', 'fg': '#000000', 'border': '#a0a0a0', 'hint': '#555555',
                'select': '#0078d7', 'selectfg': '#ffffff', 'good': '#90ee90', 'sash': '#c8c8c8'},
}
CLASSIC_FIELDS = ('Listbox', 'Text')   # input-like widgets that get the "field" shade
_current = {'mode': 'light'}


# ---------------------------------------------------------------- scaling
def enable_dpi_awareness():
    """Tell Windows the app draws itself at the real DPI, so text is sharp at 125%/150% instead of a blurry
    stretched bitmap. Must run before the first window is created."""
    if sys.platform != 'win32':
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def scale_of(widget):
    """1.0 at 100% Windows scaling, 1.5 at 150%... (TTS_UI_SCALE overrides it, for testing)"""
    forced = os.environ.get('TTS_UI_SCALE')
    if forced:
        return float(forced)
    try:
        return max(1.0, widget.winfo_fpixels('1i') / 96.0)
    except Exception:
        return 1.0


def apply_forced_scale(root):
    """TTS_UI_SCALE=1.5 makes the app draw as if Windows were at 150% (for checking layouts)"""
    forced = os.environ.get('TTS_UI_SCALE')
    if forced:
        root.tk.call('tk', 'scaling', float(forced) * 96 / 72)


def size_window(win, width, height, parent=None, minsize=None):
    """Size a window in 100%-scaling pixels (scaled for the screen), kept inside the screen, centred over parent"""
    s = scale_of(win)
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    w = min(int(width * s), int(sw * 0.95))
    h = min(int(height * s), int(sh * 0.90))
    if minsize:
        win.minsize(min(int(minsize[0] * s), w), min(int(minsize[1] * s), h))
    x = y = None
    if parent is not None:
        try:
            parent.update_idletasks()
            x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - h) // 3
        except Exception:
            x = y = None
    if x is None:
        x, y = (sw - w) // 2, (sh - h) // 3
    x = max(0, min(x, sw - w))
    y = max(0, min(y, sh - h - 40))
    win.geometry(f"{w}x{h}+{x}+{y}")
    win.after_idle(lambda: _grow_to_fit(win))  # once its contents are in


def _grow_to_fit(win):
    """Make a window at least as big as its contents need (dialogs had fixed sizes that bigger text outgrew)"""
    try:
        win.update_idletasks()
        need_w, need_h = win.winfo_reqwidth(), win.winfo_reqheight()
        w, h = win.winfo_width(), win.winfo_height()
        if need_w > w or need_h > h:
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            w, h = min(max(w, need_w), int(sw * 0.95)), min(max(h, need_h), int(sh * 0.90))
            x = max(0, min(win.winfo_x(), sw - w))
            y = max(0, min(win.winfo_y(), sh - h - 40))
            win.geometry(f"{w}x{h}+{x}+{y}")
    except tk.TclError:
        pass  # closed already


# ---------------------------------------------------------------- theme
def windows_uses_dark():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except Exception:
        return False


def resolve_mode(choice):
    """'system' / 'light' / 'dark' / 'classic' -> 'light', 'dark' or 'classic'"""
    choice = (choice or 'system').lower()
    if choice == 'classic':
        return 'classic'
    if choice == 'system':
        return 'dark' if windows_uses_dark() else 'light'
    return 'dark' if choice == 'dark' else 'light'


def colors():
    return PALETTE[_current['mode']]


def _title_bar(win, dark):
    """Dark/light Windows title bar to match (Windows 10 2004+ / 11)"""
    if sys.platform != 'win32':
        return
    try:
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id()) or win.winfo_id()
        value = ctypes.c_int(1 if dark else 0)
        for attribute in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (newer, then older builds)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(value),
                                                          ctypes.sizeof(value)) == 0:
                break
        # redraw the frame so the change shows on a window that's already open
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020)
    except Exception:
        pass


def _recolor(widget):
    """Classic Tk widgets: fields get the field shade, everything else the window colour"""
    c = colors()
    try:
        cls = widget.winfo_class()
        if cls in CLASSIC_FIELDS:
            classic = _current['mode'] == 'classic'
            widget.configure(background=c['field'], foreground=c['fg'], selectbackground=c['select'],
                             selectforeground=c['selectfg'], highlightthickness=0 if classic else 1,
                             highlightbackground=c['border'], highlightcolor=c['select'],
                             relief=tk.SUNKEN if classic else tk.FLAT, borderwidth=1 if classic else 0)
            if cls == 'Text':
                widget.configure(insertbackground=c['fg'])
        elif cls == 'Canvas':
            widget.configure(background=c['bg'], highlightthickness=0)
        elif cls == 'Panedwindow':
            widget.configure(background=c['sash'])
        elif cls in ('Toplevel', 'Tk', 'Frame'):
            widget.configure(background=c['bg'])
    except tk.TclError:
        pass
    for child in widget.winfo_children():
        _recolor(child)


_font_base = {}


def _scale_fonts(root):
    """Sun Valley's fonts are fixed pixel sizes (14 px body text), so on a 150% screen they'd stay tiny while
    everything else grows. Scale them like Windows scales normal text."""
    from tkinter import font
    s = scale_of(root)
    for name in font.names(root):
        if not name.startswith('SunValley'):
            continue
        f = font.nametofont(name, root=root)
        size = int(f.cget('size'))
        base = _font_base.setdefault(name, size)  # remember the original so repeated calls don't compound
        if base < 0:  # negative = pixels
            f.configure(size=-int(round(-base * s)))


def _compact_tabs(root, style, mode, s):
    """Sun Valley's tab headers have 16 px of padding on each side built into their image element - too wide
    for the TTS's nine tabs. Same images, slimmer sides."""
    if sv_ttk is None:
        return
    ns = 'sv_dark' if mode == 'dark' else 'sv_light'
    try:
        rest, selected, hover = (root.tk.eval(f'set ttk::theme::{ns}::I({name})')
                                 for name in ('tab-rest', 'tab-selected', 'tab-hover'))
        if 'Compact.Notebook.tab' not in style.element_names():  # elements belong to the current theme
            style.element_create('Compact.Notebook.tab', 'image', rest, ('selected', selected), ('active', hover),
                                 border=13, padding=(int(9 * s), 14, int(9 * s), 6), height=32)
        style.layout('TNotebook.Tab', [('Compact.Notebook.tab', {'sticky': 'nswe', 'children': [
            ('Notebook.padding', {'side': 'top', 'sticky': 'nswe', 'children': [
                ('Notebook.focus', {'side': 'top', 'sticky': 'nswe', 'children': [
                    ('Notebook.label', {'side': 'top', 'sticky': ''})]})]})]})])
    except tk.TclError as e:
        logger.debug(f"Compact tabs not available: {e}")


def apply_theme(root, choice='system'):
    """Switch every window to Light/Dark ('system' follows Windows) or the original Classic look. Safe to call
    again to change it."""
    mode = resolve_mode(choice)
    if mode != 'classic' and sv_ttk is None:
        mode = 'classic'  # the modern theme isn't available
    _current['mode'] = mode
    c = colors()
    if mode == 'classic':
        return _apply_classic(root, c)
    if sv_ttk is not None:
        try:
            sv_ttk.set_theme(mode, root)
        except Exception as e:
            logger.warning(f"Theme not available ({e}) - using the plain look")
    style = ttk.Style(root)
    _scale_fonts(root)
    body = 'SunValleyBodyFont' if sv_ttk is not None else 'TkDefaultFont'
    strong = 'SunValleyBodyStrongFont' if sv_ttk is not None else ('TkDefaultFont', 10, 'bold')
    style.configure('Section.TLabel', font=strong)
    style.configure('Title.TLabel', font='SunValleySubtitleFont' if sv_ttk is not None else ('TkDefaultFont', 12, 'bold'))
    style.configure('Hint.TLabel', foreground=c['hint'])
    s = scale_of(root)
    _compact_tabs(root, style, mode, s)
    # Small buttons for the list panel on the right
    style.configure('Compact.TButton', padding=(int(6 * s), int(1 * s), int(6 * s), int(2 * s)))
    # New classic widgets (dialogs, the DJ/Soundboard windows) pick these up when they're created
    for cls in CLASSIC_FIELDS:
        root.option_add(f'*{cls}.background', c['field'], 'interactive')
        root.option_add(f'*{cls}.foreground', c['fg'], 'interactive')
        root.option_add(f'*{cls}.selectBackground', c['select'], 'interactive')
        root.option_add(f'*{cls}.selectForeground', c['selectfg'], 'interactive')
        root.option_add(f'*{cls}.highlightThickness', 1, 'interactive')
        root.option_add(f'*{cls}.highlightBackground', c['border'], 'interactive')
        root.option_add(f'*{cls}.relief', 'flat', 'interactive')
        root.option_add(f'*{cls}.borderWidth', 0, 'interactive')
        root.option_add(f'*{cls}.font', body, 'interactive')
    root.option_add('*Text.insertBackground', c['fg'], 'interactive')
    root.option_add('*Canvas.background', c['bg'], 'interactive')
    root.option_add('*Canvas.highlightThickness', 0, 'interactive')
    # Dialogs and other windows get a matching title bar as they open
    root.bind_class('Toplevel', '<Map>', lambda e: _title_bar(e.widget, mode == 'dark') if e.widget.winfo_class() == 'Toplevel' else None, add='')
    _title_bar(root, mode == 'dark')
    _recolor(root)
    if sv_ttk is not None:
        _refont(root)
    root.after(60, lambda: _recolor(root))  # once more after the theme's own <<ThemeChanged>> recolouring
    for top in root.winfo_children():
        if isinstance(top, tk.Toplevel):
            _title_bar(top, mode == 'dark')
    return mode


def _apply_classic(root, c):
    """The original look: Tk's plain grey theme, Arial headings, white sunken lists - the new resizable layout"""
    style = ttk.Style(root)
    style.theme_use('default')
    # undo the modern theme's colours on the classic widgets (it set them with tk_setPalette)
    root.tk_setPalette(background=c['bg'], foreground=c['fg'], activeBackground='#ececec',
                       activeForeground=c['fg'], highlightColor='#000000', selectBackground=c['select'],
                       selectForeground=c['selectfg'])
    style.configure('Section.TLabel', font=('Arial', 10, 'bold'))
    style.configure('Title.TLabel', font=('Arial', 12, 'bold'))
    style.configure('Hint.TLabel', foreground=c['hint'])
    style.configure('Compact.TButton', padding=(2, 1))
    for cls, font in (('Listbox', 'TkDefaultFont'), ('Text', 'TkFixedFont')):
        root.option_add(f'*{cls}.background', c['field'], 'interactive')
        root.option_add(f'*{cls}.foreground', c['fg'], 'interactive')
        root.option_add(f'*{cls}.selectBackground', c['select'], 'interactive')
        root.option_add(f'*{cls}.selectForeground', c['selectfg'], 'interactive')
        root.option_add(f'*{cls}.highlightThickness', 0, 'interactive')
        root.option_add(f'*{cls}.relief', 'sunken', 'interactive')
        root.option_add(f'*{cls}.borderWidth', 1, 'interactive')
        root.option_add(f'*{cls}.font', font, 'interactive')
    root.option_add('*Text.insertBackground', c['fg'], 'interactive')
    root.option_add('*Canvas.background', c['bg'], 'interactive')
    root.bind_class('Toplevel', '<Map>', lambda e: _title_bar(e.widget, False) if e.widget.winfo_class() == 'Toplevel' else None, add='')
    _title_bar(root, False)
    _recolor(root)
    root.after(60, lambda: _recolor(root))
    _refont(root)
    for top in root.winfo_children():
        if isinstance(top, tk.Toplevel):
            _title_bar(top, False)
    return 'classic'


def _refont(widget):
    """Lists and text boxes already open take the font of the current look"""
    classic = _current['mode'] == 'classic'
    try:
        cls = widget.winfo_class()
        if cls == 'Listbox':
            widget.configure(font='TkDefaultFont' if classic else 'SunValleyBodyFont')
        elif cls == 'Text':
            widget.configure(font='TkFixedFont' if classic else 'SunValleyBodyFont')
    except tk.TclError:
        pass
    for child in widget.winfo_children():
        _refont(child)


# ---------------------------------------------------------------- scrolling
class ScrolledText(tk.Text):
    """Drop-in for tkinter.scrolledtext.ScrolledText with a themed scrollbar (the stock one is an old-style
    Windows scrollbar that sticks out, especially in dark mode)"""

    def __init__(self, master=None, **kw):
        self.frame = ttk.Frame(master)
        self.vbar = ttk.Scrollbar(self.frame, orient=tk.VERTICAL)
        self.vbar.pack(side=tk.RIGHT, fill=tk.Y)
        kw.update({'yscrollcommand': self.vbar.set})
        super().__init__(self.frame, **kw)
        self.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.vbar['command'] = self.yview
        # Geometry calls on this widget place the outer frame (same trick as the stock ScrolledText)
        text_meths = set(vars(tk.Text))
        methods = set(vars(tk.Pack)) | set(vars(tk.Grid)) | set(vars(tk.Place))
        for m in methods - text_meths:
            if m[0] != '_' and m not in ('config', 'configure'):
                setattr(self, m, getattr(self.frame, m))

    def __str__(self):
        return str(self.frame)


class ScrollFrame(ttk.Frame):
    """A frame whose contents scroll when there isn't room for them. When there is room, the contents fill it
    (so lists and text boxes inside still stretch). Put widgets in .inner."""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.vbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.hbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vbar.set, xscrollcommand=self.hbar.set)
        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window(0, 0, window=self.inner, anchor='nw')
        self.canvas.grid(row=0, column=0, sticky='nsew')
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.inner.bind('<Configure>', self._layout)
        self.canvas.bind('<Configure>', self._layout)
        self.bind('<Enter>', lambda e: self._wheel(True))
        self.bind('<Leave>', lambda e: self._wheel(False))
        self._pending = False

    def _layout(self, event=None):
        if self._pending:
            return
        self._pending = True
        self.after_idle(self._do_layout)

    def _do_layout(self):
        self._pending = False
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        rw, rh = self.inner.winfo_reqwidth(), self.inner.winfo_reqheight()
        # fill the visible area when the contents fit; keep their own size (and scroll) when they don't
        self.canvas.itemconfigure(self._win, width=max(cw, rw), height=max(ch, rh))
        self.canvas.configure(scrollregion=(0, 0, max(cw, rw), max(ch, rh)))
        self._show(self.vbar, rh > ch + 1, row=0, column=1, sticky='ns')
        self._show(self.hbar, rw > cw + 1, row=1, column=0, sticky='ew')

    @staticmethod
    def _show(bar, needed, **grid):
        if needed and not bar.winfo_ismapped():
            bar.grid(**grid)
        elif not needed and bar.winfo_ismapped():
            bar.grid_remove()

    def _wheel(self, on):
        if on:
            self.bind_all('<MouseWheel>', self._on_wheel)
            self.bind_all('<Shift-MouseWheel>', self._on_shift_wheel)
        else:
            self.unbind_all('<MouseWheel>')
            self.unbind_all('<Shift-MouseWheel>')

    def _scrollable_under_mouse(self, widget):
        """Lists, text boxes and drop-downs scroll themselves"""
        try:
            return widget.winfo_class() in ('Listbox', 'Text', 'Treeview', 'TCombobox', 'TSpinbox', 'Scale', 'TScale')
        except Exception:
            return False

    def _on_wheel(self, event):
        if self._scrollable_under_mouse(event.widget) or not self.vbar.winfo_ismapped():
            return
        self.canvas.yview_scroll(int(-event.delta / 120), 'units')

    def _on_shift_wheel(self, event):
        if self.hbar.winfo_ismapped():
            self.canvas.xview_scroll(int(-event.delta / 120), 'units')
