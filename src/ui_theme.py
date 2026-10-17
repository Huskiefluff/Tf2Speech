"""
Look and layout helpers for the TTS windows

- A flat modern look (Light, Dark or "match Windows") drawn by Tk itself - fast to redraw - or the Classic look,
  with matching Windows title bars
- Sharp text at 125%/150% Windows scaling, and window/dialog sizes that scale with it
- ScrollFrame: a tab or panel that scrolls when the window is too small for it, instead of cutting it off

(The Sun Valley theme was tried first, but it draws every control from semi-transparent pictures, which Tk on
Windows redraws very slowly: ~140 ms per tab hover and over a second per resize step on the TTS window.)
"""

import os
import sys
import ctypes
import logging
import tkinter as tk
from tkinter import ttk, font as tkfont

logger = logging.getLogger(__name__)

PALETTE = {
    'light': {'bg': '#f3f3f3', 'field': '#ffffff', 'fg': '#1b1b1b', 'border': '#cfcfcf', 'hint': '#666666',
              'select': '#0067c0', 'selectfg': '#ffffff', 'good': '#d9f2df', 'sash': '#e2e2e2',
              'btn': '#fdfdfd', 'btn_hover': '#ededed', 'btn_press': '#e0e0e0',
              'accent': '#0067c0', 'accent_hover': '#1975c5', 'accent_fg': '#ffffff',
              'tab': '#e6e6e6', 'tab_hover': '#ececec', 'disabled': '#a0a0a0'},
    'dark': {'bg': '#202020', 'field': '#2b2b2b', 'fg': '#f2f2f2', 'border': '#3d3d3d', 'hint': '#a3a3a3',
             'select': '#0a64c8', 'selectfg': '#ffffff', 'good': '#1f4d2b', 'sash': '#2a2a2a',
             'btn': '#2d2d2d', 'btn_hover': '#383838', 'btn_press': '#444444',
             'accent': '#4cc2ff', 'accent_hover': '#6dcfff', 'accent_fg': '#000000',
             'tab': '#272727', 'tab_hover': '#303030', 'disabled': '#6b6b6b'},
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


def fast_resize(root):
    """Stop Windows repainting every control each time the window changes size. Tk registers its windows as
    "redraw everything on resize" (CS_HREDRAW | CS_VREDRAW), so dragging the TTS window's edge made all ~110
    visible controls repaint on every step (~110 ms each, i.e. stutter). Controls whose size really changes still
    redraw themselves, so nothing is lost. Applies to Tk's two window classes for this program only."""
    if sys.platform != 'win32':
        return
    try:
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        user32.GetClassLongPtrW.restype = ctypes.c_size_t
        user32.GetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.SetClassLongPtrW.restype = ctypes.c_size_t
        user32.SetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_size_t]
        root.update_idletasks()
        for hwnd in (user32.GetParent(root.winfo_id()), root.winfo_id()):  # TkTopLevel frame, TkChild windows
            if hwnd:
                style = user32.GetClassLongPtrW(hwnd, -26)  # GCL_STYLE
                user32.SetClassLongPtrW(hwnd, -26, style & ~0x0003)
    except Exception as e:
        logger.debug(f"fast_resize not applied: {e}")


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
    classic = _current['mode'] == 'classic'
    try:
        cls = widget.winfo_class()
        if cls in CLASSIC_FIELDS:
            widget.configure(background=c['field'], foreground=c['fg'], selectbackground=c['select'],
                             selectforeground=c['selectfg'], highlightthickness=0 if classic else 1,
                             highlightbackground=c['border'], highlightcolor=c['select'],
                             relief=tk.SUNKEN if classic else tk.FLAT, borderwidth=1 if classic else 0,
                             font=('TkDefaultFont' if cls == 'Listbox' else 'TkFixedFont') if classic else 'TTSBody')
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


_FONTS = {}  # tkinter deletes a named font when its Font object is garbage-collected, so keep them


def _recolor_popdowns(root, c, font):
    """A combobox's drop-down list is made by Tk the first time it opens, so it isn't one of the program's own
    widgets and _recolor never sees it - after switching look it kept the old colours (e.g. grey with white text)"""
    stack = ['.']
    while stack:
        path = stack.pop()
        try:
            stack.extend(root.tk.splitlist(root.tk.call('winfo', 'children', path)))
            if root.tk.call('winfo', 'class', path) != 'TCombobox':
                continue
            listbox = f'{path}.popdown.f.l'
            if int(root.tk.call('winfo', 'exists', listbox)):
                root.tk.call(listbox, 'configure', '-background', c['field'], '-foreground', c['fg'],
                             '-selectbackground', c['select'], '-selectforeground', c['selectfg'], '-font', font)
        except tk.TclError:
            pass


def _fonts(root):
    """Named fonts for the flat look (in points, so they grow with Windows scaling)"""
    have = set(tkfont.names(root))
    for name, family, size in (('TTSBody', 'Segoe UI', 10), ('TTSStrong', 'Segoe UI Semibold', 10),
                               ('TTSTitle', 'Segoe UI Semibold', 13), ('TTSSmall', 'Segoe UI', 9)):
        if name not in have:
            _FONTS[name] = tkfont.Font(root=root, name=name, family=family, size=size)


def _option_db(root, c, body, text_font, classic):
    """New classic widgets (dialogs, the DJ/Soundboard windows, combobox drop-downs) pick these up"""
    for cls in CLASSIC_FIELDS:
        root.option_add(f'*{cls}.background', c['field'], 'interactive')
        root.option_add(f'*{cls}.foreground', c['fg'], 'interactive')
        root.option_add(f'*{cls}.selectBackground', c['select'], 'interactive')
        root.option_add(f'*{cls}.selectForeground', c['selectfg'], 'interactive')
        root.option_add(f'*{cls}.highlightThickness', 0 if classic else 1, 'interactive')
        root.option_add(f'*{cls}.highlightBackground', c['border'], 'interactive')
        root.option_add(f'*{cls}.relief', 'sunken' if classic else 'flat', 'interactive')
        root.option_add(f'*{cls}.borderWidth', 1 if classic else 0, 'interactive')
    root.option_add('*Listbox.font', body, 'interactive')
    root.option_add('*Text.font', text_font, 'interactive')
    root.option_add('*Text.insertBackground', c['fg'], 'interactive')
    root.option_add('*Canvas.background', c['bg'], 'interactive')
    root.option_add('*Canvas.highlightThickness', 0, 'interactive')
    root.option_add('*TCombobox*Listbox.background', c['field'], 'interactive')
    root.option_add('*TCombobox*Listbox.foreground', c['fg'], 'interactive')
    root.option_add('*TCombobox*Listbox.selectBackground', c['select'], 'interactive')
    root.option_add('*TCombobox*Listbox.selectForeground', c['selectfg'], 'interactive')


def _flat_theme(root, c):
    """The modern look: Tk's 'clam' theme, flattened and recoloured. Everything is drawn as plain shapes, so
    hovering and resizing redraw quickly."""
    style = ttk.Style(root)
    style.theme_use('clam')
    bg, field, fg, border = c['bg'], c['field'], c['fg'], c['border']
    style.configure('.', background=bg, foreground=fg, fieldbackground=field, bordercolor=border,
                    lightcolor=bg, darkcolor=bg, troughcolor=c['field'], selectbackground=c['select'],
                    selectforeground=c['selectfg'], insertcolor=fg, arrowcolor=fg, focuscolor=c['accent'],
                    font='TTSBody', relief='flat')
    style.map('.', foreground=[('disabled', c['disabled'])])
    flat = lambda colour: {'background': colour, 'lightcolor': colour, 'darkcolor': colour}
    # buttons
    style.configure('TButton', padding=(10, 4), bordercolor=border, focusthickness=0, **flat(c['btn']))
    style.map('TButton', background=[('disabled', bg), ('pressed', c['btn_press']), ('active', c['btn_hover'])],
              lightcolor=[('disabled', bg), ('pressed', c['btn_press']), ('active', c['btn_hover'])],
              darkcolor=[('disabled', bg), ('pressed', c['btn_press']), ('active', c['btn_hover'])],
              bordercolor=[('focus', c['accent'])])
    style.configure('Accent.TButton', foreground=c['accent_fg'], bordercolor=c['accent'], **flat(c['accent']))
    style.map('Accent.TButton',
              background=[('disabled', bg), ('pressed', c['accent_hover']), ('active', c['accent_hover'])],
              lightcolor=[('disabled', bg), ('pressed', c['accent_hover']), ('active', c['accent_hover'])],
              darkcolor=[('disabled', bg), ('pressed', c['accent_hover']), ('active', c['accent_hover'])],
              foreground=[('disabled', c['disabled'])])
    style.configure('Compact.TButton', padding=(6, 2), width=-4)  # no 11-character minimum width
    # text fields
    for name in ('TEntry', 'TCombobox', 'TSpinbox'):
        style.configure(name, padding=(6, 3), bordercolor=border, fieldbackground=field, foreground=fg,
                        lightcolor=field, darkcolor=field, background=c['btn'], arrowcolor=fg)
        style.map(name, bordercolor=[('focus', c['accent'])], lightcolor=[('focus', field)],
                  fieldbackground=[('readonly', field), ('disabled', bg)])
    style.map('TCombobox', selectbackground=[('readonly', field)], selectforeground=[('readonly', fg)],
              background=[('active', c['btn_hover'])])
    # check boxes and radio buttons
    for name in ('TCheckbutton', 'TRadiobutton'):
        style.configure(name, background=bg, foreground=fg, indicatorbackground=field, indicatorforeground=fg,
                        upperbordercolor=border, lowerbordercolor=border, focuscolor=bg)
        style.map(name, background=[('active', bg)],
                  indicatorbackground=[('selected', c['accent']), ('pressed', c['btn_press'])],
                  indicatorforeground=[('selected', c['accent_fg'])])
    # tabs
    style.configure('TNotebook', background=bg, bordercolor=border, lightcolor=bg, darkcolor=bg,
                    tabmargins=(0, 2, 0, 0))
    style.configure('TNotebook.Tab', padding=(9, 4), font='TTSSmall', bordercolor=border, foreground=c['hint'],
                    **flat(c['tab']))
    style.map('TNotebook.Tab', background=[('selected', bg), ('active', c['tab_hover'])],
              lightcolor=[('selected', bg), ('active', c['tab_hover'])],
              darkcolor=[('selected', bg), ('active', c['tab_hover'])],
              foreground=[('selected', fg), ('active', fg)])
    # lists and tables
    rowheight = tkfont.nametofont('TTSBody', root=root).metrics('linespace') + 6
    style.configure('Treeview', background=field, fieldbackground=field, foreground=fg, bordercolor=border,
                    lightcolor=field, darkcolor=field, rowheight=rowheight)
    style.map('Treeview', background=[('selected', c['select'])], foreground=[('selected', c['selectfg'])])
    style.configure('Treeview.Heading', relief='flat', foreground=fg, bordercolor=border, **flat(c['btn']))
    style.map('Treeview.Heading', background=[('active', c['btn_hover'])], lightcolor=[('active', c['btn_hover'])],
              darkcolor=[('active', c['btn_hover'])])
    # scrollbars, sliders, frames
    for name in ('Vertical.TScrollbar', 'Horizontal.TScrollbar'):
        style.configure(name, troughcolor=bg, bordercolor=bg, arrowcolor=c['hint'], gripcount=0, **flat(c['btn']))
        style.map(name, background=[('active', c['btn_hover'])], lightcolor=[('active', c['btn_hover'])],
                  darkcolor=[('active', c['btn_hover'])])
    for name in ('Horizontal.TScale', 'Vertical.TScale'):
        style.configure(name, troughcolor=field, bordercolor=border, **flat(c['btn']))
        style.map(name, background=[('active', c['btn_hover'])])
    style.configure('TLabelframe', background=bg, bordercolor=border, lightcolor=bg, darkcolor=bg)
    style.configure('TLabelframe.Label', background=bg, foreground=c['hint'], font='TTSSmall')
    style.configure('TSeparator', background=border)
    style.configure('TProgressbar', background=c['accent'], troughcolor=field, bordercolor=border,
                    lightcolor=c['accent'], darkcolor=c['accent'])
    style.configure('Sash', sashthickness=5, background=bg)
    style.configure('Section.TLabel', font='TTSStrong')
    style.configure('Title.TLabel', font='TTSTitle')
    style.configure('Hint.TLabel', foreground=c['hint'])
    root.tk_setPalette(background=bg, foreground=fg, activeBackground=c['btn_hover'], activeForeground=fg,
                       highlightColor=c['accent'], selectBackground=c['select'], selectForeground=c['selectfg'],
                       insertBackground=fg, troughColor=field)


def _classic_theme(root, c):
    """The original look: Tk's plain grey theme, Arial headings, white sunken lists"""
    style = ttk.Style(root)
    style.theme_use('default')
    root.tk_setPalette(background=c['bg'], foreground=c['fg'], activeBackground='#ececec',
                       activeForeground=c['fg'], highlightColor='#000000', selectBackground=c['select'],
                       selectForeground=c['selectfg'])
    style.configure('Section.TLabel', font=('Arial', 10, 'bold'))
    style.configure('Title.TLabel', font=('Arial', 12, 'bold'))
    style.configure('Hint.TLabel', foreground=c['hint'])
    style.configure('Compact.TButton', padding=(2, 1))


def apply_theme(root, choice='system'):
    """Switch every window to Light/Dark ('system' follows Windows) or the original Classic look. Safe to call
    again to change it. Returns 'light', 'dark' or 'classic'."""
    mode = resolve_mode(choice)
    _current['mode'] = mode
    c = colors()
    _fonts(root)
    classic = mode == 'classic'
    if classic:
        _classic_theme(root, c)
        _option_db(root, c, 'TkDefaultFont', 'TkFixedFont', True)
    else:
        _flat_theme(root, c)
        _option_db(root, c, 'TTSBody', 'TTSBody', False)
    dark = mode == 'dark'
    # Dialogs and other windows get a matching title bar as they open
    root.bind_class('Toplevel', '<Map>',
                    lambda e: _title_bar(e.widget, dark) if e.widget.winfo_class() == 'Toplevel' else None, add='')
    _title_bar(root, dark)
    _recolor(root)
    _recolor_popdowns(root, c, 'TkDefaultFont' if classic else 'TTSBody')
    for top in root.winfo_children():
        if isinstance(top, tk.Toplevel):
            _title_bar(top, dark)
    return mode


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
