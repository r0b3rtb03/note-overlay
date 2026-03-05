import os
import sys
import re
import threading
import json
import ctypes
import platform
import base64
import io
try:
    import pystray
    from PIL import Image, ImageDraw, ImageGrab
    TRAY_AVAILABLE = True
except Exception:
    TRAY_AVAILABLE = False
try:
    from dotenv import load_dotenv
    _env_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), '.env')
    if os.path.exists(_env_path):
        load_dotenv(_env_path)
    else:
        load_dotenv()
except ImportError:
    pass
try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from ctypes import wintypes


# ---------------------------------------------------------------------------
# Theme definitions
# ---------------------------------------------------------------------------
THEMES = {
    'dark': {
        'bg': '#1e1e1e',
        'fg': '#dcdcdc',
        'entry_bg': '#2b2b2b',
        'text_bg': '#121212',
        'highlight_bg': '#44475a',
        'button_bg': '#333333',
        'button_fg': '#dcdcdc',
        'selectcolor': '#1e1e1e',
        'accent': '#5865f2',
        'border': '#3a3a3a',
    },
    'light': {
        'bg': '#f5f5f5',
        'fg': '#1a1a1a',
        'entry_bg': '#ffffff',
        'text_bg': '#ffffff',
        'highlight_bg': '#fff176',
        'button_bg': '#e0e0e0',
        'button_fg': '#1a1a1a',
        'selectcolor': '#f5f5f5',
        'accent': '#4a6cf7',
        'border': '#cccccc',
    },
    'transparent': {
        'bg': '#010101',
        'fg': '#000000',
        'entry_bg': '#010101',
        'text_bg': '#010101',
        'highlight_bg': '#333333',
        'button_bg': '#010101',
        'button_fg': '#000000',
        'selectcolor': '#010101',
        'accent': '#010101',
        'border': '#010101',
    },
}


class ScreenSnipper:
    """Full-screen transparent overlay for region capture (ShareX-style)."""

    def __init__(self, parent, callback):
        self.callback = callback
        try:
            if parent is not None:
                self._top = tk.Toplevel(parent)
            else:
                self._top = tk.Toplevel()
        except Exception:
            self._top = tk.Toplevel()
        self._top.overrideredirect(True)
        try:
            self._top.attributes('-fullscreen', True)
        except Exception:
            self._top.geometry(
                f'{self._top.winfo_screenwidth()}x{self._top.winfo_screenheight()}+0+0'
            )
        try:
            self._top.attributes('-alpha', 0.3)
        except Exception:
            pass
        self._top.config(cursor='cross')
        self.canvas = tk.Canvas(self._top, bg='black', highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)

        self.start_x = self.start_y = None
        self.rect_id = None

        self._top.bind('<ButtonPress-1>', self._on_press)
        self._top.bind('<B1-Motion>', self._on_drag)
        self._top.bind('<ButtonRelease-1>', self._on_release)
        self._top.bind('<Escape>', lambda e: self._cancel())
        try:
            self._top.focus_force()
        except Exception:
            pass

    def _on_press(self, event):
        self.start_x = self._top.winfo_pointerx()
        self.start_y = self._top.winfo_pointery()
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline='red', width=2,
        )

    def _on_drag(self, event):
        if self.rect_id is None:
            return
        cx = self._top.winfo_pointerx()
        cy = self._top.winfo_pointery()
        self.canvas.coords(self.rect_id, self.start_x, self.start_y, cx, cy)

    def _on_release(self, event):
        if self.rect_id is None:
            self._destroy()
            return
        ex = self._top.winfo_pointerx()
        ey = self._top.winfo_pointery()
        left = min(self.start_x, ex)
        top = min(self.start_y, ey)
        right = max(self.start_x, ex)
        bottom = max(self.start_y, ey)
        self._destroy()
        try:
            img = ImageGrab.grab(bbox=(left, top, right, bottom))
        except Exception:
            img = None
        if callable(self.callback):
            self.callback(img)

    def _cancel(self):
        self._destroy()
        if callable(self.callback):
            self.callback(None)

    def _destroy(self):
        try:
            self._top.destroy()
        except Exception:
            pass


class NoteAssistantApp:
    def __init__(self, root, default_file='notes.txt'):
        self.root = root
        self.root.title('Service Host\uA789 Windows Helper')
        self.default_file = default_file
        self.visible = True

        # Config lives next to the exe (not in _MEIPASS temp dir)
        exe_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        self.config_path = os.path.join(exe_dir, 'note_assistant_config.json')

        # Proxy mode — single URL + key replaces direct API clients
        self.proxy_url = os.environ.get('PROXY_URL', '').rstrip('/')
        self.proxy_key = os.environ.get('PROXY_KEY', '')

        # Claude API client
        # Force IPv4 — IPv6 TLS to api.anthropic.com is broken on some networks
        self.anthropic_client = None
        api_key = os.environ.get('ANTHROPIC_API_KEY')
        if not self.proxy_url and ANTHROPIC_AVAILABLE and api_key:
            try:
                import httpx as _httpx
                transport = _httpx.HTTPTransport(local_address='0.0.0.0')
                http_client = _httpx.Client(transport=transport)
                self.anthropic_client = anthropic.Anthropic(
                    api_key=api_key, http_client=http_client,
                )
            except Exception as e:
                print(f'Failed to init Claude client: {e}')

        # Gemini API client (fallback)
        self.gemini_model = None
        self.gemini_vision_model = None
        gemini_key = os.environ.get('GEMINI_API_KEY')
        if not self.proxy_url and GEMINI_AVAILABLE and gemini_key and gemini_key != 'YOUR_GEMINI_KEY_HERE':
            try:
                genai.configure(api_key=gemini_key)
                self.gemini_model = genai.GenerativeModel('gemini-2.0-flash')
                self.gemini_vision_model = genai.GenerativeModel('gemini-2.0-flash')
            except Exception as e:
                print(f'Failed to init Gemini client: {e}')

        # Window size and appearance
        self.root.geometry('800x500')
        self.root.minsize(450, 300)

        # Apply initial theme
        self.current_theme = 'dark'
        self._apply_theme_colors('dark')
        self.root.configure(bg=self.bg)

        # ---------------------------------------------------------------
        # Use grid for main layout — only text area expands vertically
        # ---------------------------------------------------------------
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=0)  # toolbar row 1
        self.root.rowconfigure(1, weight=0)  # toolbar row 2
        self.root.rowconfigure(2, weight=3)  # main text area (gets most space)
        self.root.rowconfigure(3, weight=0)  # chat panel (fixed height)
        self.root.rowconfigure(4, weight=0)  # status bar

        # ---------------------------------------------------------------
        # Row 0 — Search + File controls
        # ---------------------------------------------------------------
        row0 = tk.Frame(self.root, bg=self.bg)
        row0.grid(row=0, column=0, sticky='ew', padx=8, pady=(6, 2))
        row0.columnconfigure(0, weight=1)
        self.top_frame = row0

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(row0, textvariable=self.search_var,
                                     bg=self.entry_bg, fg=self.fg,
                                     insertbackground=self.fg, relief='flat',
                                     font=('Segoe UI', 10))
        self.search_entry.grid(row=0, column=0, sticky='ew', padx=(0, 4))
        self.search_entry.bind('<Return>', lambda e: self.find_next())

        btn_frame = tk.Frame(row0, bg=self.bg)
        btn_frame.grid(row=0, column=1, sticky='e')

        self.find_btn = tk.Button(btn_frame, text='Find', command=self.find_all,
                                  bg=self.button_bg, fg=self.button_fg,
                                  relief='flat', padx=8, font=('Segoe UI', 9))
        self.find_btn.pack(side='left', padx=1)

        self.next_btn = tk.Button(btn_frame, text='Next', command=self.find_next,
                                  bg=self.button_bg, fg=self.button_fg,
                                  relief='flat', padx=8, font=('Segoe UI', 9))
        self.next_btn.pack(side='left', padx=1)

        sep1 = tk.Frame(btn_frame, width=2, bg=self.border, height=20)
        sep1.pack(side='left', padx=6, pady=2)

        self.open_btn = tk.Button(btn_frame, text='Open', command=self.open_file,
                                  bg=self.button_bg, fg=self.button_fg,
                                  relief='flat', padx=8, font=('Segoe UI', 9))
        self.open_btn.pack(side='left', padx=1)

        sep2 = tk.Frame(btn_frame, width=2, bg=self.border, height=20)
        sep2.pack(side='left', padx=6, pady=2)

        self.section_label = tk.Label(btn_frame, text='Section:', bg=self.bg, fg=self.fg,
                                      font=('Segoe UI', 9))
        self.section_label.pack(side='left')
        self.section_var = tk.StringVar(value='All')
        self.section_menu = tk.OptionMenu(btn_frame, self.section_var, 'All',
                                          command=self._on_section_change)
        self.section_menu.config(bg=self.entry_bg, fg=self.fg, highlightthickness=0,
                                 relief='flat', font=('Segoe UI', 9))
        self.section_menu['menu'].config(bg=self.entry_bg, fg=self.fg,
                                         font=('Segoe UI', 9))
        self.section_menu.pack(side='left', padx=(2, 0))

        # ---------------------------------------------------------------
        # Row 1 — Theme, Actions, Window controls
        # ---------------------------------------------------------------
        row1 = tk.Frame(self.root, bg=self.bg)
        row1.grid(row=1, column=0, sticky='ew', padx=8, pady=(2, 4))
        self.top2_frame = row1

        # Left group: theme + actions
        left_grp = tk.Frame(row1, bg=self.bg)
        left_grp.pack(side='left')

        self.theme_label = tk.Label(left_grp, text='Theme:', bg=self.bg, fg=self.fg,
                                    font=('Segoe UI', 9))
        self.theme_label.pack(side='left')
        self.theme_var = tk.StringVar(value='dark')
        self.theme_menu = tk.OptionMenu(left_grp, self.theme_var,
                                        'dark', 'light', 'transparent',
                                        command=self.apply_theme)
        self.theme_menu.config(bg=self.entry_bg, fg=self.fg, highlightthickness=0,
                               relief='flat', font=('Segoe UI', 9))
        self.theme_menu['menu'].config(bg=self.entry_bg, fg=self.fg,
                                       font=('Segoe UI', 9))
        self.theme_menu.pack(side='left', padx=(2, 8))

        self.format_btn = tk.Button(left_grp, text='Auto-Format',
                                    command=self.auto_format_notes,
                                    bg=self.button_bg, fg=self.button_fg,
                                    relief='flat', padx=8, font=('Segoe UI', 9))
        self.format_btn.pack(side='left', padx=1)

        self.prompt_visible = False
        self.toggle_prompt_btn = tk.Button(left_grp, text='Claude Chat',
                                           command=self.toggle_prompt_panel,
                                           bg=self.button_bg, fg=self.button_fg,
                                           relief='flat', padx=8, font=('Segoe UI', 9))
        self.toggle_prompt_btn.pack(side='left', padx=(4, 0))

        # Stealth snip settings (compact dropdown)
        self.snip_text_var = tk.StringVar(value='black')
        self.snip_font_var = tk.StringVar(value='Arial')
        self.snip_size_var = tk.StringVar(value='12')
        self.snip_hide_text_var = tk.BooleanVar(value=False)

        self.snip_settings_btn = tk.Menubutton(
            left_grp,
            text='Stealth ▾',
            bg=self.button_bg,
            fg=self.button_fg,
            relief='flat',
            padx=8,
            font=('Segoe UI', 9),
            activebackground=self.button_bg,
            activeforeground=self.button_fg,
        )
        self.snip_settings_btn.pack(side='left', padx=(8, 0))

        self.snip_settings_menu = tk.Menu(self.snip_settings_btn, tearoff=0)

        self.snip_color_menu = tk.Menu(self.snip_settings_menu, tearoff=0)
        self.snip_color_menu.add_radiobutton(label='Black', variable=self.snip_text_var, value='black')
        self.snip_color_menu.add_radiobutton(label='White', variable=self.snip_text_var, value='white')
        self.snip_settings_menu.add_cascade(label='Text Color', menu=self.snip_color_menu)

        self.snip_font_menu = tk.Menu(self.snip_settings_menu, tearoff=0)
        self.snip_font_menu.add_radiobutton(label='Arial', variable=self.snip_font_var, value='Arial')
        self.snip_font_menu.add_radiobutton(label='Consolas', variable=self.snip_font_var, value='Consolas')
        self.snip_font_menu.add_radiobutton(label='Segoe UI', variable=self.snip_font_var, value='Segoe UI')
        self.snip_settings_menu.add_cascade(label='Font Family', menu=self.snip_font_menu)

        self.snip_size_menu = tk.Menu(self.snip_settings_menu, tearoff=0)
        for _size in ('10', '11', '12', '13', '14', '16', '18'):
            self.snip_size_menu.add_radiobutton(label=_size, variable=self.snip_size_var, value=_size)
        self.snip_settings_menu.add_cascade(label='Font Size', menu=self.snip_size_menu)

        self.snip_settings_menu.add_separator()
        self.snip_settings_menu.add_checkbutton(
            label='Copy Only (Hide Text)',
            variable=self.snip_hide_text_var,
            onvalue=True,
            offvalue=False,
        )
        self.snip_settings_btn.configure(menu=self.snip_settings_menu)

        # Right group: window controls
        right_grp = tk.Frame(row1, bg=self.bg)
        right_grp.pack(side='right')

        self.topmost_var = tk.BooleanVar(value=False)
        self.topmost_cb = tk.Checkbutton(right_grp, text='On Top',
                                         variable=self.topmost_var,
                                         command=self.set_topmost,
                                         bg=self.bg, fg=self.fg,
                                         selectcolor=self.selectcolor,
                                         activebackground=self.bg,
                                         font=('Segoe UI', 8))
        self.topmost_cb.pack(side='left', padx=2)

        self.minimize_tray_var = tk.BooleanVar(value=False)
        self.tray_cb = tk.Checkbutton(right_grp, text='Tray',
                                      variable=self.minimize_tray_var,
                                      bg=self.bg, fg=self.fg,
                                      selectcolor=self.selectcolor,
                                      activebackground=self.bg,
                                      font=('Segoe UI', 8))
        self.tray_cb.pack(side='left', padx=2)

        # Hide from taskbar
        self.hide_taskbar_var = tk.BooleanVar(value=False)
        self.hide_cb = tk.Checkbutton(right_grp, text='Hide Taskbar',
                                      variable=self.hide_taskbar_var,
                                      command=self.apply_hide_taskbar,
                                      bg=self.bg, fg=self.fg,
                                      selectcolor=self.selectcolor,
                                      activebackground=self.bg,
                                      font=('Segoe UI', 8))
        self.hide_cb.pack(side='left', padx=2)

        # No-focus (browser won't detect switch)
        self.nofocus_var = tk.BooleanVar(value=False)
        self.nofocus_cb = tk.Checkbutton(right_grp, text='No-Focus',
                                         variable=self.nofocus_var,
                                         command=self.apply_nofocus_mode,
                                         bg=self.bg, fg=self.fg,
                                         selectcolor=self.selectcolor,
                                         activebackground=self.bg,
                                         font=('Segoe UI', 8))
        self.nofocus_cb.pack(side='left', padx=2)

        # ---------------------------------------------------------------
        # Row 2 — Main text area (expands to fill)
        # ---------------------------------------------------------------
        text_frame = tk.Frame(self.root, bg=self.bg)
        text_frame.grid(row=2, column=0, sticky='nsew', padx=8, pady=(0, 2))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)

        self.text = scrolledtext.ScrolledText(text_frame, wrap='word',
                                              bg=self.text_bg, fg=self.fg,
                                              insertbackground=self.fg,
                                              relief='flat', borderwidth=0,
                                              font=('Consolas', 10))
        self.text.grid(row=0, column=0, sticky='nsew')
        self.text.tag_config('highlight', background=self.highlight_bg)
        self.text.tag_config('section', background=self.highlight_bg)

        # ---------------------------------------------------------------
        # Row 3 — Chat panel (hidden by default, fixed height)
        # ---------------------------------------------------------------
        self.prompt_frame = tk.Frame(self.root, bg=self.bg, height=200)
        self.prompt_frame.grid_propagate(False)  # keep fixed height

        # Chat input row
        chat_input = tk.Frame(self.prompt_frame, bg=self.bg)
        chat_input.pack(fill='x', pady=(4, 2))

        self.prompt_label = tk.Label(chat_input, text='Ask Claude:',
                                     bg=self.bg, fg=self.fg,
                                     font=('Segoe UI', 9, 'bold'))
        self.prompt_label.pack(side='left', padx=(0, 6))

        self.snip_btn = tk.Button(chat_input, text='Snip & Ask',
                                   command=self._snip_and_ask,
                                   bg=self.button_bg, fg=self.button_fg,
                                   relief='flat', padx=8, font=('Segoe UI', 9))
        self.snip_btn.pack(side='right', padx=(4, 0))

        self.send_btn = tk.Button(chat_input, text='Send', command=self.send_prompt,
                                  bg=self.accent, fg='#ffffff',
                                  relief='flat', padx=12, font=('Segoe UI', 9, 'bold'))
        self.send_btn.pack(side='right')

        self.prompt_var = tk.StringVar()
        self.prompt_entry = tk.Entry(chat_input, textvariable=self.prompt_var,
                                     bg=self.entry_bg, fg=self.fg,
                                     insertbackground=self.fg, relief='flat',
                                     font=('Segoe UI', 10))
        self.prompt_entry.pack(side='left', fill='x', expand=True, padx=(0, 4))
        self.prompt_entry.bind('<Return>', lambda e: self.send_prompt())

        # Response area — fills the rest of the prompt_frame
        self.response_text = scrolledtext.ScrolledText(self.prompt_frame, wrap='word',
                                                       bg=self.text_bg, fg=self.fg,
                                                       insertbackground=self.fg,
                                                       relief='flat', borderwidth=0,
                                                       font=('Consolas', 10))
        self.response_text.pack(fill='both', expand=True, pady=(2, 0))
        self.response_text.config(state='disabled')

        # ---------------------------------------------------------------
        # Row 4 — Status bar
        # ---------------------------------------------------------------
        self.status = tk.Label(self.root, text='', anchor='w', bg=self.bg, fg=self.fg,
                               font=('Segoe UI', 8))
        self.status.grid(row=4, column=0, sticky='ew', padx=8, pady=(0, 4))

        # Show API status at startup
        if self.proxy_url:
            self.status.config(text=f'AI ready: Proxy mode')
        else:
            apis = []
            if self.anthropic_client:
                apis.append('Claude')
            if self.gemini_model:
                apis.append('Gemini')
            if apis:
                self.status.config(text=f'AI ready: {" + ".join(apis)}' + (' (Gemini fallback)' if len(apis) == 2 else ''))
            else:
                self.status.config(text='No AI API configured — set PROXY_URL or API keys in .env')

        # Widget collections for theme updates
        self._all_buttons = [
            self.find_btn, self.next_btn, self.open_btn,
            self.format_btn, self.toggle_prompt_btn, self.snip_btn,
            self.snip_settings_btn,
        ]
        self._all_checkbuttons = [
            self.topmost_cb, self.tray_cb, self.hide_cb, self.nofocus_cb,
        ]
        self._all_labels = [
            self.section_label, self.theme_label,
            self.prompt_label, self.status,
        ]
        self._all_frames = [
            self.top_frame, self.top2_frame, self.prompt_frame,
        ]
        self._separators = [sep1, sep2]
        self._btn_frame = btn_frame
        self._left_grp = left_grp
        self._right_grp = right_grp
        self._chat_input = chat_input
        self._text_frame = text_frame

        # Internal state
        self.current_search = ''
        self.last_found_index = None
        self._full_text = ''
        self._sections = {}

        # Load config — if no config exists, enable hide taskbar by default
        self._config_loaded = False
        self.load_config()
        if not self._config_loaded:
            self.hide_taskbar_var.set(True)
            self.root.after(200, self.apply_hide_taskbar)

        if not getattr(self, 'current_file', None):
            if os.path.exists(self.default_file):
                self.load_file(self.default_file)
            else:
                self.status.config(text=f'No {self.default_file} found \u2014 use Open to load notes')

        # Global hotkeys (Windows only)
        self._hotkey_thread = None
        if platform.system() == 'Windows':
            self._start_hotkey_listener()

        self.root.protocol('WM_DELETE_WINDOW', self.on_close)

    # -------------------------------------------------------------------
    # Theme system
    # -------------------------------------------------------------------
    def _apply_theme_colors(self, theme_name):
        theme = THEMES.get(theme_name, THEMES['dark'])
        self.bg = theme['bg']
        self.fg = theme['fg']
        self.entry_bg = theme['entry_bg']
        self.text_bg = theme.get('text_bg', theme['bg'])
        self.highlight_bg = theme['highlight_bg']
        self.button_bg = theme.get('button_bg', theme['entry_bg'])
        self.button_fg = theme.get('button_fg', theme['fg'])
        self.selectcolor = theme.get('selectcolor', theme['bg'])
        self.accent = theme.get('accent', '#5865f2')
        self.border = theme.get('border', '#3a3a3a')

    def apply_theme(self, theme_name):
        self.current_theme = theme_name
        self._apply_theme_colors(theme_name)
        self.root.configure(bg=self.bg)

        if theme_name == 'transparent':
            self.root.attributes('-transparentcolor', self.bg)
            self.root.attributes('-alpha', 1.0)
        else:
            try:
                self.root.attributes('-transparentcolor', '')
            except Exception:
                pass
            self.root.attributes('-alpha', 1.0)

        for frame in self._all_frames:
            frame.configure(bg=self.bg)
        for f in [self._btn_frame, self._left_grp, self._right_grp,
                  self._chat_input, self._text_frame]:
            f.configure(bg=self.bg)

        for sep in self._separators:
            sep.configure(bg=self.border)

        for btn in self._all_buttons:
            btn.configure(bg=self.button_bg, fg=self.button_fg)
        self.send_btn.configure(bg=self.accent, fg='#ffffff')

        for cb in self._all_checkbuttons:
            cb.configure(bg=self.bg, fg=self.fg, selectcolor=self.selectcolor,
                         activebackground=self.bg)

        for lbl in self._all_labels:
            lbl.configure(bg=self.bg, fg=self.fg)

        self.search_entry.configure(bg=self.entry_bg, fg=self.fg, insertbackground=self.fg)
        self.prompt_entry.configure(bg=self.entry_bg, fg=self.fg, insertbackground=self.fg)

        self.text.configure(bg=self.text_bg, fg=self.fg, insertbackground=self.fg)
        self.text.tag_config('highlight', background=self.highlight_bg)
        self.text.tag_config('section', background=self.highlight_bg)
        self.response_text.configure(bg=self.text_bg, fg=self.fg, insertbackground=self.fg)

        self.section_menu.config(bg=self.entry_bg, fg=self.fg)
        self.section_menu['menu'].config(bg=self.entry_bg, fg=self.fg)
        self.theme_menu.config(bg=self.entry_bg, fg=self.fg)
        self.theme_menu['menu'].config(bg=self.entry_bg, fg=self.fg)
        self.snip_settings_btn.config(bg=self.button_bg, fg=self.button_fg,
                                      activebackground=self.button_bg,
                                      activeforeground=self.button_fg)
        for _menu in (self.snip_settings_menu, self.snip_color_menu, self.snip_font_menu, self.snip_size_menu):
            _menu.config(bg=self.entry_bg, fg=self.fg)

    # -------------------------------------------------------------------
    # AI API calls (proxy mode + direct mode)
    # -------------------------------------------------------------------
    def _call_proxy(self, system_prompt, user_message, callback, image_b64=None):
        """Call the Cloudflare Worker proxy for text or vision queries."""
        import json as _json
        import socket
        import ssl
        from urllib.parse import urlparse

        def worker():
            payload = {
                'proxy_key': self.proxy_key,
                'system': system_prompt,
                'message': user_message,
            }
            if image_b64:
                payload['image_b64'] = image_b64
            try:
                parsed = urlparse(self.proxy_url)
                host = parsed.hostname
                port = parsed.port or 443
                path = parsed.path or '/'

                body = _json.dumps(payload).encode('utf-8')

                # Force IPv4 — IPv6 TLS to Cloudflare is broken on some networks
                addr = socket.getaddrinfo(host, port, socket.AF_INET)[0]
                sock = socket.create_connection(addr[4], timeout=30)
                ctx = ssl.create_default_context()
                ssock = ctx.wrap_socket(sock, server_hostname=host)

                req_str = (
                    f'POST {path} HTTP/1.1\r\n'
                    f'Host: {host}\r\n'
                    f'Content-Type: application/json\r\n'
                    f'Content-Length: {len(body)}\r\n'
                    f'Connection: close\r\n'
                    f'\r\n'
                )
                ssock.sendall(req_str.encode() + body)

                data = b''
                while True:
                    chunk = ssock.recv(8192)
                    if not chunk:
                        break
                    data += chunk
                ssock.close()

                raw = data.decode('utf-8', errors='replace')
                # Split HTTP headers from body
                if '\r\n\r\n' in raw:
                    resp_body = raw.split('\r\n\r\n', 1)[1]
                else:
                    resp_body = raw

                result = _json.loads(resp_body)
                text = result.get('text', '')
                error = result.get('error', '')
                if text:
                    self.root.after(0, lambda t=text: callback(t))
                else:
                    err = f'ERROR: Proxy: {error or "empty response"}'
                    self.root.after(0, lambda e=err: callback(e))
            except Exception as e:
                err = f'ERROR: Proxy request failed: {e}'
                self.root.after(0, lambda e=err: callback(e))

        threading.Thread(target=worker, daemon=True).start()

    def _call_gemini(self, system_prompt, user_message, callback):
        """Fallback: call Gemini API for text queries."""
        if not self.gemini_model:
            callback('ERROR: No AI API available.\n\nSet PROXY_URL or API keys in .env')
            return

        def worker():
            try:
                combined = f'{system_prompt}\n\n{user_message}'
                response = self.gemini_model.generate_content(combined)
                result_text = response.text
                self.root.after(0, lambda t=result_text: callback(t))
            except Exception as e:
                err = f'ERROR (Gemini): {e}'
                self.root.after(0, lambda e=err: callback(e))

        threading.Thread(target=worker, daemon=True).start()

    def _call_claude(self, system_prompt, user_message, callback,
                     model='claude-sonnet-4-6'):
        # Proxy mode — route everything through the worker
        if self.proxy_url:
            self._call_proxy(system_prompt, user_message, callback)
            return

        if not self.anthropic_client and not self.gemini_model:
            callback('ERROR: No AI API available.\n\nSet PROXY_URL or API keys in .env')
            return
        if not self.anthropic_client:
            self._call_gemini(system_prompt, user_message, callback)
            return

        def worker():
            try:
                response = self.anthropic_client.messages.create(
                    model=model,
                    max_tokens=4096,
                    system=system_prompt,
                    messages=[{'role': 'user', 'content': user_message}],
                )
                result_text = response.content[0].text
                self.root.after(0, lambda t=result_text: callback(t))
            except (anthropic.APIConnectionError, anthropic.APIStatusError) as e:
                # Claude failed — try Gemini fallback
                if self.gemini_model:
                    self.root.after(0, lambda: self._call_gemini(system_prompt, user_message, callback))
                else:
                    err = f'ERROR: Claude failed: {e}'
                    self.root.after(0, lambda e=err: callback(e))
            except Exception as e:
                if self.gemini_model:
                    self.root.after(0, lambda: self._call_gemini(system_prompt, user_message, callback))
                else:
                    err = f'ERROR: {e}'
                    self.root.after(0, lambda e=err: callback(e))

        threading.Thread(target=worker, daemon=True).start()

    # -------------------------------------------------------------------
    # Auto-format notes
    # -------------------------------------------------------------------
    def auto_format_notes(self):
        try:
            selected = self.text.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            selected = None

        content = selected if selected else self.text.get('1.0', tk.END).strip()
        if not content:
            self.status.config(text='No text to format')
            return

        self.status.config(text='Formatting notes with Claude...')
        self.format_btn.config(state='disabled')

        system_prompt = (
            'You are a note formatting assistant. Clean up and format the provided study notes. '
            'Preserve all content and meaning. Fix spelling, grammar, and formatting. '
            'Use consistent markdown headings (# for top level, ## for sub). '
            'Ensure multiple choice questions are properly numbered and indented. '
            'Return ONLY the formatted text with no preamble or explanation.'
        )

        def on_result(result):
            self.format_btn.config(state='normal')
            if result.startswith('ERROR:'):
                self.status.config(text=result[:80])
                return
            if selected:
                try:
                    self.text.delete(tk.SEL_FIRST, tk.SEL_LAST)
                    self.text.insert(tk.INSERT, result)
                except tk.TclError:
                    pass
            else:
                self.text.delete('1.0', tk.END)
                self.text.insert('1.0', result)
                self._parse_sections(result)
            self._full_text = self.text.get('1.0', tk.END).strip()
            self.status.config(text='Notes formatted successfully')

        self._call_claude(system_prompt, content, on_result)

    # -------------------------------------------------------------------
    # Claude chat panel
    # -------------------------------------------------------------------
    def toggle_prompt_panel(self):
        if self.prompt_visible:
            self.prompt_frame.grid_forget()
            self.prompt_visible = False
            self.toggle_prompt_btn.config(text='Claude Chat')
        else:
            self.prompt_frame.grid(row=3, column=0, sticky='ew', padx=8, pady=(0, 2))
            self.prompt_visible = True
            self.toggle_prompt_btn.config(text='Hide Chat')
            self.prompt_entry.focus_set()

    def send_prompt(self):
        question = self.prompt_var.get().strip()
        if not question:
            self.status.config(text='Enter a question first')
            return

        # Show thinking state
        self.response_text.config(state='normal')
        self.response_text.delete('1.0', tk.END)
        self.response_text.insert('1.0', 'Thinking...')
        self.response_text.config(state='disabled')

        notes_content = self.text.get('1.0', tk.END).strip()
        context_prefix = ''
        if notes_content:
            truncated = notes_content[:8000]
            context_prefix = (
                f'Here are the user\'s study notes for context:\n\n{truncated}\n\n---\n\n'
            )
        full_message = context_prefix + f'User question: {question}'

        system_prompt = (
            'You are a study assistant. '
            'For multiple choice or true/false questions, give ONLY the answer (e.g. "b. Resource pooling"). '
            'For short answer questions, write a single sentence at a high school to freshman college level. '
            'For essay questions, write the shortest paragraph possible at a high school to freshman college level. '
            'Do NOT explain your reasoning. Just provide the answer.'
        )

        self.status.config(text='Asking Claude...')
        self.send_btn.config(state='disabled')

        def on_result(result):
            self.send_btn.config(state='normal')
            self.response_text.config(state='normal')
            self.response_text.delete('1.0', tk.END)
            self.response_text.insert('1.0', result)
            self.response_text.see('1.0')
            self.response_text.config(state='disabled')
            if result.startswith('ERROR:'):
                self.status.config(text='Claude query failed \u2014 see response')
            else:
                self.status.config(text='Response received')
            self.prompt_var.set('')

        self._call_claude(system_prompt, full_message, on_result)

    # -------------------------------------------------------------------
    # Screen snip + Claude Vision
    # -------------------------------------------------------------------
    def _snip_and_ask(self):
        """Button-triggered snip — result shows in chat panel."""
        try:
            self.root.withdraw()
            self.root.update()
        except Exception:
            pass
        ScreenSnipper(self.root, lambda img: self.root.after(0, lambda i=img: self._on_snip_captured(i)))

    def _stealth_text(self):
        """Text-hotkey triggered — grab highlighted text, send to Claude, show answer like stealth snip."""
        if not self.proxy_url and not self.anthropic_client and not self.gemini_model:
            self.status.config(text='No AI API available')
            return

        import time
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        original_hwnd = user32.GetForegroundWindow()

        # ---- All blocking work runs in a background thread ----
        def _background():
            # Set up proper ctypes prototypes (64-bit safe)
            _OpenClipboard = user32.OpenClipboard
            _OpenClipboard.argtypes = [wintypes.HWND]
            _OpenClipboard.restype = wintypes.BOOL

            _CloseClipboard = user32.CloseClipboard
            _CloseClipboard.argtypes = []
            _CloseClipboard.restype = wintypes.BOOL

            _GetClipboardData = user32.GetClipboardData
            _GetClipboardData.argtypes = [wintypes.UINT]
            _GetClipboardData.restype = ctypes.c_void_p

            _GlobalLock = kernel32.GlobalLock
            _GlobalLock.argtypes = [ctypes.c_void_p]
            _GlobalLock.restype = ctypes.c_void_p

            _GlobalUnlock = kernel32.GlobalUnlock
            _GlobalUnlock.argtypes = [ctypes.c_void_p]
            _GlobalUnlock.restype = wintypes.BOOL

            def _read_clipboard_win32():
                """Read clipboard via Win32 API — works from any thread, no Tkinter needed."""
                CF_UNICODETEXT = 13
                # Retry OpenClipboard — may fail if source app still holds it
                for _ in range(6):
                    if _OpenClipboard(None):
                        break
                    time.sleep(0.03)
                else:
                    return ''
                try:
                    h = _GetClipboardData(CF_UNICODETEXT)
                    if not h:
                        return ''
                    p = _GlobalLock(h)
                    if not p:
                        return ''
                    try:
                        return ctypes.wstring_at(p)
                    finally:
                        _GlobalUnlock(h)
                finally:
                    _CloseClipboard()

            def _send_copy_combo(vk_key):
                KEYEVENTF_KEYUP = 0x0002
                INPUT_KEYBOARD = 1
                VK_CONTROL = 0x11

                class KEYBDINPUT(ctypes.Structure):
                    _fields_ = [
                        ('wVk', wintypes.WORD),
                        ('wScan', wintypes.WORD),
                        ('dwFlags', wintypes.DWORD),
                        ('time', wintypes.DWORD),
                        ('dwExtraInfo', ctypes.POINTER(ctypes.c_ulong)),
                    ]

                class INPUT(ctypes.Structure):
                    _fields_ = [
                        ('type', wintypes.DWORD),
                        ('ki', KEYBDINPUT),
                    ]

                extra = ctypes.c_ulong(0)
                inputs = (INPUT * 4)(
                    INPUT(INPUT_KEYBOARD, KEYBDINPUT(VK_CONTROL, 0, 0, 0, ctypes.pointer(extra))),
                    INPUT(INPUT_KEYBOARD, KEYBDINPUT(vk_key, 0, 0, 0, ctypes.pointer(extra))),
                    INPUT(INPUT_KEYBOARD, KEYBDINPUT(vk_key, 0, KEYEVENTF_KEYUP, 0, ctypes.pointer(extra))),
                    INPUT(INPUT_KEYBOARD, KEYBDINPUT(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0, ctypes.pointer(extra))),
                )
                sent = user32.SendInput(len(inputs), ctypes.byref(inputs), ctypes.sizeof(INPUT))
                if sent != len(inputs):
                    user32.keybd_event(VK_CONTROL, 0, 0, 0)
                    user32.keybd_event(vk_key, 0, 0, 0)
                    user32.keybd_event(vk_key, 0, KEYEVENTF_KEYUP, 0)
                    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)

            def _capture_after_copy(vk_key, old_clip, timeout_s=1.0):
                start_seq = user32.GetClipboardSequenceNumber()
                _send_copy_combo(vk_key)
                deadline = time.time() + timeout_s
                while time.time() < deadline:
                    if user32.GetClipboardSequenceNumber() != start_seq:
                        time.sleep(0.02)
                        text = _read_clipboard_win32()
                        if text and text != old_clip:
                            return text
                    time.sleep(0.04)
                return ''

            # Small delay for hotkey key-up settle
            time.sleep(0.06)

            old_clip = _read_clipboard_win32()

            VK_C = 0x43
            VK_INSERT = 0x2D

            selected = _capture_after_copy(VK_C, old_clip, timeout_s=1.0)

            if not selected:
                time.sleep(0.08)
                selected = _capture_after_copy(VK_INSERT, old_clip, timeout_s=1.0)

            if not selected:
                time.sleep(0.10)
                selected = _capture_after_copy(VK_C, old_clip, timeout_s=1.5)

            # Deliver result back to main thread
            self.root.after(0, lambda: _on_captured(selected))

        def _on_captured(selected_text):
            if not selected_text or not selected_text.strip():
                self.status.config(text='No text selected — highlight text first')
                return

            system_prompt = (
                'You are a study assistant that answers questions. '
                'For multiple choice or true/false questions, give ONLY the answer (e.g. "b. Resource pooling"). '
                'For short answer questions, write a single sentence at a high school to freshman college level. '
                'For essay questions, write the shortest paragraph possible at a high school to freshman college level. '
                'Do NOT explain your reasoning. Just provide the answer.'
            )

            def on_result(result):
                self._copy_to_clipboard(result)
                if self.snip_hide_text_var.get():
                    self.status.config(text='Stealth answer copied to clipboard')
                    return
                self._show_tooltip(result, auto_ms=3000, keep_focus_hwnd=original_hwnd)

            self._call_claude(system_prompt, selected_text.strip(), on_result)

        threading.Thread(target=_background, daemon=True).start()

    def _snip_stealth(self):
        """F10-triggered snip — app stays hidden, result shows as tooltip at cursor."""
        # Hide the window for the capture, keep it hidden afterwards
        try:
            self.root.withdraw()
            self.root.update()
        except Exception:
            pass

        def on_capture(img):
            def handle(i):
                # Do NOT restore the window — stealth mode stays hidden
                if i is None:
                    return
                self._call_claude_vision(i, tooltip=True)
            self.root.after(0, lambda: handle(img))

        # Use no explicit parent so creating the snipper cannot remap/deiconify root.
        ScreenSnipper(None, on_capture)

    def _on_snip_captured(self, img):
        try:
            self.root.deiconify()
            if self.nofocus_var.get():
                self.apply_nofocus_mode()
            else:
                self.root.lift()
                self.root.focus_force()
        except Exception:
            pass
        if img is None:
            self.status.config(text='Snip cancelled')
            return
        # Ensure chat panel is visible
        if not self.prompt_visible:
            self.toggle_prompt_panel()
        self.status.config(text='Analyzing screenshot with Claude Vision...')
        self._call_claude_vision(img, tooltip=False)

    def _call_claude_vision(self, pil_image, tooltip=False):
        if not self.proxy_url and not self.anthropic_client and not self.gemini_vision_model:
            self.status.config(text='No AI API available for vision')
            return

        buf = io.BytesIO()
        try:
            pil_image.save(buf, format='PNG')
        except Exception:
            pil_image.convert('RGBA').save(buf, format='PNG')
        img_bytes = buf.getvalue()
        b64 = base64.b64encode(img_bytes).decode('ascii')

        if not tooltip:
            self.response_text.config(state='normal')
            self.response_text.delete('1.0', tk.END)
            self.response_text.insert('1.0', 'Analyzing image...')
            self.response_text.config(state='disabled')
            self.snip_btn.config(state='disabled')

        system_prompt = (
            'You are a study assistant that reads images of questions. '
            'For multiple choice or true/false questions, give ONLY the answer (e.g. "b. Resource pooling"). '
            'For short answer questions, write a single sentence at a high school to freshman college level. '
            'For essay questions, write the shortest paragraph possible at a high school to freshman college level. '
            'Do NOT explain your reasoning. Just provide the answer.'
        )

        def on_result(result):
            if tooltip:
                self._copy_to_clipboard(result)
                if self.snip_hide_text_var.get():
                    self.status.config(text='Stealth answer copied to clipboard')
                    return
                self._show_tooltip(result, auto_ms=3000)
            else:
                self.snip_btn.config(state='normal')
                self.response_text.config(state='normal')
                self.response_text.delete('1.0', tk.END)
                self.response_text.insert('1.0', result)
                self.response_text.see('1.0')
                self.response_text.config(state='disabled')
                if result.startswith('ERROR:'):
                    self.status.config(text='Vision analysis failed \u2014 see response')
                else:
                    self.status.config(text='Vision analysis complete')

        def gemini_vision_fallback():
            try:
                img_part = {
                    'mime_type': 'image/png',
                    'data': img_bytes,
                }
                prompt = system_prompt + '\n\nAnswer the question in this image.'
                response = self.gemini_vision_model.generate_content([prompt, img_part])
                text = response.text
                self.root.after(0, lambda t=text: on_result(t))
            except Exception as e:
                err = f'ERROR (Gemini Vision): {e}'
                self.root.after(0, lambda e=err: on_result(e))

        def worker():
            # Proxy mode handles vision too
            if self.proxy_url:
                self._call_proxy(system_prompt, 'Answer the question in this image.', on_result, image_b64=b64)
                return
            if not self.anthropic_client:
                gemini_vision_fallback()
                return
            try:
                response = self.anthropic_client.messages.create(
                    model='claude-sonnet-4-6',
                    max_tokens=4096,
                    system=system_prompt,
                    messages=[{
                        'role': 'user',
                        'content': [
                            {
                                'type': 'image',
                                'source': {
                                    'type': 'base64',
                                    'media_type': 'image/png',
                                    'data': b64,
                                },
                            },
                            {
                                'type': 'text',
                                'text': 'Answer the question in this image.',
                            },
                        ],
                    }],
                )
                text = response.content[0].text
                self.root.after(0, lambda t=text: on_result(t))
            except (anthropic.APIConnectionError, anthropic.APIStatusError):
                if self.gemini_vision_model:
                    gemini_vision_fallback()
                else:
                    err = 'ERROR: Claude API down and no Gemini fallback available'
                    self.root.after(0, lambda e=err: on_result(e))
            except Exception:
                if self.gemini_vision_model:
                    gemini_vision_fallback()
                else:
                    err = 'ERROR: Claude API failed and no Gemini fallback available'
                    self.root.after(0, lambda e=err: on_result(e))

        threading.Thread(target=worker, daemon=True).start()

    def _copy_to_clipboard(self, text):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update_idletasks()
        except Exception:
            pass

    # -------------------------------------------------------------------
    # Floating tooltip — shows answer near cursor, click to dismiss
    # -------------------------------------------------------------------
    def _show_tooltip(self, text, auto_ms=3000, keep_focus_hwnd=None):
        # Dismiss any existing tooltip
        self._dismiss_tooltip()

        # --- Tooltip with transparent bg, floating text only ---
        tip = tk.Toplevel()
        tip.overrideredirect(True)
        tip.attributes('-topmost', True)
        # Transparent background – only text visible
        tip_trans = '#f0f0f0'
        tip.configure(bg=tip_trans)
        tip.attributes('-transparentcolor', tip_trans)

        # Get user's chosen stealth text color
        fg_color = getattr(self, 'snip_text_var', None)
        fg_color = fg_color.get() if fg_color else 'black'
        font_family = getattr(self, 'snip_font_var', None)
        font_family = font_family.get() if font_family else 'Arial'
        font_size = getattr(self, 'snip_size_var', None)
        try:
            font_size = int(font_size.get()) if font_size else 12
        except Exception:
            font_size = 12

        # Position near cursor
        x = tip.winfo_pointerx() + 15
        y = tip.winfo_pointery() + 15

        lbl = tk.Label(tip, text=text, bg=tip_trans, fg=fg_color,
                       font=(font_family, font_size, 'bold'), wraplength=500,
                       justify='left', padx=6, pady=4)
        lbl.pack()

        tip.geometry(f'+{x}+{y}')
        tip.lift()

        if platform.system() == 'Windows':
            try:
                hwnd = tip.winfo_id()
                GWL_EXSTYLE = -20
                WS_EX_NOACTIVATE = 0x08000000
                SWP_NOMOVE = 0x0002
                SWP_NOSIZE = 0x0001
                SWP_NOACTIVATE = 0x0010
                HWND_TOPMOST = -1
                ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex | WS_EX_NOACTIVATE)
                ctypes.windll.user32.SetWindowPos(
                    hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                )
                if keep_focus_hwnd:
                    ctypes.windll.user32.SetForegroundWindow(keep_focus_hwnd)
            except Exception:
                pass

        self._tooltip = tip
        self._tooltip_dismiss_pending = False

        # Poll for mouse clicks using GetAsyncKeyState — works globally
        self._start_click_poll()

        # Auto-dismiss fail-safe
        tip.after(auto_ms, self._dismiss_tooltip)

    def _start_click_poll(self):
        """Poll mouse button state to detect the *next* click anywhere."""
        if platform.system() != 'Windows':
            return
        user32 = ctypes.windll.user32
        VK_LBUTTON = 0x01
        VK_RBUTTON = 0x02
        VK_MBUTTON = 0x04

        # Capture initial state so we only dismiss on a new click.
        prev = {
            VK_LBUTTON: bool(user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000),
            VK_RBUTTON: bool(user32.GetAsyncKeyState(VK_RBUTTON) & 0x8000),
            VK_MBUTTON: bool(user32.GetAsyncKeyState(VK_MBUTTON) & 0x8000),
        }

        def poll():
            if getattr(self, '_tooltip', None) is None:
                return
            for vk in (VK_LBUTTON, VK_RBUTTON, VK_MBUTTON):
                down = bool(user32.GetAsyncKeyState(vk) & 0x8000)
                # Dismiss on edge: button transitioned from up -> down
                if down and not prev[vk]:
                    self._dismiss_tooltip()
                    return
                prev[vk] = down
            # Keep polling
            try:
                self.root.after(25, poll)
            except Exception:
                pass

        # Small delay avoids consuming the click that ended the snip drag.
        try:
            self.root.after(150, poll)
        except Exception:
            pass

    def _dismiss_tooltip(self):
        tip = getattr(self, '_tooltip', None)
        if tip:
            try:
                tip.destroy()
            except Exception:
                pass
            self._tooltip = None

    # -------------------------------------------------------------------
    # Stealth helpers (hide taskbar + no-focus)
    # -------------------------------------------------------------------
    def apply_hide_taskbar(self):
        if platform.system() != 'Windows':
            return
        try:
            hwnd = self._get_toplevel_hwnd()
            GWL_EXSTYLE = -20
            WS_EX_TOOLWINDOW = 0x00000080
            WS_EX_APPWINDOW = 0x00040000
            SW_HIDE = 0
            SW_SHOW = 5

            if self.hide_taskbar_var.get():
                ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                self._orig_exstyle_taskbar = ex
                ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                new_ex = (ex & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_ex)
                ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                self.status.config(text='Hidden from taskbar')
            else:
                orig = getattr(self, '_orig_exstyle_taskbar', None)
                if orig is not None:
                    ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, orig)
                    ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                self.status.config(text='Shown in taskbar')
        except Exception as e:
            self.status.config(text=f'Hide taskbar failed: {e}')

    def apply_nofocus_mode(self):
        if platform.system() != 'Windows':
            return
        try:
            hwnd = self._get_toplevel_hwnd()
            GWL_EXSTYLE = -20
            WS_EX_NOACTIVATE = 0x08000000
            WS_EX_TOPMOST = 0x00000008
            SW_HIDE = 0
            SW_SHOWNOACTIVATE = 4
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            SWP_NOACTIVATE = 0x0010
            HWND_TOPMOST = -1
            HWND_NOTOPMOST = -2

            ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)

            if self.nofocus_var.get():
                # No-Focus requires On Top to work
                self.topmost_var.set(True)
                self.root.attributes('-topmost', True)
                self._orig_exstyle_nofocus = ex
                new_ex = ex | WS_EX_NOACTIVATE | WS_EX_TOPMOST
                ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_ex)
                ctypes.windll.user32.SetWindowPos(
                    hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                )
                ctypes.windll.user32.ShowWindow(hwnd, SW_SHOWNOACTIVATE)
                self.status.config(text='No-Focus ON \u2014 browser won\'t detect switch (uncheck to type)')
            else:
                orig = getattr(self, '_orig_exstyle_nofocus', None)
                if orig is not None:
                    ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, orig)
                    # Restore topmost based on On Top checkbox
                    if self.topmost_var.get():
                        ctypes.windll.user32.SetWindowPos(
                            hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                        )
                    else:
                        ctypes.windll.user32.SetWindowPos(
                            hwnd, HWND_NOTOPMOST, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                        )
                    ctypes.windll.user32.ShowWindow(hwnd, 5)  # SW_SHOW
                self.status.config(text='No-Focus OFF \u2014 you can type again')
        except Exception as e:
            self.status.config(text=f'No-focus failed: {e}')

    # -------------------------------------------------------------------
    # Hotkey listener
    # -------------------------------------------------------------------
    def _start_hotkey_listener(self):
        HOTKEY_TOGGLE = 1
        HOTKEY_QUIT = 2
        HOTKEY_SNIP = 3
        HOTKEY_TEXT_BACKTICK = 4
        HOTKEY_TEXT_F8 = 5
        HOTKEY_TEXT_CTRL_SHIFT_Q = 6
        HOTKEY_TEXT_CTRL_ALT_Q = 7
        MOD_NONE = 0x0000
        MOD_ALT = 0x0001
        MOD_CTRL = 0x0002
        MOD_SHIFT = 0x0004
        VK_BACKTICK = 0xC0  # VK_OEM_3 (`)
        VK_F8 = 0x77
        VK_Q = 0x51
        VK_F9 = 0x78
        VK_F10 = 0x79
        WM_HOTKEY = 0x0312

        user32 = ctypes.windll.user32

        def listener():
            if not user32.RegisterHotKey(None, HOTKEY_TOGGLE, MOD_NONE, VK_F9):
                return
            if not user32.RegisterHotKey(None, HOTKEY_QUIT, MOD_CTRL | MOD_SHIFT, VK_F9):
                user32.UnregisterHotKey(None, HOTKEY_TOGGLE)
                return

            user32.RegisterHotKey(None, HOTKEY_SNIP, MOD_NONE, VK_F10)

            text_hotkey_candidates = [
                (HOTKEY_TEXT_BACKTICK, MOD_NONE, VK_BACKTICK, '`'),
                (HOTKEY_TEXT_F8, MOD_NONE, VK_F8, 'F8'),
                (HOTKEY_TEXT_CTRL_SHIFT_Q, MOD_CTRL | MOD_SHIFT, VK_Q, 'Ctrl+Shift+Q'),
                (HOTKEY_TEXT_CTRL_ALT_Q, MOD_CTRL | MOD_ALT, VK_Q, 'Ctrl+Alt+Q'),
            ]
            active_text_hotkeys = []
            for hotkey_id, mods, vk, label in text_hotkey_candidates:
                if user32.RegisterHotKey(None, hotkey_id, mods, vk):
                    active_text_hotkeys.append(label)

            if active_text_hotkeys:
                joined = ', '.join(active_text_hotkeys)
                self.root.after(0, lambda j=joined: self.status.config(text=f'Text hotkeys active: {j}'))
            else:
                self.root.after(0, lambda: self.status.config(text='Failed to register text hotkey'))

            msg = wintypes.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                if msg.message == WM_HOTKEY:
                    if msg.wParam == HOTKEY_TOGGLE:
                        self.root.after(0, self.toggle_visibility)
                    elif msg.wParam == HOTKEY_QUIT:
                        self.root.after(0, self.on_close)
                    elif msg.wParam == HOTKEY_SNIP:
                        self.root.after(0, self._snip_stealth)
                    elif msg.wParam in (
                        HOTKEY_TEXT_BACKTICK,
                        HOTKEY_TEXT_F8,
                        HOTKEY_TEXT_CTRL_SHIFT_Q,
                        HOTKEY_TEXT_CTRL_ALT_Q,
                    ):
                        # Delay avoids racing the hotkey key-up and focus handoff.
                        self.root.after(120, self._stealth_text)

            user32.UnregisterHotKey(None, HOTKEY_TOGGLE)
            user32.UnregisterHotKey(None, HOTKEY_QUIT)
            user32.UnregisterHotKey(None, HOTKEY_SNIP)
            user32.UnregisterHotKey(None, HOTKEY_TEXT_BACKTICK)
            user32.UnregisterHotKey(None, HOTKEY_TEXT_F8)
            user32.UnregisterHotKey(None, HOTKEY_TEXT_CTRL_SHIFT_Q)
            user32.UnregisterHotKey(None, HOTKEY_TEXT_CTRL_ALT_Q)

        t = threading.Thread(target=listener, daemon=True)
        t.start()
        self._hotkey_thread = t

    # -------------------------------------------------------------------
    # File management
    # -------------------------------------------------------------------
    def load_file(self, path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = f.read()
        except Exception as e:
            messagebox.showerror('Error', f'Could not open file: {e}')
            return
        self.text.delete('1.0', tk.END)
        self.text.insert('1.0', data)
        self._parse_sections(data)
        self.status.config(text=f'Loaded: {path}')
        self.current_file = path
        self.clear_search()

    def open_file(self):
        path = filedialog.askopenfilename(
            title='Open notes file',
            filetypes=[('Text/Markdown', '*.txt *.md *.markdown'), ('All files', '*.*')],
        )
        if path:
            self.load_file(path)

    # -------------------------------------------------------------------
    # Config
    # -------------------------------------------------------------------
    def load_config(self):
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        except Exception:
            return
        self._config_loaded = True
        geom = cfg.get('geometry')
        if geom:
            try:
                self.root.geometry(geom)
            except Exception:
                pass
        last = cfg.get('last_file')
        if last and os.path.exists(last):
            try:
                self.load_file(last)
            except Exception:
                pass
        if cfg.get('topmost'):
            try:
                self.topmost_var.set(True)
                self.set_topmost()
            except Exception:
                pass
        if cfg.get('minimize_to_tray'):
            try:
                self.minimize_tray_var.set(True)
            except Exception:
                pass
        theme = cfg.get('theme', 'dark')
        if theme in THEMES and theme != 'dark':
            self.theme_var.set(theme)
            self.root.after(100, lambda: self.apply_theme(theme))
        if cfg.get('hide_taskbar'):
            self.hide_taskbar_var.set(True)
            self.root.after(200, self.apply_hide_taskbar)
        if cfg.get('nofocus'):
            self.nofocus_var.set(True)
            self.root.after(300, self.apply_nofocus_mode)
        snip_color = cfg.get('snip_text_color', 'black')
        if snip_color in ('black', 'white'):
            self.snip_text_var.set(snip_color)
        snip_font = cfg.get('snip_font_family', 'Arial')
        if snip_font in ('Arial', 'Consolas', 'Segoe UI'):
            self.snip_font_var.set(snip_font)
        snip_size = str(cfg.get('snip_font_size', '12'))
        if snip_size in ('10', '11', '12', '13', '14', '16', '18'):
            self.snip_size_var.set(snip_size)
        self.snip_hide_text_var.set(bool(cfg.get('snip_hide_text', False)))

    def save_config(self):
        cfg = {
            'geometry': self.root.winfo_geometry(),
            'last_file': getattr(self, 'current_file', None),
            'topmost': bool(self.topmost_var.get()),
            'minimize_to_tray': bool(self.minimize_tray_var.get()),
            'theme': self.current_theme,
            'hide_taskbar': bool(self.hide_taskbar_var.get()),
            'nofocus': bool(self.nofocus_var.get()),
            'snip_text_color': self.snip_text_var.get(),
            'snip_font_family': self.snip_font_var.get(),
            'snip_font_size': self.snip_size_var.get(),
            'snip_hide_text': bool(self.snip_hide_text_var.get()),
        }
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    # -------------------------------------------------------------------
    # Search
    # -------------------------------------------------------------------
    def clear_search(self):
        self.text.tag_remove('highlight', '1.0', tk.END)
        self.current_search = ''
        self.last_found_index = None

    def find_all(self):
        pattern = self.search_var.get().strip()
        self.text.tag_remove('highlight', '1.0', tk.END)
        if not pattern:
            self.status.config(text='Empty search')
            return
        start = '1.0'
        count = 0
        while True:
            idx = self.text.search(pattern, start, nocase=1, stopindex=tk.END)
            if not idx:
                break
            end = f'{idx}+{len(pattern)}c'
            self.text.tag_add('highlight', idx, end)
            start = end
            count += 1
        self.current_search = pattern
        self.last_found_index = None
        self.status.config(text=f'Found {count} matches for "{pattern}"')

    def find_next(self):
        pattern = self.search_var.get().strip()
        if not pattern:
            self.status.config(text='Empty search')
            return
        if self.current_search != pattern:
            self.find_all()
        start_index = '1.0'
        if self.last_found_index:
            start_index = f'{self.last_found_index}+1c'
        idx = self.text.search(pattern, start_index, nocase=1, stopindex=tk.END)
        if not idx:
            idx = self.text.search(pattern, '1.0', nocase=1, stopindex=tk.END)
            if not idx:
                self.status.config(text=f'No matches for "{pattern}"')
                return
        end = f'{idx}+{len(pattern)}c'
        self.text.tag_remove(tk.SEL, '1.0', tk.END)
        self.text.tag_add(tk.SEL, idx, end)
        self.text.mark_set(tk.INSERT, end)
        self.text.see(idx)
        self.last_found_index = idx
        self.status.config(text=f'Match at {idx}')

    # -------------------------------------------------------------------
    # Section parsing
    # -------------------------------------------------------------------
    def _parse_sections(self, text):
        self._full_text = text
        self._sections = {}
        pattern = re.compile(r'^(#+\s+.+)$', re.MULTILINE)
        matches = list(pattern.finditer(text))
        if not matches:
            return
        for i, m in enumerate(matches):
            name = m.group(1).strip()
            start = m.start()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            self._sections[name] = text[start:end].rstrip()
        desired = ['# Multiple Choice', '# Short Answer', '# Essay']
        self._section_label_map = {}
        for d in desired:
            matched = None
            for h in self._sections:
                if h.lower().startswith(d.lower()):
                    matched = h
                    break
            self._section_label_map[d] = matched

        menu = self.section_menu['menu']
        menu.delete(0, 'end')
        menu.add_command(label='All', command=lambda: self._set_section('All'))
        for d in desired:
            menu.add_command(label=d, command=lambda l=d: self._set_section(l))
        self.section_var.set('All')

    def _set_section(self, name):
        self.section_var.set(name)
        self._on_section_change(name)

    def _on_section_change(self, value):
        self.text.config(state='normal')
        try:
            self.text.tag_remove('section', '1.0', tk.END)
        except Exception:
            pass

        if value == 'All':
            self.text.delete('1.0', tk.END)
            self.text.insert('1.0', self._full_text)
        else:
            mapped = getattr(self, '_section_label_map', {}).get(value)
            content = self._sections.get(mapped, '') if mapped else ''
            if not content:
                self.status.config(text=f'Section "{value}" not found in notes')
                return
            self.text.delete('1.0', tk.END)
            self.text.insert('1.0', content)
            idx = self.text.search(r'^(#+\s+.+)$', '1.0', regexp=True, stopindex=tk.END)
            if idx:
                line_num = idx.split('.')[0]
                line_start = f'{line_num}.0'
                line_end = f'{line_num}.end'
                self.text.tag_add('section', line_start, line_end)
                self.text.mark_set(tk.INSERT, line_start)
                self.text.see(line_start)

        self.clear_search()
        self.status.config(text=f'Section: {value}')

    # -------------------------------------------------------------------
    # Window management
    # -------------------------------------------------------------------
    def set_topmost(self):
        v = self.topmost_var.get()
        try:
            self.root.attributes('-topmost', v)
        except Exception:
            pass

    def _get_toplevel_hwnd(self):
        self.root.update_idletasks()
        return ctypes.windll.user32.GetParent(int(self.root.winfo_id()))

    # --- System tray support ---
    def _create_image(self, width=64, height=64, color1=(48, 48, 48), color2=(200, 200, 200)):
        img = Image.new('RGB', (width, height), color1)
        d = ImageDraw.Draw(img)
        d.text((width * 0.2, height * 0.15), 'N', fill=color2)
        return img

    def _tray_worker(self, icon):
        try:
            icon.run()
        except Exception:
            pass

    def show_in_tray(self):
        if not TRAY_AVAILABLE:
            self.status.config(text='Install pystray + pillow for tray support')
            return
        if getattr(self, 'tray_icon', None) is not None:
            return
        image = self._create_image()
        menu = pystray.Menu(
            pystray.MenuItem('Restore', lambda: self.root.after(0, self._tray_restore)),
            pystray.MenuItem('Quit', lambda: self.root.after(0, self._tray_quit)),
        )
        icon = pystray.Icon('Service Host', image, 'Service Host\uA789 Windows Helper', menu)
        self.tray_icon = icon
        t = threading.Thread(target=self._tray_worker, args=(icon,), daemon=True)
        t.start()
        self.tray_thread = t
        self.status.config(text='Minimized to tray')

    def remove_tray(self):
        if getattr(self, 'tray_icon', None) is None:
            return
        try:
            self.tray_icon.stop()
        except Exception:
            pass
        self.tray_icon = None
        self.tray_thread = None

    def _tray_restore(self):
        self.remove_tray()
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.visible = True
        except Exception:
            pass

    def _tray_quit(self):
        self.remove_tray()
        self.on_close()

    def toggle_visibility(self):
        try:
            if self.visible:
                if self.minimize_tray_var.get():
                    self.root.withdraw()
                    self.show_in_tray()
                else:
                    self.root.withdraw()
                self.visible = False
            else:
                if getattr(self, 'tray_icon', None) is not None:
                    self.remove_tray()
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
                self.visible = True
        except Exception:
            pass

    def on_close(self):
        try:
            self.save_config()
        except Exception:
            pass
        self.root.destroy()


def main():
    root = tk.Tk()
    app = NoteAssistantApp(root)

    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if os.path.exists(arg):
            try:
                app.load_file(arg)
            except Exception:
                pass

    root.mainloop()


if __name__ == '__main__':
    main()
