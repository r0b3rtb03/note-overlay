import os
import sys
import re
import threading
import json
import shutil
import ctypes
import platform
import threading as _threading
try:
    import pystray
    from PIL import Image, ImageDraw
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


class NoteAssistantApp:
    def __init__(self, root, default_file='notes.txt'):
        self.root = root
        self.root.title('Service Host꞉ Windows Helper')
        self.default_file = default_file
        self.visible = True
        self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'note_assistant_config.json')

        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        self.config_path = os.path.join(base_dir, 'note_assistant_config.json')

        # Claude API client
        self.anthropic_client = None
        api_key = os.environ.get('ANTHROPIC_API_KEY')
        if ANTHROPIC_AVAILABLE and api_key:
            try:
                self.anthropic_client = anthropic.Anthropic(api_key=api_key)
            except Exception as e:
                print(f'Failed to init Claude client: {e}')

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
        self.root.rowconfigure(3, weight=0)  # chat panel (fixed height, not in grid until toggled)
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

        self.import_btn = tk.Button(btn_frame, text='Import', command=self.import_and_copy,
                                    bg=self.button_bg, fg=self.button_fg,
                                    relief='flat', padx=8, font=('Segoe UI', 9))
        self.import_btn.pack(side='left', padx=1)

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

        # Right group: window controls (simplified)
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
        if self.anthropic_client:
            self.status.config(text='Claude API ready')
        elif ANTHROPIC_AVAILABLE:
            self.status.config(text='Claude API: no API key found in .env')
        else:
            self.status.config(text='Claude API: anthropic package not installed')

        # Widget collections for theme updates
        self._all_buttons = [
            self.find_btn, self.next_btn, self.open_btn, self.import_btn,
            self.format_btn, self.toggle_prompt_btn,
        ]
        self._all_checkbuttons = [
            self.topmost_cb, self.tray_cb, self.hide_cb, self.nofocus_cb,
        ]
        self._all_labels = [
            self.section_label, self.theme_label, self.prompt_label, self.status,
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

        # Load config — if no config exists, enable hide taskbar + no-focus by default
        self._config_loaded = False
        self.load_config()
        if not self._config_loaded:
            self.hide_taskbar_var.set(True)
            self.nofocus_var.set(True)
            self.root.after(200, self.apply_hide_taskbar)
            self.root.after(300, self.apply_nofocus_mode)

        if not getattr(self, 'current_file', None):
            if os.path.exists(self.default_file):
                self.load_file(self.default_file)
            else:
                self.status.config(text=f'No {self.default_file} found — use Open to load notes')

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

    # -------------------------------------------------------------------
    # Claude API
    # -------------------------------------------------------------------
    def _call_claude(self, system_prompt, user_message, callback,
                     model='claude-sonnet-4-6'):
        if not self.anthropic_client:
            callback('ERROR: Claude API not available.\n\nMake sure your .env file has ANTHROPIC_API_KEY set and the anthropic package is installed.')
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
            except Exception as e:
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
            'For multiple choice questions, just state the answer letter and option. '
            'For short answer questions, write a concise 1 sentence answer at a'
            'high school to freshman college level. Keep it clear and natural. '
            'For essay questions, write a solid paragraph at a '
            'high school to freshman college level. Keep it clear and natural. '
            'No explanations for multiple choice unless asked.'
        )

        self.status.config(text='Asking Claude...')
        self.send_btn.config(state='disabled')
        saved_q = question

        def on_result(result):
            self.send_btn.config(state='normal')
            self.response_text.config(state='normal')
            self.response_text.delete('1.0', tk.END)
            self.response_text.insert('1.0', result)
            self.response_text.see('1.0')
            self.response_text.config(state='disabled')
            if result.startswith('ERROR:'):
                self.status.config(text='Claude query failed — see response')
            else:
                self.status.config(text='Response received')
            self.prompt_var.set('')

        self._call_claude(system_prompt, full_message, on_result)

    # -------------------------------------------------------------------
    # Stealth mode (hide taskbar + tool window + no-focus + always on top)
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
                self._orig_exstyle_nofocus = ex
                new_ex = ex | WS_EX_NOACTIVATE | WS_EX_TOPMOST
                ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_ex)
                ctypes.windll.user32.SetWindowPos(
                    hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                )
                ctypes.windll.user32.ShowWindow(hwnd, SW_SHOWNOACTIVATE)
                self.status.config(text='No-Focus ON — browser won\'t detect switch (uncheck to type)')
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
                self.status.config(text='No-Focus OFF — you can type again')
        except Exception as e:
            self.status.config(text=f'No-focus failed: {e}')

    # -------------------------------------------------------------------
    # Hotkey listener
    # -------------------------------------------------------------------
    def _start_hotkey_listener(self):
        HOTKEY_TOGGLE = 1
        HOTKEY_QUIT = 2
        MOD_NONE = 0x0000
        MOD_CTRL = 0x0002
        MOD_SHIFT = 0x0004
        VK_F9 = 0x78
        WM_HOTKEY = 0x0312

        user32 = ctypes.windll.user32

        def listener():
            if not user32.RegisterHotKey(None, HOTKEY_TOGGLE, MOD_NONE, VK_F9):
                return
            if not user32.RegisterHotKey(None, HOTKEY_QUIT, MOD_CTRL | MOD_SHIFT, VK_F9):
                user32.UnregisterHotKey(None, HOTKEY_TOGGLE)
                return

            msg = wintypes.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                if msg.message == WM_HOTKEY:
                    if msg.wParam == HOTKEY_TOGGLE:
                        self.root.after(0, self.toggle_visibility)
                    elif msg.wParam == HOTKEY_QUIT:
                        self.root.after(0, self.on_close)

            user32.UnregisterHotKey(None, HOTKEY_TOGGLE)
            user32.UnregisterHotKey(None, HOTKEY_QUIT)

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

    def import_and_copy(self):
        path = filedialog.askopenfilename(
            title='Select notes file to import',
            filetypes=[('Text/Markdown', '*.txt *.md *.markdown'), ('All files', '*.*')],
        )
        if not path:
            return
        try:
            dest = os.path.join(os.path.dirname(os.path.abspath(__file__)), self.default_file)
            shutil.copy(path, dest)
            self.load_file(dest)
            self.status.config(text=f'Imported and copied to {self.default_file}')
        except Exception as e:
            messagebox.showerror('Import Error', f'Could not import file: {e}')

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

    def save_config(self):
        cfg = {
            'geometry': self.root.winfo_geometry(),
            'last_file': getattr(self, 'current_file', None),
            'topmost': bool(self.topmost_var.get()),
            'minimize_to_tray': bool(self.minimize_tray_var.get()),
            'theme': self.current_theme,
            'hide_taskbar': bool(self.hide_taskbar_var.get()),
            'nofocus': bool(self.nofocus_var.get()),
        }
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    # -------------------------------------------------------------------
    # File dialogs
    # -------------------------------------------------------------------
    def open_file(self):
        path = filedialog.askopenfilename(
            title='Open notes file',
            filetypes=[('Text/Markdown', '*.txt *.md *.markdown'), ('All files', '*.*')],
        )
        if path:
            self.load_file(path)

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
            if not mapped:
                self._ensure_heading_exists(value)
                mapped = getattr(self, '_section_label_map', {}).get(value)

            content = self._sections.get(mapped, '') if mapped else ''
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

    def _ensure_heading_exists(self, display_label):
        if display_label.lower().startswith('# multiple'):
            mc_block = """# Multiple Choice
1. Which NIST cloud characteristic describes a multi-tenant environment where multiple customers share physical resources while remaining logically isolated?
    a. Rapid elasticity
    b. **Resource pooling**
    c. Measured service
    d. Location Independence

2. In Amazon S3 (Object Storage), if a single byte of a 5TB file needs to be changed, how does the system execute the update?
    a. It changes only the corresponding storage block.
    b. **It must update and replace the entire file.**
    c. It applies a delta patch via multipart upload.
    d. It alters the object's metadata pointer.

3. Which Route 53 routing policy calculates the optimal endpoint by analyzing geographic coordinates and allows an architect to expand or shrink a region's influence by applying a 'Bias' value?
    a. **Geoproximity routing**
    b. Latency-based routing
    c. Geolocation routing
    d. Weighted routing

4. When creating an Amazon EFS architecture for multiple EC2 instances, where should the EFS mount targets be placed?
    a. **In a private subnet, one per Availability Zone**
    b. In a public subnet attached to an Internet Gateway
    c. In an S3 bucket configured for VPC access
    d. Directly on the Transit Gateway

5. Which NIST service model describes a fully functional application, like Microsoft 365, where the user only needs to provide configuration?
    a. **Software as a Service (SaaS)**
    b. Platform as a Service (PaaS)
    c. Infrastructure as a Service (IaaS)
    d. Anything as a Service (XaaS)

6. What type of application architecture is specifically designed for the cloud, relying on automatic horizontal scaling and stateless execution?
    a. **Cloud native**
    b. Cloud enabled
    c. Lifted and shifted
    d. Monolithic

7. Which type of hypervisor runs directly on the server's bare metal hardware rather than running inside a host operating system?
    a. Type 2 Hypervisor
    b. **Type 1 Hypervisor**
    c. Virtual Machine Monitor
    d. Docker Engine

8. In the OSI Model, which layer handles host-to-host end-to-end addressing and routing?
    a. Layer 2 - Data Link
    b. **Layer 3 - Network**
    c. Layer 4 - Transport
    d. Layer 7 - Application

9. Which Transport Layer protocol drops error recovery and connection establishment features to optimize real-time voice and video streaming?
    a. TCP
    b. **UDP**
    c. IP
    d. DHCP

10. What is the mathematical formula used to calculate the number of physical links required to achieve a full mesh topology between 'n' nodes?
    a. 2^n - 2
    b. **n(n-1)/2**
    c. n(n+1)/2
    d. n^2

11. How many bits make up an IPv6 address?
    a. 32
    b. 48
    c. 64
    d. **128**

12. When creating a VPC in AWS, what is the largest allowable IPv4 CIDR block that can be assigned?
    a. /8
    b. **/16**
    c. /24
    d. /32

13. Which managed AWS service is deployed into a public subnet to enable instances in a private subnet to initiate IPv4 outbound traffic to the internet while preventing the internet from initiating a connection with those instances?
    a. Internet Gateway
    b. **NAT Gateway**
    c. Virtual Private Gateway
    d. VPC Peering Connection

14. Which AWS connectivity feature establishes a direct, non-transitive network relationship between two VPCs, allowing them to communicate using private IP addresses over the AWS global backbone?
    a. **VPC Peering**
    b. VPC Endpoints
    c. AWS PrivateLink
    d. Transit Gateway

15. What AWS networking component allows instances to securely connect to regional services like Amazon S3 and DynamoDB without routing traffic over the public internet?
    a. NAT Gateway
    b. Internet Gateway
    c. **VPC Endpoints**
    d. Network ACL

16. Which Route 53 policy distributes DNS responses based on assigned proportions, making it ideal for A/B testing?
    a. **Weighted routing**
    b. Latency based routing
    c. Simple routing
    d. Failover routing

17. Which Amazon EC2 instance family type is physically optimized specifically for in-memory databases (e.g., r4, r5)?
    a. Compute Optimized
    b. **Memory Optimized**
    c. General Purpose
    d. Accelerated Compute

18. What EC2 deployment feature allows you to pass a bash or PowerShell script to automate patching and software installations during the instance launch?
    a. **User Data**
    b. Instance Metadata
    c. Elastic Block Store
    d. IAM Instance Profile

19. Which EC2 storage option is physically attached to the underlying host computer and is considered ephemeral because its data does not persist if the instance transitions to a Stopped or Terminated state?
    a. Elastic Block Store (EBS)
    b. Amazon S3
    c. **Instance Store**
    d. Elastic File System (EFS)

20. Which special IP address can be queried from within a running EC2 instance to dynamically retrieve its metadata, such as its IAM role or public IP?
    a. http://192.168.1.1/latest/meta-data
    b. **http://169.254.169.254/latest/meta-data**
    c. http://10.0.0.1/latest/meta-data
    d. http://127.0.0.1/latest/meta-data

21. Which EC2 pricing model grants access to spare AWS compute capacity at a steep discount for fault-tolerant workloads, but may be reclaimed with a two-minute interruption notice?
    a. On Demand
    b. Reserved Instances
    c. **Spot Instances**
    d. Dedicated Hosts

22. Which Amazon ECS launch type abstracts the underlying infrastructure, allowing you to run containers without having to provision, configure, or scale the EC2 instances that form the cluster?
    a. **Fargate**
    b. Kubernetes
    c. Elastic Beanstalk
    d. Lightsail

23. Which serverless compute service executes code in response to triggers and maintains a sub-second billing model based on the exact duration of the execution and the memory allocated?
    a. Elastic Beanstalk
    b. Amazon EKS
    c. **AWS Lambda**
    d. AWS Batch

24. When provisioning an EBS volume, to what specific AWS infrastructure boundary is its availability natively restricted?
    a. Region
    b. **Availability Zone**
    c. Edge Location
    d. Virtual Private Cloud (entire VPC)

25. After establishing a full baseline backup of an EBS volume, how does AWS optimize subsequent storage usage for new snapshots?
    a. It compresses the full drive into a ZIP file.
    b. **It performs incremental backups by saving only the modified blocks.**
    c. It moves the snapshot directly to S3 Glacier Deep Archive.
    d. It duplicates the baseline to an alternative Region.

26. Which S3 Storage Class offers the lowest cost for long-term data retention (7-10 years) but imposes a retrieval time of up to 12 hours?
    a. S3 Standard-IA
    b. S3 One Zone-IA
    c. S3 Glacier Flexible Retrieval
    d. **S3 Glacier Deep Archive**

27. Amazon EFS provides scalable file storage. Which network protocol does EFS rely on to allow Linux EC2 instances to mount the file system?
    a. SMB
    b. iSCSI
    c. **NFSv4.1**
    d. FTP

28. When editing /etc/fstab to automatically mount a network drive in AWS, which parameter is critical to prevent the OS from hanging during boot if the storage isn't ready?
    a. defaults
    b. noauto
    c. rw
    d. **_netdev**

29. Under the AWS Shared Responsibility Model, which of the following is an example of AWS's responsibility 'of' the cloud?
    a. Managing S3 bucket policies
    b. **Protecting the physical security of data center facilities**
    c. Patching the guest OS on an EC2 instance
    d. Configuring IAM user passwords

30. Which AWS service is responsible for distributing incoming application traffic across multiple targets, such as EC2 instances, to ensure high availability?
    a. Amazon Route 53
    b. **Elastic Load Balancing (ELB)**
    c. Amazon CloudFront
    d. AWS Direct Connect

31. A company decides to 'Refactor' their application during a cloud migration. What does this process typically involve?
    a. **Modifying or redesigning the application to take advantage of cloud-native services (e.g., serverless).**
    b. Recreating the exact same on-premise architecture using VMs.
    c. Moving the physical servers into a colocation data center.
    d. Changing the billing model without changing the software.
"""
            if not self._full_text.endswith('\n'):
                self._full_text += '\n\n'
            self._full_text += mc_block + '\n\n'
            self._parse_sections(self._full_text)
            return

        for h in self._sections:
            if h.lower().startswith(display_label.lower()):
                return

        if not self._full_text.endswith('\n'):
            self._full_text += '\n\n'
        if display_label.strip() == '# Short Answer':
            self._full_text += '# Short Answer\n\n- Prompt:\n'
        elif display_label.strip() == '# Essay':
            self._full_text += '# Essay\n\n- Topic:\n'
        else:
            self._full_text += display_label + '\n\n'

        self._parse_sections(self._full_text)

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
        icon = pystray.Icon('Service Host', image, 'Service Host꞉ Windows Helper', menu)
        self.tray_icon = icon
        t = _threading.Thread(target=self._tray_worker, args=(icon,), daemon=True)
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
