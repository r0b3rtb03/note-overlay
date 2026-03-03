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
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from ctypes import wintypes


class NoteAssistantApp:
    def __init__(self, root, default_file='notes.txt'):
        self.root = root
        self.root.title('Study Notes Assistant')
        self.default_file = default_file
        self.visible = True
        self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'note_assistant_config.json')

        # Window size and appearance
        self.root.geometry('640x360')
        self.root.minsize(400, 200)
        self.bg = '#1e1e1e'
        self.fg = '#dcdcdc'
        self.entry_bg = '#2b2b2b'
        self.highlight_bg = '#44475a'
        self.root.configure(bg=self.bg)

        # Top frame with controls
        top = tk.Frame(self.root, bg=self.bg)
        top.pack(fill='x', padx=8, pady=6)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(top, textvariable=self.search_var, bg=self.entry_bg, fg=self.fg, insertbackground=self.fg)
        self.search_entry.pack(side='left', fill='x', expand=True, padx=(0, 6))
        self.search_entry.bind('<Return>', lambda e: self.find_next())

        find_btn = tk.Button(top, text='Find', command=self.find_all, bg=self.entry_bg, fg=self.fg)
        find_btn.pack(side='left')

        next_btn = tk.Button(top, text='Next', command=self.find_next, bg=self.entry_bg, fg=self.fg)
        next_btn.pack(side='left', padx=(4, 0))

        open_btn = tk.Button(top, text='Open', command=self.open_file, bg=self.entry_bg, fg=self.fg)
        open_btn.pack(side='left', padx=(8, 0))

        import_btn = tk.Button(top, text='Import', command=self.import_and_copy, bg=self.entry_bg, fg=self.fg)
        import_btn.pack(side='left', padx=(6, 0))

        # Section filter dropdown
        tk.Label(top, text='Section:', bg=self.bg, fg=self.fg).pack(side='left', padx=(8, 2))
        self.section_var = tk.StringVar(value='All')
        self.section_menu = tk.OptionMenu(top, self.section_var, 'All', command=self._on_section_change)
        self.section_menu.config(bg=self.entry_bg, fg=self.fg, highlightthickness=0)
        self.section_menu['menu'].config(bg=self.entry_bg, fg=self.fg)
        self.section_menu.pack(side='left')

        self.topmost_var = tk.BooleanVar(value=False)
        topmost_cb = tk.Checkbutton(top, text='Always on Top', variable=self.topmost_var, command=self.set_topmost, bg=self.bg, fg=self.fg, selectcolor=self.bg, activebackground=self.bg)
        topmost_cb.pack(side='left', padx=(8, 0))

        # Option to hide from taskbar (Windows only)
        self.hide_taskbar_var = tk.BooleanVar(value=False)
        hide_cb = tk.Checkbutton(top, text='Hide from Taskbar', variable=self.hide_taskbar_var, command=self.apply_taskbar_setting, bg=self.bg, fg=self.fg, selectcolor=self.bg, activebackground=self.bg)
        hide_cb.pack(side='left', padx=(8, 0))

        # Minimize to tray option
        self.minimize_tray_var = tk.BooleanVar(value=False)
        tray_cb = tk.Checkbutton(top, text='Minimize to Tray', variable=self.minimize_tray_var, bg=self.bg, fg=self.fg, selectcolor=self.bg, activebackground=self.bg)
        tray_cb.pack(side='left', padx=(8, 0))
        
        # Stronger hide button (Windows only)
        strong_btn = tk.Button(top, text='Apply Stronger Hide', command=self.apply_stronger_hide, bg=self.entry_bg, fg=self.fg)
        strong_btn.pack(side='left', padx=(8, 0))

        # Text area for notes
        self.text = scrolledtext.ScrolledText(self.root, wrap='word', bg='#121212', fg=self.fg, insertbackground=self.fg)
        self.text.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self.text.tag_config('highlight', background=self.highlight_bg)

        # Status bar
        self.status = tk.Label(self.root, text='', anchor='w', bg=self.bg, fg=self.fg)
        self.status.pack(fill='x', padx=8, pady=(0, 6))

        # Internal state
        self.current_search = ''
        self.last_found_index = None
        self._full_text = ''
        self._sections = {}  # name -> text content

        # Load config (geometry, last file, topmost)
        self.load_config()

        # If no file loaded from config, load default file if present
        if not getattr(self, 'current_file', None):
            if os.path.exists(self.default_file):
                self.load_file(self.default_file)
            else:
                self.status.config(text=f'No {self.default_file} found — use Open to load notes')

        # Global hotkeys via Win32 RegisterHotKey (no admin, works in all apps)
        self._hotkey_thread = None
        if platform.system() == 'Windows':
            self._start_hotkey_listener()
        else:
            self.status.config(text='Global hotkeys only supported on Windows')

        # Ensure we unhook on exit
        self.root.protocol('WM_DELETE_WINDOW', self.on_close)

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
            # RegisterHotKey is thread-bound — register in this thread
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
        path = filedialog.askopenfilename(title='Select notes file to import', filetypes=[('Text/Markdown', '*.txt *.md *.markdown'), ('All files', '*.*')])
        if not path:
            return
        try:
            dest = os.path.join(os.path.dirname(os.path.abspath(__file__)), self.default_file)
            shutil.copy(path, dest)
            self.load_file(dest)
            self.status.config(text=f'Imported and copied to {self.default_file}')
        except Exception as e:
            messagebox.showerror('Import Error', f'Could not import file: {e}')

    def load_config(self):
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        except Exception:
            return
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
        topmost = cfg.get('topmost')
        if topmost:
            try:
                self.topmost_var.set(True)
                self.set_topmost()
            except Exception:
                pass
        hide_taskbar = cfg.get('hide_from_taskbar')
        if hide_taskbar:
            self.hide_taskbar_var.set(True)
            # Defer the actual Win32 call until the window is mapped
            self.root.after(100, self.apply_taskbar_setting)
        minimize_tray = cfg.get('minimize_to_tray')
        if minimize_tray:
            try:
                self.minimize_tray_var.set(True)
            except Exception:
                pass

    def save_config(self):
        cfg = {
            'geometry': self.root.winfo_geometry(),
            'last_file': getattr(self, 'current_file', None),
            'topmost': bool(self.topmost_var.get()),
            'hide_from_taskbar': bool(self.hide_taskbar_var.get()),
            'minimize_to_tray': bool(self.minimize_tray_var.get()),
        }
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    def open_file(self):
        path = filedialog.askopenfilename(title='Open notes file', filetypes=[('Text/Markdown', '*.txt *.md *.markdown'), ('All files', '*.*')])
        if path:
            self.load_file(path)

    def clear_search(self):
        self.text.tag_remove('highlight', '1.0', tk.END)
        self.current_search = ''
        self.last_found_index = None

    def _parse_sections(self, text):
        """Parse markdown-style # headings into sections."""
        self._full_text = text
        self._sections = {}
        # Find all lines starting with # (top-level headings)
        pattern = re.compile(r'^(#+\s+.+)$', re.MULTILINE)
        matches = list(pattern.finditer(text))
        if not matches:
            return
        for i, m in enumerate(matches):
            name = m.group(1).strip()
            start = m.start()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            self._sections[name] = text[start:end].rstrip()
        # Update the dropdown
        menu = self.section_menu['menu']
        menu.delete(0, 'end')
        menu.add_command(label='All', command=lambda: self._set_section('All'))
        for name in self._sections:
            menu.add_command(label=name, command=lambda n=name: self._set_section(n))
        self.section_var.set('All')

    def _set_section(self, name):
        self.section_var.set(name)
        self._on_section_change(name)

    def _on_section_change(self, value):
        self.text.config(state='normal')
        self.text.delete('1.0', tk.END)
        if value == 'All':
            self.text.insert('1.0', self._full_text)
        else:
            content = self._sections.get(value, '')
            self.text.insert('1.0', content)
        self.clear_search()
        self.status.config(text=f'Section: {value}')

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
            # new search
            self.find_all()
        start_index = '1.0'
        if self.last_found_index:
            # move after last found
            start_index = f'{self.last_found_index}+1c'
        idx = self.text.search(pattern, start_index, nocase=1, stopindex=tk.END)
        if not idx:
            # wrap around
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
        self.status.config(text=f'Showing match at {idx}')

    def set_topmost(self):
        v = self.topmost_var.get()
        try:
            self.root.attributes('-topmost', v)
        except Exception:
            pass

    def _get_toplevel_hwnd(self):
        """Get the real Win32 toplevel HWND (not tkinter's inner frame)."""
        self.root.update_idletasks()
        return ctypes.windll.user32.GetParent(int(self.root.winfo_id()))

    def apply_taskbar_setting(self):
        if platform.system() != 'Windows':
            self.status.config(text='Hide-from-taskbar only supported on Windows')
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
                # Hide, change style, show — Windows refreshes taskbar on show
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
            self.status.config(text=f'Could not change taskbar visibility: {e}')

    def apply_stronger_hide(self):
        if platform.system() != 'Windows':
            self.status.config(text='Stronger hide only supported on Windows')
            return
        try:
            hwnd = self._get_toplevel_hwnd()
            GWL_EXSTYLE = -20
            WS_EX_TOOLWINDOW = 0x00000080
            WS_EX_APPWINDOW = 0x00040000
            SW_HIDE = 0
            SW_SHOW = 5

            ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            if not getattr(self, '_strong_hide', False):
                self._orig_exstyle = ex
                new_ex = (ex & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
                ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_ex)
                ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                self._strong_hide = True
                self.status.config(text='Applied stronger hide')
            else:
                orig = getattr(self, '_orig_exstyle', None)
                if orig is not None:
                    ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, orig)
                    ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                self._strong_hide = False
                self.status.config(text='Reverted stronger hide')
        except Exception as e:
            self.status.config(text=f'Stronger hide failed: {e}')

    # --- System tray support ---
    def _create_image(self, width=64, height=64, color1=(48,48,48), color2=(200,200,200)):
        img = Image.new('RGB', (width, height), color1)
        d = ImageDraw.Draw(img)
        # simple 'N' letter
        d.text((width*0.2, height*0.15), 'N', fill=color2)
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
        icon = pystray.Icon('note_assistant', image, 'Study Notes Assistant', menu)
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
        self.status.config(text='Removed tray icon')

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
                # hide: if minimize-to-tray enabled, use tray icon; otherwise withdraw
                if self.minimize_tray_var.get():
                    self.root.withdraw()
                    self.show_in_tray()
                else:
                    self.root.withdraw()
                self.visible = False
            else:
                # restore from tray if present
                if getattr(self, 'tray_icon', None) is not None:
                    self.remove_tray()
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
                self.visible = True
        except Exception:
            pass

    def on_close(self):
        # save window state
        try:
            self.save_config()
        except Exception:
            pass
        # Hotkey thread is a daemon — exits automatically when process ends
        self.root.destroy()


def main():
    root = tk.Tk()
    app = NoteAssistantApp(root)

    # If a file path was provided on the command line, load it
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if os.path.exists(arg):
            try:
                app.load_file(arg)
            except Exception:
                pass

    # Run the tkinter mainloop in the main thread (keyboard hook runs in background)
    root.mainloop()


if __name__ == '__main__':
    main()
