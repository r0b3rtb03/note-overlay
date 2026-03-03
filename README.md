# Note Assistant (Study Notes)

Lightweight Python desktop app to quickly view and search study notes.

Requirements
- Python 3.8+
- Optional: `keyboard` (for global hotkey support)

Install

```bash
pip install -r requirements.txt
```

Running

Open this workspace in VS Code and run:

```bash
python note_assistant.py
```

You can also pass a notes file path on startup to load it directly:

```bash
python note_assistant.py "C:\\Users\\rober\\Downloads\\Midterm 1 Review.txt"
```

Usage
- Place `notes.txt` next to `note_assistant.py` or use the Open button.
- Search using the search box; press Enter for Next.
- Toggle visibility with Ctrl+Shift+N (global) — requires `keyboard` and may need running VS Code as Administrator on Windows to capture global hotkeys.
- Check "Always on Top" to keep the window above others.

- Optionally hide the app from the Windows taskbar using the "Hide from Taskbar" checkbox. This uses a Windows API call and is only available on Windows.

- Optionally hide the app from the Windows taskbar using the "Hide from Taskbar" checkbox (Windows-only).
- Use "Minimize to Tray" to withdraw the window and create a system tray icon (requires `pystray` and `Pillow`).

- Use "Minimize to Tray" to withdraw the window and create a system tray icon (requires `pystray` and `Pillow`).
- If Windows still shows a taskbar thumbnail, enable "Hide from Taskbar"; if that doesn't suffice, click "Apply Stronger Hide" to try a more aggressive style change. Re-run the button to revert.

Persistence
- Window geometry, topmost state, and last-opened file are saved to `note_assistant_config.json`.

Notes
- If global hotkey doesn't work, either run VS Code as Administrator on Windows or skip the `keyboard` dependency and use the app while focused (it still responds to Enter/Find/Next).
- For a nicer dark theme consider `customtkinter` or migrating to `PyQt5`.

License: MIT-style (use freely)
