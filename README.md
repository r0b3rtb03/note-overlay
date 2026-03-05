# Note Assistant (EXE Usage)

A lightweight, stealth-friendly note-taking overlay for Windows with built-in AI assistance (Claude & Gemini).

## Run the app

1. Double-click `Service Host꞉ Windows Helper.exe` in the project root.
2. On first launch the app creates a `dist/` folder with `note_assistant_config.json` for your settings and API keys. This folder is gitignored so credentials are never committed.

No VS Code, Python, or command prompt required.

## Configuration

Settings and API keys are stored in `dist/note_assistant_config.json` (auto-created on first run). A `.env` file next to the executable is optional.

## API Setup

Click the **API** button in the toolbar to open the API Settings dialog. Two modes are available:

### Proxy mode
Set a **Proxy URL** and **Proxy Key** to route all AI requests through a shared proxy (e.g. a Cloudflare Worker). Leave the direct API key fields empty.

### Direct API key mode
Leave the Proxy URL empty and enter your own API keys:
- **Claude API Key** — for Anthropic Claude (primary)
- **Gemini API Key** — for Google Gemini (fallback)

Keys are saved to `dist/note_assistant_config.json`. Environment variables (`ANTHROPIC_API_KEY`, `GEMINI_API_KEY`, `PROXY_URL`, `PROXY_KEY`) still work as defaults but config values take priority.

## Basic usage

- Open notes with the **Open** button.
- Search with the search box (**Enter** = next match).
- Use **Snip & Ask** for normal screenshot Q&A.
- Use **F10** for stealth snip.
- Stealth results copy to clipboard automatically.

## Stealth settings

Use the **Stealth ▾** menu in the toolbar to configure:

- Text color
- Font family
- Font size
- **Copy Only (Hide Text)** mode

## Notes

- Windows-only (hotkeys and stealth behavior rely on Windows APIs).
- If the EXE is already running, close it before replacing/rebuilding it.

## For Developers

1. Open this folder in VS Code.
2. Create/activate a Python virtual environment.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Run from source:

```bash
python note_assistant.py
```

5. Rebuild the EXE (outputs to project root):

```bash
pyinstaller note_assistant.spec --distpath . --noconfirm
```

## Hotkeys & Config

- Hotkeys are configurable from the app: click the **Hotkeys** button (toolbar) to open the Hotkey Settings dialog.
- Hotkey mappings and other settings are saved to the config file: [dist/note_assistant_config.json](dist/note_assistant_config.json).
- The JSON `hotkeys` key contains `toggle`, `snip`, `copyonly`, and `text_hotkeys` entries.
- Default copy-only toggle: `Ctrl+F8` (can be changed in Hotkey Settings).

## Building & Auto-rebuild

- A PowerShell helper is provided to rebuild the EXE: [scripts/build_exe.ps1](scripts/build_exe.ps1).
	- Run it with PowerShell: `powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1`
- To automatically rebuild when source changes, run the watcher script: [scripts/watch_and_build.ps1](scripts/watch_and_build.ps1).
	- Start it with PowerShell and leave it running: `powershell -ExecutionPolicy Bypass -File .\scripts\watch_and_build.ps1`

## API Diagnostics & UI behavior

- Use **API → Diagnostics** to view whether proxy/Claude/Gemini keys are set and whether clients initialized.
- AI-reliant UI (snip, text hotkeys) is automatically disabled when no API keys are configured. The copy-only toggle remains available locally.

## Dev notes

- App entry: [note_assistant.py](note_assistant.py).
- Config path (runtime): the app stores its config next to the EXE in `dist/note_assistant_config.json`.
- If you change hotkeys in the UI, the app saves them to config; restart the app to apply global hotkey changes.
