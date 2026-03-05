# Note Assistant (EXE Usage)

This project is intended to be run from the packaged executable in `dist/`.

## Run the app

1. Open the `dist` folder.
2. Double-click `Service Host꞉ Windows Helper.exe`.

No VS Code launch or command prompt command is required.

## Required files

- `note_assistant_config.json` is created/updated next to the executable and stores your UI/settings state (including API keys).
- A `.env` file next to the executable is optional — API keys can also be configured from within the app.

## API Setup

Click the **API** button in the toolbar to open the API Settings dialog. Two modes are available:

### Proxy mode
Set a **Proxy URL** and **Proxy Key** to route all AI requests through a shared proxy (e.g. a Cloudflare Worker). Leave the direct API key fields empty.

### Direct API key mode
Leave the Proxy URL empty and enter your own API keys:
- **Claude API Key** — for Anthropic Claude (primary)
- **Gemini API Key** — for Google Gemini (fallback)

Keys are saved to `note_assistant_config.json`. Environment variables (`ANTHROPIC_API_KEY`, `GEMINI_API_KEY`, `PROXY_URL`, `PROXY_KEY`) still work as defaults but config values take priority.

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

- This app is Windows-focused (hotkeys and stealth behavior rely on Windows APIs).
- If the EXE is already running, close it before replacing/rebuilding it.

## For Developers (VS Code)

If you want to edit or debug the source code instead of using the packaged EXE:

1. Open this folder in VS Code.
2. Create/activate a Python environment.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Run from source:

```bash
python note_assistant.py
```

5. Rebuild the EXE after changes (from this workspace root):

```bash
python -m PyInstaller note_assistant.spec
```
