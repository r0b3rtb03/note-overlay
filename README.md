# Note Assistant (EXE Usage)

This project is intended to be run from the packaged executable in `dist/`.

## Run the app

1. Open the `dist` folder.
2. Double-click `Service Host꞉ Windows Helper.exe`.

No VS Code launch or command prompt command is required.

## Required files

- Keep your `.env` file next to the executable if you use API features (Claude/Gemini/Copilot/OCR).
- `note_assistant_config.json` is created/updated next to the executable and stores your UI/settings state.

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
