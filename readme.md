# Assignment Tracker

## GUI App

Run the GUI:

`python AssignmentTrackerGUI.py`

What it does:
- Checks/installs required dependencies (`playwright`) and Chromium.
- Opens Canvas login in Chromium.
- After sign-in, shows 3 sync buttons:
  - Sync all assignments
  - Sync future assignments
  - Dry-sync (no sheet writes)
- Shows live logs in the right-side console panel.

## Build Single EXE

Install PyInstaller:

`python -m pip install pyinstaller`

One-command build (PowerShell):

`./build_exe.ps1`

Build without icon:

`./build_exe.ps1 -NoIcon`

Build one-file windowed EXE:

`pyinstaller --onefile --windowed --name AssignmentTrackerGUI --icon app.ico AssignmentTrackerGUI.py`

Icon file location:
- Put your icon at `app.ico` in the project root (or `assets/app.ico` for the runtime GUI window icon).

Output EXE path:

`dist/AssignmentTrackerGUI.exe`

## Config

For source code runs, config comes from `keys.py` defaults/env or `keys.local.json`.

For EXE runs (no rebuild needed), create `keys.local.json` in the same folder as the EXE:

```json
{
	"SHEET_API_URL": "https://script.google.com/macros/s/YOUR_DEPLOYMENT/exec",
	"CANVAS_BASE_URL": "https://umsystem.instructure.com"
}
```

Quick setup:
- Copy `keys.local.example.json` to `keys.local.json`
- Replace `SHEET_API_URL` with your deployed Apps Script URL
- Run the EXE

After updating `keys.local.json`, just run the EXE again (no rebuild required).

EXE runtime config priority:
1) `keys.local.json` (next to EXE, then current working directory)
2) external `keys.py` (next to EXE or current working directory, only `SHEET_API_URL` and `CANVAS_BASE_URL` constants)
3) environment variables
4) built-in defaults from the EXE build

At startup, the GUI log shows `Config source:` and `Sheet API URL:` so you can verify exactly which value is being used.

## Repository Cleanup Notes

Generated/runtime files are ignored by git:
- `outputs/`
- `build/`, `dist/`
- `__pycache__/`, `*.pyc`
- `.venv/`

