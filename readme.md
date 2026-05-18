# Assignment Tracker

Assignment Tracker is a Python desktop app that pulls assignments from Canvas and syncs them directly into your Google Sheet class tabs.

## What this project does

- Pulls upcoming assignments from Canvas.
- Syncs assignments into matching class tabs in your Google Sheet.
- Supports multiple sheets from a dropdown in the GUI.
- Stores Canvas and Google auth sessions locally so you do not need to sign in every launch.

## Setup Guide (with screenshots)

### 1) Copy and prepare the spreadsheet

1. Make a copy of the template sheet:
	https://docs.google.com/spreadsheets/d/17W5u-FZ-bq8ciiSIgSRu7B255P1kheF_G30hGUedHuU/edit?usp=sharing

![Template spreadsheet](images%20for%20readme/image.png)

2. Copy the `class [TEMPLATE]` tab once for each class you are taking.

![Copy template tabs](images%20for%20readme/image-1.png)

3. Rename each copied tab to your real class name, and add those same names to the `Classes` section on the dashboard.

![Rename class tabs](images%20for%20readme/image-2.png)

> **Important:** class names must match exactly between dashboard and tab names.

## Run the app

1. Open the `dist` folder.
2. Run `AssignmentTrackerGUI.exe`.
3. Sign in to Google when prompted (same account that can edit your sheet).
4. Sign in to Canvas when prompted.
5. Paste your Google Sheet URL in the top field and click **Add**.

![Add sheet URL in GUI](images%20for%20readme/image-20.png)
 
6. Return to the GUI and run sync.

![Assignment Tracker GUI](images%20for%20readme/image-21.png)

![Spreadsheet after sync](images%20for%20readme/image-18.png)

## Notes

- `DEFAULT_CANVAS_BASE_URL` is set to `https://umsystem.instructure.com` by default.
- If your school uses a different Canvas domain, replace it with your school’s Canvas base URL.
- You can add multiple spreadsheets to the top-right dropdown and switch between them.
- OAuth token is saved locally as `google_sheets_token.local.json`.
- Canvas session is saved locally and refreshed as needed.
- If Google sign-in shows "app is being tested", ask the app owner to add your account as a test user.

## Spreadsheet color behavior

- <span style="color:#8B0000;"><strong>Dark red</strong></span>: due today or overdue.
- <span style="color:#FF6B6B;"><strong>Light red</strong></span>: due within 3 days.
- <span style="color:#D4A017;"><strong>Yellow</strong></span>: due within 1 week.
- <span style="color:#228B22;"><strong>Green</strong></span>: due later than 1 week.
- Assignment cells can also be color-coded by keywords in your trackers section (example: `Exam`).

![Tracker keyword color example](images%20for%20readme/image-19.png)

Completed assignments (checked on class tabs) move to the bottom of the dashboard list.