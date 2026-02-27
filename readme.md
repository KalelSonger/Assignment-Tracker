# Assignment Tracker

Google Sheets + Python desktop app workflow that syncs your Canvas assignments into a class-organized tracker.

## What this project does

- Pulls upcoming assignments from Canvas.
- Syncs those assignments into your Google Sheet tabs.
- Gives you a simple desktop GUI (`AssignmentTrackerGUI.exe`) to run sync actions.

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

### 2) Create and deploy the Google Apps Script web app

4. Go to https://script.google.com.

![Apps Script home](images%20for%20readme/image-3.png)

5. Click **New project**.

![New Apps Script project](images%20for%20readme/image-4.png)

6. Add the **Google Sheets API** service from the Services panel.

![Add Google Sheets API service](images%20for%20readme/image-5.png)

7. Open `GoogleSheetSync.gs` from this repo, copy it, and paste it into the `Code.gs` editor.

![Paste script code](images%20for%20readme/image-6.png)

8. Replace the spreadsheet ID in the script with your own sheet ID.

Example template URL:
https://docs.google.com/spreadsheets/d/17W5u-FZ-bq8ciiSIgSRu7B255P1kheF_G30hGUedHuU/edit?usp=sharing

In any Google Sheet URL, the sheet ID is between `/d/` and `/edit`.

![Replace sheet ID](images%20for%20readme/image-7.png)

9. Click **Deploy** (top-right), then **New deployment**.

![Open deploy menu](images%20for%20readme/image-8.png)

10. Click the gear icon next to **Select type** and choose **Web app**.

![Select web app deployment type](images%20for%20readme/image-9.png)

> **Important:** set access to **Anyone**.

11. Name the deployment and click **Deploy**.

![Deploy web app](images%20for%20readme/image-10.png)

12. Click **Authorize access**.

![Authorize access prompt](images%20for%20readme/image-11.png)

13. Click **Advanced** → **Go to (your project)**, then continue.

![Advanced authorization](images%20for%20readme/image-12.png)

![Continue authorization](images%20for%20readme/image-13.png)

You should then see the deployment details page:

![Deployment details page](images%20for%20readme/image-14.png)

### 3) Configure the desktop app

14. Copy the deployed web app URL and open `keys.py`.

![Open keys.py](images%20for%20readme/image-15.png)

15. Set `DEFAULT_SHEET_API_URL` to the deployed web app URL.

![Paste web app URL into keys.py](images%20for%20readme/image-16.png)

## Run the app

1. Open the `dist` folder.
2. Run `AssignmentTrackerGUI.exe`.
3. On first launch, dependencies may install automatically.
4. Sign in to Canvas when the browser opens.
5. Return to the GUI and run sync.

![Assignment Tracker GUI](images%20for%20readme/image-17.png)

![Spreadsheet after sync](images%20for%20readme/image-18.png)

## Notes

- `DEFAULT_CANVAS_BASE_URL` is set to `https://umsystem.instructure.com` by default.
- If your school uses a different Canvas domain, replace it with your school’s Canvas base URL.
- Other Canvas architectures are not fully tested.

## Spreadsheet color behavior

- Dark red: due today or overdue.
- Light red: due within 3 days.
- Yellow: due within 1 week.
- Green: due later than 1 week.
- Assignment cells can also be color-coded by keywords in your trackers section (example: `Exam`).

![Tracker keyword color example](images%20for%20readme/image-19.png)

Completed assignments (checked on class tabs) move to the bottom of the dashboard list.