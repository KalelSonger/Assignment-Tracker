import json
import os
import urllib.parse
import urllib.request

from keys import SHEET_API_URL

OUTPUT_DIR = "outputs"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "sheet_tabs_dump.json")


def dump_sheet_tabs(max_rows: int = 300) -> dict:
    payload = {
        "action": "dump_tabs",
        "maxRows": str(max_rows),
    }
    encoded = urllib.parse.urlencode(payload).encode("utf-8")
    request = urllib.request.Request(SHEET_API_URL, data=encoded, method="POST")
    request.add_header("Content-Type", "application/x-www-form-urlencoded")

    with urllib.request.urlopen(request, timeout=60) as response:
        raw = response.read().decode("utf-8")

    return json.loads(raw)


def save_json(data: dict, file_path: str) -> None:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(file_path, "w", encoding="utf-8") as file:
        json.dump(data, file, indent=2, ensure_ascii=False)


def main() -> None:
    try:
        data = dump_sheet_tabs(max_rows=300)
        save_json(data, OUTPUT_FILE)
        print(f"Saved sheet dump to {OUTPUT_FILE}")
        print(f"Spreadsheet: {data.get('spreadsheetName', 'unknown')} ({data.get('spreadsheetId', 'unknown')})")
    except Exception as error:
        print(f"Unexpected error: {error}")


if __name__ == "__main__":
    main()
