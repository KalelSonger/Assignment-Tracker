import json
import urllib.error
import urllib.parse
import urllib.request
from keys import SHEET_API_URL


API_URL = SHEET_API_URL
OUTPUT_FILE = "write_response.json"

SUCCESS_VALUE = "SUCCESS"
TEXT_PARAM = "text"


def call_sheet_write_api(url: str, value: str):
	payload = {
		TEXT_PARAM: value,
	}

	encoded_payload = urllib.parse.urlencode(payload).encode("utf-8")
	request = urllib.request.Request(url, data=encoded_payload, method="POST")
	request.add_header("Content-Type", "application/x-www-form-urlencoded")

	with urllib.request.urlopen(request, timeout=30) as response:
		raw = response.read().decode("utf-8")
		try:
			return json.loads(raw)
		except json.JSONDecodeError:
			return {"raw_response": raw}


def save_json_file(data, file_path: str) -> None:
	with open(file_path, "w", encoding="utf-8") as file:
		json.dump(data, file, indent=2, ensure_ascii=False)


def main() -> None:
	try:
		data = call_sheet_write_api(API_URL, SUCCESS_VALUE)
		save_json_file(data, OUTPUT_FILE)
		print(f"Sent '{SUCCESS_VALUE}' to Apps Script as '{TEXT_PARAM}'.")
		print(f"Saved API response to {OUTPUT_FILE}")
	except urllib.error.URLError as error:
		print(f"Network error: {error}")
	except Exception as error:
		print(f"Unexpected error: {error}")


if __name__ == "__main__":
	main()
