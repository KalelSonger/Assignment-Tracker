# Run command:
# & ".\.venv\Scripts\python.exe" ".\PullFromCanvas.py"
import os
import re
import json
import urllib.parse
import urllib.request
from datetime import datetime
from playwright.sync_api import sync_playwright
from keys import CANVAS_BASE_URL, SHEET_API_URL


LOGIN_URL = f"{CANVAS_BASE_URL}/login/saml"
OUTPUT_DIR = "outputs"
SHEET_CLASSES_DEBUG_FILE = os.path.join(OUTPUT_DIR, "sheet_classes_debug.txt")
SHEET_SYNC_RESPONSE_FILE = os.path.join(OUTPUT_DIR, "sheet_sync_response.json")
TABS_ACTION_PARAM = "action"
TABS_ACTION_VALUE = "tabs"
SYNC_ACTION_VALUE = "sync_assignments"
CLEAR_ALL_TABS_ACTION_VALUE = "clear_all_class_tabs"
CLEAR_ONE_TAB_ACTION_VALUE = "clear_class_tab"
EXCLUDED_TAB_NAMES = {"dashboard", "class[template]"}

SYNC_MODES = {
	"1": {"name": "Sync all assignments", "include_past": True, "dry_run": False, "replace_existing": False},
	"2": {"name": "Sync future assignments", "include_past": False, "dry_run": False, "replace_existing": False},
	"3": {"name": "Dry-sync", "include_past": True, "dry_run": True, "replace_existing": False},
	"4": {"name": "Exit", "include_past": False, "dry_run": False, "exit": True},
}


def _extract_next_link(link_header: str | None) -> str | None:
	if not link_header:
		return None

	for part in link_header.split(","):
		sections = [section.strip() for section in part.split(";")]
		if len(sections) < 2:
			continue
		link_part = sections[0]
		rel_part = sections[1]
		if rel_part == 'rel="next"' and link_part.startswith("<") and link_part.endswith(">"):
			return link_part[1:-1]

	return None


def _fetch_all_pages(api_context, url: str) -> list[dict]:
	all_items: list[dict] = []
	next_url = url

	while next_url:
		response = api_context.get(next_url)
		if not response.ok:
			raise RuntimeError(f"Canvas API request failed: {response.status} {response.status_text}")

		items = response.json()
		if isinstance(items, list):
			all_items.extend(items)

		link_header = response.headers.get("link") or response.headers.get("Link")
		next_url = _extract_next_link(link_header)

	return all_items


def _is_canvas_authenticated(api_context) -> bool:
	response = api_context.get(f"{CANVAS_BASE_URL}/api/v1/users/self")
	return response.ok


def _wait_for_login(context, page, timeout_seconds: int = 300, poll_interval_ms: int = 1500) -> None:
	if _is_canvas_authenticated(context.request):
		print("Canvas session already authenticated.")
		return

	print("Opening UMSYSTEM Canvas login...")
	page.goto(LOGIN_URL, wait_until="domcontentloaded")	
	print("Complete Microsoft sign-in in the browser window. Waiting for automatic login detection...")

	started = datetime.now()
	while (datetime.now() - started).total_seconds() < timeout_seconds:
		if _is_canvas_authenticated(context.request):
			print("Login detected. Continuing...")
			return
		page.wait_for_timeout(poll_interval_ms)

	raise RuntimeError(
		"Timed out waiting for Canvas authentication. "
		"Confirm sign-in completed in the browser, then run again."
	)


def _sanitize_filename(value: str) -> str:
	cleaned = re.sub(r'[<>:"/\\|?*]', "_", value).strip().strip(".")
	return cleaned or "Unknown_Class"


def _parse_course_id(raw_course_id) -> int | None:
	if isinstance(raw_course_id, int):
		return raw_course_id
	if isinstance(raw_course_id, str) and raw_course_id.isdigit():
		return int(raw_course_id)
	return None


def _parse_canvas_datetime(raw_value: str | None) -> datetime | None:
	if not isinstance(raw_value, str) or not raw_value.strip():
		return None

	try:
		return datetime.fromisoformat(raw_value.replace("Z", "+00:00"))
	except ValueError:
		return None


def _is_current_canvas_course(course: dict) -> bool:
	workflow_state = str(course.get("workflow_state") or "").strip().lower()
	if workflow_state and workflow_state != "available":
		return False

	if course.get("access_restricted_by_date") is True:
		return False

	end_at = _parse_canvas_datetime(course.get("end_at"))
	if end_at is not None and end_at.astimezone() < datetime.now().astimezone():
		return False

	return True


def _normalize_name(value: str) -> str:
	return re.sub(r"\s+", " ", value).strip().casefold()


def _compact_name(value: str) -> str:
	return re.sub(r"\s+", "", value).casefold()


def _alnum_space(value: str) -> str:
	return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()


def _alnum_compact(value: str) -> str:
	return re.sub(r"[^a-z0-9]+", "", value.casefold())


def _build_sheet_class_patterns(tab_names: list[str]) -> list[dict]:
	patterns: list[dict] = []

	for tab_name in tab_names:
		left, _, right = tab_name.partition(" - ")
		code_source = left if right else tab_name
		title_source = right if right else tab_name

		code_compact = _alnum_compact(code_source)
		title_compact = _alnum_compact(title_source)
		title_tokens = [token for token in _alnum_space(title_source).split() if len(token) > 2]

		number_match = re.search(r"\b(\d{4})\b", _alnum_space(code_source))
		course_number = number_match.group(1) if number_match else ""

		patterns.append(
			{
				"tab_name": tab_name,
				"code_compact": code_compact,
				"title_compact": title_compact,
				"title_tokens": title_tokens,
				"course_number": course_number,
			}
		)

	return patterns


def _match_canvas_course_to_sheet_tab(canvas_course_name: str, patterns: list[dict]) -> str | None:
	canvas_compact = _alnum_compact(canvas_course_name)
	canvas_tokens = set(_alnum_space(canvas_course_name).split())

	best_tab = None
	best_score = 0

	for pattern in patterns:
		score = 0

		code_compact = pattern["code_compact"]
		if code_compact and code_compact in canvas_compact:
			score += 4

		title_compact = pattern["title_compact"]
		if title_compact and len(title_compact) >= 8 and title_compact in canvas_compact:
			score += 3

		course_number = pattern["course_number"]
		if course_number and course_number in canvas_compact:
			score += 1

		title_tokens = pattern["title_tokens"]
		if title_tokens:
			overlap_count = sum(1 for token in title_tokens if token in canvas_tokens)
			if overlap_count >= 2:
				score += 2
			elif overlap_count == 1:
				score += 1

		if score > best_score:
			best_score = score
			best_tab = pattern["tab_name"]

	return best_tab if best_score >= 3 else None


def _extract_tab_names(payload) -> list[str]:
	if isinstance(payload, list):
		return [item for item in payload if isinstance(item, str)]

	if isinstance(payload, dict):
		for key in ("tabs", "sheets", "classes", "data"):
			items = payload.get(key)
			if isinstance(items, list):
				return [item for item in items if isinstance(item, str)]

	return []


def _write_sheet_classes_debug(raw_response: str, tab_names: list[str], filtered_tabs: list[str], allowed: set[str]) -> None:
	os.makedirs(OUTPUT_DIR, exist_ok=True)
	with open(SHEET_CLASSES_DEBUG_FILE, "w", encoding="utf-8") as file:
		file.write("Sheet class debug output\n")
		file.write("========================\n\n")
		file.write("Raw API response:\n")
		file.write(raw_response)
		file.write("\n\n")

		file.write("Tabs parsed from response:\n")
		if tab_names:
			for name in tab_names:
				file.write(f"- {name}\n")
		else:
			file.write("(none)\n")
		file.write("\n")

		file.write("Tabs after excluding dashboard/class[template]:\n")
		if filtered_tabs:
			for name in filtered_tabs:
				file.write(f"- {name}\n")
		else:
			file.write("(none)\n")
		file.write("\n")

		file.write("Normalized names used for matching:\n")
		if allowed:
			for name in sorted(allowed):
				file.write(f"- {name}\n")
		else:
			file.write("(none)\n")


def fetch_allowed_sheet_classes() -> list[str]:
	payload = {TABS_ACTION_PARAM: TABS_ACTION_VALUE}
	encoded_payload = urllib.parse.urlencode(payload).encode("utf-8")
	request = urllib.request.Request(SHEET_API_URL, data=encoded_payload, method="POST")
	request.add_header("Content-Type", "application/x-www-form-urlencoded")

	with urllib.request.urlopen(request, timeout=30) as response:
		raw = response.read().decode("utf-8")

	parsed = json.loads(raw)
	tab_names = _extract_tab_names(parsed)
	excluded_compact = {_compact_name(name) for name in EXCLUDED_TAB_NAMES}
	filtered_tabs = [
		tab_name
		for tab_name in tab_names
		if _compact_name(tab_name) not in excluded_compact
	]

	allowed = {_normalize_name(tab_name) for tab_name in filtered_tabs if tab_name.strip()}
	_write_sheet_classes_debug(raw, tab_names, filtered_tabs, allowed)
	print(f"Wrote sheet class debug output to {SHEET_CLASSES_DEBUG_FILE}")
	if not filtered_tabs:
		raise RuntimeError(
			"No class tabs were returned from Google Sheet API. "
			"Expected a JSON list of tab names (or object with tabs/sheets/classes/data)."
		)

	print(f"Loaded {len(filtered_tabs)} class tab(s) from Google Sheet for filtering.")
	return filtered_tabs


def fetch_assignments_from_canvas_context(
	context,
	sheet_patterns: list[dict],
	include_past_assignments: bool = False,
) -> dict[str, list[dict]]:
	courses_url = (
		f"{CANVAS_BASE_URL}/api/v1/courses"
		"?per_page=100&enrollment_state=active&state[]=available"
	)

	print("Fetching courses...")
	courses = _fetch_all_pages(context.request, courses_url)
	current_courses = [course for course in courses if isinstance(course, dict) and _is_current_canvas_course(course)]
	print(f"Found {len(courses)} Canvas course entries total.")
	print(f"Retained {len(current_courses)} current/active courses after filtering.")

	matched_courses: list[tuple[int, str]] = []
	for course in current_courses:
		course_id = _parse_course_id(course.get("id"))
		course_name = course.get("name")
		if course_id is None or not isinstance(course_name, str):
			continue

		matched_tab = _match_canvas_course_to_sheet_tab(course_name, sheet_patterns)
		if matched_tab:
			matched_courses.append((course_id, matched_tab))

	print(f"Matched {len(matched_courses)} Canvas courses to sheet tabs. Fetching assignments only for matched courses...")

	assignments_by_course_id: dict[int, list[dict]] = {}
	course_id_to_sheet_tab: dict[int, str] = {}
	for course_id, matched_tab in matched_courses:
		course_assignments_url = (
			f"{CANVAS_BASE_URL}/api/v1/courses/{course_id}/assignments"
			"?per_page=100&order_by=due_at"
		)
		try:
			assignments_by_course_id[course_id] = _fetch_all_pages(context.request, course_assignments_url)
			course_id_to_sheet_tab[course_id] = matched_tab
		except RuntimeError as error:
			print(f"Skipping course {course_id}: {error}")
			continue

	print(f"Finished fetching assignments for {len(assignments_by_course_id)} courses.")

	today_local = datetime.now().date()
	course_names_by_id: dict[int, str] = {}
	for course in current_courses:
		course_id = _parse_course_id(course.get("id"))
		course_name = course.get("name")
		if course_id is not None and isinstance(course_name, str) and course_name:
			course_names_by_id[course_id] = course_name

	output_by_class: dict[str, list[dict]] = {}

	for course_id, assignments in assignments_by_course_id.items():
		class_name = course_id_to_sheet_tab.get(course_id, course_names_by_id.get(course_id, f"Course {course_id}"))

		for assignment in assignments:
			due_at = assignment.get("due_at")
			if not due_at:
				continue

			normalized = due_at.replace("Z", "+00:00")
			due_dt = datetime.fromisoformat(normalized)
			due_local_date = due_dt.astimezone().date()
			if not include_past_assignments and due_local_date < today_local:
				continue

			assignment_name = assignment.get("name") or ""
			record = {
				"assignment name": assignment_name,
				"due-date": due_local_date.strftime("%m/%d/%Y"),
				"Class": class_name,
			}

			output_by_class.setdefault(class_name, []).append(record)


	for class_name, records in output_by_class.items():
		records.sort(
			key=lambda item: datetime.strptime(item["due-date"], "%m/%d/%Y") if item["due-date"] else datetime.max
		)

	return output_by_class


def save_output(data: list[dict], output_path: str) -> None:
	with open(output_path, "w", encoding="utf-8") as file:
		json.dump(data, file, indent=2, ensure_ascii=False)


def write_outputs_by_class(data_by_class: dict[str, list[dict]], output_dir: str) -> int:
	os.makedirs(output_dir, exist_ok=True)
	file_count = 0

	for class_name, records in data_by_class.items():
		safe_name = _sanitize_filename(class_name)
		file_path = os.path.join(output_dir, f"{safe_name}.json")
		save_output(records, file_path)
		file_count += 1

	return file_count


def _save_sync_response(payload: dict) -> None:
	os.makedirs(OUTPUT_DIR, exist_ok=True)
	with open(SHEET_SYNC_RESPONSE_FILE, "w", encoding="utf-8") as file:
		json.dump(payload, file, indent=2, ensure_ascii=False)


def _post_sheet_action(action: str, extra_payload: dict | None = None, timeout: int = 60) -> dict:
	payload = {TABS_ACTION_PARAM: action}
	if extra_payload:
		payload.update(extra_payload)

	encoded_payload = urllib.parse.urlencode(payload).encode("utf-8")
	request = urllib.request.Request(SHEET_API_URL, data=encoded_payload, method="POST")
	request.add_header("Content-Type", "application/x-www-form-urlencoded")

	with urllib.request.urlopen(request, timeout=timeout) as response:
		raw = response.read().decode("utf-8")

	try:
		parsed = json.loads(raw)
	except json.JSONDecodeError:
		parsed = {"status": "unknown", "raw_response": raw}

	if not isinstance(parsed, dict):
		raise RuntimeError("Sheet API returned an invalid response format.")

	if parsed.get("status") not in {"success", "ok"}:
		raise RuntimeError(parsed.get("message") or f"Sheet API action '{action}' failed.")

	return parsed


def clear_all_class_tabs() -> dict:
	response = _post_sheet_action(CLEAR_ALL_TABS_ACTION_VALUE)
	return response


def clear_single_class_tab(class_name: str) -> dict:
	name = str(class_name or "").strip()
	if not name:
		raise RuntimeError("class_name is required to clear a single tab.")
	response = _post_sheet_action(CLEAR_ONE_TAB_ACTION_VALUE, {"className": name})
	return response


def sync_assignments_to_sheet(
	data_by_class: dict[str, list[dict]],
	dry_run: bool = False,
	replace_existing: bool = False,
) -> dict:
	flat_records: list[dict] = []
	for class_name, records in data_by_class.items():
		for record in records:
			flat_records.append(
				{
					"assignmentName": record.get("assignment name", ""),
					"dueDate": record.get("due-date", ""),
					"className": class_name,
				}
			)

	payload = {
		TABS_ACTION_PARAM: SYNC_ACTION_VALUE,
		"records": json.dumps(flat_records),
		"dryRun": "true" if dry_run else "false",
		"replaceExisting": "true" if replace_existing else "false",
	}
	mode = "DRY RUN" if dry_run else "LIVE"
	print(f"Sync mode: {mode}")
	print(f"Replace existing rows: {'yes' if replace_existing else 'no'}")
	print(f"Syncing {len(flat_records)} assignment rows to Google Sheet API...")
	print(f"Using endpoint: {SHEET_API_URL}")
	encoded_payload = urllib.parse.urlencode(payload).encode("utf-8")
	request = urllib.request.Request(SHEET_API_URL, data=encoded_payload, method="POST")
	request.add_header("Content-Type", "application/x-www-form-urlencoded")

	with urllib.request.urlopen(request, timeout=60) as response:
		raw = response.read().decode("utf-8")

	try:
		parsed = json.loads(raw)
	except json.JSONDecodeError:
		parsed = {"status": "unknown", "raw_response": raw}

	if not isinstance(parsed, dict) or "rowsWritten" not in parsed:
		_save_sync_response(parsed if isinstance(parsed, dict) else {"raw_response": raw})
		raise RuntimeError(
			"Sheet API did not return sync details (expected key 'rowsWritten'). "
			"This usually means the deployed Apps Script URL does not include the sync_assignments handler yet. "
			f"Check deployment for: {SHEET_API_URL}"
		)

	if replace_existing and "replaceExisting" not in parsed:
		_save_sync_response(parsed)
		raise RuntimeError(
			"Sheet API response is missing 'replaceExisting'. "
			"Your deployed Apps Script is older than the full-refresh version. "
			"Deploy a new Apps Script version with the latest GoogleSheetSync.gs and run again."
		)

	_save_sync_response(parsed)
	return parsed


def choose_sync_mode() -> dict:
	print("\nSelect sync mode:")
	print("1) Sync all assignments (past + future)")
	print("2) Sync future assignments (today onward)")
	print("3) Dry-sync (all assignments, no sheet writes)")
	print("4) Exit")

	while True:
		choice = input("Enter 1, 2, 3, or 4: ").strip()
		mode = SYNC_MODES.get(choice)
		if mode:
			print(f"Selected: {mode['name']}")
			return mode
		print("Invalid choice. Please enter 1, 2, 3, or 4.")


def main() -> None:
	try:
		allowed_sheet_tabs = fetch_allowed_sheet_classes()
		sheet_patterns = _build_sheet_class_patterns(allowed_sheet_tabs)

		with sync_playwright() as playwright:
			browser = playwright.chromium.launch(headless=False)
			context = browser.new_context()
			page = context.new_page()

			_wait_for_login(context, page)

			while True:
				mode = choose_sync_mode()
				if mode.get("exit"):
					print("Ending session.")
					break

				include_past_assignments = mode["include_past"]
				dry_run = mode["dry_run"]
				replace_existing = mode.get("replace_existing", False)

				assignments_by_class = fetch_assignments_from_canvas_context(
					context,
					sheet_patterns,
					include_past_assignments=include_past_assignments,
				)
				file_count = write_outputs_by_class(assignments_by_class, OUTPUT_DIR)
				total_assignments = sum(len(records) for records in assignments_by_class.values())
				print(f"Saved {total_assignments} assignments into {file_count} file(s) in '{OUTPUT_DIR}'.")

				sync_response = sync_assignments_to_sheet(
					assignments_by_class,
					dry_run=dry_run,
					replace_existing=replace_existing,
				)
				print(f"Sheet sync response saved to {SHEET_SYNC_RESPONSE_FILE}")
				print(f"Sheet sync status: {sync_response.get('status', 'unknown')}")
				print(f"Rows written: {sync_response.get('rowsWritten', 0)}")
				for class_name, stats in sync_response.get("classStats", {}).items():
					print(
						f"[{class_name}] incoming={stats.get('incomingCount', 0)} "
						f"existing={stats.get('existingNamedCount', 0)} matched={stats.get('matchedCount', 0)} "
						f"added={stats.get('addedCount', 0)} updated={stats.get('updatedCount', 0)}"
					)
				if sync_response.get("dryRun"):
					print("Dry run mode: no spreadsheet changes were made.")
				for message in sync_response.get("debugMessages", []):
					print(message)

			browser.close()
	except json.JSONDecodeError:
		print("Canvas returned invalid JSON.")
	except Exception as error:
		print(f"Unexpected error: {error}")


if __name__ == "__main__":
	main()
