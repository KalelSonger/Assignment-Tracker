# Run command:
# & ".\.venv\Scripts\python.exe" ".\PullFromCanvas.py"
import os
import re
import json
import sys
from datetime import datetime
from playwright.sync_api import sync_playwright
from keys import CANVAS_BASE_URL, SHEET_API_URL


LOGIN_URL = f"{CANVAS_BASE_URL}/login/saml"
CURRENT_SHEET_URL = SHEET_API_URL
CURRENT_SPREADSHEET_ID = ""
OUTPUT_DIR = "outputs"
SHEET_CLASSES_DEBUG_FILE = os.path.join(OUTPUT_DIR, "sheet_classes_debug.txt")
SHEET_SYNC_RESPONSE_FILE = os.path.join(OUTPUT_DIR, "sheet_sync_response.json")
CANVAS_ASSIGNMENTS_DEBUG_FILE = os.path.join(OUTPUT_DIR, "canvas_assignments_debug.json")
EXCLUDED_TAB_NAMES = {"dashboard", "class[template]"}
GOOGLE_SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
GOOGLE_TOKEN_FILE = "google_sheets_token.local.json"
GOOGLE_CLIENT_SECRET_CANDIDATES = (
	"google_oauth_client_secret.json",
	"client_secret.json",
)


def _project_dir() -> str:
	return os.path.dirname(os.path.abspath(__file__))


def _state_dir() -> str:
	if getattr(sys, "frozen", False):
		return os.path.dirname(sys.executable)
	return _project_dir()


def _token_path() -> str:
	return os.path.join(_state_dir(), GOOGLE_TOKEN_FILE)


def _candidate_client_secret_paths() -> list[str]:
	paths: list[str] = []
	state_dir = _state_dir()
	parent_state_dir = os.path.dirname(state_dir)

	for candidate in GOOGLE_CLIENT_SECRET_CANDIDATES:
		paths.append(os.path.join(state_dir, candidate))
		paths.append(os.path.join(parent_state_dir, candidate))
		paths.append(os.path.join(os.getcwd(), candidate))
		paths.append(os.path.join(_project_dir(), candidate))

	# Preserve order while deduplicating.
	seen: set[str] = set()
	unique_paths: list[str] = []
	for path in paths:
		norm = os.path.normcase(os.path.abspath(path))
		if norm in seen:
			continue
		seen.add(norm)
		unique_paths.append(path)

	return unique_paths


def _client_secret_path() -> str | None:
	for path in _candidate_client_secret_paths():
		if os.path.isfile(path):
			return path
	return None


def parse_spreadsheet_id(sheet_url: str) -> str:
	cleaned = str(sheet_url or "").strip()
	if not cleaned:
		return ""

	match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", cleaned)
	if match:
		return match.group(1)

	# Accept bare spreadsheet IDs as input for power users.
	if re.fullmatch(r"[a-zA-Z0-9-_]{20,}", cleaned):
		return cleaned

	return ""


def set_sheet_api_url(api_url: str) -> None:
	global CURRENT_SHEET_URL, CURRENT_SPREADSHEET_ID
	cleaned = str(api_url or "").strip()
	if not cleaned:
		raise RuntimeError("Google Sheet URL cannot be empty.")

	spreadsheet_id = parse_spreadsheet_id(cleaned)
	if not spreadsheet_id:
		raise RuntimeError(
			"Invalid Google Sheet URL. Paste a URL like "
			"https://docs.google.com/spreadsheets/d/<ID>/edit"
		)

	CURRENT_SHEET_URL = cleaned
	CURRENT_SPREADSHEET_ID = spreadsheet_id


def get_sheet_api_url() -> str:
	return CURRENT_SHEET_URL


def _require_google_dependencies() -> None:
	try:
		import importlib
		importlib.import_module("google.auth.transport.requests")
		importlib.import_module("google.oauth2.credentials")
		importlib.import_module("google_auth_oauthlib.flow")
		importlib.import_module("googleapiclient.discovery")
	except Exception as error:
		raise RuntimeError(
			"Missing Google Sheets dependencies. Install: "
			"google-api-python-client google-auth-oauthlib"
		) from error


def _google_sheets_service():
	_require_google_dependencies()

	import importlib

	request_module = importlib.import_module("google.auth.transport.requests")
	credentials_module = importlib.import_module("google.oauth2.credentials")
	flow_module = importlib.import_module("google_auth_oauthlib.flow")
	discovery_module = importlib.import_module("googleapiclient.discovery")

	Request = getattr(request_module, "Request")
	Credentials = getattr(credentials_module, "Credentials")
	InstalledAppFlow = getattr(flow_module, "InstalledAppFlow")
	build = getattr(discovery_module, "build")

	creds = None
	token_path = _token_path()
	if os.path.isfile(token_path):
		creds = Credentials.from_authorized_user_file(token_path, GOOGLE_SHEETS_SCOPES)

	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			client_secret_path = _client_secret_path()
			if not client_secret_path:
				raise RuntimeError(
					"Google OAuth client secret file not found. Add one of: "
					+ ", ".join(GOOGLE_CLIENT_SECRET_CANDIDATES)
				)
			flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, GOOGLE_SHEETS_SCOPES)
			creds = flow.run_local_server(port=0)

	with open(token_path, "w", encoding="utf-8") as token_file:
		token_file.write(creds.to_json())

	return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _require_spreadsheet_id() -> str:
	if CURRENT_SPREADSHEET_ID:
		return CURRENT_SPREADSHEET_ID
	spreadsheet_id = parse_spreadsheet_id(CURRENT_SHEET_URL)
	if not spreadsheet_id:
		raise RuntimeError("No Google Sheet is selected. Add a valid sheet URL first.")
	return spreadsheet_id

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
	service = _google_sheets_service()
	spreadsheet_id = _require_spreadsheet_id()

	parsed = service.spreadsheets().get(
		spreadsheetId=spreadsheet_id,
		fields="properties.title,sheets.properties.title",
	).execute()
	raw = json.dumps(parsed, ensure_ascii=False, indent=2)
	tab_names = [
		str((sheet.get("properties") or {}).get("title") or "").strip()
		for sheet in parsed.get("sheets", [])
		if isinstance(sheet, dict)
	]
	tab_names = [name for name in tab_names if name]
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


def infer_sheet_display_name(sheet_url: str) -> str:
	spreadsheet_id = parse_spreadsheet_id(sheet_url)
	if not spreadsheet_id:
		raise RuntimeError("Invalid Google Sheet URL.")
	service = _google_sheets_service()
	parsed = service.spreadsheets().get(
		spreadsheetId=spreadsheet_id,
		fields="properties.title",
	).execute()
	return str((parsed.get("properties") or {}).get("title") or "").strip() or f"Sheet {spreadsheet_id[:8]}"


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
			"?per_page=100&order_by=due_at&include=all_dates"
		)
		try:
			all_assignments = _fetch_all_pages(context.request, course_assignments_url)
			assignments_by_course_id[course_id] = all_assignments
			course_id_to_sheet_tab[course_id] = matched_tab
			print(f"  Course {course_id} ({matched_tab}): fetched {len(all_assignments)} assignments")
			
			# Debug: show assignments without due_at
			missing_due_date = [a for a in all_assignments if not a.get("due_at")]
			if missing_due_date:
				print(f"    Warning: {len(missing_due_date)} assignments have no due_at date:")
				for a in missing_due_date[:5]:
					print(f"      - {a.get('name', 'Unknown')}")
				if len(missing_due_date) > 5:
					print(f"      ... and {len(missing_due_date) - 5} more")
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

	# Debug: save detailed Canvas assignments data
	debug_data = {}
	for course_id, assignments in assignments_by_course_id.items():
		class_name = course_id_to_sheet_tab.get(course_id, course_names_by_id.get(course_id, f"Course {course_id}"))
		debug_data[class_name] = [
			{
				"name": a.get("name"),
				"due_at": a.get("due_at"),
				"submission_types": a.get("submission_types"),
				"published": a.get("published"),
				"unpublishable": a.get("unpublishable"),
				"id": a.get("id"),
				"html_url": a.get("html_url"),
				"description": a.get("description", "")[:100] if a.get("description") else "",
			}
			for a in assignments
		]
	
	os.makedirs(OUTPUT_DIR, exist_ok=True)
	with open(CANVAS_ASSIGNMENTS_DEBUG_FILE, "w", encoding="utf-8") as f:
		json.dump(debug_data, f, indent=2, default=str)
	print(f"Wrote Canvas API response to {CANVAS_ASSIGNMENTS_DEBUG_FILE}")

	output_by_class: dict[str, list[dict]] = {}

	for course_id, assignments in assignments_by_course_id.items():
		class_name = course_id_to_sheet_tab.get(course_id, course_names_by_id.get(course_id, f"Course {course_id}"))
		skipped_no_due_date = 0
		skipped_past_date = 0

		for assignment in assignments:
			due_at = assignment.get("due_at")
			if not due_at:
				skipped_no_due_date += 1
				continue

			normalized = due_at.replace("Z", "+00:00")
			due_dt = datetime.fromisoformat(normalized)
			due_local_date = due_dt.astimezone().date()
			if not include_past_assignments and due_local_date < today_local:
				skipped_past_date += 1
				continue

			assignment_name = assignment.get("name") or ""
			record = {
				"assignment name": assignment_name,
				"due-date": due_local_date.strftime("%m/%d/%Y"),
				"Class": class_name,
			}

			output_by_class.setdefault(class_name, []).append(record)
		
		# Debug output
		total_for_class = len([a for a in assignments if a.get("due_at")])
		synced_count = len(output_by_class.get(class_name, []))
		print(f"  {class_name}: syncing {synced_count}/{total_for_class} assignments", end="")
		if skipped_no_due_date:
			print(f" (skipped {skipped_no_due_date} without due dates)", end="")
		if skipped_past_date:
			print(f" (skipped {skipped_past_date} past due)", end="")
		print()


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


def _quote_sheet_name(sheet_name: str) -> str:
	return "'" + sheet_name.replace("'", "''") + "'"


def _sheet_assignment_rows(service, spreadsheet_id: str, sheet_name: str) -> list[dict]:
	range_name = f"{_quote_sheet_name(sheet_name)}!A2:D"
	values = service.spreadsheets().values().get(
		spreadsheetId=spreadsheet_id,
		range=range_name,
	).execute().get("values", [])

	rows: list[dict] = []
	for idx, row in enumerate(values, start=2):
		assignment_name = str(row[0]).strip() if len(row) > 0 else ""
		due_date = str(row[1]).strip() if len(row) > 1 else ""
		class_name = str(row[3]).strip() if len(row) > 3 else ""
		rows.append(
			{
				"rowNumber": idx,
				"assignmentName": assignment_name,
				"dueDate": format_due_date(due_date),
				"dueDateKey": normalize_due_date_key(due_date),
				"className": class_name,
				"matched": False,
			}
		)
	return rows


def clear_all_class_tabs() -> dict:
	service = _google_sheets_service()
	spreadsheet_id = _require_spreadsheet_id()
	tabs = fetch_allowed_sheet_classes()

	cleared_tabs: list[dict] = []
	total_cleared_rows = 0
	for tab_name in tabs:
		rows = _sheet_assignment_rows(service, spreadsheet_id, tab_name)
		class_cleared_rows = sum(
			1 for row in rows if row["assignmentName"] or row["dueDate"] or row["className"]
		)
		if class_cleared_rows > 0:
			for col in ("A", "B", "D"):
				service.spreadsheets().values().clear(
					spreadsheetId=spreadsheet_id,
					range=f"{_quote_sheet_name(tab_name)}!{col}2:{col}",
					body={},
				).execute()

		total_cleared_rows += class_cleared_rows
		cleared_tabs.append({"sheetName": tab_name, "clearedRows": class_cleared_rows})

	return {
		"status": "success",
		"action": "clear_all_class_tabs",
		"clearedRows": total_cleared_rows,
		"clearedTabs": cleared_tabs,
	}


def clear_single_class_tab(class_name: str) -> dict:
	name = str(class_name or "").strip()
	if not name:
		raise RuntimeError("class_name is required to clear a single tab.")

	service = _google_sheets_service()
	spreadsheet_id = _require_spreadsheet_id()
	rows = _sheet_assignment_rows(service, spreadsheet_id, name)
	class_cleared_rows = sum(1 for row in rows if row["assignmentName"] or row["dueDate"] or row["className"])

	if class_cleared_rows > 0:
		for col in ("A", "B", "D"):
			service.spreadsheets().values().clear(
				spreadsheetId=spreadsheet_id,
				range=f"{_quote_sheet_name(name)}!{col}2:{col}",
				body={},
			).execute()

	return {
		"status": "success",
		"action": "clear_class_tab",
		"className": name,
		"clearedRows": class_cleared_rows,
	}


def normalize_name(value: str) -> str:
	return (
		str(value or "")
		.lower()
		.replace("-", " ")
		.replace("_", " ")
	)


def normalize_name_tokens(value: str) -> str:
	text = normalize_name(value)
	text = re.sub(r"([a-z])(\d)", r"\1 \2", text)
	text = re.sub(r"(\d)([a-z])", r"\1 \2", text)
	text = re.sub(r"[^a-z0-9]+", " ", text)
	return text.strip()


def build_assignment_key(normalized_name: str) -> str:
	tokens = normalized_name.split()
	numbers = [str(int(token)) for token in tokens if token.isdigit()]

	has_hw = "hw" in tokens or "homework" in tokens
	if has_hw and numbers:
		return f"hw:{'-'.join(numbers)}"
	if "attendance" in tokens and len(numbers) >= 2:
		return f"attendance:{numbers[0]}-{numbers[1]}"
	if "quiz" in tokens and numbers:
		return f"quiz:{numbers[0]}"
	if "exam" in tokens and numbers:
		return f"exam:{numbers[0]}"
	if "problem" in tokens and numbers:
		return f"problem:{'-'.join(numbers)}"
	return ""


def similarity_score(existing_name: str, incoming_name: str) -> int:
	a = normalize_name_tokens(existing_name)
	b = normalize_name_tokens(incoming_name)
	if not a or not b:
		return 0
	if a == b:
		return 10

	a_key = build_assignment_key(a)
	b_key = build_assignment_key(b)
	if a_key and b_key and a_key == b_key:
		return 9

	score = 0
	if len(a) >= 6 and len(b) >= 6 and (a in b or b in a):
		score += 3

	a_tokens = [token for token in a.split() if token]
	b_set = set(token for token in b.split() if token)
	overlap = 0
	numeric_overlap = 0
	for token in a_tokens:
		if token not in b_set:
			continue
		overlap += 1
		if token.isdigit():
			numeric_overlap += 1

	if overlap >= 3:
		score += 4
	elif overlap == 2:
		score += 3

	if numeric_overlap >= 1:
		score += 2

	return score


def find_best_matching_row(existing_rows: list[dict], assignment_name: str) -> dict | None:
	best_row = None
	best_score = 0
	for row in existing_rows:
		if row.get("matched"):
			continue
		score = similarity_score(str(row.get("assignmentName") or ""), assignment_name)
		if score > best_score:
			best_score = score
			best_row = row
	return best_row if best_score >= 7 else None


def parse_date_value(value: str):
	text = str(value or "").strip()
	if not text:
		return None
	mmdd = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", text)
	if mmdd:
		month = int(mmdd.group(1))
		day = int(mmdd.group(2))
		year = int(mmdd.group(3))
		return datetime(year, month, day)
	try:
		return datetime.fromisoformat(text)
	except ValueError:
		return None


def normalize_due_date_key(value: str) -> str:
	parsed = parse_date_value(value)
	if not parsed:
		return ""
	return parsed.strftime("%Y-%m-%d")


def format_due_date(value: str) -> str:
	parsed = parse_date_value(value)
	if not parsed:
		return str(value or "").strip()
	return parsed.strftime("%m/%d/%Y")


def first_empty_assignment_row(all_existing_rows: list[dict]) -> int:
	if not all_existing_rows:
		return 2
	for row in all_existing_rows:
		if not str(row.get("assignmentName") or "").strip():
			return int(row.get("rowNumber") or 2)
	return int(all_existing_rows[-1].get("rowNumber") or 1) + 1


def cache_written_assignment_row(
	all_existing_rows: list[dict],
	row_number: int,
	assignment_name: str,
	due_date: str,
	class_name: str,
) -> None:
	for row in all_existing_rows:
		if int(row.get("rowNumber") or -1) != row_number:
			continue
		row["assignmentName"] = assignment_name
		row["dueDate"] = due_date
		row["dueDateKey"] = normalize_due_date_key(due_date)
		row["className"] = class_name
		return

	all_existing_rows.append(
		{
			"rowNumber": row_number,
			"assignmentName": assignment_name,
			"dueDate": due_date,
			"dueDateKey": normalize_due_date_key(due_date),
			"className": class_name,
			"matched": False,
		}
	)


def sync_assignments_to_sheet(
	data_by_class: dict[str, list[dict]],
	dry_run: bool = False,
	replace_existing: bool = False,
) -> dict:
	service = _google_sheets_service()
	spreadsheet_id = _require_spreadsheet_id()

	flat_records: list[dict] = []
	grouped: dict[str, list[dict]] = {}
	for class_name, records in data_by_class.items():
		for record in records:
			item = {
				"assignmentName": str(record.get("assignment name") or "").strip(),
				"dueDate": str(record.get("due-date") or "").strip(),
				"className": class_name,
			}
			flat_records.append(item)
			grouped.setdefault(class_name, []).append(item)

	mode = "DRY RUN" if dry_run else "LIVE"
	print(f"Sync mode: {mode}")
	print(f"Replace existing rows: {'yes' if replace_existing else 'no'}")
	print(f"Syncing {len(flat_records)} assignment rows to Google Sheet...")
	print(f"Using sheet: {CURRENT_SHEET_URL}")

	class_stats: dict[str, dict] = {}
	debug_messages: list[str] = []
	added_rows = 0
	updated_rows = 0
	updated_classes: list[str] = []

	for class_name, class_records in grouped.items():
		class_records = [item for item in class_records if item["assignmentName"]]
		class_records.sort(key=lambda x: parse_date_value(x["dueDate"]) or datetime.max)
		class_updates: list[dict] = []

		all_existing_rows = _sheet_assignment_rows(service, spreadsheet_id, class_name)
		existing_rows = [row for row in all_existing_rows if row["assignmentName"]]

		class_added = 0
		class_updated = 0
		class_matched = 0

		if replace_existing:
			if not dry_run:
				service.spreadsheets().values().batchClear(
					spreadsheetId=spreadsheet_id,
					body={
						"ranges": [
							f"{_quote_sheet_name(class_name)}!A2:A",
							f"{_quote_sheet_name(class_name)}!B2:B",
							f"{_quote_sheet_name(class_name)}!D2:D",
						]
					},
				).execute()
				all_existing_rows = []

			for item in class_records:
				incoming_due_date = format_due_date(item["dueDate"])
				if not dry_run:
					new_row = first_empty_assignment_row(all_existing_rows)
					class_updates.append(
						{
							"range": f"{_quote_sheet_name(class_name)}!A{new_row}:D{new_row}",
							"values": [[item["assignmentName"], incoming_due_date, "", class_name]],
						}
					)
					cache_written_assignment_row(
						all_existing_rows,
						new_row,
						item["assignmentName"],
						incoming_due_date,
						class_name,
					)
				added_rows += 1
				class_added += 1

			if not dry_run and class_updates:
				service.spreadsheets().values().batchUpdate(
					spreadsheetId=spreadsheet_id,
					body={
						"valueInputOption": "USER_ENTERED",
						"data": class_updates,
					},
				).execute()

			class_stats[class_name] = {
				"incomingCount": len(class_records),
				"existingNamedCount": 0,
				"matchedCount": 0,
				"addedCount": class_added,
				"updatedCount": 0,
				"replaceMode": True,
			}
			updated_classes.append(class_name)
			continue

		for item in class_records:
			best_match = find_best_matching_row(existing_rows, item["assignmentName"])
			incoming_due_date = format_due_date(item["dueDate"])
			incoming_due_key = normalize_due_date_key(item["dueDate"])

			if best_match:
				best_match["matched"] = True
				class_matched += 1
				if best_match["dueDateKey"] != incoming_due_key:
					if not dry_run:
						class_updates.append(
							{
								"range": f"{_quote_sheet_name(class_name)}!B{best_match['rowNumber']}",
								"values": [[incoming_due_date]],
							}
						)
					debug_messages.append(
						f"assignment {item['assignmentName']} date updated from "
						f"{best_match['dueDate'] or '(blank)'} to {incoming_due_date or '(blank)'}"
					)
					best_match["dueDate"] = incoming_due_date
					best_match["dueDateKey"] = incoming_due_key
					updated_rows += 1
					class_updated += 1

				if best_match["className"] != class_name and not dry_run:
					class_updates.append(
						{
							"range": f"{_quote_sheet_name(class_name)}!D{best_match['rowNumber']}",
							"values": [[class_name]],
						}
					)
					best_match["className"] = class_name
			else:
				new_row = first_empty_assignment_row(all_existing_rows)
				if not dry_run:
					class_updates.append(
						{
							"range": f"{_quote_sheet_name(class_name)}!A{new_row}:D{new_row}",
							"values": [[item["assignmentName"], incoming_due_date, "", class_name]],
						}
					)
				added_rows += 1
				class_added += 1
				cache_written_assignment_row(
					all_existing_rows,
					new_row,
					item["assignmentName"],
					incoming_due_date,
					class_name,
				)

		if not dry_run and class_updates:
			service.spreadsheets().values().batchUpdate(
				spreadsheetId=spreadsheet_id,
				body={
					"valueInputOption": "USER_ENTERED",
					"data": class_updates,
				},
			).execute()

		class_stats[class_name] = {
			"incomingCount": len(class_records),
			"existingNamedCount": len(existing_rows),
			"matchedCount": class_matched,
			"addedCount": class_added,
			"updatedCount": class_updated,
			"replaceMode": False,
		}
		updated_classes.append(class_name)

	response = {
		"status": "success",
		"dryRun": dry_run,
		"replaceExisting": replace_existing,
		"updatedClasses": updated_classes,
		"rowsWritten": added_rows + updated_rows,
		"addedRows": added_rows,
		"updatedRows": updated_rows,
		"classStats": class_stats,
		"debugMessages": debug_messages,
		"requestedClasses": list(grouped.keys()),
	}
	_save_sync_response(response)
	return response


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
