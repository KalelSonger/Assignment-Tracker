import json
import os
import sys
import ast


DEFAULT_SHEET_API_URL = "https://script.google.com/macros/s/AKfycbzEgPUMcTkDvFHmSfmYGLQF75-F9pVmz3GD9sXA5IvXttrh1enRzRrgx7iEbXYkvScIZA/exec"
DEFAULT_CANVAS_BASE_URL = "https://umsystem.instructure.com"


def _candidate_config_paths() -> list[str]:
	paths = []

	if getattr(sys, "frozen", False):
		exe_dir = os.path.dirname(sys.executable)
		project_root = os.path.dirname(exe_dir)
		paths.append(os.path.join(exe_dir, "keys.local.json"))
		paths.append(os.path.join(project_root, "keys.local.json"))

	paths.append(os.path.join(os.getcwd(), "keys.local.json"))
	paths.append(os.path.join(os.path.dirname(__file__), "keys.local.json"))
	return paths


def _candidate_python_override_paths() -> list[str]:
	paths = []

	if getattr(sys, "frozen", False):
		exe_dir = os.path.dirname(sys.executable)
		project_root = os.path.dirname(exe_dir)
		paths.append(os.path.join(exe_dir, "keys.py"))
		paths.append(os.path.join(project_root, "keys.py"))

	paths.append(os.path.join(os.getcwd(), "keys.py"))
	return paths


def _load_external_config() -> tuple[dict, str]:
	for path in _candidate_config_paths():
		if not os.path.isfile(path):
			continue

		try:
			with open(path, "r", encoding="utf-8") as file:
				data = json.load(file)
			if isinstance(data, dict):
				return data, path
		except Exception:
			continue

	return {}, ""


def _load_python_overrides() -> tuple[dict, str]:
	this_file = os.path.abspath(__file__)

	for path in _candidate_python_override_paths():
		if not os.path.isfile(path):
			continue

		if os.path.abspath(path) == this_file:
			continue

		try:
			with open(path, "r", encoding="utf-8") as file:
				source = file.read()
			tree = ast.parse(source, filename=path)

			overrides = {}
			for node in tree.body:
				if not isinstance(node, ast.Assign):
					continue
				if len(node.targets) != 1 or not isinstance(node.targets[0], ast.Name):
					continue

				name = node.targets[0].id
				if name not in {
					"SHEET_API_URL",
					"CANVAS_BASE_URL",
					"DEFAULT_SHEET_API_URL",
					"DEFAULT_CANVAS_BASE_URL",
				}:
					continue

				if isinstance(node.value, ast.Constant) and isinstance(node.value.value, str):
					overrides[name] = node.value.value

			if "SHEET_API_URL" not in overrides and "DEFAULT_SHEET_API_URL" in overrides:
				overrides["SHEET_API_URL"] = overrides["DEFAULT_SHEET_API_URL"]
			if "CANVAS_BASE_URL" not in overrides and "DEFAULT_CANVAS_BASE_URL" in overrides:
				overrides["CANVAS_BASE_URL"] = overrides["DEFAULT_CANVAS_BASE_URL"]

			if overrides:
				return overrides, path
		except Exception:
			continue

	return {}, ""


_external_json, _external_json_path = _load_external_config()
_external_py, _external_py_path = _load_python_overrides()

_effective = {}
if _external_py:
	_effective.update(_external_py)
if _external_json:
	_effective.update(_external_json)

SHEET_API_URL = (
	_effective.get("SHEET_API_URL")
	or os.getenv("SHEET_API_URL")
	or DEFAULT_SHEET_API_URL
)

CANVAS_BASE_URL = (
	_effective.get("CANVAS_BASE_URL")
	or os.getenv("CANVAS_BASE_URL")
	or DEFAULT_CANVAS_BASE_URL
)

if _external_json_path:
	CONFIG_SOURCE = _external_json_path
elif _external_py_path:
	CONFIG_SOURCE = _external_py_path
elif os.getenv("SHEET_API_URL") or os.getenv("CANVAS_BASE_URL"):
	CONFIG_SOURCE = "environment"
else:
	CONFIG_SOURCE = "built-in defaults"
