"""
Settings, logging, update checks, and recovery helpers for SmartCutList.

This module persists user preferences to JSON, registers a Fusion 360 command
for editing those settings, configures file-based logging, and provides
batch-recovery utilities that exporters can reuse.
"""

from __future__ import annotations

import copy
import json
import logging
import os
import threading
import traceback
import urllib.error
import urllib.request
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Mapping, Optional, Tuple

import adsk.core

ADDIN_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
)
SETTINGS_FILE = os.path.join(ADDIN_DIR, "settings.json")
LOG_FILE = os.path.join(ADDIN_DIR, "smartcutlist.log")
MANIFEST_FILE = os.path.join(ADDIN_DIR, "SmartCutList.manifest")

PANEL_ID = "SolidMakePanel"
DROPDOWN_ID = "SmartCutListDropDown"
DROPDOWN_NAME = "Smart Cut List"
SETTINGS_CMD_ID = "SmartCutListSettingsCmd"
SETTINGS_CMD_NAME = "Settings"
SETTINGS_CMD_DESCRIPTION = "Configure Smart Cut List defaults and behavior."
SELECT_CMD_ID = "SmartCutListSelectCmd"

DEFAULT_UPDATE_URL = ""

DEFAULT_SETTINGS: Dict[str, Any] = {
    "default_naming_template": "{name}_{material}_{dims}",
    "default_units": "mm",
    "include_bend_lines": True,
    "use_ai_classification": False,
    "ai_backend": "claude",
    "anthropic_api_key": "",
    "ollama_model": "llama3.2",
    "confidence_threshold": 0.6,
    "auto_group_identical": True,
    "dimension_tolerance_mm": 0.1,
    "custom_type_definitions": [],
}

_app: Optional[adsk.core.Application] = None
_ui: Optional[adsk.core.UserInterface] = None
_handlers: List[Any] = []
_cmd_handlers: List[Any] = []
_settings_cache: Optional[Dict[str, Any]] = None
_logging_ready = False

logger = logging.getLogger("SmartCutList")


def start(update_url: str = DEFAULT_UPDATE_URL, current_version: Optional[str] = None) -> None:
    """Initialize logging, register the settings command, and check for updates."""
    global _app, _ui

    _app = adsk.core.Application.get()
    _ui = _app.userInterface if _app else None

    configure_logging()
    register_settings_command()
    load_settings(force_reload=True)

    version = current_version or read_current_version()
    if update_url:
        check_for_updates_async(version, update_url)


def stop() -> None:
    """Remove the settings command and release event-handler references."""
    global _handlers, _cmd_handlers

    try:
        if _ui:
            dropdown = _find_dropdown()
            if dropdown:
                ctrl = dropdown.controls.itemById(SETTINGS_CMD_ID)
                if ctrl:
                    ctrl.deleteMe()

            cmd_def = _ui.commandDefinitions.itemById(SETTINGS_CMD_ID)
            if cmd_def:
                cmd_def.deleteMe()
    except Exception:
        logger.exception("Failed while stopping settings command")
    finally:
        _handlers = []
        _cmd_handlers = []


def configure_logging() -> logging.Logger:
    """Configure file logging once for the whole add-in."""
    global _logging_ready

    if _logging_ready:
        return logger

    os.makedirs(ADDIN_DIR, exist_ok=True)

    formatter = logging.Formatter(
        fmt="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    for handler in root_logger.handlers:
        if isinstance(handler, logging.FileHandler):
            try:
                if os.path.normcase(handler.baseFilename) == os.path.normcase(LOG_FILE):
                    _logging_ready = True
                    return logger
            except Exception:
                continue

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    logger.info("Logging initialized at %s", LOG_FILE)
    _logging_ready = True
    return logger


def log_classification_decision(
    part_name: str,
    classified_type: str,
    confidence: float,
    reason: str,
) -> None:
    """Record one classification decision without leaking sensitive settings."""
    logger.info(
        'Classification: part="%s" type="%s" confidence=%.3f reason="%s"',
        part_name,
        classified_type,
        confidence,
        reason,
    )


def log_export_action(action: str, path: str, detail: str = "") -> None:
    """Record an export action to the add-in log file."""
    if detail:
        logger.info('Export: action="%s" path="%s" detail="%s"', action, path, detail)
    else:
        logger.info('Export: action="%s" path="%s"', action, path)


def log_error(message: str, exc: Optional[BaseException] = None) -> None:
    """Record an error and optional exception traceback."""
    if exc is None:
        logger.error(message)
        return
    logger.error("%s: %s", message, exc)
    logger.debug(traceback.format_exc())


def default_settings() -> Dict[str, Any]:
    """Return a deep copy of the default settings payload."""
    return copy.deepcopy(DEFAULT_SETTINGS)


def load_settings(force_reload: bool = False) -> Dict[str, Any]:
    """Load settings from disk, merge defaults, and cache the result."""
    global _settings_cache

    if _settings_cache is not None and not force_reload:
        return copy.deepcopy(_settings_cache)

    settings = default_settings()

    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as handle:
                loaded = json.load(handle)
            if isinstance(loaded, dict):
                settings.update(loaded)
        except Exception:
            logger.exception("Failed to load settings from %s", SETTINGS_FILE)

    _settings_cache = _normalize_settings(settings)
    return copy.deepcopy(_settings_cache)


def save_settings(settings: Mapping[str, Any]) -> Dict[str, Any]:
    """Normalize and persist settings to disk."""
    global _settings_cache

    normalized = _normalize_settings(settings)
    os.makedirs(ADDIN_DIR, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as handle:
        json.dump(normalized, handle, indent=2)

    _settings_cache = normalized
    logger.info("Settings saved to %s", SETTINGS_FILE)
    return copy.deepcopy(normalized)


def get_setting(key: str, default: Any = None) -> Any:
    """Retrieve one setting value."""
    return load_settings().get(key, default)


def update_settings(values: Mapping[str, Any]) -> Dict[str, Any]:
    """Merge and save a partial settings update."""
    settings = load_settings()
    settings.update(dict(values))
    return save_settings(settings)


def read_current_version() -> str:
    """Read the add-in version from the manifest, or return 0.0.0."""
    try:
        with open(MANIFEST_FILE, "r", encoding="utf-8") as handle:
            manifest = json.load(handle)
        return str(manifest.get("version", "0.0.0"))
    except Exception:
        logger.exception("Failed to read manifest version from %s", MANIFEST_FILE)
        return "0.0.0"


def check_for_updates_async(current_version: str, update_url: str) -> None:
    """Check the remote version JSON in a background thread."""
    thread = threading.Thread(
        target=_check_for_updates_worker,
        args=(current_version, update_url),
        name="SmartCutListUpdateCheck",
        daemon=True,
    )
    thread.start()


def recoverable_batch(
    items: List[Any],
    operation: Callable[[Any], Any],
    action_name: str = "Export",
    success_predicate: Optional[Callable[[Any], bool]] = None,
    show_summary_dialog: bool = True,
) -> Tuple[List[Any], List[Tuple[Any, str]]]:
    """
    Run a batch operation item-by-item, logging failures and continuing.

    Returns:
      (successful_results, failed_items_with_error_message)
    """
    successes: List[Any] = []
    failures: List[Tuple[Any, str]] = []

    for item in items:
        try:
            result = operation(item)
            succeeded = success_predicate(result) if success_predicate else bool(result)
            if succeeded:
                successes.append(result)
            else:
                failures.append((item, "Operation returned no result"))
                logger.error('%s failed for "%s": no result', action_name, _safe_item_label(item))
        except Exception as exc:
            failures.append((item, str(exc)))
            logger.exception('%s failed for "%s"', action_name, _safe_item_label(item))

    if show_summary_dialog:
        show_operation_summary(action_name, len(successes), len(items), len(failures))

    return successes, failures


def show_operation_summary(
    action_name: str,
    successful_count: int,
    total_count: int,
    failed_count: int,
) -> str:
    """Show the end-of-batch summary requested by the exporter workflow."""
    summary = "{}ed {}/{} parts.".format(action_name.rstrip(), successful_count, total_count)
    if failed_count:
        summary += " {} parts failed (see log).".format(failed_count)

    logger.info("%s summary: %s", action_name, summary)

    if _ui:
        try:
            _ui.messageBox(summary, "SmartCutList")
        except Exception:
            logger.exception("Failed to display summary dialog")

    return summary


def register_settings_command() -> None:
    """Create the settings command definition and place it in the add-in dropdown."""
    if _ui is None:
        return

    try:
        existing = _ui.commandDefinitions.itemById(SETTINGS_CMD_ID)
        if existing:
            existing.deleteMe()

        cmd_def = _ui.commandDefinitions.addButtonDefinition(
            SETTINGS_CMD_ID,
            SETTINGS_CMD_NAME,
            SETTINGS_CMD_DESCRIPTION,
            "",
        )

        on_created = SettingsCommandCreatedHandler()
        cmd_def.commandCreated.add(on_created)
        _handlers.append(on_created)

        dropdown = _ensure_dropdown()
        if dropdown and dropdown.controls.itemById(SETTINGS_CMD_ID) is None:
            dropdown.controls.addCommand(cmd_def)
    except Exception:
        logger.exception("Failed to register SmartCutList settings command")


def _ensure_dropdown():
    panel = _find_make_panel()
    if panel is None:
        return None

    dropdown = panel.controls.itemById(DROPDOWN_ID)
    if dropdown is None:
        dropdown = panel.controls.addDropDown(DROPDOWN_NAME, "", DROPDOWN_ID)
        try:
            dropdown.isPromoted = True
        except Exception:
            pass

    _move_main_command_into_dropdown(panel, dropdown)
    return dropdown


def _move_main_command_into_dropdown(panel, dropdown) -> None:
    try:
        main_cmd_def = _ui.commandDefinitions.itemById(SELECT_CMD_ID) if _ui else None
        if main_cmd_def and dropdown.controls.itemById(SELECT_CMD_ID) is None:
            dropdown.controls.addCommand(main_cmd_def)

        standalone = panel.controls.itemById(SELECT_CMD_ID)
        if standalone:
            standalone.deleteMe()
    except Exception:
        logger.exception("Failed to move main SmartCutList command into dropdown")


def _find_dropdown():
    panel = _find_make_panel()
    if panel is None:
        return None
    return panel.controls.itemById(DROPDOWN_ID)


def _find_make_panel():
    if _ui is None:
        return None

    panel = _ui.allToolbarPanels.itemById(PANEL_ID)
    if panel is not None:
        return panel

    workspace = _ui.workspaces.itemById("FusionSolidEnvironment")
    if workspace:
        return workspace.toolbarPanels.itemById(PANEL_ID)

    return None


class SettingsCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """Build the settings dialog when the user clicks Settings."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandCreatedEventArgs) -> None:
        global _cmd_handlers

        try:
            _cmd_handlers = []
            settings = load_settings()
            inputs = args.command.commandInputs

            _add_string_input(
                inputs,
                "default_naming_template",
                "Naming Template",
                settings["default_naming_template"],
            )
            _add_dropdown(
                inputs,
                "default_units",
                "Units",
                ("mm", "inches"),
                settings["default_units"],
            )
            inputs.addBoolValueInput(
                "include_bend_lines",
                "Include Bend Lines",
                True,
                "",
                bool(settings["include_bend_lines"]),
            )
            inputs.addBoolValueInput(
                "use_ai_classification",
                "Use AI Classification",
                True,
                "",
                bool(settings["use_ai_classification"]),
            )
            _add_dropdown(
                inputs,
                "ai_backend",
                "AI Backend",
                ("claude", "ollama"),
                settings["ai_backend"],
            )

            api_key_input = _add_string_input(
                inputs,
                "anthropic_api_key",
                "Anthropic API Key",
                settings["anthropic_api_key"],
            )
            try:
                api_key_input.isPassword = True
            except Exception:
                pass

            _add_string_input(
                inputs,
                "ollama_model",
                "Ollama Model",
                settings["ollama_model"],
            )
            _add_string_input(
                inputs,
                "confidence_threshold",
                "Confidence Threshold",
                str(settings["confidence_threshold"]),
            )
            inputs.addBoolValueInput(
                "auto_group_identical",
                "Auto Group Identical",
                True,
                "",
                bool(settings["auto_group_identical"]),
            )
            _add_string_input(
                inputs,
                "dimension_tolerance_mm",
                "Dimension Tolerance (mm)",
                str(settings["dimension_tolerance_mm"]),
            )
            inputs.addTextBoxCommandInput(
                "custom_type_definitions",
                "Custom Type Definitions (JSON)",
                json.dumps(settings["custom_type_definitions"], indent=2),
                10,
                False,
            )

            on_execute = SettingsCommandExecuteHandler()
            args.command.execute.add(on_execute)
            _cmd_handlers.append(on_execute)

            on_destroy = SettingsCommandDestroyHandler()
            args.command.destroy.add(on_destroy)
            _cmd_handlers.append(on_destroy)

        except Exception:
            logger.exception("Failed to build settings dialog")
            if _ui:
                _ui.messageBox(
                    "SmartCutList settings failed to open:\n{}".format(traceback.format_exc())
                )


class SettingsCommandExecuteHandler(adsk.core.CommandEventHandler):
    """Persist settings when the settings dialog is accepted."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs) -> None:
        try:
            inputs = args.command.commandInputs
            custom_type_definitions = json.loads(
                inputs.itemById("custom_type_definitions").formattedText or "[]"
            )

            if not isinstance(custom_type_definitions, list):
                raise ValueError("Custom type definitions must be a JSON list")

            updated = {
                "default_naming_template": inputs.itemById("default_naming_template").value.strip(),
                "default_units": _selected_dropdown_value(inputs.itemById("default_units")),
                "include_bend_lines": bool(inputs.itemById("include_bend_lines").value),
                "use_ai_classification": bool(inputs.itemById("use_ai_classification").value),
                "ai_backend": _selected_dropdown_value(inputs.itemById("ai_backend")),
                "anthropic_api_key": inputs.itemById("anthropic_api_key").value,
                "ollama_model": inputs.itemById("ollama_model").value.strip(),
                "confidence_threshold": float(inputs.itemById("confidence_threshold").value),
                "auto_group_identical": bool(inputs.itemById("auto_group_identical").value),
                "dimension_tolerance_mm": float(inputs.itemById("dimension_tolerance_mm").value),
                "custom_type_definitions": custom_type_definitions,
            }

            save_settings(updated)
            _notify_text("SmartCutList settings saved.")
        except Exception as exc:
            logger.exception("Failed to save settings")
            if _ui:
                _ui.messageBox(
                    "Could not save SmartCutList settings:\n{}".format(exc),
                    "SmartCutList",
                )


class SettingsCommandDestroyHandler(adsk.core.CommandEventHandler):
    """Release per-dialog handler references after the settings command closes."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs) -> None:
        global _cmd_handlers
        _cmd_handlers = []


def _add_string_input(inputs, input_id: str, label: str, value: str):
    return inputs.addStringValueInput(input_id, label, value)


def _add_dropdown(inputs, input_id: str, label: str, items, selected: str):
    dropdown = inputs.addDropDownCommandInput(
        input_id,
        label,
        adsk.core.DropDownStyles.TextListDropDownStyle,
    )
    for item in items:
        dropdown.listItems.add(item, item == selected, "")
    return dropdown


def _selected_dropdown_value(dropdown) -> str:
    try:
        return dropdown.selectedItem.name
    except Exception:
        try:
            return dropdown.selectedItem.item
        except Exception:
            return ""


def _normalize_settings(raw_settings: Mapping[str, Any]) -> Dict[str, Any]:
    settings = default_settings()
    settings.update(dict(raw_settings))

    settings["default_naming_template"] = str(
        settings.get("default_naming_template") or DEFAULT_SETTINGS["default_naming_template"]
    )

    units = str(settings.get("default_units") or "mm").strip().lower()
    settings["default_units"] = "inches" if units in ("in", "inch", "inches") else "mm"

    settings["include_bend_lines"] = bool(settings.get("include_bend_lines", True))
    settings["use_ai_classification"] = bool(settings.get("use_ai_classification", False))

    backend = str(settings.get("ai_backend") or "claude").strip().lower()
    settings["ai_backend"] = backend if backend in ("claude", "ollama") else "claude"

    settings["anthropic_api_key"] = str(settings.get("anthropic_api_key") or "")
    settings["ollama_model"] = str(settings.get("ollama_model") or DEFAULT_SETTINGS["ollama_model"])

    try:
        threshold = float(settings.get("confidence_threshold", 0.6))
    except (TypeError, ValueError):
        threshold = 0.6
    settings["confidence_threshold"] = min(max(threshold, 0.0), 1.0)

    settings["auto_group_identical"] = bool(settings.get("auto_group_identical", True))

    try:
        tolerance = float(settings.get("dimension_tolerance_mm", 0.1))
    except (TypeError, ValueError):
        tolerance = 0.1
    settings["dimension_tolerance_mm"] = max(tolerance, 0.0)

    custom_types = settings.get("custom_type_definitions")
    settings["custom_type_definitions"] = custom_types if isinstance(custom_types, list) else []

    return settings


def _check_for_updates_worker(current_version: str, update_url: str) -> None:
    try:
        remote = _fetch_remote_version_info(update_url)
        remote_version = str(remote.get("version") or "").strip()
        if not remote_version:
            return

        if _compare_versions(remote_version, current_version) > 0:
            download_url = str(remote.get("download_url") or remote.get("url") or update_url)
            logger.info(
                "Update available: current=%s remote=%s url=%s",
                current_version,
                remote_version,
                download_url,
            )
            _notify_text(
                "SmartCutList update available: {} (current {}). {}".format(
                    remote_version,
                    current_version,
                    download_url,
                )
            )
    except Exception:
        logger.exception("Update check failed for %s", update_url)


def _fetch_remote_version_info(update_url: str) -> Dict[str, Any]:
    request = urllib.request.Request(
        update_url,
        headers={"User-Agent": "SmartCutList/{}".format(read_current_version())},
    )
    with urllib.request.urlopen(request, timeout=5) as response:
        return json.loads(response.read().decode("utf-8"))


def _compare_versions(left: str, right: str) -> int:
    left_parts = _version_parts(left)
    right_parts = _version_parts(right)

    max_len = max(len(left_parts), len(right_parts))
    left_parts.extend([0] * (max_len - len(left_parts)))
    right_parts.extend([0] * (max_len - len(right_parts)))

    if left_parts > right_parts:
        return 1
    if left_parts < right_parts:
        return -1
    return 0


def _version_parts(value: str) -> List[int]:
    parts = []
    for token in str(value or "").split("."):
        digits = "".join(ch for ch in token if ch.isdigit())
        parts.append(int(digits or 0))
    return parts or [0]


def _notify_text(message: str) -> None:
    timestamp = datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M:%S")
    full_message = "[{}] {}".format(timestamp, message)

    logger.info(message)

    if _app:
        try:
            _app.log(full_message)
        except Exception:
            pass

    if _ui:
        try:
            palette = _ui.palettes.itemById("TextCommands")
            if palette:
                palette.writeText(full_message)
        except Exception:
            logger.debug("Failed to write notification to Text Commands")


def _safe_item_label(item: Any) -> str:
    if isinstance(item, Mapping):
        for key in ("component_name", "display_name", "body_name", "name"):
            value = item.get(key)
            if value:
                return str(value)
    return str(item)
