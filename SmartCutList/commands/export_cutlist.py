"""
Pure-Python cut list export helpers for SmartCutList.

The module accepts the final reviewed/overridden part list and exports it to
CSV and JSON without depending on the Fusion 360 API.
"""

from __future__ import annotations

import csv
import json
import logging
import os
import re
from collections import namedtuple
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Iterable, List, Mapping, Optional, Tuple

MM_PER_INCH = 25.4
DEFAULT_FILENAME_TEMPLATE = "{name}_{material}_{dims}"
SHEET_METAL_FILENAME_TEMPLATE = "{name}_{material}_{thickness}"
_INVALID_FILENAME_CHARS = re.compile(r"[^A-Za-z0-9._-]+")
_TOKEN_PATTERN = re.compile(r"{([^{}]+)}")
logger = logging.getLogger(__name__)
LINEAR_STOCK_TYPES = {
    "HOLLOWRECTANGULARCHANNEL",
    "HOLLOWCIRCULARCYLINDER",
    "SOLIDCYLINDER",
}


ExportResult = namedtuple(
    "ExportResult",
    ("filepath", "parts_exported", "parts_skipped"),
)


@dataclass
class ExportSettings:
    """Settings that control filename generation and export formatting."""

    project_name: str = ""
    unit: str = "mm"
    filename_template: str = DEFAULT_FILENAME_TEMPLATE
    dim_precision: int = 1

    @classmethod
    def from_mapping(cls, value: Mapping[str, Any]) -> "ExportSettings":
        fields = cls.__dataclass_fields__
        normalized = dict(value)
        if "filename_template" not in normalized and "naming_template" in normalized:
            normalized["filename_template"] = normalized["naming_template"]
        return cls(**{key: normalized[key] for key in normalized if key in fields})


def _coerce_settings(settings: Optional[Any]) -> ExportSettings:
    if settings is None:
        return ExportSettings()
    if isinstance(settings, ExportSettings):
        return settings
    if isinstance(settings, Mapping):
        return ExportSettings.from_mapping(settings)
    raise TypeError(
        "settings must be None, a mapping, or ExportSettings; got {}".format(
            type(settings).__name__
        )
    )


def _sanitize_for_filename(value: Any) -> str:
    text = str(value or "").strip().replace(" ", "_")
    text = _INVALID_FILENAME_CHARS.sub("", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("._-") or "Unknown"


def _normalize_unit(unit: str) -> str:
    normalized = str(unit or "mm").strip().lower()
    if normalized in ("inch", "in", "inches"):
        return "inches"
    return "mm"


def _part_name(part: Mapping[str, Any]) -> str:
    return (
        part.get("display_name")
        or part.get("body_name")
        or part.get("component_name")
        or "Unknown"
    )


def _resolve_overrides(part: Mapping[str, Any]) -> Dict[str, Any]:
    resolved = dict(part)
    overrides = resolved.get("user_overrides") or {}

    name_override = (
        overrides.get("component_name")
        or overrides.get("part_name")
        or overrides.get("name")
    )
    type_override = overrides.get("classified_type") or overrides.get("type")
    material_override = overrides.get("material_name") or overrides.get("material")

    if name_override:
        resolved["component_name"] = name_override
    if type_override:
        resolved["classified_type"] = type_override
    if material_override:
        resolved["material_name"] = material_override

    return resolved


def _should_export(part: Mapping[str, Any]) -> bool:
    if "include_in_export" in part:
        return bool(part.get("include_in_export"))
    if "include" in part:
        return bool(part.get("include"))
    return True


def _quantity(part: Mapping[str, Any]) -> int:
    try:
        return max(0, int(part.get("quantity", 1)))
    except (TypeError, ValueError):
        return 0


def _dimensions_mm(part: Mapping[str, Any]) -> Tuple[float, float, float]:
    raw_dims = part.get("bounding_box_mm") or part.get("dimensions_mm") or []

    try:
        dims = [float(value) for value in raw_dims if value is not None]
    except (TypeError, ValueError):
        dims = []

    if len(dims) >= 3:
        smallest, middle, largest = sorted(dims[:3])
        return largest, middle, smallest
    if len(dims) == 2:
        smaller, larger = sorted(dims)
        return larger, smaller, 0.0
    if len(dims) == 1:
        return dims[0], 0.0, 0.0
    return 0.0, 0.0, 0.0


def _normalize_type_key(value: Any) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", str(value or "").upper())


def _profile_label(part: Mapping[str, Any], settings: ExportSettings) -> str:
    length_mm, width_mm, height_mm = _dimensions_mm(part)
    type_key = _normalize_type_key(part.get("classified_type"))

    if type_key in ("HOLLOWCIRCULARCYLINDER", "SOLIDCYLINDER"):
        return "{} dia".format(_format_dimension_string(width_mm, settings))

    if width_mm > 0 and height_mm > 0:
        return "{} x {}".format(
            _format_dimension_string(width_mm, settings),
            _format_dimension_string(height_mm, settings),
        )

    if length_mm > 0:
        return _format_dimension_string(length_mm, settings)

    return "Unknown"


def _convert_dimension(mm_value: float, settings: ExportSettings) -> float:
    unit = _normalize_unit(settings.unit)
    if unit == "inches":
        return mm_value / MM_PER_INCH
    return mm_value


def _rounded_dimension(mm_value: float, settings: ExportSettings) -> float:
    return round(_convert_dimension(mm_value, settings), settings.dim_precision)


def _format_dimension_string(mm_value: float, settings: ExportSettings) -> str:
    return "{value:.{precision}f}".format(
        value=_convert_dimension(mm_value, settings),
        precision=settings.dim_precision,
    )


def _format_dims_token(part: Mapping[str, Any], settings: ExportSettings) -> str:
    length_mm, width_mm, height_mm = _dimensions_mm(part)
    return "x".join(
        (
            _format_dimension_string(length_mm, settings),
            _format_dimension_string(width_mm, settings),
            _format_dimension_string(height_mm, settings),
        )
    )


def _sheet_metal_thickness_token(
    part: Mapping[str, Any],
    settings: ExportSettings,
) -> str:
    thickness = part.get("sheet_metal_thickness_mm")
    if thickness is None:
        thickness = part.get("estimated_wall_thickness_mm")
    try:
        return _format_dimension_string(float(thickness), settings)
    except (TypeError, ValueError):
        return "unknown"


def build_filename(
    part: Mapping[str, Any], settings: Optional[Any] = None
) -> str:
    """
    Build an export filename from a part dict and template settings.

    Supported template tokens:
    {name}, {material}, {dims}, {type}, {qty}
    """

    resolved_settings = _coerce_settings(settings)
    resolved_settings.unit = _normalize_unit(resolved_settings.unit)

    resolved_part = _resolve_overrides(part)
    tokens = {
        "name": _sanitize_for_filename(_part_name(resolved_part)),
        "material": _sanitize_for_filename(resolved_part.get("material_name", "Unknown")),
        "dims": _format_dims_token(resolved_part, resolved_settings),
        "type": _sanitize_for_filename(resolved_part.get("classified_type", "UNKNOWN")),
        "qty": str(_quantity(resolved_part)),
        "thickness": _sanitize_for_filename(
            _sheet_metal_thickness_token(resolved_part, resolved_settings)
        ),
    }

    def replace_token(match: re.Match[str]) -> str:
        token_name = match.group(1)
        return tokens.get(token_name, "")

    template = resolved_settings.filename_template
    if (
        str(resolved_part.get("classified_type") or "") == "SheetMetal"
        and "{thickness}" not in template
    ):
        template = SHEET_METAL_FILENAME_TEMPLATE

    filename = _TOKEN_PATTERN.sub(replace_token, template)
    filename = _sanitize_for_filename(filename.replace("__", "_"))
    return filename or "Unknown"


def _dimension_headers(settings: ExportSettings) -> Tuple[str, str, str]:
    unit = _normalize_unit(settings.unit)
    unit_label = "in" if unit == "inches" else "mm"
    return (
        "Length ({})".format(unit_label),
        "Width ({})".format(unit_label),
        "Height/Thickness ({})".format(unit_label),
    )


def _build_export_rows(
    parts: Iterable[Mapping[str, Any]],
    settings: ExportSettings,
) -> Tuple[List[Dict[str, Any]], int]:
    rows: List[Dict[str, Any]] = []
    skipped = 0

    for raw_part in parts:
        if not _should_export(raw_part):
            skipped += 1
            continue
        # Fasteners are exported separately in their own section
        if _normalize_type_key(raw_part.get("classified_type")) == "FASTENER":
            skipped += 1
            continue

        part = _resolve_overrides(raw_part)
        length_mm, width_mm, height_mm = _dimensions_mm(part)

        row = {
            "part_name": _part_name(part),
            "type": str(part.get("classified_type") or "UNKNOWN"),
            "material": str(part.get("material_name") or "Unknown"),
            "length": _rounded_dimension(length_mm, settings),
            "width": _rounded_dimension(width_mm, settings),
            "height": _rounded_dimension(height_mm, settings),
            "quantity": _quantity(part),
            "export_filename": build_filename(part, settings),
            "notes": "",
        }
        rows.append(row)

    for item_number, row in enumerate(rows, start=1):
        row["item_number"] = item_number

    return rows, skipped


def build_linear_stock_summary(
    parts: Iterable[Mapping[str, Any]],
    settings: Optional[Any] = None,
) -> List[Dict[str, Any]]:
    """Summarize total buy-length required for profile and stock-based parts."""

    resolved_settings = _coerce_settings(settings)
    resolved_settings.unit = _normalize_unit(resolved_settings.unit)
    summary_map: Dict[Tuple[str, str, str], Dict[str, Any]] = {}

    for raw_part in parts:
        if not _should_export(raw_part):
            continue

        part = _resolve_overrides(raw_part)
        type_value = str(part.get("classified_type") or "UNKNOWN")
        type_key = _normalize_type_key(type_value)
        if type_key not in LINEAR_STOCK_TYPES:
            continue

        length_mm, _width_mm, _height_mm = _dimensions_mm(part)
        quantity = _quantity(part)
        if length_mm <= 0 or quantity <= 0:
            continue

        material = str(part.get("material_name") or "Unknown")
        profile = _profile_label(part, resolved_settings)
        key = (type_value, material, profile)

        if key not in summary_map:
            summary_map[key] = {
                "type": type_value,
                "material": material,
                "profile": profile,
                "piece_count": 0,
                "total_length": 0.0,
            }

        entry = summary_map[key]
        entry["piece_count"] += quantity
        entry["total_length"] += _convert_dimension(length_mm * quantity, resolved_settings)

    return sorted(
        summary_map.values(),
        key=lambda item: (item["type"], item["material"], item["profile"]),
    )


def export_linear_stock_csv(
    parts: Iterable[Mapping[str, Any]],
    filepath: str,
    settings: Optional[Any] = None,
) -> str:
    """Write the linear stock summary to a standalone CSV file."""

    resolved_settings = _coerce_settings(settings)
    resolved_settings.unit = _normalize_unit(resolved_settings.unit)
    summary_rows = build_linear_stock_summary(parts, resolved_settings)
    unit_label = "in" if resolved_settings.unit == "inches" else "mm"

    os.makedirs(os.path.dirname(os.path.abspath(filepath)) or ".", exist_ok=True)
    with open(filepath, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "Type",
                "Material",
                "Profile",
                "Piece Count",
                "Total Length ({})".format(unit_label),
            ]
        )
        for row in summary_rows:
            writer.writerow(
                [
                    row["type"],
                    row["material"],
                    row["profile"],
                    row["piece_count"],
                    round(row["total_length"], resolved_settings.dim_precision),
                ]
            )

    logger.info('Linear stock summary export complete: path="%s"', os.path.abspath(filepath))
    return os.path.abspath(filepath)


def build_weld_assemblies(
    parts: Iterable[Mapping[str, Any]],
) -> List[Dict[str, Any]]:
    """Detect components with multiple bodies (weld assemblies)."""
    path_bodies: Dict[str, List[Dict[str, Any]]] = {}
    for part in parts:
        if not _should_export(part):
            continue
        path = part.get("component_path", "")
        if not path:
            continue
        path_bodies.setdefault(path, []).append(part)

    welds: List[Dict[str, Any]] = []
    for path, grps in path_bodies.items():
        all_bodies: List[str] = []
        for g in grps:
            all_bodies.extend(g.get("bodies", []))
        if len(all_bodies) > 1:
            comp_name = path.rsplit("/", 1)[-1] if "/" in path else path
            welds.append({
                "component": comp_name,
                "component_path": path,
                "body_count": len(all_bodies),
                "bodies": ", ".join(all_bodies),
            })
    welds.sort(key=lambda w: w["component"])
    return welds


def export_csv(
    parts: Iterable[Mapping[str, Any]],
    filepath: str,
    settings: Optional[Any] = None,
) -> ExportResult:
    """Export the included parts to a CSV file."""

    resolved_settings = _coerce_settings(settings)
    resolved_settings.unit = _normalize_unit(resolved_settings.unit)
    parts_list = list(parts)  # materialize so we can iterate multiple times
    rows, skipped = _build_export_rows(parts_list, resolved_settings)
    stock_rows = build_linear_stock_summary(parts_list, resolved_settings)
    weld_rows = build_weld_assemblies(parts_list)

    os.makedirs(os.path.dirname(os.path.abspath(filepath)) or ".", exist_ok=True)
    length_header, width_header, height_header = _dimension_headers(resolved_settings)
    unit_label = "in" if resolved_settings.unit == "inches" else "mm"

    with open(filepath, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "Item Number",
                "Part Name",
                "Type",
                "Material",
                length_header,
                width_header,
                height_header,
                "Quantity",
                "Export Filename",
                "Notes",
            ]
        )

        for row in rows:
            writer.writerow(
                [
                    row["item_number"],
                    row["part_name"],
                    row["type"],
                    row["material"],
                    row["length"],
                    row["width"],
                    row["height"],
                    row["quantity"],
                    row["export_filename"],
                    row["notes"],
                ]
            )

        if stock_rows:
            writer.writerow([])
            writer.writerow(["Linear Stock Summary"])
            writer.writerow(
                [
                    "Type",
                    "Material",
                    "Profile",
                    "Piece Count",
                    "Total Length ({})".format(unit_label),
                ]
            )
            for row in stock_rows:
                writer.writerow(
                    [
                        row["type"],
                        row["material"],
                        row["profile"],
                        row["piece_count"],
                        round(row["total_length"], resolved_settings.dim_precision),
                    ]
                )

        if weld_rows:
            writer.writerow([])
            writer.writerow(["Weld Assemblies"])
            writer.writerow(["Component", "Body Count", "Bodies"])
            for row in weld_rows:
                writer.writerow(
                    [row["component"], row["body_count"], row["bodies"]]
                )

        # Fastener procurement summary
        fastener_parts = [
            _resolve_overrides(p) for p in parts_list
            if _should_export(p)
            and _normalize_type_key(p.get("classified_type")) == "FASTENER"
        ]
        if fastener_parts:
            writer.writerow([])
            writer.writerow(["Fastener Procurement List"])
            writer.writerow(["Part Name", "Material", "Length ({})".format(unit_label), "Quantity"])
            for fp in fastener_parts:
                length_mm, _w, _h = _dimensions_mm(fp)
                writer.writerow([
                    _part_name(fp),
                    str(fp.get("material_name") or "Unknown"),
                    round(_convert_dimension(length_mm, resolved_settings), resolved_settings.dim_precision),
                    _quantity(fp),
                ])

    result = ExportResult(
        filepath=os.path.abspath(filepath),
        parts_exported=len(rows),
        parts_skipped=skipped,
    )
    logger.info(
        'CSV export complete: path="%s" exported=%s skipped=%s',
        result.filepath,
        result.parts_exported,
        result.parts_skipped,
    )
    return result


def export_json(
    parts: Iterable[Mapping[str, Any]],
    filepath: str,
    settings: Optional[Any] = None,
) -> ExportResult:
    """Export the included parts to a JSON file."""

    resolved_settings = _coerce_settings(settings)
    resolved_settings.unit = _normalize_unit(resolved_settings.unit)
    parts_list = list(parts)
    rows, skipped = _build_export_rows(parts_list, resolved_settings)

    types_breakdown: Dict[str, int] = {}
    total_parts = 0
    json_parts: List[Dict[str, Any]] = []

    for row in rows:
        types_breakdown[row["type"]] = types_breakdown.get(row["type"], 0) + 1
        total_parts += row["quantity"]

        json_parts.append(
            {
                "item_number": row["item_number"],
                "part_name": row["part_name"],
                "type": row["type"],
                "material": row["material"],
                "dimensions": {
                    "length": row["length"],
                    "width": row["width"],
                    "height": row["height"],
                },
                "quantity": row["quantity"],
                "export_filename": row["export_filename"],
                "notes": row["notes"],
            }
        )

    payload = {
        "project_name": resolved_settings.project_name or "SmartCutList",
        "export_date": datetime.now().isoformat(timespec="seconds"),
        "units": resolved_settings.unit,
        "parts": json_parts,
        "linear_stock_summary": build_linear_stock_summary(parts_list, resolved_settings),
        "weld_assemblies": build_weld_assemblies(parts_list),
        "summary": {
            "total_unique_parts": len(rows),
            "total_parts": total_parts,
            "types_breakdown": dict(sorted(types_breakdown.items())),
        },
    }

    os.makedirs(os.path.dirname(os.path.abspath(filepath)) or ".", exist_ok=True)
    with open(filepath, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)

    result = ExportResult(
        filepath=os.path.abspath(filepath),
        parts_exported=len(rows),
        parts_skipped=skipped,
    )
    logger.info(
        'JSON export complete: path="%s" exported=%s skipped=%s',
        result.filepath,
        result.parts_exported,
        result.parts_skipped,
    )
    return result


def export_fasteners_csv(
    parts: Iterable[Mapping[str, Any]],
    filepath: str,
    settings: Optional[Any] = None,
) -> ExportResult:
    """Export fastener-type parts to a standalone procurement CSV."""
    resolved_settings = _coerce_settings(settings)
    resolved_settings.unit = _normalize_unit(resolved_settings.unit)

    fastener_parts = [
        p for p in parts
        if _should_export(p)
        and _normalize_type_key(p.get("classified_type")) == "FASTENER"
    ]

    os.makedirs(os.path.dirname(os.path.abspath(filepath)) or ".", exist_ok=True)
    with open(filepath, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow([
            "Part Name", "Size", "Length (mm)", "Material", "Quantity",
        ])
        for part in fastener_parts:
            resolved = _resolve_overrides(part)
            length_mm, _w, _h = _dimensions_mm(resolved)
            writer.writerow([
                _part_name(resolved),
                str(resolved.get("classified_type", "Fastener")),
                round(length_mm, resolved_settings.dim_precision),
                str(resolved.get("material_name") or "Unknown"),
                _quantity(resolved),
            ])

    result = ExportResult(
        filepath=os.path.abspath(filepath),
        parts_exported=len(fastener_parts),
        parts_skipped=0,
    )
    logger.info(
        'Fastener CSV export complete: path="%s" exported=%s',
        result.filepath,
        result.parts_exported,
    )
    return result


def export_sourced_csv(
    parts: Iterable[Mapping[str, Any]],
    filepath: str,
    settings: Optional[Any] = None,
) -> ExportResult:
    """Export sourced (externally linked) components to a standalone CSV."""
    resolved_settings = _coerce_settings(settings)
    resolved_settings.unit = _normalize_unit(resolved_settings.unit)

    sourced_parts = [
        p for p in parts
        if _should_export(p)
        and _normalize_type_key(p.get("classified_type")) == "SOURCEDCOMPONENT"
    ]

    os.makedirs(os.path.dirname(os.path.abspath(filepath)) or ".", exist_ok=True)
    with open(filepath, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow([
            "Part Name", "Component", "Material",
            "Length (mm)", "Width (mm)", "Height (mm)", "Quantity",
        ])
        for part in sourced_parts:
            resolved = _resolve_overrides(part)
            length_mm, width_mm, height_mm = _dimensions_mm(resolved)
            writer.writerow([
                _part_name(resolved),
                str(resolved.get("component_path") or ""),
                str(resolved.get("material_name") or "Unknown"),
                round(length_mm, resolved_settings.dim_precision),
                round(width_mm, resolved_settings.dim_precision),
                round(height_mm, resolved_settings.dim_precision),
                _quantity(resolved),
            ])

    result = ExportResult(
        filepath=os.path.abspath(filepath),
        parts_exported=len(sourced_parts),
        parts_skipped=0,
    )
    logger.info(
        'Sourced CSV export complete: path="%s" exported=%s',
        result.filepath,
        result.parts_exported,
    )
    return result


def export_all(
    parts: Iterable[Mapping[str, Any]],
    base_path: str,
    settings: Optional[Any] = None,
) -> Dict[str, ExportResult]:
    """Export both CSV and JSON using the same base path."""

    return {
        "csv": export_csv(parts, base_path + ".csv", settings),
        "json": export_json(parts, base_path + ".json", settings),
    }
