"""
Fusion 360 DXF export helpers for SmartCutList.

Exports sheet metal flat patterns, extrusion cross-sections, flat-bar faces,
and optional milled-block orthographic views based on the part classification.
"""

from __future__ import annotations

import logging
import os
import re
import traceback
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

import adsk.core
import adsk.fusion

from .classifier import PartType
from . import settings as settings_module

# Legacy type string → current PartType constant (for any saved data using old names)
_LEGACY_TYPE_MAP = {
    'RectangularTube':    PartType.HOLLOW_RECTANGULAR_CHANNEL,
    'RoundTube':          PartType.HOLLOW_CIRCULAR_CYLINDER,
    'RoundBar':           PartType.SOLID_CYLINDER,
    'MilledBlock':        PartType.SOLID_BLOCK,
    'FlatBar':            PartType.SOLID_BLOCK,
    'AluminiumExtrusion': PartType.SOLID_BLOCK,
    'AngleSection':       PartType.SOLID_BLOCK,
    'CChannel':           PartType.SOLID_BLOCK,
    'Fastener':           PartType.FASTENER,
}

logger = logging.getLogger(__name__)

MM_PER_INCH = 25.4
_VECTOR_TOL = 1e-6
_PARALLEL_TOL = 0.999
_DEFAULT_FILENAME_TEMPLATE = "{name}_{material}_{dims}"
_SHEET_METAL_FILENAME_TEMPLATE = "{name}_{material}_{thickness}"
_INVALID_FILENAME_CHARS = re.compile(r"[^A-Za-z0-9._-]+")
_TOKEN_PATTERN = re.compile(r"{([^{}]+)}")

_PROFILE_TYPES = {
    PartType.HOLLOW_RECTANGULAR_CHANNEL,
    PartType.HOLLOW_CIRCULAR_CYLINDER,
    PartType.SOLID_CYLINDER,
}

_SKIPPED_TYPES = {
    PartType.THREE_D_PRINTED,
    PartType.FASTENER,
    PartType.SOURCED_COMPONENT,
    PartType.UNKNOWN,
}

_TYPE_ALIASES = {
    # Current type names
    "SHEETMETAL":                  PartType.SHEET_METAL,
    "SHEET_METAL":                 PartType.SHEET_METAL,
    "HOLLOWRECTANGULARCHANNEL":    PartType.HOLLOW_RECTANGULAR_CHANNEL,
    "HOLLOW_RECTANGULAR_CHANNEL":  PartType.HOLLOW_RECTANGULAR_CHANNEL,
    "HOLLOWCIRCULARCYLINDER":      PartType.HOLLOW_CIRCULAR_CYLINDER,
    "HOLLOW_CIRCULAR_CYLINDER":    PartType.HOLLOW_CIRCULAR_CYLINDER,
    "SOLIDCYLINDER":               PartType.SOLID_CYLINDER,
    "SOLID_CYLINDER":              PartType.SOLID_CYLINDER,
    "SOLIDBLOCK":                  PartType.SOLID_BLOCK,
    "SOLID_BLOCK":                 PartType.SOLID_BLOCK,
    "3DPRINTED":                   PartType.THREE_D_PRINTED,
    "FASTENER":                    PartType.FASTENER,
    "SOURCEDCOMPONENT":            PartType.SOURCED_COMPONENT,
    "SOURCED_COMPONENT":           PartType.SOURCED_COMPONENT,
    "UNKNOWN":                     PartType.UNKNOWN,
    # Legacy names for backward compatibility
    "RECTANGULARTUBE":             PartType.HOLLOW_RECTANGULAR_CHANNEL,
    "RECTANGULAR_TUBE":            PartType.HOLLOW_RECTANGULAR_CHANNEL,
    "ROUNDTUBE":                   PartType.HOLLOW_CIRCULAR_CYLINDER,
    "ROUND_TUBE":                  PartType.HOLLOW_CIRCULAR_CYLINDER,
    "ROUNDBAR":                    PartType.SOLID_CYLINDER,
    "ROUND_BAR":                   PartType.SOLID_CYLINDER,
    "MILLEDBLOCK":                 PartType.SOLID_BLOCK,
    "MILLED_BLOCK":                PartType.SOLID_BLOCK,
    "FLATBAR":                     PartType.SOLID_BLOCK,
    "FLAT_BAR":                    PartType.SOLID_BLOCK,
    "ALUMINIUMEXTRUSION":          PartType.SOLID_BLOCK,
    "ALUMINIUM_EXTRUSION":         PartType.SOLID_BLOCK,
    "ANGLESECTION":                PartType.SOLID_BLOCK,
    "ANGLE_SECTION":               PartType.SOLID_BLOCK,
    "CCHANNEL":                    PartType.SOLID_BLOCK,
    "C_CHANNEL":                   PartType.SOLID_BLOCK,
    "FASTENER":                    PartType.UNKNOWN,
}

Vec3 = Tuple[float, float, float]


def export_dxfs(
    classified_parts: List[dict],
    output_dir: str,
    naming_template: str,
    options: Dict[str, Any],
) -> List[str]:
    """
    Export DXF files for the supplied classified parts.

    Each part dict should contain at least:
      - classified_type
      - body/entity/brep_body OR body_name/bodies
      - bounding_box_mm (recommended for naming and annotations)

    Supported options:
      include_bend_lines: bool
      include_dimensions: bool
      sheet_metal_only: bool
      cross_section_for_profiles: bool
      milled_block_mode: "skip" or "orthographic" (default: "skip")
      unit: "mm" or "inches" for filename dimension tokens
    """

    os.makedirs(output_dir, exist_ok=True)

    exported_paths: List[str] = []
    design = _active_design()
    unit = _normalize_unit(options.get("unit", "mm"))
    milled_block_mode = str(options.get("milled_block_mode", "skip")).strip().lower()
    total_candidates = 0
    exported_parts = 0
    failed_parts = 0

    for part in classified_parts or []:
        if not _should_export(part):
            logger.info("Skipping DXF export for excluded part %s", _part_label(part))
            continue

        part_type = _normalize_part_type(part.get("classified_type"))
        if options.get("sheet_metal_only") and part_type != PartType.SHEET_METAL:
            logger.info("Skipping non-sheet-metal part %s", _part_label(part))
            continue
        if part_type in _SKIPPED_TYPES:
            logger.info("Skipping unsupported DXF type %s for %s", part_type, _part_label(part))
            continue

        total_candidates += 1
        try:
            body = _resolve_body(part, design)
            if body is None:
                raise RuntimeError("Unable to resolve Fusion body for part")

            base_filename = _build_filename(
                part,
                naming_template or _DEFAULT_FILENAME_TEMPLATE,
                unit=unit,
            )

            if part_type == PartType.SHEET_METAL:
                filepath = _unique_path(os.path.join(output_dir, base_filename + ".dxf"))
                target_face = resolve_sheet_metal_face(part, design, body)
                _export_sheet_metal_dxf(body, filepath, options, target_face)
                exported_paths.append(filepath)
                exported_parts += 1
                logger.info("Exported sheet metal DXF: %s", filepath)
                continue

            if part_type in _PROFILE_TYPES:
                if not options.get("cross_section_for_profiles", True):
                    logger.info("Skipping profile cross-section for %s", _part_label(part))
                    continue
                filepath = _unique_path(os.path.join(output_dir, base_filename + ".dxf"))
                _export_profile_cross_section(part, body, filepath, options)
                exported_paths.append(filepath)
                exported_parts += 1
                logger.info("Exported profile DXF: %s", filepath)
                continue

            if part_type == PartType.SOLID_BLOCK:
                if milled_block_mode != "orthographic":
                    logger.info("Skipping solid block DXF for %s", _part_label(part))
                    continue
                view_paths = _export_milled_block_views(part, body, output_dir, base_filename, options)
                exported_paths.extend(view_paths)
                exported_parts += 1
                for filepath in view_paths:
                    logger.info("Exported solid block view DXF: %s", filepath)
                continue

            logger.info("No DXF export strategy for %s (%s)", _part_label(part), part_type)

        except Exception as exc:
            logger.error(
                "DXF export failed for %s (%s): %s",
                _part_label(part),
                part.get("classified_type", "Unknown"),
                exc,
            )
            logger.debug(traceback.format_exc())
            failed_parts += 1

    if options.get("show_summary", True) and total_candidates:
        settings_module.show_operation_summary(
            "Export",
            exported_parts,
            total_candidates,
            failed_parts,
        )

    return exported_paths


def export_steps(
    classified_parts: List[dict],
    output_dir: str,
    naming_template: str,
    options: Dict[str, Any],
) -> List[str]:
    """
    Export STEP files for 3D-printed parts.

    Uses Fusion 360's exportManager.createSTEPExportOptions() to produce
    industry-standard STEP files suitable for 3D-printing services.
    """
    os.makedirs(output_dir, exist_ok=True)
    exported_paths: List[str] = []
    design = _active_design()
    unit = _normalize_unit(options.get("unit", "mm"))
    total_candidates = 0
    exported_parts = 0
    failed_parts = 0

    for part in classified_parts or []:
        if not _should_export(part):
            continue
        part_type = _normalize_part_type(part.get("classified_type"))
        if part_type != PartType.THREE_D_PRINTED:
            continue

        total_candidates += 1
        try:
            body = _resolve_body(part, design)
            if body is None:
                raise RuntimeError("Unable to resolve Fusion body for STEP export")

            base_filename = _build_filename(
                part,
                naming_template or _DEFAULT_FILENAME_TEMPLATE,
                unit=unit,
            )
            filepath = _unique_path(os.path.join(output_dir, base_filename + ".step"))

            export_mgr = design.exportManager
            step_options = export_mgr.createSTEPExportOptions(
                filepath, body.parentComponent
            )
            export_mgr.execute(step_options)
            exported_paths.append(filepath)
            exported_parts += 1
            logger.info("Exported STEP: %s", filepath)

        except Exception as exc:
            logger.error(
                "STEP export failed for %s (%s): %s",
                _part_label(part),
                part.get("classified_type", "Unknown"),
                exc,
            )
            logger.debug(traceback.format_exc())
            failed_parts += 1

    if options.get("show_summary", True) and total_candidates:
        settings_module.show_operation_summary(
            "STEP Export",
            exported_parts,
            total_candidates,
            failed_parts,
        )

    return exported_paths


def _active_design() -> adsk.fusion.Design:
    app = adsk.core.Application.get()
    design = adsk.fusion.Design.cast(app.activeProduct)
    if design is None:
        raise RuntimeError("Active Fusion product is not a design")
    return design


def resolve_body_for_part(part: Mapping[str, Any]) -> Optional[adsk.fusion.BRepBody]:
    """Public helper for review flows that need to act on a part's body."""
    return _resolve_body(part, _active_design())


def resolve_body_token(token: str) -> Optional[adsk.fusion.BRepBody]:
    """Resolve a body entity token in the active design."""
    if not token:
        return None
    return _find_entity_by_token(_active_design(), token, adsk.fusion.BRepBody)


def resolve_sheet_metal_face(
    part: Mapping[str, Any],
    design: Optional[adsk.fusion.Design] = None,
    body: Optional[adsk.fusion.BRepBody] = None,
) -> Optional[adsk.fusion.BRepFace]:
    """Resolve the configured sheet-metal base face token if one exists."""

    face_token = part.get("sheet_metal_base_face_token")
    if not face_token:
        return None

    if design is None:
        design = _active_design()
    if body is None:
        body = _resolve_body(part, design)

    face = _find_entity_by_token(design, face_token, adsk.fusion.BRepFace)
    if face is None:
        return None
    if body is None:
        return face
    return _coerce_face_for_body(face, body)


def validate_sheet_metal_configuration(
    part: Mapping[str, Any],
    face: adsk.fusion.BRepFace,
) -> None:
    """
    Validate that the chosen face can be used to create a flat pattern.

    The temporary flat pattern is removed immediately after validation.
    """

    design = _active_design()
    body = _resolve_body(part, design)
    if body is None:
        raise RuntimeError("Unable to resolve Fusion body for sheet metal configuration")
    sheet_face = _coerce_face_for_body(face, body)
    if sheet_face is None:
        raise RuntimeError("Selected face does not belong to the chosen sheet metal body")

    _validate_sheet_metal_face(body, sheet_face)


def resolve_sheet_metal_thickness_mm(
    part: Mapping[str, Any],
    face: Optional[adsk.fusion.BRepFace] = None,
) -> Optional[float]:
    """Resolve a sheet-metal thickness value in mm for naming/export."""

    thickness = part.get("sheet_metal_thickness_mm")
    try:
        if thickness is not None:
            return round(float(thickness), 3)
    except (TypeError, ValueError):
        pass

    body = resolve_body_for_part(part)
    if body is None:
        return None
    return _sheet_metal_thickness_from_body(body)


def _resolve_body(
    part: Mapping[str, Any],
    design: adsk.fusion.Design,
) -> Optional[adsk.fusion.BRepBody]:
    body_token = part.get("body_token")
    if body_token:
        body = _find_entity_by_token(design, body_token, adsk.fusion.BRepBody)
        if body is not None:
            return body

    for token in part.get("body_tokens", []) or []:
        body = _find_entity_by_token(design, token, adsk.fusion.BRepBody)
        if body is not None:
            return body

    for key in ("body", "entity", "brep_body", "body_ref"):
        candidate = part.get(key)
        body = adsk.fusion.BRepBody.cast(candidate) if candidate is not None else None
        if body is not None:
            return body

    bodies = part.get("bodies") or []
    component_name = part.get("component_name")

    for body_name in bodies:
        body = _find_body_by_name(design, body_name, component_name)
        if body is not None:
            return body

    body_name = part.get("body_name")
    if body_name:
        return _find_body_by_name(design, body_name, component_name)

    return None


def _find_body_by_name(
    design: adsk.fusion.Design,
    body_name: str,
    component_name: Optional[str] = None,
) -> Optional[adsk.fusion.BRepBody]:
    fallback = None
    for component in design.allComponents:
        for body in component.bRepBodies:
            if body.name != body_name:
                continue
            if component_name and component.name == component_name:
                return body
            if fallback is None:
                fallback = body
    return fallback


def _normalize_part_type(value: Any) -> str:
    raw = str(value or "").strip()
    compact = re.sub(r"[^A-Za-z0-9]+", "", raw).upper()
    # Check aliases first (covers current and legacy names)
    if compact in _TYPE_ALIASES:
        return _TYPE_ALIASES[compact]
    # Check legacy display names
    if raw in _LEGACY_TYPE_MAP:
        return _LEGACY_TYPE_MAP[raw]
    return raw or PartType.UNKNOWN


def _normalize_unit(value: Any) -> str:
    unit = str(value or "mm").strip().lower()
    if unit in ("inch", "in", "inches"):
        return "inches"
    return "mm"


def _should_export(part: Mapping[str, Any]) -> bool:
    if "include_in_export" in part:
        return bool(part.get("include_in_export"))
    if "include" in part:
        return bool(part.get("include"))
    return True


def _part_label(part: Mapping[str, Any]) -> str:
    return (
        part.get("component_name")
        or part.get("display_name")
        or part.get("body_name")
        or "Unknown"
    )


def _sanitize_token(value: Any) -> str:
    text = str(value or "").strip().replace(" ", "_")
    text = _INVALID_FILENAME_CHARS.sub("", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("._-") or "Unknown"


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


def _format_dim(value_mm: float, unit: str) -> str:
    value = value_mm / MM_PER_INCH if unit == "inches" else value_mm
    return "{:.1f}".format(round(value, 1))


def _build_filename(part: Mapping[str, Any], template: str, unit: str) -> str:
    dims = _dimensions_mm(part)
    part_type = _normalize_part_type(part.get("classified_type"))
    thickness_mm = resolve_sheet_metal_thickness_mm(part) if part_type == PartType.SHEET_METAL else None
    tokens = {
        "name": _sanitize_token(_part_label(part)),
        "material": _sanitize_token(part.get("material_name", "Unknown")),
        "dims": "x".join(_format_dim(value, unit) for value in dims),
        "type": _sanitize_token(part.get("classified_type", PartType.UNKNOWN)),
        "qty": str(part.get("quantity", 1)),
        "thickness": _sanitize_token(
            _format_dim(thickness_mm, unit) if thickness_mm is not None else "unknown"
        ),
    }

    def replace_token(match: re.Match[str]) -> str:
        return tokens.get(match.group(1), "")

    active_template = template or _DEFAULT_FILENAME_TEMPLATE
    if part_type == PartType.SHEET_METAL and "{thickness}" not in active_template:
        active_template = _SHEET_METAL_FILENAME_TEMPLATE
    filename = _TOKEN_PATTERN.sub(replace_token, active_template)
    return _sanitize_token(filename)


def _unique_path(filepath: str) -> str:
    if not os.path.exists(filepath):
        return filepath

    root, ext = os.path.splitext(filepath)
    index = 2
    while True:
        candidate = "{}_{}{}".format(root, index, ext)
        if not os.path.exists(candidate):
            return candidate
        index += 1


def _export_sheet_metal_dxf(
    body: adsk.fusion.BRepBody,
    filepath: str,
    options: Mapping[str, Any],
    target_face: Optional[adsk.fusion.BRepFace] = None,
) -> None:
    if target_face is None:
        target_face = _largest_planar_face(body)
    else:
        target_face = _coerce_face_for_body(target_face, body)
    if target_face is None:
        raise RuntimeError("No planar face available for flat pattern creation")

    flat_owner = _sheet_metal_owner(body.parentComponent)
    existing_flat = _existing_flat_pattern(flat_owner)
    created_flat = None
    flat_pattern = existing_flat

    try:
        if flat_pattern is None:
            if not hasattr(flat_owner, "createFlatPattern"):
                raise RuntimeError("Component does not support createFlatPattern")
            flat_pattern = flat_owner.createFlatPattern(target_face)
            if flat_pattern is None:
                flat_pattern = _existing_flat_pattern(flat_owner)
            created_flat = flat_pattern

        if flat_pattern is None:
            raise RuntimeError("Failed to create or access flat pattern")

        export_manager = _resolve_export_manager(body.parentComponent.parentDesign)
        dxf_options = export_manager.createDXFFlatPatternExportOptions(filepath, flat_pattern)
        _set_if_present(
            dxf_options,
            bool(options.get("include_bend_lines", True)),
            "isBendLinesVisible",
            "bendLinesVisible",
            "showBendLines",
        )
        result = export_manager.execute(dxf_options)
        if result is False:
            raise RuntimeError("Fusion flat-pattern DXF export returned False")

    finally:
        if created_flat is not None:
            try:
                created_flat.deleteMe()
            except Exception:
                logger.debug("Failed to delete temporary flat pattern for %s", body.name)


def _validate_sheet_metal_face(
    body: adsk.fusion.BRepBody,
    face: adsk.fusion.BRepFace,
) -> None:
    flat_owner = _sheet_metal_owner(body.parentComponent)
    existing_flat = _existing_flat_pattern(flat_owner)
    created_flat = None

    try:
        if existing_flat is None:
            if not hasattr(flat_owner, "createFlatPattern"):
                raise RuntimeError("Component does not support createFlatPattern")
            created_flat = flat_owner.createFlatPattern(face)
            if created_flat is None:
                created_flat = _existing_flat_pattern(flat_owner)
        if created_flat is None and existing_flat is None:
            raise RuntimeError("Unable to create a flat pattern from the selected face")
    finally:
        if created_flat is not None:
            try:
                created_flat.deleteMe()
            except Exception:
                logger.debug("Failed to delete validation flat pattern for %s", body.name)


def _sheet_metal_owner(component: adsk.fusion.Component) -> Any:
    try:
        owner = adsk.fusion.SheetMetalComponent.cast(component)
        if owner is not None:
            return owner
    except Exception:
        pass
    return component


def _existing_flat_pattern(owner: Any) -> Any:
    for attr in ("flatPattern", "activeFlatPattern"):
        try:
            value = getattr(owner, attr)
            if value is not None:
                return value
        except Exception:
            pass

    try:
        flat_patterns = getattr(owner, "flatPatterns")
        if flat_patterns and flat_patterns.count:
            return flat_patterns.item(0)
    except Exception:
        pass

    return None


def _find_entity_by_token(design: adsk.fusion.Design, token: str, cast_type) -> Any:
    try:
        entities = design.findEntityByToken(token)
    except Exception:
        return None

    if entities is None:
        return None

    try:
        for entity in entities:
            cast_entity = cast_type.cast(entity)
            if cast_entity is not None:
                return cast_entity
    except TypeError:
        cast_entity = cast_type.cast(entities)
        if cast_entity is not None:
            return cast_entity

    return None


def _face_belongs_to_body(face: adsk.fusion.BRepFace, body: adsk.fusion.BRepBody) -> bool:
    face = _canonical_face(face)
    body = _canonical_body(body)
    if face is None or body is None:
        return False

    try:
        return face.body.entityToken == body.entityToken
    except Exception:
        try:
            return face.body == body
        except Exception:
            return False


def _coerce_face_for_body(
    face: adsk.fusion.BRepFace,
    body: adsk.fusion.BRepBody,
) -> Optional[adsk.fusion.BRepFace]:
    candidates = []
    for candidate in (face, getattr(face, "nativeObject", None)):
        cast_face = adsk.fusion.BRepFace.cast(candidate)
        if cast_face is not None:
            candidates.append(cast_face)

    for candidate in candidates:
        if _face_belongs_to_body(candidate, body):
            return candidate
    return None


def _canonical_body(body: Optional[adsk.fusion.BRepBody]) -> Optional[adsk.fusion.BRepBody]:
    cast_body = adsk.fusion.BRepBody.cast(body)
    if cast_body is None:
        return None
    try:
        native = cast_body.nativeObject
        if native is not None:
            return native
    except Exception:
        pass
    return cast_body


def _canonical_face(face: Optional[adsk.fusion.BRepFace]) -> Optional[adsk.fusion.BRepFace]:
    cast_face = adsk.fusion.BRepFace.cast(face)
    if cast_face is None:
        return None
    try:
        native = cast_face.nativeObject
        if native is not None:
            return native
    except Exception:
        pass
    return cast_face


def _sheet_metal_thickness_from_body(body: adsk.fusion.BRepBody) -> Optional[float]:
    component = body.parentComponent

    try:
        sm_rules = component.sheetMetalRules
        if sm_rules is not None and sm_rules.count:
            for index in range(sm_rules.count):
                rule = sm_rules.item(index)
                if getattr(rule, "isActive", False):
                    thickness = getattr(rule, "thickness", None)
                    if thickness is not None:
                        return round(float(thickness.value) * 10.0, 3)
    except Exception:
        pass

    try:
        bounding = body.boundingBox
        dims_mm = [
            abs(bounding.maxPoint.x - bounding.minPoint.x) * 10.0,
            abs(bounding.maxPoint.y - bounding.minPoint.y) * 10.0,
            abs(bounding.maxPoint.z - bounding.minPoint.z) * 10.0,
        ]
        dims_mm = [value for value in dims_mm if value > 0]
        if dims_mm:
            return round(min(dims_mm), 3)
    except Exception:
        pass

    return None


def _resolve_export_manager(design: adsk.fusion.Design) -> Any:
    for candidate in (design, adsk.core.Application.get().activeProduct):
        try:
            export_manager = getattr(candidate, "exportManager", None)
            if export_manager is not None:
                return export_manager
        except Exception:
            continue
    raise RuntimeError("Unable to resolve Fusion export manager")


def _export_profile_cross_section(
    part: Mapping[str, Any],
    body: adsk.fusion.BRepBody,
    filepath: str,
    options: Mapping[str, Any],
) -> None:
    end_face = _find_end_face(body)
    if end_face is None:
        raise RuntimeError("Unable to identify an end face for profile export")
    _export_face_projection(part, body, end_face, filepath, options, project_entire_body=False)


def _export_flat_bar_face(
    part: Mapping[str, Any],
    body: adsk.fusion.BRepBody,
    filepath: str,
    options: Mapping[str, Any],
) -> None:
    face = _largest_planar_face(body)
    if face is None:
        raise RuntimeError("Unable to find a planar face for flat bar export")
    _export_face_projection(part, body, face, filepath, options, project_entire_body=False)


def _export_milled_block_views(
    part: Mapping[str, Any],
    body: adsk.fusion.BRepBody,
    output_dir: str,
    base_filename: str,
    options: Mapping[str, Any],
) -> List[str]:
    faces = _dominant_orthographic_faces(body)
    if len(faces) < 3:
        raise RuntimeError("Unable to identify three orthographic faces")

    view_names = ("top", "front", "right")
    exported_paths: List[str] = []
    for view_name, face in zip(view_names, faces):
        filepath = _unique_path(os.path.join(output_dir, "{}_{}.dxf".format(base_filename, view_name)))
        _export_face_projection(part, body, face, filepath, options, project_entire_body=True)
        exported_paths.append(filepath)

    return exported_paths


def _export_face_projection(
    part: Mapping[str, Any],
    body: adsk.fusion.BRepBody,
    face: adsk.fusion.BRepFace,
    filepath: str,
    options: Mapping[str, Any],
    project_entire_body: bool,
) -> None:
    sketch = body.parentComponent.sketches.add(face)
    try:
        if project_entire_body:
            _project_body_edges(sketch, body)
        else:
            _project_face(sketch, face)

        if options.get("include_dimensions"):
            _add_dimension_annotation(sketch, part)

        result = sketch.saveAsDXF(filepath)
        if result is False:
            raise RuntimeError("Sketch DXF export returned False")
    finally:
        try:
            sketch.deleteMe()
        except Exception:
            logger.debug("Failed to delete temporary sketch for %s", filepath)


def _project_face(sketch: adsk.fusion.Sketch, face: adsk.fusion.BRepFace) -> None:
    try:
        sketch.project(face)
    except Exception:
        pass

    curve_count = 0
    try:
        curve_count = sketch.sketchCurves.count
    except Exception:
        curve_count = 0

    if curve_count:
        return

    projected = 0
    for edge in face.edges:
        try:
            sketch.project(edge)
            projected += 1
        except Exception:
            continue

    if projected == 0:
        raise RuntimeError("Failed to project face geometry onto sketch")


def _project_body_edges(sketch: adsk.fusion.Sketch, body: adsk.fusion.BRepBody) -> None:
    projected = 0
    for edge in body.edges:
        try:
            sketch.project(edge)
            projected += 1
        except Exception:
            continue

    if projected == 0:
        raise RuntimeError("Failed to project body edges onto sketch")


def _add_dimension_annotation(sketch: adsk.fusion.Sketch, part: Mapping[str, Any]) -> None:
    text_collection = getattr(sketch, "sketchTexts", None)
    if text_collection is None:
        return

    bbox = getattr(sketch, "boundingBox", None)
    if bbox is None:
        return

    try:
        dims = _dimensions_mm(part)
        text = "{} mm".format(" x ".join("{:.1f}".format(value) for value in dims if value > 0))
        if not text.strip():
            return

        width = max(abs(bbox.maxPoint.x - bbox.minPoint.x), abs(bbox.maxPoint.y - bbox.minPoint.y))
        height = max(width * 0.05, 0.2)
        point = adsk.core.Point3D.create(bbox.minPoint.x, bbox.maxPoint.y + height * 2.0, 0)

        if hasattr(text_collection, "add"):
            try:
                text_collection.add(text, point, height)
                return
            except Exception:
                pass

        if hasattr(text_collection, "createInput2") and hasattr(text_collection, "add"):
            try:
                text_input = text_collection.createInput2(text, height)
                if hasattr(text_input, "position"):
                    text_input.position = point
                text_collection.add(text_input)
            except Exception:
                pass
    except Exception:
        logger.debug("Dimension annotation failed for %s", _part_label(part))


def _largest_planar_face(body: adsk.fusion.BRepBody) -> Optional[adsk.fusion.BRepFace]:
    best_face = None
    best_area = -1.0
    plane_type = adsk.core.Plane.classType()

    for face in body.faces:
        try:
            if face.geometry.objectType != plane_type:
                continue
            if face.area > best_area:
                best_face = face
                best_area = face.area
        except Exception:
            continue

    return best_face


def _find_end_face(body: adsk.fusion.BRepBody) -> Optional[adsk.fusion.BRepFace]:
    planar_faces = _planar_faces(body)
    if not planar_faces:
        return None

    best_face = None
    best_score = (-1.0, -1.0)

    for normal in _unique_normals(planar_faces):
        extent_min, extent_max = _body_projection_range(body, normal)
        extent = extent_max - extent_min
        if extent <= _VECTOR_TOL:
            continue

        tolerance = max(extent * 1e-4, 1e-6)
        end_faces: List[adsk.fusion.BRepFace] = []
        for face in planar_faces:
            face_normal = _face_normal(face)
            if face_normal is None or abs(_dot(face_normal, normal)) < _PARALLEL_TOL:
                continue
            face_pos = _face_projection(face, normal)
            if face_pos is None:
                continue
            if abs(face_pos - extent_min) <= tolerance or abs(face_pos - extent_max) <= tolerance:
                end_faces.append(face)

        if not end_faces:
            continue

        candidate = max(end_faces, key=lambda item: _safe_area(item))
        score = (extent, _safe_area(candidate))
        if score > best_score:
            best_score = score
            best_face = candidate

    if best_face is not None:
        return best_face

    return _largest_planar_face(body)


def _dominant_orthographic_faces(body: adsk.fusion.BRepBody) -> List[adsk.fusion.BRepFace]:
    planar_faces = _planar_faces(body)
    axis_groups: List[Tuple[Vec3, float, adsk.fusion.BRepFace]] = []

    for normal in _unique_normals(planar_faces):
        matching_faces = []
        total_area = 0.0
        for face in planar_faces:
            face_normal = _face_normal(face)
            if face_normal is None or abs(_dot(face_normal, normal)) < _PARALLEL_TOL:
                continue
            matching_faces.append(face)
            total_area += _safe_area(face)
        if matching_faces:
            representative = max(matching_faces, key=lambda item: _safe_area(item))
            axis_groups.append((normal, total_area, representative))

    axis_groups.sort(key=lambda item: item[1], reverse=True)

    selected: List[Tuple[Vec3, adsk.fusion.BRepFace]] = []
    for normal, _area, face in axis_groups:
        if any(abs(_dot(normal, existing_normal)) > 0.2 for existing_normal, _ in selected):
            continue
        selected.append((normal, face))
        if len(selected) == 3:
            break

    return [face for _normal, face in selected]


def _planar_faces(body: adsk.fusion.BRepBody) -> List[adsk.fusion.BRepFace]:
    faces = []
    plane_type = adsk.core.Plane.classType()
    for face in body.faces:
        try:
            if face.geometry.objectType == plane_type:
                faces.append(face)
        except Exception:
            continue
    return faces


def _unique_normals(faces: Sequence[adsk.fusion.BRepFace]) -> List[Vec3]:
    normals: List[Vec3] = []
    for face in faces:
        normal = _face_normal(face)
        if normal is None:
            continue
        canonical = _canonical_vector(normal)
        if any(abs(_dot(canonical, existing)) > _PARALLEL_TOL for existing in normals):
            continue
        normals.append(canonical)
    return normals


def _face_normal(face: adsk.fusion.BRepFace) -> Optional[Vec3]:
    try:
        plane = adsk.core.Plane.cast(face.geometry)
        if plane is None:
            return None
        return _normalize_vector((plane.normal.x, plane.normal.y, plane.normal.z))
    except Exception:
        return None


def _canonical_vector(vector: Vec3) -> Vec3:
    x, y, z = vector
    if x < -_VECTOR_TOL:
        return (-x, -y, -z)
    if abs(x) <= _VECTOR_TOL and y < -_VECTOR_TOL:
        return (-x, -y, -z)
    if abs(x) <= _VECTOR_TOL and abs(y) <= _VECTOR_TOL and z < -_VECTOR_TOL:
        return (-x, -y, -z)
    return vector


def _body_projection_range(body: adsk.fusion.BRepBody, axis: Vec3) -> Tuple[float, float]:
    min_value = None
    max_value = None
    for vertex in body.vertices:
        try:
            point = vertex.geometry
            value = _dot(axis, (point.x, point.y, point.z))
        except Exception:
            continue
        if min_value is None:
            min_value = value
            max_value = value
        else:
            min_value = min(min_value, value)
            max_value = max(max_value, value)

    if min_value is None or max_value is None:
        raise RuntimeError("Unable to project body vertices")
    return min_value, max_value


def _face_projection(face: adsk.fusion.BRepFace, axis: Vec3) -> Optional[float]:
    for point in (_face_centroid(face), _face_bbox_center(face), _face_plane_origin(face)):
        if point is not None:
            return _dot(axis, point)
    return None


def _face_centroid(face: adsk.fusion.BRepFace) -> Optional[Vec3]:
    try:
        props = face.evaluator.getAreaProperties()
        centroid = props.centroid
        return (centroid.x, centroid.y, centroid.z)
    except Exception:
        return None


def _face_bbox_center(face: adsk.fusion.BRepFace) -> Optional[Vec3]:
    try:
        bbox = face.boundingBox
        return (
            (bbox.minPoint.x + bbox.maxPoint.x) * 0.5,
            (bbox.minPoint.y + bbox.maxPoint.y) * 0.5,
            (bbox.minPoint.z + bbox.maxPoint.z) * 0.5,
        )
    except Exception:
        return None


def _face_plane_origin(face: adsk.fusion.BRepFace) -> Optional[Vec3]:
    try:
        plane = adsk.core.Plane.cast(face.geometry)
        if plane is None:
            return None
        return (plane.origin.x, plane.origin.y, plane.origin.z)
    except Exception:
        return None


def _safe_area(face: adsk.fusion.BRepFace) -> float:
    try:
        return float(face.area)
    except Exception:
        return 0.0


def _normalize_vector(vector: Vec3) -> Optional[Vec3]:
    mag_sq = vector[0] * vector[0] + vector[1] * vector[1] + vector[2] * vector[2]
    if mag_sq <= _VECTOR_TOL:
        return None
    mag = mag_sq ** 0.5
    return (vector[0] / mag, vector[1] / mag, vector[2] / mag)


def _dot(a: Vec3, b: Vec3) -> float:
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _set_if_present(target: Any, value: Any, *attrs: str) -> None:
    for attr in attrs:
        try:
            if hasattr(target, attr):
                setattr(target, attr, value)
                return
        except Exception:
            continue
