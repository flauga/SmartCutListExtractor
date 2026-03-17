"""
hole_detection.py — Detect and analyse holes on hollow rectangular channel bodies.

Hollow rectangular channels (tube stock) may have drilled/cut holes that must
be completed *before* the channels are welded into an assembly.  This module
detects those holes from cylindrical BRepFaces, groups concentric through-holes
(e.g. counterbored holes), computes manufacturing-relevant measurements
(diameter, depth, edge distances, centre-to-centre spacing, thread hints),
and associates each hole with the sketch/feature that created it for
perpendicular photo overlays.

Public API
----------
    detect_holes_on_channels(groups) -> list[BodyHoleSummary dict]
"""

from __future__ import annotations

import logging
import math
import traceback
from typing import Any, Dict, List, Optional, Tuple

import adsk.core
import adsk.fusion

from . import export_dxf
from .classifier import PartType

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CM_TO_MM: float = 10.0
PARALLEL_TOL: float = 0.9998       # |cos θ| above → parallel
CENTER_PROXIMITY_MM: float = 0.15  # centres within this → concentric
THREAD_HINT_TOL_MM: float = 0.15   # ± tolerance for thread-size matching

# Minimum cylindrical face radius (mm) to consider as a real hole.
# Tiny radii are usually fillet/chamfer artefacts.
MIN_HOLE_RADIUS_MM: float = 0.5

# ---------------------------------------------------------------------------
# Standard metric thread lookup tables — separated by purpose.
# Tap drill: used for blind holes (likely threaded).
# Clearance: used for through-holes (likely bolt pass-through).
# ---------------------------------------------------------------------------

_TAP_DRILL_TABLE: List[Tuple[float, str]] = [
    (2.5,  '~M3 tap'),
    (3.3,  '~M4 tap'),
    (4.2,  '~M5 tap'),
    (5.0,  '~M6 tap'),
    (6.8,  '~M8 tap'),
    (8.5,  '~M10 tap'),
    (10.2, '~M12 tap'),
    (14.0, '~M16 tap'),
    (17.5, '~M20 tap'),
]

_CLEARANCE_TABLE: List[Tuple[float, str]] = [
    (3.4,  '~M3 clearance'),
    (4.5,  '~M4 clearance'),
    (5.5,  '~M5 clearance'),
    (6.6,  '~M6 clearance'),
    (9.0,  '~M8 clearance'),
    (11.0, '~M10 clearance'),
    (13.5, '~M12 clearance'),
    (17.5, '~M16 clearance'),
    (22.0, '~M20 clearance'),
]


# ---------------------------------------------------------------------------
# Vector helpers (same conventions as feature_extraction.py)
# ---------------------------------------------------------------------------

Vec3 = Tuple[float, float, float]


def _dot(a: Vec3, b: Vec3) -> float:
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross(a: Vec3, b: Vec3) -> Vec3:
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _length(v: Vec3) -> float:
    return math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2])


def _normalize(v: Vec3) -> Optional[Vec3]:
    mag = _length(v)
    if mag < 1e-12:
        return None
    return (v[0] / mag, v[1] / mag, v[2] / mag)


def _sub(a: Vec3, b: Vec3) -> Vec3:
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _scale(v: Vec3, s: float) -> Vec3:
    return (v[0] * s, v[1] * s, v[2] * s)


def _dist(a: Vec3, b: Vec3) -> float:
    return _length(_sub(a, b))


def _project_onto_axis(point: Vec3, axis_origin: Vec3, axis_dir: Vec3) -> float:
    """Project *point* onto an axis and return the signed scalar distance from *axis_origin*."""
    return _dot(_sub(point, axis_origin), axis_dir)


# ===========================================================================
# Public API
# ===========================================================================

def detect_holes_on_channels(groups: list) -> List[Dict[str, Any]]:
    """Detect holes on all HollowRectangularChannel bodies in *groups*.

    Parameters
    ----------
    groups : list
        The grouped classified body list from ``review_palette.group_classified_bodies``.

    Returns
    -------
    list of BodyHoleSummary dicts.  Bodies with zero holes are omitted.
    """
    summaries: List[Dict[str, Any]] = []
    for group in groups:
        effective_type = group.get('override_type') or group.get('classified_type', '')
        if effective_type != PartType.HOLLOW_RECTANGULAR_CHANNEL:
            continue

        for token in group.get('body_tokens', []):
            try:
                summary = _detect_holes_on_body(token, group)
                if summary and summary['total_holes'] > 0:
                    summaries.append(summary)
            except Exception:
                logger.exception('Hole detection failed for token %s', token)
    return summaries


# ===========================================================================
# Per-body hole detection
# ===========================================================================

def _detect_holes_on_body(
    token: str,
    group: dict,
) -> Optional[Dict[str, Any]]:
    """Detect holes on a single body resolved by *token*."""
    body = export_dxf.resolve_body_token(token)
    if body is None:
        return None

    cylinder_type = adsk.core.Cylinder.classType()
    holes: List[Dict[str, Any]] = []
    hole_idx = 0

    # --- Determine the channel's long axis from its bounding box -----------
    channel_axis, channel_origin, channel_length_mm, bb_min_cm, bb_max_cm = \
        _channel_axis_from_bb(body)

    # --- Collect face normals for labelling (Top/Bottom/Left/Right) --------
    face_label_axes = _build_face_label_axes(body, channel_axis)

    # --- Single pass over faces to find cylindrical (hole) faces -----------
    for face in body.faces:
        try:
            geom = face.geometry
            if geom.objectType != cylinder_type:
                continue

            radius_mm = geom.radius * CM_TO_MM
            if radius_mm < MIN_HOLE_RADIUS_MM:
                continue

            axis_raw = geom.axis
            axis = _normalize((axis_raw.x, axis_raw.y, axis_raw.z))
            if axis is None:
                continue

            origin = geom.origin
            center_cm: Vec3 = (origin.x, origin.y, origin.z)

            # Depth: axial extent between the two circular edge loops
            depth_mm = _compute_hole_depth(face, axis)

            # Face label (which wall of the channel this hole is on)
            face_label = _classify_face_direction(axis, face_label_axes)

            # Edge distances from channel ends
            dist_a, dist_b = _edge_distances(
                center_cm, channel_axis, channel_origin, channel_length_mm
            )

            # Cross-section offset — how far the hole is from the centre
            # of the face it goes through.
            cross_offset_mm, face_width_mm = _cross_section_offset(
                center_cm, axis, channel_axis, bb_min_cm, bb_max_cm
            )

            # Thread hint — assigned later based on group context (tap vs clearance)
            thread_hint = ''

            face_token = ''
            try:
                face_token = face.entityToken
            except Exception:
                pass

            holes.append({
                'hole_id': 'h_{}'.format(hole_idx),
                'axis': axis,
                'center_cm': center_cm,
                'radius_mm': round(radius_mm, 2),
                'diameter_mm': round(radius_mm * 2.0, 2),
                'depth_mm': round(depth_mm, 2) if depth_mm else None,
                'is_through': False,              # updated during grouping
                'concentric_group': -1,           # updated during grouping
                'face_tokens': [face_token],
                'face_label': face_label,
                'creating_feature_name': '',      # updated during feature scan
                'creating_sketch_token': '',      # updated during feature scan
                'thread_hint': thread_hint,
                'distance_from_end_a_mm': round(dist_a, 1) if dist_a is not None else None,
                'distance_from_end_b_mm': round(dist_b, 1) if dist_b is not None else None,
                'cross_offset_mm': round(cross_offset_mm, 1) if cross_offset_mm is not None else None,
                'face_width_mm': round(face_width_mm, 1) if face_width_mm is not None else None,
            })
            hole_idx += 1

        except Exception:
            logger.debug('Skipping face during hole detection: %s', traceback.format_exc())

    if not holes:
        return None

    # --- Group concentric holes -------------------------------------------
    concentric_groups = _group_concentric_holes(holes)

    # --- Assign context-aware thread hints after grouping ------------------
    # Blind holes → tap drill lookup; through-holes → clearance lookup.
    for cg in concentric_groups:
        is_through = cg.get('is_through', False)
        hint_type = 'clearance' if is_through else 'tap'
        for h in cg.get('holes', []):
            h['thread_hint'] = _get_thread_hint(h['diameter_mm'], hint_type)

    # --- Centre-to-centre distances between groups -------------------------
    _compute_center_distances(concentric_groups)

    # --- Feature / sketch association (best-effort) ------------------------
    _find_creating_features(body, holes)

    body_name = ''
    try:
        body_name = body.name
    except Exception:
        pass

    return {
        'body_token': token,
        'body_name': body_name or group.get('display_name', ''),
        'component_path': group.get('component_path', ''),
        'material_name': group.get('material_name', 'Unknown'),
        'dimensions_mm': group.get('dimensions_mm'),
        'total_holes': sum(len(cg.get('holes', [])) for cg in concentric_groups),
        'concentric_groups': concentric_groups,
        'face_images': {},  # populated later by weld_plan_generator
        # Channel geometry for annotation overlays
        'channel_axis': channel_axis,
        'channel_length_mm': channel_length_mm,
        'bb_min_cm': bb_min_cm,
        'bb_max_cm': bb_max_cm,
        'estimated_wall_thickness_mm': group.get('estimated_wall_thickness_mm'),
    }


# ===========================================================================
# Channel axis detection (from bounding box)
# ===========================================================================

def _channel_axis_from_bb(
    body: adsk.fusion.BRepBody,
) -> Tuple[Optional[Vec3], Optional[Vec3], float, Optional[Vec3], Optional[Vec3]]:
    """Determine the channel's long axis from its AABB.

    Returns (axis_unit, bb_min_cm, channel_length_mm, bb_min_cm, bb_max_cm).
    The long axis is the bounding-box axis with the largest extent.
    """
    try:
        bb = body.boundingBox
        mn = bb.minPoint
        mx = bb.maxPoint
        dx = mx.x - mn.x
        dy = mx.y - mn.y
        dz = mx.z - mn.z
        bb_min_cm: Vec3 = (mn.x, mn.y, mn.z)
        bb_max_cm: Vec3 = (mx.x, mx.y, mx.z)

        extents = [(dx, (1, 0, 0)), (dy, (0, 1, 0)), (dz, (0, 0, 1))]
        extents.sort(key=lambda e: e[0], reverse=True)
        length_cm = extents[0][0]
        axis = extents[0][1]
        return axis, bb_min_cm, length_cm * CM_TO_MM, bb_min_cm, bb_max_cm
    except Exception:
        return None, None, 0.0, None, None


def _build_face_label_axes(
    body: adsk.fusion.BRepBody,
    channel_axis: Optional[Vec3],
) -> Dict[str, Vec3]:
    """Build axis→label mapping for the four cross-section walls.

    Uses the two largest-area planar face normal directions perpendicular to
    the channel axis.  Labels: C/D (largest pair) and E/F (next pair).
    A/B are reserved for the lengthwise channel ends.
    """
    labels: Dict[str, Vec3] = {}
    if channel_axis is None:
        return labels

    plane_type = adsk.core.Plane.classType()
    normals_by_area: List[Tuple[Vec3, float]] = []

    try:
        for face in body.faces:
            try:
                geom = face.geometry
                if geom.objectType != plane_type:
                    continue
                n_raw = geom.normal
                n = _normalize((n_raw.x, n_raw.y, n_raw.z))
                if n is None:
                    continue
                # Only perpendicular to channel axis (wall faces)
                if abs(_dot(n, channel_axis)) > 0.1:
                    continue
                normals_by_area.append((n, face.area))
            except Exception:
                pass
    except Exception:
        pass

    if not normals_by_area:
        return labels

    # Merge anti-parallel normals
    merged: List[Tuple[Vec3, float]] = []
    consumed = [False] * len(normals_by_area)
    for i, (ni, ai) in enumerate(normals_by_area):
        if consumed[i]:
            continue
        total = ai
        for j in range(i + 1, len(normals_by_area)):
            if consumed[j]:
                continue
            nj = normals_by_area[j][0]
            if abs(_dot(ni, nj)) > PARALLEL_TOL:
                total += normals_by_area[j][1]
                consumed[j] = True
        consumed[i] = True
        merged.append((ni, total))

    merged.sort(key=lambda x: x[1], reverse=True)

    # Assign labels: largest-area pair → C/D, next → E/F
    # (A/B are reserved for the lengthwise channel ends)
    if len(merged) >= 1:
        n1 = merged[0][0]
        max_comp = max(range(3), key=lambda c: abs(n1[c]))
        if n1[max_comp] >= 0:
            labels['C'] = n1
            labels['D'] = _scale(n1, -1.0)
        else:
            labels['D'] = n1
            labels['C'] = _scale(n1, -1.0)

    if len(merged) >= 2:
        n2 = merged[1][0]
        max_comp = max(range(3), key=lambda c: abs(n2[c]))
        if n2[max_comp] >= 0:
            labels['E'] = n2
            labels['F'] = _scale(n2, -1.0)
        else:
            labels['F'] = n2
            labels['E'] = _scale(n2, -1.0)

    return labels


# ===========================================================================
# Hole depth computation
# ===========================================================================

def _compute_hole_depth(
    face: adsk.fusion.BRepFace,
    axis: Vec3,
) -> Optional[float]:
    """Compute the axial depth of a cylindrical face from its edge loops.

    The two bounding circles of a cylindrical face have centres at different
    positions along the cylinder axis.  The distance between those projections
    is the depth.
    """
    try:
        min_proj = None
        max_proj = None
        circle_type = adsk.core.Circle3D.classType()
        arc_type = adsk.core.Arc3D.classType()

        for loop in face.loops:
            for edge in loop.edges:
                try:
                    eg = edge.geometry
                    ot = eg.objectType
                    if ot == circle_type:
                        c = eg.center
                    elif ot == arc_type:
                        c = eg.center
                    else:
                        # Use edge midpoint for non-circular edges
                        ev = edge.evaluator
                        _, sp, ep = ev.getEndPoints()
                        c = adsk.core.Point3D.create(
                            (sp.x + ep.x) / 2, (sp.y + ep.y) / 2, (sp.z + ep.z) / 2
                        )

                    proj = _dot(axis, (c.x, c.y, c.z))
                    if min_proj is None or proj < min_proj:
                        min_proj = proj
                    if max_proj is None or proj > max_proj:
                        max_proj = proj
                except Exception:
                    pass

        if min_proj is not None and max_proj is not None:
            depth_cm = abs(max_proj - min_proj)
            return depth_cm * CM_TO_MM if depth_cm > 1e-6 else None
    except Exception:
        pass
    return None


# ===========================================================================
# Face direction labelling
# ===========================================================================

def _classify_face_direction(
    hole_axis: Vec3,
    label_axes: Dict[str, Vec3],
) -> str:
    """Classify which channel face a hole points through.

    Returns 'C', 'D', 'E', 'F', or 'End'.
    """
    if not label_axes:
        return 'Unknown'

    best_label = 'Unknown'
    best_dot = -1.0
    for label, direction in label_axes.items():
        d = abs(_dot(hole_axis, direction))
        if d > best_dot:
            best_dot = d
            best_label = label
    return best_label if best_dot > 0.7 else 'End'


# ===========================================================================
# Cross-section offset (centered vs off-centre hole)
# ===========================================================================

def _cross_section_offset(
    center_cm: Vec3,
    hole_axis: Vec3,
    channel_axis: Optional[Vec3],
    bb_min_cm: Optional[Vec3],
    bb_max_cm: Optional[Vec3],
) -> Tuple[Optional[float], Optional[float]]:
    """Compute how far a hole is from the centre of the channel face it pierces.

    The "transverse" direction is perpendicular to both the channel long axis
    and the hole axis.  We project the hole centre onto this direction and
    compare with the bounding-box centre.

    Returns (offset_mm, face_width_mm).
    offset_mm = 0 means centred.  Positive/negative is arbitrary (magnitude matters).
    face_width_mm = extent of the channel in the transverse direction.
    """
    if channel_axis is None or bb_min_cm is None or bb_max_cm is None:
        return None, None

    try:
        # Transverse direction = perpendicular to both channel axis and hole axis
        trans = _cross(channel_axis, hole_axis)
        trans = _normalize(trans)
        if trans is None:
            # Hole axis is parallel to channel axis (unusual) — fall back
            return None, None

        # Bounding-box centre in world coordinates
        bb_center = (
            (bb_min_cm[0] + bb_max_cm[0]) / 2.0,
            (bb_min_cm[1] + bb_max_cm[1]) / 2.0,
            (bb_min_cm[2] + bb_max_cm[2]) / 2.0,
        )

        # Project hole centre and BB centre onto transverse axis
        hole_proj = _dot(center_cm, trans)
        bb_proj = _dot(bb_center, trans)
        offset_cm = hole_proj - bb_proj

        # Face width = extent of BB along transverse direction
        bb_extent_cm = abs(_dot(_sub(bb_max_cm, bb_min_cm), trans))
        return offset_cm * CM_TO_MM, bb_extent_cm * CM_TO_MM
    except Exception:
        return None, None


# ===========================================================================
# Edge distances
# ===========================================================================

def _edge_distances(
    center_cm: Vec3,
    channel_axis: Optional[Vec3],
    channel_origin: Optional[Vec3],
    channel_length_mm: float,
) -> Tuple[Optional[float], Optional[float]]:
    """Distance from hole centre to each end of the channel.

    Returns (dist_end_a_mm, dist_end_b_mm).
    End A is always the bb_min side along the channel axis (consistent per channel).
    End B is always the bb_max side.
    """
    if channel_axis is None or channel_origin is None or channel_length_mm <= 0:
        return None, None

    proj_mm = _project_onto_axis(center_cm, channel_origin, channel_axis) * CM_TO_MM
    dist_a = proj_mm
    dist_b = channel_length_mm - proj_mm

    return max(dist_a, 0.0), max(dist_b, 0.0)


# ===========================================================================
# Thread hint lookup
# ===========================================================================

def _get_thread_hint(diameter_mm: float, hint_type: str = 'tap') -> str:
    """Return an advisory thread-size hint for the given hole diameter.

    Parameters
    ----------
    diameter_mm : float
        Hole diameter.
    hint_type : str
        ``'tap'`` for blind holes (likely threaded),
        ``'clearance'`` for through-holes (likely bolt pass-through).
    """
    table = _TAP_DRILL_TABLE if hint_type == 'tap' else _CLEARANCE_TABLE
    best_label = ''
    best_delta = THREAD_HINT_TOL_MM + 1.0

    for size_mm, label in table:
        delta = abs(diameter_mm - size_mm)
        if delta < best_delta:
            best_delta = delta
            best_label = label

    return best_label if best_delta <= THREAD_HINT_TOL_MM else ''


# ===========================================================================
# Concentric grouping
# ===========================================================================

def _group_concentric_holes(
    holes: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """Group holes that share the same axis position (concentric through-holes).

    Two holes belong to the same concentric group when:
    1. Their axes are parallel (|dot| > PARALLEL_TOL)
    2. Their projected centres are within CENTER_PROXIMITY_MM

    Returns a list of concentric-group dicts.
    """
    n = len(holes)
    group_of = [-1] * n
    next_group = 0

    for i in range(n):
        if group_of[i] >= 0:
            continue
        group_of[i] = next_group
        axis_i = holes[i]['axis']
        center_i = holes[i]['center_cm']

        for j in range(i + 1, n):
            if group_of[j] >= 0:
                continue
            axis_j = holes[j]['axis']
            center_j = holes[j]['center_cm']

            # Axes must be parallel
            if abs(_dot(axis_i, axis_j)) < PARALLEL_TOL:
                continue

            # Projected centres must be close (perpendicular distance)
            diff = _sub(center_j, center_i)
            along = _dot(diff, axis_i)
            perp = _sub(diff, _scale(axis_i, along))
            perp_dist_mm = _length(perp) * CM_TO_MM

            if perp_dist_mm <= CENTER_PROXIMITY_MM:
                group_of[j] = next_group

        next_group += 1

    # Build group dicts
    groups_map: Dict[int, List[Dict[str, Any]]] = {}
    for i, gi in enumerate(group_of):
        holes[i]['concentric_group'] = gi
        groups_map.setdefault(gi, []).append(holes[i])

    result: List[Dict[str, Any]] = []
    for gi in sorted(groups_map.keys()):
        group_holes = groups_map[gi]
        # Sort by radius descending (counterbore first)
        group_holes.sort(key=lambda h: h['radius_mm'], reverse=True)

        largest_d = group_holes[0]['diameter_mm']
        smallest_d = group_holes[-1]['diameter_mm']
        has_multiple_radii = abs(largest_d - smallest_d) > 0.05

        # Through-hole detection: if there are multiple concentric faces
        # of the same radius, one on each side, it's a through hole.
        is_through = len(group_holes) >= 2
        for h in group_holes:
            h['is_through'] = is_through

        # Type label
        if has_multiple_radii:
            hole_type = 'Counterbored'
        elif is_through:
            hole_type = 'Through Hole'
        else:
            hole_type = 'One Side'

        # Merge through-holes of the same radius into a single representative
        # entry.  The user sees one row per physical hole, not two rows for
        # the entry and exit cylindrical faces on the same channel wall.
        if is_through and not has_multiple_radii and len(group_holes) > 1:
            merged = group_holes[0].copy()
            merged['face_tokens'] = []
            for h in group_holes:
                merged['face_tokens'].extend(h.get('face_tokens', []))
            group_holes = [merged]

        # Group centre and axis (use first hole's values)
        result.append({
            'group_index': gi,
            'center_cm': group_holes[0]['center_cm'],
            'axis': group_holes[0]['axis'],
            'holes': group_holes,
            'largest_diameter_mm': largest_d,
            'smallest_diameter_mm': smallest_d,
            'is_through': is_through,
            'hole_type_label': hole_type,
            'center_to_center_mm': {},  # populated by _compute_center_distances
        })

    return result


# ===========================================================================
# Centre-to-centre distances
# ===========================================================================

def _compute_center_distances(groups: List[Dict[str, Any]]) -> None:
    """Compute pairwise Euclidean distances between group centres (mutates in place)."""
    for i, gi in enumerate(groups):
        ci = gi['center_cm']
        for j in range(i + 1, len(groups)):
            cj = groups[j]['center_cm']
            d_mm = _dist(ci, cj) * CM_TO_MM
            d_mm = round(d_mm, 1)
            gi['center_to_center_mm'][groups[j]['group_index']] = d_mm
            groups[j]['center_to_center_mm'][gi['group_index']] = d_mm


# ===========================================================================
# Feature / sketch association
# ===========================================================================

def _find_creating_features(
    body: adsk.fusion.BRepBody,
    holes: List[Dict[str, Any]],
) -> None:
    """Associate each hole with the timeline feature that created it.

    Scans the design timeline once and checks face membership to find
    HoleFeature or ExtrudeFeature (cut) that owns each hole face.
    Mutates hole dicts in place (sets creating_feature_name, creating_sketch_token).
    """
    # Build reverse index: face_token → hole dict
    face_to_hole: Dict[str, Dict[str, Any]] = {}
    for h in holes:
        for ft in h['face_tokens']:
            if ft:
                face_to_hole[ft] = h

    if not face_to_hole:
        return

    try:
        app = adsk.core.Application.get()
        design = adsk.fusion.Design.cast(app.activeProduct)
        if not design or not design.timeline:
            return

        hole_feature_type = adsk.fusion.HoleFeature.classType()
        extrude_feature_type = adsk.fusion.ExtrudeFeature.classType()

        for tl_obj in design.timeline:
            try:
                entity = tl_obj.entity
                if entity is None:
                    continue

                ot = entity.objectType
                is_hole_feature = (ot == hole_feature_type)
                is_extrude = (ot == extrude_feature_type)

                if not is_hole_feature and not is_extrude:
                    continue

                # For extrude features, only care about cuts
                if is_extrude:
                    try:
                        op = entity.operation
                        if op not in (
                            adsk.fusion.FeatureOperations.CutFeatureOperation,
                            adsk.fusion.FeatureOperations.IntersectFeatureOperation,
                        ):
                            continue
                    except Exception:
                        continue

                # Check if any of this feature's faces match our hole faces
                feature_name = 'HoleFeature' if is_hole_feature else 'ExtrudeFeature (Cut)'
                sketch_token = ''
                try:
                    sketch = entity.sketch if hasattr(entity, 'sketch') else None
                    if sketch:
                        sketch_token = sketch.entityToken
                except Exception:
                    pass

                try:
                    for feat_face in entity.faces:
                        try:
                            ft = feat_face.entityToken
                            if ft in face_to_hole:
                                h = face_to_hole[ft]
                                h['creating_feature_name'] = feature_name
                                h['creating_sketch_token'] = sketch_token
                        except Exception:
                            pass
                except Exception:
                    pass

            except Exception:
                pass

    except Exception:
        logger.debug('Feature/sketch association failed: %s', traceback.format_exc())
