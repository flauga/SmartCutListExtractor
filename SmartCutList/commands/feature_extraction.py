"""
feature_extraction.py  —  SmartCutList Phase 2
===============================================
Given a list of adsk.fusion.BRepBody objects (from Phase 1 selection), extract
geometric, topological, and material properties and return them as a list of
plain Python dicts suitable for downstream processing or export.

Unit convention
---------------
  Fusion 360 internal units  :  centimetres (cm)
  Length / thickness output  :  millimetres (mm)   [ value × CM_TO_MM ]
  Volume output              :  cubic centimetres (cm³)   [ body.volume as-is ]
  Surface area output        :  square centimetres (cm²)  [ body.area  as-is ]
"""
from __future__ import annotations

import math
import traceback
from collections import namedtuple
from typing import Any, Dict, List, Optional, Tuple

import adsk.core
import adsk.fusion

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CM_TO_MM: float = 10.0      # Fusion internal cm → output mm
MM_TO_CM: float = 0.1       # mm → cm (for bounding-box volume in cm³)

# Tolerances for axis-alignment tests  (|cos θ| thresholds)
PARALLEL_TOL: float = 0.9998   # above → parallel or anti-parallel
PERP_TOL:     float = 0.05     # below → perpendicular

# Minimum meaningful thickness (filters numerical noise from face-distance calc)
MIN_THICKNESS_MM: float = 0.01

# bb_fill_ratio above this → body treated as solid for wall-thickness purposes
SOLID_FILL_RATIO: float = 0.85

# Timeline feature-history extraction is disabled for the cut-list workflow.
# It adds noticeable overhead on larger parametric designs and is not required
# for the current review/export pipeline.
INCLUDE_FEATURE_HISTORY: bool = False

# ---------------------------------------------------------------------------
# Geometry-type name tables  (built once at module import)
# ---------------------------------------------------------------------------

# face.geometry.objectType → human-readable label
_FACE_GEOM_TYPES: Dict[str, str] = {}
for _cls, _lbl in [
    (adsk.core.Plane,        'Plane'),
    (adsk.core.Cylinder,     'Cylinder'),
    (adsk.core.Cone,         'Cone'),
    (adsk.core.Sphere,       'Sphere'),
    (adsk.core.Torus,        'Torus'),
    (adsk.core.NurbsSurface, 'NurbsSurface'),
]:
    try:
        _FACE_GEOM_TYPES[_cls.classType()] = _lbl
    except Exception:
        pass

# edge.geometry.objectType → human-readable label
_EDGE_GEOM_TYPES: Dict[str, str] = {}
for _cls, _lbl in [
    (adsk.core.Line3D,          'Line'),
    (adsk.core.Arc3D,           'Arc'),
    (adsk.core.Circle3D,        'Circle'),
    (adsk.core.Ellipse3D,       'Ellipse'),
    (adsk.core.EllipticalArc3D, 'EllipticalArc'),
    (adsk.core.NurbsCurve3D,    'NurbsCurve'),
    (adsk.core.InfiniteLine3D,  'InfiniteLine'),
]:
    try:
        _EDGE_GEOM_TYPES[_cls.classType()] = _lbl
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Type aliases
# ---------------------------------------------------------------------------

Vec3 = Tuple[float, float, float]      # (x, y, z) float tuple

# Single-pass face analysis result — computed once per body and shared across
# all downstream functions (_get_obb_mm, _has_constant_cross_section,
# _estimate_wall_thickness_mm) to avoid re-iterating body.faces.
#
# Fields
# ------
# planar      : list of dicts {'n': Vec3, 'area': float, 'o': Vec3}
#               unit normal, face area (cm²), plane origin (cm) — planar faces only
# non_planar  : list of BRepFace objects whose geometry is not a Plane
# type_counts : {label: count} for all faces (already tallied in the same pass)
# total       : total face count
_FaceAnalysis = namedtuple('_FaceAnalysis', ('planar', 'non_planar', 'type_counts', 'total'))


# ===========================================================================
# Public API
# ===========================================================================

def extract_features(bodies: List[adsk.fusion.BRepBody]) -> List[Dict[str, Any]]:
    """
    Extract geometric and material properties from each BRepBody.

    Builds the design timeline feature-map once (rather than once per body)
    and passes it down to avoid O(T × B) timeline scans.

    :param bodies:  List of adsk.fusion.BRepBody objects from Phase 1.
    :returns:       List of property dicts, one per body.  On per-body failure
                    the dict contains ``'body_name'`` and ``'extraction_error'``
                    keys so the caller can log and continue.
    """
    if not bodies:
        return []

    # One-shot timeline scan: {entityToken: [featureTypeName, ...]}
    feature_map = _build_feature_map() if INCLUDE_FEATURE_HISTORY else {}

    results: List[Dict[str, Any]] = []
    for body in bodies:
        try:
            results.append(_extract_body_features(body, feature_map))
        except Exception:
            results.append({
                'body_name':        _safe_attr(body, 'name', '<unknown>'),
                'extraction_error': traceback.format_exc(),
            })
    return results


def extract_features_with_context(body_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Like extract_features() but accepts the full item dicts from select_components
    (each has 'entity', 'component_path', 'name', etc.) and copies component_path
    into the extracted feature dict.

    :param body_items: List of dicts with at least {'entity': BRepBody, 'component_path': str}
    :returns:          List of property dicts with 'component_path' included.
    """
    if not body_items:
        return []

    feature_map = _build_feature_map() if INCLUDE_FEATURE_HISTORY else {}

    results: List[Dict[str, Any]] = []
    for item in body_items:
        body = item.get('entity')
        component_path = item.get('component_path', '')
        try:
            props = _extract_body_features(body, feature_map)
            props['component_path'] = component_path
            props['is_content_library_fastener'] = item.get('is_content_library_fastener', False)
            props['is_sourced_component'] = item.get('is_sourced_component', False)
            results.append(props)
        except Exception:
            results.append({
                'body_name':        _safe_attr(body, 'name', '<unknown>') if body else '<unknown>',
                'component_path':   component_path,
                'is_content_library_fastener': item.get('is_content_library_fastener', False),
                'is_sourced_component': item.get('is_sourced_component', False),
                'extraction_error': traceback.format_exc(),
            })
    return results


# ===========================================================================
# Core per-body extractor
# ===========================================================================

def _extract_body_features(
    body: adsk.fusion.BRepBody,
    feature_map: Dict[str, List[str]],
) -> Dict[str, Any]:
    """Build the complete property dict for a single BRepBody."""
    props: Dict[str, Any] = {}

    # ------------------------------------------------------------------
    # 1 & 2.  Identity
    # ------------------------------------------------------------------
    props['component_name'] = _safe_attr(body.parentComponent, 'name')
    props['body_name']      = _safe_attr(body, 'name')
    props['body_token']     = _safe_attr(body, 'entityToken')

    # ------------------------------------------------------------------
    # 3.  Material  (body → component → 'Unknown')
    # ------------------------------------------------------------------
    props['material_name'] = _get_material_name(body)

    # ------------------------------------------------------------------
    # Single-pass face analysis (shared by properties 4, 8, 14, 15)
    # ------------------------------------------------------------------
    face_data = _analyze_faces(body)

    # ------------------------------------------------------------------
    # 4.  Bounding box  [L, W, H] mm, sorted descending.
    #     AABB is computed once here and reused inside _get_obb_mm to
    #     avoid a second body.boundingBox read for the sanity check.
    # ------------------------------------------------------------------
    aabb  = _get_aabb_mm(body)
    bb_mm = _get_obb_mm(body, face_data, aabb) or aabb
    props['bounding_box_mm'] = bb_mm

    # ------------------------------------------------------------------
    # 5 & 6.  Scalar physical properties
    #         Fusion returns body.volume in cm³, body.area in cm².
    # ------------------------------------------------------------------
    props['volume_cm3']       = _safe_attr(body, 'volume')
    props['surface_area_cm2'] = _safe_attr(body, 'area')

    # ------------------------------------------------------------------
    # 7.  Sheet metal flag
    # ------------------------------------------------------------------
    props['is_sheet_metal'] = _safe_attr(body, 'isSheetMetal', False)

    # ------------------------------------------------------------------
    # 8 & 10.  Face topology counts  (from the pre-computed face_data)
    # ------------------------------------------------------------------
    props['total_faces']      = face_data.total
    props['face_type_counts'] = face_data.type_counts

    # ------------------------------------------------------------------
    # 9 & 11.  Edge topology counts  (edges still need their own pass)
    # ------------------------------------------------------------------
    props['total_edges'], props['edge_type_counts'] = \
        _count_geom_types(body.edges, _EDGE_GEOM_TYPES)

    # ------------------------------------------------------------------
    # 12 & 13.  BB fill ratio and aspect ratios.
    #           Guard is evaluated once; L/W/H are unpacked once.
    # ------------------------------------------------------------------
    vol_cm3  = props['volume_cm3']
    bb_valid = bb_mm is not None and vol_cm3 is not None and all(d and d > 0 for d in bb_mm)

    if bb_valid:
        L, W, H = bb_mm
        bb_vol_cm3 = (L * MM_TO_CM) * (W * MM_TO_CM) * (H * MM_TO_CM)
        props['bb_fill_ratio'] = vol_cm3 / bb_vol_cm3 if bb_vol_cm3 > 0 else None
        props['aspect_ratios'] = [_safe_div(L, W), _safe_div(L, H), _safe_div(W, H)]
    else:
        props['bb_fill_ratio'] = None
        props['aspect_ratios'] = None

    # ------------------------------------------------------------------
    # 14.  Constant cross-section heuristic  (extrusion detector)
    # ------------------------------------------------------------------
    props['has_constant_cross_section'] = _has_constant_cross_section(face_data)

    # ------------------------------------------------------------------
    # 15.  Estimated wall thickness  (None for solid bodies)
    # ------------------------------------------------------------------
    props['estimated_wall_thickness_mm'] = _estimate_wall_thickness_mm(
        body, bb_mm, props['bb_fill_ratio'], face_data
    )

    # ------------------------------------------------------------------
    # 16.  Feature history  (O(1) lookup into the pre-built map)
    # ------------------------------------------------------------------
    props['feature_history'] = _get_feature_history(body, feature_map)

    return props


# ===========================================================================
# Single-pass face analysis
# ===========================================================================

def _analyze_faces(body: adsk.fusion.BRepBody) -> _FaceAnalysis:
    """
    Iterate body.faces exactly once and return all geometry data needed by
    _get_obb_mm, _has_constant_cross_section, _estimate_wall_thickness_mm,
    and the face_type_counts / total_faces properties.
    """
    plane_type = adsk.core.Plane.classType()

    planar:     List[Dict[str, Any]] = []   # {'n': Vec3, 'area': float, 'o': Vec3}
    non_planar: List                 = []   # BRepFace objects
    type_counts: Dict[str, int]      = {}
    total = 0

    try:
        for face in body.faces:
            total += 1
            try:
                geom = face.geometry
                ot   = geom.objectType
                label = _FACE_GEOM_TYPES.get(ot, 'Other')
                type_counts[label] = type_counts.get(label, 0) + 1

                if ot == plane_type:
                    n    = geom.normal
                    o    = geom.origin
                    unit = _normalize3((n.x, n.y, n.z))
                    if unit is not None:
                        planar.append({
                            'n':    unit,
                            'area': face.area,
                            'o':    (o.x, o.y, o.z),
                        })
                else:
                    non_planar.append(face)

            except Exception:
                type_counts['Unknown'] = type_counts.get('Unknown', 0) + 1
    except Exception:
        pass

    return _FaceAnalysis(
        planar=planar,
        non_planar=non_planar,
        type_counts=type_counts,
        total=total,
    )


# ===========================================================================
# Individual property extractors
# ===========================================================================

# ---------------------------------------------------------------------------
# 3. Material
# ---------------------------------------------------------------------------

def _get_material_name(body: adsk.fusion.BRepBody) -> str:
    """Return material name: body material → component material → 'Unknown'."""
    for source in (body, body.parentComponent):
        try:
            mat = source.material
            if mat is not None:
                return mat.name
        except Exception:
            pass
    return 'Unknown'


# ---------------------------------------------------------------------------
# 4. Bounding box
# ---------------------------------------------------------------------------

def _get_aabb_mm(body: adsk.fusion.BRepBody) -> Optional[List[float]]:
    """Axis-aligned bounding box from body.boundingBox, returned in mm."""
    try:
        bb = body.boundingBox
        dx = (bb.maxPoint.x - bb.minPoint.x) * CM_TO_MM
        dy = (bb.maxPoint.y - bb.minPoint.y) * CM_TO_MM
        dz = (bb.maxPoint.z - bb.minPoint.z) * CM_TO_MM
        return sorted([dx, dy, dz], reverse=True)
    except Exception:
        return None


def _get_obb_mm(
    body:      adsk.fusion.BRepBody,
    face_data: _FaceAnalysis,
    aabb:      Optional[List[float]],
) -> Optional[List[float]]:
    """
    Oriented bounding box (OBB) from dominant planar-face principal axes.

    Accepts the pre-computed face_data (planar normals/areas from a single
    body.faces pass) and the AABB (already computed by the caller) so that
    neither body.faces nor body.boundingBox is read a second time.

    Algorithm
    ---------
    1. Use face_data.planar normals/areas.
    2. Merge anti-parallel normals (same physical axis) and rank by total area.
    3. Primary axis (ax1) = highest total-area direction.
    4. Secondary axis (ax2) = highest-area planar face perpendicular to ax1,
       then Gram–Schmidt orthogonalised against ax1.
    5. Tertiary axis (ax3) = ax1 × ax2.
    6. Project every body vertex onto {ax1, ax2, ax3}; min/max extents → OBB.
    7. Sanity-check: if OBB volume > AABB volume (numerical drift), discard OBB.

    Returns None when fewer than 3 planar faces exist (caller uses AABB).
    """
    try:
        planar = face_data.planar      # List of {'n': Vec3, 'area': float, ...}

        if len(planar) < 3:
            return None

        # ---- Step 2: merge anti-parallel normals & rank by total area -------
        consumed = [False] * len(planar)
        axes: List[Tuple[Vec3, float]] = []   # (direction, total_area)
        for i, fd_i in enumerate(planar):
            if consumed[i]:
                continue
            ni    = fd_i['n']
            total = fd_i['area']
            for j in range(i + 1, len(planar)):
                if consumed[j]:
                    continue
                if abs(_dot3(ni, planar[j]['n'])) > PARALLEL_TOL:
                    total += planar[j]['area']
                    consumed[j] = True
            consumed[i] = True
            axes.append((ni, total))

        axes.sort(key=lambda x: x[1], reverse=True)
        ax1 = axes[0][0]

        # ---- Step 3: secondary axis — most-area face ⊥ to ax1 --------------
        ax2: Optional[Vec3] = None
        best_area = -1.0
        for fd in planar:
            if abs(_dot3(ax1, fd['n'])) < PERP_TOL and fd['area'] > best_area:
                best_area = fd['area']
                ax2       = fd['n']

        if ax2 is None:
            return None

        # Gram–Schmidt: remove ax1 component from ax2 for numerical stability
        dot = _dot3(ax1, ax2)
        ax2 = _normalize3((ax2[0] - dot * ax1[0],
                            ax2[1] - dot * ax1[1],
                            ax2[2] - dot * ax1[2]))
        if ax2 is None:
            return None

        # ---- Step 4: tertiary axis ------------------------------------------
        ax3 = _normalize3(_cross3(ax1, ax2))
        if ax3 is None:
            return None

        # ---- Step 5 & 6: project vertices, compute extents ------------------
        e1_min = e1_max = e2_min = e2_max = e3_min = e3_max = None
        for vertex in body.vertices:
            try:
                p  = vertex.geometry          # adsk.core.Point3D (cm)
                pt = (p.x, p.y, p.z)
                p1 = _dot3(ax1, pt)
                p2 = _dot3(ax2, pt)
                p3 = _dot3(ax3, pt)
                if e1_min is None:
                    e1_min = e1_max = p1
                    e2_min = e2_max = p2
                    e3_min = e3_max = p3
                else:
                    if p1 < e1_min: e1_min = p1
                    if p1 > e1_max: e1_max = p1
                    if p2 < e2_min: e2_min = p2
                    if p2 > e2_max: e2_max = p2
                    if p3 < e3_min: e3_min = p3
                    if p3 > e3_max: e3_max = p3
            except Exception:
                continue

        if e1_min is None:
            return None

        dims_mm = sorted([
            (e1_max - e1_min) * CM_TO_MM,
            (e2_max - e2_min) * CM_TO_MM,
            (e3_max - e3_min) * CM_TO_MM,
        ], reverse=True)

        # ---- Step 7: sanity check — OBB must not exceed AABB ----------------
        if aabb:
            obb_vol  = dims_mm[0] * dims_mm[1] * dims_mm[2]
            aabb_vol = aabb[0]   * aabb[1]    * aabb[2]
            if obb_vol > aabb_vol * 1.01:
                return None          # unexpected; caller falls back to AABB

        return dims_mm

    except Exception:
        return None


# ---------------------------------------------------------------------------
# 8 & 9. Topology counts  (edges only; faces come from _analyze_faces)
# ---------------------------------------------------------------------------

def _count_geom_types(
    collection,
    type_map: Dict[str, str],
) -> Tuple[int, Dict[str, int]]:
    """
    Tally geometry-type labels across a BRepEdges collection.

    :param collection:  body.edges
    :param type_map:    _EDGE_GEOM_TYPES
    :returns:           (total_count, {label: count})
    """
    counts: Dict[str, int] = {}
    total = 0
    try:
        for item in collection:
            total += 1
            try:
                label = type_map.get(item.geometry.objectType, 'Other')
            except Exception:
                label = 'Unknown'
            counts[label] = counts.get(label, 0) + 1
    except Exception:
        pass
    return total, counts


# ---------------------------------------------------------------------------
# 14. Constant cross-section heuristic
# ---------------------------------------------------------------------------

def _has_constant_cross_section(face_data: _FaceAnalysis) -> bool:
    """
    Return True when the body appears to have been created by extruding a
    2-D profile (i.e. all cross-sections perpendicular to one axis are equal).

    Uses the pre-computed _FaceAnalysis to avoid re-iterating body.faces.

    Heuristic
    ---------
    (a) There exist ≥ 2 anti-parallel planar faces whose areas agree within 5 %
        — these are the "end caps" that bracket the extrusion.
    (b) Every other planar face has its normal perpendicular to the cap axis
        (i.e. it is a "wall" face parallel to the extrusion direction).
    (c) Every cylindrical face has its axis parallel to the cap axis (handles
        circular profiles and rounds on prismatic parts).

    Non-cylindrical, non-planar faces (e.g. fillets modelled as NURBS) are
    tolerated to avoid false negatives on chamfered parts.
    """
    try:
        planar     = face_data.planar
        non_planar = face_data.non_planar

        if len(planar) < 2:
            return False

        # Sort by area (largest first) to find the most-likely caps quickly
        planar_sorted = sorted(planar, key=lambda f: f['area'], reverse=True)

        # (a) Find the best anti-parallel pair (cap-axis candidate) -----------
        cap_axis: Optional[Vec3] = None
        for i in range(len(planar_sorted)):
            n1, a1 = planar_sorted[i]['n'], planar_sorted[i]['area']
            for j in range(i + 1, len(planar_sorted)):
                n2, a2 = planar_sorted[j]['n'], planar_sorted[j]['area']
                if _dot3(n1, n2) > -PARALLEL_TOL:
                    continue                       # not anti-parallel
                if max(a1, a2) > 0 and abs(a1 - a2) / max(a1, a2) > 0.05:
                    continue                       # areas differ >5 %
                cap_axis = n1
                break
            if cap_axis is not None:
                break

        if cap_axis is None:
            return False

        # (b) All non-cap planar faces must be ⊥ to cap_axis -----------------
        for f in planar:
            dot_abs = abs(_dot3(cap_axis, f['n']))
            if dot_abs > PARALLEL_TOL:
                continue              # this IS a cap face
            if dot_abs > PERP_TOL:
                return False          # oblique wall → not constant cross-section

        # (c) Cylindrical side-faces must be co-axial with cap_axis -----------
        cyl_type = adsk.core.Cylinder.classType()
        for face in non_planar:
            try:
                geom = face.geometry
                if geom.objectType != cyl_type:
                    continue          # non-cylindrical non-planar → tolerate
                ax = (geom.axis.x, geom.axis.y, geom.axis.z)
                if abs(_dot3(cap_axis, ax)) < PARALLEL_TOL:
                    return False      # tilted cylinder → not a simple extrusion
            except Exception:
                continue

        return True

    except Exception:
        return False


# ---------------------------------------------------------------------------
# 15. Wall-thickness estimation
# ---------------------------------------------------------------------------

def _estimate_wall_thickness_mm(
    body:        adsk.fusion.BRepBody,
    bb_dims_mm:  Optional[List[float]],
    fill_ratio:  Optional[float],
    face_data:   _FaceAnalysis,
) -> Optional[float]:
    """
    Estimate the wall thickness of hollow / thin-walled bodies in mm.

    Returns None when the body appears solid (bb_fill_ratio > SOLID_FILL_RATIO).

    Uses the pre-computed face_data.planar list (already has normals and
    origins) to avoid re-iterating body.faces.

    Strategy
    --------
    For sheet-metal bodies, attempt to read the active sheet-metal rule
    thickness directly (accurate, avoids geometry heuristic).

    General method: find every pair of anti-parallel planar faces and compute
    their perpendicular separation.  The minimum separation that is also less
    than 40 % of the smallest bounding-box dimension is returned as the
    estimated wall thickness (a larger value would be the part's overall
    extent, not a wall).
    """
    try:
        # Solid bodies — no meaningful wall thickness
        if fill_ratio is not None and fill_ratio > SOLID_FILL_RATIO:
            return None

        # Sheet-metal shortcut
        if _safe_attr(body, 'isSheetMetal', False):
            t = _get_sheet_metal_thickness_mm(body)
            if t is not None:
                return t
            # fall through to general method if rule not accessible

        planes = face_data.planar    # {'n': Vec3, 'area': float, 'o': Vec3}
        if len(planes) < 2:
            return None

        min_bb = min(bb_dims_mm) if bb_dims_mm else None
        candidates = []  # list of (distance_mm, area_weight)

        for i in range(len(planes)):
            ni, oi, ai = planes[i]['n'], planes[i]['o'], planes[i]['area']
            for j in range(i + 1, len(planes)):
                nj, oj, aj = planes[j]['n'], planes[j]['o'], planes[j]['area']
                if _dot3(ni, nj) > -PARALLEL_TOL:
                    continue           # not anti-parallel
                # Perpendicular distance = |projection of (oj - oi) onto ni|
                diff    = (oj[0] - oi[0], oj[1] - oi[1], oj[2] - oi[2])
                dist_mm = abs(_dot3(ni, diff)) * CM_TO_MM
                if dist_mm < MIN_THICKNESS_MM:
                    continue
                # Reject distances that are the body's overall extent
                if min_bb is not None and dist_mm >= min_bb * 0.4:
                    continue
                candidates.append((dist_mm, min(ai, aj)))

        if not candidates:
            return None

        # Cluster by rounding to 0.5mm buckets, pick the bucket with the
        # largest total face area.  This ensures the dominant wall thickness
        # (large structural faces) wins over tiny chamfer/fillet distances.
        buckets = {}  # rounded_dist → total_area
        for dist, area in candidates:
            bucket = round(dist * 2) / 2  # 0.5mm resolution
            buckets[bucket] = buckets.get(bucket, 0.0) + area

        best_dist = max(buckets, key=buckets.get)
        return round(best_dist, 4)

    except Exception:
        return None


def _get_sheet_metal_thickness_mm(body: adsk.fusion.BRepBody) -> Optional[float]:
    """
    Read wall thickness from the parent component's active sheet-metal rule.

    The rule's ``thickness`` is a ModelParameter whose ``value`` is in cm.
    Returns None if the component has no sheet-metal rules or if the API
    raises (e.g. the component is not sheet-metal).
    """
    try:
        comp     = body.parentComponent
        sm_rules = comp.sheetMetalRules
        if sm_rules is None or sm_rules.count == 0:
            return None
        for i in range(sm_rules.count):
            try:
                rule = sm_rules.item(i)
                if _safe_attr(rule, 'isActive', False):
                    t = rule.thickness
                    if t is not None:
                        return t.value * CM_TO_MM
            except Exception:
                continue
        return None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 16. Feature history (parametric timeline)
# ---------------------------------------------------------------------------

def _build_feature_map() -> Dict[str, List[str]]:
    """
    Build a reverse index from the design timeline: {entityToken: [featureTypeName, ...]}.

    Called once per extract_features() invocation.  Every body token is mapped
    to the list of feature types that produced it, so each per-body lookup
    is O(1) instead of O(timeline × bodies).

    Returns {} for direct-modelling designs or on any access error.
    """
    try:
        app    = adsk.core.Application.get()
        design = adsk.fusion.Design.cast(app.activeProduct)
        if design is None:
            return {}
        if design.designType != adsk.fusion.DesignTypes.ParametricDesignType:
            return {}

        result: Dict[str, List[str]] = {}
        timeline = design.timeline

        for i in range(timeline.count):
            try:
                entity = timeline.item(i).entity
                if entity is None or not hasattr(entity, 'bodies'):
                    continue
                feat_bodies = entity.bodies
                if feat_bodies is None:
                    continue

                # 'adsk::fusion::ExtrudeFeature' → 'ExtrudeFeature'
                label = entity.objectType.rsplit('::', 1)[-1]

                for j in range(feat_bodies.count):
                    try:
                        token = feat_bodies.item(j).entityToken
                        if token not in result:
                            result[token] = []
                        if label not in result[token]:
                            result[token].append(label)
                    except Exception:
                        continue
            except Exception:
                continue

        return result

    except Exception:
        return {}


def _get_feature_history(
    body:        adsk.fusion.BRepBody,
    feature_map: Dict[str, List[str]],
) -> List[str]:
    """
    Look up this body's feature history in the pre-built map (O(1)).

    Returns [] if the body has no entity token or is absent from the map.
    """
    try:
        return list(feature_map.get(body.entityToken, []))
    except Exception:
        return []


# ===========================================================================
# Math utilities
# ===========================================================================

def _dot3(a: Vec3, b: Vec3) -> float:
    """Dot product of two 3-tuples."""
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross3(a: Vec3, b: Vec3) -> Vec3:
    """Cross product  a × b."""
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _normalize3(a: Vec3) -> Optional[Vec3]:
    """Return unit vector, or None for degenerate (near-zero length) input."""
    mag = math.sqrt(a[0] * a[0] + a[1] * a[1] + a[2] * a[2])
    if mag < 1e-10:
        return None
    inv = 1.0 / mag
    return (a[0] * inv, a[1] * inv, a[2] * inv)


# ===========================================================================
# General helpers
# ===========================================================================

def _safe_attr(obj: Any, attr: str, default: Any = None) -> Any:
    """Return obj.attr, or default if the attribute access raises or is None."""
    try:
        val = getattr(obj, attr)
        return default if val is None else val
    except Exception:
        return default


def _safe_div(numerator: float, denominator: float) -> Optional[float]:
    """Return numerator / denominator, or None if denominator ≈ 0."""
    return None if abs(denominator) < 1e-10 else numerator / denominator
