"""
classifier.py — Rule-based structural part classifier for SmartCutList.

Public API
----------
    classify_bodies(features: list[dict]) -> list[dict]

Each input dict is the output of feature_extraction.extract_features().
Each output dict is the input dict extended with:
    classified_type         : str   (one of PartType constants)
    confidence              : float (0.0 – 1.0)
    classification_reason   : str
    needs_review            : bool  (True when confidence < CONFIDENCE_THRESHOLD)
"""

from __future__ import annotations

import math
import re
from typing import Optional


# ---------------------------------------------------------------------------
# Part type constants
# ---------------------------------------------------------------------------

class PartType:
    SHEET_METAL        = 'SheetMetal'
    RECTANGULAR_TUBE   = 'RectangularTube'
    ROUND_TUBE         = 'RoundTube'
    ALUMINIUM_EXTRUSION = 'AluminiumExtrusion'
    ANGLE_SECTION      = 'AngleSection'
    C_CHANNEL          = 'CChannel'
    FLAT_BAR           = 'FlatBar'
    MILLED_BLOCK       = 'MilledBlock'
    ROUND_BAR          = 'RoundBar'
    FASTENER           = 'Fastener'
    UNKNOWN            = 'Unknown'


CONFIDENCE_THRESHOLD = 0.6


# ---------------------------------------------------------------------------
# Name-keyword lookup tables
# (pattern -> (type, confidence))
# ---------------------------------------------------------------------------

# Each entry: (compiled_regex, PartType constant, confidence)
_NAME_KEYWORDS: list[tuple] = [
    # Fasteners — check before generic geometry rules
    (re.compile(r'\b(bolt|screw|nut|washer|rivet|pin|stud|anchor|cap.?head|hex.?head)\b', re.I),
     PartType.FASTENER, 0.85),

    # Aluminium extrusion profiles
    (re.compile(r'\b(extrusion|extrud|8020|80.?20|t.?slot|v.?slot|aluminium.?profile|aluminum.?profile|alu.?profile)\b', re.I),
     PartType.ALUMINIUM_EXTRUSION, 0.90),

    # Angle / L-bracket
    (re.compile(r'\b(angle|l.?bracket|l.?section|equal.?angle|unequal.?angle)\b', re.I),
     PartType.ANGLE_SECTION, 0.80),

    # C / U channel
    (re.compile(r'\b(c.?channel|u.?channel|channel|purlin|lip.?channel)\b', re.I),
     PartType.C_CHANNEL, 0.80),

    # Tube / hollow section
    (re.compile(r'\b(rhs|shs|rectangular.?tube|square.?tube|box.?section|hollow.?section)\b', re.I),
     PartType.RECTANGULAR_TUBE, 0.85),
    (re.compile(r'\b(chs|round.?tube|circular.?tube|pipe|hollow.?bar)\b', re.I),
     PartType.ROUND_TUBE, 0.85),

    # Flat bar / plate
    (re.compile(r'\b(flat.?bar|flat.?plate|flat.?stock|strap)\b', re.I),
     PartType.FLAT_BAR, 0.80),

    # Round bar / rod
    (re.compile(r'\b(round.?bar|rod|shaft|axle|spindle)\b', re.I),
     PartType.ROUND_BAR, 0.80),
]


# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------

def _dim(features: dict, key: str, index: int) -> Optional[float]:
    """Safely retrieve a dimension from the bounding box or aspect_ratio list."""
    val = features.get(key)
    if val is None:
        return None
    try:
        return float(val[index])
    except (IndexError, TypeError, ValueError):
        return None


def _sorted_dims(features: dict) -> Optional[tuple[float, float, float]]:
    """Return (small, mid, large) sorted bounding-box dimensions in mm, or None."""
    bb = features.get('bounding_box_mm')
    if not bb or len(bb) < 3:
        return None
    try:
        dims = sorted(float(d) for d in bb)
        if any(d <= 0 for d in dims):
            return None
        return tuple(dims)  # (small, mid, large)
    except (TypeError, ValueError):
        return None


def _has_cylindrical_face(features: dict) -> bool:
    """True when at least one cylindrical face is recorded."""
    counts = features.get('face_type_counts') or {}
    return int(counts.get('Cylinder', 0)) > 0


def _cylindrical_face_ratio(features: dict) -> float:
    """Fraction of faces that are cylindrical."""
    counts = features.get('face_type_counts') or {}
    total  = features.get('total_faces') or 0
    if total == 0:
        return 0.0
    return int(counts.get('Cylinder', 0)) / total


def _planar_face_ratio(features: dict) -> float:
    """Fraction of faces that are planar."""
    counts = features.get('face_type_counts') or {}
    total  = features.get('total_faces') or 0
    if total == 0:
        return 0.0
    return int(counts.get('Plane', 0)) / total


def _fill_ratio(features: dict) -> Optional[float]:
    return features.get('bb_fill_ratio')


def _aspect(features: dict) -> Optional[tuple[float, float, float]]:
    """Return aspect ratios [L/W, L/H, W/H] as a tuple, or None."""
    ar = features.get('aspect_ratios')
    if ar and len(ar) == 3 and all(v is not None for v in ar):
        try:
            return tuple(float(v) for v in ar)
        except (TypeError, ValueError):
            return None
    return None


def _is_elongated(features: dict, ratio_threshold: float = 3.0) -> bool:
    """True when the largest dimension is at least ratio_threshold × the smallest."""
    dims = _sorted_dims(features)
    if dims is None:
        return False
    small, _, large = dims
    return large / small >= ratio_threshold if small > 0 else False


def _wall_thickness_mm(features: dict) -> Optional[float]:
    return features.get('estimated_wall_thickness_mm')


def _is_hollow(features: dict, fill_threshold: float = 0.80) -> bool:
    """True when the fill ratio suggests a hollow interior."""
    fr = _fill_ratio(features)
    if fr is None:
        return False
    return fr < fill_threshold


# ---------------------------------------------------------------------------
# Rule chain
# Each rule returns Optional[tuple[str, float, str]] = (type, confidence, reason)
# or None to pass to the next rule.
# ---------------------------------------------------------------------------

def _rule_sheet_metal(features: dict):
    """Fusion's own sheet-metal flag — highest confidence."""
    if features.get('is_sheet_metal'):
        return PartType.SHEET_METAL, 0.98, 'Fusion sheet-metal body flag is set'
    return None


def _rule_name_keywords(features: dict):
    """Match component or body name against keyword groups."""
    name = ' '.join(filter(None, [
        features.get('component_name', ''),
        features.get('body_name', ''),
    ]))
    for pattern, part_type, confidence in _NAME_KEYWORDS:
        if pattern.search(name):
            match = pattern.search(name).group(0)
            return part_type, confidence, f'Name keyword match: "{match}"'
    return None


def _rule_fastener_geometry(features: dict):
    """
    Small, roughly equidimensional part with cylindrical faces and threads.
    Catches fasteners whose names are generic (e.g. "Body1").
    """
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, large = dims
    # Must be small (< 50 mm longest dim) and not too elongated
    if large > 60 or large / small > 5:
        return None
    cyl_ratio = _cylindrical_face_ratio(features)
    if cyl_ratio >= 0.3:
        # Check feature history for thread features
        history = features.get('feature_history') or []
        has_thread = any('thread' in f.lower() for f in history)
        confidence = 0.70 if has_thread else 0.60
        reason = 'Small cylindrical geometry' + (' with thread feature' if has_thread else '')
        return PartType.FASTENER, confidence, reason
    return None


def _rule_round_bar(features: dict):
    """Solid cylinder: high cylindrical face ratio + elongated + solid fill."""
    if not _is_elongated(features, ratio_threshold=2.0):
        return None
    cyl_ratio = _cylindrical_face_ratio(features)
    if cyl_ratio < 0.40:
        return None
    fr = _fill_ratio(features)
    if fr is not None and fr < 0.70:
        return None  # likely hollow → round tube
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, _ = dims
    # Cross-section should be roughly circular (small ≈ mid)
    if mid > 0 and small / mid < 0.70:
        return None
    confidence = 0.75 if (fr is not None and fr >= 0.85) else 0.65
    return PartType.ROUND_BAR, confidence, 'Solid elongated cylinder'


def _rule_round_tube(features: dict):
    """Hollow cylinder: cylindrical faces + hollow fill ratio + elongated."""
    if not _has_cylindrical_face(features):
        return None
    if not _is_elongated(features, ratio_threshold=2.0):
        return None
    if not _is_hollow(features, fill_threshold=0.80):
        return None
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, _ = dims
    # Roughly circular cross-section
    if mid > 0 and small / mid < 0.70:
        return None
    cyl_ratio = _cylindrical_face_ratio(features)
    confidence = 0.80 if cyl_ratio >= 0.40 else 0.65
    return PartType.ROUND_TUBE, confidence, 'Hollow elongated cylinder'


def _rule_rectangular_tube(features: dict):
    """
    Hollow rectangular/square section:
    - All planar faces
    - Elongated
    - Hollow fill ratio
    - Cross-section has 4 faces per end → ~8 planar faces total + 2 end caps
    """
    if _has_cylindrical_face(features):
        return None
    if not _is_elongated(features, ratio_threshold=2.5):
        return None
    if not _is_hollow(features, fill_threshold=0.80):
        return None
    planar_ratio = _planar_face_ratio(features)
    if planar_ratio < 0.85:
        return None
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, large = dims
    # Cross-section must not be too thin (that would be flat bar)
    if mid > 0 and small / mid < 0.20:
        return None
    ccs = features.get('has_constant_cross_section')
    confidence = 0.82 if ccs else 0.68
    reason = 'Hollow rectangular section' + (' with constant cross-section' if ccs else '')
    return PartType.RECTANGULAR_TUBE, confidence, reason


def _rule_flat_bar(features: dict):
    """
    Thin, wide, elongated solid plate:
    - High fill ratio (solid)
    - All planar faces
    - thickness << width << length  (small/mid < 0.30, large/mid > 2.5)
    """
    if _has_cylindrical_face(features):
        return None
    fr = _fill_ratio(features)
    if fr is not None and fr < 0.80:
        return None
    planar_ratio = _planar_face_ratio(features)
    if planar_ratio < 0.85:
        return None
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, large = dims
    if mid <= 0:
        return None
    thin = small / mid < 0.30
    long = large / mid > 2.5
    if not (thin and long):
        return None
    return PartType.FLAT_BAR, 0.78, 'Thin flat plate with high fill ratio'


def _rule_milled_block(features: dict):
    """
    Compact, solid, all-planar — a machined/milled block:
    - High fill ratio
    - Low elongation
    - All planar faces
    - Feature history contains extrude/cut/fillet/chamfer
    """
    if _has_cylindrical_face(features):
        return None
    fr = _fill_ratio(features)
    if fr is not None and fr < 0.75:
        return None
    planar_ratio = _planar_face_ratio(features)
    if planar_ratio < 0.75:
        return None
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, large = dims
    if large / small > 8:
        return None  # too elongated
    history = features.get('feature_history') or []
    machining_ops = {'ExtrudeFeature', 'CutFeature', 'FilletFeature',
                     'ChamferFeature', 'HoleFeature', 'PocketFeature',
                     'MillingFeature'}
    has_machining = any(f in machining_ops for f in history)
    confidence = 0.72 if has_machining else 0.60
    reason = 'Compact solid block' + (' with machining features' if has_machining else '')
    return PartType.MILLED_BLOCK, confidence, reason


def _rule_angle_section(features: dict):
    """
    L-shaped cross-section:
    - Elongated
    - Roughly solid fill (solid metal in L shape)
    - Two dominant perpendicular planar face groups
    - Total face count in range typical for L-section (6–12)
    """
    if _has_cylindrical_face(features):
        return None
    if not _is_elongated(features, ratio_threshold=2.5):
        return None
    total_faces = features.get('total_faces') or 0
    # An L-section extruded has 6 planar faces minimum; with chamfers more
    if not (5 <= total_faces <= 20):
        return None
    ccs = features.get('has_constant_cross_section')
    fr = _fill_ratio(features)
    if fr is not None and fr < 0.40:
        return None  # too hollow
    planar_ratio = _planar_face_ratio(features)
    if planar_ratio < 0.80:
        return None
    confidence = 0.68 if ccs else 0.55
    reason = 'Elongated L-section geometry' + (' with constant cross-section' if ccs else '')
    return PartType.ANGLE_SECTION, confidence, reason


def _rule_c_channel(features: dict):
    """
    C/U-shaped cross-section:
    - Elongated, hollow, planar faces
    - More faces than a rectangular tube (open section adds extra faces)
    - Fill ratio between tube and solid bar
    """
    if _has_cylindrical_face(features):
        return None
    if not _is_elongated(features, ratio_threshold=2.5):
        return None
    total_faces = features.get('total_faces') or 0
    # C-channel: 3 webs + 2 flanges + 2 ends = 7 min; more with fillets
    if not (6 <= total_faces <= 24):
        return None
    planar_ratio = _planar_face_ratio(features)
    if planar_ratio < 0.80:
        return None
    ccs = features.get('has_constant_cross_section')
    fr = _fill_ratio(features)
    # C-channel fill ratio is typically between flat bar and tube
    if fr is not None and (fr < 0.25 or fr > 0.88):
        return None
    confidence = 0.65 if ccs else 0.52
    reason = 'Elongated open-section channel geometry' + (' with constant cross-section' if ccs else '')
    return PartType.C_CHANNEL, confidence, reason


def _rule_aluminium_extrusion_geometry(features: dict):
    """
    Complex constant cross-section extrusion not caught by simpler rules:
    - Elongated
    - Constant cross-section
    - Higher face count (T-slot profiles are complex)
    """
    if not _is_elongated(features, ratio_threshold=2.5):
        return None
    ccs = features.get('has_constant_cross_section')
    if not ccs:
        return None
    total_faces = features.get('total_faces') or 0
    if total_faces < 12:
        return None  # simple enough that another rule should have matched
    return PartType.ALUMINIUM_EXTRUSION, 0.62, 'Complex constant cross-section extrusion'


def _rule_fallback(_features: dict):
    """Always matches — provides UNKNOWN with zero confidence."""
    return PartType.UNKNOWN, 0.0, 'No rule matched'


# Priority-ordered rule chain
_RULES = [
    _rule_sheet_metal,
    _rule_name_keywords,
    _rule_fastener_geometry,
    _rule_round_bar,
    _rule_round_tube,
    _rule_rectangular_tube,
    _rule_flat_bar,
    _rule_milled_block,
    _rule_angle_section,
    _rule_c_channel,
    _rule_aluminium_extrusion_geometry,
    _rule_fallback,
]


# ---------------------------------------------------------------------------
# Per-body classification
# ---------------------------------------------------------------------------

def _classify_one(features: dict) -> tuple[str, float, str]:
    """Run the rule chain and return the first match."""
    for rule in _RULES:
        result = rule(features)
        if result is not None:
            return result
    # _rule_fallback always returns, so this line is unreachable
    return PartType.UNKNOWN, 0.0, 'No rule matched'  # pragma: no cover


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def classify_bodies(features: list[dict]) -> list[dict]:
    """
    Classify a list of feature dicts produced by feature_extraction.extract_features().

    Returns a new list of dicts, each containing all original fields plus:
        classified_type       : str
        confidence            : float  (0.0 – 1.0)
        classification_reason : str
        needs_review          : bool
    """
    results = []
    for feat in features:
        out = dict(feat)

        # Bodies that errored during extraction get flagged immediately
        if 'extraction_error' in feat:
            out.update({
                'classified_type':       PartType.UNKNOWN,
                'confidence':            0.0,
                'classification_reason': 'Skipped — feature extraction failed',
                'needs_review':          True,
            })
            results.append(out)
            continue

        part_type, confidence, reason = _classify_one(feat)
        out['classified_type']       = part_type
        out['confidence']            = round(confidence, 4)
        out['classification_reason'] = reason
        out['needs_review']          = confidence < CONFIDENCE_THRESHOLD
        results.append(out)

    return results
