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

import logging
import math
import re
from typing import Optional


# ---------------------------------------------------------------------------
# Part type constants
# ---------------------------------------------------------------------------

class PartType:
    HOLLOW_RECTANGULAR_CHANNEL = 'HollowRectangularChannel'
    HOLLOW_CIRCULAR_CYLINDER   = 'HollowCircularCylinder'
    SOLID_CYLINDER             = 'SolidCylinder'
    SOLID_BLOCK                = 'SolidBlock'
    SHEET_METAL                = 'SheetMetal'
    THREE_D_PRINTED            = '3DPrinted'
    FASTENER                   = 'Fastener'
    SOURCED_COMPONENT          = 'SourcedComponent'
    UNKNOWN                    = 'Unknown'


CONFIDENCE_THRESHOLD = 0.6
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Plastic material detection
# ---------------------------------------------------------------------------

_PLASTIC_MATERIALS = {
    'pla', 'abs', 'petg', 'nylon', 'asa', 'tpu', 'resin',
    'pc', 'peek', 'pei', 'ultem', 'polycarbonate', 'polyamide',
    'polylactic', 'acrylonitrile', 'polypropylene', 'polyethylene',
    'plastic', 'fdm', 'sla', 'sls',
}


def _is_plastic_material(features: dict) -> bool:
    """Return True if the material name contains a known plastic keyword."""
    name = (features.get('material_name') or '').lower()
    return any(plastic in name for plastic in _PLASTIC_MATERIALS)


# ---------------------------------------------------------------------------
# Fastener detection
# ---------------------------------------------------------------------------

_FASTENER_NAME_KEYWORDS = {
    'bolt', 'screw', 'nut', 'washer', 'rivet',
    'hex cap', 'socket head', 'flat head', 'pan head', 'button head',
    'lock nut', 'flange nut', 'wing nut', 'acorn nut',
    'cap screw', 'set screw', 'machine screw', 'self tapping',
}

_FASTENER_NAME_RE = re.compile(
    r'\b(?:' + '|'.join(re.escape(k) for k in _FASTENER_NAME_KEYWORDS) + r')\b'
    r'|'
    r'\b[Mm]\d{1,2}(?:\.\d)?\b',   # M2, M3, M4, M5, M6, M8, M10, M12, M16, M20
    re.I,
)

_MAX_FASTENER_DIM_MM = 200.0  # fasteners are typically < 200mm


def _is_fastener_name(features: dict) -> Optional[str]:
    """Return the matched keyword if fastener name detected, else None.

    Checks component_path in addition to component_name and body_name because
    Fusion 360 content-library parts often have descriptive names in their
    occurrence hierarchy (e.g. "M8 Socket Head Cap Screw:1") while the body
    itself may be named generically ("Body" or "Solid1").
    """
    name = ' '.join(filter(None, [
        features.get('component_path', ''),
        features.get('component_name', ''),
        features.get('body_name', ''),
    ]))
    m = _FASTENER_NAME_RE.search(name)
    return m.group(0) if m else None


def _is_fastener_geometry(features: dict) -> bool:
    """Return True if geometry looks like a fastener (small, solid, cylindrical).

    Thresholds are deliberately loose: nuts and washers are not elongated,
    and many fasteners have mixed planar+cylindrical faces (hex heads, etc.).
    """
    dims = _sorted_dims(features)
    if dims is None:
        return False
    small, mid, large = dims
    if large > _MAX_FASTENER_DIM_MM:
        return False
    fr = _fill_ratio(features)
    if fr is not None and fr < 0.35:
        return False  # very hollow → not a solid fastener
    cyl_ratio = _cylindrical_face_ratio(features)
    if cyl_ratio < 0.15:
        return False  # no significant cylindrical faces
    # Allow non-elongated fasteners (nuts, washers, short bolts):
    # only reject if it's flat/sheet-like (large/small > 10 with low cyl_ratio)
    if mid > 0 and large / mid > 8.0 and cyl_ratio < 0.25:
        return False
    return True


# ---------------------------------------------------------------------------
# Name-keyword lookup tables
# (pattern -> (type, confidence))
# ---------------------------------------------------------------------------

# Each entry: (compiled_regex, PartType constant, confidence)
_NAME_KEYWORDS: list[tuple] = [
    # Hollow rectangular sections
    (re.compile(r'\b(rhs|shs|rectangular.?tube|square.?tube|box.?section|hollow.?section|hollow.?rectangular)\b', re.I),
     PartType.HOLLOW_RECTANGULAR_CHANNEL, 0.85),

    # Hollow circular sections
    (re.compile(r'\b(chs|round.?tube|circular.?tube|pipe|hollow.?bar|hollow.?circle|round.?pipe)\b', re.I),
     PartType.HOLLOW_CIRCULAR_CYLINDER, 0.85),

    # Solid cylinders / round bars
    (re.compile(r'\b(round.?bar|rod|shaft|axle|spindle|solid.?cylinder|solid.?round)\b', re.I),
     PartType.SOLID_CYLINDER, 0.80),

    # 3D printing indicators
    (re.compile(r'\b(3d.?print|printed|fdm|sla|sls|fff|filament)\b', re.I),
     PartType.THREE_D_PRINTED, 0.85),
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


def _sorted_dims(features: dict) -> Optional[tuple]:
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


def _is_elongated(features: dict, ratio_threshold: float = 3.0) -> bool:
    """True when the largest dimension is at least ratio_threshold × the smallest."""
    dims = _sorted_dims(features)
    if dims is None:
        return False
    small, _, large = dims
    return large / small >= ratio_threshold if small > 0 else False


def _is_hollow(features: dict, fill_threshold: float = 0.82) -> bool:
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


def _rule_3d_printed(features: dict):
    """Plastic material detected → 3D printed part."""
    if _is_plastic_material(features):
        mat = features.get('material_name', '')
        return PartType.THREE_D_PRINTED, 0.90, 'Plastic material detected: "{}"'.format(mat)
    return None


def _rule_content_library_fastener(features: dict):
    """Content-library part flagged by isLibraryItem — definitive."""
    if features.get('is_content_library_fastener'):
        return PartType.FASTENER, 0.98, 'Content-library referenced component'
    return None


def _rule_sourced_component(features: dict):
    """Externally referenced (linked) component — NOT a content-library fastener."""
    if features.get('is_sourced_component'):
        return PartType.SOURCED_COMPONENT, 0.95, 'Externally referenced component'
    return None


def _rule_fastener(features: dict):
    """Detect fasteners by name keywords, optionally reinforced by geometry."""
    kw = _is_fastener_name(features)
    if not kw:
        return None
    geo = _is_fastener_geometry(features)
    if geo:
        return PartType.FASTENER, 0.92, 'Fastener name "{}" + geometry match'.format(kw)
    return PartType.FASTENER, 0.85, 'Fastener name keyword: "{}"'.format(kw)


def _rule_hollow_circular_cylinder(features: dict):
    """Hollow cylinder: cylindrical faces + hollow fill ratio + elongated."""
    if not _has_cylindrical_face(features):
        return None
    if not _is_elongated(features, ratio_threshold=2.0):
        return None
    if not _is_hollow(features, fill_threshold=0.82):
        return None
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, _ = dims
    # Roughly circular cross-section
    if mid > 0 and small / mid < 0.70:
        return None
    # Predominantly planar faces → rectangular section, not circular
    # Use 0.70 threshold to allow holes in circular tubes (holes add planar faces)
    planar_ratio = _planar_face_ratio(features)
    if planar_ratio > 0.70:
        return None
    cyl_ratio = _cylindrical_face_ratio(features)
    confidence = 0.82 if cyl_ratio >= 0.40 else 0.67
    return PartType.HOLLOW_CIRCULAR_CYLINDER, confidence, 'Hollow elongated cylinder'


def _rule_solid_cylinder(features: dict):
    """Solid cylinder: high cylindrical face ratio + elongated + solid fill."""
    if not _is_elongated(features, ratio_threshold=2.0):
        return None
    cyl_ratio = _cylindrical_face_ratio(features)
    if cyl_ratio < 0.40:
        return None
    fr = _fill_ratio(features)
    if fr is not None and fr < 0.70:
        return None  # likely hollow — caught by hollow rule above
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, _ = dims
    # Cross-section should be roughly circular (small ≈ mid)
    if mid > 0 and small / mid < 0.70:
        return None
    confidence = 0.78 if (fr is not None and fr >= 0.85) else 0.65
    return PartType.SOLID_CYLINDER, confidence, 'Solid elongated cylinder'


def _rule_hollow_rectangular_channel(features: dict):
    """
    Hollow rectangular/square section:
    - Predominantly planar faces (small proportion of cylindrical faces
      from holes is tolerated)
    - Elongated
    - Hollow fill ratio
    """
    cyl_ratio = _cylindrical_face_ratio(features)
    if cyl_ratio >= 0.40:
        return None
    if not _is_elongated(features, ratio_threshold=1.5):
        return None
    if not _is_hollow(features, fill_threshold=0.82):
        return None
    planar_ratio = _planar_face_ratio(features)
    if planar_ratio < 0.70:
        return None
    dims = _sorted_dims(features)
    if dims is None:
        return None
    small, mid, _ = dims
    # Cross-section must not be too thin (a very thin hollow flat object is unusual)
    if mid > 0 and small / mid < 0.20:
        return None
    ccs = features.get('has_constant_cross_section')
    confidence = 0.83 if ccs else 0.68
    reason = 'Hollow rectangular section' + (' with constant cross-section' if ccs else '')
    if cyl_ratio > 0:
        confidence = max(confidence - 0.05, 0.55)
        reason += ' (with holes)'
    return PartType.HOLLOW_RECTANGULAR_CHANNEL, confidence, reason


def _rule_solid_block(features: dict):
    """
    Catch-all for non-cylinder, non-hollow solids: milled blocks, flat bars,
    angle sections, extrusions, and any other planar-face solid body.
    """
    # Skip if predominantly cylindrical (should have been caught above)
    if _has_cylindrical_face(features) and _cylindrical_face_ratio(features) > 0.5:
        return None
    fr = _fill_ratio(features)
    if fr is not None and fr < 0.70:
        return None  # too hollow to be a solid block
    planar_ratio = _planar_face_ratio(features)
    if planar_ratio < 0.60:
        return None  # mostly curved surfaces
    ccs = features.get('has_constant_cross_section')
    history = features.get('feature_history') or []
    machining_ops = {'ExtrudeFeature', 'CutFeature', 'FilletFeature',
                     'ChamferFeature', 'HoleFeature', 'PocketFeature'}
    has_machining = any(f in machining_ops for f in history)
    if ccs and fr is not None and fr >= 0.75:
        confidence = 0.78
    elif has_machining:
        confidence = 0.72
    elif fr is not None and fr >= 0.70:
        confidence = 0.65
    else:
        confidence = 0.52
    reason = 'Solid planar body (raw material / milled block)'
    if has_machining:
        reason += ' with machining features'
    return PartType.SOLID_BLOCK, confidence, reason


def _rule_name_keywords(features: dict):
    """Match component or body name against keyword groups."""
    name = ' '.join(filter(None, [
        features.get('component_name', ''),
        features.get('body_name', ''),
    ]))
    for pattern, part_type, confidence in _NAME_KEYWORDS:
        if pattern.search(name):
            match = pattern.search(name).group(0)
            return part_type, confidence, 'Name keyword match: "{}"'.format(match)
    return None


def _rule_fallback(_features: dict):
    """Always matches — provides UNKNOWN with zero confidence."""
    return PartType.UNKNOWN, 0.0, 'No rule matched'


# Priority-ordered rule chain.
# Sheet metal first (definitive Fusion API flag).
# Content-library fastener second — the isLibraryItem flag is definitive and
# must run before _rule_3d_printed so that nylon/plastic fasteners from the
# content library are not misclassified as 3D-printed parts.
# Name keywords run before geometry rules so explicitly named parts
# (e.g. "RHS 100x50") are classified correctly even if geometry is ambiguous.
_RULES = [
    _rule_sheet_metal,
    _rule_content_library_fastener,
    _rule_sourced_component,
    _rule_3d_printed,
    _rule_fastener,
    _rule_name_keywords,
    _rule_hollow_circular_cylinder,
    _rule_solid_cylinder,
    _rule_hollow_rectangular_channel,
    _rule_solid_block,
    _rule_fallback,
]


# ---------------------------------------------------------------------------
# Per-body classification
# ---------------------------------------------------------------------------

def _classify_one(features: dict) -> tuple:
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

def classify_bodies(
    features: list,
    confidence_threshold: float = CONFIDENCE_THRESHOLD,
) -> list:
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
        out['needs_review']          = confidence < confidence_threshold
        logger.info(
            'Classification: part="%s" body="%s" type="%s" confidence=%.4f review=%s reason="%s"',
            feat.get('component_name', ''),
            feat.get('body_name', ''),
            part_type,
            out['confidence'],
            out['needs_review'],
            reason,
        )
        results.append(out)

    return results
