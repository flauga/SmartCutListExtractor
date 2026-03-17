"""
review_palette.py — Cut list review UI via a Fusion 360 HTML Palette.

Lifecycle
---------
    start(classified_features)   — create palette, group bodies, send data
    stop()                       — destroy palette, release handlers

Messaging protocol (HTML ↔ Python)
-----------------------------------
HTML → Python  :  adsk.fusionSendData(action, JSON.stringify(data))
                  args.action = action string
                  args.data   = JSON payload string

Python → HTML  :  palette.sendInfoToHTML(action, json_string)
                  window.fusionJavaScriptHandler.handle(action, json_string)

Actions the HTML sends to Python
---------------------------------
  'htmlReady'           {}                                → trigger initial data push
  'override_type'       {group_id, value}                → reclassify a group
  'override_material'   {group_id, value}                → rename material
  'toggle_include'      {group_id, value: bool}          → include/exclude from export
  'bulk_include'        {group_ids: [...], value: bool}  → bulk include/exclude
  'filter_by_type'      {type: str, include: bool}       → include/exclude all of type
  'export_csv'          {settings}                       → write CSV to disk
  'export_json'         {settings}                       → write JSON to disk
  'export_dxf'          {settings}                       → write DXF files to disk
  'export_fasteners'    {settings}                       → write fastener CSV to disk
  'export_step'         {settings}                       → write STEP files for 3D prints
  'export_sourced'      {settings}                       → write sourced components CSV
  'weld_reorder_steps'  {weld_id, step_order: [...]}     → reorder steps in a weld assembly
  'weld_add_body'       {weld_id, group_id, body_index}  → add a body to a weld assembly
  'weld_remove_body'    {weld_id, step_index}            → remove a step from a weld assembly
  'weld_create_assembly' {name, group_ids}               → create a new manual weld assembly
  'weld_delete_assembly' {weld_id}                       → delete a weld assembly
  'weld_update_description' {weld_id, step_index, desc}  → update step description
  'weld_update_notes'   {weld_id, notes}                 → update assembly-level notes
  'weld_nest_assembly'  {child_weld_id, parent_weld_id}  → nest a weld inside another
  'weld_unnest_assembly' {child_weld_id}                 → promote a nested weld to top-level
  'generate_weld_plan'  {settings}                       → generate weld plan document
  'highlight_weld_step' {weld_id, step_index}            → highlight step body in viewport
  'weld_preview_step'   {weld_id, step_index}            → preview cumulative weld progress

Actions Python sends to HTML
-----------------------------
  'initData'            {groups: [...], part_types: [...], weld_assemblies: [...], weld_plans: [...]}
  'updateWeldPlans'     [...]                             → push updated weld plan state
"""

from __future__ import annotations

import json
import logging
import os
import re
import traceback
from datetime import datetime
from typing import Optional

import adsk.core
import adsk.fusion

from .classifier import PartType
from . import export_cutlist
from . import export_dxf
from . import hole_detection
from . import weld_plan_generator


# ---------------------------------------------------------------------------
# Module-level globals
# ---------------------------------------------------------------------------

_app:     adsk.core.Application   = None
_ui:      adsk.core.UserInterface = None
_palette: adsk.core.Palette       = None
_handlers: list                   = []
_groups:   list                   = []   # current grouped cut list data
_weld_plans: list                 = []   # enhanced weld assembly data with ordering
_hole_summaries: list             = []   # hole data for channel bodies
_palette_settings: dict           = {}
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PALETTE_ID   = 'SmartCutListReviewPalette'
PALETTE_NAME = 'Smart Cut List - Review'

# HTML file sits in SmartCutList/resources/ relative to this commands/ directory
_HTML_PATH = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)),
                 '..', 'resources', 'review_palette.html')
)
_RUNTIME_HTML_PATH = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)),
                 '..', 'resources', 'review_palette_runtime.html')
)

# Types that support DXF flat-pattern export
_DXF_TYPES = {PartType.SHEET_METAL}

# Dimension grouping tolerance in mm (round to nearest value)
_DEFAULT_DIM_TOLERANCE = 0.1


# ---------------------------------------------------------------------------
# Part grouping
# ---------------------------------------------------------------------------

def _round_dim(value: float, tolerance_mm: float) -> float:
    """Round a single dimension to the nearest tolerance for grouping."""
    if tolerance_mm <= 0:
        return round(float(value), 4)
    return round(round(value / tolerance_mm) * tolerance_mm, 4)


def _dims_key(bb_mm, tolerance_mm: float) -> Optional[tuple]:
    """Return a hashable, sorted, tolerance-snapped dimension tuple, or None."""
    if not bb_mm or len(bb_mm) < 3:
        return None
    try:
        return tuple(_round_dim(float(d), tolerance_mm) for d in sorted(bb_mm))
    except (TypeError, ValueError):
        return None


# Regex that matches Fusion 360's default body name pattern: "Body", "Body1", "Body12", etc.
_DEFAULT_BODY_RE = re.compile(r'^Body\d*$', re.IGNORECASE)


def _pick_display_name(feat: dict) -> str:
    """Choose the most informative name for a group's display_name.

    Content-library fasteners have generic body names ("Body", "Body1")
    but their component_path / component_name carries the full description
    (e.g. "Broached Hexagon Socket Head Cap Screw M10x1.5x50:1").

    For standard parts whose body name is still a Fusion default
    ("Body", "Body1", etc.) the component name is prepended so filenames
    and palette rows are meaningful (e.g. "Frame_Tube_Body1").
    """
    if feat.get('is_content_library_fastener') or feat.get('is_sourced_component'):
        # External components have generic body names ("Body", "Body1")
        # but their component_path carries the full description.
        # Strip the trailing ":N" occurrence index for cleaner display.
        path = feat.get('component_path', '')
        if path:
            leaf = path.rsplit('/', 1)[-1]
            clean = re.sub(r':\d+$', '', leaf)
            if clean:
                return clean
        comp = feat.get('component_name', '')
        if comp:
            return comp
    # Standard parts: prefer body_name, fall back to component_name.
    # If the body carries a Fusion default name ("Body1") prefix it with
    # the component name so it is meaningful in the cut list & filenames.
    body_name = feat.get('body_name', '')
    comp_name = feat.get('component_name', '')
    if not body_name:
        return comp_name or 'Unknown'
    if _DEFAULT_BODY_RE.match(body_name.strip()):
        prefix = comp_name.strip()
        return '{0}_{1}'.format(prefix, body_name.strip()) if prefix else body_name
    return body_name


def group_classified_bodies(
    classified_features: list,
    auto_group_identical: bool = True,
    dimension_tolerance_mm: float = _DEFAULT_DIM_TOLERANCE,
) -> list:
    """
    Group classified feature dicts by (classified_type, material, sorted dims,
    component_path).  Parts in different components are never merged even when
    their geometry and material are identical.

    Returns a list of group dicts.
    """
    groups: list  = []
    group_map: dict = {}   # key → index into groups

    for feat in classified_features:
        if 'extraction_error' in feat:
            # Error body → standalone group flagged for review
            groups.append({
                'group_id':             'g_{}'.format(len(groups)),
                'display_name':         feat.get('body_name') or '<unknown>',
                'component_name':       feat.get('component_name', ''),
                'component_path':       feat.get('component_path', ''),
                'classified_type':      PartType.UNKNOWN,
                'confidence':           0.0,
                'classification_reason': 'Feature extraction failed',
                'needs_review':         True,
                'material_name':        'Unknown',
                'dimensions_mm':        None,
                'quantity':             1,
                'include':              False,
                'bodies':               [_resolve_body_display_name(
                                            feat.get('body_name', ''),
                                            feat.get('component_name', ''))],
                'body_token':           feat.get('body_token', ''),
                'body_tokens':          [feat.get('body_token', '')] if feat.get('body_token') else [],
                'override_type':        None,
                'override_material':    None,
                'feature_history':      [],
                'is_sheet_metal':       False,
                'sheet_metal_configured': False,
                'sheet_metal_base_face_token': '',
                'sheet_metal_thickness_mm': None,
                'estimated_wall_thickness_mm': None,
            })
            continue

        ctype    = feat.get('classified_type', PartType.UNKNOWN)
        material = feat.get('material_name') or 'Unknown'
        bb       = feat.get('bounding_box_mm')
        dk       = _dims_key(bb, dimension_tolerance_mm)
        cpath    = feat.get('component_path', '')

        if auto_group_identical:
            # Scope grouping to same component path so identical parts in
            # different sub-assemblies remain separate rows.
            # Include face count and fill-ratio bucket to avoid merging
            # bodies that share dimensions but differ structurally.
            total_faces = feat.get('total_faces') or 0
            fill_bucket = round(
                round((feat.get('bb_fill_ratio') or 0.0) / 0.05) * 0.05, 2
            )
            key = (ctype, material, dk, cpath, total_faces, fill_bucket)
        else:
            key = ('body', cpath, feat.get('body_name', ''))

        if key in group_map:
            g = groups[group_map[key]]
            g['quantity'] += 1
            g['bodies'].append(_resolve_body_display_name(
                feat.get('body_name', ''), feat.get('component_name', '')))
            if feat.get('body_token'):
                g['body_tokens'].append(feat.get('body_token'))
        else:
            idx = len(groups)
            groups.append({
                'group_id':             'g_{}'.format(idx),
                'display_name':         _pick_display_name(feat),
                'component_name':       feat.get('component_name', ''),
                'component_path':       cpath,
                'classified_type':      ctype,
                'confidence':           feat.get('confidence', 0.0),
                'classification_reason': feat.get('classification_reason', ''),
                'needs_review':         feat.get('needs_review', False),
                'material_name':        material,
                'dimensions_mm':        sorted(bb) if bb else None,
                'quantity':             1,
                'include':              True,
                'bodies':               [_resolve_body_display_name(
                                            feat.get('body_name', ''),
                                            feat.get('component_name', ''))],
                'body_token':           feat.get('body_token', ''),
                'body_tokens':          [feat.get('body_token', '')] if feat.get('body_token') else [],
                'override_type':        None,
                'override_material':    None,
                'feature_history':      feat.get('feature_history', []),
                'is_sheet_metal':       bool(feat.get('is_sheet_metal', False)),
                'sheet_metal_configured': False,
                'sheet_metal_base_face_token': '',
                'sheet_metal_thickness_mm': feat.get('sheet_metal_thickness_mm'),
                'estimated_wall_thickness_mm': feat.get('estimated_wall_thickness_mm'),
            })
            group_map[key] = idx

    return groups


# ---------------------------------------------------------------------------
# Weld assembly detection
# ---------------------------------------------------------------------------

def _detect_weld_assemblies(groups: list) -> list:
    """
    Identify components containing multiple bodies (weld assemblies).

    Groups bodies by their parent component path.  When a component has more
    than one body it is treated as a welded assembly.
    """
    path_bodies: dict = {}  # parent_path → list of (group_id, body_names)
    for g in groups:
        path = g.get('component_path', '')
        if not path:
            continue
        path_bodies.setdefault(path, []).append(g)

    welds = []
    for path, grps in path_bodies.items():
        all_bodies = []
        for g in grps:
            all_bodies.extend(g.get('bodies', []))
        if len(all_bodies) > 1:
            welds.append({
                'component_path': path,
                'component_name': path.rsplit('/', 1)[-1] if '/' in path else path,
                'body_count': len(all_bodies),
                'body_names': list(all_bodies),
                'group_ids': [g['group_id'] for g in grps],
            })
    return welds


def _resolve_body_display_name(body_name: str, component_name: str) -> str:
    """Prefix default-named bodies (Body, Body1, …) with their parent component name.

    Fusion 360 assigns generic names like "Body1" when the user has not renamed a body.
    These are meaningless in isolation, so we prepend the component name to give context
    (e.g. "Frame_Tube_50x50_Body1").
    """
    stripped = (body_name or '').strip()
    if not stripped:
        return component_name or 'Unknown'
    if _DEFAULT_BODY_RE.match(stripped):
        prefix = (component_name or '').strip()
        return '{0}_{1}'.format(prefix, stripped) if prefix else stripped
    return stripped


def _build_weld_plan_data(groups: list) -> list:
    """
    Build enhanced weld assembly data with per-body steps and nesting.

    Each weld assembly gets an ordered ``steps`` list that defines the welding
    sequence.  Bodies are resolved from ``_groups`` so that each step carries
    the body token, material, and dimensions needed for viewport capture.
    """
    raw_welds = _detect_weld_assemblies(groups)
    weld_plans: list = []

    for i, raw in enumerate(raw_welds):
        steps: list = []
        for gid in raw['group_ids']:
            group = _find_group(gid)
            if not group:
                continue
            bodies = group.get('bodies', [])
            tokens = group.get('body_tokens', [])
            grp_component = group.get('component_name') or raw['component_name']
            for bi, bname in enumerate(bodies):
                token = tokens[bi] if bi < len(tokens) else ''
                steps.append({
                    'step_index': len(steps),
                    'body_name': _resolve_body_display_name(bname, grp_component),
                    'body_token': token,
                    'group_id': gid,
                    'body_index': bi,
                    'material_name': group.get('material_name', 'Unknown'),
                    'dimensions_mm': group.get('dimensions_mm'),
                    'is_sub_assembly': False,
                    'sub_assembly_weld_id': None,
                    'description': '',
                })

        weld_plans.append({
            'weld_id': 'w_{}'.format(i),
            'component_path': raw['component_path'],
            'component_name': raw['component_name'],
            'parent_weld_id': None,
            'children_weld_ids': [],
            'steps': steps,
            'notes': '',
            'is_manual': False,
            'user_modified': False,
        })

    _resolve_nesting(weld_plans)
    return weld_plans


def _resolve_nesting(weld_plans: list) -> None:
    """Detect parent/child relationships from component_path hierarchy.

    If ``path_a`` is a strict prefix of ``path_b`` (separated by ``/``),
    then weld B is a child of weld A.  A composite sub-assembly step is
    inserted into A's steps list so the user can position the entire nested
    assembly within the parent sequence.
    """
    by_path = {w['component_path']: w for w in weld_plans}

    for child in weld_plans:
        cpath = child['component_path']
        # Walk up the path looking for a parent weld
        parts = cpath.rsplit('/', 1)
        while len(parts) == 2:
            parent_path = parts[0]
            if parent_path in by_path:
                parent = by_path[parent_path]
                if parent['weld_id'] != child['weld_id']:
                    child['parent_weld_id'] = parent['weld_id']
                    if child['weld_id'] not in parent['children_weld_ids']:
                        parent['children_weld_ids'].append(child['weld_id'])
                    # Insert a composite step into the parent
                    already = any(
                        s.get('sub_assembly_weld_id') == child['weld_id']
                        for s in parent['steps']
                    )
                    if not already:
                        parent['steps'].append({
                            'step_index': len(parent['steps']),
                            'body_name': child['component_name'] + ' (sub-assembly)',
                            'body_token': '',
                            'group_id': '',
                            'material_name': '',
                            'dimensions_mm': None,
                            'is_sub_assembly': True,
                            'sub_assembly_weld_id': child['weld_id'],
                            'description': '',
                        })
                    break
            parts = parent_path.rsplit('/', 1)


def _find_weld(weld_id: str) -> Optional[dict]:
    """Look up a weld plan by ID."""
    for w in _weld_plans:
        if w['weld_id'] == weld_id:
            return w
    return None


def _push_weld_plans() -> None:
    """Send updated weld plan data to the HTML palette."""
    if _palette is None:
        return
    try:
        _palette.sendInfoToHTML(
            'updateWeldPlans', json.dumps(_weld_plans, default=str)
        )
    except Exception:
        logger.exception('Failed to push weld plan update to palette')


def _next_weld_id() -> str:
    """Generate the next available weld_id."""
    existing = {w['weld_id'] for w in _weld_plans}
    idx = len(_weld_plans)
    while 'w_{}'.format(idx) in existing:
        idx += 1
    return 'w_{}'.format(idx)


# ---------------------------------------------------------------------------
# Palette lifecycle
# ---------------------------------------------------------------------------

def start(classified_features: list, palette_settings: Optional[dict] = None) -> None:
    """Create the review palette and populate it with grouped body data."""
    global _app, _ui, _palette, _groups, _weld_plans, _hole_summaries, _handlers, _palette_settings

    try:
        _app = adsk.core.Application.get()
        _ui  = _app.userInterface
        _palette_settings = dict(palette_settings or {})

        _groups = group_classified_bodies(
            classified_features,
            auto_group_identical=bool(_palette_settings.get('auto_group_identical', True)),
            dimension_tolerance_mm=float(
                _palette_settings.get('dimension_tolerance_mm', _DEFAULT_DIM_TOLERANCE)
            ),
        )
        _weld_plans = _build_weld_plan_data(_groups)

        # Detect holes on channel bodies for the weld-plan drilling section
        _hole_summaries = hole_detection.detect_holes_on_channels(_groups)
        if _hole_summaries:
            logger.info('Detected holes on %d channel bodies', len(_hole_summaries))

        # Remove any stale palette from a previous invocation
        existing = _ui.palettes.itemById(PALETTE_ID)
        if existing:
            existing.deleteMe()

        html_url = 'file:///' + _build_runtime_html().replace('\\', '/')

        _palette = _ui.palettes.add(
            PALETTE_ID,
            PALETTE_NAME,
            html_url,
            True,    # isVisible
            True,    # showCloseButton
            True,    # isResizable
            960,     # width  (px)
            680,     # height (px)
        )
        _palette.dockingState = adsk.core.PaletteDockingStates.PaletteDockStateFloating

        on_html = HTMLEventHandler()
        _palette.incomingFromHTML.add(on_html)
        _handlers.append(on_html)
        logger.info('Review palette created; waiting for HTML handshake')

        # Some Fusion palette environments miss the first HTML -> Python event,
        # so push init data once optimistically as well.
        try:
            _send_init_data()
        except Exception:
            logger.exception('Initial optimistic palette data push failed')

    except Exception:
        if _ui:
            _ui.messageBox('review_palette start() error:\n{}'.format(
                traceback.format_exc()))


def stop() -> None:
    """Destroy the palette and release all handler references."""
    global _palette, _handlers, _groups, _weld_plans, _hole_summaries, _palette_settings
    try:
        if _palette:
            _palette.deleteMe()
            _palette = None
    except Exception:
        pass
    finally:
        _handlers = []
        _groups   = []
        _weld_plans = []
        _hole_summaries = []
        _palette_settings = {}


# ---------------------------------------------------------------------------
# Data → HTML
# ---------------------------------------------------------------------------

def _send_init_data() -> None:
    """Push the current group list and available part types to the palette."""
    if _palette is None:
        return
    try:
        logger.info('Sending %s grouped parts to review palette', len(_groups))
        payload = _palette_payload()
        _palette.sendInfoToHTML('initData', json.dumps(payload, default=str))
    except Exception:
        if _ui:
            _ui.messageBox('_send_init_data error:\n{}'.format(
                traceback.format_exc()))


def _palette_payload() -> dict:
    return {
        'groups': _groups,
        'settings': {
            'naming_template': _palette_settings.get(
                'default_naming_template',
                '{name}_{material}_{dims}',
            ),
            'unit': _palette_settings.get('default_units', 'mm'),
        },
        'linear_stock_summary': export_cutlist.build_linear_stock_summary(
            _groups,
            {'unit': _palette_settings.get('default_units', 'mm')},
        ),
        'part_types': [
            PartType.HOLLOW_RECTANGULAR_CHANNEL,
            PartType.HOLLOW_CIRCULAR_CYLINDER,
            PartType.SOLID_CYLINDER,
            PartType.SOLID_BLOCK,
            PartType.SHEET_METAL,
            PartType.THREE_D_PRINTED,
            PartType.FASTENER,
            PartType.SOURCED_COMPONENT,
            PartType.UNKNOWN,
        ],
        'weld_assemblies': _detect_weld_assemblies(_groups),
        'weld_plans': _weld_plans,
        'hole_summaries': _hole_summaries,
        'include_fasteners': bool(_palette_settings.get('include_fasteners', True)),
    }


def _build_runtime_html() -> str:
    with open(_HTML_PATH, 'r', encoding='utf-8') as handle:
        template = handle.read()

    bootstrap = '<script>window.__SMARTCUTLIST_INIT__ = {};</script>'.format(
        json.dumps(_palette_payload(), default=str)
    )
    runtime_html = template.replace('<!-- SMARTCUTLIST_BOOTSTRAP -->', bootstrap, 1)

    with open(_RUNTIME_HTML_PATH, 'w', encoding='utf-8') as handle:
        handle.write(runtime_html)

    logger.info('Wrote runtime review palette HTML to %s', _RUNTIME_HTML_PATH)
    return _RUNTIME_HTML_PATH


# ---------------------------------------------------------------------------
# Message handlers (HTML → Python)
# ---------------------------------------------------------------------------

def _find_group(group_id: str) -> Optional[dict]:
    for g in _groups:
        if g['group_id'] == group_id:
            return g
    return None


def _handle_html_ready(_data: dict) -> None:
    logger.info('Review palette HTML is ready; sending init data for %s groups', len(_groups))
    _send_init_data()


def _handle_override_type(data: dict) -> None:
    g = _find_group(data.get('group_id', ''))
    if g:
        g['classified_type'] = data.get('value', PartType.UNKNOWN)
        g['override_type']   = g['classified_type']
        if g['classified_type'] != PartType.SHEET_METAL:
            g['sheet_metal_configured'] = False
            g['sheet_metal_base_face_token'] = ''
            g['sheet_metal_thickness_mm'] = None


def _handle_override_material(data: dict) -> None:
    g = _find_group(data.get('group_id', ''))
    if g:
        g['material_name']    = data.get('value', 'Unknown')
        g['override_material'] = g['material_name']


def _handle_toggle_include(data: dict) -> None:
    g = _find_group(data.get('group_id', ''))
    if g:
        g['include'] = bool(data.get('value', True))


def _handle_bulk_include(data: dict) -> None:
    """Set include flag on multiple groups at once."""
    group_ids = data.get('group_ids') or []
    value = bool(data.get('value', True))
    for gid in group_ids:
        g = _find_group(gid)
        if g:
            g['include'] = value


def _handle_filter_by_type(data: dict) -> None:
    """Set include flag on all groups matching a given part type."""
    part_type = data.get('type', '')
    value = bool(data.get('include', True))
    for g in _groups:
        if g.get('classified_type') == part_type:
            g['include'] = value


def _handle_export_csv(data: dict) -> None:
    _export_csv(data.get('settings', {}))


def _handle_export_json(data: dict) -> None:
    _export_json(data.get('settings', {}))


def _handle_highlight_group(data: dict) -> None:
    group = _find_group(data.get('group_id', ''))
    if not group:
        return

    try:
        _highlight_group(group)
    except Exception:
        logger.exception('Failed to highlight group %s', group.get('group_id'))


def _handle_highlight_weld(data: dict) -> None:
    """Highlight all bodies in a weld assembly by iterating its group_ids."""
    group_ids = data.get('group_ids') or []
    if not group_ids:
        return

    try:
        _ui.activeSelections.clear()
    except Exception:
        pass

    highlighted = 0
    for gid in group_ids:
        group = _find_group(gid)
        if not group:
            continue
        for token in group.get('body_tokens', []) or []:
            body = export_dxf.resolve_body_token(token)
            if body is None:
                continue
            try:
                _ui.activeSelections.add(body)
                highlighted += 1
            except Exception:
                continue

    if highlighted:
        try:
            _app.activeViewport.refresh()
        except Exception:
            pass


def _handle_configure_sheet_metal(data: dict) -> None:
    group = _find_group(data.get('group_id', ''))
    if not group:
        return

    if group.get('classified_type') != PartType.SHEET_METAL:
        _ui.messageBox('Only sheet metal groups need flat-pattern configuration.')
        return

    saved_visibility = []  # type: list
    try:
        _highlight_group(group)
        body = export_dxf.resolve_body_for_part(group)
        if body is None:
            raise RuntimeError('Unable to resolve the selected Fusion body')

        # ── Isolate the target body so the user can see only it ──────
        design = adsk.fusion.Design.cast(_app.activeProduct) if _app else None
        if design:
            for comp in design.allComponents:
                for bi in range(comp.bRepBodies.count):
                    b = comp.bRepBodies.item(bi)
                    saved_visibility.append((b, b.isVisible))
                    b.isVisible = (b.entityToken == body.entityToken)
            try:
                _app.activeViewport.refresh()
                _app.activeViewport.fit()
            except Exception:
                pass

        palette_was_visible = bool(_palette and _palette.isVisible)
        if _palette:
            _palette.isVisible = False
        try:
            selection = _ui.selectEntity(
                'Select the base face for "{}" in the Fusion canvas'.format(
                    group.get('display_name', 'sheet metal part')
                ),
                'PlanarFaces',
            )
        finally:
            if _palette:
                _palette.isVisible = palette_was_visible
        if not selection:
            return

        face = adsk.fusion.BRepFace.cast(selection.entity)
        if face is None:
            raise RuntimeError('Please select a planar face on the chosen body')

        export_dxf.validate_sheet_metal_configuration(group, face)
        group['sheet_metal_base_face_token'] = face.entityToken
        group['sheet_metal_configured'] = True
        group['sheet_metal_thickness_mm'] = export_dxf.resolve_sheet_metal_thickness_mm(group, face)
        logger.info(
            'Configured sheet metal flat-pattern face for group %s (%s)',
            group.get('group_id'),
            group.get('display_name'),
        )
        _send_init_data()
    except Exception:
        logger.exception(
            'Sheet metal configuration failed for group %s',
            group.get('group_id'),
        )
        _ui.messageBox(
            'Sheet metal configuration failed:\n{}'.format(traceback.format_exc())
        )
    finally:
        # ── Restore original body visibility ─────────────────────────
        for b, was_visible in saved_visibility:
            try:
                b.isVisible = was_visible
            except Exception:
                pass
        if saved_visibility:
            try:
                _app.activeViewport.refresh()
            except Exception:
                pass


def _handle_export_fasteners(data: dict) -> None:
    _export_fasteners(data.get('settings', {}))


def _handle_save_template(data: dict) -> None:
    """Persist the naming template from the palette back to settings.json."""
    from . import settings as settings_mod
    template = data.get('template', '').strip()
    if template:
        settings_mod.update_settings({'default_naming_template': template})
        _palette_settings['default_naming_template'] = template
        logger.info('Naming template saved: %s', template)


def _handle_export_step(data: dict) -> None:
    """Export STEP files for 3D printed parts."""
    try:
        folder_dlg = _ui.createFolderDialog()
        folder_dlg.title = 'Select STEP Export Folder'
        if folder_dlg.showDialog() != adsk.core.DialogResults.DialogOK:
            return
        output_dir = folder_dlg.folder
        settings = data.get('settings', {})
        export_settings = _export_settings_from_ui(settings)
        step_paths = export_dxf.export_steps(
            _groups,
            output_dir,
            export_settings.filename_template,
            {'unit': export_settings.unit, 'show_summary': True},
        )
        if step_paths:
            _ui.messageBox(
                'Exported {} STEP file(s) to:\n{}'.format(len(step_paths), output_dir),
                'STEP Export Complete',
            )
        else:
            _ui.messageBox(
                'No 3D printed parts found to export.',
                'STEP Export',
            )
    except Exception:
        logger.exception('STEP export failed')
        _ui.messageBox('STEP export failed:\n{}'.format(traceback.format_exc()))


def _handle_export_sourced(data: dict) -> None:
    """Export sourced components CSV."""
    _export_sourced(data.get('settings', {}))


def _handle_export_package(data: dict) -> None:
    _export_package(data.get('settings', {}))


# ---------------------------------------------------------------------------
# Weld plan handlers
# ---------------------------------------------------------------------------

def _handle_weld_reorder_steps(data: dict) -> None:
    """Reorder steps in a weld assembly via a new index list."""
    weld = _find_weld(data.get('weld_id', ''))
    if not weld:
        return
    step_order = data.get('step_order')
    if not step_order or len(step_order) != len(weld['steps']):
        return
    try:
        reordered = [weld['steps'][i] for i in step_order]
        for idx, step in enumerate(reordered):
            step['step_index'] = idx
        weld['steps'] = reordered
        weld['user_modified'] = True
        _push_weld_plans()
    except (IndexError, TypeError):
        logger.exception('Invalid step_order for weld %s', data.get('weld_id'))


def _handle_weld_add_body(data: dict) -> None:
    """Add a body from a group to a weld assembly."""
    weld = _find_weld(data.get('weld_id', ''))
    if not weld:
        return
    group = _find_group(data.get('group_id', ''))
    if not group:
        return
    body_index = data.get('body_index', 0)
    bodies = group.get('bodies', [])
    tokens = group.get('body_tokens', [])
    if body_index < 0 or body_index >= len(bodies):
        return
    token = tokens[body_index] if body_index < len(tokens) else ''
    # Prevent adding a body that is already part of any weld assembly
    if token:
        for wp in _weld_plans:
            for s in wp.get('steps', []):
                if s.get('body_token') == token and not s.get('is_sub_assembly'):
                    return  # silently reject duplicate
    comp_name = group.get('component_name', '')
    weld['steps'].append({
        'step_index': len(weld['steps']),
        'body_name': _resolve_body_display_name(bodies[body_index], comp_name),
        'body_token': token,
        'group_id': group['group_id'],
        'body_index': body_index,
        'material_name': group.get('material_name', 'Unknown'),
        'dimensions_mm': group.get('dimensions_mm'),
        'is_sub_assembly': False,
        'sub_assembly_weld_id': None,
        'description': '',
    })
    weld['user_modified'] = True
    _push_weld_plans()


def _handle_weld_remove_body(data: dict) -> None:
    """Remove a step from a weld assembly."""
    weld = _find_weld(data.get('weld_id', ''))
    if not weld:
        return
    step_index = data.get('step_index', -1)
    if step_index < 0 or step_index >= len(weld['steps']):
        return
    if len(weld['steps']) <= 1:
        return  # don't allow removing the last step
    weld['steps'].pop(step_index)
    for idx, step in enumerate(weld['steps']):
        step['step_index'] = idx
    weld['user_modified'] = True
    _push_weld_plans()


def _handle_weld_create_assembly(data: dict) -> None:
    """Create a new manual weld assembly from selected groups."""
    name = data.get('name', '').strip()
    group_ids = data.get('group_ids') or []
    if not name or not group_ids:
        return
    steps: list = []
    for gid in group_ids:
        group = _find_group(gid)
        if not group:
            continue
        bodies = group.get('bodies', [])
        tokens = group.get('body_tokens', [])
        for bi, bname in enumerate(bodies):
            token = tokens[bi] if bi < len(tokens) else ''
            steps.append({
                'step_index': len(steps),
                'body_name': bname,
                'body_token': token,
                'group_id': gid,
                'material_name': group.get('material_name', 'Unknown'),
                'dimensions_mm': group.get('dimensions_mm'),
                'is_sub_assembly': False,
                'sub_assembly_weld_id': None,
                'description': '',
            })
    if not steps:
        return
    _weld_plans.append({
        'weld_id': _next_weld_id(),
        'component_path': '',
        'component_name': name,
        'parent_weld_id': None,
        'children_weld_ids': [],
        'steps': steps,
        'notes': '',
        'is_manual': True,
        'user_modified': True,
    })
    _push_weld_plans()


def _handle_weld_delete_assembly(data: dict) -> None:
    """Delete a weld assembly."""
    global _weld_plans
    weld_id = data.get('weld_id', '')
    weld = _find_weld(weld_id)
    if not weld:
        return
    # Remove from parent if nested
    if weld['parent_weld_id']:
        parent = _find_weld(weld['parent_weld_id'])
        if parent:
            parent['children_weld_ids'] = [
                c for c in parent['children_weld_ids'] if c != weld_id
            ]
            parent['steps'] = [
                s for s in parent['steps']
                if s.get('sub_assembly_weld_id') != weld_id
            ]
            for idx, step in enumerate(parent['steps']):
                step['step_index'] = idx
    _weld_plans = [w for w in _weld_plans if w['weld_id'] != weld_id]
    _push_weld_plans()


def _handle_weld_update_description(data: dict) -> None:
    """Update the description on a weld step."""
    weld = _find_weld(data.get('weld_id', ''))
    if not weld:
        return
    step_index = data.get('step_index', -1)
    if step_index < 0 or step_index >= len(weld['steps']):
        return
    weld['steps'][step_index]['description'] = data.get('description', '')
    weld['user_modified'] = True


def _handle_weld_update_notes(data: dict) -> None:
    """Update assembly-level notes."""
    weld = _find_weld(data.get('weld_id', ''))
    if not weld:
        return
    weld['notes'] = data.get('notes', '')


def _handle_weld_nest_assembly(data: dict) -> None:
    """Nest a weld assembly inside a parent."""
    child = _find_weld(data.get('child_weld_id', ''))
    parent = _find_weld(data.get('parent_weld_id', ''))
    if not child or not parent or child['weld_id'] == parent['weld_id']:
        return
    # Prevent circular nesting
    if parent['parent_weld_id'] == child['weld_id']:
        return
    # Remove from old parent
    if child['parent_weld_id']:
        old_parent = _find_weld(child['parent_weld_id'])
        if old_parent:
            old_parent['children_weld_ids'] = [
                c for c in old_parent['children_weld_ids']
                if c != child['weld_id']
            ]
            old_parent['steps'] = [
                s for s in old_parent['steps']
                if s.get('sub_assembly_weld_id') != child['weld_id']
            ]
            for idx, step in enumerate(old_parent['steps']):
                step['step_index'] = idx
    child['parent_weld_id'] = parent['weld_id']
    if child['weld_id'] not in parent['children_weld_ids']:
        parent['children_weld_ids'].append(child['weld_id'])
    # Add composite step
    already = any(
        s.get('sub_assembly_weld_id') == child['weld_id']
        for s in parent['steps']
    )
    if not already:
        parent['steps'].append({
            'step_index': len(parent['steps']),
            'body_name': child['component_name'] + ' (sub-assembly)',
            'body_token': '',
            'group_id': '',
            'material_name': '',
            'dimensions_mm': None,
            'is_sub_assembly': True,
            'sub_assembly_weld_id': child['weld_id'],
            'description': '',
        })
    _push_weld_plans()


def _handle_weld_unnest_assembly(data: dict) -> None:
    """Remove a weld assembly from its parent, promoting to top-level."""
    child = _find_weld(data.get('child_weld_id', ''))
    if not child or not child['parent_weld_id']:
        return
    parent = _find_weld(child['parent_weld_id'])
    if parent:
        parent['children_weld_ids'] = [
            c for c in parent['children_weld_ids']
            if c != child['weld_id']
        ]
        parent['steps'] = [
            s for s in parent['steps']
            if s.get('sub_assembly_weld_id') != child['weld_id']
        ]
        for idx, step in enumerate(parent['steps']):
            step['step_index'] = idx
    child['parent_weld_id'] = None
    _push_weld_plans()


def _handle_generate_weld_plan(data: dict) -> None:
    """Generate the weld plan document with progressive viewport screenshots."""
    if not _weld_plans:
        _ui.messageBox('No weld assemblies to generate a plan for.', 'Weld Plan')
        return
    try:
        folder_dlg = _ui.createFolderDialog()
        folder_dlg.title = 'Select Weld Plan Output Folder'
        if folder_dlg.showDialog() != adsk.core.DialogResults.DialogOK:
            return
        output_dir = folder_dlg.folder
        project_name = ''
        if _app and _app.activeDocument:
            project_name = _app.activeDocument.name
        settings = {
            'project_name': project_name or 'SmartCutList',
            'unit': _palette_settings.get('default_units', 'mm'),
        }
        result_path = weld_plan_generator.generate_weld_plan(
            _weld_plans, output_dir, settings, groups=_groups
        )
        _ui.messageBox(
            'Weld plan generated:\n{}'.format(result_path),
            'Weld Plan Complete',
        )
    except Exception:
        logger.exception('Weld plan generation failed')
        _ui.messageBox(
            'Weld plan generation failed:\n{}'.format(traceback.format_exc())
        )


def _handle_highlight_weld_step(data: dict) -> None:
    """Highlight a specific step's body in the viewport.

    When ``cumulative`` is True (Shift+click), highlights **all** steps
    from index 0 up to and including the requested step, so the user can
    see the progressive weld state.
    """
    weld = _find_weld(data.get('weld_id', ''))
    if not weld:
        return
    step_index = data.get('step_index', -1)
    if step_index < 0 or step_index >= len(weld['steps']):
        return
    cumulative = data.get('cumulative', False)
    try:
        _ui.activeSelections.clear()
    except Exception:
        pass

    # Decide which steps to highlight
    if cumulative:
        steps_to_highlight = weld['steps'][: step_index + 1]
    else:
        steps_to_highlight = [weld['steps'][step_index]]

    for step in steps_to_highlight:
        if step.get('is_sub_assembly'):
            sub = _find_weld(step.get('sub_assembly_weld_id', ''))
            if sub:
                for s in sub['steps']:
                    _select_body_token(s.get('body_token', ''))
        else:
            _select_body_token(step.get('body_token', ''))

    try:
        _app.activeViewport.refresh()
    except Exception:
        pass


def _select_body_token(token: str) -> None:
    """Resolve a body token and add it to the active selection."""
    if not token:
        return
    body = export_dxf.resolve_body_token(token)
    if body:
        try:
            _ui.activeSelections.add(body)
        except Exception:
            pass


def _handle_weld_preview_step(data: dict) -> None:
    """Preview cumulative weld progress up to a given step in the viewport."""
    weld = _find_weld(data.get('weld_id', ''))
    if not weld:
        return
    up_to = data.get('step_index', -1)
    if up_to < 0 or up_to >= len(weld['steps']):
        return
    try:
        weld_plan_generator.preview_weld_step(_weld_plans, weld['weld_id'], up_to)
    except Exception:
        logger.exception('Weld step preview failed')


def _handle_capture_hole_images(data: dict) -> None:
    """Capture perpendicular hole images for channel bodies and push to HTML."""
    if not _hole_summaries:
        return
    try:
        folder_dlg = _ui.createFolderDialog()
        folder_dlg.title = 'Temporary folder for hole images'
        # Use temp directory instead of asking user
        import tempfile
        output_dir = tempfile.mkdtemp(prefix='smartcutlist_holes_')
        weld_plan_generator.capture_hole_images(_hole_summaries, output_dir)
        # Push updated summaries (now with face_images) to HTML
        if _palette:
            _palette.sendInfoToHTML(
                'updateHoleImages',
                json.dumps(_hole_summaries, default=str),
            )
    except Exception:
        logger.exception('Hole image capture failed')
        if _ui:
            _ui.messageBox('Hole image capture failed:\n{}'.format(traceback.format_exc()))


def _handle_highlight_hole(data: dict) -> None:
    """Highlight hole faces in the viewport for a specific hole."""
    face_tokens = data.get('face_tokens', [])
    if not face_tokens:
        return
    try:
        app = adsk.core.Application.get()
        design = adsk.fusion.Design.cast(app.activeProduct)
        _ui.activeSelections.clear()
        for ft in face_tokens:
            if not ft:
                continue
            try:
                entities = design.findEntityByToken(ft)
                if entities:
                    _ui.activeSelections.add(entities[0])
            except Exception:
                pass
    except Exception:
        logger.debug('Highlight hole failed')


def _handle_export_drill_csv(data: dict) -> None:
    """Export a standalone drill-list CSV for channel holes."""
    if not _hole_summaries:
        _ui.messageBox('No holes detected on channel bodies.', 'Drill Export')
        return
    try:
        file_dlg = _ui.createFileDialog()
        file_dlg.filter = 'CSV Files (*.csv)'
        file_dlg.initialFilename = 'DrillList.csv'
        file_dlg.title = 'Save Drill List as CSV'
        if file_dlg.showSave() != adsk.core.DialogResults.DialogOK:
            return
        filepath = file_dlg.filename
        settings = data.get('settings', {})
        export_settings = _export_settings_from_ui(settings)
        result = export_cutlist.export_drill_csv(
            _hole_summaries, filepath, export_settings
        )
        _ui.messageBox(
            'Drill list saved to:\n{}'.format(result.filepath),
            'Export Complete',
        )
    except Exception:
        _ui.messageBox('Drill CSV export error:\n{}'.format(traceback.format_exc()))


def _handle_export_dxfs_only(data: dict) -> None:
    """Export DXF flat patterns only (no CSV)."""
    try:
        pending = _sheet_metal_groups_needing_configuration()
        if pending:
            names = ', '.join(g.get('display_name', 'Unknown') for g in pending[:5])
            if len(pending) > 5:
                names += ', ...'
            _ui.messageBox(
                'Configure all included sheet metal groups before exporting DXFs.\n\nPending: {}'.format(names),
                'Sheet Metal Setup Required',
            )
            return

        folder_dlg = _ui.createFolderDialog()
        folder_dlg.title = 'Select DXF Output Folder'
        if folder_dlg.showDialog() != adsk.core.DialogResults.DialogOK:
            return

        output_dir = folder_dlg.folder
        settings = data.get('settings', {})
        export_settings = _export_settings_from_ui(settings)
        dxf_paths = export_dxf.export_dxfs(
            _groups,
            output_dir,
            export_settings.filename_template,
            {
                'unit': export_settings.unit,
                'sheet_metal_only': True,
                'cross_section_for_profiles': False,
                'include_bend_lines': False,
                'include_dimensions': False,
                'show_summary': False,
            },
        )
        _ui.messageBox(
            'Exported {} DXF file(s) to:\n{}'.format(len(dxf_paths), output_dir),
            'DXF Export Complete',
        )
    except Exception:
        logger.exception('DXF-only export failed')
        _ui.messageBox('DXF export error:\n{}'.format(traceback.format_exc()))


def _handle_export_all(data: dict) -> None:
    """Export all outputs (CSVs, DXFs, STEPs, weld plan) into one folder."""
    try:
        folder_dlg = _ui.createFolderDialog()
        folder_dlg.title = 'Select Output Folder for All Exports'
        if folder_dlg.showDialog() != adsk.core.DialogResults.DialogOK:
            return

        output_dir = folder_dlg.folder
        settings = data.get('settings', {})
        export_settings = _export_settings_from_ui(settings)
        base_name = _sanitise_stem(export_settings.project_name or 'SmartCutList')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        exported: list = []

        # 1. Cut list CSV
        try:
            csv_path = os.path.join(output_dir, '{}_cut_list_{}.csv'.format(base_name, timestamp))
            export_cutlist.export_csv(
                _groups, csv_path, export_settings, hole_summaries=_hole_summaries
            )
            exported.append('Cut List CSV')
        except Exception:
            logger.exception('Export All: CSV failed')

        # 2. Drill CSV
        if _hole_summaries:
            try:
                drill_path = os.path.join(output_dir, '{}_drill_list_{}.csv'.format(base_name, timestamp))
                export_cutlist.export_drill_csv(_hole_summaries, drill_path, export_settings)
                exported.append('Drill List CSV')
            except Exception:
                logger.exception('Export All: Drill CSV failed')

        # 3. Fasteners CSV
        try:
            fast_path = os.path.join(output_dir, '{}_fasteners_{}.csv'.format(base_name, timestamp))
            export_cutlist.export_fasteners_csv(_groups, fast_path, export_settings)
            exported.append('Fasteners CSV')
        except Exception:
            logger.debug('Export All: Fasteners CSV skipped (no fasteners or error)')

        # 4. Sourced CSV
        try:
            sourced_path = os.path.join(output_dir, '{}_sourced_{}.csv'.format(base_name, timestamp))
            export_cutlist.export_sourced_csv(_groups, sourced_path, export_settings)
            exported.append('Sourced CSV')
        except Exception:
            logger.debug('Export All: Sourced CSV skipped')

        # 5. DXFs
        try:
            dxf_paths = export_dxf.export_dxfs(
                _groups, output_dir, export_settings.filename_template,
                {'unit': export_settings.unit, 'sheet_metal_only': True,
                 'cross_section_for_profiles': False, 'include_bend_lines': False,
                 'include_dimensions': False, 'show_summary': False},
            )
            if dxf_paths:
                exported.append('{} DXF file(s)'.format(len(dxf_paths)))
        except Exception:
            logger.exception('Export All: DXF export failed')

        # 6. STEPs
        try:
            step_paths = export_dxf.export_steps(
                _groups, output_dir, export_settings.filename_template,
                {'unit': export_settings.unit, 'show_summary': False},
            )
            if step_paths:
                exported.append('{} STEP file(s)'.format(len(step_paths)))
        except Exception:
            logger.debug('Export All: STEP export skipped')

        # 7. Weld plan
        if _weld_plans:
            try:
                project_name = ''
                if _app and _app.activeDocument:
                    project_name = _app.activeDocument.name
                weld_settings = {
                    'project_name': project_name or 'SmartCutList',
                    'unit': _palette_settings.get('default_units', 'mm'),
                }
                weld_plan_generator.generate_weld_plan(
                    _weld_plans, output_dir, weld_settings, groups=_groups
                )
                exported.append('Drill & Weld Plan')
            except Exception:
                logger.exception('Export All: Weld plan failed')

        summary = '\n'.join('  ✓ {}'.format(e) for e in exported)
        _ui.messageBox(
            'Export complete to:\n{}\n\n{}'.format(output_dir, summary),
            'Export All Complete',
        )
    except Exception:
        logger.exception('Export All failed')
        _ui.messageBox('Export All error:\n{}'.format(traceback.format_exc()))


_MESSAGE_HANDLERS = {
    'htmlReady':             _handle_html_ready,
    'override_type':         _handle_override_type,
    'override_material':     _handle_override_material,
    'toggle_include':        _handle_toggle_include,
    'bulk_include':          _handle_bulk_include,
    'filter_by_type':        _handle_filter_by_type,
    'export_csv':            _handle_export_csv,
    'export_json':           _handle_export_json,
    'highlight_group':       _handle_highlight_group,
    'highlight_weld':        _handle_highlight_weld,
    'configure_sheet_metal': _handle_configure_sheet_metal,
    'export_fasteners':      _handle_export_fasteners,
    'export_step':           _handle_export_step,
    'export_sourced':        _handle_export_sourced,
    'export_package':        _handle_export_package,
    'save_template':         _handle_save_template,
    # Weld plan actions
    'weld_reorder_steps':     _handle_weld_reorder_steps,
    'weld_add_body':          _handle_weld_add_body,
    'weld_remove_body':       _handle_weld_remove_body,
    'weld_create_assembly':   _handle_weld_create_assembly,
    'weld_delete_assembly':   _handle_weld_delete_assembly,
    'weld_update_description': _handle_weld_update_description,
    'weld_update_notes':      _handle_weld_update_notes,
    'weld_nest_assembly':     _handle_weld_nest_assembly,
    'weld_unnest_assembly':   _handle_weld_unnest_assembly,
    'generate_weld_plan':     _handle_generate_weld_plan,
    'highlight_weld_step':    _handle_highlight_weld_step,
    'weld_preview_step':      _handle_weld_preview_step,
    # Hole detection actions
    'capture_hole_images':    _handle_capture_hole_images,
    'highlight_hole':         _handle_highlight_hole,
    'export_drill_csv':       _handle_export_drill_csv,
    'export_dxfs_only':       _handle_export_dxfs_only,
    'export_all':             _handle_export_all,
}


# ---------------------------------------------------------------------------
# Event handler class
# ---------------------------------------------------------------------------

class HTMLEventHandler(adsk.core.HTMLEventHandler):
    """Routes all incoming palette messages to the appropriate handler."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.HTMLEventArgs):
        try:
            action = args.action
            data   = json.loads(args.data) if args.data else {}
            logger.info('Review palette received HTML action: %s', action)
            handler = _MESSAGE_HANDLERS.get(action)
            if handler:
                handler(data)
            # Unknown actions are silently ignored (future-proofing)
        except Exception:
            if _ui:
                _ui.messageBox('HTMLEventHandler error (action={}):\n{}'.format(
                    getattr(args, 'action', '?'), traceback.format_exc()))


# ---------------------------------------------------------------------------
# Filename / path helpers
# ---------------------------------------------------------------------------

def _sanitise(s: str) -> str:
    """Replace filesystem-unsafe characters with underscores."""
    return re.sub(r'[^\w\-.]', '_', s or '')


def _build_filename(group: dict, template: str, settings: dict) -> str:
    """Expand the naming template for a given group and unit setting."""
    unit       = settings.get('unit', 'mm')
    use_inches = unit == 'inches'
    factor     = 1 / 25.4 if use_inches else 1.0
    dims       = group.get('dimensions_mm')

    if dims and len(dims) >= 3:
        dims_str = 'x'.join('{:.1f}'.format(d * factor) for d in dims)
    else:
        dims_str = 'unknown'

    return (template
            .replace('{name}',     _sanitise(group.get('display_name', '')))
            .replace('{material}', _sanitise(group.get('material_name', '')))
            .replace('{dims}',     dims_str)
            .replace('{type}',     _sanitise(group.get('classified_type', '')))
            .replace('{qty}',      str(group.get('quantity', 1))))


def _included_groups() -> list:
    return [g for g in _groups if g.get('include', True)]


def _highlight_group(group: dict) -> None:
    try:
        _ui.activeSelections.clear()
    except Exception:
        pass

    highlighted = 0
    for token in group.get('body_tokens', []) or []:
        body = export_dxf.resolve_body_token(token)
        if body is None:
            continue
        try:
            _ui.activeSelections.add(body)
            highlighted += 1
        except Exception:
            continue

    if highlighted:
        try:
            _app.activeViewport.refresh()
        except Exception:
            pass


def _export_settings_from_ui(settings: dict) -> export_cutlist.ExportSettings:
    return export_cutlist.ExportSettings.from_mapping({
        'project_name': _app.activeDocument.name if _app and _app.activeDocument else 'SmartCutList',
        'unit': settings.get('unit', _palette_settings.get('default_units', 'mm')),
        'filename_template': settings.get(
            'naming_template',
            _palette_settings.get('default_naming_template', '{name}_{material}_{dims}')
        ),
    })


def _sanitise_stem(value: str) -> str:
    return _sanitise(os.path.splitext(value or 'SmartCutList')[0]) or 'SmartCutList'


# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------

def _export_csv(settings: dict) -> None:
    try:
        file_dlg = _ui.createFileDialog()
        file_dlg.filter           = 'CSV Files (*.csv)'
        file_dlg.initialFilename  = 'SmartCutList.csv'
        file_dlg.title            = 'Save Cut List as CSV'
        if file_dlg.showSave() != adsk.core.DialogResults.DialogOK:
            return

        filepath = file_dlg.filename
        export_settings = _export_settings_from_ui(settings)
        result = export_cutlist.export_csv(
            _groups, filepath, export_settings, hole_summaries=_hole_summaries
        )
        logger.info('Review palette CSV export: %s', result.filepath)
        _ui.messageBox('CSV saved to:\n{}'.format(result.filepath), 'Export Complete')

    except Exception:
        _ui.messageBox('CSV export error:\n{}'.format(traceback.format_exc()))


# ---------------------------------------------------------------------------
# JSON export
# ---------------------------------------------------------------------------

def _export_json(settings: dict) -> None:
    try:
        file_dlg = _ui.createFileDialog()
        file_dlg.filter          = 'JSON Files (*.json)'
        file_dlg.initialFilename = 'SmartCutList.json'
        file_dlg.title           = 'Save Cut List as JSON'
        if file_dlg.showSave() != adsk.core.DialogResults.DialogOK:
            return

        filepath = file_dlg.filename
        export_settings = _export_settings_from_ui(settings)
        result = export_cutlist.export_json(
            _groups, filepath, export_settings,
            weld_plans=_weld_plans, hole_summaries=_hole_summaries,
        )
        logger.info('Review palette JSON export: %s', result.filepath)
        _ui.messageBox('JSON saved to:\n{}'.format(result.filepath), 'Export Complete')

    except Exception:
        _ui.messageBox('JSON export error:\n{}'.format(traceback.format_exc()))


# ---------------------------------------------------------------------------
# Fastener export
# ---------------------------------------------------------------------------

def _export_fasteners(settings: dict) -> None:
    try:
        file_dlg = _ui.createFileDialog()
        file_dlg.filter           = 'CSV Files (*.csv)'
        file_dlg.initialFilename  = 'SmartCutList_Fasteners.csv'
        file_dlg.title            = 'Save Fastener Procurement List as CSV'
        if file_dlg.showSave() != adsk.core.DialogResults.DialogOK:
            return

        filepath = file_dlg.filename
        export_settings = _export_settings_from_ui(settings)
        result = export_cutlist.export_fasteners_csv(_groups, filepath, export_settings)
        logger.info('Fastener CSV export: %s', result.filepath)
        _ui.messageBox('Fastener CSV saved to:\n{}'.format(result.filepath), 'Export Complete')

    except Exception:
        _ui.messageBox('Fastener CSV export error:\n{}'.format(traceback.format_exc()))


def _export_sourced(settings: dict) -> None:
    try:
        file_dlg = _ui.createFileDialog()
        file_dlg.filter           = 'CSV Files (*.csv)'
        file_dlg.initialFilename  = 'SmartCutList_Sourced.csv'
        file_dlg.title            = 'Save Sourced Components as CSV'
        if file_dlg.showSave() != adsk.core.DialogResults.DialogOK:
            return

        filepath = file_dlg.filename
        export_settings = _export_settings_from_ui(settings)
        result = export_cutlist.export_sourced_csv(_groups, filepath, export_settings)
        logger.info('Sourced CSV export: %s', result.filepath)
        _ui.messageBox('Sourced CSV saved to:\n{}'.format(result.filepath), 'Export Complete')

    except Exception:
        _ui.messageBox('Sourced CSV export error:\n{}'.format(traceback.format_exc()))


# ---------------------------------------------------------------------------
# DXF / package export
# ---------------------------------------------------------------------------

def _find_body(design: adsk.fusion.Design, body_name: str) -> Optional[adsk.fusion.BRepBody]:
    """Search all components in the design for a body with the given name."""
    for comp in design.allComponents:
        for body in comp.bRepBodies:
            if body.name == body_name:
                return body
    return None


def _sheet_metal_groups_needing_configuration() -> list:
    pending = []
    for group in _included_groups():
        if group.get('classified_type') == PartType.SHEET_METAL and not group.get('sheet_metal_configured'):
            pending.append(group)
    return pending


def _export_package(settings: dict) -> None:
    try:
        pending = _sheet_metal_groups_needing_configuration()
        if pending:
            names = ', '.join(g.get('display_name', 'Unknown') for g in pending[:5])
            if len(pending) > 5:
                names += ', ...'
            _ui.messageBox(
                'Configure all included sheet metal groups before exporting DXFs.\n\nPending: {}'.format(names),
                'Sheet Metal Setup Required',
            )
            return

        folder_dlg = _ui.createFolderDialog()
        folder_dlg.title = 'Select Export Folder'
        if folder_dlg.showDialog() != adsk.core.DialogResults.DialogOK:
            return

        output_dir = folder_dlg.folder
        export_settings = _export_settings_from_ui(settings)
        base_name = _sanitise_stem(export_settings.project_name or 'SmartCutList')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        csv_path = os.path.join(output_dir, '{}_cut_list_{}.csv'.format(base_name, timestamp))

        csv_result = export_cutlist.export_csv(
            _groups, csv_path, export_settings, hole_summaries=_hole_summaries
        )
        dxf_paths = export_dxf.export_dxfs(
            _groups,
            output_dir,
            export_settings.filename_template,
            {
                'unit': export_settings.unit,
                'sheet_metal_only': True,
                'cross_section_for_profiles': False,
                'include_bend_lines': False,
                'include_dimensions': False,
                'show_summary': False,
            },
        )

        included_sheet_metal = [
            group for group in _included_groups()
            if group.get('classified_type') == PartType.SHEET_METAL
        ]
        dxf_failed = max(0, len(included_sheet_metal) - len(dxf_paths))

        # STEP export for 3D printed parts
        step_paths = export_dxf.export_steps(
            _groups,
            output_dir,
            export_settings.filename_template,
            {'unit': export_settings.unit, 'show_summary': False},
        )

        # Sourced components CSV
        sourced_csv_path = os.path.join(
            output_dir, '{}_sourced_{}.csv'.format(base_name, timestamp)
        )
        sourced_result = export_cutlist.export_sourced_csv(
            _groups, sourced_csv_path, export_settings
        )

        logger.info(
            'Review package export complete: csv=%s dxfs=%s failed_dxfs=%s steps=%s sourced=%s',
            csv_result.filepath,
            len(dxf_paths),
            dxf_failed,
            len(step_paths),
            sourced_result.parts_exported,
        )
        _ui.messageBox(
            'Package export complete.\n\n'
            'Cut list CSV: {}\n'
            'Sheet metal DXFs: {}\n'
            'Sheet metal DXF failures: {}\n'
            '3D print STEPs: {}\n'
            'Sourced components: {}\n\n'
            'See the SmartCutList log for any skipped or failed parts.'.format(
                csv_result.filepath,
                len(dxf_paths),
                dxf_failed,
                len(step_paths),
                sourced_result.parts_exported,
            ),
            'Export Complete',
        )
    except Exception:
        logger.exception('Package export failed')
        _ui.messageBox('Package export error:\n{}'.format(traceback.format_exc()))
