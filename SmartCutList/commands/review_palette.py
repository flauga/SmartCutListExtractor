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

Actions Python sends to HTML
-----------------------------
  'initData'            {groups: [...], part_types: [...], weld_assemblies: [...]}
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


# ---------------------------------------------------------------------------
# Module-level globals
# ---------------------------------------------------------------------------

_app:     adsk.core.Application   = None
_ui:      adsk.core.UserInterface = None
_palette: adsk.core.Palette       = None
_handlers: list                   = []
_groups:   list                   = []   # current grouped cut list data
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


def _pick_display_name(feat: dict) -> str:
    """Choose the most informative name for a group's display_name.

    Content-library fasteners have generic body names ("Body", "Body1")
    but their component_path / component_name carries the full description
    (e.g. "Broached Hexagon Socket Head Cap Screw M10x1.5x50:1").
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
    # Standard parts: prefer body_name, fall back to component_name
    return feat.get('body_name') or feat.get('component_name') or 'Unknown'


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
                'bodies':               [feat.get('body_name', '')],
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
            g['bodies'].append(feat.get('body_name', ''))
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
                'bodies':               [feat.get('body_name', '')],
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


# ---------------------------------------------------------------------------
# Palette lifecycle
# ---------------------------------------------------------------------------

def start(classified_features: list, palette_settings: Optional[dict] = None) -> None:
    """Create the review palette and populate it with grouped body data."""
    global _app, _ui, _palette, _groups, _handlers, _palette_settings

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
    global _palette, _handlers, _groups, _palette_settings
    try:
        if _palette:
            _palette.deleteMe()
            _palette = None
    except Exception:
        pass
    finally:
        _handlers = []
        _groups   = []
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

    try:
        _highlight_group(group)
        body = export_dxf.resolve_body_for_part(group)
        if body is None:
            raise RuntimeError('Unable to resolve the selected Fusion body')

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
        result = export_cutlist.export_csv(_groups, filepath, export_settings)
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
        result = export_cutlist.export_json(_groups, filepath, export_settings)
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

        csv_result = export_cutlist.export_csv(_groups, csv_path, export_settings)
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
