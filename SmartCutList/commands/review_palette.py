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
  'htmlReady'         {}                                → trigger initial data push
  'override_type'     {group_id, value}                → reclassify a group
  'override_material' {group_id, value}                → rename material
  'toggle_include'    {group_id, value: bool}          → include/exclude from export
  'export_csv'        {settings}                       → write CSV to disk
  'export_json'       {settings}                       → write JSON to disk
  'export_dxf'        {settings}                       → write DXF files to disk

Actions Python sends to HTML
-----------------------------
  'initData'          {groups: [...], part_types: [...]}
"""

from __future__ import annotations

import csv
import json
import os
import re
import traceback
from typing import Optional

import adsk.core
import adsk.fusion

from .classifier import PartType


# ---------------------------------------------------------------------------
# Module-level globals
# ---------------------------------------------------------------------------

_app:     adsk.core.Application   = None
_ui:      adsk.core.UserInterface = None
_palette: adsk.core.Palette       = None
_handlers: list                   = []
_groups:   list                   = []   # current grouped cut list data

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PALETTE_ID   = 'SmartCutListReviewPalette'
PALETTE_NAME = 'Smart Cut List — Review'

# HTML file sits in SmartCutList/resources/ relative to this commands/ directory
_HTML_PATH = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)),
                 '..', 'resources', 'review_palette.html')
)

# All exported types that support DXF (flat-profile parts)
_DXF_TYPES = {PartType.SHEET_METAL, PartType.FLAT_BAR}

# Dimension grouping tolerance in mm (round to nearest value)
_DIM_ROUND = 0.1


# ---------------------------------------------------------------------------
# Part grouping
# ---------------------------------------------------------------------------

def _round_dim(value: float) -> float:
    """Round a single dimension to the nearest _DIM_ROUND for tolerance matching."""
    return round(round(value / _DIM_ROUND) * _DIM_ROUND, 4)


def _dims_key(bb_mm) -> Optional[tuple]:
    """Return a hashable, sorted, tolerance-snapped dimension tuple, or None."""
    if not bb_mm or len(bb_mm) < 3:
        return None
    try:
        return tuple(_round_dim(float(d)) for d in sorted(bb_mm))
    except (TypeError, ValueError):
        return None


def group_classified_bodies(classified_features: list) -> list:
    """
    Group classified feature dicts by (classified_type, material, sorted dims).

    Dimensions are compared within _DIM_ROUND tolerance.  Parts with
    extraction errors are included as singleton error groups so they are
    still visible in the UI.

    Returns a list of group dicts, each containing:
        group_id              str
        display_name          str   — name of the first body in the group
        component_name        str
        classified_type       str   — active type (may be overridden later)
        confidence            float
        classification_reason str
        needs_review          bool
        material_name         str
        dimensions_mm         list[float] | None  — sorted [small, mid, large]
        quantity              int
        include               bool
        bodies                list[str]  — body names in this group
        override_type         str | None
        override_material     str | None
        feature_history       list[str]
        is_sheet_metal        bool
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
                'classified_type':      PartType.UNKNOWN,
                'confidence':           0.0,
                'classification_reason': 'Feature extraction failed',
                'needs_review':         True,
                'material_name':        'Unknown',
                'dimensions_mm':        None,
                'quantity':             1,
                'include':              False,
                'bodies':               [feat.get('body_name', '')],
                'override_type':        None,
                'override_material':    None,
                'feature_history':      [],
                'is_sheet_metal':       False,
            })
            continue

        ctype    = feat.get('classified_type', PartType.UNKNOWN)
        material = feat.get('material_name') or 'Unknown'
        bb       = feat.get('bounding_box_mm')
        dk       = _dims_key(bb)
        key      = (ctype, material, dk)

        if key in group_map:
            g = groups[group_map[key]]
            g['quantity'] += 1
            g['bodies'].append(feat.get('body_name', ''))
        else:
            idx = len(groups)
            groups.append({
                'group_id':             'g_{}'.format(idx),
                'display_name':         (feat.get('body_name')
                                         or feat.get('component_name')
                                         or 'Unknown'),
                'component_name':       feat.get('component_name', ''),
                'classified_type':      ctype,
                'confidence':           feat.get('confidence', 0.0),
                'classification_reason': feat.get('classification_reason', ''),
                'needs_review':         feat.get('needs_review', False),
                'material_name':        material,
                'dimensions_mm':        sorted(bb) if bb else None,
                'quantity':             1,
                'include':              True,
                'bodies':               [feat.get('body_name', '')],
                'override_type':        None,
                'override_material':    None,
                'feature_history':      feat.get('feature_history', []),
                'is_sheet_metal':       bool(feat.get('is_sheet_metal', False)),
            })
            group_map[key] = idx

    return groups


# ---------------------------------------------------------------------------
# Palette lifecycle
# ---------------------------------------------------------------------------

def start(classified_features: list) -> None:
    """Create the review palette and populate it with grouped body data."""
    global _app, _ui, _palette, _groups, _handlers

    try:
        _app = adsk.core.Application.get()
        _ui  = _app.userInterface

        _groups = group_classified_bodies(classified_features)

        # Remove any stale palette from a previous invocation
        existing = _ui.palettes.itemById(PALETTE_ID)
        if existing:
            existing.deleteMe()

        # Convert Windows path to a file:// URL Fusion can load
        html_url = 'file:///' + _HTML_PATH.replace('\\', '/')

        _palette = _ui.palettes.add(
            PALETTE_ID,
            PALETTE_NAME,
            html_url,
            True,    # isVisible
            True,    # showCloseButton
            True,    # isResizable
            960,     # width  (px)
            640,     # height (px)
        )
        _palette.dockingState = adsk.core.PaletteDockingStates.PaletteDockStateFloating

        on_html = HTMLEventHandler()
        _palette.incomingFromHTML.add(on_html)
        _handlers.append(on_html)

    except Exception:
        if _ui:
            _ui.messageBox('review_palette start() error:\n{}'.format(
                traceback.format_exc()))


def stop() -> None:
    """Destroy the palette and release all handler references."""
    global _palette, _handlers, _groups
    try:
        if _palette:
            _palette.deleteMe()
            _palette = None
    except Exception:
        pass
    finally:
        _handlers = []
        _groups   = []


# ---------------------------------------------------------------------------
# Data → HTML
# ---------------------------------------------------------------------------

def _send_init_data() -> None:
    """Push the current group list and available part types to the palette."""
    if _palette is None:
        return
    try:
        payload = {
            'groups': _groups,
            'part_types': [
                PartType.SHEET_METAL,
                PartType.RECTANGULAR_TUBE,
                PartType.ROUND_TUBE,
                PartType.ALUMINIUM_EXTRUSION,
                PartType.ANGLE_SECTION,
                PartType.C_CHANNEL,
                PartType.FLAT_BAR,
                PartType.MILLED_BLOCK,
                PartType.ROUND_BAR,
                PartType.FASTENER,
                PartType.UNKNOWN,
            ],
        }
        _palette.sendInfoToHTML('initData', json.dumps(payload, default=str))
    except Exception:
        if _ui:
            _ui.messageBox('_send_init_data error:\n{}'.format(
                traceback.format_exc()))


# ---------------------------------------------------------------------------
# Message handlers (HTML → Python)
# ---------------------------------------------------------------------------

def _find_group(group_id: str) -> Optional[dict]:
    for g in _groups:
        if g['group_id'] == group_id:
            return g
    return None


def _handle_html_ready(_data: dict) -> None:
    _send_init_data()


def _handle_override_type(data: dict) -> None:
    g = _find_group(data.get('group_id', ''))
    if g:
        g['classified_type'] = data.get('value', PartType.UNKNOWN)
        g['override_type']   = g['classified_type']


def _handle_override_material(data: dict) -> None:
    g = _find_group(data.get('group_id', ''))
    if g:
        g['material_name']    = data.get('value', 'Unknown')
        g['override_material'] = g['material_name']


def _handle_toggle_include(data: dict) -> None:
    g = _find_group(data.get('group_id', ''))
    if g:
        g['include'] = bool(data.get('value', True))


def _handle_export_csv(data: dict) -> None:
    _export_csv(data.get('settings', {}))


def _handle_export_json(data: dict) -> None:
    _export_json(data.get('settings', {}))


def _handle_export_dxf(data: dict) -> None:
    _export_dxf(data.get('settings', {}))


_MESSAGE_HANDLERS = {
    'htmlReady':        _handle_html_ready,
    'override_type':    _handle_override_type,
    'override_material':_handle_override_material,
    'toggle_include':   _handle_toggle_include,
    'export_csv':       _handle_export_csv,
    'export_json':      _handle_export_json,
    'export_dxf':       _handle_export_dxf,
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
    unit_suffix = 'in' if use_inches else 'mm'
    dims       = group.get('dimensions_mm')

    if dims and len(dims) >= 3:
        dims_str = 'x'.join('{:.1f}'.format(d * factor) for d in dims) + unit_suffix
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
        template = settings.get('naming_template', '{name}_{material}_{dims}')
        unit     = settings.get('unit', 'mm')
        use_in   = unit == 'inches'
        factor   = 1 / 25.4 if use_in else 1.0
        unit_lbl = 'in' if use_in else 'mm'

        def fmt(d):
            return '{:.4f}'.format(d * factor) if d is not None else ''

        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([
                'Part Name', 'Component', 'Type', 'Material',
                'Length ({})'.format(unit_lbl),
                'Width ({})'.format(unit_lbl),
                'Thickness ({})'.format(unit_lbl),
                'Qty', 'Confidence', 'Needs Review',
                'Export Filename', 'Notes',
            ])
            for g in _included_groups():
                dims = (g.get('dimensions_mm') or []) + [None, None, None]
                # dims is sorted [small, mid, large]; expose as L/W/T
                writer.writerow([
                    g['display_name'],
                    g.get('component_name', ''),
                    g['classified_type'],
                    g['material_name'],
                    fmt(dims[2]),   # length (largest)
                    fmt(dims[1]),   # width
                    fmt(dims[0]),   # thickness (smallest)
                    g['quantity'],
                    '{:.2f}'.format(g['confidence']),
                    'Yes' if g['needs_review'] else 'No',
                    _build_filename(g, template, settings),
                    g.get('classification_reason', ''),
                ])

        _ui.messageBox('CSV saved to:\n{}'.format(filepath), 'Export Complete')

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
        template = settings.get('naming_template', '{name}_{material}_{dims}')

        # Omit the large feature_history from the export payload
        export_data = []
        for g in _included_groups():
            entry = {k: v for k, v in g.items() if k != 'feature_history'}
            entry['export_filename'] = _build_filename(g, template, settings)
            export_data.append(entry)

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, indent=2, default=str)

        _ui.messageBox('JSON saved to:\n{}'.format(filepath), 'Export Complete')

    except Exception:
        _ui.messageBox('JSON export error:\n{}'.format(traceback.format_exc()))


# ---------------------------------------------------------------------------
# DXF export
# ---------------------------------------------------------------------------

def _find_body(design: adsk.fusion.Design, body_name: str) -> Optional[adsk.fusion.BRepBody]:
    """Search all components in the design for a body with the given name."""
    for comp in design.allComponents:
        for body in comp.bRepBodies:
            if body.name == body_name:
                return body
    return None


def _export_body_dxf(design: adsk.fusion.Design, group: dict,
                     filepath: str, settings: dict) -> bool:
    """
    Export one body to DXF.

    Sheet metal bodies → Fusion flat-pattern DXF (respects bend lines setting).
    All others         → temporary sketch on largest face, saved as DXF.

    Returns True on success, False on failure.
    """
    body_name = group['bodies'][0] if group['bodies'] else ''
    body = _find_body(design, body_name)
    if body is None:
        return False

    export_mgr = design.exportManager
    comp       = body.parentComponent

    # --- Sheet metal flat-pattern path ---
    if group['classified_type'] == PartType.SHEET_METAL and body.isSheetMetal:
        try:
            sm = adsk.fusion.SheetMetalComponent.cast(comp)
            if sm and sm.hasActiveFlatPattern:
                flat_pattern = sm.flatPattern
                dxf_opts = export_mgr.createDXF2DExportOptions(filepath, flat_pattern)
                # Bend-line visibility (API attribute name varies by Fusion version)
                for attr in ('isBendLinesVisible', 'bendLinesVisible'):
                    if hasattr(dxf_opts, attr):
                        setattr(dxf_opts, attr, bool(settings.get('include_bend_lines', True)))
                        break
                export_mgr.execute(dxf_opts)
                return True
        except Exception:
            pass   # Fall through to generic sketch approach

    # --- Generic: project largest face onto a temporary sketch ---
    try:
        largest_face = max(body.faces, key=lambda f: f.area)
    except ValueError:
        return False   # body has no faces

    sketch = comp.sketches.add(largest_face)
    sketch.name = '__SCL_DXF_tmp'
    try:
        dxf_opts = export_mgr.createDXF2DExportOptions(filepath, sketch)
        export_mgr.execute(dxf_opts)
        return True
    except Exception:
        return False
    finally:
        try:
            sketch.deleteMe()
        except Exception:
            pass


def _export_dxf(settings: dict) -> None:
    try:
        dxf_groups = [g for g in _included_groups()
                      if g['classified_type'] in _DXF_TYPES]

        if not dxf_groups:
            _ui.messageBox(
                'No sheet metal or flat bar parts are included for DXF export.\n'
                'Enable "Include" for at least one such part and try again.',
                'DXF Export')
            return

        folder_dlg = _ui.createFolderDialog()
        folder_dlg.title = 'Select DXF Output Folder'
        if folder_dlg.showDialog() != adsk.core.DialogResults.DialogOK:
            return

        folder   = folder_dlg.folder
        template = settings.get('naming_template', '{name}_{material}_{dims}')
        design   = adsk.fusion.Design.cast(_app.activeProduct)

        succeeded, failed = 0, 0
        for g in dxf_groups:
            filename = _build_filename(g, template, settings) + '.dxf'
            filepath = os.path.join(folder, filename)
            if _export_body_dxf(design, g, filepath, settings):
                succeeded += 1
            else:
                failed += 1

        msg = 'DXF export complete.\n{} file(s) written to:\n{}'.format(
            succeeded, folder)
        if failed:
            msg += '\n\n{} part(s) could not be exported (check body visibility).'.format(failed)
        _ui.messageBox(msg, 'DXF Export')

    except Exception:
        _ui.messageBox('DXF export error:\n{}'.format(traceback.format_exc()))
