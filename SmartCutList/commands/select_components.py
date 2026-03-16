import logging
import traceback

import adsk.core
import adsk.fusion

from . import classifier
from . import feature_extraction
from . import review_palette
from . import settings

# ---------------------------------------------------------------------------
# Module-level globals
# ---------------------------------------------------------------------------

_app: adsk.core.Application = None
_ui: adsk.core.UserInterface = None
_handlers: list = []
_cmd_handlers: list = []
selected_items: list = []

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PANEL_ID = 'SolidMakePanel'
CMD_ID = 'SmartCutListSelectCmd'
CMD_NAME = 'Smart Cut List'
CMD_DESCRIPTION = 'Select components and bodies to generate a smart cut list.'

# ---------------------------------------------------------------------------
# Add-in lifecycle
# ---------------------------------------------------------------------------

def start():
    """Register the main SmartCutList command."""
    global _app, _ui
    try:
        _app = adsk.core.Application.get()
        _ui = _app.userInterface

        existing = _ui.commandDefinitions.itemById(CMD_ID)
        if existing:
            existing.deleteMe()

        cmd_def = _ui.commandDefinitions.addButtonDefinition(
            CMD_ID, CMD_NAME, CMD_DESCRIPTION, ''
        )

        on_created = CommandCreatedHandler()
        cmd_def.commandCreated.add(on_created)
        _handlers.append(on_created)

        make_panel = _ui.allToolbarPanels.itemById(PANEL_ID)
        if make_panel is None:
            ws = _ui.workspaces.itemById('FusionSolidEnvironment')
            if ws:
                make_panel = ws.toolbarPanels.itemById(PANEL_ID)

        if make_panel is None:
            _ui.messageBox(
                'SmartCutList: Could not find the Make panel (SolidMakePanel).\n'
                'The button will not be visible in the toolbar.'
            )
            return

        control = make_panel.controls.addCommand(cmd_def)
        control.isPromoted = True

    except Exception:
        if _ui:
            _ui.messageBox('SmartCutList start() error:\n{}'.format(traceback.format_exc()))


def stop():
    """Remove the command definition and release event handlers."""
    global _handlers, _cmd_handlers, selected_items

    try:
        make_panel = _ui.allToolbarPanels.itemById(PANEL_ID) if _ui else None
        if make_panel:
            ctrl = make_panel.controls.itemById(CMD_ID)
            if ctrl:
                ctrl.deleteMe()

        cmd_def = _ui.commandDefinitions.itemById(CMD_ID) if _ui else None
        if cmd_def:
            cmd_def.deleteMe()

    except Exception:
        if _ui:
            _ui.messageBox('SmartCutList stop() error:\n{}'.format(traceback.format_exc()))
    finally:
        _handlers = []
        _cmd_handlers = []
        selected_items = []


# ---------------------------------------------------------------------------
# Event handler classes
# ---------------------------------------------------------------------------

class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """Build the selection dialog when the command is launched."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandCreatedEventArgs):
        global _cmd_handlers
        try:
            _cmd_handlers = []
            saved_settings = settings.load_settings()
            inputs = args.command.commandInputs

            sel = inputs.addSelectionInput(
                'sel_input',
                'Components / Bodies',
                'Select components (occurrences) or solid bodies'
            )
            sel.addSelectionFilter('Occurrences')
            sel.addSelectionFilter('SolidBodies')
            sel.setSelectionLimits(1, 0)

            inputs.addBoolValueInput(
                'include_sub',
                'Include sub-components',
                True,
                '',
                True,
            )
            inputs.addBoolValueInput(
                'include_fasteners',
                'Include fasteners',
                True,
                '',
                bool(saved_settings.get('include_fasteners', True)),
            )

            on_execute = CommandExecuteHandler()
            args.command.execute.add(on_execute)
            _cmd_handlers.append(on_execute)

            on_destroy = CommandDestroyHandler()
            args.command.destroy.add(on_destroy)
            _cmd_handlers.append(on_destroy)

        except Exception:
            _ui.messageBox('CommandCreated error:\n{}'.format(traceback.format_exc()))


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    """Resolve selected bodies, classify them, and open the review palette."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            global selected_items
            selected_items = []

            saved_settings = settings.load_settings()
            inputs = args.command.commandInputs
            sel_input = inputs.itemById('sel_input')
            include_sub = inputs.itemById('include_sub').value
            include_fasteners = inputs.itemById('include_fasteners').value
            seen_tokens = set()

            for index in range(sel_input.selectionCount):
                collect_items(
                    sel_input.selection(index).entity,
                    include_sub,
                    selected_items,
                    seen_tokens,
                )

            # When include_fasteners is on, also scan the root component for
            # fastener occurrences (content-library fasteners appear as a
            # separate "Fasteners" section in the browser, not as children of
            # the selected components, so they are missed by the selection loop).
            if include_fasteners:
                try:
                    design = adsk.fusion.Design.cast(_app.activeProduct)
                    if design:
                        fastener_items = _collect_root_fasteners(design, seen_tokens)
                        selected_items.extend(fastener_items)
                        if fastener_items:
                            logger.info(
                                'Found %d fastener body/bodies from root-level occurrences',
                                len(fastener_items),
                            )
                except Exception:
                    logger.exception('Root-level fastener scan failed (non-fatal)')

            # Scan for sourced (externally linked) components
            try:
                design = adsk.fusion.Design.cast(_app.activeProduct)
                if design:
                    sourced_items = _collect_sourced_components(design, seen_tokens)
                    selected_items.extend(sourced_items)
                    if sourced_items:
                        logger.info(
                            'Found %d sourced body/bodies from externally referenced components',
                            len(sourced_items),
                        )
            except Exception:
                logger.exception('Sourced component scan failed (non-fatal)')

            # Pass full item dicts (including component_path) to feature extraction
            body_items = [item for item in selected_items if item.get('type') == 'BRepBody']
            if not body_items:
                _ui.messageBox(
                    'No solid bodies were found in the selected components.',
                    'Smart Cut List',
                )
                return

            logger.info(
                'SmartCutList selection resolved to %s bodies (include_sub=%s)',
                len(body_items),
                include_sub,
            )

            features = feature_extraction.extract_features_with_context(body_items)
            classified = classifier.classify_bodies(
                features,
                confidence_threshold=saved_settings.get('confidence_threshold', 0.6),
            )

            palette_settings = dict(saved_settings)
            palette_settings['include_fasteners'] = include_fasteners
            review_palette.start(classified, palette_settings)

        except Exception:
            _ui.messageBox('Execute error:\n{}'.format(traceback.format_exc()))


class CommandDestroyHandler(adsk.core.CommandEventHandler):
    """Release per-invocation handlers when the command closes."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs):
        global _cmd_handlers
        _cmd_handlers = []


# ---------------------------------------------------------------------------
# Body collection helpers
# ---------------------------------------------------------------------------

def collect_items(root_entity, include_sub: bool, result: list, seen_tokens: set = None,
                  component_path: str = ''):
    """
    Collect solid bodies from selected occurrences and direct body picks,
    tracking the full component hierarchy path for each body.

    The stack holds (entity, current_path) tuples so component nesting
    is preserved without recursion.
    """
    if seen_tokens is None:
        seen_tokens = set()

    occ_type  = adsk.fusion.Occurrence.classType()
    body_type = adsk.fusion.BRepBody.classType()
    stack = [(root_entity, component_path)]

    while stack:
        entity, current_path = stack.pop()

        if entity.objectType == occ_type:
            occ = adsk.fusion.Occurrence.cast(entity)
            occ_path = (current_path + '/' + occ.name) if current_path else occ.name
            _append_occurrence_bodies(occ, result, seen_tokens, occ_path)
            if include_sub:
                children = list(occ.childOccurrences)
                stack.extend([(child, occ_path) for child in children[::-1]])

        elif entity.objectType == body_type:
            body = adsk.fusion.BRepBody.cast(entity)
            _append_body(body, result, seen_tokens, current_path)


def _append_occurrence_bodies(occ, result: list, seen_tokens: set,
                               component_path: str) -> None:
    bodies = []
    try:
        bodies = list(occ.bRepBodies)
    except Exception:
        bodies = []

    if not bodies:
        try:
            bodies = list(occ.component.bRepBodies)
        except Exception:
            bodies = []

    for body in bodies:
        body_ref = adsk.fusion.BRepBody.cast(body)
        if body_ref is not None:
            _append_body(body_ref, result, seen_tokens, component_path)


def _append_body(body, result: list, seen_tokens: set,
                 component_path: str = '') -> None:
    token = _body_token(body)
    if token in seen_tokens:
        return

    seen_tokens.add(token)
    # Fall back to component name if no occurrence path was provided
    path = component_path or body.parentComponent.name
    result.append({
        'name':           body.name,
        'type':           'BRepBody',
        'parent':         body.parentComponent.name,
        'component_path': path,
        'entity':         body,
    })


def _body_token(body) -> str:
    try:
        return body.entityToken
    except Exception:
        return '{}::{}'.format(body.parentComponent.name, body.name)


def _collect_root_fasteners(design: adsk.fusion.Design, seen_tokens: set) -> list:
    """
    Collect fasteners inserted via Insert > Fasteners (content library).

    Fusion 360 groups these under a "Fasteners" folder in the browser but
    exposes no direct API for that folder.  Instead we check each
    occurrence's ``component.isLibraryItem`` which returns True only for
    components from the built-in content/fastener library — much more
    specific than ``isReferencedComponent`` which matches ANY external
    reference (bearings, purchased parts, sub-assemblies, etc.).
    """
    result = []
    try:
        all_occs = design.rootComponent.allOccurrences
    except Exception:
        logger.exception('_collect_root_fasteners: could not read allOccurrences')
        return result

    for occ in all_occs:
        # Only collect content-library items (fasteners inserted via
        # Insert > Fasteners).  isLibraryItem is far more specific than
        # isReferencedComponent which matches all external references.
        try:
            comp = occ.component
            if comp is None or not comp.isLibraryItem:
                continue
        except Exception:
            continue

        try:
            occ_name = occ.name or ''
        except Exception:
            continue

        # Try the occurrence's own bodies first, then fall back to the component.
        bodies = []
        try:
            bodies = list(occ.bRepBodies)
        except Exception:
            pass
        if not bodies:
            try:
                bodies = list(comp.bRepBodies)
            except Exception:
                pass

        for body in bodies:
            body_ref = adsk.fusion.BRepBody.cast(body)
            if body_ref is None:
                continue
            token = _body_token(body_ref)
            if token in seen_tokens:
                continue
            seen_tokens.add(token)
            try:
                parent_name = body_ref.parentComponent.name
            except Exception:
                parent_name = occ_name
            result.append({
                'name':           body_ref.name,
                'type':           'BRepBody',
                'parent':         parent_name,
                'component_path': occ_name,
                'entity':         body_ref,
                'is_content_library_fastener': True,
            })

    return result


def _collect_sourced_components(design: adsk.fusion.Design, seen_tokens: set) -> list:
    """
    Collect externally linked/referenced components (NOT content-library fasteners).

    These are components saved as external files and linked into the design —
    purchased parts, bearings, motors, etc. that are not from the Fusion 360
    content library.  Identified by ``isReferencedComponent=True`` AND
    ``isLibraryItem=False``.
    """
    result = []
    try:
        all_occs = design.rootComponent.allOccurrences
    except Exception:
        logger.exception('_collect_sourced_components: could not read allOccurrences')
        return result

    for occ in all_occs:
        try:
            comp = occ.component
            if comp is None:
                continue
            # Must be externally referenced but NOT a content-library item
            if not comp.isReferencedComponent:
                continue
            if comp.isLibraryItem:
                continue  # fasteners handled by _collect_root_fasteners
        except Exception:
            continue

        try:
            occ_name = occ.name or ''
        except Exception:
            continue

        bodies = []
        try:
            bodies = list(occ.bRepBodies)
        except Exception:
            pass
        if not bodies:
            try:
                bodies = list(comp.bRepBodies)
            except Exception:
                pass

        for body in bodies:
            body_ref = adsk.fusion.BRepBody.cast(body)
            if body_ref is None:
                continue
            token = _body_token(body_ref)
            if token in seen_tokens:
                continue
            seen_tokens.add(token)
            try:
                parent_name = body_ref.parentComponent.name
            except Exception:
                parent_name = occ_name
            result.append({
                'name':           body_ref.name,
                'type':           'BRepBody',
                'parent':         parent_name,
                'component_path': occ_name,
                'entity':         body_ref,
                'is_sourced_component': True,
            })

    return result
