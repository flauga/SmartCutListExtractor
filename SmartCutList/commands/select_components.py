import traceback
import adsk.core
import adsk.fusion

# ---------------------------------------------------------------------------
# Module-level globals
# ---------------------------------------------------------------------------

# _app and _ui are None until start() is called.  Assigning them here (rather
# than at import time) prevents failures when the module is imported before
# Fusion's API is fully initialised.
_app: adsk.core.Application  = None
_ui:  adsk.core.UserInterface = None

# CRITICAL: event handler objects must be kept in module-level lists so that
# Python's garbage collector does not destroy them while they are still active.
#
# _handlers     — permanent handlers (survive until stop() is called).
# _cmd_handlers — per-invocation handlers (reset on every command open/close).
#   Separating the two removes the fragile [:-2] slice that assumed a fixed
#   count of per-invocation handlers appended in a specific order.
_handlers:     list = []
_cmd_handlers: list = []

# Output list populated on each successful execute; consumed by later phases.
selected_items: list = []

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PANEL_ID        = 'SolidMakePanel'   # "Make" panel in Design workspace Utilities tab
CMD_ID          = 'SmartCutListSelectCmd'
CMD_NAME        = 'Smart Cut List'
CMD_DESCRIPTION = 'Select components and bodies to generate a smart cut list.'


# ---------------------------------------------------------------------------
# Add-in lifecycle
# ---------------------------------------------------------------------------

def start():
    """Register the button definition and attach it to the Make panel."""
    global _app, _ui
    try:
        _app = adsk.core.Application.get()
        _ui  = _app.userInterface

        # Clean up any leftover definition from a previously failed stop()
        existing = _ui.commandDefinitions.itemById(CMD_ID)
        if existing:
            existing.deleteMe()

        cmd_def = _ui.commandDefinitions.addButtonDefinition(
            CMD_ID, CMD_NAME, CMD_DESCRIPTION, ''
        )

        on_created = CommandCreatedHandler()
        cmd_def.commandCreated.add(on_created)
        _handlers.append(on_created)

        # Primary: workspace-agnostic panel lookup (most robust across versions)
        make_panel = _ui.allToolbarPanels.itemById(PANEL_ID)

        # Fallback: traverse workspace → panel
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
    """Remove the button and command definition; clear all handler references."""
    global _handlers, _cmd_handlers, selected_items

    try:
        make_panel = _ui.allToolbarPanels.itemById(PANEL_ID)
        if make_panel:
            ctrl = make_panel.controls.itemById(CMD_ID)
            if ctrl:
                ctrl.deleteMe()

        cmd_def = _ui.commandDefinitions.itemById(CMD_ID)
        if cmd_def:
            cmd_def.deleteMe()

    except Exception:
        if _ui:
            _ui.messageBox('SmartCutList stop() error:\n{}'.format(traceback.format_exc()))
    finally:
        # Always release handler references so GC can reclaim them
        _handlers     = []
        _cmd_handlers = []
        selected_items = []


# ---------------------------------------------------------------------------
# Event handler classes
# ---------------------------------------------------------------------------

class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """Fired when the user clicks the Smart Cut List button."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandCreatedEventArgs):
        global _cmd_handlers
        try:
            # Release any handlers from the previous invocation before
            # registering new ones for this dialog opening.
            _cmd_handlers = []

            inputs = args.command.commandInputs

            # --- Selection input ---
            sel = inputs.addSelectionInput(
                'sel_input',
                'Components / Bodies',
                'Select components (occurrences) or solid bodies'
            )
            sel.addSelectionFilter('Occurrences')   # component instances
            sel.addSelectionFilter('SolidBodies')   # solid BRep bodies
            sel.setSelectionLimits(1, 0)            # min 1, unlimited max (0)

            # --- Checkboxes ---
            inputs.addBoolValueInput(
                'include_sub', 'Include sub-components',
                True,  # isCheckBox
                '',    # resourceFolder (no custom icon)
                True   # initialValue: checked
            )
            inputs.addBoolValueInput(
                'use_ai', 'Use AI classification',
                True,   # isCheckBox
                '',
                False   # initialValue: unchecked
            )

            # Attach per-invocation handlers to _cmd_handlers (not _handlers),
            # so CommandDestroyHandler can release them cleanly without any
            # index arithmetic.
            on_execute = CommandExecuteHandler()
            args.command.execute.add(on_execute)
            _cmd_handlers.append(on_execute)

            on_destroy = CommandDestroyHandler()
            args.command.destroy.add(on_destroy)
            _cmd_handlers.append(on_destroy)

        except Exception:
            _ui.messageBox('CommandCreated error:\n{}'.format(traceback.format_exc()))


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    """Fired when the user clicks OK in the command dialog."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            global selected_items
            selected_items = []

            inputs      = args.command.commandInputs
            sel_input   = inputs.itemById('sel_input')
            include_sub = inputs.itemById('include_sub').value
            use_ai      = inputs.itemById('use_ai').value

            for i in range(sel_input.selectionCount):
                collect_items(sel_input.selection(i).entity, include_sub, selected_items)

            # Feedback summary
            lines = ['Collected {} item(s):'.format(len(selected_items))]
            for item in selected_items:
                lines.append('  - {} ({})'.format(item['name'], item['type']))
            if use_ai:
                lines.append('\n[AI classification enabled — not yet implemented]')

            _ui.messageBox('\n'.join(lines), 'Smart Cut List')

        except Exception:
            _ui.messageBox('Execute error:\n{}'.format(traceback.format_exc()))


class CommandDestroyHandler(adsk.core.CommandEventHandler):
    """Fired when the command dialog closes (OK or Cancel)."""

    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs):
        global _cmd_handlers
        try:
            # Release per-invocation handlers.  The permanent CommandCreatedHandler
            # in _handlers is intentionally left in place so the button works again.
            _cmd_handlers = []
        except Exception:
            pass   # nothing meaningful to report at destroy time


# ---------------------------------------------------------------------------
# Recursive item collection helper
# ---------------------------------------------------------------------------

def collect_items(root_entity, include_sub: bool, result: list):
    """
    Collect entities into result using an explicit stack.

    An iterative (stack-based) approach is used instead of recursion to avoid
    Python's default recursion limit on deeply nested assemblies.

    Each result entry is a dict:
      {'name': str, 'type': 'Occurrence'|'BRepBody',
       'component'|'parent': str, 'entity': <api object>}
    """
    occ_type  = adsk.fusion.Occurrence.classType()
    body_type = adsk.fusion.BRepBody.classType()

    stack = [root_entity]
    while stack:
        entity = stack.pop()

        if entity.objectType == occ_type:
            occ = adsk.fusion.Occurrence.cast(entity)
            result.append({
                'name':      occ.name,
                'type':      'Occurrence',
                'component': occ.component.name,
                'entity':    occ,
            })
            if include_sub:
                # Reverse children before pushing so that the first child is
                # processed first (stack is LIFO), preserving original order.
                children = list(occ.childOccurrences)
                stack.extend(children[::-1])

        elif entity.objectType == body_type:
            body = adsk.fusion.BRepBody.cast(entity)
            result.append({
                'name':   body.name,
                'type':   'BRepBody',
                'parent': body.parentComponent.name,
                'entity': body,
            })
