import traceback
import adsk.core
from . import commands


def run(context):
    try:
        commands.start()
    except:
        app = adsk.core.Application.get()
        app.userInterface.messageBox(
            'SmartCutList failed to start:\n{}'.format(traceback.format_exc())
        )


def stop(context):
    try:
        commands.stop()
    except:
        app = adsk.core.Application.get()
        app.userInterface.messageBox(
            'SmartCutList failed to stop:\n{}'.format(traceback.format_exc())
        )
