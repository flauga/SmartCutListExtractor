from . import select_components
from . import settings


def start():
    select_components.start()
    settings.start()


def stop():
    settings.stop()
    select_components.stop()
