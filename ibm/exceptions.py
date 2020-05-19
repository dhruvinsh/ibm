class InvalidSessionName(Exception):
    """Invalid session name get passed"""


class ScreenTransition(Exception):
    """Fail to transit to new screen"""


class WindowNotFound(Exception):
    """AS400 window not found"""
