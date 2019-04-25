"""New IBM library to handle AS400 sessions with BC06 window

highly customizable with only help of InstanceAction and subclassing
InstanceHandler

Tow modules are available, core and application_system"""
__version__ = '1.0'

from .core import Instance, InstanceActions
from .exceptions import InvalidSerial, InvalidSessionName, ScreenTransition
