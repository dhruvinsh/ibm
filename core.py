import os
import subprocess
import time

from comtypes.client import CreateObject
from .exceptions import InvalidSessionName


class Instance(object):
    def __call__(self, session=None):
        """Initalize AS400 session connection and PS and OIA as well,
        param: session: alphabatical instance name required like 'A', 'B'.."""

        if (session is None or
                not isinstance(session, str) or
                len(session) != 1 or
                ord(session.upper()) not in range(65, 91)):
            raise InvalidSessionName

        conn = CreateObject('PCOMM.autECLConnList')
        conn.Refresh

        self.session = CreateObject('PCOMM.autECLSession')

        if (conn.FindConnectionByName(session) is not None):
            self.session.SetConnectionByName(session)
        else:
            session_loc = os.path.join('C:\\', 'Users', 'Public',
                                       'Desktop',
                                       'Session {}.ws'.format(session))
            if not os.path.isfile(session_loc):
                raise InvalidSessionName(
                    "Provided session Name {} is not valid".format(session))
            subprocess.Popen(
                '"{}"'.format(session_loc),
                shell=True,
                creationflags=subprocess.CREATE_NEW_CONSOLE)
            time.sleep(5)
            conn.Refresh
            self.session.SetConnectionByName(session)

        # pointer initialization
        self.ps = self.session.autECLPS
        self.oia = self.session.autECLOIA


class InstanceActions(Instance):
    def __init__(self, session):
        super(InstanceActions, self).__call__(session=session)

    @property
    def _input_inhibited(self):
        "get input inhibited status"
        return self.oia.InputInhibited

    def _input_wait(self, seconds):
        "wait for input to be ready"
        self.oia.WaitForInputReady(seconds * 1000)

    def _key_event_setup(self):
        "Make proper setup befor key events"
        self._input_wait(2)
        if self._input_inhibited != 0:
            self.ps.SendKeys('[ENTER]')
            self._input_wait(2)

    def create_screeen(self):
        "Created screen description object"
        return CreateObject('PCOMM.autECLScreenDesc')

    def gettext(self, row, column, length):
        "Return trimmed string value from AS400 instance"
        self._key_event_setup()
        return self.ps.GetText(row, column, length).encode('UTF-8').strip(' ')

    def sendkeys(self, key):
        "Send keys to as400 instance"
        self._key_event_setup()
        self.ps.SendKeys(key)

    def settext(self, text, row, column):
        "Set text to specified row and column"
        self._key_event_setup()
        self.ps.SetText(text, row, column)

    def screen_wait(self, screen, second):
        "for given screen object wait for specified time"
        return self.ps.WaitForScreen(screen, second * 1000)

    def wait(self, seconds):
        "for given seconds make pointer sleep"
        self.ps.Wait(int(seconds * 1000))
