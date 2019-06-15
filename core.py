from comtypes.client import CreateObject

from .exceptions import InvalidSessionName, WindowNotFound


class Instance(object):
    def __call__(self, session=None):
        """Initalize AS400 session connection and PS and OIA as well,
        param: session: alphabatical instance name required like 'A', 'B'.."""
        if (session is None or
                not isinstance(session, str) or
                len(session) != 1 or
                ord(session.upper()) not in range(65, 91)):
            raise InvalidSessionName("Session not found")

        conn = CreateObject('PCOMM.autECLConnList')
        # Refresh the connection with available sessions
        conn.Refresh

        self.session = CreateObject('PCOMM.autECLSession')

        if conn.FindConnectionByName(session) is None:
            raise WindowNotFound("AS400 not found, Run AS400 first.")

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
        """It creates screen object, and for created screen
        some description get provided like below,
        screen = create_screen()
        screen.AddCursorPos(23, 1)"""
        return CreateObject('PCOMM.autECLScreenDesc')

    def gettext(self, row, column, length):
        """Return trimmed string from AS400 instance
        row and column divide screen into x and y axis
        param: row: location of x axis
               column: location of y axis
               length: alphabets that need to be fetch from screen"""
        self._key_event_setup()
        return self.ps.GetText(row, column, length).encode('UTF-8').strip(' ')

    def sendkeys(self, key):
        """Send keys to as400 instance.
        param: key: keystrokes that need to be send to the session

        Mnemonic Keywords for a list of these keystrokes can be found
        here: https://ibm.co/31yC100"""
        self._key_event_setup()
        self.ps.SendKeys(key)

    def settext(self, text, row, column):
        """Set text to specified row and column
        row and column divide screen into x and y axis
        param: text: string that needs to be set
               row: location of x axis
               column: location of y axis"""
        self._key_event_setup()
        self.ps.SetText(text, row, column)

    def screen_wait(self, screen, second):
        """for given screen object wait for specified time
        param: screen: pass the screen descriptive object
               second: wait time in second"""
        return self.ps.WaitForScreen(screen, second * 1000)

    def wait(self, seconds):
        """for given seconds make pointer sleep implicitly"""
        self.ps.Wait(int(seconds * 1000))
