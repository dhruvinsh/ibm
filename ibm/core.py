#
# Author: Dhruvin Shah
# Email: dhruvin3@gmail.com
#


from comtypes.client import CreateObject

from .exceptions import InvalidSessionName, WindowNotFound

class Instance(object):
    """A wrapper to AS400 session. Primitive methods like read, write is
    provided."""

    def __init__(self, session):
        """Initialize AS400 session connection and PS and OIA as well.
        param: session: alphabetical instance name required like 'A', 'B'"""
        if (
            not isinstance(session, str)
            or len(session) != 1
            or ord(session.upper()) not in range(65, 91)
        ):
            raise InvalidSessionName("Session: {} not found".format(session))

        conn = CreateObject("PCOMM.autECLConnList")
        # Refresh the connection with available sessions
        conn.Refresh

        self.session = CreateObject("PCOMM.autECLSession")

        if conn.FindConnectionByName(session) is None:
            raise WindowNotFound("AS400 not found, Run AS400 first.")

        self.session_desc = session
        self.session.SetConnectionByName(self.session_desc)

        # pointer initialization
        self.ps = self.session.autECLPS
        self.oia = self.session.autECLOIA

    def __repr__(self):
        return "<Session: {}>".format(self.session_desc)

    def _input_inhibited(self):
        """Get input inhibited status from AS400."""
        return self.oia.InputInhibited

    def _input_wait(self, second):
        """Wait for input to be ready.
        param: second: max wait time in second before raising the error."""
        self.oia.WaitForInputReady(second * 1000)

    def _key_event_setup(self, wait):
        """Get instance ready for the read/write task."""
        self._input_wait(wait)
        if self._input_inhibited() != 0:
            self.ps.SendKeys("[ENTER]")
            self._input_wait(wait)

    def get_text(self, row, column, length=1, wait=60):
        """Screen is divided in two axis: X and Y. For a given length, pass
        x, y location to read AS400 screen. Default word length is set to 1.

        :param row: x axis location.
        :param column: y axis location.
        :param length: length of alphabets that need to be read from screen.
        :param wait: max wait time, before raising the error, default is 2 sec.

        :return: stripped string from AS400 instance."""
        self._key_event_setup(wait)
        return self.ps.GetText(row, column, length).encode("UTF-8")

    def get_text_rect(self, startRow, startColumn, endRow, endColumn, wait=60):
        """Screen is divided in two axis: X and Y. For a given length, pass
        x, y location to read AS400 screen. Default word length is set to 1.

        :param startRow: start x axis location.
        :param startColumn: start y axis location.
        :param endRow: start x axis location.
        :param endColumn: start y axis location.
        :param wait: max wait time, before raising the error, default is 2 sec.

        :return: stripped string from AS400 instance."""
        self._key_event_setup(wait)
        return self.ps.GetTextRect(startRow, startColumn, endRow, endColumn).encode("UTF-8")

    def set_text(self, text, row, column, wait=60):
        """Set text to a specified location on AS400 screen.
        :param text: string that needs to be set.
        :param row: location of x axis.
        :param column: location of y axis.

        :return: None"""
        self._key_event_setup(wait)   #no tenia ningún parámetro
        self.ps.SetText(text, row, column)

    def set_cursor(self, row, column, wait=60):
        """Set cursor to a specified location on AS400 screen.
        :param row: location of x axis.
        :param column: location of y axis.

        :return: None"""
        self._key_event_setup(wait)   #no tenia ningún parámetro
        self.ps.SetCursorPos(row, column)

    def get_cursor_pos_row(self, wait=60):
        """Get cursor row on AS400 screen.

        :return: Cursor row"""
        self._key_event_setup(wait)   #no tenia ningún parámetro      
        return self.ps.CursorPosRow

    def get_cursor_pos_col(self, wait=60):
        """Get cursor row on AS400 screen.

        :return: Cursor col"""
        self._key_event_setup(wait)   #no tenia ningún parámetro      
        return self.ps.CursorPosCol

    def send_keys(self, key, wait=60):
        """Send keystrokes to AS400 screen.

        :param key: Mnemonic keystrokes that need to be send to the session.
                    List of these keystrokes can be found at
                    https://ibm.co/31yC100

        :return: None"""
        self._key_event_setup(wait)
        self.ps.SendKeys(key)

    def wait(self, seconds):
        """For given seconds make pointer sleep implicitly.

        :param seconds: time in seconds to wait."""
        self.ps.Wait(int(seconds * 1000))

    def waitForCursor(self, row, column, seconds=60):
        """Wait for the cursor to be in a given position.
        Afert the given seconds, timeout.

        :param row: Row where the cursor should be.
        :param column: Column where the cursor should be.
        :param seconds: time in seconds to wait."""
        return self.ps.WaitForCursor(row, column, int(seconds * 1000))

    def waitForString(self, string, row, column, seconds=60):
        """Wait for the cursor to be in a given position.
        Afert the given seconds, timeout.

        :param row: Row where the cursor should be.
        :param column: Column where the cursor should be.
        :param seconds: time in seconds to wait."""
        return self.ps.WaitForString(string, row, column, int(seconds * 1000))

    def search_text(self, string='', dir=1, row=1, column=1, wait=1):
        """Searches for the first ocurrence of the string.

        :param string: text to search.
        :param dir: 1: forward, 2: backward.
        :param row: row from where to start the search.
        :param column: column from where to start the search.
        :param second: wait time in second.

        :return: None"""
        self._key_event_setup(wait)  
        return self.ps.SearchText(string, dir, row, column)

class Screen:
    """Allow to create screen objects For AS400 session. Screen objects are
    descriptive object of AS400 that describe how the screen looks like.
    Example would be a cursor location for a given screen.

    Various descriptive attributes:

    AddCursorPos: Sets the cursor position for the screen description to the
                  given position.

    AddString: Adds a string at the given location to the screen description.

    Usage:

    >>> cur_screen = Screen()
    >>> cur_screen.describe("cursor", x=10, y=20)
    >>> cur_screen.wait(5)
    >>> "wait up to 5 seconds for cursor to appear on (10, 20)"

    >>> text_screen = Screen()
    >>> text_screen.describe("text", string="Exit", x=10, y=10, case=True)
    >>> text_screen.wait(5)
    >>> "wait up to 5 seconds for Exit (case sensitive) on (10, 10) location"


    Below are extra descriptions method availabel at users despose.

    AddNumFields: Adds the number of fields to the screen description.

    AddNumInputFields: Adds the number of fields to the screen description.

    AddOIAInhibitStatus: Sets the type of OIA monitoring for the screen
                         description.

    AddStringInRect: Adds a string in the given rectangle to the screen
                     description.

    Clear: Removes all description elements from the screen description."""

    def __init__(self, instance):
        """Initialize screen objects, description added by `describe` method."""
        self.instance = instance
        self.screen = CreateObject("PCOMM.autECLScreenDesc")
        self.desc = {}

    def __repr__(self):
        return "<Screen: {}>".format(self.desc)

    def describe(self, dtype, x, y, **kwargs):
        """allows to describe the screen object. allowed methods are
        AddCursorPos and AddString.

        :param dtype: descriptive object type: cursor or text.
        :param x: x-axis location.
        :param y: y-axis location.
        :param kwargs: usable in case of text description. Acceptable parameter
                       are string: str and case: bool.
                       by default text matching is not case sensitive

        :return: None"""
        valid_desc = ["cursor", "text"]
        if dtype not in valid_desc:
            raise ValueError("Only cursor or text object are allowed.")

        if dtype == "cursor":
            self.desc = {"cursor": {"x": x, "y": y}}
            self.screen.AddCursorPos(x, y)
        else:
            string = kwargs.get("string")
            case = kwargs.get("bool", False)
            self.desc = {"text": {"string": string, "x": x, "y": y, "case": case}}
            self.screen.AddString(string, x, y, case)

    def wait(self, second):
        """For given screen object wait for specified time.

        :param instance: pass the instance object.
        :param second: wait time in second.

        :return: None"""
        return self.instance.ps.WaitForScreen(self.screen, second * 1000)
