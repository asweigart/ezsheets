# EZSheets
# By Al Sweigart al@inventwithpython.com

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import googleapiclient

__version__ = '0.0.1'

#SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE = None


class EZSheetsException(Exception):
    pass # This class exists for this module to raise for EZSheets-specific problems.


class Spreadsheet():
    def __init__(self, spreadsheetId):
        self.spreadsheetId = spreadsheetId

        request = SERVICE.spreadsheets().get(spreadsheetId=self.spreadsheetId)
        response = request.execute()

        """
        Example of a response object here:
        {'properties': {'autoRecalc': 'ON_CHANGE',
                        'defaultFormat': {'backgroundColor': {'blue': 1,
                                                              'green': 1,
                                                              'red': 1},
                                          'padding': {'bottom': 2,
                                                      'left': 3,
                                                      'right': 3,
                                                      'top': 2},
                                          'textFormat': {'bold': False,
                                                         'fontFamily': 'arial,sans,sans-serif',
                                                         'fontSize': 10,
                                                         'foregroundColor': {},
                                                         'italic': False,
                                                         'strikethrough': False,
                                                         'underline': False},
                                          'verticalAlignment': 'BOTTOM',
                                          'wrapStrategy': 'LEGACY_WRAP'},
                        'locale': 'en_US',
                        'timeZone': 'America/New_York',
                        'title': 'Copy of Example Spreadsheet'},
         'sheets': [{'properties': {'gridProperties': {'columnCount': 22,
                                                       'frozenRowCount': 1,
                                                       'rowCount': 101},
                                    'index': 0,
                                    'sheetId': 0,
                                    'sheetType': 'GRID',
                                    'title': 'Class Data'}},
                    {'properties': {'gridProperties': {'columnCount': 26,
                                                       'rowCount': 1000},
                                    'index': 1,
                                    'sheetId': 2075929783,
                                    'sheetType': 'GRID',
                                    'title': 'Sheet1'}}],
         'spreadsheetId': '16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c',
         'spreadsheetUrl': 'https://docs.google.com/spreadsheets/d/16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c/edit'}
         """
        #self.locale = response['properties']['locale'] # Leave these commented until it's clear I actually need them.
        #self.timeZone = response['properties']['timeZone']
        self.title = response['properties']['title']

        self.sheets = []
        for sheetInfo in response['sheets']:
            title = sheetInfo['properties']['title']
            sheetId = sheetInfo['properties']['sheetId']
            gp = sheetInfo['properties']['gridProperties'] # syntactic sugar
            additionalArgs = {'rowCount':                gp.get('rowCount'),
                              'columnCount':             gp.get('columnCount'),
                              'frozenRowCount':          gp.get('frozenRowCount'),
                              'frozenColumnCount':       gp.get('frozenColumnCount'),
                              'hideGridlines':           gp.get('hideGridlines'),
                              'rowGroupControlAfter':    gp.get('rowGroupControlAfter'),
                              'columnGroupControlAfter': gp.get('columnGroupControlAfter'),}
            sheet = Sheet(title, sheetId, **additionalArgs)
            self.sheets.append(sheet)

    def __str__(self):
        return '<%s title="%d", %s sheets>' % (type(self).__name__, self.title, len(self.sheets))

    def __repr__(self):
        return '%s(spreadsheetId=%s)' % (type(self).__name__, self.spreadsheetId)

class Sheet():
    def __init__(self, spreadsheetId, title, sheetId, **kwargs):
        self._spreadsheetId = spreadsheetId
        self._title = title
        self._sheetId = sheetId
        self._rowCount                = kwargs.get('rowCount')
        self._columnCount             = kwargs.get('columnCount')
        self._frozenRowCount          = kwargs.get('frozenRowCount')
        self._frozenColumnCount       = kwargs.get('frozenColumnCount')
        self._hideGridlines           = kwargs.get('hideGridlines')
        self._rowGroupControlAfter    = kwargs.get('rowGroupControlAfter')
        self._columnGroupControlAfter = kwargs.get('columnGroupControlAfter')

        self._dataLoaded = False
        self._data = None

    # Set up the read-only attributes.
    @property
    def _spreadsheetId(self):
        return self._spreadsheetId
    @property
    def title(self):
        return self._title
    @property
    def sheetId(self):
        return self._sheetId
    @property
    def rowCount(self):
        return self._rowCount
    @property
    def columnCount(self):
        return self._columnCount
    @property
    def frozenRowCount(self):
        return self._frozenRowCount
    @property
    def frozenColumnCount(self):
        return self._frozenColumnCount
    @property
    def hideGridlines(self):
        return self._hideGridlines
    @property
    def rowGroupControlAfter(self):
        return self._rowGroupControlAfter
    @property
    def columnGroupControlAfter(self):
        return self._columnGroupControlAfter

    def __str__(self):
        return '<%s title=%r, sheetId=%r, rowCount=%r, columnCount=%r>' % (type(self).__name__, self._title, self._sheetId, self._rowCount, self._columnCount)

    def __repr__(self):
        args = ['title=%r' % (self.title), 'sheetId=%r' % (self.sheetId)]
        if self._rowCount is not None:
            args.append('rowCount=%r' % (self._rowCount))
        if self._columnCount is not None:
            args.append('columnCount=%r' % (self._columnCount))
        if self._frozenRowCount is not None:
            args.append('frozenRowCount=%r' % (self._frozenRowCount))
        if self._frozenColumnCount is not None:
            args.append('frozenColumnCount=%r' % (self._frozenColumnCount))
        if self._hideGridlines is not None:
            args.append('hideGridlines=%r' % (self._hideGridlines))
        if self._rowGroupControlAfter is not None:
            args.append('rowGroupControlAfter=%r' % (self._rowGroupControlAfter))
        if self._columnGroupControlAfter is not None:
            args.append('columnGroupControlAfter=%r' % (self._columnGroupControlAfter))

        return '%s(%s)' % (type(self).__name__, ', '.join(args))


    def get(self, column, row):
        if not self._dataLoaded:
            response = SERVICE.spreadsheets().values().get(spreadsheetId='', range='%s!%s:%s').execute()
            self._data = response['values']
            self._majorDimension = response['majorDimension']
            # LEFT OFF


def getColumnLetter(columnNumber):
    """getColumnLetter(1) => 'A', getColumnLetter(27) => 'AA'"""
    letters = []
    while columnNumber > 0:
        columnNumber, remainder = divmod(columnNumber, 26)
        if remainder == 0:
            remainder = 26
            columnNumber -= 1
        letters.append(chr(remainder + 64))
    return ''.join(reversed(letters))

def getColumnNumber(columnLetter):
    """getColumnNumber('A') => 1, getColumnNumber('AA') => 27"""
    columnLetter = columnLetter.upper()
    digits = []

    while columnLetter:
        digits.append(ord(columnLetter[0]) - 64)
        columnLetter = columnLetter[1:]

    number = 0
    place = 0
    for digit in reversed(digits):
        number += digit * (26 ** place)
        place += 1

    return number







def init(tokenFile='token.pickle', credentialsFile='credentials.json'):
    global SERVICE

    if not os.path.exists(credentialsFile):
        raise EZSheetsException('Can\'t find credentials file at %s. You can download this file from https://developers.google.com/gmail/api/quickstart/python and clicking "Enable the Gmail API"' % (os.path.abspath(credentialsFile)))

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    SERVICE = build('sheets', 'v4', credentials=creds)

init()
