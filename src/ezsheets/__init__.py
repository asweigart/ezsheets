# EZSheets
# By Al Sweigart al@inventwithpython.com

import pickle, copy
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import googleapiclient

try:
    # For Python 3:
    from collections.abc import Iterable
except:
    # For Python 2:
    from collections import Iterable


__version__ = '0.0.1'

#SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE = None

from ezsheets.colorvalues import COLORS

"""
Features to add:

- get values from cells
- modify values in cells
- change title of spreadsheet and sheets
- download as csv/excel/whatever

Done:
- get sheet names
"""


# Sample spreadsheet id: 16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c


class EZSheetsException(Exception):
    pass # This class exists for this module to raise for EZSheets-specific problems.


class Spreadsheet():
    def __init__(self, spreadsheetId):
        self._spreadsheetId = getIdFromUrl(spreadsheetId)
        self.refresh()

    def refresh(self):
        request = SERVICE.spreadsheets().get(spreadsheetId=self._spreadsheetId)
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
        self._title = response['properties']['title']

        self.sheets = []
        for i, sheetInfo in enumerate(response['sheets']):
            title = sheetInfo['properties']['title']
            sheetId = sheetInfo['properties']['sheetId']
            gp = sheetInfo['properties']['gridProperties'] # syntactic sugar
            additionalArgs = {'index':                   i,
                              'rowCount':                gp.get('rowCount'),
                              'columnCount':             gp.get('columnCount'),
                              'frozenRowCount':          gp.get('frozenRowCount'),
                              'frozenColumnCount':       gp.get('frozenColumnCount'),
                              'hideGridlines':           gp.get('hideGridlines'),
                              'rowGroupControlAfter':    gp.get('rowGroupControlAfter'),
                              'columnGroupControlAfter': gp.get('columnGroupControlAfter'),}
            sheet = Sheet(self, title, sheetId, **additionalArgs)
            self.sheets.append(sheet)
        self.sheets = tuple(self.sheets) # Make sheets attribute an immutable tuple.

    @property
    def spreadsheetId(self):
        return self._spreadsheetId

    @property
    def sheetTitles(self):
        return tuple([sheet.title for sheet in self.sheets])

    def __str__(self):
        return '<%s title="%d", %s sheets>' % (type(self).__name__, self.title, len(self.sheets))

    def __repr__(self):
        return '%s(spreadsheetId=%s)' % (type(self).__name__, self.spreadsheetId)

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, value):
        value = str(value)
        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheetId,
        body={
            'requests': [{'updateSpreadsheetProperties': {'properties': {'title': value},
                                                          'fields': 'title'}}],
            'includeSpreadsheetInResponse': False,
            'responseRanges': [''], # This value is meaningful only if includeSpreadsheetInResponse is True
            'responseIncludeGridData': False # This value is meaningful only if includeSpreadsheetInResponse is True
        })
        request.execute()
        s._title = value


class Sheet():
    def __init__(self, spreadsheet, title, sheetId, **kwargs):
        self._spreadsheet = spreadsheet
        self._title = title
        self._sheetId = sheetId
        self._index                   = kwargs.get('index')
        self._rowCount                = kwargs.get('rowCount')
        self._columnCount             = kwargs.get('columnCount')
        self._frozenRowCount          = kwargs.get('frozenRowCount')
        self._frozenColumnCount       = kwargs.get('frozenColumnCount')
        self._hideGridlines           = kwargs.get('hideGridlines')
        self._rowGroupControlAfter    = kwargs.get('rowGroupControlAfter')
        self._columnGroupControlAfter = kwargs.get('columnGroupControlAfter')

        self._dataIsFresh = False
        self._data = None

    # Set up the read-only attributes.
    @property
    def spreadsheet(self):
        return self._spreadsheet

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, value):
        value = str(value)
        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet.spreadsheetId,
        body={
            'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
                                                                   'title': value},
                                                    'fields': 'title'}}],
            'includeSpreadsheetInResponse': False,
            'responseRanges': [''], # This value is meaningful only if includeSpreadsheetInResponse is True
            'responseIncludeGridData': False # This value is meaningful only if includeSpreadsheetInResponse is True
        })
        request.execute()
        self._title = value

    @property
    def tabColor(self):
        return self._tabColor

    @tabColor.setter
    def tabColor(self, value):

        if isinstance(value, str):
            tabColorArg = {
                'red': COLORS[value][0],
                'green': COLORS[value][1],
                'blue': COLORS[value][2],
                'alpha': COLORS[value][3],
            }

        #elif value is None: # TODO - apparently there's no way to reset the color through the api?
        #    tabColorArg = {} # Reset the color
        else:
            tabColorArg = {
                'red': float(value[0]),
                'green': float(value[1]),
                'blue': float(value[2]),
            }
            try:
                tabColorArg['alpha'] = value[3]
            except:
                tabColorArg['alpha'] = 1.0

        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet.spreadsheetId,
        body={
            'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
                                                                   'tabColor': tabColorArg},
                                                    'fields': 'tabColor'}}],
            'includeSpreadsheetInResponse': False,
            'responseRanges': [''], # This value is meaningful only if includeSpreadsheetInResponse is True
            'responseIncludeGridData': False # This value is meaningful only if includeSpreadsheetInResponse is True
        })
        request.execute()
        self._tabColor = value


    @property
    def index(self):
        return self._index


    @index.setter
    def index(self, value):
        if not isinstance(value, int):
            raise TypeError('indices must be integers, not %s' % (type(value).__name__))
        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
        body={
            'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
                                                                   'index': value},
                                                    'fields': 'index'}}],
            'includeSpreadsheetInResponse': False,
            'responseRanges': [''], # This value is meaningful only if includeSpreadsheetInResponse is True
            'responseIncludeGridData': False # This value is meaningful only if includeSpreadsheetInResponse is True
        })
        request.execute()
        self._spreadsheet.refresh() # Update the spreadsheet's tuple of Sheet objects to reflect the new order.


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
        if not isinstance(column, int):
            raise TypeError('column indices must be integers, not %s' % (type(column).__name__))
        if not isinstance(row, int):
            raise TypeError('row indices must be integers, not %s' % (type(row).__name__))
        if column < 1 or row < 1:
            raise IndexError('Column %s, row %s does not exist. Google Sheets\' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index. Negative indices are not supported by ezsheets.' % (column, row))

        if not self._dataIsFresh: # Download the sheet, if it hasn't been downloaded already.
            self.refresh()

        try:
            if self._majorDimension == 'ROWS':
                return self._cells[row-1][column-1] # -1 because _cells is 0-based while Google Sheets is 1-based.
            elif self._majorDimension == 'COLUMNS':
                return self._cells[column-1][row-1]
        except IndexError:
            return ''


    def getAll(self):
        if not self._dataIsFresh: # Download the sheet, if it hasn't been downloaded already.
            self.refresh()

        if self._majorDimension == 'ROWS':
            return copy.deepcopy(self._cells)
        elif self._majorDimension == 'COLUMNS':
            # self._cells' inner lists represent columns, not rows. But getAll() should always return a ROWS-major dimensioned structure.
            cells = []

            longestColumnLength = max([len(column) for column in self._cells])
            for rowIndex in range(longestColumnLength):
                rowList = []
                for columnData in self._cells:
                    if rowIndex < len(columnData):
                        rowList.append(columnData[rowIndex])
                    else:
                        rowList.append('')
                cells.append(rowList)

            return cells
        else:
            assert False, 'self._majorDimension is set to %s instead of "ROWS" or "COLUMNS"' % (self._majorDimension)


    def getRow(self, row):
        if not isinstance(row, int):
            raise TypeError('row indices must be integers, not %s' % (type(row).__name__))
        if row < 1:
            raise IndexError('Row %s does not exist. Google Sheets\' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index.' % (row))

        if not self._dataIsFresh: # Download the sheet, if it hasn't been downloaded already.
            self.refresh()

        if self._majorDimension == 'ROWS':
            try:
                return self._cells[row-1] # -1 because _cells is 0-based while Google Sheets is 1-based.
            except IndexError:
                return []
        elif self._majorDimension == 'COLUMNS':
            rowList = []
            for row in range(len(self._cells)):
                if (len(self._cells[row]) == 0) or (row-1 >= len(self._cells[row])):
                    rowList.append('')
                else:
                    rowList.append(self._cells[row][row-1]) # rows don't have -1 because they're based on _cells which is a 0-based Python lists.
            return rowList


    def getColumn(self, column):
        if not isinstance(column, (int, str)):
            raise TypeError('column indices must be integers or str, not %s' % (type(column).__name__))
        if isinstance(column, int) and column < 1:
            raise IndexError('Column %s does not exist. Google Sheets\' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index.' % (column))
        if isinstance(column, str) and not column.isalpha():
            raise ValueError('Column %s does not exist. Columns must be a 1-based int or a letters-only str.')

        if isinstance(column, str):
            column = getColumnNumber(column)

        if not self._dataIsFresh: # Download the sheet, if it hasn't been downloaded already.
            self.refresh()

        if self._majorDimension == 'COLUMNS':
            try:
                return self._cells[column-1] # -1 because _cells is 0-based while Google Sheets is 1-based.
            except IndexError:
                return []
        elif self._majorDimension == 'ROWS':
            columnList = []
            for row in range(len(self._cells)):
                if (len(self._cells[row]) == 0) or (column-1 >= len(self._cells[row])):
                    columnList.append('')
                else:
                    columnList.append(self._cells[row][column-1]) # columns don't have -1 because they're based on _cells which is a 0-based Python lists.
            return columnList


    def refresh(self):
        request = SERVICE.spreadsheets().values().get(
            spreadsheetId=self._spreadsheet._spreadsheetId,
            range='%s!A1:%s%s' % (self._title, getColumnLetterOf(self._columnCount), self._rowCount))
        response = request.execute()

        self._cells = response['values']
        self._majorDimension = response['majorDimension']

        self._dataIsFresh = True


    def update(self, row, column, value):
        if not isinstance(column, int):
            raise TypeError('column indices must be integers, not %s' % (type(column).__name__))
        if not isinstance(row, int):
            raise TypeError('row indices must be integers, not %s' % (type(row).__name__))
        if column < 1 or row < 1:
            raise IndexError('Column %s, row %s does not exist. Google Sheets\' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index. Negative indices are not supported by ezsheets.' % (column, row))

        cellLocation = getColumnLetterOf(column) + str(row)
        request = SERVICE.spreadsheets().values().update(
            spreadsheetId=self._spreadsheet._spreadsheetId,
            range='%s!%s:%s' % (self._title, cellLocation, cellLocation),
            valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
            body={
                'majorDimension': 'ROWS',
                'values': [[value]],
                'range': '%s!%s:%s' % (self._title, cellLocation, cellLocation),
                }
            )
        request.execute()

        self._dataIsFresh = False


    def updateRow(self, row, values):
        if not isinstance(row, int):
            raise TypeError('row indices must be integers, not %s' % (type(row).__name__))
        if row < 1:
            raise IndexError('Row %s does not exist. Google Sheets\' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index.' % (row))
        if not isinstance(values, Iterable):
            raise TypeError('values must be an iterable, not %s' % (type(values).__name__))

        if len(values) < self._columnCount:
            values.extend([''] * (self._columnCount - len(values)))

        request = SERVICE.spreadsheets().values().update(
            spreadsheetId=self._spreadsheet._spreadsheetId,
            range='%s!A%s:%s%s' % (self._title, row, getColumnLetterOf(len(values)), row),
            valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
            body={
                'majorDimension': 'ROWS',
                'values': [values],
                'range': '%s!A%s:%s%s' % (self._title, row, getColumnLetterOf(len(values)), row),
                }
            )
        request.execute()

        self._dataIsFresh = False


    def updateColumn(self, column, values):
        if not isinstance(column, (int, str)):
            raise TypeError('column indices must be integers, not %s' % (type(column).__name__))
        if isinstance(column, int) and column < 1:
            raise IndexError('Column %s does not exist. Google Sheets\' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index.' % (column))
        if not isinstance(values, Iterable):
            raise TypeError('values must be an iterable, not %s' % (type(values).__name__))
        if isinstance(column, str) and not column.isalpha():
            raise ValueError('Column %s does not exist. Columns must be a 1-based int or a letters-only str.')

        if isinstance(column, str):
            column = getColumnNumber(column)

        if len(values) < self._rowCount:
            values.extend([''] * (self._rowCount - len(values)))

        request = SERVICE.spreadsheets().values().update(
            spreadsheetId=self._spreadsheet._spreadsheetId,
            range='%s!%s1:%s%s' % (self._title, getColumnLetterOf(column), getColumnLetterOf(column), len(values)),
            valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
            body={
                'majorDimension': 'COLUMNS',
                'values': [values],
                'range': '%s!%s1:%s%s' % (self._title, getColumnLetterOf(column), getColumnLetterOf(column), len(values)),
                }
            )
        response = request.execute()
        print(response)

        self._dataIsFresh = False

    def updateAll(self, values):
        # Ensure that `values` is a list of lists:


        # Find out the dimensions of `values`
        valRowCount = len(values)
        #valColumnCount = max([len(row) for row in values])
        # LEFT OFF

        if valRowCount < self._rowCount:
            values.extend([''] * (self._rowCount - valRowCount))

        for row in values:
            row.extend([''] * (self.__columnCount - len(row)))



        if self._majorDimension == 'ROWS':
            return copy.deepcopy(self._cells)
        elif self._majorDimension == 'COLUMNS':
            # self._cells' inner lists represent columns, not rows. But getAll() should always return a ROWS-major dimensioned structure.
            cells = []

            longestColumnLength = max([len(column) for column in self._cells])
            for rowIndex in range(longestColumnLength):
                rowList = []
                for columnData in self._cells:
                    if rowIndex < len(columnData):
                        rowList.append(columnData[rowIndex])
                    else:
                        rowList.append('')
                cells.append(rowList)

            return cells
        else:
            assert False, 'self._majorDimension is set to %s instead of "ROWS" or "COLUMNS"' % (self._majorDimension)


    def downloadAsCSV(self):
        pass # TODO
    def downloadAsExcel(self):
        pass # TODO
    def downloadAsODS(self):
        pass # TODO
    def downloadAsPDF(self):
        pass # TODO
    def downloadAsHTML(self):
        pass # TODO
    def downloadAsTSV(self):
        pass # TODO


def getIdFromUrl(url):
    # https://docs.google.com/spreadsheets/d/16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c/edit#gid=0
    if url.startswith('https://docs.google.com/spreadsheets/d/'):
        return url[39:url.find('/', 39)]
    else:
        return url

def getColumnLetterOf(columnNumber):
    """getColumnLetterOf(1) => 'A', getColumnLetterOf(27) => 'AA'"""
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
s = Spreadsheet('16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c')