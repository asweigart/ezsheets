# EZSheets
# By Al Sweigart al@inventwithpython.com

import pickle, copy, re
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

__version__ = '0.0.2'

#SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE = None
IS_INITIALIZED = False

DEFAULT_NEW_ROW_COUNT = 1000  # This is the Google Sheets default for a new Sheet.
DEFAULT_NEW_COLUMN_COUNT = 26 # This is the Google Sheets default for a new Sheet.
DEFAULT_FROZEN_ROW_COUNT = 0
DEFAULT_FROZEN_COLUMN_COUNT = 0
DEFAULT_HIDE_GRID_LINES = False
DEFAULT_ROW_GROUP_CONTROL_AFTER = False
DEFAULT_COLUMN_GROUP_CONTROL_AFTER = False

from ezsheets.colorvalues import COLORS

"""
Features to add:
- delete spreadsheets
- download as csv/excel/whatever
"""


# Sample spreadsheet id: 16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c


class EZSheetsException(Exception):
    pass # This class exists for this module to raise for EZSheets-specific problems.


class Spreadsheet():
    def __init__(self, spreadsheetId):
        if not IS_INITIALIZED: init() # Initialize this module if not done so already.

        self._dataIsFresh = False

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

        # TODO!!!!!!!! Every time we call refresh, we're recreating the Sheet objects. BAD!!!!

        sheets = []
        if not hasattr(self, 'sheets'):
            self.sheets = ()
        for i, sheetInfo in enumerate(response['sheets']):
            sheetId = sheetInfo['properties']['sheetId']
            gp = sheetInfo['properties']['gridProperties'] # syntactic sugar
            additionalArgs = {'title':                   sheetInfo['properties']['title'],
                              'index':                   i,
                              'tabColor':                _getTabColorArg(sheetInfo['properties'].get('tabColor')),
                              'rowCount':                gp.get('rowCount', DEFAULT_NEW_ROW_COUNT),
                              'columnCount':             gp.get('columnCount', DEFAULT_NEW_COLUMN_COUNT),
                              'frozenRowCount':          gp.get('frozenRowCount', 0),
                              'frozenColumnCount':       gp.get('frozenColumnCount', 0),
                              'hideGridlines':           gp.get('hideGridlines'), # TODO - add default
                              'rowGroupControlAfter':    gp.get('rowGroupControlAfter'), # TODO - add default
                              'columnGroupControlAfter': gp.get('columnGroupControlAfter'),} # TODO - add default

            existingSheetIndex = None
            try:
                # If the sheet has been previously loaded, reuse that Sheet object:
                existingSheetIndex = [sh.sheetId for sh in self.sheets].index(sheetId)
            except ValueError:
                pass # Do nothing.

            if existingSheetIndex is not None:
                # Update the info in the Sheet object:
                for k, v in additionalArgs.items():
                    setattr(self.sheets[existingSheetIndex], '_' + k, v) # Set the _backing variable for the property directly, otherwise it causes an infinite loop if we try to set the property.
                sheets.append(self.sheets[existingSheetIndex])
            else:
                # If the sheet hasn't been seen before, create a new Sheet object:
                sheets.append(Sheet(self, sheetId, **additionalArgs))

        self.sheets = tuple(sheets) # Make sheets attribute an immutable tuple.
        self._dataIsFresh = True

    def __getitem__(self, key):
        try:
            i = self.sheetTitles.index(key)
            return self.sheets[i]
        except ValueError:
            pass # Do nothing if the title isn't found.


        if isinstance(key, int) and (-len(self.sheets) <= key < len(self.sheets)):
            return self.sheets[key]
        if isinstance(key, slice):
            return self.sheets[key]

        raise KeyError('key must be an int between %s and %s or a str matching a title: %r' % (-(len(self.sheets)), len(self.sheets) - 1, self.sheetTitles))

    def __delitem__(self, key):
        if isinstance(key, (int, str)):
            # Key is an int index or a str title.
            self[key].delete()
        elif isinstance(key, slice):
            # TODO - there's got to be a better way to do this.
            start = key.start if key.start is not None else 0
            stop  = key.stop  if key.stop  is not None else len(self.sheets)
            step  = key.step  if key.step  is not None else 1

            if start < 0 or stop < 0:
                return # When deleting list items with a slice, a negative start or stop results in a no-op. I'll mimic that behavior here.

            indexesToDelete = [i for i in range(start, stop, step) if i >= 0 and i < len(self.sheets)] # Don't include invalid or negative indexes.
            if len(indexesToDelete) == len(self.sheets):
                raise ValueError('Cannot delete all sheets; spreadsheets must have at least one sheet')

            if indexesToDelete[0] < indexesToDelete[-1]:
                indexesToDelete.reverse() # We want this is descending order.

            for i in indexesToDelete:
                self.sheets[i].delete()

        else:
            raise TypeError('key must be an int index, str sheet title, or slice object, not %r' % (type(key).__name__))

    def __len__(self):
        return len(self.sheets)

    def __iter__(self):
        return iter(self.sheets)

    @property
    def spreadsheetId(self):
        return self._spreadsheetId

    @property
    def sheetTitles(self):
        return tuple([sheet.title for sheet in self.sheets])

    def __str__(self):
        return '<%s title="%s", %d sheets>' % (type(self).__name__, self.title, len(self.sheets))

    def __repr__(self):
        return '%s(spreadsheetId=%r)' % (type(self).__name__, self.spreadsheetId)

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, value):
        value = str(value)
        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheetId,
        body={
            'requests': [{'updateSpreadsheetProperties': {'properties': {'title': value},
                                                          'fields': 'title'}}]})
        request.execute()
        self._title = value


    def addSheet(self, title='', rowCount=DEFAULT_NEW_ROW_COUNT, columnCount=DEFAULT_NEW_COLUMN_COUNT, index=None, tabColor=None):
        if not self._dataIsFresh: # Download the sheet, if it hasn't been downloaded already.
            self.refresh()

        if index is None:
            # Set the index to make this new sheet be the last sheet:
            index = len(self.sheets)

        propertiesDictValue = {'title': title,
                               'index': index,
                               'gridProperties' : {'rowCount': rowCount,
                                                   'columnCount': columnCount}}
        if tabColor is not None:
            propertiesDictValue['tabColor'] = _getTabColorArg(tabColor)

        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheetId,
        body={
            'requests': [{'addSheet': {'properties': propertiesDictValue}}]})
        request.execute()

        self.refresh()
        return self.sheets[index]



class Sheet():
    def __init__(self, spreadsheet, sheetId, **kwargs):
        #if not IS_INITIALIZED: init() # Initialize this module if not done so already. # This line might not be needed? Sheet objects can only exist when you've already made a Spreadsheet object.

        self._spreadsheet = spreadsheet
        self._sheetId = sheetId
        self._title                   = kwargs.get('title')
        self._index                   = kwargs.get('index')
        self._rowCount                = kwargs.get('rowCount')
        self._columnCount             = kwargs.get('columnCount')
        self._frozenRowCount          = kwargs.get('frozenRowCount')
        self._frozenColumnCount       = kwargs.get('frozenColumnCount')
        self._hideGridlines           = kwargs.get('hideGridlines')
        self._rowGroupControlAfter    = kwargs.get('rowGroupControlAfter')
        self._columnGroupControlAfter = kwargs.get('columnGroupControlAfter')

        if kwargs.get('tabColor') is None:
            self._tabColor = None
        else:
            self._tabColor = _getTabColorArg(kwargs.get('tabColor'))

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
                                                    'fields': 'title'}}]})
        request.execute()
        self._title = value


    @property
    def tabColor(self):
        return self._tabColor

    @tabColor.setter
    def tabColor(self, value):
        tabColorArg = _getTabColorArg(value)

        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet.spreadsheetId,
        body={
            'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
                                                                   'tabColor': tabColorArg},
                                                    'fields': 'tabColor'}}]})
        request.execute()
        self._tabColor = tabColorArg


    @property
    def index(self):
        return self._index


    @index.setter
    def index(self, value):
        if not isinstance(value, int):
            raise TypeError('indices must be integers, not %s' % (type(value).__name__))

        if value < 0: # Handle negative indexes the way Python lists do.
            if value < -len(self.spreadsheet.sheets):
                raise IndexError('%r is out of range (-1 to %d)' % (value, -len(self.spreadsheet.sheets)))
            value = len(self.spreadsheet.sheets) + value # convert this negative index into its corresponding positive index
        if value >= len(self.spreadsheet.sheets):
            raise IndexError('%r is out of range (0 to %d)' % (value, len(self.spreadsheet.sheets) - 1))


        if value == self.index:
            return # No change needed.
        if value > self.index:
            value += 1 # Google Sheets uses "before the move" indexes, which is confusing and I don't want to do it here.

        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
        body={
            'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
                                                                   'index': value},
                                                    'fields': 'index'}}]})
        request.execute()

        self._spreadsheet.refresh() # Update the spreadsheet's tuple of Sheet objects to reflect the new order.
        self._index = self._spreadsheet.sheets.index(self) # Update the local Sheet object's index.


    def __eq__(self, other):
        if not isinstance(other, Sheet):
            return False
        return self._sheetId == other._sheetId

    @property
    def sheetId(self):
        return self._sheetId


    @property
    def rowCount(self):
        return self._rowCount

    @rowCount.setter
    def rowCount(self, value):
        # Validate arguments:
        if not isinstance(value, int):
            raise TypeError('value arg must be an int, not %s' % (type(value).__name__))
        if value < 1:
            raise TypeError('value arg must be a positive nonzero int, not %r' % (value))

        self._refreshGridProperties() # Retrieve up-to-date grid properties from Google Sheets.
        self._rowCount = value        # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.


    @property
    def columnCount(self):
        return self._columnCount


    @columnCount.setter
    def columnCount(self, value):
        # Validate arguments:
        if not isinstance(value, int):
            raise TypeError('value arg must be an int, not %s' % (type(value).__name__))
        if value < 1:
            raise TypeError('value arg must be a positive nonzero int, not %r' % (value))

        self._refreshGridProperties() # Retrieve up-to-date grid properties from Google Sheets.
        self._columnCount = value     # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.


    @property
    def frozenRowCount(self):
        return self._frozenRowCount


    @frozenRowCount.setter
    def frozenRowCount(self, value):
        # Validate arguments:
        if not isinstance(value, int):
            raise TypeError('value arg must be an int, not %s' % (type(value).__name__))
        if value < 1:
            raise TypeError('value arg must be a positive nonzero int, not %r' % (value))

        self._refreshGridProperties() # Retrieve up-to-date grid properties from Google Sheets.
        self._frozenRowCount = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.


    @property
    def frozenColumnCount(self):
        return self._frozenColumnCount


    @frozenColumnCount.setter
    def frozenColumnCount(self, value):
        # Validate arguments:
        if not isinstance(value, int):
            raise TypeError('value arg must be an int, not %s' % (type(value).__name__))
        if value < 1:
            raise TypeError('value arg must be a positive nonzero int, not %r' % (value))

        self._refreshGridProperties() # Retrieve up-to-date grid properties from Google Sheets.
        self._frozenRowCount = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.


    @property
    def hideGridlines(self):
        return self._hideGridlines


    @hideGridlines.setter
    def hideGridlines(self, value):
        value = bool(value)

        self._refreshGridProperties() # Retrieve up-to-date grid properties from Google Sheets.
        self._hideGridlines = value   # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.

    @property
    def rowGroupControlAfter(self):
        return self._rowGroupControlAfter


    @rowGroupControlAfter.setter
    def rowGroupControlAfter(self, value):
        value = bool(value)

        self._refreshGridProperties() # Retrieve up-to-date grid properties from Google Sheets.
        self._rowGroupControlAfter = value # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.


    @property
    def columnGroupControlAfter(self):
        return self._columnGroupControlAfter


    @columnGroupControlAfter.setter
    def columnGroupControlAfter(self, value):
        value = bool(value)

        self._refreshGridProperties() # Retrieve up-to-date grid properties from Google Sheets.
        self._columnGroupControlAfter = value # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.


    def __str__(self):
        return '<%s title=%r, sheetId=%r, rowCount=%r, columnCount=%r>' % (type(self).__name__, self._title, self._sheetId, self._rowCount, self._columnCount)


    def __repr__(self):
        return '%s(sheetId=%r, title=%r, rowCount=%r, columnCount=%r' % (type(self).__name__, self.sheetId, self._title, self._rowCount, self._columnCount)


    def get(self, *args):
        if len(args) == 2: # args are column, row like (2, 5)
            column, row = args
        elif len(args) == 1: # args is a string of a grid cell like ('B5',)
            column, row = convertToColumnRowInts(args[0])
        else:
            raise TypeError("get() takes one or two arguments, like ('A1',) or (2, 5)")


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


    def getRows(self, startRow=0, stopRow=None, step=1):
        if stopRow is None:
            stopRow = self._rowCount + 1

        return [self.getRow(rowNum) for rowNum in range(startRow, stopRow,  step)]


    def __getitem__(self, key):
        if isinstance(key, int) and (-len(self.sheets) <= key < len(self.sheets)):
            return self.sheets[key]
        if isinstance(key, slice):
            start = key.start if key.start is not None else 0
            stop =  key.stop # if key.stop is None, then pass it anyways. It's fine.
            step = key.step if key.step is not None else 1
            return self.getRows(startRow=start, stopRow=stop, step=step)

        raise KeyError('key must be an int between %s and %s or a str matching a title: %r' % (-(len(self.sheets)), len(self.sheets) - 1, self.sheetTitles))


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
        # Refresh the local sheet data based on what's on Google Sheets.
        self._refreshGridProperties()

        request = SERVICE.spreadsheets().values().get(
            spreadsheetId=self._spreadsheet._spreadsheetId,
            range='%s!A1:%s%s' % (self._title, getColumnLetterOf(self._columnCount), self._rowCount))
        response = request.execute()

        if 'values' in response:
            self._cells = response['values']
        else:
            # There's no 'values' key, it's because there's no data in the sheet.
            self._cells = [[]]

        self._majorDimension = response['majorDimension']

        self._dataIsFresh = True


    def _refreshGridProperties(self):
        for sheetDict in SERVICE.spreadsheets().get(spreadsheetId=self._spreadsheet._spreadsheetId).execute()['sheets']:
            if sheetDict['properties']['sheetId'] == self._sheetId: # Find this sheet in the returned spreadsheet json data.
                gridProps = sheetDict['properties']['gridProperties']
                self._rowCount                = gridProps.get('rowCount', DEFAULT_NEW_ROW_COUNT)
                self._columnCount             = gridProps.get('columnCount', DEFAULT_NEW_COLUMN_COUNT)
                self._frozenRowCount          = gridProps.get('frozenRowCount', DEFAULT_FROZEN_ROW_COUNT)
                self._frozenColumnCount       = gridProps.get('frozenColumnCount', DEFAULT_FROZEN_COLUMN_COUNT)
                self._hideGridlines           = gridProps.get('hideGridlines', DEFAULT_HIDE_GRID_LINES)
                self._rowGroupControlAfter    = gridProps.get('rowGroupControlAfter', DEFAULT_ROW_GROUP_CONTROL_AFTER)
                self._columnGroupControlAfter = gridProps.get('columnGroupControlAfter', DEFAULT_COLUMN_GROUP_CONTROL_AFTER)
                return # No need to check the rest of the sheets.

    def _updateGridProperties(self):
        gridProperties = {'rowCount':                self._rowCount,
                          'columnCount':             self._columnCount,
                          'frozenRowCount':          self._frozenRowCount,
                          'frozenColumnCount':       self._frozenColumnCount,
                          'hideGridlines':           self._hideGridlines,
                          'rowGroupControlAfter':    self._rowGroupControlAfter,
                          'columnGroupControlAfter': self._columnGroupControlAfter}
        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
            body={
            'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
                                                                   'gridProperties': gridProperties},
                                                    'fields': 'gridProperties'}}]})
        request.execute()


    def update(self, *args):
        if len(args) == 3: # args are column, row like (2, 5)
            column, row, value = args
        elif len(args) == 2: # args is a string of a grid cell like ('B5',)
            column, row = convertToColumnRowInts(args[0])
            value = args[1]
        else:
            raise TypeError("get() takes one or two arguments, like ('A1',) or (2, 5)")

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
        if not isinstance(values, (list, tuple)):
            raise TypeError('values must be a list or tuple, not %s' % (type(values).__name__))

        if isinstance(values, tuple):
            values = list(values)
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
        if not isinstance(values, (list, tuple)):
            raise TypeError('values must be a list or tuple, not %s' % (type(values).__name__))
        if isinstance(column, str) and not column.isalpha():
            raise ValueError('Column %s does not exist. Columns must be a 1-based int or a letters-only str.')

        if isinstance(values, tuple):
            values = list(values)
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
        request.execute()

        self._dataIsFresh = False

    def updateAll(self, values):
        # Ensure that `values` is a list of lists:
        if not isinstance(values, list):
            if isinstance(values, tuple):
                values = list(values)
            else:
                raise TypeError('values arg must be a list or tuple, not %s' % (type(values).__name__))
            for i, innerList in enumerate(values):
                if not isinstance(innerList, list):
                    if isinstance(innerList, tuple):
                        values[i] = list(innerList)
                    else:
                        raise TypeError('values[%r] must be a list or tuple, not %s' % (type(innerList).__name__))


        # Find out the dimensions of `values`, lengthen them to the size of the sheet if needed.
        valRowCount = len(values)
        if valRowCount < self._rowCount:
            values.extend([[''] for i in range(self._rowCount - valRowCount)])
        for row in values:
            row.extend([''] * (self._columnCount - len(row)))

        if self._majorDimension == 'ROWS':
            pass # Nothing needs to be done if majorDimension is already 'ROWS'
        elif self._majorDimension == 'COLUMNS':
            # self._cells' inner lists represent columns, not rows. But getAll() should always return a ROWS-major dimensioned structure.
            cells = []

            longestColumnLength = max([len(column) for column in self._cells])
            for rowIndex in range(longestColumnLength):
                rowList = [] # create the data for the row
                for columnData in self._cells:
                    if rowIndex < len(columnData):
                        rowList.append(columnData[rowIndex])
                    else:
                        rowList.append('')
                cells.append(rowList)
            values = cells
        else:
            assert False, 'self._majorDimension is set to %r instead of "ROWS" or "COLUMNS"' % (self._majorDimension)

        # Send the API request that updates the Google sheet.
        request = SERVICE.spreadsheets().values().update(
            spreadsheetId=self._spreadsheet._spreadsheetId,
            range='%s!A1:%s%s' % (self._title, getColumnLetterOf(len(values[0])), len(values)),
            valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
            body={
                'majorDimension': 'ROWS',
                'values': values,
                'range': '%s!A1:%s%s' % (self._title, getColumnLetterOf(len(values[0])), len(values)),
                }
            )
        response = request.execute()
        print(response)

        self._dataIsFresh = False


    def copyTo(self, destinationSpreadsheetId):
        request = SERVICE.spreadsheets().sheets().copyTo(spreadsheetId=self._spreadsheet._spreadsheetId,
                                                         sheetId=self._sheetId,
                                                         body={'destinationSpreadsheetId': destinationSpreadsheetId})
        request.execute()

    def delete(self):
        if len(self._spreadsheet.sheets) == 1:
            raise ValueError('Cannot delete all sheets; spreadsheets must have at least one sheet')

        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
            body={
                'requests': [{'deleteSheet': {'sheetId': self._sheetId}}]})
        request.execute()
        self._dataIsFresh = False
        self._spreadsheet.refresh() # Refresh the spreadsheet's list of sheets.


    def resize(self, rowCount=None, columnCount=None):
        # NOTE: If you try to specify the rowCount without the columnCount
        # (and vice versa), Google Sheets thinks you want to set the
        # columnCount to 0 and then complains that you can't delete all the
        # columns.
        # We have a resize() method so that the user doesn't set the row/column
        # count back to the local setting in this Sheet object when it has
        # been changed on Google Sheets by another user. The rowCount and
        # columnCount property setters will make a request to get the current
        # sizes so they don't mistakenly change the other dimension, but
        # this won't be an atomic operation like resize() is.

        # As of Feb 2019, Google Sheets has a cell max of 5,000,000, but
        # this could change so ezsheets won't catch it.

        # Google Sheets size limits are documented here:
        #   https://support.google.com/drive/answer/37603?hl=en
        #   https://www.quora.com/What-are-the-limits-of-Google-Sheets
        if rowCount is None and columnCount is None:
            return # No resizing is taking place, so this function is a no-op.

        # A None value means "use the current setting"
        if rowCount is None:
            rowCount = self._rowCount
        if columnCount is None:
            columnCount = self._columnCount

        if isinstance(columnCount, str):
            columnCount = getColumnNumber(columnCount)

        if not isinstance(rowCount, int):
            raise TypeError('rowCount arg must be an int, not %s' % (type(rowCount).__name__))
        if not isinstance(columnCount, int):
            raise TypeError('columnCount arg must be an int, not %s' % (type(columnCount).__name__))

        if rowCount < 1:
            raise TypeError('rowCount arg must be a positive nonzero int, not %r' % (rowCount))
        if columnCount < 1:
            raise TypeError('columnCount arg must be a positive nonzero int, not %r' % (columnCount))


        request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
        body={
            'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
                                                                   'gridProperties': {'rowCount': rowCount,
                                                                                      'columnCount': columnCount}},
                                                    'fields': 'gridProperties'}}]})
        request.execute()
        self._rowCount = rowCount
        self._columnCount = columnCount

    def __iter__(self):
        return iter(self.getRows())

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


def _getTabColorArg(value):
    if isinstance(value, str) and value in COLORS:
        # value is a color string from colorvalues.py, like 'red' or 'black'
        tabColorArg = {
            'red':   COLORS[value][0],
            'green': COLORS[value][1],
            'blue':  COLORS[value][2],
            'alpha': COLORS[value][3],
        }

    #elif value is None: # TODO - apparently there's no way to reset the color through the api?
    #    tabColorArg = {} # Reset the color
    elif isinstance(value, (list, tuple)) and len(value) in (3, 4):
        # value is a tuple of three or four floats (ranged from 0.0 to 1.0)
        tabColorArg = {
            'red': float(value[0]),
            'green': float(value[1]),
            'blue': float(value[2]),
        }
        try:
            tabColorArg['alpha'] = value[3]
        except:
            tabColorArg['alpha'] = 1.0
    elif value is None:
        return None # Represents no tabColor setting.
    elif type(value) == dict:
        tabColorArg = value
    else:
        raise ValueError("value argument must be a color string like 'red', a 3- or 4-float tuple for an RGB or RGBA value, or a dict")

    # Set any remaining unspecified defaults.
    tabColorArg.setdefault('red', 0.0)
    tabColorArg.setdefault('green', 0.0)
    tabColorArg.setdefault('blue', 0.0)
    tabColorArg.setdefault('alpha', 1.0)
    tabColorArg['red']   = float(tabColorArg['red'])
    tabColorArg['green'] = float(tabColorArg['green'])
    tabColorArg['blue']  = float(tabColorArg['blue'])
    tabColorArg['alpha'] = float(tabColorArg['alpha'])
    return tabColorArg


def convertToColumnRowInts(arg):
    if not isinstance(arg, str):
        raise TypeError("argument must be a grid cell str, like 'A1', not of type %s" % (type(arg).__name__))
    if not arg.isalnum() or not arg[0].isalpha() or not arg[-1].isdecimal():
        raise ValueError("argument must be a grid cell str, like 'A1', not %r" % (arg))

    for i in range(1, len(arg)):
        if arg[i].isdecimal():
            column = getColumnNumber(arg[:i])
            row = int(arg[i:])
            return (column, row)

    assert False # pragma: no cover We know this will always return before this point because arg[-1].isdecimal().


def createSpreadsheet(title=''):
    if not IS_INITIALIZED: init() # Initialize this module if not done so already.
    request = SERVICE.spreadsheets().create(body={
        'properties': {'title': title}
        })
    response = request.execute()

    return Spreadsheet(response['spreadsheetId'])


def getIdFromUrl(url):
    # https://docs.google.com/spreadsheets/d/16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c/edit#gid=0
    if url.startswith('https://docs.google.com/spreadsheets/d/'):
        spreadsheetId = url[39:url.find('/', 39)]
    else:
        spreadsheetId = url

    if re.match('^([a-zA-Z0-9]|_|-)+$', spreadsheetId) is None:
        raise ValueError('url argument must be an alphanumeric id or a full URL')
    return spreadsheetId


def getColumnLetterOf(columnNumber):
    """getColumnLetterOf(1) => 'A', getColumnLetterOf(27) => 'AA'"""
    if not isinstance(columnNumber, int):
        raise TypeError('columnNumber must be an int, not a %r' % (type(columnNumber).__name__))
    if columnNumber < 1:
        raise ValueError('columnNumber must be an int value of at least 1')

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
    if not isinstance(columnLetter, str):
        raise TypeError('columnLetter must be a str, not a %r' % (type(columnLetter).__name__))
    if not columnLetter.isalpha():
        raise ValueError('columnLetter must be composed of only letters')

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


def init(credentialsFile='credentials.json', tokenFile='token.pickle'):
    global SERVICE, IS_INITIALIZED

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
    IS_INITIALIZED = True




init()
s = Spreadsheet('https://docs.google.com/spreadsheets/d/1GfFDkD7LfwlVSLQMVQILaz2BPARG7Ott5Ui-frh0m2Y/edit#gid=0')
