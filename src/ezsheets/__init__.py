# EZSheets
# By Al Sweigart al@inventwithpython.com

# IMPORTANT NOTE: This module has not been stress-tested for performance
# and should not be considered "thread-safe" if multiple users are

# TODO - figure out drive quotas
# TODO - batch mode?

import collections
import json
import os.path
import pickle
import re
import time
import webbrowser
import http.client
from urllib.parse import urlparse

from apiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from ezsheets.colorvalues import COLORS

__version__ = "2024.8.9"

# SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES_SHEETS = ["https://www.googleapis.com/auth/spreadsheets"]
SHEETS_SERVICE = None
IS_INITIALIZED = False

# SCOPES_DRIVE = ['https://www.googleapis.com/auth/drive.readonly']
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive"]
DRIVE_SERVICE = None
IS_INITIALIZED = False


DEFAULT_NEW_ROW_COUNT = 1000  # This is the Google Sheets default for a new Sheet.
DEFAULT_NEW_COLUMN_COUNT = 26  # This is the Google Sheets default for a new Sheet.
DEFAULT_FROZEN_ROW_COUNT = 0
DEFAULT_FROZEN_COLUMN_COUNT = 0
DEFAULT_HIDE_GRID_LINES = False
DEFAULT_ROW_GROUP_CONTROL_AFTER = False
DEFAULT_COLUMN_GROUP_CONTROL_AFTER = False

# Quota throttling:
_READ_REQUESTS = collections.deque()
_WRITE_REQUESTS = collections.deque()
READ_QUOTA = 90  # 50 reads per 100 seconds
WRITE_QUOTA = 90  # 50 writes per 100 seconds
IGNORE_QUOTA = False
""" TODO - create a context manager to wrap calls, so that we can do both
preventative throttling and automated retries if it somehow raises an exception.
Also, use a sqlite database so that multiple scripts use the same queue.

Features to add:
- delete spreadsheets
- download as csv/excel/whatever
"""


# Sample spreadsheet id: 16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c


def _logWriteRequest():
    """
    Logs a write request to the `_WRITE_REQUESTS` deque. This function should be
    called whenever a Google Sheets write request is made. It will also throttle
    requests based on the quota in the global WRITE_QUOTA constant.

    By default, WRITE_QUOTA is set to 100 so that only 100 read requests can be
    made in the last 101 seconds.
    """
    _WRITE_REQUESTS.append(time.time())
    while _WRITE_REQUESTS[0] < time.time() - 101:  # 101 seconds rather than 100 in case of general inaccuracy
        _WRITE_REQUESTS.popleft()  # Get rid of all entries older than 100 seconds.

    if IGNORE_QUOTA:
        return  # Don't throttle.

    # Throttle if necessary:
    while len(_WRITE_REQUESTS) > (
        WRITE_QUOTA - 1
    ):  # pragma: no cover Note that the actual quota is one less than WRITE_QUOTA
        time.sleep(1)
        while _WRITE_REQUESTS[0] < time.time() - 101:
            _WRITE_REQUESTS.popleft()  # Get rid of all entries older than 100 seconds.


def _logReadRequest():
    """
    Logs a read request to the `_READ_REQUESTS` deque. This function should be
    called whenever a Google Sheets read request is made. It will also throttle
    requests based on the quota in the global READ_QUOTA constant.

    By default, READ_QUOTA is set to 50 so that only 50 read requests can be
    made in the last 101 seconds.
    """
    _READ_REQUESTS.append(time.time())
    while _READ_REQUESTS[0] < time.time() - 101:  # 101 seconds rather than 100 in case of general inaccuracy
        _READ_REQUESTS.popleft()  # Get rid of all entries older than 100 seconds

    if IGNORE_QUOTA:
        return  # Don't throttle.

    while len(_READ_REQUESTS) > (
        READ_QUOTA - 1
    ):  # pragma: no cover Note that the actual quota is one less than READ_QUOTA
        time.sleep(1)
        while _READ_REQUESTS[0] < time.time() - 101:
            _READ_REQUESTS.popleft()  # Get rid of all entries older than 100 seconds


def _makeRequest(requestType, **kwargs):
    pauseLength = 10
    while True:
        # TODO - do some of these requests count as a read AND write?
        if requestType == "get":
            request = SHEETS_SERVICE.spreadsheets().get(**kwargs)
            _logReadRequest()
        elif requestType == "batchUpdate":
            request = SHEETS_SERVICE.spreadsheets().batchUpdate(**kwargs)
            _logWriteRequest()
        elif requestType == "values.get":
            request = SHEETS_SERVICE.spreadsheets().values().get(**kwargs)
            _logReadRequest()
        elif requestType == "values.update":
            request = SHEETS_SERVICE.spreadsheets().values().update(**kwargs)
            _logWriteRequest()
        elif requestType == "sheets.copyTo":
            request = SHEETS_SERVICE.spreadsheets().sheets().copyTo(**kwargs)
            _logWriteRequest()
        elif requestType == "create":
            request = SHEETS_SERVICE.spreadsheets().create(**kwargs)
            _logWriteRequest()
        elif requestType == "drive.export":
            request = DRIVE_SERVICE.files().export(**kwargs)
            _logReadRequest()
        elif requestType == "drive.delete":
            request = DRIVE_SERVICE.files().delete(**kwargs)
            _logWriteRequest()
        elif requestType == "drive.update":
            request = DRIVE_SERVICE.files().update(**kwargs)
            _logWriteRequest()
        elif requestType == "drive.list":
            request = DRIVE_SERVICE.files().list(**kwargs)
            _logReadRequest()
        elif requestType == "drive.create":
            request = DRIVE_SERVICE.files().create(**kwargs)
            _logWriteRequest()
        else:
            assert False, "Invalid requestType: %r" % (requestType)

        try:
            return request.execute()
        except HttpError as e:
            errorContent = json.loads(str(e.content, encoding="utf-8"))
            if errorContent['error']['status'] != 'RESOURCE_EXHAUSTED':
                raise  # Some other, non-quota-related HttpError was raised, so we'll just re-raise it here.
            if pauseLength == 50:
                raise  # Throttling doesn't seem to work. Give up, and re-raise the error.
            time.sleep(pauseLength)
            pauseLength += 5


class EZSheetsException(Exception):
    """The base class for all EZSheets-specific problems. If the ``ezsheets`` module raises something that isn't this
    or a subclass of this exception, you can assume it is caused by a bug in EZSheets."""

    pass


class Spreadsheet:
    """
    This class represents a Spreadsheet on Google Sheets. Spreadsheets can
    contain one or more sheets, also called worksheets.
    """

    def __init__(self, spreadsheetId=None, credentialsFile='.'):
        """
        Initializer for Spreadsheet objects.

        :param spreadsheetId: The ID or URL of the spreadsheet on Google Sheets. E.g. `'https://docs.google.com/spreadsheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0'` or `'10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng'`
        """
        if not IS_INITIALIZED:
            init(credentialsFile=credentialsFile)  # Initialize this module if not done so already.

        if spreadsheetId is None:
            # Create a new spreadsheet.
            ss = createSpreadsheet()
            self._spreadsheetId = ss.id
            self.sheets = ()
            self.refresh()
            return

        try:
            # Figure out if this URL redirects to the Google Sheets URL.
            # NOTE: Restricted spreadsheets will redirect a docs.google.com url to their https://accounts.google.com/v3/signin/... URL, which
            # we don't want, so if it begins with docs.google.com just use it and don't check for redirects. (This doesn't apply to shared spreadsheets.)
            while spreadsheetId.lower().startswith('http') and not spreadsheetId.lower().startswith('https://docs.google.com'):

                redirects = [spreadsheetId]
                while True:
                    parsed_url = urlparse(spreadsheetId)
                    http_conn = http.client.HTTPConnection(parsed_url.netloc)
                    http_conn.request("GET", parsed_url.path)
                    response = http_conn.getresponse()
                    redirects.append(spreadsheetId)

                    spreadsheetId = response.getheader('Location')
                    if spreadsheetId == redirects[-1] or response.status not in (301, 302):
                        """There's some weird behavior where the google doc url keeps redirecting to itself forever,
                        hence why I added the spreadsheetId == redirects[-1] check. I'm not sure what causes this.
                        Requests doesn't have this problem, nor does `wget https://bit.ly/3D34nDh 2>&1 | grep Location:`
                        so I can't quite fix it. But this works well enough for now."""
                        break

            try:
                spreadsheetId = getIdFromUrl(spreadsheetId)
            except ValueError:
                pass  # No problem if it's not a valid ID or URL; it could be a title.

            # request = SHEETS_SERVICE.spreadsheets().get(spreadsheetId=spreadsheetId)
            # _logReadRequest(); request.execute()
            _makeRequest("get", **{"spreadsheetId": spreadsheetId})
        except HttpError:
            # URL/ID wasn't found, so check if this is the title of a spreadsheet returned by listSpreadsheets()
            sheetIDsWithTitle = []
            for listedId, listedTitle in listSpreadsheets().items():
                if listedTitle == spreadsheetId:
                    sheetIDsWithTitle.append(listedId)
            if len(sheetIDsWithTitle) == 1:
                spreadsheetId = sheetIDsWithTitle[0]
            elif len(sheetIDsWithTitle) == 0:
                raise EZSheetsException(
                    "No spreadsheet with id, url, or title of %r found for the Google account in this token file."
                    % (spreadsheetId)
                )
            elif len(sheetIDsWithTitle) > 1:
                raise EZSheetsException(
                    "Multiple spreadsheets with title of %r found. Specify the id or url instead." % (spreadsheetId)
                )

        self._spreadsheetId = spreadsheetId
        self.sheets = ()
        self.refresh()

    def refresh(self):
        """
        Updates this Spreadsheet object's Sheet objects with the current data
        of the spreadsheet and sheets on Google sheets.
        """
        # _logReadRequest(); response = SHEETS_SERVICE.spreadsheets().get(spreadsheetId=self._spreadsheetId).execute()
        response = _makeRequest("get", **{"spreadsheetId": self._spreadsheetId})

        self._title = response["properties"]["title"]

        # Delete local Sheets that are no longer in the online spreadsheet:
        remoteSheetIds = [sheetInfo["properties"]["sheetId"] for sheetInfo in response["sheets"]]
        replacementSheetsAttr = []  # We will replace self.sheets with this list.
        for sheet in self.sheets:
            if sheet._sheetId in remoteSheetIds:
                replacementSheetsAttr.append(sheet)
        self.sheets = replacementSheetsAttr

        # Update/create Sheet objects:
        replacementSheetsAttr = []  # We will replace self.sheets with this list.
        for i, sheetInfo in enumerate(response["sheets"]):
            sheetId = sheetInfo["properties"]["sheetId"]

            # Find the index of the sheet if it already exists in `self.sheets`
            try:
                existingSheetIndex = [sh.sheetId for sh in self.sheets].index(sheetId)
            except ValueError:
                existingSheetIndex = None

            if existingSheetIndex is not None:
                # If the sheet has been previously loaded, reuse that Sheet object:
                replacementSheetsAttr.append(self.sheets[existingSheetIndex])
                self.sheets[existingSheetIndex]._refreshPropertiesWithSheetPropertiesDict(sheetInfo["properties"])
                self.sheets[existingSheetIndex]._refreshData()
            else:
                # If the sheet hasn't been seen before, create a new Sheet object:
                replacementSheetsAttr.append(
                    Sheet(self, sheetId)
                )  # TODO - would be nice to reuse the info in `response` for this instead of letting the ctor make another request, but this isn't that important.

        self.sheets = tuple(replacementSheetsAttr)  # Make sheets attribute an immutable tuple.

    def __getitem__(self, key):
        """
        Retrieve the Sheet object at index `key`.

        :param key An integer index of the Sheet to retrieve.
        """
        if isinstance(key, str):
            # Assume key is a sheet title:
            try:
                i = self.sheetTitles.index(key)
                return self.sheets[i]
            except ValueError:
                pass  # Do nothing if the title isn't found.
        elif isinstance(key, int) and (-len(self.sheets) <= key < len(self.sheets)):
            # Assume key is an integer index:
            return self.sheets[key]
        elif isinstance(key, slice):
            # Assume key is a slice object:
            return self.sheets[key]

        raise KeyError(
            "key must be an int between %s and %s or a str matching a title: %r"
            % (-(len(self.sheets)), len(self.sheets) - 1, self.sheetTitles)
        )

    def __delitem__(self, key):
        """
        Delete the Sheet object at index `key`.

        :param key An integer index of the Sheet to delete.
        """
        if isinstance(key, (int, str)):
            # Key is an int index or a str title.
            self[key].delete()
        elif isinstance(key, slice):
            # TODO - there's got to be a better way to do this.
            start = key.start if key.start is not None else 0
            stop = key.stop if key.stop is not None else len(self.sheets)
            step = key.step if key.step is not None else 1

            if start < 0 or stop < 0:
                return  # When deleting list items with a slice, a negative start or stop results in a no-op. I'll mimic that behavior here.

            indexesToDelete = [
                i for i in range(start, stop, step) if i >= 0 and i < len(self.sheets)
            ]  # Don't include invalid or negative indexes.
            if len(indexesToDelete) == len(self.sheets):
                raise ValueError("Cannot delete all sheets; spreadsheets must have at least one sheet")

            if indexesToDelete[0] < indexesToDelete[-1]:
                indexesToDelete.reverse()  # We want this is descending order.

            for i in indexesToDelete:
                self.sheets[i].delete()

        else:
            raise TypeError("key must be an int index, str sheet title, or slice object, not %r" % (type(key).__name__))

    def __len__(self):
        """
        Return the number of Sheet objects in this Spreadsheet.
        """
        return len(self.sheets)

    def __iter__(self):
        """
        Return an iterable of the Sheet objects in this Spreadsheet.
        """
        return iter(self.sheets)

    @property
    def spreadsheetId(self):
        """
        The unique, read-only id for this Spreadsheet on Google Sheets. (This is the old, deprecated name. Use id instead.)
        """
        return self._spreadsheetId


    @property
    def id(self):
        """
        The unique, read-only id for this Spreadsheet on Google Sheets.
        """
        return self._spreadsheetId


    @property
    def url(self):
        """
        The URL for this Spreadsheet on Google Sheets.
        """
        return "https://docs.google.com/spreadsheets/d/" + self._spreadsheetId + "/"

    @property
    def sheetTitles(self):
        """
        A tuple of the Sheet objects' titles (as strings) in this Spreadsheet object.
        """
        return tuple([sheet.title for sheet in self.sheets])

    def __str__(self):
        """
        A human-readable string representation of this Spreadsheet object.
        """
        return '<%s title="%s", %d sheets>' % (type(self).__name__, self.title, len(self.sheets))

    def __repr__(self):
        """
        A string representation of code that can recreate this Spreadsheet object.
        """
        return "%s(spreadsheetId=%r)" % (type(self).__name__, self.spreadsheetId)
        # NOTE that the __str__ function will still use "spreadsheetId" instead of "id" to maintain backwards compatibility.

    @property
    def title(self):
        """
        The string title for this Spreadsheet on Google Sheets. Both Spreadsheets and Sheets have titles.
        """
        return self._title

    @title.setter
    def title(self, value):
        value = str(value)
        # request = SHEETS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheetId,
        # body={
        #    'requests': [{'updateSpreadsheetProperties': {'properties': {'title': value},
        #                                                  'fields': 'title'}}]})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "batchUpdate",
            **{
                "spreadsheetId": self._spreadsheetId,
                "body": {
                    "requests": [{"updateSpreadsheetProperties": {"properties": {"title": value}, "fields": "title"}}]
                },
            }
        )
        self._title = value

    def createSheet(self, title="", index=None, columnCount=DEFAULT_NEW_COLUMN_COUNT, rowCount=DEFAULT_NEW_ROW_COUNT):
        """
        Create a new Sheet object in this Spreadsheet.
        """
        if index is None:
            # Set the index to make this new sheet be the last sheet:
            index = len(self.sheets)

        # request = SHEETS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheetId,
        # body={
        #    'requests': [{'addSheet': {'properties': {'title': title, 'index': index}}}]})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "batchUpdate",
            **{
                "spreadsheetId": self._spreadsheetId,
                "body": {"requests": [{"addSheet": {"properties": {"title": title, "index": index}}}]},
            }
        )

        self.refresh()
        self.sheets[index].resize(columnCount, rowCount)
        return self.sheets[index]

    def _download(self, filename=None, _fileType="spreadsheet"):
        fileTypes = {
            "csv": "text/csv",
            "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "ods": "application/x-vnd.oasis.opendocument.spreadsheet",
            "pdf": "application/pdf",
            "zip": "application/zip",  # a zip file of html files
            "tsv": "text/tab-separated-values",
        }

        if filename is None:
            filename = _makeFilenameSafe(self._title) + "." + _fileType

        request = DRIVE_SERVICE.files().export(fileId=self._spreadsheetId, mimeType=fileTypes[_fileType])
        fh = open(filename, "wb")
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()

        return filename

    def downloadAsCSV(self, filename=None):
        return self._download(filename, "csv")

    def downloadAsExcel(self, filename=None):
        return self._download(filename, "xlsx")

    def downloadAsODS(self, filename=None):
        return self._download(filename, "ods")

    def downloadAsPDF(self, filename=None):
        return self._download(filename, "pdf")

    def downloadAsHTML(self, filename=None):
        return self._download(filename, "zip")

    def downloadAsTSV(self, filename=None):
        return self._download(filename, "tsv")

    def delete(self, permanent=False):
        if permanent:
            # Delete spreadsheet without moving it to Trashed folder:
            # DRIVE_SERVICE.files().delete(fileId=self._spreadsheetId).execute()
            _makeRequest("drive.delete", **{"fileId": self._spreadsheetId})
        else:
            # Delete spreadsheet by moving it to Trashed folder:
            # DRIVE_SERVICE.files().update(fileId=self._spreadsheetId,
            #                             body={'trashed': True}).execute()
            _makeRequest("drive.update", **{"fileId": self._spreadsheetId, "body": {"trashed": True}})

    def open(self):
        webbrowser.open(self.url)


    def __eq__(self, other):
        """
        A Spreadsheet object is only considered equal to Spreadsheet objects with the same ID.
        """
        if not isinstance(other, Spreadsheet):
            return False
        return self.spreadsheetId == other.spreadsheetId

    def Sheet(self, title="", index=None, columnCount=DEFAULT_NEW_COLUMN_COUNT, rowCount=DEFAULT_NEW_ROW_COUNT):
        # Wrapper for createSheet(). Now that we can create Spreadsheet
        # objects with by calling the Spreadsheet() class's init, we
        # should have similar code for create sheets.
        # The createSpreadsheet() and createSheet() functions remain,
        # but they were always awkward from an API design point of view.
        self.createSheet(title, index, columnCount, rowCount)



def _makeFilenameSafe(filename):
    for replaceChar in ' \\/:*?"<>|':
        filename = filename.replace(replaceChar, "_")
    return filename


class Sheet:
    """
    This class represents an individual worksheet inside a spreadsheet. Sheets
    are composed of columns and rows of cells, which contain a single string value.
    """

    def __init__(self, spreadsheet, sheetId):
        """
        TODO
        """
        if not IS_INITIALIZED:
            init()  # Initialize this module if not done so already. # This line might not be needed? Sheet objects can only exist when you've already made a Spreadsheet object.

        # Set the properties of this sheet
        self._spreadsheet = spreadsheet
        self._sheetId = sheetId
        self._cells = (
            {}
        )  # To ease development, internally the local copy of the sheet data is stored in a dict with 1-based (column, row) keys.
        self.refresh()

    # Set up the read-only attributes.
    @property
    def spreadsheet(self):
        """
        The Spreadsheet object that contains this Sheet object.
        """
        return self._spreadsheet

    @property
    def title(self):
        """
        The title of this Sheet on Google Sheets. Both Spreadsheets and Sheets have titles.
        """
        return self._title

    @title.setter
    def title(self, value):
        value = str(value)
        # request = SHEETS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet.spreadsheetId,
        # body={
        #    'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
        #                                                           'title': value},
        #                                            'fields': 'title'}}]})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "batchUpdate",
            **{
                "spreadsheetId": self._spreadsheet.spreadsheetId,
                "body": {
                    "requests": [
                        {
                            "updateSheetProperties": {
                                "properties": {"sheetId": self._sheetId, "title": value},
                                "fields": "title",
                            }
                        }
                    ]
                },
            }
        )
        self._title = value

    @property
    def tabColor(self):
        """
        The color of the Sheet's tab as displayed in the browser.
        """
        return self._tabColor

    @tabColor.setter
    def tabColor(self, value):
        tabColorArg = _getTabColorArg(value)

        # request = SHEETS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet.spreadsheetId,
        # body={
        #    'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
        #                                                           'tabColor': tabColorArg},
        #                                            'fields': 'tabColor'}}]})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "batchUpdate",
            **{
                "spreadsheetId": self._spreadsheet.spreadsheetId,
                "body": {
                    "requests": [
                        {
                            "updateSheetProperties": {
                                "properties": {"sheetId": self._sheetId, "tabColor": tabColorArg},
                                "fields": "tabColor",
                            }
                        }
                    ]
                },
            }
        )
        self._tabColor = tabColorArg

    @property
    def index(self):
        """
        The integer index of this Sheet in it's Spreadsheet's tuple of Sheet objects.
        """
        return self._index

    @index.setter
    def index(self, value):
        if value == self._index:
            return  # No change needed.

        if not isinstance(value, int):
            raise TypeError("indices must be integers, not %s" % (type(value).__name__))

        if value < 0:  # Handle negative indexes the way Python lists do.
            if value < -len(self.spreadsheet.sheets):
                raise IndexError("%r is out of range (-1 to %d)" % (value, -len(self.spreadsheet.sheets)))
            value = (
                len(self.spreadsheet.sheets) + value
            )  # convert this negative index into its corresponding positive index
        if value >= len(self.spreadsheet.sheets):
            raise IndexError("%r is out of range (0 to %d)" % (value, len(self.spreadsheet.sheets) - 1))

        # Update the index:
        if value > self._index:
            # Google Sheets uses "before the move" indexes, which is confusing and I don't want to do it here.
            value += 1


        # request = SHEETS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
        # body={
        #    'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
        #                                                           'index': value},
        #                                            'fields': 'index'}}]})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "batchUpdate",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "body": {
                    "requests": [
                        {
                            "updateSheetProperties": {
                                "properties": {"sheetId": self._sheetId, "index": value},
                                "fields": "index",
                            }
                        }
                    ]
                },
            }
        )

        self._spreadsheet.refresh()  # Update the spreadsheet's tuple of Sheet objects to reflect the new order.
        # self._index = self._spreadsheet.sheets.index(self) # Update the local Sheet object's index.

    def __eq__(self, other):
        """
        A Sheet object is only considered equal to itself.
        """
        if not isinstance(other, Sheet):
            return False
        return self._sheetId == other._sheetId

    @property
    def sheetId(self):
        """
        The unique, read-only ID string of this Sheet object in its Spreadsheet. (This is the old, deprecated name. Use id instead.)
        """
        return self._sheetId

    @property
    def id(self):
        """
        The unique, read-only ID string of this Sheet object in its Spreadsheet. (This is the old, deprecated name. Use id instead.)
        """
        # NOTE that the __str__ function will still use "sheetId" instead of "id" to maintain backwards compatibility.
        return self._sheetId


    @property
    def rowCount(self):
        """
        The number of rows in this Sheet object.
        """
        return self._rowCount

    @rowCount.setter
    def rowCount(self, value):
        # Validate arguments:
        if not isinstance(value, int):
            raise TypeError("value arg must be an int, not %s" % (type(value).__name__))
        if value < 1:
            raise TypeError("value arg must be a positive nonzero int, not %r" % (value))
        if value <= self._frozenRowCount:
            raise ValueError(
                "You cannot have all rows on the sheet frozen (sheet %r has %s frozen rows)"
                % (self.title, self._frozenRowCount)
            )

        self.refresh()  # Retrieve up-to-date grid properties from Google Sheets.
        self._rowCount = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.

    @property
    def columnCount(self):
        """
        The number of columns in this Sheet object.
        """
        return self._columnCount

    @columnCount.setter
    def columnCount(self, value):
        # Validate arguments:
        if not isinstance(value, int):
            raise TypeError("value arg must be an int, not %s" % (type(value).__name__))
        if value < 1:
            raise TypeError("value arg must be a positive nonzero int, not %r" % (value))
        if value <= self._frozenColumnCount:
            raise ValueError(
                "You cannot have all columns on the sheet frozen (sheet %r has %s frozen columns)"
                % (self.title, self._frozenColumnCount)
            )

        self.refresh()  # Retrieve up-to-date grid properties from Google Sheets.
        self._columnCount = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.

    @property
    def frozenRowCount(self):
        """
        The integer number of frozen rows in this Sheet object. Frozen rows remain visible
        in the browser even as the user scrolls down the Sheet. There must be at
        least one non-frozen row in the Sheet.
        """
        return self._frozenRowCount

    @frozenRowCount.setter
    def frozenRowCount(self, value):
        # Validate arguments:
        if not isinstance(value, int):
            raise TypeError("value arg must be an int, not %s" % (type(value).__name__))
        if value < 1:
            raise TypeError("value arg must be a positive nonzero int, not %r" % (value))
        if value >= self._rowCount:
            raise ValueError(
                "You cannot freeze all rows on the sheet (sheet %r has %s rows)" % (self.title, self._rowCount)
            )

        self.refresh()  # Retrieve up-to-date grid properties from Google Sheets.
        self._frozenRowCount = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.

    @property
    def frozenColumnCount(self):
        """
        The integer number of frozen columns in this Sheet object. Frozen columns remain visible
        in the browser even as the user scrolls the Sheet. There must be at
        least one non-frozen column in the Sheet.
        """
        return self._frozenColumnCount

    @frozenColumnCount.setter
    def frozenColumnCount(self, value):
        # Validate arguments:
        if not isinstance(value, int):
            raise TypeError("value arg must be an int, not %s" % (type(value).__name__))
        if value < 1:
            raise TypeError("value arg must be a positive nonzero int, not %r" % (value))
        if value >= self._columnCount:
            raise ValueError(
                "You cannot freeze all columns on the sheet (sheet %r has %s columns)" % (self.title, self._columnCount)
            )

        self.refresh()  # Retrieve up-to-date grid properties from Google Sheets.
        self._frozenColumnCount = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.

    @property
    def hideGridlines(self):
        """
        The Boolean setting of whether the gridlines in the Sheet are visible or not.
        """
        return self._hideGridlines

    @hideGridlines.setter
    def hideGridlines(self, value):
        value = bool(value)

        self.refresh()  # Retrieve up-to-date grid properties from Google Sheets.
        self._hideGridlines = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.

    @property
    def rowGroupControlAfter(self):
        """
        TODO
        """
        return self._rowGroupControlAfter

    @rowGroupControlAfter.setter
    def rowGroupControlAfter(self, value):
        value = bool(value)

        self.refresh()  # Retrieve up-to-date grid properties from Google Sheets.
        self._rowGroupControlAfter = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.

    @property
    def columnGroupControlAfter(self):
        """
        TODO
        """
        return self._columnGroupControlAfter

    @columnGroupControlAfter.setter
    def columnGroupControlAfter(self, value):
        value = bool(value)

        self.refresh()  # Retrieve up-to-date grid properties from Google Sheets.
        self._columnGroupControlAfter = value  # Change local grid property.
        self._updateGridProperties()  # Upload grid properties to Google Sheets.

    def __str__(self):
        """
        TODO
        """
        return "<%s title=%r, sheetId=%r, rowCount=%r, columnCount=%r>" % (
            type(self).__name__,
            self._title,
            self._sheetId,
            self._rowCount,
            self._columnCount,
        )
        # NOTE that the __str__ function will still use "sheetId" instead of "id" to maintain backwards compatibility.

    def __repr__(self):
        """
        TODO
        """
        return "<%s sheetId=%r, title=%r, rowCount=%r, columnCount=%r>" % (
            type(self).__name__,
            self._sheetId,
            self._title,
            self._rowCount,
            self._columnCount,
        )
        # NOTE that the __str__ function will still use "sheetId" instead of "id" to maintain backwards compatibility.

    def get(self, *args):
        """
        Retrieve the value in a cell. The arguments to `get()` can either be two
        integers (column and row) or a single string such as 'A1'.
        """
        # TODO!!!! Add a switch or a mode or something so that all the ezsheets functions call refresh() before running.
        if len(args) == 2:  # args are column, row like (2, 5)
            column, row = args
        elif len(args) == 1:  # args is a string of a grid cell like ('B5',)
            column, row = convertToColumnRowInts(args[0])
        else:
            raise TypeError("get() takes one or two arguments, like ('A1',) or (2, 5)")

        if not isinstance(column, int):
            raise TypeError("column indices must be integers, not %s" % (type(column).__name__))
        if not isinstance(row, int):
            raise TypeError("row indices must be integers, not %s" % (type(row).__name__))
        if column < 1 or row < 1:
            raise IndexError(
                "Column %s, row %s does not exist. Google Sheets' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index. Negative indices are not supported by ezsheets."
                % (column, row)
            )

        return self._cells.get((column, row), "")

    """
    def getAllRows(self):
        rows = []
        for rowNum in range(1, self._rowCount + 1):
            row = []
            for colNum in range(1, self._columnCount + 1):
                row.append(self._cells.get((colNum, rowNum), ''))
            rows.append(row)
        return rows


    def getAllColumns(self):
        cols = []
        for colNum in range(1, self._columnCount + 1):
            col = []
            for rowNum in range(1, self._rowCount + 1):
                col.append(self._cells.get((colNum, rowNum), ''))
            cols.append(col)
        return cols
    """

    def getRow(self, rowNum):
        # NOTE: getRow() and getCol() do not support negative indexes.
        if not isinstance(rowNum, int):
            raise TypeError("rowNum indices must be integers, not %s" % (type(rowNum).__name__))
        if rowNum < 1:
            raise IndexError(
                "Row %s does not exist. Google Sheets' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index."
                % (rowNum)
            )

        row = []
        for colNum in range(1, self._columnCount + 1):
            row.append(self._cells.get((colNum, rowNum), ""))
        return row

    def getRows(self, startRow=1, stopRow=None):
        # Validate arguments:
        if stopRow is None:
            stopRow = self._rowCount + 1
        if not isinstance(startRow, int):
            raise TypeError("startRow arg must be an int, not %s" % (type(startRow).__name__))
        if startRow < 1:
            raise ValueError("startRow arg must be at least 1, not %s" % (startRow))
        if not isinstance(stopRow, int):
            raise TypeError("stopRow arg must be an int, not %s" % (type(stopRow).__name__))
        if stopRow < 1:
            raise ValueError("stopRow arg must be at least 1, not %s" % (stopRow))

        # Get rows by calling getRow():
        return [self.getRow(rowNum) for rowNum in range(startRow, stopRow)]

    def __contains__(self, item):
        """Returns `True` if the `str` representation of `item` is equal to or
        within the `str` representation of a cell in this sheet."""
        for cell in self._cells:
            if str(item) in str(cell):
                return True
        return False

    def find(self, needle):
        pass

    def getColumn(self, colNum):
        # NOTE: getRow() and getCol() do not support negative indexes.
        if isinstance(colNum, str):
            colNum = getColumnNumberOf(colNum)

        if not isinstance(colNum, int):
            raise TypeError("colNum indices must be integers, not %s" % (type(colNum).__name__))
        if colNum < 1:
            raise IndexError(
                "Column %s does not exist. Google Sheets' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index."
                % (colNum)
            )

        column = []
        for rowNum in range(1, self._rowCount + 1):
            column.append(self._cells.get((colNum, rowNum), ""))
        return column

    def getColumns(self, startColumn=1, stopColumn=None):
        # Validate arguments:
        if stopColumn is None:
            stopColumn = self._columnCount + 1
        if not isinstance(startColumn, int):
            raise TypeError("startColumn arg must be an int, not %s" % (type(startColumn).__name__))
        if startColumn < 1:
            raise ValueError("startColumn arg must be at least 1, not %s" % (startColumn))
        if not isinstance(stopColumn, int):
            raise TypeError("stopColumn arg must be an int, not %s" % (type(stopColumn).__name__))
        if stopColumn < 1:
            raise ValueError("stopColumn arg must be at least 1, not %s" % (stopColumn))

        # Get columns by calling getColumn():
        return [self.getColumn(colNum) for colNum in range(startColumn, stopColumn)]

    def refresh(self):
        self._refreshProperties()
        self._refreshData()

    def _refreshProperties(self):
        # Get all the sheet properties:
        # _logReadRequest(); response = SHEETS_SERVICE.spreadsheets().get(spreadsheetId=self._spreadsheet._spreadsheetId).execute()
        response = _makeRequest("get", **{"spreadsheetId": self._spreadsheet._spreadsheetId})

        for sheetDict in response["sheets"]:
            if (
                sheetDict["properties"]["sheetId"] == self._sheetId
            ):  # Find this sheet in the returned spreadsheet json data.
                self._refreshPropertiesWithSheetPropertiesDict(sheetDict["properties"])

    def _refreshPropertiesWithSheetPropertiesDict(self, sheetPropsDict):
        self._title = sheetPropsDict["title"]
        self._index = sheetPropsDict["index"]
        self._tabColor = _getTabColorArg(sheetPropsDict.get("tabColor"))  # Set to None if there is no tabColor.

        # These attrs we don't have properties for yet, I'm not sure if we'll keep them:
        self._sheetType = sheetPropsDict.get("sheetType")
        self._hidden = sheetPropsDict.get("hidden")
        self._rightToLeft = sheetPropsDict.get("rightToLeft")

        gridProps = sheetPropsDict["gridProperties"]
        self._rowCount = gridProps.get("rowCount", DEFAULT_NEW_ROW_COUNT)
        self._columnCount = gridProps.get("columnCount", DEFAULT_NEW_COLUMN_COUNT)
        self._frozenRowCount = gridProps.get("frozenRowCount", DEFAULT_FROZEN_ROW_COUNT)
        self._frozenColumnCount = gridProps.get("frozenColumnCount", DEFAULT_FROZEN_COLUMN_COUNT)
        self._hideGridlines = gridProps.get("hideGridlines", DEFAULT_HIDE_GRID_LINES)
        self._rowGroupControlAfter = gridProps.get("rowGroupControlAfter", DEFAULT_ROW_GROUP_CONTROL_AFTER)
        self._columnGroupControlAfter = gridProps.get("columnGroupControlAfter", DEFAULT_COLUMN_GROUP_CONTROL_AFTER)

    def _refreshData(self):
        # Get all the sheet data:
        # _logReadRequest(); response = SHEETS_SERVICE.spreadsheets().values().get(
        #    spreadsheetId=self._spreadsheet._spreadsheetId,
        #    range='%s!A1:%s%s' % (self._title, getColumnLetterOf(self._columnCount), self._rowCount)).execute()
        response = _makeRequest(
            "values.get",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "range": "%s!A1:%s%s" % (self._title, getColumnLetterOf(self._columnCount), self._rowCount),
            }
        )

        sheetData = response.get("values", [[]])
        self._cells = {}
        if response["majorDimension"] == "ROWS":
            for rowNumBase0, row in enumerate(sheetData):
                for colNumBase0, sheetDatum in enumerate(row):
                    self._cells[(colNumBase0 + 1, rowNumBase0 + 1)] = sheetDatum
        elif response["majorDimension"] == "COLUMNS":
            for colNumBase0, column in enumerate(sheetData):
                for rowNumBase0, sheetDatum in enumerate(column):
                    self._cells[(colNumBase0 + 1, rowNumBase0 + 1)] = sheetDatum

    def _updateGridProperties(self):
        gridProperties = {
            "rowCount": self._rowCount,
            "columnCount": self._columnCount,
            "frozenRowCount": self._frozenRowCount,
            "frozenColumnCount": self._frozenColumnCount,
            "hideGridlines": self._hideGridlines,
            "rowGroupControlAfter": self._rowGroupControlAfter,
            "columnGroupControlAfter": self._columnGroupControlAfter,
        }
        # request = SHEETS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
        #    body={
        #    'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
        #                                                           'gridProperties': gridProperties},
        #                                            'fields': 'gridProperties'}}]})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "batchUpdate",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "body": {
                    "requests": [
                        {
                            "updateSheetProperties": {
                                "properties": {"sheetId": self._sheetId, "gridProperties": gridProperties},
                                "fields": "gridProperties",
                            }
                        }
                    ]
                },
            }
        )

    def _enlargeIfNeeded(self, requestedColumn=None, requestedRow=None):
        # Increase rowCount or columnCount if needed.
        if requestedColumn is None:
            requestedColumn = self._columnCount
        if requestedRow is None:
            requestedRow = self._rowCount

        # Enlarge the sheet:
        self.resize(max(requestedColumn, self._columnCount), max(requestedRow, self._rowCount))

    def update(self, *args):
        if len(args) == 3:  # args are column, row like (2, 5)
            column, row, value = args
        elif len(args) == 2:  # args is a string of a grid cell like ('B5',)
            if isinstance(args[0], int) and isinstance(args[1], int):
                raise TypeError("You most likely have forgotten to supply a value to update the this cell with.")
            column, row = convertToColumnRowInts(args[0])
            value = args[1]
        else:
            raise TypeError("get() takes one or two arguments, like ('A1',) or (2, 5)")

        if not isinstance(column, int):
            raise TypeError("column indices must be integers, not %s" % (type(column).__name__))
        if not isinstance(row, int):
            raise TypeError("row indices must be integers, not %s" % (type(row).__name__))
        if column < 1 or row < 1:
            raise IndexError(
                "Column %s, row %s does not exist. Google Sheets' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index. Negative indices are not supported by ezsheets."
                % (column, row)
            )

        if value is None:
            value == ""  # Pass None or '' for value to delete the cell's content.

        self._enlargeIfNeeded(column, row)

        cellLocation = getColumnLetterOf(column) + str(row)
        # request = SHEETS_SERVICE.spreadsheets().values().update(
        #    spreadsheetId=self._spreadsheet._spreadsheetId,
        #    range='%s!%s:%s' % (self._title, cellLocation, cellLocation),
        #    valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
        #    body={
        #        'majorDimension': 'ROWS',
        #        'values': [[value]],
        #        #'range': '%s!%s:%s' % (self._title, cellLocation, cellLocation),
        #        }
        #    )
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "values.update",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "range": "%s!%s:%s" % (self._title, cellLocation, cellLocation),
                "valueInputOption": "USER_ENTERED",  # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
                "body": {"majorDimension": "ROWS", "values": [[value]]},
            }
        )

        if value == "":
            del self._cells[(column, row)]
        else:
            # Google Sheets seem to only store strings (TODO: verify this), but we can't
            # do a simple str() call here because True and False are stored as 'TRUE' and 'FALSE'
            # I don't want to have to do a refresh on each setting, so for the _cells cache
            # I'll just hard code some known rules and we can hunt down the edge cases later.
            if isinstance(value, bool):
                value = str(value).upper()
            else:
                value = str(value)

            self._cells[(column, row)] = value

    def updateRow(self, row, values):
        if not isinstance(row, int):
            raise TypeError("row indices must be integers, not %s" % (type(row).__name__))
        if row < 1:
            raise IndexError(
                "Row %s does not exist. Google Sheets' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index."
                % (row)
            )
        if not isinstance(values, (list, tuple)):
            raise TypeError("values must be a list or tuple, not %s" % (type(values).__name__))

        if isinstance(values, tuple):
            values = list(values)
        if len(values) < self._columnCount:
            values.extend([""] * (self._columnCount - len(values)))

        self._enlargeIfNeeded(None, row)

        # request = SHEETS_SERVICE.spreadsheets().values().update(
        #    spreadsheetId=self._spreadsheet._spreadsheetId,
        #    range='%s!A%s:%s%s' % (self._title, row, getColumnLetterOf(len(values)), row),
        #    valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
        #    body={
        #        'majorDimension': 'ROWS',
        #        'values': [values],
        #        #'range': '%s!A%s:%s%s' % (self._title, row, getColumnLetterOf(len(values)), row),
        #        }
        #    )
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "values.update",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "range": "%s!A%s:%s%s" % (self._title, row, getColumnLetterOf(len(values)), row),
                "valueInputOption": "USER_ENTERED",  # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
                "body": {"majorDimension": "ROWS", "values": [values]},
            }
        )

        # Update the local data in `_cells`:
        for colNumBase1 in range(1, self._columnCount + 1):
            self._cells[(colNumBase1, row)] = values[colNumBase1 - 1]

    def updateColumn(self, column, values):
        if not isinstance(column, (int, str)):
            raise TypeError("column indices must be integers, not %s" % (type(column).__name__))
        if isinstance(column, int) and column < 1:
            raise IndexError(
                "Column %s does not exist. Google Sheets' columns and rows are 1-based, not 0-based. Use index 1 instead of index 0 for row and column index."
                % (column)
            )
        if not isinstance(values, (list, tuple)):
            raise TypeError("values must be a list or tuple, not %s" % (type(values).__name__))
        if isinstance(column, str) and not column.isalpha():
            raise ValueError("Column %s does not exist. Columns must be a 1-based int or a letters-only str.")

        if isinstance(values, tuple):
            values = list(values)
        if isinstance(column, str):
            column = getColumnNumberOf(column)

        if len(values) < self._rowCount:
            values.extend([""] * (self._rowCount - len(values)))

        self._enlargeIfNeeded(column, None)

        # request = SHEETS_SERVICE.spreadsheets().values().update(
        #    spreadsheetId=self._spreadsheet._spreadsheetId,
        #    range='%s!%s1:%s%s' % (self._title, getColumnLetterOf(column), getColumnLetterOf(column), len(values)),
        #    valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
        #    body={
        #        'majorDimension': 'COLUMNS',
        #        'values': [values],
        #        #'range': '%s!%s1:%s%s' % (self._title, getColumnLetterOf(column), getColumnLetterOf(column), len(values)),
        #        }
        #    )
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "values.update",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "range": "%s!%s1:%s%s"
                % (self._title, getColumnLetterOf(column), getColumnLetterOf(column), len(values)),
                "valueInputOption": "USER_ENTERED",  # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
                "body": {"majorDimension": "COLUMNS", "values": [values]},
            }
        )

        # Update the local data in `_cells`:
        for rowNumBase1 in range(1, self._rowCount + 1):
            self._cells[(column, rowNumBase1)] = values[rowNumBase1 - 1]

    def updateRows(self, rows, startRow=1):
        # Argument validation:
        # Ensure that `rows` is a list of lists:
        if not isinstance(rows, (list, tuple)):
            raise TypeError("rows arg must be a list/tuple of lists/tuples, not %s" % (type(rows).__name__))
        for row in rows:
            if not isinstance(row, (list, tuple)):
                raise TypeError("rows arg contains a non-list/tuple")

        if not isinstance(startRow, int):
            raise TypeError("startRow arg must be an int, not %s" % (type(startRow).__name__))
        if startRow < 1:
            raise ValueError("startRow arg is 1-based, and must be 1 or greater, not %r" % (startRow))

        if startRow > self._rowCount:
            return  # No rows to update, so return.

        # Find out the max length of a row in `rows`. This will be the new columnCount for the sheet:
        maxColumnCount = self._columnCount
        for row in rows:
            maxColumnCount = max(maxColumnCount, len(row))

        # Lengthen rows to the length of self._rowCount, and each row to the length of self._columnCount:
        for row in rows:
            row.extend([""] * (maxColumnCount - len(row)))  # pad each row
        while len(rows) < (
            self._rowCount - startRow + 1
        ):  # TODO - this could probably be made more performant if we use extend().
            rows.append([""] * self._columnCount)  # pad extra rows

        self._enlargeIfNeeded(None, len(rows) + startRow - 1)

        # Send the API request that updates the Google sheet.
        # rangeCells = '%s!A%s:%s%s' % (self._title, startRow, getColumnLetterOf(maxColumnCount), stopRow - 1)
        # request = SHEETS_SERVICE.spreadsheets().values().update(
        #    spreadsheetId=self._spreadsheet._spreadsheetId,
        #    range='%s!A%s:%s%s' % (self._title, startRow, getColumnLetterOf(maxColumnCount), startRow + len(rows) - 1),
        #    valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
        #    body={
        #        'majorDimension': 'ROWS',
        #        'values': rows,
        #        #'range': rangeCells,
        #        }
        #    )
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "values.update",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "range": "%s!A%s:%s%s"
                % (self._title, startRow, getColumnLetterOf(maxColumnCount), startRow + len(rows) - 1),
                "valueInputOption": "USER_ENTERED",  # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
                "body": {"majorDimension": "ROWS", "values": rows},
            }
        )

        # Update the local data in `_cells`:
        for rowNumBase1 in range(startRow, startRow + len(rows)):
            for colNumBase0 in range(maxColumnCount):
                self._cells[(colNumBase0 + 1, rowNumBase1)] = rows[rowNumBase1 - startRow][colNumBase0]

    def updateColumns(self, columns, startColumn=1):
        # Argument validation:
        # Ensure that `columns` is a list of lists:
        if not isinstance(columns, (list, tuple)):
            raise TypeError("columns arg must be a list/tuple of lists/tuples, not %s" % (type(columns).__name__))
        for column in columns:
            if not isinstance(column, (list, tuple)):
                raise TypeError("columns arg contains a non-list/tuple")

        if not isinstance(startColumn, int):
            raise TypeError("startColumn arg must be an int, not %s" % (type(startColumn).__name__))
        if startColumn < 1:
            raise ValueError("startColumn arg is 1-based, and must be 1 or greater, not %r" % (startColumn))

        if startColumn > self._columnCount:
            return  # No rows to update, so return.

        # Find out the max length of a column in `columns`. This will be the new rowCount for the sheet:
        maxRowCount = self._rowCount
        for column in columns:
            maxRowCount = max(maxRowCount, len(column))

        # Lengthen columns to the length of self._columnCount, and each column to the length of self._rowCount:
        for column in columns:
            column.extend([""] * (maxRowCount - len(column)))  # pad each column
        while len(columns) < (
            self._columnCount - startColumn + 1
        ):  # TODO - this could probably be made more performant if we use extend().
            columns.append([""] * self._rowCount)  # pad extra columns

        self._enlargeIfNeeded(len(columns) + startColumn - 1, None)

        # Send the API request that updates the Google sheet.
        # rangeCells = '%s!A%s:%s%s' % (self._title, startRow, getColumnLetterOf(maxColumnCount), stopRow - 1)
        # request = SHEETS_SERVICE.spreadsheets().values().update(
        #    spreadsheetId=self._spreadsheet._spreadsheetId,
        #    range='%s!%s1:%s%s' % (self._title, getColumnLetterOf(startColumn), getColumnLetterOf(startColumn + len(columns) - 1), maxRowCount),
        #    valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
        #    body={
        #        'majorDimension': 'COLUMNS',
        #        'values': columns,
        #        #'range': rangeCells,
        #        }
        #    )
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "values.update",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "range": "%s!%s1:%s%s"
                % (
                    self._title,
                    getColumnLetterOf(startColumn),
                    getColumnLetterOf(startColumn + len(columns) - 1),
                    maxRowCount,
                ),
                "valueInputOption": "USER_ENTERED",  # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
                "body": {"majorDimension": "COLUMNS", "values": columns},
            }
        )

        # Update the local data in `_cells`:
        for colNumBase1 in range(startColumn, startColumn + len(columns)):
            for rowNumBase0 in range(maxRowCount):
                self._cells[(colNumBase1, rowNumBase0 + 1)] = columns[colNumBase1 - startColumn][rowNumBase0]

    """
    def updateColumns(self, columns, startColumn=0, stopColumn=None, step=1):
        # Ensure that `columns` is a list of lists:
        if not isinstance(columns, (list, tuple)):
            raise TypeError('columns arg must be a list/tuple of lists/tuples, not %s' % (type(columns).__name__))
        for value in columns:
            if not isinstance(columns, (list, tuple)):
                raise TypeError('columns arg must be a list/tuple of lists/tuples, not %s' % (type(columns).__name__))

        if stopColumn is None:
            stopColumn = self._columnCount + 1

        # Lengthen columns to the length of self._columnCount, and each column to the length of self._rowCount:
        for column in columns:
            column.extend([''] * (self._columnCount - len(column))) # pad each column
        if len(columns) < self._columnCount:
            columns.extend([[''] * self._columnCount for i in range(self.stopColumn - len(columns) - 1)]) # pad extra columns

        self._enlargeIfNeeded(len(columns) + startColumn - 1, len(columns[0]))

        # Send the API request that updates the Google sheet.
        rangeCells = '%s!%s1:%s%s' % (self._title, getColumnLetterOf(startColumn), getColumnLetterOf(len(columns)), len(columns[0]))
        request = SHEETS_SERVICE.spreadsheets().values().update(
            spreadsheetId=self._spreadsheet._spreadsheetId,
            range=rangeCells,
            valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
            body={
                'majorDimension': 'COLUMNS',
                'values': columns,
                #'range': rangeCells,
                }
            )
        _logWriteRequest(); request.execute()

        # Update the local data in `_cells`:
        for colNumBase0 in range(len(columns)):
            for rowNumBase0 in range(len(columns[0])):
                self._cells[(colNumBase0+1, rowNumBase0+1)] = columns[colNumBase0][rowNumBase0]
    """

    def clear(self):
        # request = SHEETS_SERVICE.spreadsheets().values().update(
        #    spreadsheetId=self._spreadsheet._spreadsheetId,
        #    range='%s!A1:%s%s' % (self._title, getColumnLetterOf(self._columnCount), self._rowCount),
        #    valueInputOption='USER_ENTERED', # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
        #    body={
        #        'majorDimension': 'ROWS',
        #        'values': [[''] * self._columnCount for i in range(self._rowCount)],
        #        #'range': rangeCells,
        #        }
        #    )
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "values.update",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "range": "%s!A1:%s%s" % (self._title, getColumnLetterOf(self._columnCount), self._rowCount),
                "valueInputOption": "USER_ENTERED",  # Details at https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
                "body": {"majorDimension": "ROWS", "values": [[""] * self._columnCount for i in range(self._rowCount)]},
            }
        )

        # Update the local data in `_cells`:
        self._cells = {}

    def copyTo(self, destinationSpreadsheet):
        # NOTE: Don't update this method to allow ID or URL strings to be
        # passed, because we'll always need to call refresh() on the
        # spreadsheet object itself.

        if not isinstance(destinationSpreadsheet, Spreadsheet):
            raise TypeError(
                "destinationSpreadsheet must be of type Spreadsheet, not %s" % (type(destinationSpreadsheet).__name__)
            )

        # request = SHEETS_SERVICE.spreadsheets().sheets().copyTo(spreadsheetId=self._spreadsheet._spreadsheetId,
        #                                                 sheetId=self._sheetId,
        #                                                 body={'destinationSpreadsheetId': destinationSpreadsheet._spreadsheetId})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "sheets.copyTo",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "sheetId": self._sheetId,
                "body": {"destinationSpreadsheetId": destinationSpreadsheet._spreadsheetId},
            }
        )
        destinationSpreadsheet.refresh()  # Refresh the spreadsheet since its sheets list has changed.

    def delete(self):
        if len(self._spreadsheet.sheets) == 1:
            raise ValueError("Cannot delete all sheets; spreadsheets must have at least one sheet")

        # request = SHEETS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
        #    body={
        #        'requests': [{'deleteSheet': {'sheetId': self._sheetId}}]})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "batchUpdate",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "body": {"requests": [{"deleteSheet": {"sheetId": self._sheetId}}]},
            }
        )

        self._spreadsheet.refresh()  # Refresh the spreadsheet's list of sheets.

    def resize(self, columnCount=None, rowCount=None):
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
            return  # No resizing is taking place, so this function is a no-op.
        if rowCount == self._rowCount and columnCount == self._columnCount:
            return  # No change needed, so just return.

        # A None value means "use the current setting"
        if rowCount is None:
            rowCount = self._rowCount
        if columnCount is None:
            columnCount = self._columnCount

        if isinstance(columnCount, str):
            columnCount = getColumnNumberOf(columnCount)

        if not isinstance(rowCount, int):
            raise TypeError("rowCount arg must be an int, not %s" % (type(rowCount).__name__))
        if not isinstance(columnCount, int):
            raise TypeError("columnCount arg must be an int, not %s" % (type(columnCount).__name__))

        if rowCount < 1:
            raise TypeError("rowCount arg must be a positive nonzero int, not %r" % (rowCount))
        if columnCount < 1:
            raise TypeError("columnCount arg must be a positive nonzero int, not %r" % (columnCount))

        # request = SHEETS_SERVICE.spreadsheets().batchUpdate(spreadsheetId=self._spreadsheet._spreadsheetId,
        # body={
        #    'requests': [{'updateSheetProperties': {'properties': {'sheetId': self._sheetId,
        #                                                           'gridProperties': {'rowCount': rowCount,
        #                                                                              'columnCount': columnCount}},
        #                                            'fields': 'gridProperties'}}]})
        # _logWriteRequest(); request.execute()
        _makeRequest(
            "batchUpdate",
            **{
                "spreadsheetId": self._spreadsheet._spreadsheetId,
                "body": {
                    "requests": [
                        {
                            "updateSheetProperties": {
                                "properties": {
                                    "sheetId": self._sheetId,
                                    "gridProperties": {"rowCount": rowCount, "columnCount": columnCount},
                                },
                                "fields": "gridProperties",
                            }
                        }
                    ]
                },
            }
        )
        self._rowCount = rowCount
        self._columnCount = columnCount

    def __getitem__(self, *key):
        if isinstance(key[0], str):
            # Key is assumed to be an address like 'A1'
            return self.get(key[0])
        elif len(key[0]) == 2:
            # Key is assumed to be a tuple of column, row addresses
            return self.get(*key[0])
        else:
            # Let get() handle raising an error.
            return self.get(*key)

    def __setitem__(self, *args):
        if len(args) < 2:
            raise TypeError("__setitem__() requires at least two arguments")
        key = args[:-1]
        value = args[-1]

        if isinstance(key[0], str):
            # Key is assumed to be an address like 'A1'
            return self.update(key[0], value)
        elif len(key[0]) == 2:
            # Key is assumed to be a tuple of column, row addresses
            return self.update(*key[0], value)
        else:
            # Let update() handle raising an error.
            return self.update(*key, value)

    def __iter__(self):
        return iter(self.getRows())


def _getTabColorArg(value):
    if isinstance(value, str) and value in COLORS:
        # value is a color string from colorvalues.py, like 'red' or 'black'
        tabColorArg = {
            "red": COLORS[value][0],
            "green": COLORS[value][1],
            "blue": COLORS[value][2],
            "alpha": COLORS[value][3],
        }

    # elif value is None: # TODO - apparently there's no way to reset the color through the api?
    #    tabColorArg = {} # Reset the color
    elif isinstance(value, (list, tuple)) and len(value) in (3, 4):
        # value is a tuple of three or four floats (ranged from 0.0 to 1.0)
        tabColorArg = {"red": float(value[0]), "green": float(value[1]), "blue": float(value[2])}
        try:
            tabColorArg["alpha"] = value[3]
        except:
            tabColorArg["alpha"] = 1.0
    elif value is None:
        return None  # Represents no tabColor setting.
    elif type(value) == dict:
        tabColorArg = value
    else:
        raise ValueError(
            "value argument must be a color string like 'red', a 3- or 4-float tuple for an RGB or RGBA value, or a dict"
        )

    # Set any remaining unspecified defaults.
    tabColorArg.setdefault("red", 0.0)
    tabColorArg.setdefault("green", 0.0)
    tabColorArg.setdefault("blue", 0.0)
    tabColorArg.setdefault("alpha", 1.0)
    tabColorArg["red"] = float(tabColorArg["red"])
    tabColorArg["green"] = float(tabColorArg["green"])
    tabColorArg["blue"] = float(tabColorArg["blue"])
    tabColorArg["alpha"] = float(tabColorArg["alpha"])
    return tabColorArg


def convertToColumnRowInts(arg):
    if not isinstance(arg, str):
        raise TypeError("argument must be a grid cell str, like 'A1', not of type %s" % (type(arg).__name__))
    if not arg.isalnum() or not arg[0].isalpha() or not arg[-1].isdecimal():
        raise ValueError("argument must be a grid cell str, like 'A1', not %r" % (arg))

    for i in range(1, len(arg)):
        if arg[i].isdecimal():
            column = getColumnNumberOf(arg[:i])
            row = int(arg[i:])
            return (column, row)

    assert False  # pragma: no cover We know this will always return before this point because arg[-1].isdecimal().


def createSpreadsheet(title="Untitled spreadsheet"):
    if not IS_INITIALIZED:
        init()  # Initialize this module if not done so already.
    # request = SHEETS_SERVICE.spreadsheets().create(body={
    #    'properties': {'title': title}
    #    })
    # _logWriteRequest(); response = request.execute()
    response = _makeRequest("create", **{"body": {"properties": {"title": title}}})

    return Spreadsheet(response["spreadsheetId"])


def getIdFromUrl(url):
    # https://docs.google.com/spreadsheets/d/16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c/edit#gid=0
    if url.startswith("https://docs.google.com/spreadsheets/d/"):
        spreadsheetId = url[39 : url.find("/", 39)]

        # TODO URL could also have been in the form: https://docs.google.com/spreadsheets/u/3/d/1_MNPzTbsGQsMWVG9Di91U03sjbs_1SUUKPZVcnzlPbA/edit
    else:
        spreadsheetId = url

    if re.match("^([a-zA-Z0-9]|_|-)+$", spreadsheetId) is None:
        raise ValueError("url argument must be an alphanumeric id or a full URL")
    return spreadsheetId


def getColumnLetterOf(columnNumber):
    """getColumnLetterOf(1) => 'A', getColumnLetterOf(27) => 'AA'"""
    if not isinstance(columnNumber, int):
        raise TypeError("columnNumber must be an int, not a %r" % (type(columnNumber).__name__))
    if columnNumber < 1:
        raise ValueError("columnNumber must be an int value of at least 1")

    letters = []
    while columnNumber > 0:
        columnNumber, remainder = divmod(columnNumber, 26)
        if remainder == 0:
            remainder = 26
            columnNumber -= 1
        letters.append(chr(remainder + 64))
    return "".join(reversed(letters))


def getColumnNumberOf(columnLetter):
    """getColumnNumberOf('A') => 1, getColumnNumberOf('AA') => 27"""
    if not isinstance(columnLetter, str):
        raise TypeError("columnLetter must be a str, not a %r" % (type(columnLetter).__name__))
    if not columnLetter.isalpha():
        raise ValueError("columnLetter must be composed of only letters")

    columnLetter = columnLetter.upper()
    digits = []

    while columnLetter:
        digits.append(ord(columnLetter[0]) - 64)
        columnLetter = columnLetter[1:]

    number = 0
    place = 0
    for digit in reversed(digits):
        number += digit * (26**place)
        place += 1

    return number


def convertAddress(*address):
    if len(address) < 1 or len(address) > 2:
        raise TypeError('The address argument must be a singe string like "A1" or a tuple of two 1-based integers.')

    if isinstance(address[0], str):
        # Convert 'A2' to (1, 2)
        return convertToColumnRowInts(address[0])

    if isinstance(address[0], (tuple, list)) and len(address[0]) == 2:
        # If a tuple was passed, split it into two ints:
        if not isinstance(address[0][0], int) or not isinstance(address[0][1], int):
            raise TypeError('The address argument must be a singe string like "A1" or a tuple of two 1-based integers.')
        address = address[0][0], address[0][1]

    if isinstance(address[0], int) and isinstance(address[1], int) and address[0] > 0 and address[1] > 0:
        # Convert (1, 2) to 'A2'
        return getColumnLetterOf(address[0]) + str(address[1])

    raise TypeError('The address argument must be a singe string like "A1" or a tuple of two 1-based integers.')


def init(
    credentialsFile='.',
    sheetsTokenFile="token-sheets.pickle",
    driveTokenFile="token-drive.pickle",
    _raiseException=True,
):
    global SHEETS_SERVICE, DRIVE_SERVICE, IS_INITIALIZED

    # Set this to False, in case module was initialized before but this current initialization fails.
    IS_INITIALIZED = False

    # If the credentialsFile parameter is None, assume the credentials json file in the cwd.
    # In version 2023.3.14 and before (and in Automate the Boring Stuff
    # 2nd Edition), the credentials file had to be credentials-sheets.json.
    # But this isn't the name it has when you download it from Google
    # Cloud Console, so we'll just use the client_secret_*.json filename
    # format it already has, and fall back on credentials-sheets.json.
    # If credentialsFile is a folder name, use that folder to search for the credentials file.

    # credentialsFile is a bit misleading of a name because it can be a file or a folder (that contains the credentials file)
    if os.path.isdir(os.path.abspath(credentialsFile)):
        # If credentialsFile is a folder, search that folder for credentials-sheets.json or client_secret_*.json files:
        possibleCredentialsFiles = []
        for filename in os.listdir(os.path.abspath(credentialsFile)):
            if (filename.startswith('client_secret_') and filename.endswith('.json')) or filename == 'credentials-sheets.json':
                possibleCredentialsFiles.append(filename)
        if len(possibleCredentialsFiles) == 0:
            credentialsFile = 'credentials-sheets.json'  # Setting it to this nonexistant file will trigger the later EZSheetsException.
        elif len(possibleCredentialsFiles) > 1:
            raise EZSheetsException('You must specify a credentialsFile argument to init() because multiple possible credential files exist in ' + str(os.getcwd()) + ': ' + ', '.join(possibleCredentialsFiles))
        elif len(possibleCredentialsFiles) == 1:
            credentialsFile = os.path.join(os.path.abspath(credentialsFile), possibleCredentialsFiles[0])

    try:
        if not os.path.exists(credentialsFile):
            raise EZSheetsException(
                'Can\'t find credentials file at %s. You can download this file from https://developers.google.com/sheets/api/quickstart/python  and clicking "Enable the Google Sheets API". Rename the downloaded file to credentials-sheets.json.'
                % (os.path.abspath(credentialsFile))
            )

        # Find the token files, assume they are in the same folder as the credentials file:
        if not os.path.isabs(sheetsTokenFile):
            sheetsTokenFile = os.path.join(os.path.dirname(os.path.abspath(credentialsFile)), sheetsTokenFile)
        if not os.path.isabs(driveTokenFile):
            driveTokenFile = os.path.join(os.path.dirname(os.path.abspath(credentialsFile)), driveTokenFile)

        # Log in to Google Sheets API to generate token-sheets.pickle.
        creds = None
        # The file token-sheets.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(sheetsTokenFile):
            with open(sheetsTokenFile, "rb") as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(credentialsFile, SCOPES_SHEETS)
                creds = flow.run_local_server()
            # Save the credentials for the next run
            with open(sheetsTokenFile, "wb") as token:
                pickle.dump(creds, token)

        SHEETS_SERVICE = build("sheets", "v4", credentials=creds)

        # Log in to Google Drive API to generate token-drive.pickle.
        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(driveTokenFile):
            with open(driveTokenFile, "rb") as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(credentialsFile, SCOPES_DRIVE)
                creds = flow.run_local_server()
            # Save the credentials for the next run
            with open(driveTokenFile, "wb") as token:
                pickle.dump(creds, token)

        DRIVE_SERVICE = build("drive", "v3", credentials=creds)

        IS_INITIALIZED = True
        return IS_INITIALIZED
    except:
        if _raiseException:
            raise
        else:
            return False


def listSpreadsheets():
    if not IS_INITIALIZED:
        init()

    spreadsheets = {}  # key is ID, value is title
    page_token = None
    while True:
        # response = DRIVE_SERVICE.files().list(q="mimeType='application/vnd.google-apps.spreadsheet'",
        #                                      spaces='drive',
        #                                      fields='nextPageToken, files(id, name)',
        #                                      pageToken=page_token).execute()
        response = _makeRequest(
            "drive.list",
            **{
                "q": "mimeType='application/vnd.google-apps.spreadsheet'",
                "spaces": "drive",
                "fields": "nextPageToken, files(id, name)",
                "pageToken": page_token,
            }
        )
        for file in response.get("files", []):
            spreadsheets[file.get("id")] = file.get("name")
        page_token = response.get("nextPageToken", None)
        if page_token is None:
            break
    return spreadsheets


def upload(filename):
    if not IS_INITIALIZED:
        init()
    # TODO - be able to pass a file object for `filename`, not just a string name of a file on the hard drive.

    if filename.lower().endswith(".xlsx"):
        mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif filename.lower().endswith(".ods"):
        mimeType = "application/x-vnd.oasis.opendocument.spreadsheet"
    elif filename.lower().endswith(".csv"):
        mimeType = "text/csv"
    elif filename.lower().endswith(".tsv"):
        mimeType = "text/tab-separated-values"
    else:
        raise ValueError(
            "File to upload must be a .xlsx (Excel), .ods (OpenOffice), .csv (Comma-separated), or .tsv (Tab-separated) file type."
        )

    if not os.path.exists(filename):
        raise FileNotFoundError("Unable to find a file named %s" % (os.path.abspath(filename)))

    media = MediaFileUpload(filename, mimetype=mimeType)
    # file = DRIVE_SERVICE.files().create(body={'name': filename, 'mimeType': 'application/vnd.google-apps.spreadsheet'},
    #                                    media_body=media,
    #                                    fields='id').execute()
    file = _makeRequest(
        "drive.create",
        **{
            "body": {"name": os.path.basename(filename), "mimeType": "application/vnd.google-apps.spreadsheet"},
            "media_body": media,
            "fields": "id",
        }
    )
    return Spreadsheet(file.get("id"))


init(_raiseException=False)
# s = Spreadsheet('https://docs.google.com/spreadsheets/d/1lRyPHuaLIgqYwkCTJYexbZUO1dcWeunm69B0L7L4ZQ8/edit#gid=0')
