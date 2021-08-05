EZSheets
========

A Pythonic interface to the Google Sheets API that actually works as of August 2021.

Installation
------------

To install with pip, run:

    pip install ezsheets

Quickstart Guide
----------------

First you need to enable the Google Sheets API for your Google account:

* Go to https://developers.google.com/sheets/api/quickstart/python
* Click the blue "Enable the Google Sheets API" button.
* If you aren't already logged in, log in to your Google account on the login page that appears. (I recommend using a separate Google account specifically made for your Python scripts.)
* On the window that appears, click the blue "Download Client Configuration" button to download the credentials.json file.
* Rename this file to credentials-sheets.json.
* Place this file in the same folder as your Python script.

Next install the EZSheets module:

    pip install --upgrade ezsheets

(Use `pip3` on macOS and Linux.)

The first time you call an EZSheets function, the module will use your credentials-sheets.json file to generate a token-sheets.pickle and token-drive.pickle file. Don't share these files: Treat these files the same as you would your Google account password.

Create a `Spreadsheet` object by using the Spreadsheet's URL:

    >>> import ezsheets
    >>> s = ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c/edit#gid=0')

You can also just provide the spreadsheet ID part of the URL:

    >>> s = ezsheets.Spreadsheet('16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c')

`Spreadsheet` objects have a `title` and `spreadsheetId` attributes:

    >>> s.title
    'Class Data Example'
    >>> s.title = 'Class Data'
    >>> s.title
    'Class Data'
    >>> s.spreadsheetId
    '16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c'

`Spreadsheet` objects also have a `sheets` attribute, which is a list of `Sheet` objects:

    >>> s.sheets
    (Sheet(title='Sheet3', sheetId=314007586, rowCount=1000, columnCount=26), Sheet(title='Foobar', sheetId=2075929783, rowCount=1000, columnCount=27), Sheet(title='Class Data', sheetId=0, rowCount=101, columnCount=22, frozenRowCount=1), Sheet(title='Sheet2', sheetId=880141843, rowCount=1000, columnCount=26))
    >>> s.sheetTitles
    ('Sheet3', 'Foobar', 'Class Data', 'Sheet2')
    >>> sh = s.sheets[0]

You can then view the size and title of a sheet:

    >>> sh = s.sheets[0]
    >>> sh.title
    'Sheet3'
    >>> sh.title = 'My New Title'
    >>> sh.title
    'My New Title'
    >>> sh.columnCount, sh.rowCount
    (26, 1000)

You can also get or update data in a specific cell, row, or column:

    >>> sh.get(1,1)
    'fads'
    >>> sh.update(1, 1, 'New cell value')
    >>> sh.getRow(1)
    ['New cell value', 'fe', 'fa', 'ewafwe', 'f', 'ew', 'ewafawef', 'ewf', 'ewf', 'ew', 'fewa', 'f', 'ew', '', '', '', '', '', '', 'ewf', 'ewafewaf', 'ewfewf', '', 'f', 'ewfewafewaf', 'ewfew']
    >>> sh.updateRow(['cell A', 'cell B', 'cell C'])
    Traceback (most recent call last):
      File "<stdin>", line 1, in <module>
    TypeError: updateRow() missing 1 required positional argument: 'values'
    >>> sh.updateRow(1, ['cell A', 'cell B', 'cell C'])
    >>> sh.getColumn(1)
    ['cell A']
    >>> sh.update(1, 2, 'another value')
    >>> sh.getColumn(1)
    ['cell A', 'another value']
    >>> sh.updateAll([['CELL A', 'ANOTHER VALUE', 'CELL C'], ['ANOTHER VALUE']])
    >>> sh.getRows()
    [['CELL A', 'ANOTHER VALUE', 'CELL C'], ['ANOTHER VALUE']]

If the data on the Google Sheet changes, you can refresh your local copy of the data:

    >>> sh.refresh() # Updates the Sheet object.
    >>> s.refresh()  # Updates the Spreadsheet object and all its sheets.

You can rearrange the order of the sheets in the spreadsheet:

    >>> s.sheetTitles
    ('My New Title', 'Foobar', 'Class Data', 'Sheet2')
    >>> s.sheets[0].index
    0
    >>> s.sheets[0].index = 2
    >>> s.sheetTitles
    ('Foobar', 'Class Data', 'My New Title', 'Sheet2')
    >>> s.sheets[2].index = 0
    >>> s.sheetTitles
    ('My New Title', 'Foobar', 'Class Data', 'Sheet2')

You can recolor the tabs as well. (Currently you can't reset the tab color back to no color.)



Contribute
----------

If you'd like to contribute to EZSheets, check out https://github.com/asweigart/ezsheets
