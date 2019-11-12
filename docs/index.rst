EZSheets
========

A Pythonic interface to the Google Sheets API that actually works as of November 2019.


Installation
------------

EZSheets can be installed from PyPI using `pip`:

    ``pip install ezsheets``

On macOS and Linux, installing EZSheets for Python 3 is done with `pip3`:

    ``pip3 install ezsheets``

If you run into permissions errors, try installing with the `--user` option:

    ``pip install --user ezsheets``

Before you can use EZSheets, you need to enable the Google Sheets and
Google Drive APIs for your Google account. Visit the following web pages
and click the Enable API buttons at the top of each:

* https://console.developers.google.com/apis/library/sheets.googleapis.com/
* https://console.developers.google.com/apis/library/drive.googleapis.com/

You'll also need to obtain three files, which you should save in the same
folder as your .py Python script that uses EZSheets:

* A credentials file named credentials-sheets.json
* A token for Google Sheets named token-sheets.pickle
* A token for Google Drive named token-drive.pickle

The credentials file will generate the token files. The easiest way to
obtain a credentials file is to go to the Google Sheets Python Quickstart
page at https://developers.google.com/sheets/api/quickstart/python/ and click the
blue **Enable the Google Sheets API** button. You'll
need to log in to your Google account to view this page.

Clicking this button will bring up a window with a **Download Client
Configuration** link that lets you download a *credentials.json* file. Rename
this file to *credentials-sheets.json* and place it in the same folder as your
Python scripts.

Once you have a *credentials-sheets.json* file, run the import ezsheets module.
The first time you import the EZSheets module, it will open a new
browser window for you to log in to your Google account. Click **Allow**.

The message about Quickstart comes from the fact that you downloaded
the credentials file from the Google Sheets Python Quickstart page.
Note that this window will open twice: first for Google Sheets access and second
for Google Drive access. EZSheets uses Google Drive access to upload,
download, and delete spreadsheets.

After you log in, the browser window will prompt you to close it, and
the *token-sheets.pickle* and *token-drive.pickle* files will appear in the same folder
as *credentials-sheets.json*. You only need to go through this process the first
time you run ``import ezsheets``.

If you encounter an error after clicking **Allow** and the page seems to
hang, make sure you have first enabled the Google Sheets and Drive APIs
from the links at the start of this section. It may take a few minutes for
Google's servers to register this change, so you may have to wait before
you can use EZSheets.

Don't share the credential or token files with anyone: treat them
like passwords.


Authentication and Credentials Setup
------------------------------------

You need a Google account to access Google Sheets.

.. code-block::
    >>> import ezsheets

    >>> ss = ezsheets.createSpreadsheet(title='My New Spreadsheet')

    >>> ss.sheetTitles
    ('Sheet1',)

    >>> ss.sheets
    (Sheet(sheetId=0, title='Sheet1', rowCount=1000, columnCount=26),)

    >>> sh = ss.sheets[0]
    >>> sh.title
    'Sheet1'

    >>> sh.updateRow(1, ['Name', 'Species', 'Color', 'Weight'])
    >>> sh.updateRow(2, ('Zophie', 'Cat', 'Gray', 11))
    >>> sh[1, 1]
    'Name'
    >>> sh[1, 2]
    'Zophie'
    >>> sh['A2']
    'Zophie'
    >>> sh['A2'] = 'Pooka'
    >>> sh['A2']
    'Pooka'

The API section contains complete documentation.

Unit Tests
----------

The unit test suite takes approximately 7 minutes to run, due to the throttling.


API
---

.. automodule:: ezsheets
    :members:

