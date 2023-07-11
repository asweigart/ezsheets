"""
These tests check to make sure that the interactive shell examples in Automate the Boring Stuff 2nd Edition still run. It is critically important
that this code still works as it is printed in the book.
"""

import os

THIS_TEST_SCRIPT_FOLDER = os.path.dirname(os.path.abspath(__file__))


def test_page333():
    """
    Page 333
    >>> import ezsheets
    >>> ss = ezsheets.Spreadsheet('1J-Jx6Ne2K_vqI9J2SO-TAXOFbxx_9tUjwnkPC22LjeU')
    >>> ss
    Spreadsheet(spreadsheetId='1J-Jx6Ne2K_vqI9J2SO-TAXOFbxx_9tUjwnkPC22LjeU')
    >>> ss.title
    'Education Data'
    """
    import ezsheets
    dummy_ss = ezsheets.createSpreadsheet('Education Data')  # (setup)
    ss = ezsheets.Spreadsheet(dummy_ss.spreadsheetId)  # (adjusted)
    assert ss.title == 'Education Data'
    ss.delete(permanent=True)  # (teardown)


def test_page334_create():
    """
    Page 334
    >>> import ezsheets
    >>> ss = ezsheets.createSpreadsheet('Title of My New Spreadsheet')
    >>> ss.title
    'Title of My New Spreadsheet'
    """

    import ezsheets
    ss = ezsheets.createSpreadsheet('Title of My New Spreadsheet')
    assert ss.title == 'Title of My New Spreadsheet'
    ss.delete(permanent=True)  # (teardown)


def test_page334_upload_list():
    """
    >>> import ezsheets
    >>> ss = ezsheets.upload('my_spreadsheet.xlsx')
    >>> ss.title
    'my_spreadsheet'

    >>> ezsheets.listSpreadsheets()
    {'1J-Jx6Ne2K_vqI9J2SO-TAXOFbxx_9tUjwnkPC22LjeU': 'Education Data'}
    """
    import ezsheets
    ss = ezsheets.upload(os.path.join(THIS_TEST_SCRIPT_FOLDER, 'my_spreadsheet.xlsx'))  # (adjusted)
    assert ss.title == 'my_spreadsheet'
    assert (ss.spreadsheetId, 'my_spreadsheet') in ezsheets.listSpreadsheets().items()
    ss.delete(permanent=True)  # (teardown)


def test_page334_335():
    """
    Page 334-335
    >>> import ezsheets
    >>> ss = ezsheets.Spreadsheet('1J-Jx6Ne2K_vqI9J2SO-TAXOFbxx_9tUjwnkPC22LjeU')
    >>> ss.title # The title of the spreadsheet.
    'Education Data'
    >>> ss.title = 'Class Data' # Change the title.
    >>> ss.spreadsheetId # The unique ID (this is a read-only attribute).
    '1J-Jx6Ne2K_vqI9J2SO-TAXOFbxx_9tUjwnkPC22LjeU'
    >>> ss.url # The original URL (this is a read-only attribute).
    'https://docs.google.com/spreadsheets/d/1J-Jx6Ne2K_vqI9J2SO-
    TAXOFbxx_9tUjwnkPC22LjeU/'
    >>> ss.sheetTitles # The titles of all the Sheet objects
    ('Students', 'Classes', 'Resources')
    >>> ss.sheets # The Sheet objects in this Spreadsheet, in order.
    (<Sheet sheetId=0, title='Students', rowCount=1000, columnCount=26>, <Sheet
    sheetId=1669384683, title='Classes', rowCount=1000, columnCount=26>, <Sheet
    sheetId=151537240, title='Resources', rowCount=1000, columnCount=26>)
    >>> ss[0] # The first Sheet object in this Spreadsheet.
    <Sheet sheetId=0, title='Students', rowCount=1000, columnCount=26>
    >>> ss['Students'] # Sheets can also be accessed by title.
    <Sheet sheetId=0, title='Students', rowCount=1000, columnCount=26>
    >>> del ss[0] # Delete the first Sheet object in this Spreadsheet.
    >>> ss.sheetTitles # The "Students" Sheet object has been deleted:
    ('Classes', 'Resources')
    """

    import ezsheets
    dummy_ss = ezsheets.createSpreadsheet('Education Data')  # (setup)
    ss = ezsheets.Spreadsheet(dummy_ss.spreadsheetId)  # (adjusted)
    assert ss.title == 'Education Data'  # The title of the spreadsheet.
    ss.title = 'Class Data'  # Change the title.
    assert ss.spreadsheetId
    assert ss.url == 'https://docs.google.com/spreadsheets/d/' + ss.spreadsheetId + '/'
    ss.createSheet('Students')  # (setup)
    ss.createSheet('Classes')  # (setup)
    ss.createSheet('Resources')  # (setup)
    ss['Sheet1'].delete()  # (setup)
    assert ss.sheetTitles == ('Students', 'Classes', 'Resources')
    assert len(ss.sheets) == 3  # (adjusted)
    assert ss[0]
    assert ss['Students']
    del ss[0]
    assert ss.sheetTitles == ('Classes', 'Resources')

    ss.delete(permanent=True)  # (teardown)


def test_page335_refresh():
    """
    >>> ss.refresh()
    """
    import ezsheets
    dummy_ss = ezsheets.createSpreadsheet('Class Data')  # (setup)
    dummy_ss.refresh()
    dummy_ss.delete(permanent=True)  # (teardown)


def test_page335_download_upload():
    """
    >>> import ezsheets
    >>> ss = ezsheets.Spreadsheet('1J-Jx6Ne2K_vqI9J2SO-TAXOFbxx_9tUjwnkPC22LjeU')
    >>> ss.title
    'Class Data'
    >>> ss.downloadAsExcel() # Downloads the spreadsheet as an Excel file.
    'Class_Data.xlsx'
    >>> ss.downloadAsODS() # Downloads the spreadsheet as an OpenOffice file.
    'Class_Data.ods'
    >>> ss.downloadAsCSV() # Only downloads the first sheet as a CSV file.
    'Class_Data.csv'
    >>> ss.downloadAsTSV() # Only downloads the first sheet as a TSV file.
    'Class_Data.tsv'
    >>> ss.downloadAsPDF() # Downloads the spreadsheet as a PDF.
    'Class_Data.pdf'
    >>> ss.downloadAsHTML() # Downloads the spreadsheet as a ZIP of HTML files.
    'Class_Data.zip'

    >>> ss.downloadAsExcel('a_different_filename.xlsx')
    'a_different_filename.xlsx'
    """

    import ezsheets
    dummy_ss = ezsheets.createSpreadsheet('Class Data')  # (setup)
    ss = ezsheets.Spreadsheet(dummy_ss.spreadsheetId)  # (adjusted)
    assert ss.title == 'Class Data'

    assert ss.downloadAsExcel() == 'Class_Data.xlsx'
    assert os.path.exists('Class_Data.xlsx')  # (teardown)
    os.unlink('Class_Data.xlsx')  # (teardown)

    assert ss.downloadAsODS() == 'Class_Data.ods'
    assert os.path.exists('Class_Data.ods')  # (teardown)
    os.unlink('Class_Data.ods')  # (teardown)

    assert ss.downloadAsCSV() == 'Class_Data.csv'
    assert os.path.exists('Class_Data.csv')  # (teardown)
    os.unlink('Class_Data.csv')  # (teardown)

    assert ss.downloadAsTSV() == 'Class_Data.tsv'
    assert os.path.exists('Class_Data.tsv')  # (teardown)
    os.unlink('Class_Data.tsv')  # (teardown)

    assert ss.downloadAsPDF() == 'Class_Data.pdf'
    assert os.path.exists('Class_Data.pdf')  # (teardown)
    os.unlink('Class_Data.pdf')  # (teardown)

    assert ss.downloadAsHTML() == 'Class_Data.zip'
    assert os.path.exists('Class_Data.zip')  # (teardown)
    os.unlink('Class_Data.zip')  # (teardown)

    assert ss.downloadAsExcel('a_different_filename.xlsx') == 'a_different_filename.xlsx'
    assert os.path.exists('a_different_filename.xlsx')  # (teardown)
    os.unlink('a_different_filename.xlsx')  # (teardown)

    ss.delete(permanent=True)  # (teardown)


def test_page336_deleting_spreadsheets():
    """
    >>> import ezsheets
    >>> ss = ezsheets.createSpreadsheet('Delete me') # Create the spreadsheet.
    >>> ezsheets.listSpreadsheets() # Confirm that we've created a spreadsheet.
    {'1aCw2NNJSZblDbhygVv77kPsL3djmgV5zJZllSOZ_mRk': 'Delete me'}
    >>> ss.delete()  # Delete the spreadsheet.
    >>> ezsheets.listSpreadsheets()
    {}

    >>> ss.delete(permanent=True)
    """

    # NOTE - delete() moves stuff to trash, and ss in the trash will still show up in listSpreadsheets()!!!!

    import ezsheets
    ss = ezsheets.createSpreadsheet('Delete me PAGE336') # Create the spreadsheet. (adjusted since there can be other 'Delete me' spreadsheets on this account)
    assert (ss.spreadsheetId, 'Delete me PAGE336') in ezsheets.listSpreadsheets().items()  # Confirm that we've created a spreadsheet.
    savedId = ss.spreadsheetId  # (setup)
    ss.delete()  # Delete the spreadsheet.
    assert (savedId, 'Delete me PAGE336') in ezsheets.listSpreadsheets().items()  # (teardown)
    ss.delete(permanent=True)  # (teardown)
    assert (savedId, 'Delete me PAGE336') not in ezsheets.listSpreadsheets().items()  # (teardown)

    #ss = ezsheets.createSpreadsheet('Delete me PAGE336') # Create the spreadsheet.
    #assert (ss.spreadsheetId, 'Delete me PAGE336') in ezsheets.listSpreadsheets().items()  # Confirm that we've created a spreadsheet.
    #savedId = ss.spreadsheetId  # (setup)
    #ss.delete(permanent=True)  # Delete the spreadsheet.
    #assert (savedId, 'Delete me PAGE336') not in ezsheets.listSpreadsheets().items()  # (teardown)

def test_page336_337_sheet_objects():
    """
    >>> import ezsheets
    >>> ss = ezsheets.Spreadsheet('1J-Jx6Ne2K_vqI9J2SO-TAXOFbxx_9tUjwnkPC22LjeU')
    >>> ss.sheets # The Sheet objects in this Spreadsheet, in order.
    (<Sheet sheetId=1669384683, title='Classes', rowCount=1000, columnCount=26>,
    <Sheet sheetId=151537240, title='Resources', rowCount=1000, columnCount=26>)
    >>> ss.sheets[0] # Gets the first Sheet object in this Spreadsheet.
    <Sheet sheetId=1669384683, title='Classes', rowCount=1000, columnCount=26>
    >>> ss[0] # Also gets the first Sheet object in this Spreadsheet.
    <Sheet sheetId=1669384683, title='Classes', rowCount=1000, columnCount=26>

    >>> ss.sheetTitles # The titles of all the Sheet objects in this Spreadsheet.
    ('Classes', 'Resources')
    >>> ss['Classes'] # Sheets can also be accessed by title.
    <Sheet sheetId=1669384683, title='Classes', rowCount=1000, columnCount=26>
    """

    import ezsheets
    dummy_ss = ezsheets.createSpreadsheet('Education Data')  # (setup)
    ss = ezsheets.Spreadsheet(dummy_ss.spreadsheetId)  # (adjusted)
    ss.createSheet('Classes')  # (setup)
    ss.createSheet('Resources')  # (setup)
    del ss['Sheet1']  # (setup)

    assert len(ss.sheets) == 2  # (adjusted)
    assert ss.sheets[0] == ss[0]

    assert ss.sheetTitles == ('Classes', 'Resources')


def test_page337_read_write():
    """
    >>> import ezsheets
    >>> ss = ezsheets.createSpreadsheet('My Spreadsheet')
    >>> sheet = ss[0] # Get the first sheet in this spreadsheet.
    >>> sheet.title
    'Sheet1'
    >>> sheet = ss[0]
    >>> sheet['A1'] = 'Name' # Set the value in cell A1.
    >>> sheet['B1'] = 'Age'
    >>> sheet['C1'] = 'Favorite Movie'
    >>> sheet['A1'] # Read the value in cell A1.
    'Name'
    >>> sheet['A2'] # Empty cells return a blank string.
    ''
    >>> sheet[2, 1] # Column 2, Row 1 is the same address as B1.
    'Age'
    >>> sheet['A2'] = 'Alice'
    >>> sheet['B2'] = 30
    >>> sheet['C2'] = 'RoboCop'
    """
    import ezsheets
    ss = ezsheets.createSpreadsheet('My Spreadsheet')
    sheet = ss[0] # Get the first sheet in this spreadsheet.
    assert sheet.title == 'Sheet1'
    sheet = ss[0]
    sheet['A1'] = 'Name' # Set the value in cell A1.
    sheet['B1'] = 'Age'
    sheet['C1'] = 'Favorite Movie'
    assert sheet['A1'] == 'Name' # Read the value in cell A1.

    assert sheet['A2'] == '' # Empty cells return a blank string.
    assert sheet[2, 1] == 'Age' # Column 2, Row 1 is the same address as B1.
    sheet['A2'] = 'Alice'
    assert sheet['A2'] == 'Alice'
    sheet['B2'] = 30
    assert sheet['B2'] == 30
    sheet['C2'] = 'RoboCop'
    assert sheet['C2'] == 'RoboCop'

    ss.delete(permanent=True)


def test_page338():
    """
    >>> import ezsheets
    >>> ezsheets.convertAddress('A2') # Converts addresses...
    (1, 2)
    >>> ezsheets.convertAddress(1, 2) # ...and converts them back, too.
    'A2'
    >>> ezsheets.getColumnLetterOf(2)
    'B'
    >>> ezsheets.getColumnNumberOf('B')
    2
    >>> ezsheets.getColumnLetterOf(999)
    'ALK'


    >>> ezsheets.getColumnNumberOf('ZZZ')
    18278
    """
    import ezsheets
    assert ezsheets.convertAddress('A2') == (1, 2)# Converts addresses...
    assert ezsheets.convertAddress(1, 2) == 'A2'# ...and converts them back, too.
    assert ezsheets.getColumnLetterOf(2) == 'B'
    assert ezsheets.getColumnNumberOf('B') == 2
    assert ezsheets.getColumnLetterOf(999) == 'ALK'
    assert ezsheets.getColumnNumberOf('ZZZ') == 18278

def test_page339_341():
    """
    >>> import ezsheets
    >>> ss = ezsheets.upload('produceSales.xlsx')
    >>> sheet = ss[0]
    >>> sheet.getRow(1) # The first row is row 1, not row 0.
    ['PRODUCE', 'COST PER POUND', 'POUNDS SOLD', 'TOTAL', '', '']
    >>> sheet.getRow(2)
    ['Potatoes', '0.86', '21.6', '18.58', '', '']
    >>> columnOne = sheet.getColumn(1)
    >>> sheet.getColumn(1)
    ['PRODUCE', 'Potatoes', 'Okra', 'Fava beans', 'Watermelon', 'Garlic',
    --snip--
    >>> sheet.getColumn('A') # Same result as getColumn(1)
    ['PRODUCE', 'Potatoes', 'Okra', 'Fava beans', 'Watermelon', 'Garlic',
    --snip--
    >>> sheet.getRow(3)
    ['Okra', '2.26', '38.6', '87.24', '', '']
    >>> sheet.updateRow(3, ['Pumpkin', '11.50', '20', '230'])
    >>> sheet.getRow(3)
    ['Pumpkin', '11.50', '20', '230', '', '']
    >>> columnOne = sheet.getColumn(1)
    >>> for i, value in enumerate(columnOne):
    ... # Make the Python list contain uppercase strings:
    ... columnOne[i] = value.upper()
    ...
    >>> sheet.updateColumn(1, columnOne) # Update the entire column in one
    request.


    >>> rows = sheet.getRows() # Get every row in the spreadsheet.
    >>> rows[0] # Examine the values in the first row.
    ['PRODUCE', 'COST PER POUND', 'POUNDS SOLD', 'TOTAL', '', '']
    >>> rows[1]
    ['POTATOES', '0.86', '21.6', '18.58', '', '']
    >>> rows[1][0] = 'PUMPKIN' # Change the produce name.
    >>> rows[1]
    ['PUMPKIN', '0.86', '21.6', '18.58', '', '']
    >>> rows[10]
    ['OKRA', '2.26', '40', '90.4', '', '']
    >>> rows[10][2] = '400' # Change the pounds sold.
    >>> rows[10][3] = '904' # Change the total.
    >>> rows[10]
    ['OKRA', '2.26', '400', '904', '', '']
    >>> sheet.updateRows(rows) # Update the online spreadsheet with the changes.


    >>> sheet.rowCount # The number of rows in the sheet.
    23758
    >>> sheet.columnCount # The number of columns in the sheet.
    6
    >>> sheet.columnCount = 4 # Change the number of columns to 4.
    >>> sheet.columnCount # Now the number of columns in the sheet is 4.
    4
    """

    import ezsheets
    ss = ezsheets.upload(os.path.join(THIS_TEST_SCRIPT_FOLDER, 'produceSales.xlsx'))  # (adjusted)
    sheet = ss[0]
    assert sheet.getRow(1) == ['PRODUCE', 'COST PER POUND', 'POUNDS SOLD', 'TOTAL', '', ''] # The first row is row 1, not row 0.

    assert sheet.getRow(2) == ['Potatoes', '0.86', '21.6', '18.58', '', '']

    columnOne = sheet.getColumn(1)
    assert sheet.getColumn(1)[0:6] == ['PRODUCE', 'Potatoes', 'Okra', 'Fava beans', 'Watermelon', 'Garlic',]  # (adjusted)

    assert sheet.getColumn('A')[0:6] == ['PRODUCE', 'Potatoes', 'Okra', 'Fava beans', 'Watermelon', 'Garlic',] # Same result as getColumn(1) # (adjusted)

    assert sheet.getRow(3) == ['Okra', '2.26', '38.6', '87.24', '', '']
    sheet.updateRow(3, ['Pumpkin', '11.50', '20', '230'])
    assert sheet.getRow(3) == ['Pumpkin', '11.50', '20', '230', '', '']

    columnOne = sheet.getColumn(1)
    for i, value in enumerate(columnOne):
        # Make the Python list contain uppercase strings:
        columnOne[i] = value.upper()

    sheet.updateColumn(1, columnOne) # Update the entire column in one request.


    rows = sheet.getRows() # Get every row in the spreadsheet.
    assert rows[0] == ['PRODUCE', 'COST PER POUND', 'POUNDS SOLD', 'TOTAL', '', ''] # Examine the values in the first row.
    assert rows[1] == ['POTATOES', '0.86', '21.6', '18.58', '', '']
    rows[1][0] = 'PUMPKIN' # Change the produce name.
    assert rows[1] == ['PUMPKIN', '0.86', '21.6', '18.58', '', '']
    assert rows[10] == ['OKRA', '2.26', '40', '90.4', '', '']
    rows[10][2] = '400' # Change the pounds sold.
    rows[10][3] = '904' # Change the total.
    assert rows[10] == ['OKRA', '2.26', '400', '904', '', '']
    sheet.updateRows(rows) # Update the online spreadsheet with the changes.


    assert sheet.rowCount == 23758 # The number of rows in the sheet.

    assert sheet.columnCount == 6 # The number of columns in the sheet.

    sheet.columnCount = 4 # Change the number of columns to 4.
    assert sheet.columnCount == 4 # Now the number of columns in the sheet is 4.


    ss.delete(permanent=True)  # (teardown)


def test_page341_342():
    """
    >>> import ezsheets
    >>> ss = ezsheets.createSpreadsheet('Multiple Sheets')
    >>> ss.sheetTitles
    ('Sheet1',)
    >>> ss.createSheet('Spam') # Create a new sheet at the end of the list of
    sheets.
    <Sheet sheetId=2032744541, title='Spam', rowCount=1000, columnCount=26>
    >>> ss.createSheet('Eggs') # Create another new sheet.
    <Sheet sheetId=417452987, title='Eggs', rowCount=1000, columnCount=26>
    >>> ss.sheetTitles
    ('Sheet1', 'Spam', 'Eggs')
    >>> ss.createSheet('Bacon', 0) # Create a sheet at index 0 in the list of
    sheets.
    <Sheet sheetId=814694991, title='Bacon', rowCount=1000, columnCount=26>
    >>> ss.sheetTitles
    ('Bacon', 'Sheet1', 'Spam', 'Eggs')


    >>> ss.sheetTitles
    ('Bacon', 'Sheet1', 'Spam', 'Eggs')
    >>> ss[0].delete() # Delete the sheet at index 0: the "Bacon" sheet.
    >>> ss.sheetTitles
    ('Sheet1', 'Spam', 'Eggs')
    >>> ss['Spam'].delete() # Delete the "Spam" sheet.
    >>> ss.sheetTitles
    ('Sheet1', 'Eggs')
    >>> sheet = ss['Eggs'] # Assign a variable to the "Eggs" sheet.
    >>> sheet.delete() # Delete the "Eggs" sheet.
    >>> ss.sheetTitles
    ('Sheet1',)
    >>> ss[0].clear() # Clear all the cells on the "Sheet1" sheet.
    >>> ss.sheetTitles # The "Sheet1" sheet is empty but still exists.
    ('Sheet1',)
    """
    import ezsheets
    ss = ezsheets.createSpreadsheet('Multiple Sheets')
    assert ss.sheetTitles == ('Sheet1',)
    ss.createSheet('Spam') # Create a new sheet at the end of the list of sheets.
    ss.createSheet('Eggs') # Create another new sheet.
    assert ss.sheetTitles == ('Sheet1', 'Spam', 'Eggs')
    ss.createSheet('Bacon', 0) # Create a sheet at index 0 in the list of sheets.
    assert ss.sheetTitles == ('Bacon', 'Sheet1', 'Spam', 'Eggs')

    ss[0].delete() # Delete the sheet at index 0: the "Bacon" sheet.
    assert ss.sheetTitles == ('Sheet1', 'Spam', 'Eggs')
    ss['Spam'].delete() # Delete the "Spam" sheet.
    assert ss.sheetTitles == ('Sheet1', 'Eggs')
    sheet = ss['Eggs'] # Assign a variable to the "Eggs" sheet.
    sheet.delete() # Delete the "Eggs" sheet.
    assert ss.sheetTitles == ('Sheet1',)
    ss[0].clear() # Clear all the cells on the "Sheet1" sheet.
    assert ss.sheetTitles == ('Sheet1',)

    ss.delete(permanent=True)  # (teardown)


def test_page343():
    """
    >>> import ezsheets
    >>> ss1 = ezsheets.createSpreadsheet('First Spreadsheet')
    >>> ss2 = ezsheets.createSpreadsheet('Second Spreadsheet')
    >>> ss1[0]
    <Sheet sheetId=0, title='Sheet1', rowCount=1000, columnCount=26>
    >>> ss1[0].updateRow(1, ['Some', 'data', 'in', 'the', 'first', 'row'])
    >>> ss1[0].copyTo(ss2) # Copy the ss1's Sheet1 to the ss2 spreadsheet.
    >>> ss2.sheetTitles # ss2 now contains a copy of ss1's Sheet1.
    ('Sheet1', 'Copy of Sheet1')
    """

    import ezsheets
    ss1 = ezsheets.createSpreadsheet('First Spreadsheet')
    ss2 = ezsheets.createSpreadsheet('Second Spreadsheet')
    ss1[0]
    ss1[0].updateRow(1, ['Some', 'data', 'in', 'the', 'first', 'row'])
    ss1[0].copyTo(ss2) # Copy the ss1's Sheet1 to the ss2 spreadsheet.
    assert ss2.sheetTitles == ('Sheet1', 'Copy of Sheet1') # ss2 now contains a copy of ss1's Sheet1.

    ss1.delete(permanent=True)  # (teardown)
    ss2.delete(permanent=True)  # (teardown)


def test_bean_count():
    # Just test to make sure this is still there. This URL is referenced by the "Finding Mistakes in a Spreadsheet" practice project section.
    import ezsheets
    ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/1jDZEdvSIh4TmZxccyy0ZXrH-ELlrwq8_YYiZrEOB4jg/edit?usp=sharing/')
