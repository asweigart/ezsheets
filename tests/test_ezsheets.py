from __future__ import division, print_function
import random
import pytest
import ezsheets

#now = time.time()
#random.seed(now)
#print('Random seed: %s' % (now))

"""
NOTE: This test requires a credentials.json and token.pickle file to be in the
same folder as this script. A new spreadsheet will be created for use by this
script.
"""


def checkIfSpreadsheetInOriginalState():
    # Check if there exists only one Sheet in the FIXED_SPREADSHEET Spreadsheet object.
    assert FIXED_SPREADSHEET.title == 'Delete Me'
    assert len(FIXED_SPREADSHEET) == 1
    assert FIXED_SPREADSHEET[0].title == 'Sheet1'
    print('READS=%s, WRITES=%s' % (ezsheets._READ_REQUESTS, ezsheets._WRITE_REQUESTS))


def addOriginalSheet():
    # If we delete all the sheets in FIXED_SPREADSHEET, call this to restore the original Sheet1 Sheet object.
    assert 'Sheet1' not in FIXED_SPREADSHEET.sheetTitles
    FIXED_SPREADSHEET.addSheet(title='Sheet1')
    checkIfSpreadsheetInOriginalState()


@pytest.fixture
def checkPreAndPostCondition():
    # Check that Spreadsheet is in original condition before the test.
    checkIfSpreadsheetInOriginalState()
    yield
    # Check that Spreadsheet is in original condition after the test.
    checkIfSpreadsheetInOriginalState()


def test_basic():
    pass # TODO - add unit tests


def test_getIdFromUrl():
    assert ezsheets.getIdFromUrl("https://docs.google.com/spreadsheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0") == "10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng"

    with pytest.raises(ValueError):
        ezsheets.getIdFromUrl(r"https://docs.google.com/spread sheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0")


def test_getColumnLetterOf():
    assert ezsheets.getColumnLetterOf(1) == 'A'
    assert ezsheets.getColumnLetterOf(26) == 'Z'
    assert ezsheets.getColumnLetterOf(27) == 'AA'
    assert ezsheets.getColumnLetterOf(702) == 'ZZ'

    with pytest.raises(TypeError):
        ezsheets.getColumnLetterOf('invalid arg')

    with pytest.raises(ValueError):
        ezsheets.getColumnLetterOf(0)

    with pytest.raises(ValueError):
        ezsheets.getColumnLetterOf(-1)


def test_getColumnNumber():
    assert ezsheets.getColumnNumber('A') == 1
    assert ezsheets.getColumnNumber('Z') == 26
    assert ezsheets.getColumnNumber('AA') == 27
    assert ezsheets.getColumnNumber('ZZ') == 702

    with pytest.raises(TypeError):
        ezsheets.getColumnNumber(1)

    with pytest.raises(ValueError):
        ezsheets.getColumnNumber('123')


def test_columnNumberLetterTranslation():
    for i in range(1, 5000):
        assert ezsheets.getColumnNumber(ezsheets.getColumnLetterOf(i)) == i


def test_convertToColumnRowInts():
    assert ezsheets.convertToColumnRowInts('A1') == (1, 1)
    assert ezsheets.convertToColumnRowInts('Z1') == (26, 1)
    assert ezsheets.convertToColumnRowInts('AA1') == (27, 1)
    assert ezsheets.convertToColumnRowInts('ZZ1') == (702, 1)

    assert ezsheets.convertToColumnRowInts('A10') == (1, 10)
    assert ezsheets.convertToColumnRowInts('ZZ10') == (702, 10)

    for i in range(1, 1000):
        for j in range(101):
            assert ezsheets.convertToColumnRowInts(ezsheets.getColumnLetterOf(i) + str(j)) == (i, j)

    with pytest.raises(ValueError):
        ezsheets.convertToColumnRowInts('1')
    with pytest.raises(ValueError):
        ezsheets.convertToColumnRowInts('A')
    with pytest.raises(ValueError):
        ezsheets.convertToColumnRowInts('')

    with pytest.raises(TypeError):
        ezsheets.convertToColumnRowInts(123)


def test__getTabColorArg():
    RED_COLOR = {'red': 1.0, 'green': 0.0, 'blue': 0.0, 'alpha': 1.0}
    assert ezsheets._getTabColorArg('red') == RED_COLOR
    assert ezsheets._getTabColorArg((1.0, 0.0, 0.0)) == RED_COLOR
    assert ezsheets._getTabColorArg((1.0, 0.0, 0.0, 1.0)) == RED_COLOR
    assert ezsheets._getTabColorArg((1.0, 0.0, 0.0, 1.0)) == RED_COLOR
    assert ezsheets._getTabColorArg([1.0, 0.0, 0.0]) == RED_COLOR

    with pytest.raises(ValueError):
        ezsheets._getTabColorArg('invalid value')


@pytest.fixture(scope='module')
def init():
    global FIXED_SPREADSHEET
    ezsheets.init()
    FIXED_SPREADSHEET = ezsheets.createSpreadsheet(title='Delete Me')
    checkIfSpreadsheetInOriginalState()


def test_Spreadsheet_attr(init, checkPreAndPostCondition):
    assert FIXED_SPREADSHEET.title == 'Delete Me'
    assert FIXED_SPREADSHEET.spreadsheetId != ''
    assert FIXED_SPREADSHEET.sheetTitles == ('Sheet1',)
    assert len(FIXED_SPREADSHEET) == 1


def test_addSheet_deleteSheet(init, checkPreAndPostCondition):
    newSheet1 = FIXED_SPREADSHEET.addSheet(title='New Sheet 1')

    assert 'New Sheet 1' in FIXED_SPREADSHEET.sheetTitles
    assert newSheet1 == FIXED_SPREADSHEET.sheets[1]
    assert newSheet1.title == 'New Sheet 1'
    assert newSheet1.index == 1

    newSheet2 = FIXED_SPREADSHEET.addSheet(title='New Sheet 2', index=1)
    assert newSheet2 == FIXED_SPREADSHEET.sheets[1]
    assert newSheet2.index == 1
    assert newSheet1 == FIXED_SPREADSHEET.sheets[2]
    assert newSheet1.index == 2

    newSheet1.delete()
    newSheet2.delete()


def test_getitem_delitem(init, checkPreAndPostCondition):
    assert FIXED_SPREADSHEET['Sheet1'].title == 'Sheet1'

    # Attempt get with invalid index:
    with pytest.raises(KeyError):
        FIXED_SPREADSHEET[99]
    with pytest.raises(KeyError):
        FIXED_SPREADSHEET[-99]
    with pytest.raises(KeyError):
        FIXED_SPREADSHEET['nonexistent title']
    with pytest.raises(KeyError):
        FIXED_SPREADSHEET[['invalid', 'key', 'type']]

    # Attempt delete with invalid index:
    with pytest.raises(KeyError):
        del FIXED_SPREADSHEET[99]
    with pytest.raises(KeyError):
        del FIXED_SPREADSHEET[-99]
    with pytest.raises(KeyError):
        del FIXED_SPREADSHEET['nonexistent title']
    with pytest.raises(TypeError):
        del FIXED_SPREADSHEET[['invalid', 'key', 'type']]

    newSheet = FIXED_SPREADSHEET.addSheet(title='Added Sheet')
    assert FIXED_SPREADSHEET[1] == newSheet # Get by int index.
    assert FIXED_SPREADSHEET[-1] == newSheet # Get by negative index.
    assert FIXED_SPREADSHEET[1].title == 'Added Sheet'

    assert FIXED_SPREADSHEET[1:2] == (newSheet,) # Get by slice

    # Delete by int index:
    del FIXED_SPREADSHEET[1]
    checkIfSpreadsheetInOriginalState()

    FIXED_SPREADSHEET.addSheet(title='Added Sheet 2')
    # Get by str title:
    assert FIXED_SPREADSHEET['Added Sheet 2'].title == 'Added Sheet 2'

    # Delete by str title:
    del FIXED_SPREADSHEET['Added Sheet 2']
    checkIfSpreadsheetInOriginalState()

    # Get multiple sheets with slice:
    newSheet3 = FIXED_SPREADSHEET.addSheet(title='Added Sheet 3')
    newSheet4 = FIXED_SPREADSHEET.addSheet(title='Added Sheet 4')
    assert FIXED_SPREADSHEET[1:3] == (newSheet3, newSheet4)

    # Delete multiple sheets with slice:
    del FIXED_SPREADSHEET[1:3] # deleting newSheet3 and newSheet4
    checkIfSpreadsheetInOriginalState()

    FIXED_SPREADSHEET.addSheet(title='Added Sheet 5')
    FIXED_SPREADSHEET.addSheet(title='Added Sheet 6')
    del FIXED_SPREADSHEET[1:] # deleting newSheet5 and newSheet6
    checkIfSpreadsheetInOriginalState()

    FIXED_SPREADSHEET.addSheet(title='Added Sheet 7', index=0)
    FIXED_SPREADSHEET.addSheet(title='Added Sheet 8', index=1)
    del FIXED_SPREADSHEET[:2]
    checkIfSpreadsheetInOriginalState()

    FIXED_SPREADSHEET.addSheet(title='Added Sheet 9')
    FIXED_SPREADSHEET.addSheet(title='Added Sheet 10')
    del FIXED_SPREADSHEET[3:0:-1]
    checkIfSpreadsheetInOriginalState()

    # Attempt to delete all sheets:
    with pytest.raises(ValueError):
        del FIXED_SPREADSHEET[0]

    # Attempt to delete all sheets:
    with pytest.raises(ValueError):
        del FIXED_SPREADSHEET[0:1]

    # Deleting with negative start or stop in slice is a no-op
    del FIXED_SPREADSHEET[-1:]
    del FIXED_SPREADSHEET[:-1]


def test_len(init, checkPreAndPostCondition):
    # The length of a Spreadsheet object is how many Sheet objects it contains.
    assert len(FIXED_SPREADSHEET) == len(FIXED_SPREADSHEET.sheets)
    FIXED_SPREADSHEET.addSheet(title='Length Test Sheet')
    assert len(FIXED_SPREADSHEET) == len(FIXED_SPREADSHEET.sheets)
    del FIXED_SPREADSHEET[-1]
    assert len(FIXED_SPREADSHEET) == len(FIXED_SPREADSHEET.sheets)


def test_changeSheetIndex(init, checkPreAndPostCondition):
    newSheet1 = FIXED_SPREADSHEET.addSheet(title='New Sheet 1')
    newSheet2 = FIXED_SPREADSHEET.addSheet(title='New Sheet 2')

    assert FIXED_SPREADSHEET.sheetTitles == ('Sheet1', 'New Sheet 1', 'New Sheet 2')

    # Move 'New Sheet 2' to front:

    newSheet2.index = 0  # LEFT OFF with error happening here, some quota issue? Am I just making too many calls with this test suite?
    assert newSheet2.index == 0
    assert FIXED_SPREADSHEET[0] == newSheet2
    assert FIXED_SPREADSHEET.sheetTitles == ('New Sheet 2', 'Sheet1', 'New Sheet 1')

    newSheet2.index = 1
    assert newSheet2.index == 1
    assert FIXED_SPREADSHEET[1] == newSheet2
    assert FIXED_SPREADSHEET[1].title == 'New Sheet 2'
    assert FIXED_SPREADSHEET.sheetTitles == ('Sheet1', 'New Sheet 2', 'New Sheet 1')

    newSheet1.index = 1
    assert newSheet1.index == 1
    assert FIXED_SPREADSHEET[1] == newSheet1
    assert FIXED_SPREADSHEET[1].title == 'New Sheet 1'
    assert FIXED_SPREADSHEET.sheetTitles == ('Sheet1', 'New Sheet 1', 'New Sheet 2')

    # Test no change to index:
    newSheet1.index = 1
    assert newSheet1.index == 1
    assert FIXED_SPREADSHEET[1] == newSheet1
    assert FIXED_SPREADSHEET[1].title == 'New Sheet 1'
    assert FIXED_SPREADSHEET.sheetTitles == ('Sheet1', 'New Sheet 1', 'New Sheet 2')

    # Test negative index:
    newSheet1.index = -1
    assert newSheet1.index == 2
    assert FIXED_SPREADSHEET[2] == newSheet1
    assert FIXED_SPREADSHEET[2].title == 'New Sheet 1'
    assert FIXED_SPREADSHEET.sheetTitles == ('Sheet1', 'New Sheet 2', 'New Sheet 1')

    # Check setting to invalid index:
    with pytest.raises(TypeError):
        FIXED_SPREADSHEET[0].index = 1.0
    with pytest.raises(IndexError):
        FIXED_SPREADSHEET[0].index = 9999
    with pytest.raises(IndexError):
        FIXED_SPREADSHEET[0].index = -9999

    newSheet1.delete()
    newSheet2.delete()


def test_iter(init, checkPreAndPostCondition):
    FIXED_SPREADSHEET.addSheet(title='New Sheet 1')
    FIXED_SPREADSHEET.addSheet(title='New Sheet 2')

    for i, sheet in enumerate(FIXED_SPREADSHEET):
        if i == 0:
            assert sheet.title == 'Sheet1'
        elif i == 1:
            assert sheet.title == 'New Sheet 1'
        elif i == 2:
            assert sheet.title == 'New Sheet 2'

    del FIXED_SPREADSHEET[1] # delete New Sheet 1
    del FIXED_SPREADSHEET[1] # delete New Sheet 2
    checkIfSpreadsheetInOriginalState()


def test_str_spreadsheet(init, checkPreAndPostCondition):
    assert str(FIXED_SPREADSHEET) == '<Spreadsheet title="Delete Me", 1 sheets>'


def test_repr_spreadsheet(init, checkPreAndPostCondition):
    assert repr(FIXED_SPREADSHEET) == 'Spreadsheet(spreadsheetId=%r)' % (FIXED_SPREADSHEET.spreadsheetId)


def test_title_spreadsheet(init, checkPreAndPostCondition):
    assert FIXED_SPREADSHEET.title == 'Delete Me'
    FIXED_SPREADSHEET.title = 'New Title'
    assert FIXED_SPREADSHEET.title == 'New Title'
    FIXED_SPREADSHEET.refresh()
    assert FIXED_SPREADSHEET.title == 'New Title'
    FIXED_SPREADSHEET.title = 'Delete Me'
    checkIfSpreadsheetInOriginalState()


def test_title_sheet(init, checkPreAndPostCondition):
    assert FIXED_SPREADSHEET[0].title == 'Sheet1'
    FIXED_SPREADSHEET[0].title = 'New Title'
    assert FIXED_SPREADSHEET[0].title == 'New Title'
    FIXED_SPREADSHEET.refresh()
    assert FIXED_SPREADSHEET[0].title == 'New Title'
    FIXED_SPREADSHEET[0].title = 'Sheet1'
    checkIfSpreadsheetInOriginalState()


def test_spreadsheet_attr(init, checkPreAndPostCondition):
    assert FIXED_SPREADSHEET[0].spreadsheet == FIXED_SPREADSHEET


def test_tabColor(init, checkPreAndPostCondition):
    newSheet = FIXED_SPREADSHEET.addSheet(title='New Sheet 1')
    newSheet.tabColor = 'red'
    assert newSheet.tabColor == {'red': 1.0, 'green': 0.0, 'blue': 0.0, 'alpha': 1.0}
    FIXED_SPREADSHEET.refresh()
    assert newSheet.tabColor == {'red': 1.0, 'green': 0.0, 'blue': 0.0, 'alpha': 1.0}

    newSheet.tabColor = {'red': 0.0, 'green': 1.0, 'blue': 0.0, 'alpha': 1.0}
    assert newSheet.tabColor == {'red': 0.0, 'green': 1.0, 'blue': 0.0, 'alpha': 1.0}

    newSheet.delete()


def test_eq(init, checkPreAndPostCondition):
    assert FIXED_SPREADSHEET[0] == FIXED_SPREADSHEET[0]
    assert FIXED_SPREADSHEET[0] != 'some misc value'


def test_sheet_attrs(init, checkPreAndPostCondition):
    sheet1 = FIXED_SPREADSHEET[0]
    assert sheet1.rowCount == 1000
    assert sheet1.columnCount == 26
    assert sheet1.frozenRowCount == 0
    assert sheet1.frozenColumnCount == 0
    assert sheet1.hideGridlines == False
    assert sheet1.rowGroupControlAfter == False
    assert sheet1.columnGroupControlAfter == False


def test_str_sheet(init, checkPreAndPostCondition):
    assert str(FIXED_SPREADSHEET[0]) == '<Sheet title=\'Sheet1\', sheetId=%r, rowCount=1000, columnCount=26>' % (FIXED_SPREADSHEET[0].sheetId)


def test_repr_sheet(init, checkPreAndPostCondition):
    assert repr(FIXED_SPREADSHEET[0]) == "Sheet(sheetId=0, title='Sheet1', rowCount=1000, columnCount=26)"


def test_update_and_get(init, checkPreAndPostCondition):
    newSheet = FIXED_SPREADSHEET.addSheet(title='New Sheet 1')

    # Update a cell with column/row coordinates:
    newSheet.update(1, 1, 'new value')
    assert newSheet.get(1, 1) == 'new value'

    # Update a cell with A1 coordinates:
    newSheet.update('B2', 'b2 value')
    assert newSheet.get('B2') == 'b2 value'

    # Update a column:
    newSheet.updateColumn(1, ['a', 'b', 'c'])
    assert newSheet.getColumn(1) == ['a', 'b', 'c']

    # Update a column:
    newSheet.updateColumn('C', ['x', 'y', 'z'])
    assert newSheet.getColumn('C') == ['x', 'y', 'z']

    # Update a row:
    newSheet.updateRow(1, ['d', 'e', 'f'])
    assert newSheet.getRow(1) == ['d', 'e', 'f']

    # Update all:
    newSheet.updateAll([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']])
    assert newSheet.getAll() == [['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']]

    # Get a cell that is outside the original range of the rowCount and columnCount
    assert newSheet.get(9999, 1) == ''
    assert newSheet.get(1, 9999) == ''
    assert newSheet.get(9999, 9999) == ''

    # TODO test enlarging the sheet with an update*() call.


    # Test with invalid coordinate.
    with pytest.raises(TypeError):
        newSheet.get(1, 2, 3, 4, 5) # Test with too many arguments.
    with pytest.raises(TypeError):
        newSheet.get() # Test with too few arguments.
    with pytest.raises(TypeError):
        newSheet.get(4.2, 1) # Test with wrong argument type.
    with pytest.raises(TypeError):
        newSheet.get(1, 4.2) # Test with wrong argument type.

    with pytest.raises(TypeError):
        newSheet.update(1, 1) # Test with missing update value.
    with pytest.raises(TypeError):
        newSheet.update(1, 2, 3, 4, 5) # Test with too many arguments.
    with pytest.raises(TypeError):
        newSheet.update() # Test with too few arguments.
    with pytest.raises(TypeError):
        newSheet.update(4.2, 1) # Test with wrong argument type.
    with pytest.raises(TypeError):
        newSheet.update(1, 4.2) # Test with wrong argument type.

    # Test with bad indexes
    with pytest.raises(IndexError):
        newSheet.get(-1, 1) # Test with wrong argument type.
    with pytest.raises(IndexError):
        newSheet.get(1, -1) # Test with wrong argument type.

    with pytest.raises(IndexError):
        newSheet.update(-1, 1, 'value') # Test with wrong argument type.
    with pytest.raises(IndexError):
        newSheet.update(1, -1, 'value') # Test with wrong argument type.


    newSheet.delete()

if __name__ == '__main__':
    pytest.main()
