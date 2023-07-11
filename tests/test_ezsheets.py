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
    # Check if there exists only one Sheet in the TEST_SS Spreadsheet object.
    assert TEST_SS.title == 'Delete Me'
    assert len(TEST_SS) == 1
    assert TEST_SS[0].title == 'Sheet1'
    #print('READS=%s, WRITES=%s' % (ezsheets._READ_REQUESTS, ezsheets._WRITE_REQUESTS))


def addOriginalSheet():
    # If we delete all the sheets in TEST_SS, call this to restore the original Sheet1 Sheet object.
    assert 'Sheet1' not in TEST_SS.sheetTitles
    TEST_SS.createSheet(title='Sheet1')
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
    # Fake url and id
    assert ezsheets.getIdFromUrl("https://docs.google.com/spreadsheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0") == "10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng"

    with pytest.raises(ValueError):
        # Fake url and id
        ezsheets.getIdFromUrl(r"https://docs.google.com/spread           sheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0")


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


def test_getColumnNumberOf():
    assert ezsheets.getColumnNumberOf('A') == 1
    assert ezsheets.getColumnNumberOf('Z') == 26
    assert ezsheets.getColumnNumberOf('AA') == 27
    assert ezsheets.getColumnNumberOf('ZZ') == 702

    with pytest.raises(TypeError):
        ezsheets.getColumnNumberOf(1)

    with pytest.raises(ValueError):
        ezsheets.getColumnNumberOf('123')


def test_columnNumberLetterTranslation():
    for i in range(1, 5000):
        assert ezsheets.getColumnNumberOf(ezsheets.getColumnLetterOf(i)) == i


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
    global TEST_SS
    ezsheets.init()

    # Create a new spreadsheet
    TEST_SS = ezsheets.createSpreadsheet(title='Delete Me')
    # Use an existing spreadsheet:
    #TEST_SS = ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/1lRyPHuaLIgqYwkCTJYexbZUO1dcWeunm69B0L7L4ZQ8/edit#gid=0')

    tempName = 'temp_%s' % (random.sample('abcdefg'*10, 10))
    TEST_SS.createSheet(tempName, index=0)
    while len(TEST_SS) > 1:
        del TEST_SS[1]
    TEST_SS.createSheet('Sheet1', index=0)
    del TEST_SS[tempName]

    checkIfSpreadsheetInOriginalState()


def test_Spreadsheet_url(init):
    assert TEST_SS.url == 'https://docs.google.com/spreadsheets/d/' + TEST_SS.spreadsheetId + '/'


def test_Spreadsheet_attr(init, checkPreAndPostCondition):
    assert TEST_SS.title == 'Delete Me'
    assert TEST_SS.spreadsheetId != ''
    assert TEST_SS.sheetTitles == ('Sheet1',)
    assert len(TEST_SS) == 1

    assert TEST_SS.spreadsheetId == TEST_SS.id


def test_createSheet_deleteSheet(init, checkPreAndPostCondition):
    newSheet1 = TEST_SS.createSheet(title='New Sheet 1')

    assert 'New Sheet 1' in TEST_SS.sheetTitles
    assert newSheet1 == TEST_SS.sheets[1]
    assert newSheet1.title == 'New Sheet 1'
    assert newSheet1.index == 1

    newSheet2 = TEST_SS.createSheet(title='New Sheet 2', index=1)
    assert newSheet2 == TEST_SS.sheets[1]
    assert newSheet2.index == 1
    assert newSheet1 == TEST_SS.sheets[2]
    assert newSheet1.index == 2

    newSheet1.delete()
    newSheet2.delete()


def test_getitem_delitem(init, checkPreAndPostCondition):
    assert TEST_SS['Sheet1'].title == 'Sheet1'

    # Attempt get with invalid index:
    with pytest.raises(KeyError):
        TEST_SS[99]
    with pytest.raises(KeyError):
        TEST_SS[-99]
    with pytest.raises(KeyError):
        TEST_SS['nonexistent title']
    with pytest.raises(KeyError):
        TEST_SS[['invalid', 'key', 'type']]

    # Attempt delete with invalid index:
    with pytest.raises(KeyError):
        del TEST_SS[99]
    with pytest.raises(KeyError):
        del TEST_SS[-99]
    with pytest.raises(KeyError):
        del TEST_SS['nonexistent title']
    with pytest.raises(TypeError):
        del TEST_SS[['invalid', 'key', 'type']]

    newSheet = TEST_SS.createSheet(title='Added Sheet')
    assert TEST_SS[1] == newSheet # Get by int index.
    assert TEST_SS[-1] == newSheet # Get by negative index.
    assert TEST_SS[1].title == 'Added Sheet'

    assert TEST_SS[1:2] == (newSheet,) # Get by slice

    # Delete by int index:
    del TEST_SS[1]
    checkIfSpreadsheetInOriginalState()

    TEST_SS.createSheet(title='Added Sheet 2')
    # Get by str title:
    assert TEST_SS['Added Sheet 2'].title == 'Added Sheet 2'

    # Delete by str title:
    del TEST_SS['Added Sheet 2']
    checkIfSpreadsheetInOriginalState()

    # Get multiple sheets with slice:
    newSheet3 = TEST_SS.createSheet(title='Added Sheet 3')
    newSheet4 = TEST_SS.createSheet(title='Added Sheet 4')
    assert TEST_SS[1:3] == (newSheet3, newSheet4)

    # Delete multiple sheets with slice:
    del TEST_SS[1:3] # deleting newSheet3 and newSheet4
    checkIfSpreadsheetInOriginalState()

    TEST_SS.createSheet(title='Added Sheet 5')
    TEST_SS.createSheet(title='Added Sheet 6')
    del TEST_SS[1:] # deleting newSheet5 and newSheet6
    checkIfSpreadsheetInOriginalState()

    TEST_SS.createSheet(title='Added Sheet 7', index=0)
    TEST_SS.createSheet(title='Added Sheet 8', index=1)
    del TEST_SS[:2]
    checkIfSpreadsheetInOriginalState()

    TEST_SS.createSheet(title='Added Sheet 9')
    TEST_SS.createSheet(title='Added Sheet 10')
    del TEST_SS[3:0:-1]
    checkIfSpreadsheetInOriginalState()

    # Attempt to delete all sheets:
    with pytest.raises(ValueError):
        del TEST_SS[0]

    # Attempt to delete all sheets:
    with pytest.raises(ValueError):
        del TEST_SS[0:1]

    # Deleting with negative start or stop in slice is a no-op
    del TEST_SS[-1:]
    del TEST_SS[:-1]


def test_len(init, checkPreAndPostCondition):
    # The length of a Spreadsheet object is how many Sheet objects it contains.
    assert len(TEST_SS) == len(TEST_SS.sheets)
    TEST_SS.createSheet(title='Length Test Sheet')
    assert len(TEST_SS) == len(TEST_SS.sheets)
    del TEST_SS[-1]
    assert len(TEST_SS) == len(TEST_SS.sheets)


def test_changeSheetIndex(init, checkPreAndPostCondition):
    newSheet1 = TEST_SS.createSheet(title='New Sheet 1')
    newSheet2 = TEST_SS.createSheet(title='New Sheet 2')

    assert TEST_SS.sheetTitles == ('Sheet1', 'New Sheet 1', 'New Sheet 2')

    # Move 'New Sheet 2' to front:

    newSheet2.index = 0  # LEFT OFF with error happening here, some quota issue? Am I just making too many calls with this test suite?
    assert newSheet2.index == 0
    assert TEST_SS[0] == newSheet2
    assert TEST_SS.sheetTitles == ('New Sheet 2', 'Sheet1', 'New Sheet 1')

    newSheet2.index = 1
    assert newSheet2.index == 1
    assert TEST_SS[1] == newSheet2
    assert TEST_SS[1].title == 'New Sheet 2'
    assert TEST_SS.sheetTitles == ('Sheet1', 'New Sheet 2', 'New Sheet 1')

    newSheet1.index = 1
    assert newSheet1.index == 1
    assert TEST_SS[1] == newSheet1
    assert TEST_SS[1].title == 'New Sheet 1'
    assert TEST_SS.sheetTitles == ('Sheet1', 'New Sheet 1', 'New Sheet 2')

    # Test no change to index:
    newSheet1.index = 1
    assert newSheet1.index == 1
    assert TEST_SS[1] == newSheet1
    assert TEST_SS[1].title == 'New Sheet 1'
    assert TEST_SS.sheetTitles == ('Sheet1', 'New Sheet 1', 'New Sheet 2')

    # Test negative index:
    newSheet1.index = -1
    assert newSheet1.index == 2
    assert TEST_SS[2] == newSheet1
    assert TEST_SS[2].title == 'New Sheet 1'
    assert TEST_SS.sheetTitles == ('Sheet1', 'New Sheet 2', 'New Sheet 1')

    # Check setting to invalid index:
    with pytest.raises(TypeError):
        TEST_SS[0].index = 1.0
    with pytest.raises(IndexError):
        TEST_SS[0].index = 9999
    with pytest.raises(IndexError):
        TEST_SS[0].index = -9999

    newSheet1.delete()
    newSheet2.delete()


def test_iter(init, checkPreAndPostCondition):
    TEST_SS.createSheet(title='New Sheet 1')
    TEST_SS.createSheet(title='New Sheet 2')

    for i, sheet in enumerate(TEST_SS):
        if i == 0:
            assert sheet.title == 'Sheet1'
        elif i == 1:
            assert sheet.title == 'New Sheet 1'
        elif i == 2:
            assert sheet.title == 'New Sheet 2'

    del TEST_SS[1] # delete New Sheet 1
    del TEST_SS[1] # delete New Sheet 2
    checkIfSpreadsheetInOriginalState()


def test_str_spreadsheet(init, checkPreAndPostCondition):
    assert str(TEST_SS) == '<Spreadsheet title="Delete Me", 1 sheets>'


def test_repr_spreadsheet(init, checkPreAndPostCondition):
    assert repr(TEST_SS) == 'Spreadsheet(spreadsheetId=%r)' % (TEST_SS.spreadsheetId)


def test_title_spreadsheet(init, checkPreAndPostCondition):
    assert TEST_SS.title == 'Delete Me'
    TEST_SS.title = 'New Title'
    assert TEST_SS.title == 'New Title'
    TEST_SS.refresh()
    assert TEST_SS.title == 'New Title'
    TEST_SS.title = 'Delete Me'
    checkIfSpreadsheetInOriginalState()


def test_title_sheet(init, checkPreAndPostCondition):
    assert TEST_SS[0].title == 'Sheet1'
    TEST_SS[0].title = 'New Title'
    assert TEST_SS[0].title == 'New Title'
    TEST_SS.refresh()
    assert TEST_SS[0].title == 'New Title'
    TEST_SS[0].title = 'Sheet1'
    checkIfSpreadsheetInOriginalState()


def test_spreadsheet_attr(init, checkPreAndPostCondition):
    assert TEST_SS[0].spreadsheet == TEST_SS


def test_tabColor(init, checkPreAndPostCondition):
    newSheet = TEST_SS.createSheet(title='New Sheet 1')
    newSheet.tabColor = 'red'
    assert newSheet.tabColor == {'red': 1.0, 'green': 0.0, 'blue': 0.0, 'alpha': 1.0}
    TEST_SS.refresh()
    assert newSheet.tabColor == {'red': 1.0, 'green': 0.0, 'blue': 0.0, 'alpha': 1.0}

    newSheet.tabColor = {'red': 0.0, 'green': 1.0, 'blue': 0.0, 'alpha': 1.0}
    assert newSheet.tabColor == {'red': 0.0, 'green': 1.0, 'blue': 0.0, 'alpha': 1.0}

    newSheet.delete()


def test_eq(init, checkPreAndPostCondition):
    assert TEST_SS[0] == TEST_SS[0]
    assert TEST_SS[0] != 'some misc value'


def test_sheet_attrs(init, checkPreAndPostCondition):
    sheet1 = TEST_SS[0]
    assert sheet1.rowCount == 1000
    assert sheet1.columnCount == 26
    assert sheet1.frozenRowCount == 0
    assert sheet1.frozenColumnCount == 0
    assert sheet1.hideGridlines == False
    assert sheet1.rowGroupControlAfter == False
    assert sheet1.columnGroupControlAfter == False


def test_str_sheet(init, checkPreAndPostCondition):
    assert str(TEST_SS[0]) == '<Sheet title=\'Sheet1\', sheetId=%r, rowCount=1000, columnCount=26>' % (TEST_SS[0].sheetId)


def test_repr_sheet(init, checkPreAndPostCondition):
    assert repr(TEST_SS[0]) == "<Sheet sheetId=%r, title='Sheet1', rowCount=1000, columnCount=26>" % (TEST_SS[0].sheetId)


def test_updateRows(init, checkPreAndPostCondition):
    newSheet = TEST_SS.createSheet(title='New Sheet 1', columnCount=5, rowCount=4)

    newSheet.updateRows([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']])
    assert newSheet.getRows() == [['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', ''], ['', '', '', '', '']]
    assert newSheet.getRows(startRow=2) == [['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', ''], ['', '', '', '', '']]
    assert newSheet.getRows(stopRow=4) == [['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', '']]
    assert newSheet.getRows(startRow=2, stopRow=4) == [['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', '']]

    newSheet.clear()
    newSheet.updateRows([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']], startRow=2)
    assert newSheet.getRows() == [['', '', '', '', ''], ['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', '']]
    assert newSheet.getRows(startRow=2) == [['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', '']]
    assert newSheet.getRows(stopRow=4) == [['', '', '', '', ''], ['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', '']]
    assert newSheet.getRows(startRow=2, stopRow=4) == [['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', '']]

    newSheet.clear()
    newSheet.resize(columnCount=4, rowCount=5)

    newSheet.updateRows([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']])
    assert newSheet.getRows() == [['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', ''], ['', '', '', '']]
    assert newSheet.getRows(startRow=2) == [['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', ''], ['', '', '', '']]
    assert newSheet.getRows(stopRow=4) == [['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', '']]
    assert newSheet.getRows(startRow=2, stopRow=4) == [['d', 'e', 'f', ''], ['g', 'h', 'i', '']]

    newSheet.clear()
    newSheet.updateRows([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']], startRow=2)
    assert newSheet.getRows() == [['', '', '', ''], ['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', '']]
    assert newSheet.getRows(startRow=2) == [['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', '']]
    assert newSheet.getRows(stopRow=4) == [['', '', '', ''], ['a', 'b', 'c', ''], ['d', 'e', 'f', '']]
    assert newSheet.getRows(startRow=2, stopRow=4) == [['a', 'b', 'c', ''], ['d', 'e', 'f', '']]

    # Test invalid argumnets:
    with pytest.raises(TypeError):
        newSheet.updateRows('not a list or tuple')
    with pytest.raises(TypeError):
        newSheet.updateRows(['inner value not a list or tuple'])
    with pytest.raises(TypeError):
        newSheet.updateRows(startRow='invalid arg')
    with pytest.raises(TypeError):
        newSheet.updateRows(startRow=0)
    with pytest.raises(TypeError):
        newSheet.updateRows(startRow=9999)

    with pytest.raises(TypeError):
        newSheet.getRows(startRow='invalid arg')
    with pytest.raises(TypeError):
        newSheet.getRows(stopRow='invalid arg')
    with pytest.raises(ValueError):
        newSheet.getRows(startRow=0)
    with pytest.raises(ValueError):
        newSheet.getRows(startRow=-9999)
    with pytest.raises(ValueError):
        newSheet.getRows(stopRow=0)

    newSheet.delete()



def test_updateColumns(init, checkPreAndPostCondition):
    newSheet = TEST_SS.createSheet(title='New Sheet 1', columnCount=4, rowCount=5)

    newSheet.updateColumns([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']])
    assert newSheet.getColumns() == [['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', ''], ['', '', '', '', '']]
    assert newSheet.getColumns(startColumn=2) == [['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', ''], ['', '', '', '', '']]
    assert newSheet.getColumns(stopColumn=4) == [['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', '']]
    assert newSheet.getColumns(startColumn=2, stopColumn=4) == [['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', '']]

    newSheet.clear()
    newSheet.updateColumns([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']], startColumn=2)
    assert newSheet.getColumns() == [['', '', '', '', ''], ['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', '']]
    assert newSheet.getColumns(startColumn=2) == [['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', ''], ['g', 'h', 'i', '', '']]
    assert newSheet.getColumns(stopColumn=4) == [['', '', '', '', ''], ['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', '']]
    assert newSheet.getColumns(startColumn=2, stopColumn=4) == [['a', 'b', 'c', '', ''], ['d', 'e', 'f', '', '']]

    newSheet.clear()
    newSheet.resize(columnCount=5, rowCount=4)

    newSheet.updateColumns([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']])
    assert newSheet.getColumns() == [['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', ''], ['', '', '', '']]
    assert newSheet.getColumns(startColumn=2) == [['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', ''], ['', '', '', '']]
    assert newSheet.getColumns(stopColumn=4) == [['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', '']]
    assert newSheet.getColumns(startColumn=2, stopColumn=4) == [['d', 'e', 'f', ''], ['g', 'h', 'i', '']]

    newSheet.clear()
    newSheet.updateColumns([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']], startColumn=2)
    assert newSheet.getColumns() == [['', '', '', ''], ['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', '']]
    assert newSheet.getColumns(startColumn=2) == [['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', '']]
    assert newSheet.getColumns(stopColumn=4) == [['', '', '', ''], ['a', 'b', 'c', ''], ['d', 'e', 'f', '']]
    assert newSheet.getColumns(startColumn=2, stopColumn=4) == [['a', 'b', 'c', ''], ['d', 'e', 'f', '']]

    # Test invalid argumnets:
    with pytest.raises(TypeError):
        newSheet.updateColumns('not a list or tuple')
    with pytest.raises(TypeError):
        newSheet.updateColumns(['inner value not a list or tuple'])
    with pytest.raises(TypeError):
        newSheet.updateColumns(startRow='invalid arg')
    with pytest.raises(TypeError):
        newSheet.updateColumns(startRow=0)
    with pytest.raises(TypeError):
        newSheet.updateColumns(startRow=9999)

    with pytest.raises(TypeError):
        newSheet.getColumns(startColumn='invalid arg')
    with pytest.raises(TypeError):
        newSheet.getColumns(stopColumn='invalid arg')
    with pytest.raises(ValueError):
        newSheet.getColumns(startColumn=0)
    with pytest.raises(ValueError):
        newSheet.getColumns(startColumn=-9999)
    with pytest.raises(ValueError):
        newSheet.getColumns(stopColumn=0)

    newSheet.delete()


def test_update_and_get(init, checkPreAndPostCondition):
    newSheet = TEST_SS.createSheet(title='New Sheet 1', columnCount=4, rowCount=4)

    # Update a cell with column/row coordinates:
    newSheet.update(1, 1, 'new value')
    assert newSheet.get(1, 1) == 'new value'

    # Update a cell with A1 coordinates:
    newSheet.update('B2', 'b2 value')
    assert newSheet.get('B2') == 'b2 value'

    # Update a column:
    newSheet.updateColumn(1, ['a', 'b', 'c'])
    assert newSheet.getColumn(1) == ['a', 'b', 'c', '']
    newSheet.updateColumn(1, ('A', 'B', 'C')) # send tuple instead of list
    assert newSheet.getColumn(1) == ['A', 'B', 'C', '']

    # Update a column:
    newSheet.updateColumn('C', ['x', 'y', 'z'])
    assert newSheet.getColumn('C') == ['x', 'y', 'z', '']

    # Update a row:
    newSheet.updateRow(1, ['d', 'e', 'f'])
    assert newSheet.getRow(1) == ['d', 'e', 'f', '']
    newSheet.updateRow(1, ('D', 'E', 'F')) # send tuple instead of list
    assert newSheet.getRow(1) == ['D', 'E', 'F', '']

    # Update all:
    newSheet.updateRows([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']])
    assert newSheet.getRows() == [['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', '']]

    newSheet.updateColumns([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']])
    assert newSheet.getColumns() == [['a', 'b', 'c', ''], ['d', 'e', 'f', ''], ['g', 'h', 'i', ''], ['', '', '', '']]

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
    with pytest.raises(TypeError):
        newSheet.update('invalid', 1, 'value') # Test with wrong argument type.
    with pytest.raises(TypeError):
        newSheet.update(1, 'invalid', 'value') # Test with wrong argument type.


    # Test getRow() and getColumn() with invalid arguments.
    with pytest.raises(TypeError):
        newSheet.getRow('invalid arg')
    with pytest.raises(IndexError):
        newSheet.getRow(0)
    with pytest.raises(IndexError):
        newSheet.getRow(-9999)
    with pytest.raises(TypeError):
        newSheet.getColumn({})
    with pytest.raises(ValueError):
        newSheet.getColumn('AAAAinvalid arg')
    with pytest.raises(IndexError):
        newSheet.getColumn(0)
    with pytest.raises(IndexError):
        newSheet.getColumn(-9999)

    newSheet.delete()


def test_sheet_getitem_setitem(init, checkPreAndPostCondition):
    newSheet = TEST_SS.createSheet(title='New Sheet 1', columnCount=4, rowCount=4)

    # Update a cell with column/row coordinates:
    newSheet[1, 1] = 'new value'
    assert newSheet[1, 1] == 'new value'

    # Update a cell with A1 coordinates:
    newSheet['B2'] = 'b2 value'
    assert newSheet['B2'] == 'b2 value'

    # Get a cell that is outside the original range of the rowCount and columnCount
    assert newSheet[9999, 1] == ''
    assert newSheet[1, 9999] == ''
    assert newSheet[9999, 9999] == ''

    # TODO test enlarging the sheet with an update*() call.

    # Test with invalid coordinate.
    with pytest.raises(TypeError):
        newSheet[1, 2, 3, 4, 5] # Test with too many arguments.
    with pytest.raises(TypeError):
        newSheet[1] # Test with too few arguments.
    with pytest.raises(TypeError):
        newSheet[4.2, 1] # Test with wrong argument type.
    with pytest.raises(TypeError):
        newSheet[1, 4.2] # Test with wrong argument type.

    with pytest.raises(TypeError):
        newSheet[1, 2, 3, 4] = 5 # Test with too many arguments.
    with pytest.raises(TypeError):
        newSheet[4.2, 1] = 5 # Test with wrong argument type.
    with pytest.raises(TypeError):
        newSheet[1, 4.2] = 5 # Test with wrong argument type.

    # Test with bad indexes
    with pytest.raises(IndexError):
        newSheet[-1, 1] # Test with wrong argument type.
    with pytest.raises(IndexError):
        newSheet[1, -1] # Test with wrong argument type.

    with pytest.raises(IndexError):
        newSheet[-1, 1] = 'value' # Test with wrong argument type.
    with pytest.raises(IndexError):
        newSheet[1, -1] = 'value' # Test with wrong argument type.
    with pytest.raises(TypeError):
        newSheet['invalid', 1] = 'value' # Test with wrong argument type.
    with pytest.raises(TypeError):
        newSheet[1, 'invalid'] = 'value' # Test with wrong argument type.

    newSheet.delete()


def test_gridProperties_settersGetters(init, checkPreAndPostCondition):
    # Test rowCount
    sheet = TEST_SS[0]

    # Test setters and getters:
    sheet.rowCount = 10
    assert sheet.rowCount == 10

    # Test invalid arguments:
    with pytest.raises(TypeError):
        sheet.rowCount = 'invalid arg'
    with pytest.raises(TypeError):
        sheet.rowCount = 0
    with pytest.raises(TypeError):
        sheet.rowCount = -9999

    # Test columnCount
    # Test setters and getters:
    sheet.columnCount = 10
    assert sheet.columnCount == 10

    # Test invalid arguments:
    with pytest.raises(TypeError):
        sheet.columnCount = 'invalid arg'
    with pytest.raises(TypeError):
        sheet.columnCount = 0
    with pytest.raises(TypeError):
        sheet.columnCount = -9999

    # Test frozenRowCount
    # Test setters and getters:
    sheet.frozenRowCount = 1
    assert sheet.frozenRowCount == 1

    # Test invalid arguments:
    with pytest.raises(TypeError):
        sheet.frozenRowCount = 'invalid arg'
    with pytest.raises(TypeError):
        sheet.frozenRowCount = 0
    with pytest.raises(TypeError):
        sheet.frozenRowCount = -9999

    # Test frozenColumnCount
    # Test setters and getters:
    sheet.frozenColumnCount = 1
    assert sheet.frozenColumnCount == 1

    # Test invalid arguments:
    with pytest.raises(TypeError):
        sheet.frozenColumnCount = 'invalid arg'
    with pytest.raises(TypeError):
        sheet.frozenColumnCount = 0
    with pytest.raises(TypeError):
        sheet.frozenColumnCount = -9999

    # Test freezing all rows and columns
    with pytest.raises(ValueError):
        sheet.rowCount = 1
    with pytest.raises(ValueError):
        sheet.columnCount = 1
    with pytest.raises(ValueError):
        sheet.frozenRowCount = 10
    with pytest.raises(ValueError):
        sheet.frozenColumnCount = 10

    # Test hideGridlines
    sheet.hideGridlines = True
    assert sheet.hideGridlines == True
    sheet.hideGridlines = False
    assert sheet.hideGridlines == False
    sheet.hideGridlines = 'truthy value'
    assert sheet.hideGridlines == True
    sheet.hideGridlines = '' # Falsey value
    assert sheet.hideGridlines == False

    # Test rowGroupControlAfter
    sheet.rowGroupControlAfter = True
    assert sheet.rowGroupControlAfter == True
    sheet.rowGroupControlAfter = False
    assert sheet.rowGroupControlAfter == False
    sheet.rowGroupControlAfter = 'truthy value'
    assert sheet.rowGroupControlAfter == True
    sheet.rowGroupControlAfter = '' # Falsey value
    assert sheet.rowGroupControlAfter == False

    # Test columnGroupControlAfter
    sheet.columnGroupControlAfter = True
    assert sheet.columnGroupControlAfter == True
    sheet.columnGroupControlAfter = False
    assert sheet.columnGroupControlAfter == False
    sheet.columnGroupControlAfter = 'truthy value'
    assert sheet.columnGroupControlAfter == True
    sheet.columnGroupControlAfter = '' # Falsey value
    assert sheet.columnGroupControlAfter == False


def test_updateRow_updateColumn(init, checkPreAndPostCondition):
    sheet = TEST_SS[0]

    with pytest.raises(TypeError):
        sheet.updateRow('invalid', [1, 2, 3])
    with pytest.raises(IndexError):
        sheet.updateRow(0, [1, 2, 3])
    with pytest.raises(TypeError):
        sheet.updateRow(1, 'invalid')

    with pytest.raises(TypeError):
        sheet.updateColumn(b'123', [1, 2, 3])
    with pytest.raises(IndexError):
        sheet.updateColumn(0, [1, 2, 3])
    with pytest.raises(TypeError):
        sheet.updateColumn(1, 'invalid')
    with pytest.raises(ValueError):
        sheet.updateColumn('123', [1, 2, 3]) # Not a valid column string.


def test_resize(init, checkPreAndPostCondition):
    sheet = TEST_SS[0]

    # Test no-change resize
    sheet.resize()

    sheet.resize(rowCount=10)
    assert sheet.rowCount == 10
    sheet.resize(columnCount=10)
    assert sheet.columnCount == 10

    sheet.resize(columnCount='Z')
    assert sheet.columnCount == 26

    with pytest.raises(TypeError):
        sheet.resize(rowCount='invalid')
    with pytest.raises(TypeError):
        sheet.resize(columnCount=3.14)
    with pytest.raises(ValueError):
        sheet.resize(columnCount='invalid123')

    with pytest.raises(TypeError):
        sheet.resize(rowCount=0)
    with pytest.raises(TypeError):
        sheet.resize(columnCount=0)


def test_convertAddress():
    assert ezsheets.convertAddress('A2') == (1, 2)
    assert ezsheets.convertAddress('B1') == (2, 1)
    assert ezsheets.convertAddress(1, 2) == 'A2'
    assert ezsheets.convertAddress((1, 2)) == 'A2'
    assert ezsheets.convertAddress(2, 1) == 'B1'
    assert ezsheets.convertAddress((2, 1)) == 'B1'

    with pytest.raises(TypeError):
        ezsheets.convertAddress()
    with pytest.raises(ValueError):
        ezsheets.convertAddress('A')
    with pytest.raises(ValueError):
        ezsheets.convertAddress('1')
    with pytest.raises(TypeError):
        ezsheets.convertAddress(0, 1)
    with pytest.raises(TypeError):
        ezsheets.convertAddress(1, 0)

    for col in range(1, 1000):
        for row in range(1, 101):
            assert ezsheets.convertAddress(ezsheets.convertAddress(col, row)) == (col, row)





def test_getitem(init, checkPreAndPostCondition):
    pass # TODO LEFT OFF


def test_restricted_and_link_viewable():
    # Restricted spreadsheet:
    ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/16O3PLFhA8EUH7CdmSeNtxie03kitBfLoi_aOex93WIA/edit#gid=0')

    # Viewable by anyone with the link:
    ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/1WCKpKAgUjVJv2RM23dc4LHj9mOonK8VsvSXi_U4eqLk/edit#gid=0')


def test_eq():
    ss1 = ezsheets.Spreadsheet('https://autbor.com/examplegs')
    ss2 = ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/1TzOJxhNKr15tzdZxTqtQ3EmDP6em_elnbtmZIcyu8vI/')
    ss3 = ezsheets.Spreadsheet('1TzOJxhNKr15tzdZxTqtQ3EmDP6em_elnbtmZIcyu8vI')

    other_ss = ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/1WCKpKAgUjVJv2RM23dc4LHj9mOonK8VsvSXi_U4eqLk/edit#gid=0')

    assert ss1 == ss2 == ss3
    assert ss1 != other_ss

    assert ss1[0].id == ss1[0].sheetId




