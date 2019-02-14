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


def test_Spreadsheet_attr(init):
    assert FIXED_SPREADSHEET.title == 'Delete Me'
    assert FIXED_SPREADSHEET.spreadsheetId != ''
    assert FIXED_SPREADSHEET.sheetTitles == ('Sheet1',)
    assert len(FIXED_SPREADSHEET.sheets) == 1


def test_addSheet_deleteSheet(init):
    numOfSheetsBeforeAdding = len(FIXED_SPREADSHEET.sheets)
    FIXED_SPREADSHEET.addSheet(title='Added Sheet', rowCount=101, columnCount=13, tabColor='red')
    assert len(FIXED_SPREADSHEET.sheets) == numOfSheetsBeforeAdding + 1

    assert 'Added Sheet' in FIXED_SPREADSHEET.sheetTitles
    newSheet = FIXED_SPREADSHEET.sheets[-1]
    assert newSheet.title == 'Added Sheet'
    assert newSheet.index == numOfSheetsBeforeAdding
    assert newSheet.rowCount == 101
    assert newSheet.columnCount == 13
    assert newSheet.tabColor == {'red': 1.0, 'green': 0.0, 'blue': 0.0, 'alpha': 1.0}

    assert newSheet.delete()
    assert len(FIXED_SPREADSHEET.sheets) == numOfSheetsBeforeAdding
    assert 'Added Sheet' not in FIXED_SPREADSHEET.sheetTitles

if __name__ == '__main__':
    pytest.main()
