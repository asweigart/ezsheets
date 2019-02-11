import unittest
import __init__ as init # for helper functions
from __init__ import Spreadsheet
from __init__ import Sheet

class test_getIdFromUrl(unittest.TestCase):

	# Tests should change when behavior for bad urls passed in is decided.
	# If bad urls are allowed the 404 should be caught in refresh(),
	# Else more handling should be done to URL input.

	# Standard use case
	def test_IdFromUrl(self):
		self.assertEqual(init.getIdFromUrl(r"https://docs.google.com/spreadsheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0"), r"10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng")
		
		
	# URL with a space in it after https://docs.google.com/spreadsheets/d/
	def test_IdFromUrl1(self):
		self.assertEqual(init.getIdFromUrl(r"https://docs.google.com/spreadsheets/d/10tRbpHZYkfRecHyRHR jBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0"), r"10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng")
		
	
	# URL with a space in it before the id https://docs.google.com/spreadsheets/d/
	def test_IdFromUrl2(self):
		self.assertEqual(init.getIdFromUrl(r"https://docs.google.com/spread sheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0"), r"10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng")
		
		
	# Invalid spreadsheet URL
	def test_IdFromUrl3(self):
		self.assertEqual(init.getIdFromUrl(r"https://google.com"), r"bad url")
		
		
class test_getColumnLetterOf(unittest.TestCase):

	# Test from documentation in the code.
	def test_ColumnLetterOf(self):
		self.assertEqual(init.getColumnLetterOf(1), 'A')
		
	# Test from documentation in the code.
	def test_ColumnLetterOf1(self):
		self.assertEqual(init.getColumnLetterOf(27), 'AA')
	
	# Test when the column doesn't exist (Negative number).
	def test_ColumnLetterOf2(self):
		self.assertEqual(init.getColumnLetterOf(-1), '')
		
	# Test for accuracy when dealing with a large number of columns.
	def test_ColumnLetterOf3(self):
		self.assertEqual(init.getColumnLetterOf(702), 'ZZ')
	
	# If a letter is given as input, the letter should probably just be returned.
	def test_ColumnLetterOf4(self):
		self.assertEqual(init.getColumnLetterOf('A'), 'A')
		
		
class test_getColumnNumber(unittest.TestCase):
	
	# Test from documentation in the code.
	def test_ColumnNumber(self):
		self.assertEqual(init.getColumnNumber('A'), 1)
		
	# Test from documentation in the code.
	def test_ColumnNumber1(self):
		self.assertEqual(init.getColumnNumber('AA'), 27)
		
	# Test when column doesn't exist.
	def test_ColumnNumber2(self):
		self.assertEqual(init.getColumnNumber(''), 0)
		
	# Test for accuracy when dealing with large number of columns.
	def test_ColumnNumber3(self):
		self.assertEqual(init.getColumnNumber('ZZ'), 702)
		
	# If a number is given as input, the number should probably be returned.
	def test_ColumnNumber4(self):
		self.assertEqual(init.getColumnNumber(1), 1)
	
if __name__ == '__main__':
	unittest.main()
	
	
	
	
	
	
	
	
	
	
	
	