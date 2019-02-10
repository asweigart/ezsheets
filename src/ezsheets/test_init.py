import unittest
import __init__ as init # for helper functions
from __init__ import Spreadsheet
from __init__ import Sheet

class test_thing(unittest.TestCase):

	# Tests should change when behavior for bad urls passed in is decided.
	# If bad urls are allowed the 404 should be caught in refresh(),
	# Else more filtering should be done to URL input.

	# Standard use case
	def test_url_getter(self):
		self.assertEqual(init.getIdFromUrl(r"https://docs.google.com/spreadsheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0"), r"10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng")
		
		
	# URL with a space in it after https://docs.google.com/spreadsheets/d/
	def test_url_getter1(self):
		self.assertEqual(init.getIdFromUrl(r"https://docs.google.com/spreadsheets/d/10tRbpHZYkfRecHyRHR jBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0"), r"10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng")
		
	
	# URL with a space in it before the id https://docs.google.com/spreadsheets/d/
	def test_url_getter2(self):
		self.assertEqual(init.getIdFromUrl(r"https://docs.google.com/spread sheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0"), r"10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng")
		
		
	# Invalid spreadsheet URL
	def test_url_getter3(self):
		self.assertEqual(init.getIdFromUrl(r"https://google.com"), r"bad url")
		
		
# class test_get_column_letter_of(unittest.TestCase):
	
if __name__ == '__main__':
	init.Spreadsheet(r"https://docs.google.com/spreadsheets/d/10tRbpHZYkfRecHyRHRjBLdQYoq5QWNBqZmH9tt4Tjng/edit#gid=0")
	# unittest.main()