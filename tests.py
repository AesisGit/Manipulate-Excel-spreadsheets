import unittest
from cleaningTelNum import *

class CleningTelNumbersTestCase(unittest.TestCase):

    def test_remove_first_space_from_tel(self):
        result = remove_first_space_from_tel(" 988123453")
        self.assertEqual(result, "988123453")


    def test_remove_plus_from_tel(self):
       result = remove_plus_from_tel("+99312132")
       self.assertEqual(result, "99312132")