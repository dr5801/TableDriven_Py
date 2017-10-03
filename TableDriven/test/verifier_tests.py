import unittest
from Modules import Verifiers

"""
* Tests the functionality of the Verifier modules for:
  1. is_digit
  2. is_minus
  3. is_plus
  4. is_period
"""
class VerifierTests(unittest.TestCase):
    def test_isDigitValid(self):
        for number in range(48, 57):
            character = chr(number)
            self.assertTrue(Verifiers.is_digit(character))

    def test_isDigitInvalid(self):
        invalid_ranges = [range(32, 47), range(58, 126)]
        for invalid_range in invalid_ranges:
            for number in invalid_range:
                character = chr(number)
                self.assertFalse(Verifiers.is_digit(character))

    def test_isMinusValid(self):
        self.assertTrue(Verifiers.is_minus("-"))

    def test_isMinusInvalid(self):
        invalid_ranges = [range(32, 44), range(46, 126)]
        for invalid_range in invalid_ranges:
            for number in invalid_range:
                character = chr(number)
                self.assertFalse(Verifiers.is_minus(character))

    def test_isDecimalValid(self):
        self.assertTrue(Verifiers.is_period('.'))

    def test_isDecimalInvalid(self):
        invalid_ranges = [range(32, 45), range(47, 126)]
        for invalid_range in invalid_ranges:
            for number in invalid_range:
                character = chr(number)
                self.assertFalse(Verifiers.is_period(character))

    def test_isPlusValid(self):
        self.assertTrue(Verifiers.is_plus('+'))

    def test_isPlusInvalid(self):
        invalid_ranges = [range(32, 42), range(44, 126)]
        for invalid_range in invalid_ranges:
            for number in invalid_range:
                character = chr(number)
                self.assertFalse(Verifiers.is_plus(character))