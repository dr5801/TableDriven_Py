from src.ConvertingMachine import ConvertingMachine
import unittest

"""
* Tests the functionality of ConvertingMachine
"""

class TestConvertingMachine(unittest.TestCase):
    def test_positiveWholeNumber(self):
        machine = ConvertingMachine()
        text = "99901"
        self.assertEqual(int(text), machine.parseText(text))

    def test_negativeWholeNumber(self):
        machine = ConvertingMachine()
        text = "-101010101"
        self.assertEqual(int(text), machine.parseText(text))

    def test_positiveRealNumber(self):
        machine = ConvertingMachine()
        text = "1.00000002"
        self.assertEqual(float(text), machine.parseText(text))

    def test_numberWithPositiveSign(self):
        machine = ConvertingMachine()
        text = "+20004.5"
        self.assertEqual(float(text), machine.parseText(text))

    def test_numberWithBadInput(self):
        machine = ConvertingMachine()
        text = "23abc4b5n6"
        self.assertEqual(23456, machine.parseText(text))