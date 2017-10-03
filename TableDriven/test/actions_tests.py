from src.InterimResult import InterimResult
from Modules import Actions
import unittest

"""
* Tests the functions in the Actions module for:
  1. noAction
  2. valueIsDigitAction
  3. negateAction
  4. start fraction
  5. continuingIntegerAction
  6. continuingFractionAction
"""
class ActionsTests(unittest.TestCase):

    def test_noAction(self):
        interim_result = InterimResult(p=0, s=1, v=0)
        self.assertIsNone(Actions.noAction(interim_result, "1"))
        self.assertEqual(0, interim_result.getP())
        self.assertEqual(1, interim_result.getS())
        self.assertEqual(0, interim_result.getV())

    def test_valueIsDigitAction(self):
        interim_result = InterimResult(p=0, s=1, v=0)
        self.assertIsNone(Actions.valueIsDigitAction(interim_result, "1"))
        self.assertEqual(0, interim_result.getP())
        self.assertEqual(1, interim_result.getS())
        self.assertEqual(1, interim_result.getV())

    def test_negateAction(self):
        interim_result = InterimResult(p=0, s=1, v=0)
        self.assertIsNone(Actions.negateAction(interim_result, "-"))
        self.assertEqual(0, interim_result.getP())
        self.assertEqual(-1, interim_result.getS())
        self.assertEqual(0, interim_result.getV())

    def test_startFraction(self):
        interim_result = InterimResult(p=0, s=1, v=0)
        self.assertIsNone(Actions.startFraction(interim_result, "."))
        self.assertEqual(.1, interim_result.getP())
        self.assertEqual(1, interim_result.getS())
        self.assertEqual(0, interim_result.getV())

    def test_continuingIntegerAction(self):
        interim_result = InterimResult(p=0, s=1, v=23)
        self.assertIsNone(Actions.continuingIntegerAction(interim_result, "5"))
        self.assertEqual(0, interim_result.getP())
        self.assertEqual(1, interim_result.getS())
        self.assertEqual(235, interim_result.getV())

    def test_continuingFractionAction(self):
        interim_result = InterimResult(p=.1, s=1, v=24)
        self.assertIsNone(Actions.continuingFractionAction(interim_result, "5"))
        self.assertEqual(.01, interim_result.getP())
        self.assertEqual(1, interim_result.getS())
        self.assertEqual(24.5, interim_result.getV())