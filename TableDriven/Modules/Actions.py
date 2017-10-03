from src.InterimResult import InterimResult

# no action required
def noAction(interim_result, character):
    return None

# when the value is a digit, set the value to the digit
def valueIsDigitAction(interim_result, character):
    interim_result.setV(int(character))

# negative sign was the input, set the sign to -1
def negateAction(interim_result: InterimResult, character):
    interim_result.setS(-1)

# decimal was found, set the point to 0.1
def startFraction(interim_result: InterimResult, character):
    interim_result.setP(0.1)

# another digit was found, multiply by 10 to the interim result and add the new character to it
def continuingIntegerAction(interim_result: InterimResult, character):
    interim_result.setV(10 * interim_result.getV() + int(character))

# any digit after the decimal was found, calculate the value to v and set the point by dividing by 10
def continuingFractionAction(interim_result: InterimResult, character):
    interim_result.setV(interim_result.getV() + interim_result.getP() * int(character))
    interim_result.setP(interim_result.getP() / 10)