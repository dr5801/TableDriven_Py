from Modules import Actions
from Modules import Verifiers
from src.InterimResult import InterimResult
import attr


# create transition states
class State:
    START, INTEGER, DECIMAL, END = range(4)

# Represent an edge in a state machine
@attr.s
class Edge:
    current_state = attr.ib()
    input_verifier = attr.ib()
    action = attr.ib()
    next_state = attr.ib()

def generateList():
    machine = [
        Edge(current_state=State.START, input_verifier=Verifiers.is_digit, action=Actions.valueIsDigitAction, next_state=State.INTEGER),
        Edge(current_state=State.START, input_verifier=Verifiers.is_minus, action=Actions.negateAction, next_state=State.INTEGER),
        Edge(current_state=State.START, input_verifier=Verifiers.is_plus, action=Actions.noAction, next_state=State.INTEGER),
        Edge(current_state=State.START, input_verifier=Verifiers.is_period, action=Actions.startFraction, next_state=State.DECIMAL),
        Edge(current_state=State.INTEGER, input_verifier=Verifiers.is_digit, action=Actions.continuingIntegerAction, next_state=State.INTEGER),
        Edge(current_state=State.INTEGER, input_verifier=Verifiers.is_period, action=Actions.startFraction, next_state=State.DECIMAL),
        Edge(current_state=State.DECIMAL, input_verifier=Verifiers.is_digit, action=Actions.continuingFractionAction, next_state=State.DECIMAL)
    ]

    return machine

# convert the text to it's numerical value
@attr.s
class ConvertingMachine:
    machine = generateList()

    # parse the user input number
    def parseText(self, text):
        current_state = State.START
        interim_result = InterimResult(p=0, s=1, v=0)

        for character in list(text):
            next_edge = self.searchForEdge(current_state, character)

            if next_edge is not None:
                next_edge.action(interim_result, character)
                current_state = next_edge.next_state
            else:
                print("Error: You didn't type a real number!")

        final_result = interim_result.getS() * interim_result.getV()
        return final_result

    # search for the next edge in the machine
    def searchForEdge(self, current_state, character):
        next_edge = None

        for edge in self.machine:
            if edge.current_state == current_state and edge.input_verifier(character):
                next_edge = edge

        return next_edge
