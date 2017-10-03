from src.ConvertingMachine import ConvertingMachine

if __name__ == "__main__":
    number_input = input("Please enter a number : ")

    converter = ConvertingMachine()
    final_result = converter.parseText(number_input)
    print(final_result)