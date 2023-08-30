#!/usr/bin/python

from modules.spreadsheet import create_result_book, load_input_workbook
from modules.cell import get_numeric_value
from os.path import splitext
from re import split, search

__all__ = [
    "ask_file_paths",
    "ask_worksheet",
    "ask_api_key",
    "ask_model",
    "ask_system_prompt",
    "ask_prompt",
    "ask_inputs",
    "ask_input_start",
    "ask_output_row",
    "ask_output_column",
    "ask_skip",
    "ask_output_placement",
    "ask_until",
    "ask_limit"
]

# Constants

DEFAULT_SYSTEM_PROMPT = "You are a helpful assistant. You give only the answer, without forming a sentence. If you are not sure, try guessing. If you are still unsure about the answer, output '?'. If you don't know the answer, or if you cannot give a correct answer output '?'."

MODELS = ["gpt-3.5-turbo", "gpt-3.5-turbo-16k", "gpt-4", "gpt-4-32k"]

INPUTS_REGEX = "(^([0-9]+)([, ]([0-9]+))*$)|(^([a-zA-Z]+)([, ]([a-zA-Z]+))*$)"

# Functions

def ask_file_paths() -> tuple:
    file_path = ""
    result_path = ""
    workbook = None
    result_book = None

    while True:
        file_path = input("\nEnter the full or relative file path (.xlsx/.xlsm): ").strip()

        workbook = load_input_workbook(file_path)

        if workbook != None and workbook != False:
            break

    while True:
        result_path = input("\nEnter the full or relative output file path (.xlsx/.xlsm) (press ENTER to use default): ").strip()

        if result_path == "":
            pre, ext = splitext(file_path)
            result_path = pre + "_output" + ext

        result_book = create_result_book(file_path, result_path)

        if result_book != None and result_book != False:
            break

    return (file_path, result_path, workbook, result_book)

def ask_worksheet(workbook, result_book) -> tuple:
    if workbook == None:
        raise TypeError("workbook must not be NoneType")

    if result_book == None:
        raise TypeError("result_book must not be NoneType")

    if len(workbook.sheetnames) == 0:
        print("\nNo sheets found. Exiting...\n")
        exit(0)

    while True:
        print("\nSheets in the worksheet:\n")

        index = 1

        for sheet in workbook.sheetnames:
            print(str(index) + ") " + sheet)
            index = index + 1

        choice = input("\nChoose sheet (press ENTER to select the first sheet): ").strip()

        if choice == "":
            sheet_name = workbook.sheetnames[0]

            break
        elif choice.isnumeric() and int(choice) >= 1 and int(choice) <= len(workbook.sheetnames):
            sheet_name = workbook.sheetnames[int(choice) - 1]

            print("\nUsing the sheet with name: ")
            print(sheet_name)
            print("\n")

            break

        print("\nPlease choose a valid sheet")

    return(workbook[sheet_name], result_book[sheet_name])

def ask_api_key() -> str:
    api_key = ""

    while True:
        api_key = input("\nEnter your OpenAI API key: ").strip()

        if api_key != "": break

        print("\nPlease enter a valid API key\n")

    return api_key

def ask_model() -> str:
    global MODELS

    model = ""

    while True:
        choice = input("\nChoose a ChatGPT model:\n\n1) gpt-3.5-turbo\n2) gpt-3.5-turbo-16k (default)\n3) gpt-4\n4) gpt-4-32k\n\nChoice (1-4, ENTER for default): ").strip()

        if choice == "":
            model = MODELS[1]
            print("\nUsing the default model\n")
            break

        if choice.isnumeric() and int(choice) > 0 and int(choice) < 5:
            model = MODELS[int(choice) - 1]
            print("\nUsing the model: ")
            print(model)
            break

        print("\nPlease choose a valid model\n")

    return model

def ask_system_prompt() -> str:
    global DEFAULT_SYSTEM_PROMPT

    system_prompt = input("\nEnter a system prompt to fine-tune the response (press ENTER for default):\n\n").strip()

    if system_prompt == "":
        system_prompt = DEFAULT_SYSTEM_PROMPT

        print("\nUsing the default prompt:\n")
        print(system_prompt)
        print("\n")

    return system_prompt

def ask_prompt() -> str:
    prompt = ""

    while True:
        prompt = input("\nEnter a prompt to be used with each request. Use zero-based '$0' placeholder(s) to insert inputs (i.e. Assuming you have entered 'a,b' as the inputs, and '2' as the starting row: 'Something $0 and $0, $1' will be converted to 'Something <VALUE OF A2> and <VALUE OF A2>, <VALUE OF B2>')\n\n").strip()

        if prompt != "":
            break

        print("\nPlease enter a valid prompt\n")

    return prompt

def ask_inputs() -> object:
    global INPUTS_REGEX

    inputs = []

    while True:
        inputs = input("\nEnter one or more input columns (A-Z)/rows (1-9) separated by a space ('1 3') or a comma ('1,3')  character:\n\n").strip().strip(",")

        if inputs != "" and search(INPUTS_REGEX, inputs) != None:
            inputs = split(",| ", inputs.upper())
            break

        print("\nPlease enter a valid input. You must choose either row(s) (i.e. '2' or '4,2,0,6,9'), or column(s) (i.e. 'd' or 'p,r,n,d'). Not both (i.e. '1,a,c,4').\n")

    return inputs

def ask_input_start(inputs: object) -> str:
    if not hasattr(inputs, "__len__"):
        raise TypeError("inputs must be an array")

    input_pos = ""

    while True:
        question = "\nEnter the starting row for the inputs (i.e. '2'): "

        if inputs[0].isnumeric():
            question = "\nEnter the starting column for the inputs (i.e. 'C'): "

        choice = input(question).strip()

        if inputs[0].isnumeric() and choice.isalpha():
            input_pos = choice.upper()
            break

        if inputs[0].isalpha() and choice.isnumeric():
            input_pos = choice
            break

        if inputs[0].isnumeric():
            print("\nPlease enter a valid column (must be a character)\n")
        else:
            print("\nPlease enter a valid row (must be a number)\n")

    return input_pos

def ask_output_row() -> str:
    result_row = ""

    while True:
        result_row = input("\nEnter the starting row for the results (i.e. '2'): ").strip()

        if result_row != "" and result_row.isnumeric():
            result_row = result_row
            break

        print("\nPlease enter a valid row.\n")

    return result_row

def ask_output_column() -> str:
    result_col = ""

    while True:
        result_col = input("\nEnter the starting column for the results (i.e. 'B'): ").strip().upper()

        if result_col != "" and result_col.isalpha():
            break

        print("\nPlease enter a valid column.\n")

    return result_col

def ask_skip() -> bool:
    while True:
        choice = input("\nShould the existing results be skipped? Results will be overwritten if skipping is disabled. (y/n)  ").strip().lower()

        if choice == "y":
            return True
        elif choice == "n":
            return False

        print("\nPlease enter a valid choice.\n")

def ask_output_placement() -> str:
    choice = ""

    while True:
        choice = input("\nShould the result be output to the next row, or column? (type 'r' for next row, 'c' for next column) ").strip().lower()

        if choice == "r" or choice == "c": break

        print("\nPlease enter a valid choice ('r' or 'c')\n")

    return choice.lower()

def ask_until(input_pos: str) -> str:
    if not input_pos.isalpha() and not input_pos.isnumeric():
        raise Exception("input_pos must be either numeric or alphabetic. not both")

    question = "\nEnter the ending row for the inputs (inclusive, i.e. '2', press ENTER to continue until empty input): "
    until = ""

    if input_pos.isalpha():
        question = "\nEnter the ending column for the inputs (inclusive, i.e. 'C', press ENTER to continue until empty input): "

    while True:
        choice = input(question).strip().upper()

        if choice == "":
            return ""

        if input_pos.isnumeric() and choice.isnumeric() and get_numeric_value(input_pos) >= int(choice):
            print("\nEnding row cannot be equal or smaller than the starting row\n")
            continue

        if input_pos.isalpha() and choice.isalpha() and get_numeric_value(input_pos) >= int(choice):
            print("\nEnding column cannot be equal or before the starting column\n")
            continue

        if input_pos.isalpha() and choice.isalpha():
            until = choice
            break

        if input_pos.isnumeric() and choice.isnumeric():
            until = choice
            break

        if input_pos.isalpha():
            print("\nPlease enter a valid column (must be a character)\n")
        else:
            print("\nPlease enter a valid row (must be a number)\n")

    return until

def ask_limit() -> int:
    limit = 0

    while True:
        limit = input("\nEnter a limit for number of items to be processed by ChatGPT. Cached or existing results do not count towards the limit. (type 0 to remove limit): ").strip()

        if limit.isnumeric() and int(limit) >= 0:
            limit = int(limit)
            break

        print("\nPlease enter a valid limit.\n")

    return limit
