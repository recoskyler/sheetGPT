#!/usr/bin/python

from os import name, system
from modules.spreadsheet import create_result_book, load_workbook
from modules.cell import set_cell_value, get_cell_value, get_inputs, get_next_column, get_numeric_value
from modules.assistant import get_formatted_prompt, get_answer
from modules.io import save_progress, save_and_close
from modules.user_input import *

if __name__ != '__main__':
    print("Not supported as a module")
    exit()

# Constants

SAVE_INTERVAL = 25 # Save after every 25 processed items

# Variables

file_path = ""
result_path = ""
api_key = ""
model = ""
sheet_name = "Sheet 1"
system_prompt = ""
prompt = ""
result_row = 1
result_col = ""
inputs = ""
workbook = None
worksheet = None
result_book = None
result_sheet = None
result_placement = None
input_pos = ""
cache = dict()
limit = 0
processed = 0
until = ""
skip = True

# Functions

def get_user_inputs():
    global file_path
    global api_key
    global model
    global sheet_name
    global system_prompt
    global prompt
    global result_path
    global result_book
    global result_sheet
    global result_placement
    global result_row
    global result_col
    global inputs
    global worksheet
    global workbook
    global input_pos
    global limit
    global skip
    global until

    file_paths = ask_file_paths()

    file_path = file_paths[0]
    result_path = file_paths[1]
    workbook = file_paths[2]
    result_book = file_paths[3]

    print("\nInput file:  " + file_path)
    print("\nOutput file: " + result_path)

    sheets = ask_worksheet(workbook, result_book)

    worksheet = sheets[0]
    result_sheet = sheets[1]

    api_key = ask_api_key()
    model = ask_model()
    system_prompt = ask_system_prompt()
    prompt = ask_prompt()
    inputs = ask_inputs()
    input_pos = ask_input_start(inputs)
    result_col = ask_output_column()
    result_row = ask_output_row()

    print("\n\nStarting point for results: ")
    print(result_col + str(result_row))
    print("\n")

    result_placement = ask_output_placement()
    limit = ask_limit()
    skip = ask_skip()
    until = ask_until(input_pos)

def process_worksheet():
    global worksheet
    global result_sheet
    global input_pos
    global prompt
    global result_book
    global result_sheet
    global result_placement
    global result_row
    global result_col
    global inputs
    global limit
    global processed
    global skip
    global until

    item_no = 1
    processed = 0

    print("Processing... Press CTRL + C to cancel")

    if limit == 0:
        print("\nWARNING: No limit set!\n")
    else:
        print("\nLimiting to " + str(limit) + " items...\n")

    while limit == 0 or processed <= limit:
        if processed != 0 and processed % SAVE_INTERVAL == 0:
            save_progress(result_book, result_path)

        if until != "" and get_numeric_value(input_pos) > get_numeric_value(until): break

        print("Processing item " + str(processed) + "/" + str(item_no) + " | Position: " + input_pos + "...")

        input_values = get_inputs(worksheet, inputs, input_pos)

        if len(input_values) == 0: break

        current_prompt = get_formatted_prompt(prompt, input_values)
        result_value = get_cell_value(worksheet=result_sheet, cell=(result_col + result_row))

        if result_value == None or str(result_value).strip() == "" or not skip:
            print("Prompt: " + current_prompt)

            res = get_answer(current_prompt, system_prompt, model, api_key, cache)
            answer = res[0]
            processed_count = res[1]

            set_cell_value(worksheet=result_sheet, cell=(result_col + result_row), value=answer)

            print("Answer: " + answer)

            processed = processed + processed_count
        else:
            print("Already has result. Skipping...")

            prompt_hash = str(hash(current_prompt))

            if prompt_hash not in cache.keys():
                print("Caching answer...")
                cache[prompt_hash] = str(result_value).strip()

        if result_placement == "r": # Place next result on the next row
            result_row = str(int(result_row) + 1)
        else: # Place next result on the next column
            result_col = get_next_column(result_col)

        if input_pos.isnumeric():
            input_pos = str(int(input_pos) + 1)
        else:
            input_pos = get_next_column(input_pos)

        item_no = item_no + 1

    print("Done!")

def clear():
    if name == 'nt':
        _ = system('cls')
    else:
        _ = system('clear')

# Program

try:
    clear()

    print("\n\nSheetGPT v1.0.0 by recoskyler\n\n")

    get_user_inputs()
    process_worksheet()
except KeyboardInterrupt:
    print("\n\nCtrl+C detected. Exiting...\n\n")
except Exception as e:
    print("\n\nAn error occurred\n\n")
    print(e)
finally:
    save_and_close(workbook, result_book, result_path)

    print("\n======= Adios, Cowboy =======\n")
