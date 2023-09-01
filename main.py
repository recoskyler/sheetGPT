#!/usr/bin/python

from os import name, system
from os.path import exists, isfile
from modules.spreadsheet import create_output_book, load_input_workbook, generate_result_path
from modules.cell import set_cell_value, get_cell_value, get_inputs, get_next_column, get_numeric_value
from modules.assistant import get_formatted_prompt, get_answer
from modules.io import save_progress, save_and_close
from modules.user_input import *
from re import split, search

import flet as ft

if __name__ != '__main__':
    print("Not supported as a module")
    exit()

# Constants

SAVE_INTERVAL = 25 # Save after every 25 processed items
MODELS = ["gpt-3.5-turbo", "gpt-3.5-turbo-16k", "gpt-4", "gpt-4-32k"]
DEFAULT_SYSTEM_PROMPT = "You are a helpful assistant. You give only the answer, without forming a sentence. If you are not sure, try guessing. If you are still unsure about the answer, output '?'. If you don't know the answer, or if you cannot give a correct answer output '?'."
INPUTS_REGEX = "(^([0-9]+)([, ]([0-9]+))*$)|(^([a-zA-Z]+)([, ]([a-zA-Z]+))*$)"

# Variables

exit_app = False
processing = False

file_path = ""
result_path = ""
api_key = ""
model = MODELS[1]
sheet_name = "Sheet 1"
system_prompt = DEFAULT_SYSTEM_PROMPT
prompt = ""
result_row = ""
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
overwrite_output = False
process_task = None
ended = False
valid_fields = {
    "input_file": False,
    "output_file": False,
    "inputs": False,
    "input_pos": False,
    "output_row": False,
    "output_col": False,
    "api_key": False,
    "system_prompt": True,
    "until": True,
    "prompt": False,
    "limit": True,
    "sheet": False
}

# Functions

def reset_globals():
    global file_path
    global result_path
    global api_key
    global model
    global sheet_name
    global system_prompt
    global prompt
    global result_row
    global result_col
    global inputs
    global workbook
    global worksheet
    global result_book
    global result_sheet
    global result_placement
    global input_pos
    global cache
    global limit
    global processed
    global until
    global skip
    global overwrite_output
    global process_task
    global ended
    global valid_fields
    global MODELS
    global DEFAULT_SYSTEM_PROMPT

    file_path = ""
    result_path = ""
    api_key = ""
    model = MODELS[1]
    sheet_name = "Sheet 1"
    system_prompt = DEFAULT_SYSTEM_PROMPT
    prompt = ""
    result_row = ""
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
    overwrite_output = False
    process_task = None
    ended = False
    valid_fields = {
        "input_file": False,
        "output_file": False,
        "inputs": False,
        "input_pos": False,
        "output_row": False,
        "output_col": False,
        "api_key": False,
        "system_prompt": True,
        "until": True,
        "prompt": False,
        "limit": True,
        "sheet": False
    }

def clear():
    if name == 'nt':
        _ = system('cls')
    else:
        _ = system('clear')

# GUI

def main(page: ft.Page):
    global MODELS
    global DEFAULT_SYSTEM_PROMPT
    global model
    global skip

    def process_item(current_pos, current_col, current_row, input_values):
        global worksheet
        global result_sheet
        global prompt
        global result_sheet
        global result_placement
        global inputs
        global skip
        global until
        global ended
        global processed
        global exit_app

        if ended or exit_app: return False

        input_hash = str(hash(frozenset(input_values)))
        result_value = get_cell_value(worksheet=result_sheet, cell=(current_col + current_row))

        if result_value == None or str(result_value).strip() == "" or not skip:
            answer = ""
            processed_count = 0

            if input_hash in cache.keys():
                print("Using cached answer...")
                answer = cache[input_hash]
            else:
                current_prompt = get_formatted_prompt(prompt, input_values)

                print("Prompt: " + current_prompt)

                res = get_answer(current_prompt, system_prompt, model, api_key)
                answer = res[0]
                processed_count = res[1]

            set_cell_value(worksheet=result_sheet, cell=(current_col + current_row), value=answer)

            print("Answer: " + answer)

            processed = processed + processed_count
        else:
            print("Already has result. Skipping...")

            if input_hash not in cache.keys():
                print("Caching answer...")
                cache[input_hash] = str(result_value).strip()

        return True

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
        global skip
        global until
        global processed
        global processing
        global exit_app

        if processing or exit_app: return

        item_no = 1
        processed = 0
        current_pos = input_pos
        current_col = result_col
        current_row = result_row
        processing = True

        print("Processing... Press CTRL + C to cancel")

        if limit == 0:
            print("\nWARNING: No limit set!\n")
        else:
            print("\nLimiting to " + str(limit) + " items...\n")

        while limit == 0 or processed <= limit:
            if processed != 0 and processed % SAVE_INTERVAL == 0:
                save_progress(result_book, result_path)

            if until != "" and get_numeric_value(current_pos) > get_numeric_value(until): break

            input_values = get_inputs(worksheet, inputs, current_pos)

            if len(input_values) == 0: break

            print("Processing item " + str(processed) + "/" + str(item_no) + " | Position: " + current_pos + "...")

            process_message.value = "Processing item " + str(processed) + "/" + str(item_no) + " | Position: " + current_pos + "..."
            page.update()

            if not process_item(current_pos, current_col, current_row, input_values): break

            if result_placement == "r": # Place next result on the next row
                current_row = str(int(current_row) + 1)
            else: # Place next result on the next column
                current_col = get_next_column(current_col)

            if current_pos.isnumeric():
                current_pos = str(int(current_pos) + 1)
            else:
                current_pos = get_next_column(current_pos)

            item_no = item_no + 1

        processing = False

        print("Done!")

    def stop_processing(e):
        global ended

        ended = True
        process_message.value = "Stopping..."

        page.update()

    def handle_window_event(e):
        global exit_app
        global processing

        if e.data == "close":
            if not processing:
                page.window_destroy()
                return

            exit_app = True
            stop_processing("")

    page.title = "SheetGPT"
    page.scroll = True
    page.window_width = 790
    page.window_height = 600
    page.window_max_width = 790
    page.window_min_width = 510
    page.window_min_height = 300
    page.window_prevent_close = True
    page.on_window_event = handle_window_event

    input_text = ft.Text(
        "No input file selected",
        overflow=ft.TextOverflow.ELLIPSIS,
        max_lines=1,
        col={"md": 9}
    )

    output_text = ft.Text(
        "No output file selected",
        overflow=ft.TextOverflow.ELLIPSIS,
        max_lines=1,
        col={"md": 9}
    )

    input_error = ft.Text("", color=ft.colors.RED, visible=False)
    output_error = ft.Text("", color=ft.colors.RED, visible=False)

    input_progress = ft.Row(
        [ft.ProgressRing(width=16, height=16)],
        col={"md": 9},
        visible=False
    )

    output_progress = ft.Row(
        [ft.ProgressRing(width=16, height=16)],
        col={"md": 9},
        visible=False
    )

    process_progress = ft.ProgressRing(visible=False)

    def check_validity():
        global valid_fields
        global result_book
        global result_sheet
        global workbook
        global worksheet
        global sheet_name
        global result_path
        global file_path

        valid = True

        # print("Checking validity ======")

        for item in valid_fields:
            if not valid_fields[item]:
                valid = False
                # print("Not valid: ", item)

        if result_book == None or workbook == None or worksheet == None or result_sheet == None or sheet_name == "" or file_path == "" or result_path == "":
            valid = False

        process_button.disabled = not valid

        page.update()

    def open_url(e):
        page.launch_url(e.data)

    def on_choose_file(e: ft.FilePickerResultEvent):
        global file_path
        global workbook
        global worksheet
        global result_path
        global result_book
        global result_sheet
        global valid_fields
        global sheet_name

        if e.files == None or len(e.files) != 1:
            input_progress.visible = False
            input_text.visible = True

            page.update()
            check_validity()

            return

        workbook = None
        worksheet = None
        result_book = None
        result_sheet = None
        sheet_name = ""
        file_path = e.files[0].path
        result_path = ""
        valid_fields["output_file"] = False
        valid_fields["input_file"] = False

        input_error.visible = False
        input_text.visible = False
        input_text.value = "No input file selected"

        output_button.disabled = True
        output_error.visible = False
        output_text.visible = False
        output_text.value = "No output file selected"

        sheet_dropdown.value = None
        sheet_dropdown.options = []
        sheet_dropdown.disabled = True

        try:
            workbook = load_input_workbook(file_path)
            input_text.value = file_path

            options = []

            for item in workbook.sheetnames:
                options.append(ft.dropdown.Option(item))

            sheet_dropdown.options = options
            sheet_dropdown.disabled = False

            if len(options) > 0:
                sheet_name = workbook.sheetnames[0]
                sheet_dropdown.value = sheet_name
                worksheet = workbook[sheet_name]

                valid_fields["input_file"] = True
                valid_fields["sheet"] = True
            else:
                sheet_dropdown.error_text = "No sheets found. Please use another file"
                valid_fields["sheet"] = False
        except Exception as e:
            input_error.visible = True
            input_error.value = str(e)

        input_text.visible = True
        input_progress.visible = False
        output_button.disabled = False
        output_text.visible = True

        page.update()
        check_validity()

    def on_choose_save(e: ft.FilePickerResultEvent):
        global result_path
        global result_book
        global valid_fields
        global sheet_name
        global result_sheet
        global file_path

        if e.path == None:
            output_progress.visible = False
            output_text.visible = True

            page.update()
            check_validity()

            return

        result_book = None
        result_sheet = None
        result_path = e.path

        output_error.visible = False
        output_text.visible = False
        output_text.value = "No output file selected"
        valid_fields["output_file"] = False

        try:
            result_book = create_output_book(file_path, result_path, True)
            output_text.value = result_path
            result_sheet = result_book[sheet_name]
            valid_fields["output_file"] = True
        except Exception as e:
            output_error.visible = True
            output_error.value = str(e)

        output_text.visible = True
        output_progress.visible = False

        page.update()
        check_validity()

    file_picker = ft.FilePicker(on_result=on_choose_file)
    save_picker = ft.FilePicker(on_result=on_choose_save)

    page.overlay.append(file_picker)
    page.overlay.append(save_picker)

    def on_input_button_clicked(e):
        input_text.visible = False
        input_progress.visible = True

        page.update()
        check_validity()

        file_picker.pick_files(
            allow_multiple=False,
            dialog_title="Choose input file",
            allowed_extensions=["xlsx", "xlsm", "xltx", "xltm"],
            file_type=ft.FilePickerFileType.CUSTOM
        )

    input_button = ft.FilledButton(
        "Choose input file",
        on_click=on_input_button_clicked,
        col={"md": 3}
    )

    def on_output_button_clicked(e):
        output_text.visible = False
        output_progress.visible = True

        page.update()
        check_validity()

        save_picker.save_file(
            dialog_title="Choose output file",
            allowed_extensions=["xlsx", "xlsm", "xltx", "xltm"],
            file_type=ft.FilePickerFileType.CUSTOM,
            file_name=generate_result_path(file_path)
        ),

    output_button = ft.OutlinedButton(
        "Choose output file",
        on_click=on_output_button_clicked,
        disabled=True,
        col={"md": 3}
    )

    def sheet_changed(e):
        global sheet_name
        global worksheet
        global workbook
        global result_book
        global result_sheet

        sheet_name = sheet_dropdown.value
        worksheet = workbook[sheet_name]
        result_sheet = result_book[sheet_name]

    sheet_dropdown = ft.Dropdown(
        label="Sheet *",
        on_change=sheet_changed,
        disabled=True,
        col={"md": 4}
    )

    def inputs_changed(e):
        global inputs
        global until
        global input_pos
        global INPUTS_REGEX
        global valid_fields

        inputs = []
        input_pos = ""
        until = ""
        input_pos_field.disabled = True
        until_field.disabled = True
        input_pos_field.value = ""
        until_field.value = ""
        valid_fields["inputs"] = False

        if e.control.value != "" and search(INPUTS_REGEX, e.control.value) != None:
            inputs_field.error_text = ""
            inputs = split(",| ", e.control.value.upper())
            input_pos_field.disabled = False
            valid_fields["inputs"] = True
        else:
            inputs_field.error_text = "You must choose either row(s) (i.e. '2' or '4,2,0,6,9'), or column(s) (i.e. 'd' or 'p,r,n,d'). Not both (i.e. '1,a,c,4')."

        page.update()
        check_validity()

    inputs_field = ft.TextField(
        label="Input columns or rows (comma ',' or space ' ' separated) *",
        on_change=inputs_changed,
        max_length=500,
        multiline=False,
        col={"md": 8}
    )

    def input_pos_changed(e):
        global input_pos
        global inputs
        global until
        global valid_fields

        input_pos = ""
        input_pos_field.error_text = ""
        until_field.error_text = ""
        until_field.disabled = True
        valid_fields["input_pos"] = False

        if inputs[0].isnumeric():
            input_pos_field.error_text = "Please enter a valid column (must be a character)"
        else:
            input_pos_field.error_text = "Please enter a valid row (must be a number)"

        if inputs[0].isnumeric() and e.control.value.strip().isalpha():
            input_pos = e.control.value.strip().upper()
            until_field.disabled = False
            input_pos_field.error_text = ""
            valid_fields["input_pos"] = True

        if inputs[0].isalpha() and e.control.value.strip().isnumeric():
            input_pos = e.control.value.strip()
            until_field.disabled = False
            input_pos_field.error_text = ""
            valid_fields["input_pos"] = True

        if input_pos != "" and until != "" and input_pos.isalpha() and until.isnumeric():
            until_field.error_text = "Please enter a valid column (must be a character)"
            valid_fields["until"] = False
        elif input_pos != "" and until != "" and input_pos.isnumeric() and until.isalpha():
            until_field.error_text = "Please enter a valid row (must be a number)"
            valid_fields["until"] = False

        if input_pos.isnumeric() and until.isnumeric() and int(input_pos) >= int(until):
            until_field.error_text = "Ending row cannot be equal or smaller than the starting row"
            valid_fields["until"] = False

        if input_pos.isalpha() and until.isalpha() and get_numeric_value(input_pos) >= get_numeric_value(until):
            until_field.error_text = "Ending column cannot be equal or before the starting column"
            valid_fields["until"] = False

        page.update()
        check_validity()

    input_pos_field = ft.TextField(
        label="Input starting position *",
        on_change=input_pos_changed,
        max_length=10,
        col={"md": 6},
        multiline=False,
        disabled=True
    )

    def until_changed(e):
        global input_pos
        global inputs
        global until
        global valid_fields

        until = ""
        valid_fields["until"] = False

        if input_pos.isalpha():
            until_field.error_text = "Please enter a valid column (must be a character)"
        else:
            until_field.error_text = "Please enter a valid row (must be a number)"

        if e.control.value.strip() == "":
            valid_fields["until"] = True
            until_field.error_text = ""

        if input_pos.isalpha() and e.control.value.strip().isalpha():
            until = e.control.value.strip().upper()
            until_field.error_text = ""
            valid_fields["until"] = True

        if input_pos.isnumeric() and e.control.value.strip().isnumeric():
            until = e.control.value.strip().upper()
            until_field.error_text = ""
            valid_fields["until"] = True

        if until != "" and input_pos.isnumeric() and until.isnumeric() and get_numeric_value(input_pos) >= get_numeric_value(until):
            until = ""
            until_field.error_text = "Ending row cannot be equal or smaller than the starting row"

        if until != "" and input_pos.isalpha() and until.isalpha() and get_numeric_value(input_pos) >= get_numeric_value(until):
            until = ""
            until_field.error_text = "Ending column cannot be equal or before the starting column"

        page.update()
        check_validity()

    until_field = ft.TextField(
        label="Input ending",
        on_change=until_changed,
        max_length=10,
        col={"md": 6},
        multiline=False,
        disabled=True
    )

    def output_row_changed(e):
        global result_row
        global valid_fields

        result_row = e.control.value.strip()
        output_row_field.error_text = ""
        valid_fields["output_row"] = True

        if result_row == "" or not result_row.isnumeric():
            output_row_field.error_text = "Please enter a valid row"
            result_row = ""
            valid_fields["output_row"] = False

        page.update()
        check_validity()

    output_row_field = ft.TextField(
        label="Output row *",
        on_change=output_row_changed,
        max_length=10,
        col={"md": 4},
        multiline=False
    )

    def output_col_changed(e):
        global result_col
        global valid_fields

        result_col = e.control.value.strip()
        output_col_field.error_text = ""
        valid_fields["output_col"] = True

        if result_col == "" or not result_col.isalpha():
            output_col_field.error_text = "Please enter a valid column"
            result_col = ""
            valid_fields["output_col"] = False

        page.update()
        check_validity()

    output_col_field = ft.TextField(
        label="Output column *",
        on_change=output_col_changed,
        max_length=10,
        col={"md": 4},
        multiline=False
    )

    def output_placement_changed(e):
        global result_placement

        if output_placement_dropdown.value == "Place on the next row":
            result_placement = "r"
        elif output_placement_dropdown.value == "Place on the next column":
            result_placement = "c"

    output_placement_dropdown = ft.Dropdown(
        label="Output placement *",
        on_change=output_placement_changed,
        options=[
            ft.dropdown.Option("Place on the next row"),
            ft.dropdown.Option("Place on the next column")
        ],
        col={"md": 4},
        value="Place on the next row"
    )

    def api_key_changed(e):
        global api_key
        global valid_fields

        api_key = e.control.value.strip()
        valid_fields["api_key"] = False
        api_key_field.error_text = ""

        if api_key == "":
            api_key_field.error_text = "Please enter a valid API key"
        else:
            valid_fields["api_key"] = True

        page.update()
        check_validity()

    api_key_field = ft.TextField(
        label="OpenAI API Key *",
        on_change=api_key_changed,
        expand=False,
        max_length=64,
        col={"md": 8},
        multiline=False
    )

    def model_changed(e):
        global model
        model = model_dropdown.value

    dropdown_options = []

    for item in MODELS:
        dropdown_options.append(ft.dropdown.Option(item))

    model_dropdown = ft.Dropdown(
        label="ChatGPT Model *",
        on_change=model_changed,
        options=dropdown_options,
        col={"md": 4},
        value=model
    )

    api_key_info = ft.Markdown(
        "You can get your OpenAI API key [here](https://platform.openai.com/account/api-keys)",
        on_tap_link=open_url
    )

    def system_prompt_changed(e):
        global system_prompt
        global valid_fields

        system_prompt = e.control.value.strip()
        valid_fields["system_prompt"] = False
        system_prompt_field.error_text = ""

        if system_prompt == "":
            system_prompt_field.error_text = "Please enter a valid system prompt"
        else:
            valid_fields["system_prompt"] = True

        page.update()
        check_validity()

    system_prompt_field = ft.TextField(
        label="System prompt *",
        on_change=system_prompt_changed,
        max_length=1024,
        multiline=True,
        value=DEFAULT_SYSTEM_PROMPT
    )

    def prompt_changed(e):
        global prompt
        global valid_fields

        prompt = e.control.value.strip()
        valid_fields["prompt"] = False
        prompt_field.error_text = ""

        if prompt == "":
            prompt_field.error_text = "Please enter a valid prompt"
        else:
            valid_fields["prompt"] = True

        page.update()
        check_validity()

    prompt_field = ft.TextField(
        label="Prompt *",
        on_change=prompt_changed,
        max_length=1024,
        multiline=True
    )

    prompt_info_md = "Enter a prompt to be used with each request. Use zero-based `$0` placeholder(s) to insert inputs.\n### Example\nAssuming you have entered `a,b` as the inputs, and `2` as the starting row:\n\n_\"Bla bla $0 and $0, $1\"_\n\nwill be converted to\n\n_\"Bla bla <VALUE OF A2> and <VALUE OF A2>, <VALUE OF B2>\"_"

    prompt_info = ft.Markdown(prompt_info_md)

    def limit_changed(e):
        global limit
        global valid_fields

        limit = 0
        limit_field.error_text = ""
        valid_fields["limit"] = False

        if e.control.value.strip() == "" or not e.control.value.strip().isnumeric() or int(e.control.value.strip()) < 0:
            limit_field.error_text = "Please enter a valid limit"
        else:
            limit = int(e.control.value.strip())
            valid_fields["limit"] = True

        page.update()
        check_validity()

    limit_field = ft.TextField(
        label="Processing limit *",
        on_change=limit_changed,
        max_length=10,
        col={"md": 3},
        multiline=False,
        value="0"
    )

    def on_skip_changed(e):
        global skip
        skip = skip_switch.value

    skip_switch = ft.Switch(
        label="Skip the existing results",
        value=skip,
        on_change=on_skip_changed
    )

    def start_processing(e):
        global file_path
        global result_path
        global api_key
        global model
        global sheet_name
        global system_prompt
        global prompt
        global result_row
        global result_col
        global inputs
        global workbook
        global worksheet
        global result_book
        global result_sheet
        global result_placement
        global input_pos
        global cache
        global limit
        global processed
        global until
        global skip
        global overwrite_output
        global process_task
        global ended
        global valid_fields
        global MODELS
        global DEFAULT_SYSTEM_PROMPT
        global exit_app

        process_error.value = ""
        process_message.value = "Processing..."
        ended = False
        process_spinner.visible = True
        stop_button.disabled = False
        stop_button.visible = True
        process_button.disabled = True
        process_button.visible = False
        skip_switch.disabled = True
        limit_field.disabled = True
        api_key_field.disabled = True
        prompt_field.disabled = True
        system_prompt_field.disabled = True
        until_field.disabled = True
        inputs_field.disabled = True
        input_pos_field.disabled = True
        output_col_field.disabled = True
        output_row_field.disabled = True
        output_placement_dropdown.disabled = True
        model_dropdown.disabled = True
        sheet_dropdown.disabled = True
        input_button.disabled = True
        output_button.disabled = True

        page.update()

        try:
            process_worksheet()
        except Exception as e:
            print("Failed processing")
            print(e)
            process_error.value = "An error occurred"

        process_message.value = "Saving..."

        page.update()

        process_message.value = "Failed"

        try:
            save_and_close(workbook, result_book, result_path)
            process_message.value = "Saved to " + result_path
        except Exception as e:
            print("Failed saving")
            print(e)
            process_error.value = "An error occurred while saving"

        process_spinner.visible = False
        stop_button.disabled = True
        stop_button.visible = False
        process_button.disabled = False
        process_button.visible = True
        skip_switch.disabled = False
        limit_field.disabled = False
        api_key_field.disabled = False
        prompt_field.disabled = False
        system_prompt_field.disabled = False
        until_field.disabled = True
        inputs_field.disabled = False
        input_pos_field.disabled = True
        output_col_field.disabled = False
        output_row_field.disabled = False
        output_placement_dropdown.disabled = False
        model_dropdown.disabled = False
        sheet_dropdown.disabled = True
        input_button.disabled = False
        output_button.disabled = True

        reset_globals()

        limit_field.value = limit
        api_key_field.value = api_key
        prompt_field.value = prompt
        system_prompt_field.value = DEFAULT_SYSTEM_PROMPT
        until_field.value = until
        inputs_field.value = ",".join(inputs)
        input_pos_field.value = input_pos
        output_col_field.value = result_col
        output_row_field.value = result_row
        output_placement_dropdown.value = "Place on the next row"
        model_dropdown.value = model
        sheet_dropdown.value = None
        input_text.value = "No input file selected"
        output_text.value = "No output file selected"
        skip_switch.value = skip

        page.update()
        check_validity()

        if exit_app: page.window_destroy()

    process_spinner = ft.Row(
        [ft.ProgressRing(width=16, height=16)],
        col={"md": 1},
        alignment=ft.MainAxisAlignment.END,
        visible=False
    )

    process_message = ft.Text(
        "(*) Required fields",
        overflow=ft.TextOverflow.ELLIPSIS,
        max_lines=1
    )

    process_button = ft.ElevatedButton(
        "Start processing",
        disabled=True,
        on_click=start_processing,
        col={"md": 3}
    )

    stop_button = ft.OutlinedButton(
        "Stop processing",
        disabled=True,
        visible=False,
        on_click=stop_processing,
        col={"md": 3}
    )

    process_error = ft.Text(
        "",
        color=ft.colors.RED
    )

    root_column = ft.Column(
        [
            ft.Column(
                controls=[
                    ft.Text(
                        "SheetGPT",
                        style=ft.TextThemeStyle.HEADLINE_MEDIUM,
                        text_align=ft.TextAlign.CENTER
                    ),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            ),
            ft.Text("File configuration", style=ft.TextThemeStyle.TITLE_LARGE),
            ft.Text("Input file", style=ft.TextThemeStyle.TITLE_MEDIUM),
            ft.ResponsiveRow(
                [
                    input_button,
                    input_progress,
                    input_text
                ],
                spacing=10,
                vertical_alignment=ft.CrossAxisAlignment.CENTER
            ),
            input_error,
            ft.Text("Output file", style=ft.TextThemeStyle.TITLE_MEDIUM),
            ft.ResponsiveRow(
                [
                    output_button,
                    output_progress,
                    output_text
                ],
                spacing=10,
                vertical_alignment=ft.CrossAxisAlignment.CENTER
            ),
            output_error,
            ft.Divider(),
            ft.Text("Input configuration", style=ft.TextThemeStyle.TITLE_LARGE),
            ft.ResponsiveRow(
                [
                    sheet_dropdown,
                    inputs_field
                ],
                vertical_alignment=ft.CrossAxisAlignment.START,
                spacing=10
            ),
            ft.ResponsiveRow(
                [
                    input_pos_field,
                    until_field
                ],
                vertical_alignment=ft.CrossAxisAlignment.START,
                spacing=10
            ),
            ft.Divider(),
            ft.Text("Output configuration", style=ft.TextThemeStyle.TITLE_LARGE),
            ft.ResponsiveRow(
                [
                    output_col_field,
                    output_row_field,
                    output_placement_dropdown
                ],
                vertical_alignment=ft.CrossAxisAlignment.START,
                spacing=10
            ),
            ft.Divider(),
            ft.Text("ChatGPT configuration", style=ft.TextThemeStyle.TITLE_LARGE),
            api_key_info,
            ft.ResponsiveRow(
                [
                    model_dropdown,
                    api_key_field
                ],
                vertical_alignment=ft.CrossAxisAlignment.START,
                spacing=10
            ),
            system_prompt_field,
            ft.Divider(),
            ft.Text("Prompt configuration", style=ft.TextThemeStyle.TITLE_LARGE),
            prompt_info,
            prompt_field,
            ft.Divider(),
            ft.Text("Options", style=ft.TextThemeStyle.TITLE_LARGE),
            skip_switch,
            ft.Text("Results will be overwritten if skipping is disabled"),
            ft.Divider(),
            ft.Text("Processing limit", style=ft.TextThemeStyle.TITLE_MEDIUM),
            ft.Text("Enter a limit for number of items to be processed by ChatGPT. Cached or existing results do not count towards the limit. (type 0 to remove limit)"),
            limit_field,
            ft.Divider(),
            ft.ResponsiveRow(
                [
                    stop_button,
                    process_button,
                    ft.Row(
                        [process_message],
                        col={"md": 8},
                        vertical_alignment=ft.CrossAxisAlignment.CENTER
                    ),
                    process_spinner
                ],
                spacing=10,
                vertical_alignment=ft.CrossAxisAlignment.CENTER
            ),
            process_error
        ],
        spacing=15,
        run_spacing=15
    )

    root_container = ft.SafeArea(
        content=ft.Container(
            padding=5,
            content=root_column
        ),
    )

    page.add(root_container)

ft.app(target=main)
