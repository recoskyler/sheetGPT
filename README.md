# SheetGPT

<img src="https://github.com/recoskyler/sheetGPT/blob/main/icon.png" data-canonical-src="https://github.com/recoskyler/sheetGPT/blob/main/icon.png" width="128" height="128" />

A visual tool for completing a specific column/row using the data from other columns/rows in a XLSX/XLSM file utilizing ChatGPT.

## Status

Still in development, but can be used.

## Usage

1. Go to releases, and download the latest executable suitable for your platform.
2. Launch the executable.
3. Choose an input file.
4. Choose an output file.
5. Choose the sheet you would like to use.
6. Enter the input column(s)/row(s), separated by a comma "," or a space " " character. (*i.e. `a,b,c` or `1 2 5`*)
7. Enter the input starting position. You must enter a row if you have entered a column in the previous step, vice-versa.
8. *(Optional) Enter the input ending position. You must enter a row if you have entered a column in the previous step, vice-versa. SheetGPT will stop once it reaches this position (inclusive).*
9. Enter the output column and row.
10. Choose the output placement strategy (place the result on the next row, or column).
11. Enter your OpenAI API key. You can create one [here](https://platform.openai.com/account/api-keys).
12. Choose the ChatGPT model you would like to use.
13. Enter a system prompt to fine-tune the model to your needs. You can also use the default system prompt.
14. Enter a prompt to be used with each request. You can use `$0`, `$1`... placeholders (*zero-indexed, starting from 0*) to insert the inputs you have specified on the 6th step above.
15. *(Optional) Enter a processing limit, enter `0` to remove the limit. Only the items which were processed explicitly by ChatGPT will count towards the limit. Cached results, or existing results (if you leave* **Skip the existing results** *enabled) will not count towards the limit.*
16. Press **Start processing** to start. This might take a while depending on the data volume and device specs.

## Requirements

- Python 3.7+
- pip
- venv
- Flet
- ezpyi (for creating AppImage)

## Development

1. Clone the project and open the folder:

    ```bash
    git clone https://github.com/recoskyler/sheetGPT

    cd sheetGPT
    ```

2. Create a new virtual environment:

    ```bash
    python3 -m venv env
    ```

3. Activate the virtual environment:

    - Linux/MacOS

    ```bash
    source env/bin/activate
    ```

    - Windows (Powershell)

    ```ps1
    env/Scripts/Activate.ps1
    ```

    - Windows (CMD)

    ```cmd
    env/Scripts/activate.bat
    ```

4. Install dependencies:

    ```bash
    pip3 install -r requirements.txt
    ```

5. Run the script:

    ```bash
    python3 main.py
    ```

    Or to use hot-reload

    ```bash
    flet run -d main.py
    ```

### Packaging

#### Executable

Executable will be created for your platform (for Linux if you run it on Linux, for Windows if you run it on Windows...).

```bash
flet pack main.py --icon icon.png --name SheetGPT --add-data "modules:modules"
```

#### AppImage

An AppImage will be created to be used with Linux.

```bash
ezpyi -A -i icon.png main.py SheetGPT
```

## [LICENSE](https://github.com/recoskyler/sheetGPT/blob/main/LICENSE)

[MIT License](https://github.com/recoskyler/sheetGPT/blob/main/LICENSE)

## About

Made by [recoskyler](https://github.com/recoskyler) - 2023
