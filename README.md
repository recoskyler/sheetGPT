# SheetGPT

<img src="https://github.com/recoskyler/sheetGPT/blob/main/assets/icon.png" data-canonical-src="https://github.com/recoskyler/sheetGPT/blob/main/assets/icon.png" width="128" height="128" />

A visual tool for completing a specific column/row using the data from other columns/rows in a XLSX/XLSM file utilizing ChatGPT.

<img src="https://github.com/recoskyler/sheetGPT/blob/main/assets/screenshot-1.png" data-canonical-src="https://github.com/recoskyler/sheetGPT/blob/main/assets/screenshot-1.png" />

<img src="https://github.com/recoskyler/sheetGPT/blob/main/assets/screenshot-2.png" data-canonical-src="https://github.com/recoskyler/sheetGPT/blob/main/assets/screenshot-2.png" />

## Example

### Input table

On "**Sheet A**"

|Row v/Column >|A|B|C|D|
|--:|--:|:-:|--:|--:|
|**1**|**Release Year**|**Game**|**Character 1**|**Character 2**|
|**2**|*1996*| |*Mario*|*Peach*|
|**3**|*2015*|*SOMA*|*Simon Jarrett*|*Catherine Chun*|
|**4**|*2017*| |*Link*|*Zelda*|
|**5**|*2018*| |*Madeline*|*Theo*|
``
### SheetGPT configuration

|Setting|Value|
|---|--|
|**Sheet**|`Sheet A`|
|**Input columns**|`a,c,d`|
|**Input starting position**|`2`|
|**Input ending**|`5`|
|**Output column**|`B`|
|**Output row**|`2`|
|**Output placement**|`Place on the next row`|
|**ChatGPT Model**|`gpt-3.5-turbo-16k`|
|**OpenAI API Key**|`<MY API KEY>`|
|**System Prompt**|`You are a helpful assistant. You give only the answer, without forming a sentence. If you are not sure, try guessing. If you are still unsure about the answer, output '?'. If you don't know the answer, or if you cannot give a correct answer output '?'.`|
|**Prompt**|`What is the game with the characters $1 and $2, released in $0? Output only the name of the game. If you are unsure, or you don't know, output a question mark "?"`|
|**Skip existing results**|`Enabled`|
|**Processing limit**|`2`|

### Output table

|Row v/Column >|A|B|C|D|
|--:|--:|:-:|--:|--:|
|**1**|**Release Year**|**Game**|**Character 1**|**Character 2**|
|**2**|*1996*|Super Mario 64|*Mario*|*Peach*|
|**3**|*2015*|*SOMA*|*Simon Jarrett*|*Catherine Chun*|
|**4**|*2017*|The Legend of Zelda: Breath of the Wild|*Link*|*Zelda*|
|**5**|*2018*| |*Madeline*|*Theo*|

*As you can see, the last row was not completed as the* "**Processing limit**" *was set to* `2`

## Status

**Stable**, can be used.

**DO NOT FORGET TO BACK UP YOUR DATA BEFOREHAND**, even though a separate output file is being created, it is best to be safe.

## Usage

1. [Download the latest release](https://github.com/recoskyler/sheetGPT/releases/latest) suitable for your platform.
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
    python -m venv env
    ```

3. Activate the virtual environment:

    - Linux/MacOS

    ```bash
    source env/bin/activate
    ```

    - Windows (Powershell)

    ```ps1
    .\env\Scripts\Activate.ps1
    ```

    - Windows (CMD)

    ```cmd
    .\env\Scripts\activate.bat
    ```

4. Install dependencies:

    ```bash
    pip install -r requirements.txt
    ```

5. Run the script:

    ```bash
    python main.py
    ```

    Or to use hot-reload

    ```bash
    flet run -d main.py
    ```

### Packaging

Executable will be created for your platform (for MacOS if you run it on MacOS, for Windows if you run it on Windows...).

#### MacOS Executable

```bash
flet pack main.py --icon assets/icon.png --name SheetGPT --product-name SheetGPT --product-version v1.0.2 --copyright MIT --bundle-id com.recoskyler.sheetgpt --add-data "assets:assets"
```

#### Windows Executable

```bat
flet pack main.py --icon assets\icon.png --name SheetGPT --product-name SheetGPT --product-version v1.0.2 --file-version v1.0.2 --file-description SheetGPT --copyright MIT --add-data "assets;assets"
```

#### Linux Executable

```bash
flet pack main.py --icon assets/icon.png --name SheetGPT --add-data "assets:assets"
```

## [LICENSE](https://github.com/recoskyler/sheetGPT/blob/main/LICENSE)

[MIT License](https://github.com/recoskyler/sheetGPT/blob/main/LICENSE)

## About

Made by [recoskyler](https://github.com/recoskyler) - 2023
