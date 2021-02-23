# Marking-script

A marking script to conver the marking excel/google sheets to txt files in a nice format. Uses openpyxl to interact with the google sheets,

Converts all worksheets in the excel workbook to nice looking format which can later be used to submit the feedback.

## How to use

1. Install python 3. This should also install pip.
2. Download the code - main.py is the only file that is needed.
3. In the code directory, run `pip install openpyxl`
4. Run `python main.py <excel sheet file path>`.
5. This will create a **feedback.txt** file in the working directory.

## Excel sheet format

> The **sample** folder has a sample excel sheet and its output file.
