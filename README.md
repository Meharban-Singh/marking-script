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

Configs: You can open the code and change these constants depending on your excel workbook:

| Line | Constant name     | What it affects                                                                                                                                                                                                                                                                                                                    |
| ---- | ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 6    | HEADING_COLUMN    | Change the heading column in the excel sheet. This is the excel column with all question headings. Any cell that is bold in this column will appear as an underlined heading in the feedback.txt file. This also can have a cell with test 'total' and that would show the total marks of each student in feedback.txt. Default: A |
| 7    | FEEDBACK_COLUMN   | Change the feedback column. This is the excel column where the markers will write the feedback for a question's part. All of this column will be copied in the feedback.txt under its specific question heading as specified in HEADING_COLUMN. Default: E                                                                         |
| 8    | MAX_ROWS_TO_CHECK | The lower the value, the faster the script. This is the maximum number of rows that the excel sheet will have. Default: 90                                                                                                                                                                                                         |
