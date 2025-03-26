# First time set up

1. Install python 3. This should also install pip.
2. Download the code.
3. In the code directory, run `pip install openpyxl`

# Main Script

Converts the marking excel/google sheets to txt files in a nice format, and adds statistics + student totals to the source excel file. 
 
## How to use

Run `python main.py <excel sheet file path>`.

## Script settings

Configs: You can open the code and change these constants depending on your excel workbook:

| Line | Constant name     | What it affects                                                                                                                                                                                                                                                                                                                            |
| ---- | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 6    | HEADING_COLUMN    | Change the heading column in the excel sheet. This is the excel column with all question headings. Any cell that is **bold** in this column will appear as an underlined heading in the _feedback.txt_ file. This also can have a cell with test 'total' and that would show the total marks of each student in _feedback.txt_. Default: A |
| 7    | FEEDBACK_COLUMN   | Change the feedback column. This is the excel column where the markers will write the feedback for a question's part. All of this column will be copied in the _feedback.txt_ under its specific question heading as specified in HEADING_COLUMN. Default: E                                                                               |
| 8    | MAX_ROWS_TO_CHECK | The lower the value, the faster the script. This is the maximum number of rows that the excel sheet will have. Default: 90                                                                                                                                                                                                                 |

## Excel sheet format

> The **sample** folder has a sample excel sheet and its output file.

The excel file can have many worksheets, each for a student. Once rubrics are set up in one sheet, it can be duplicated but right-clicking the sheet and selecting `move or copy`. The name of the sheet will be the name that shows up on the _feedback.txt_ file. 

All question headings ( Example: **Question 1**, or **Question 2 (10 marks)** ) can be added to the HEADING_COLUMN. Any value in this column which is bold, will be printed to the _feedback.txt_ file as a heading. Any other non-bold text in this column is for marker's reference only (example: File format should be .pdf). 

The next columns could be anything that is needed for  marker's reference. There can, for instance, be columns for sub-heading, max marks, marks achieved etc. 

All feedback for a certain question(or a question's sub-part) is to be written in FEEDBACK_COLUMN. All contents of this column will be copied down as such to the _feedback.txt_ file. 

Any more columns can be added for marker's reference.

The HEADING_COLUMN can have a value 'total' (should NOT be bold) and if it does, it will add the total of all the questions to the _feedback.txt_ file (total will be grabbed from the column that comes before the FEEDBACK_COLUMN). So its a goooood idea to have the marks achieved column right before the FEEDBACK_COLUMN. See sample folder for examples. 

# Update Script

Creates a new duplicated excel file where a set range of cells from one source sheet are copied over to all the other sheets in the workbook. Useful when there is a need to update marks rubrics for all students after part-marking. 
 
## How to use

Run `python update.py <excel sheet file path>`.