import openpyxl as xl
from openpyxl.styles import Font
import sys

# Some constants 
HEADINGS_COLUMN = 'A'
FEEDBACK_COLUMN = 'E'
MAX_ROWS_TO_CHECK = 90 # efficiency

def main():
    
    if len(sys.argv) < 2:
        print('File name not provided')
        return 

    # Load Workbook with only data - no formulas
    wb = xl.load_workbook(sys.argv[1], data_only=True)

    # write this to the final txt file
    output = ""

    running_total = 0

    # Loop through all worksheets in the workbook
    for sheet in wb.worksheets:
  
        # number of rows to check
        max_rows = min(MAX_ROWS_TO_CHECK, sheet.max_row)

        # Add student name to the output
        output += "<" + sheet.title + ">\n\n"

        # For each row in the sheet - Leave first heading row. 
        for row in range(1, max_rows):
            # Get the heading and feedback cells in the current row
            current_heading_cell = sheet[HEADINGS_COLUMN][row]
            current_feedback_cell = sheet[FEEDBACK_COLUMN][row]
            
            # Check if we have a heading => the font will be bold
            if current_heading_cell.font.bold and current_heading_cell.value is not None and str(current_heading_cell.value).strip().lower() != 'total':     
                output += "\n\n" + current_heading_cell.value + ":\n"

                # After each heading, underline with 'n' hyphens 
                for _ in range(0, len(current_heading_cell.value)):
                    output += "-"

                output +="\n"                


            #  If the feedback coloumn is not None, we can latee print it to file.   
            if current_feedback_cell.value is not None:
                output += current_feedback_cell.value + "\n"

            # Print total marks 
            if str(current_heading_cell.value).strip().lower() == 'total':
                total = sheet[chr(ord(FEEDBACK_COLUMN) - 1)][row].value
                running_total += total

                output += "\n\n\n TOTAL: " + str(total)

        output += "\n\n\n=====================================\n\n\n"
    
    # Statistics 
    #=====================================================
    # Average = running total / total sheets - 1.
    # total sheets - 1 since 1 sheet is the empty rubric sheet. 
    avg = round(running_total / (len(wb.sheetnames) - 1), 2)
    output = f"Statistics\n==========\n\nTotal: {(len(wb.sheetnames) - 1)}\nAverage: {avg}\n\n\n=====================================\n{output}"
    #=====================================================

    if write_to_file(output) == 0:
        print('Success')
    else:
        print('Something went wrong!')

def write_to_file(content):
    '''
    Writes to a txt file and returns 0, if an Exception occurs, returns 1
    '''
    
    try:
        # open a new file in write mide 
        with open("feedback.txt", "w") as text_file:
            text_file.write(content)
        
        return 0
    except Exception:
        return 1


# run main method if the current file is run
if __name__ == "__main__":
    main()