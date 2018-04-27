import xlrd
import random

file = "list.xlsx"
workbook = xlrd.open_workbook(file)     # Load the workbook in specified file.
sheet = workbook.sheet_by_name("list")  # Load the sheet in specified workbook.
rows = sheet.nrows                      # Counts how many rows in sheet.

print("There are currently", rows, "tiles in the list. Let's pick one.")

randomNumber = random.randint(1, rows)              # Create a random number based on total rows in the list.
randomPick = sheet.cell_value(randomNumber-1, 0)    # Matches the random number with the cell value.

if randomPick.istitle == True:
    # Checks if title format is followed.
    print("How about watching " + randomPick + "?")
else:
    # If title format was previously not met, it prints it out in a correct title form.
    print("How about watching " + randomPick.title() + "?")