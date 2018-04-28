import xlrd
import random
import string

file = "list.xlsx"
workbook = xlrd.open_workbook(file)                 # Load the workbook in specified file.
sheetMovies = workbook.sheet_by_name("Movies")      # Load the sheet in specified workbook.
rows = sheetMovies.nrows                                  # Counts how many rows in sheet.

print("There are currently", rows, "tiles in the list. Let's pick one.")

randomNumber = random.randint(1, rows)              # Create a random number based on total rows in the list.
randomPick = sheetMovies.cell_value(randomNumber-1, 0)    # Matches the random number with the cell value.

if randomPick.istitle == True:
    # Checks if title format is followed.
    print("How about watching '" + randomPick + "'?")
else:
    # If title format was previously not met, it prints it out in a correct title form.
    print("How about watching '" + string.capwords(randomPick) + "'?")