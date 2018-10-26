import xlrd
import random
import string

file = "list.xlsx"
workbook = xlrd.open_workbook(file)       # Load the workbook in specified file.
sheetMovies = workbook.sheet_by_name("Movies") # Load the sheet in specified workbook.
sheetTvshows = workbook.sheet_by_name("TvShows")
sheetMDL = workbook.sheet_by_name("MDL")


def movie_pick():
    """Creates a random number which is then matched with the cell value
    in the list of the specified sheet."""
    rowsM = sheetMovies.nrows             # Counts how many rows in the sheet.
    randomNumber = random.randint(1, rowsM)
    randomPick = sheetMovies.cell_value(randomNumber - 1, 0)
    print("There are currently", rowsM, "tiles in the Movies list. Let's pick one.")
    if randomPick.istitle == True:
        # Checks if title format is followed.
        print("How about watching '" + randomPick + "'?")
    else:
        # If title format was previously not met, it prints it out in a correct title form.
        print("How about watching '" + string.capwords(randomPick) + "'?")


def tvshow_pick():
    """Creates a random number which is then matched with the cell value
    in the list of the specified sheet."""
    rowsT = sheetTvshows.nrows
    randomNumber = random.randint(1, rowsT)
    randomPick = sheetTvshows.cell_value(randomNumber - 1, 0)
    print("There are currently", rowsT, "tiles in the TV Shows list. Let's pick one.")
    if randomPick.istitle == True:
        # Checks if title format is followed.
        print("How about watching '" + randomPick + "'?")
    else:
        # If title format was previously not met, it prints it out in a correct title form.
        print("How about watching '" + string.capwords(randomPick) + "'?")


def mdl_pick():
    rowsMDL = sheetMDL.nrows
    randomNumber = random.randint(1, rowsMDL)
    randomPick = sheetMDL.cell_value(randomNumber - 1, 0)
    print("There are currently", rowsMDL, "titles in the MDL list. Let's pick one.")
    if randomPick.istitle == True:
        print("How about watching '" + randomPick + "'?")
    else:
        print("How about watching '" + string.capwords(randomPick) + "'?")


movie_pick()
print()
tvshow_pick()
print()
mdl_pick()
