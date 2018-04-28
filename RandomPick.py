import xlrd
import random
import string

file = "list.xlsx"
workbook = xlrd.open_workbook(file)                 # Load the workbook in specified file.
sheetMovies = workbook.sheet_by_name("Movies")      # Load the sheet in specified workbook.
sheetTvshows = workbook.sheet_by_name("TvShows")


def movie_pick():
    """Creates a random number which is then matched with the cell value in the list of the specified sheet."""
    randomNumber = random.randint(1, rows)
    randomPick = sheetMovies.cell_value(randomNumber - 1, 0)
    print("There are currently", rows, "tiles in the list. Let's pick one.")
    if randomPick.istitle == True:
        # Checks if title format is followed.
        print("How about watching '" + randomPick + "'?")
    else:
        # If title format was previously not met, it prints it out in a correct title form.
        print("How about watching '" + string.capwords(randomPick) + "'?")


def tvshow_pick():
    """Creates a random number which is then matched with the cell value in the list of the specified sheet."""
    randomNumber = random.randint(1, rows)
    randomPick = sheetTvshows.cell_value(randomNumber - 1, 0)
    print("There are currently", rows, "tiles in the list. Let's pick one.")
    if randomPick.istitle == True:
        # Checks if title format is followed.
        print("How about watching '" + randomPick + "'?")
    else:
        # If title format was previously not met, it prints it out in a correct title form.
        print("How about watching '" + string.capwords(randomPick) + "'?")


choice = input("What do you want to watch? A Movie or A TV Show? ").upper()


# Whatever the user enters will be stored in uppercase and matched with the predefined condition which is already
# in uppercase. This will avoid the condition below to fail if the user simply doesn't follow the predefined
# capitalisation.
if choice == "Movie".upper():
    rows = sheetMovies.nrows   # Counts how many rows in sheet.
    movie_pick()
elif choice == "TV Show".upper():
    rows = sheetTvshows.nrows
    tvshow_pick()
else:
    print("Please enter a valid choice: Movie or TV Show")