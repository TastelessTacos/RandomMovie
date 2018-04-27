import xlrd
import random

file = "list.xlsx"
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)