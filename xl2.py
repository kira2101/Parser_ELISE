import time

import openpyxl
from openpyxl.worksheet import worksheet
import pandas as pd


def SaveRowToExcel(row, excel_file):

    book = openpyxl.load_workbook(filename=excel_file)
    sheet : worksheet = book.worksheets[0]
    sheet.append(row)
    book.save(excel_file)

def ProgressBar(max_range,max_lenght):
    for i in range(max_lenght):
        time.sleep(0.5)
        print(f'*', end="\r")





