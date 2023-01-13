import os
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.worksheet import worksheet
import time


file = 'Catalog_ELSIE.xls'
exel_file = "test.xlsx"
def GetCodeFromExcel():
    try:
        excel_data_df = pd.read_excel(file, sheet_name='Автостёкла и аксессуары')
        return excel_data_df['Код Элси'].tolist()
    except:
        return False

def SavePriceToExcel(df):
    #sheet_name = 'price ' + datetime.now().strftime("%m/%d/%Y")
    writer = pd.ExcelWriter('test.xlsx', engine='openpyxl', mode='w')
    df.to_excel(writer, sheet_name="sheet_name", index=False)
    writer.save()

def GetExcel(path):
    excel_data_df = pd.read_excel(path, sheet_name='Автостёкла и аксессуары')
    return excel_data_df


def SaveRowToExcel(row, excel_file):

    book = openpyxl.load_workbook(filename=excel_file)
    sheet: worksheet = book.worksheets[0]
    sheet.append(row)
    book.save(excel_file)

def CreatExelFile():
    path = os.path.expanduser("~/Documents")

    exelFileName = path + "/Catalog_ELSIE_Price_" + datetime.now().strftime('%m-%d-%Y_%H:%M:%S') + '.xlsx'
    print(exelFileName)

    book = openpyxl.Workbook()
    book.save(exelFileName)
    return  exelFileName
#df = pd.read_excel(file, sheet_name='Автостёкла и аксессуары')


def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()