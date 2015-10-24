__author__ = 'fatosismali'

import xlrd
import os, fnmatch


mypath = "/Users/fatosismali/Desktop/DataScience/MinistryOfData/datasets/Kosovo datasets/"


def find_files(directory, pattern):
    cols = ""
    for root, dirs, files in os.walk(directory):
        for basename in files:
            if fnmatch.fnmatch(basename, pattern):

                filename = os.path.join(root, basename)

                workbook = xlrd.open_workbook(filename)
                worksheet = workbook.sheet_by_index(0)
                for idx, cell_obj in enumerate(worksheet.row(0)):
                    cols =  cols + str(cell_obj.value) + ","
                print(cols)
                cols = ""


find_files(mypath, '*.xlsx')
