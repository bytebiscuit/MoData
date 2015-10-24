__author__ = 'fatosismali'

import xlrd
import os, fnmatch


mypath = "/Users/fatosismali/Desktop/DataScience/MinistryOfData/datasets/Kosovo datasets/"

column_array = []

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
                cols = cols[:-1]
                column_array.append(cols.split(','))
                cols = ""
def create_umbrella_model(columns_array):
    flattened = [col for list_of_cols in columns_array for col in list_of_cols]
    trimmed = [x.strip() for x in flattened]
    lowercase = [y.lower() for y in trimmed]
    unique_set = set(lowercase)
    print(unique_set)


find_files(mypath, '*.xlsx')

create_umbrella_model(column_array)



