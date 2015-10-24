__author__ = 'fatosismali'

import xlrd
import os
from os import listdir
from os.path import isfile, join


mypath = "/Users/fatosismali/Desktop/DataScience/MinistryOfData/datasets/Kosovo datasets/THV_owns source revenues/"
os.chdir(mypath)
onlyfiles = [ f for f in listdir(mypath) if isfile(join(mypath,f)) ]
cols = ""
for e in onlyfiles:
    workbook = xlrd.open_workbook(e)
    worksheet = workbook.sheet_by_index(0)
    for idx, cell_obj in enumerate(worksheet.row(0)):
        cols =  cols + cell_obj.value + ","
    print(cols)
    cols = ""