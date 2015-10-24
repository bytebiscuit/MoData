__author__ = 'fatosismali'

import xlrd
import os, fnmatch
import xlwt



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
    print(list(unique_set))


# def combine_excel_files(directory,pattern):
#     with open('combined.csv','wb') as csvfile:
#         csv_writer = csv.writer(csvfile, delimiter = ',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
#         csv_writer.writerow(['', 'tipi 2', 'kolona 5', 'viti', 'budget of kosovo - planning', 'data', 'kolona 3', '4', 'tipi 5', 'departamenti', 'shuma', 'tipi 6', 'grupi', 'komuna', 'tipi 3', 'kolona 2', 'kolona 1', 'tipi 1', 'muaji', 'tipi 4', 'parashikimi', 'institucioni'])
#         for root, dirs, files in os.walk(directory):
#             for basename in files:
#                 if fnmatch.fnmatch(basename, pattern):
#                     filename = os.path.join(root, basename)
#                     workbook = xlrd.open_workbook(filename)
# #                     print(filename)
# #                     workbook.l



def combine_excel_files(directory,pattern):
    counter = 0
    book = xlwt.Workbook(encoding="utf-8")

    sheet1 = book.add_sheet("Sheet 1")
    for root, dirs, files in os.walk(directory):
        for basename in files:
            if fnmatch.fnmatch(basename, pattern):
                filename = os.path.join(root, basename)
                workbook = xlrd.open_workbook(filename)
                worksheet = workbook.sheet_by_index(0)
                umbrella_headers = ['empty1','empty2','empty3','empty4','empty5','empty6','empty7','empty8','empty9', 'tipi 2', 'groupname', 'parashikim', 'kolona 1', 'kolona 5', 'departamenti', 'instituti', 'data', 'institucioni', 'tipi 1', 'tipi 3', 'tipi 5', 'viti', 'muaji', 'parashikimi', 'shuma', 'tipi 4', 'tipi 6', '4.0', 'grupi', 'budget of kosovo - planning', 'kolona 2', 'komuna', 'kolona 3']
                for u_index in range(0,len(umbrella_headers)):
                    sheet1.write(0,u_index,umbrella_headers[u_index])


                headers = grab_column_headers(workbook)

                headers = [x.lower() for x in headers]
                headers = [y.strip() for y in headers]
                print(headers)

                values = []
                for row in range(1, worksheet.nrows):
                    counter = 0
                    col_value = []
                    for col in range(worksheet.ncols):

                        value  = worksheet.cell(row,col).value
                        if headers[col] == '':
                            counter = counter + 1
                            headers_val = 'empty'+ str(counter)
                        else:
                            headers_val = headers[col]
                        column_nr = umbrella_headers.index(headers_val)
                        print(row)
                        print(column_nr)
                        print(value)
                        sheet1.write(row,column_nr, value)
                        try : value = str(value)
                        except : pass
                        col_value.append(value)
                    values.append(col_value)
                print(values)
                book.save("combined.xls")


def grab_column_headers(workbook):
    headers = []
    worksheet = workbook.sheet_by_index(0)
    for idx, cell_obj in enumerate(worksheet.row(0)):
        headers.append(str(cell_obj.value))
    return headers


#
# find_files(mypath, '*.xlsx')
#
# create_umbrella_model(column_array)

combine_excel_files(mypath, '*.xlsx')



