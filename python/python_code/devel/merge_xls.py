#!/usr/bin/env python
import os
import xlwt
import xlrd

def getYear(worksheet):
    c0 = []
    c_ret = []
    idx = []
    idx_ret = []
    for i in range(worksheet.nrows):
        c0.append(worksheet.row_values(i)[0])
        idx.append(i)
    if 'Year' or 'year' in c0:
        for i, val in enumerate(c0):
            #print(val)
            #print(type(val))
            if type(val) == str:
                if val.isdigit():
                    val = int(val)
            elif type(val) == unicode:
                if val.isnumeric():
                    val = float(val)
            if (val > 1900) and (val < 2100):
                #print(val)
                c_ret.append(val)
                idx_ret.append(i)
        return (c_ret, idx_ret)
    else:
        return (c_0, idx)

def getValue(worksheet, idx):
    row = []
    for i in range(worksheet.ncols):
        row.append(worksheet.col_values(i)[idx])
    return row[1:]
def nameCol(row, name):
    col_names = []
    for i,val in enumerate(row):
        retStr = '\\' + name + '-' + str(i+1)
        col_names.append(retStr)
    return col_names
  
def writeRow(ws, row, offset, row_number):
    for col_index, col_value in enumerate(row):
        ws.write(row_number, col_index+offset, col_value)
    

path= '/home/nochi/PCA/python/2013'

wb          = xlwt.Workbook()
ws          = wb.add_sheet('raw_sheet', cell_overwrite_ok=True)

# Style for excel
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
            num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

# Write first column 
for idx, val in enumerate(range(1978,2015)):
    ws.write(1+idx, 0, val)
    
    

sorted_list = sorted(os.listdir(path))
new_list = []
for file in sorted_list:
    if file.endswith("xls"):
        new_list.append(file)

# Each Excel file
offset = 1
for idx, file in enumerate(new_list):
        book = xlrd.open_workbook(path + '/' + file)
        sheet = book.sheet_by_index(0)
        a1 = sheet.cell_value(rowx=0, colx=0)
        splitted_a1 = a1.rsplit()
        name = splitted_a1[0]
        #ws.write(0, idx+1, name)
        years_of_sheet, idx = getYear(sheet)
        first_row = getValue(sheet, 0)
        col_names = nameCol(first_row, name)
        writeRow(ws, col_names, offset, 0)

        #for i, val in enumerate(idx):
        for n, year in enumerate(range(1978,2015)):
            #year_of_row = sheet.cell_value(rowx = n+1, colx=0)
            for k, year_of_sheet in enumerate(years_of_sheet):
                if year_of_sheet == year:
                    row = getValue(sheet, idx[k])
                    writeRow(ws, row, offset, n+1)

        #print(col_names)
        print(col_names)
        print(offset)
        offset += len(col_names)

        #print(col_names[1])
        #print(first_row)
        #print(years)
        #print(len(years))
        #print(idx)

    


wb.save('../raw_workbook_2013.xls')
