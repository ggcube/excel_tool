import xlrd, openpyxl, os
count = 0
for input_file in os.listdir('input'):
    xlsBook = xlrd.open_workbook(os.path.join('input',input_file))
    workbook = openpyxl.Workbook()
    for i in range(0, xlsBook.nsheets):
        xlsSheet = xlsBook.sheet_by_index(i)
        sheet = workbook.active if i == 0 else workbook.create_sheet()
        sheet.title = xlsSheet.name
        for row in range(0, xlsSheet.nrows):
            for col in range(0, xlsSheet.ncols):
                sheet.cell(row=row+1, column=col+1).value = xlsSheet.cell_value(row, col)
    workbook.save(os.path.join('output',input_file.replace('.xls','.xlsx')))
    count += 1
print('轉換結束，共完成',count,'個xls檔案轉換為xlsx檔案。')
