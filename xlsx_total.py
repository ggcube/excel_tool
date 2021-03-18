import os, sys, openpyxl, logging, datetime
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
# logging.disable(logging.DEBUG)
logging.debug('程式開始～')
# 參數設定
start_row = 4 # 讀取資料的開始列
title_row = start_row - 1 # 彙整資料的標題列數
# 設置計數器
write_row = 1 # 寫入列號flag
data_file_count = 0 # 讀取檔案數量
# 預設資料都放置在 input 資料夾內
data_dir = 'input'
# 開啟新的活頁簿
workbook = openpyxl.Workbook()
sheet = workbook.active
for xlsxfile in os.listdir(data_dir):
    data_file_count += 1
    print('讀取第',data_file_count,'個檔案。')
    data_file = os.path.join(data_dir,xlsxfile) # 組成檔案路徑
    data_workbook = openpyxl.load_workbook(data_file)
    data_sheet = data_workbook.active
# 寫入標題列,只有第1個檔案執行
    if data_file_count == 1:
        print('正在寫入標題...')
        for row in range(1,title_row+1):
            for col in range(1,data_sheet.max_column+1):
                sheet.cell(row = row,column= col).value = data_sheet.cell(row = row, column = col).value
        write_row = start_row
# 寫入資料
    for row in range(start_row,data_sheet.max_row+1):
        for col in range(1,data_sheet.max_column+1):
            sheet.cell(row = write_row, column = col).value = data_sheet.cell(row = row,column = col ).value
        write_row += 1
    print('共寫入',data_sheet.max_row-(start_row-1),'筆資料。')
workbook.save(os.path.join('output','資料彙整'+datetime.datetime.now().strftime('%Y%m%d-%H%M%S')+'.xlsx'))
os.system('pause')
    



