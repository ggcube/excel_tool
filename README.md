# excel 小工具 by Python
## Python 相依套件
請先安裝```openpyxl, xlrd```等套件
## 前置作業
1. 點擊```初始化.bat```會出現```input```及```output```資料夾。
2. 將要處理的資料放在```input```資料夾內。
3. 輸出檔案將會輸出在```output```資料夾。
## 小工具說明
### xls轉xlsx 
#### 功能
將```.xls```檔全部轉成```.xlsx```檔，但格式會跑掉。
#### 使用方法
點擊```xls轉xlsx.bat```即可，輸出資料在```output```資料夾。
### xlsx資料彙整
#### 功能
將所有```xlsx```檔的內容彙整成一個```xlsx```檔。
#### 注意事項
每個```xlsx```檔的工作表只能有一個，第二個之後會讀取不到。
#### 使用方法
點擊```xlsx資料彙整.bat```即可，輸出資料在```output```資料夾。