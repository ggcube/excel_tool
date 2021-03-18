# 生成 input, output 資料夾
import os
if not(os.path.isdir('output')):
    os.makedirs('output')
if not(os.path.isdir('input')):
    os.makedirs('input')