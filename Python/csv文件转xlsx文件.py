import pandas as pd
from tkinter import filedialog
from tkinter import Tk

# 创建一个Tkinter的窗口对象
root = Tk()
# 隐藏这个窗口
root.withdraw()

# 打开一个文件选择对话框，让用户选定一个csv文件
file_path = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')])

if file_path:  # 如果用户选定了一个文件
    # 读取csv文件
    df = pd.read_csv(file_path)

    # 选择需要的列
    df = df[['Name', 'DisplayCondition', 'code']]

    # 创建xlsx文件名（与csv文件同名，路径也相同）
    xlsx_file_path = file_path.rsplit('.', 1)[0] + '.xlsx'
    
    # 将数据写入xlsx文件
    df.to_excel(xlsx_file_path, index=False)

    print(f'已将数据写入 {xlsx_file_path}')