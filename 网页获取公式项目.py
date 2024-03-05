import json
import re
import chardet
import pandas as pd
from tkinter import filedialog
from tkinter import Tk

# 探测文件编码
def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    return chardet.detect(raw_data)['encoding']

# 定义一个函数，用于从TXT文件中提取DisplayCondition信息
def get_display_conditions(txt_file_path):
    conditions = {}

    with open(txt_file_path, 'r', encoding='utf-8') as file:
        text = file.read()

        # 直接在这里解码unicode转义序列
        def decode_match(match):
            return chr(int(match.group(1), 16))

        text = re.sub(r'\\u([0-9a-fA-F]{4})', decode_match, text)

        json_matches = re.findall(r'\[\{.*?\}\]', text, re.DOTALL)
        for json_str in json_matches:
            try:
                data = json.loads(json_str)
                for item in data:
                    name = item.get('Name')
                    display_condition = item.get('DisplayCondition', '')
                    conditions[name] = display_condition
            except json.JSONDecodeError as e:
                print(f'JSON解析错误: {e}')
                continue

    return conditions

# 使用Tkinter创建一个应用程序窗口并立即隐藏
root = Tk()
root.withdraw()

# 弹出文件对话框，让用户选择JSON文件
json_file_path = filedialog.askopenfilename(
    title='请选择JSON文件',
    filetypes=[('JSON文件', '*.json')]
)

# 弹出文件对话框，让用户选择TXT文件
txt_file_path = ''
if json_file_path:  # 如果选中了JSON文件
    txt_file_path = filedialog.askopenfilename(
        title='请选择TXT文件',
        filetypes=[('TXT文件', '*.txt')]
    )

if json_file_path and txt_file_path:  # 如果两个文件都被选中了
    with open(json_file_path, 'r', encoding='utf-8') as json_file:  # 打开JSON文件
        data = json.load(json_file)  # 加载JSON数据

    df = pd.DataFrame(data['rows'])  # 将JSON数据转换为DataFrame

    display_conditions = get_display_conditions(txt_file_path)  # 获取DisplayCondition映射
    
    # 为DataFrame添加DisplayCondition列
    df['DisplayCondition'] = df['Name'].map(display_conditions)

    # 去除 'code' 列值开头的所有0
    df['code'] = df['code'].astype(str).str.lstrip('0')

    # 筛选出指定的三列
    df = df[['Name', 'DisplayCondition', 'code']]

    # 弹出保存文件对话框，让用户选择保存Excel文件的路径
    xlsx_file_path = filedialog.asksaveasfilename(
        title='保存Excel文件',
        filetypes=[('Excel文件', '*.xlsx')],
        defaultextension='.xlsx'
    )

    if xlsx_file_path:  # 如果用户选择了文件路径
        df.to_excel(xlsx_file_path, index=False)  # 将DataFrame保存到Excel文件
        print(f'数据已成功保存到 "{xlsx_file_path}"。')  # 打印成功消息
    else:
        print("操作已取消，没有选择Excel文件的保存位置。")
else:
    if not json_file_path:
        print("操作已取消，没有选择JSON文件。")
    if not txt_file_path:
        print("操作已取消，没有选择TXT文件。")

# 清理并关闭Tkinter应用程序窗口
root.destroy()