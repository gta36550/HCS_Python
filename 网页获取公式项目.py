# 导入必要的库
import json
import re
import pandas as pd
from tkinter import Tk, Text, Button, Label, END, filedialog, mainloop

# 定义函数，从输入的文本中提取DisplayCondition信息
def get_display_conditions(txt_input):
    conditions = {}  # 初始化一个字典来存储DisplayCondition信息
    text = txt_input  # 获取文本输入

    # 定义内部函数，将16进制Unicode字符串转换为对应的字符
    def decode_match(match):
        return chr(int(match.group(1), 16))

    # 使用正则表达式替换文本中的Unicode编码为实际字符
    text = re.sub(r'\\u([0-9a-fA-F]{4})', decode_match, text)
    print("已完成Text内容的转换")

    # 在处理过的文本中查找所有的json格式字符串
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
    print("已完成DisplayCondition信息的获取")
    return conditions

# 定义函数，用于获取大量输入，并处理保存结果
def get_large_input():
    def return_input_and_close():
        # 获取json_text和txt_text的内容
        json_content = json_text.get(1.0, END)
        txt_content = txt_text.get(1.0, END)
        print("已获取JSON和Text内容")

        # 关闭GUI窗口
        root.destroy()
   
        # 解析JSON内容，并使用数据构造DataFrame
        data = json.loads(json_content)
        df = pd.DataFrame(data['rows'])
        print("已完成JSON内容的处理")

        # 使用文本内容获取DisplayCondition信息，并映射到DataFrame
        display_conditions = get_display_conditions(txt_content)
        df['DisplayCondition'] = df['Name'].map(display_conditions)
        # 去除'code'列的前导零，并只保留指定的三列
        df['code'] = df['code'].astype(str).str.lstrip('0')
        df = df[['Name', 'DisplayCondition', 'code']]
        print("已完成DataFrame的处理")
        
        # 弹出文件保存对话框，保存结果为Excel文件
        f = filedialog.asksaveasfile(mode='w', defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if f is None:  # 如果用户取消保存，则退出函数
            return
        df.to_excel(f.name, index=False)
        print("已将DataFrame保存为xlsx文件")

    # 创建GUI窗口
    root = Tk()

    # 添加JSON字符串输入的标签和文本框
    json_label = Label(root, text="请输入JSON字符串")
    json_label.pack()
    json_text = Text(root)
    json_text.pack()

    # 添加TXT内容输入的标签和文本框
    txt_label = Label(root, text="请输入TXT内容")
    txt_label.pack()
    txt_text = Text(root)
    txt_text.pack()
    
    # 添加确定按钮，点击后执行return_input_and_close函数
    button = Button(root, text='确定', command=return_input_and_close)
    button.pack()

    # 进入GUI事件循环
    mainloop()

# 运行程序
get_large_input()