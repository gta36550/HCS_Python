# 导入所需的库
import tkinter as tk
import pandas as pd
import re
import chardet
from tkinter import filedialog
from tkinter import Tk

def extract_and_fill_chinese_words(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)

    #获取列数
    num_columns = len(df.columns)

    # 如果列数小于7，就添加缺失的列
    while num_columns < 8:
        df.insert(loc=num_columns, column=f'列{num_columns + 1}', value='')    # 列名为"列X"
        num_columns += 1  # 更新列数

    # 重命名列名
    df.rename(columns={df.columns[2]: '引用项目',
                    df.columns[3]: '项目编号',
                    df.columns[4]: '公式序号',
                    df.columns[5]: '原序号',
                    df.columns[6]: '是否一致',
                    df.columns[7]: '除数检查'}, inplace=True)

    # 检查F列的每个元素是否都为空字符串
    if all(df[df.columns[5]] == ''):
        # 如果F列全部由空字符串组成，则将C列的数据复制到F列
        df[df.columns[5]] = df[df.columns[2]]

    # 提取第B列中的中文词语（带数字的也算，会去掉纯数字和与单引号相邻的词），并填写到第C列，同时去重
    pattern = r'\b[\u4e00-\u9fa5]*\d+[\u4e00-\u9fa5]*\b|[\u4e00-\u9fa5]+'
    df[df.columns[2]] = df[df.columns[1]].apply(
        lambda x: ' '.join(
            set(word for word in re.findall(pattern, str(x)) 
                if not (word.isdigit() or re.search(fr"'{word}", str(x)) or re.search(fr"{word}'", str(x)))
            )
        )
    )

    # 去掉C列中与A列相同的字符串
    df[df.columns[2]] = df.apply(lambda row: ' '.join(set(row[df.columns[2]].split()) - set(row[df.columns[0]].split())), axis=1)

    # 是否检查C列中是否和A列中含有对应的字符串，填写D列
    df[df.columns[3]] = df.apply(lambda row: ', '.join(['是' if any(word in a_column for a_column in df[df.columns[0]].dropna()) else '否' for word in str(row[df.columns[2]]).split()]), axis=1)
    
    # 将E列全部填写为“是”
    df[df.columns[4]] = "是"

    # 循环检查E列是否包含“是”
    while '是' in df[df.columns[4]].values:
        # 检查C列中是否有“是”和数字，有“是”则在E列填“是”，没有则检查是否有数字，有数字填最大数字+1，都没有则填1
        df[df.columns[4]] = df.apply(lambda row: '是' if '是' in row[df.columns[3]].split(', ') else str(max([int(num) for num in re.findall(r'\b\d+\b', row[df.columns[3]])], default=0) + 1), axis=1)

        # 将其对应的A列那一行E列的数字填写到C列那一行的D列，并用逗号分隔。如果没有匹配到，就填写字符串“否”
        df[df.columns[3]] = df.apply(lambda row: ', '.join(
            [str(df.loc[df[df.columns[0]] == word, df.columns[4]].values[0])
                if any(word == a_column for a_column in df[df.columns[0]].dropna()) else '否'
                for word in str(row[df.columns[2]]).split()]), 
            axis=1)

    # 对E列和F列进行检查，如果E列不等于F列，则G列填写不一致
    df[df.columns[6]] = df.apply(lambda row: '不一致' if int(row[df.columns[4]]) != int(row[df.columns[5]]) else '', axis=1)

    # 检查G列的每个元素是否都为空字符串
    if all(df[df.columns[6]] == ''):
        # 如果G列全部由空字符串组成，则G列全部填写“全部一致”
        df[df.columns[6]] = '全部一致'
        print('序号全部一致')

    # 该函数检查被除数是否有大于0的检查
    def check_conditions(text):
        errors = []  # 创建一个空列表，用于保存错误
        lines = text.split('\n')
        for i in range(1, len(lines)):  # 从第二行开始
            match = re.search(r'/ (\w+)', lines[i])
            if match:  # 如果找到了 "/ 字符串" 格式
                word = match.group(1)  # 获取字符串
                if not word.isdigit():  # 如果字符串不是数字
                    # 检查上一行是否包含 "字符串 > 0"
                    if f'{word} > 0' not in lines[i-1]:
                        errors.append(f'第{i+1}行: {word}缺少检查')  # 将错误添加到列表中
        return ',\n'.join(errors) if errors else ''  # 整理错误信息并返回

    # 进行被除数检查
    df[df.columns[7]] = df[df.columns[1]].apply(check_conditions)

    # 检查H列的每个元素是否都为空字符串
    if all(df[df.columns[7]] == ''):
        # 如果H列全部由空字符串组成，则G列全部填写“全部通过”
        df[df.columns[7]] = '全部通过'
        print('被除数检查全部通过')

    # 保存修改后的内容
    df.to_excel(file_path, index=False)

def choose_file():
    # 创建一个Tkinter的窗口对象
    root = Tk()
    # 隐藏这个窗口
    root.withdraw()

    # 打开一个文件选择对话框，让用户选择csv文件或xlsx文件
    file_path = filedialog.askopenfilename(filetypes=[
        ('CSV & Excel Files', '*.csv; *.xlsx'),
        ('CSV Files', '*.csv'),
        ('XLSX Files', '*.xlsx')
    ])

    # 如果用户选定了一个文件
    if file_path:  
        # 根据文件后缀处理
        extension = file_path.rsplit('.', 1)[1]

        # 如果选择csv文件
        if extension == "csv":
            print('已选择csv文件')
            # 读取csv文件
            try:
                rawdata = open(file_path, 'rb').read()
                result = chardet.detect(rawdata)
                encoding = result['encoding']
                df = pd.read_csv(file_path, encoding=encoding)
            except UnicodeDecodeError:
                print(f'无法用{encoding}编码读取文件 {file_path}，请尝试其他编码。')
            except Exception as e:
                print(f'读取文件 {file_path} 时发生错误：', e)

            # 选择需要的列
            df = df[['Name', 'DisplayCondition', 'code']]

            # 创建xlsx文件名（与csv文件同名，路径也相同）
            xlsx_file_path = file_path.rsplit('.', 1)[0] + '.xlsx'

            # 将数据写入xlsx文件
            try:
                df.to_excel(xlsx_file_path, index=False)
                print(f'已生成xlsx文件')
            except PermissionError:
                print(
                    f'无法写入文件 {xlsx_file_path}。可能是由于文件正在被另一个程序使用，或者你没有写入该文件的权限。')
            except Exception as e:
                print(f'在尝试写入文件 {xlsx_file_path} 时发生未知错误: ', e)

        # 如果是xlsx文件
        elif extension == "xlsx":
            print('已选择xlsx文件')
            # 直接使用xlsx文件路径
            xlsx_file_path = file_path

        # 新增代码 - 在这里添加对新生成xlsx文件的操作
        try:
            df_xlsx = pd.read_excel(xlsx_file_path)
        except FileNotFoundError:
            print(f'文件 {xlsx_file_path} 未找到。')
        except PermissionError:
            print(f'无法读取文件 {xlsx_file_path}。可能是由于文件正在被另一个程序使用，或者你没有打开该文件的权限。')
        except Exception as e:
            print(f'在尝试读取文件 {xlsx_file_path} 时发生错误：', e)

    # 对df_xlsx进行操作
    try:
        if file_path:
            extract_and_fill_chinese_words(xlsx_file_path)
            print("已完成处理")
    except Exception as e:
        print("出现错误：", e)

# 主程序
if __name__ == "__main__":
    # 手动选择文件并处理
    choose_file()