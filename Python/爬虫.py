import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

# 目标网址
url = 'https://www.mcmod.cn/modlist.html'

# 发送 HTTP 请求
response = requests.get(url)

# 解析网页
soup = BeautifulSoup(response.text, 'html.parser')

# 存储模组名称和英文名称的列表
mod_names = []
mod_enames = []

# 查找并存储模组的中文名称和英文名称
for title_div in soup.find_all('div', class_='title'):
    mod_name = title_div.find('p', class_='name')
    mod_ename = title_div.find('p', class_='ename')

    # 检查是否找到相应的标签
    if mod_name and mod_ename:
        mod_names.append(mod_name.text.strip())
        mod_enames.append(mod_ename.text.strip())

# 创建 DataFrame
df = pd.DataFrame({
    '中文名称': mod_names,
    '英文名称': mod_enames
})

# 获取桌面路径
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# 保存为 Excel 文件
df.to_excel(os.path.join(desktop_path, 'mods_list.xlsx'), index=False)
