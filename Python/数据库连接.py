import pyodbc

# 数据库连接配置
server = 'PC-202111111555\MSSQLSERVER1'  # 或者使用你的本地数据库服务器名称
database = 'testDB'
username = 'sa'
password = 'sasa'
driver = '{ODBC Driver 17 for SQL Server}'  # 这是 SQL Server 的 ODBC 驱动


# 构建连接字符串
conn_str = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'

try:
    # 建立数据库连接
    conn = pyodbc.connect(conn_str)

    # 创建一个游标
    cursor = conn.cursor()

    # 执行 SQL 查询
    cursor.execute('SELECT Name, DisplayCondition, code FROM SlrRelateItem')

    # 获取查询结果
    rows = cursor.fetchall()

    # 输出查询结果
    for row in rows:
        print(f"Name: {row.Name}, DisplayCondition: {row.DisplayCondition}, Code: {row.code}")

except Exception as e:
    print(f"Error: {e}")

finally:
    # 关闭连接
    conn.close()




