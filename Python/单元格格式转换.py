import re

def convert_to_correct_excel_formula(input_str):
    # 去掉输入字符串开头的 "="，并按非字母数字字符分割字符串
    cleaned_input = input_str[1:]
    parts = re.split(r'(\W+)', cleaned_input)  # 使用原始字符串

    # 从部分列表中移除空字符串
    parts = [part for part in parts if part.strip() != '']

    # 形成 Excel 公式
    formula_parts = []
    i = 0
    while i < len(parts):
        if parts[i].isalnum() and not parts[i].isalpha():  # 只有字母数字组合不使用引号
            formula_parts.append(parts[i])
        else:  # 其他情况（如函数名和运算符）放在引号内
            # 合并连续的非单元格部分
            combined_part = parts[i]
            i += 1
            while i < len(parts) and (not parts[i].isalnum() or parts[i].isalpha()):
                combined_part += parts[i]
                i += 1
            formula_parts.append(f'"{combined_part}"')
            continue
        i += 1

    formula = '=CONCATENATE(' + ','.join(formula_parts) + ')'
    return formula

# 示例使用
input_str = "=IF((M12+M14+M16-M18)<M57,M62/M57*(M12+M14+M16-M18),M62)"
output = convert_to_correct_excel_formula(input_str)
print(output)
