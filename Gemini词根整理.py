import re
import openpyxl
from openpyxl.styles import Font, Alignment
import os # 用于检查文件是否存在

def parse_markdown_to_roots_data(markdown_content):
    """
    使用正则表达式解析 Markdown 文本内容，提取词根数据。

    Markdown 格式预期为:
    1.  **RootName**
        * 语言：Language Info
        * 释义：Meaning Info
        * 词根形式：Forms Info (可选)

    返回: 包含提取数据的字典列表。
    """
    extracted_data = []
    # 正则表达式：
    # - 匹配序号、点、空格、**词根**
    # - 捕获“语言”后的内容
    # - 捕获“释义”后的内容
    # - 可选捕获“词根形式”后的内容
    # re.MULTILINE 使得 ^ 和 $ 匹配每行的开始和结束
    # re.DOTALL 使得 . 可以匹配换行符 (虽然这里主要靠非贪婪匹配和明确的换行符 \n)
    # 使用命名捕获组 (?P<name>...)
    pattern = re.compile(
        r"^\s*\d+\.\s+\*\*(?P<root>.+?)\*\*\s*?\n"
        r"\s*\*\s*语言：\s*(?P<language>.+?)\s*?\n"
        r"\s*\*\s*释义：\s*(?P<meaning>.+?)\s*?\n"
        r"(?:\s*\*\s*词根形式：\s*(?P<forms>.+?)\s*?\n)?" # 词根形式行为可选
        , re.MULTILINE
    )

    for match in pattern.finditer(markdown_content):
        data = match.groupdict()
        # 清理捕获到的数据，并处理可选组可能为 None 的情况
        root = data.get("root", "").strip()
        language = data.get("language", "").strip()
        meaning = data.get("meaning", "").strip()
        forms = data.get("forms", "").strip() if data.get("forms") else "" # 如果forms为None则为空字符串

        if root: # 确保至少捕获到了词根
            extracted_data.append({
                "词根": root,
                "语言": language,
                "释义": meaning,
                "词根形式": forms
            })
            
    return extracted_data

def create_excel_from_data(data_list, excel_filename="词根列表.xlsx"):
    """
    使用 openpyxl 将数据列表创建为 Excel 文件。
    """
    if not data_list:
        print("没有数据可写入 Excel。")
        return

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "词根数据"

    # 写入表头并设置样式
    headers = ["词根", "语言", "释义", "词根形式"]
    sheet.append(headers)
    for cell in sheet[1]: # 获取第一行的所有单元格 (表头)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # 写入数据行
    for item in data_list:
        row_data = [
            item.get("词根", ""),
            item.get("语言", ""),
            item.get("释义", ""),
            item.get("词根形式", "")
        ]
        sheet.append(row_data)

    # 调整列宽 (可选，但能改善可读性)
    for col_idx, column_letter in enumerate(['A', 'B', 'C', 'D'], 1):
        max_length = 0
        column = sheet[column_letter]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) if max_length > 0 else 15 # 最小宽度15
        sheet.column_dimensions[column_letter].width = adjusted_width


    # 保存文件
    try:
        workbook.save(excel_filename)
        print(f"成功将 {len(data_list)} 条数据写入到 '{excel_filename}'")
    except Exception as e:
        print(f"保存 Excel 文件 '{excel_filename}' 时发生错误: {e}")
        print("请确保您有写入该位置的权限，并且文件名是合法的。")

def main():
    md_filename_input = input("请输入包含词根数据的 Markdown 文件名 (例如: my_roots.md): ").strip()
    if not md_filename_input:
        print("错误：未提供 Markdown 文件名。")
        return
    
    # 检查 Markdown 文件是否存在
    if not os.path.exists(md_filename_input):
        print(f"错误: Markdown 文件 '{md_filename_input}' 不存在。请检查文件名和路径。")
        return

    excel_filename_input = input("请输入您希望保存的 Excel 文件名 (例如: output_roots.xlsx) [默认为: 词根列表.xlsx]: ").strip()
    if not excel_filename_input:
        excel_filename_to_save = "词根列表.xlsx"
    else:
        if not excel_filename_input.lower().endswith(".xlsx"):
            excel_filename_to_save = excel_filename_input + ".xlsx"
        else:
            excel_filename_to_save = excel_filename_input
    
    try:
        with open(md_filename_input, 'r', encoding='utf-8') as f:
            markdown_content = f.read()
    except Exception as e:
        print(f"读取 Markdown 文件 '{md_filename_input}' 时发生错误: {e}")
        return

    roots_data = parse_markdown_to_roots_data(markdown_content)

    if roots_data:
        create_excel_from_data(roots_data, excel_filename_to_save)
    else:
        print(f"未能从 '{md_filename_input}' 中解析到任何词根数据。请检查文件内容是否符合预期的 Markdown 格式。")
        print("预期的 Markdown 格式示例：")
        print("1.  **词根名称**")
        print("    * 语言：语言信息")
        print("    * 释义：释义信息")
        print("    * 词根形式：形式信息 (此行可选)")

if __name__ == "__main__":
    main()