"""
低级方案实现
"""
import traceback

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

home_address = "湖北省武汉市"
name = "duoduo"
age = "18"
date = "2025年5月"

substitution_text = [home_address,name,age,date]
placeholder = "xxxxxx"
filename = "lower_plan_test_text"

local_file_path = "../docs/" + filename + ".docx"
new_file_path = "../test/" + filename + "_result.docx"
try:
    # 将docs中的docx文档初始化为Document对象
    doc = Document(local_file_path)

    # 将docx文档切分paragraphs 和 runs
    substitution_text_order = iter(substitution_text)  # 将需要替换的信息放入迭代器中,以便后续按序替换占位符
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, next(substitution_text_order, placeholder))

    doc.save(new_file_path)
#     文件不存在、权限问题、文档损坏
except PackageNotFoundError:
    print(f"错误: 文件不存在或不是有效的docx文件: {local_file_path}")
    # 返回错误代码或抛出异常
except ValueError as ve:
    print(f"占位符替换错误: {str(ve)}")
except Exception as e:
    traceback.print_exc()
    print(f"未知错误: {str(e)}")
