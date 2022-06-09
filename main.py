# 这是一个示例 Python 脚本。

# 按 ⌃R 执行或将其替换为您的代码。
# 按 双击 ⇧ 在所有地方搜索类、文件、工具窗口、操作和设置。

from Report import *


def main(xlsx_name: str, docx_name: str, results_path: str, model: str) -> dict:

    error_dict = {}
    dictlist = Report.xlsx_to_dictlist(xlsx_name)
    docx_names_list = []
    if model == '电能表':
        for dic in dictlist:
            meter = MeterReport(xlsx_name, docx_name, results_path, dic)
            meter.value_to_str()
            meter.image_deal_to_dic()
            meter.render_to_save()
            error_dict.update(meter.error)
    if model == 'CT':
        docx_names_list = docx_name.split(';')   # 6种电流模板分别是0~5
        docx_names_list.sort()
        for dic in dictlist:
            ct = CTReport(xlsx_name, docx_names_list[0], results_path, dic)
            ct.xlsx_value_to_dic()
            ct.docx = DocxTemplate(docx_names_list[ct.r_number - 1])  # 更新docx
            ct.image_deal_to_dic()
            ct.render_to_save()
            error_dict.update(ct.error)
    if model == 'PT':
        for dic in dictlist:
            pt = PTReport(xlsx_name, docx_name, results_path, dic)
            pt.xlsx_value_to_dic()
            pt.image_deal_to_dic()
            pt.render_to_save()
            error_dict.update(pt.error)
    return error_dict
