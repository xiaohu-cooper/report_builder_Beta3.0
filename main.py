# 这是一个示例 Python 脚本。

# 按 ⌃R 执行或将其替换为您的代码。
# 按 双击 ⇧ 在所有地方搜索类、文件、工具窗口、操作和设置。
import queue
import time

from Report import *

q = queue.Queue()
start_pro = 10


def mark(walk):
    global start_pro
    start_pro += walk
    q.put(start_pro)



def main(xlsx_name: str, docx_name: str, results_path: str, model: str, error_dict: dict):
    global start_pro

    q.put(start_pro)
    time.sleep(0.1)

    dictlist = Report.xlsx_to_dictlist(xlsx_name)
    walk_1 = (100 - start_pro) // (2 * len(dictlist))
    if model == 'METER':
        for dic in dictlist:
            meter = MeterReport(xlsx_name, docx_name, results_path, dic)
            mark(walk_1)
            meter.value_to_str()
            meter.image_deal_to_dic()
            meter.render_to_save()
            error_dict.update(meter.error)
            mark(walk_1)
    if model == 'CT':
        docx_names_list = docx_name.split(';')  # 6种电流模板分别是0~5
        docx_names_list.sort()  # 和上一行必须分写两行
        for dic in dictlist:
            ct = CTReport(xlsx_name, docx_names_list[0], results_path, dic)
            mark(walk_1)
            ct.xlsx_value_to_dic()
            ct.value_to_str()
            ct.docx = DocxTemplate(docx_names_list[ct.r_number - 1])  # 更新docx
            ct.image_deal_to_dic()
            ct.render_to_save()
            error_dict.update(ct.error)
            mark(walk_1)
    if model == 'PT':
        for dic in dictlist:
            if not flag:
                break
            pt = PTReport(xlsx_name, docx_name, results_path, dic)
            mark(walk_1)
            pt.xlsx_value_to_dic()
            pt.value_to_str()
            pt.image_deal_to_dic()
            pt.render_to_save()
            error_dict.update(pt.error)
            mark(walk_1)
    time.sleep(0.1)
    start_pro = 100
    q.put(start_pro)
