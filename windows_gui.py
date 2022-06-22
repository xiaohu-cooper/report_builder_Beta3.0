"""
windows_gui- 图形界面

Author:肖虎
Date:2022/6/4
"""
import queue
import threading

import PySimpleGUI as sg

from Report import pop_up
from main import main, q

# 修改初始路径
default_meter_xlsx = 'resources/电能表/电能表项目信息（XX项目）.xlsx'
default_meter_docx = 'resources/电能表/电能表模板.docx'
default_CT_xlsx = 'resources/电流/电流项目信息（XX项目）.xlsx'
default_CT_docxs = 'resources/电流/1_电流模板.docx;resources/电流/2_电流模板.' \
                   'docx;resources/电流/3_电流模板.docx;resources/电流/4_电流模板.docx;' \
                   'resources/电流/5_电流模板.docx;resources/电流/6_电流模板.docx'
default_PT_xlsx = 'resources/电压/电压项目信息（XX项目）.xlsx'
default_PT_docx = 'resources/电压/电压模板.docx'
default_results_path = 'results'

ls = ['电能表报告', '电流互感器报告', '电压互感器报告']
icon = b'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'

sg.set_global_icon(icon)
sg.theme('LightGray')
sg.theme_input_background_color('white')
sg.theme_input_text_color('black')


# 创建主窗口函数 tkinter只能这样创建多窗口，其他的报错
def make_window1():
    layout = [
        [sg.Text('')],
        [sg.Text('选择项目信息（.xlsx文件）', font=('楷体', 11))],
        [sg.FileBrowse('选择一个文件', size=14, target='-XLSX-', file_types=(('Excel文件', '.xlsx'),), font=('楷体', 11)),
         sg.In(default_meter_xlsx, key='-XLSX-', font=('楷体', 11))],

        [sg.Text('')],
        [sg.Text('选择报告模板（.docx文件）', font=('楷体', 11))],
        [sg.FileBrowse('选择一个文件', size=14, target='-DOCX-', key='-DOCX_FB-', file_types=(('Word文件', '.docx'),), font=('楷体', 11)),
         sg.In(default_meter_docx, key='-DOCX-', disabled_readonly_background_color='LightGray', font=('楷体', 11))],

        [sg.Text('')],
        [sg.Text('选择存放文件的文件夹', font=('楷体', 11))],
        [sg.FolderBrowse('选择一个文件夹', size=14, target='-PATH-', font=('楷体', 11)),
         sg.In(default_results_path, key='-PATH-', font=('楷体', 11))],

        [sg.Text('')],
        [sg.Text('选择生成模式', font=('楷体', 11))],
        # 单选按钮 group_id=1
        [sg.R(ls[0], group_id=1, default=True, enable_events=True, key='-Radio11-', font=('楷体', 11)),
         sg.R(ls[1], group_id=1, enable_events=True, key='-Radio12-', font=('楷体', 11)),
         sg.R(ls[2], group_id=1, enable_events=True, key='-Radio13-', font=('楷体', 11))],

        # 编辑电流互感器选项
        [sg.Frame('CT选项', title_color='DarkGray', key='-CT_menu-', relief=sg.RELIEF_SUNKEN, visible=False, layout=[
            [sg.Text('选择电流互感器报告模板（默认是六个）', font=('楷体', 11))],
            [sg.FilesBrowse('选择文件', size=10, target='-DOCXS-', file_types=(('Word文件', '.docx'),), font=('楷体', 11)),
             sg.In(default_CT_docxs, key='-DOCXS-', font=('楷体', 11))]
        ], font=('楷体', 11))],

        [sg.Text('')],
        [sg.Button('确认生成', size=(8, 1), font=('楷体', 11)), sg.Button('取消', size=(8, 1), font=('楷体', 11))],
    ]
    return sg.Window("报告生成器", layout, font=10, finalize=True)


# 创建副窗口函数
def make_window2():
    layout = [
        [sg.Text(f'完成进度0%', key='-P-', font=('楷体', 11))],
        [sg.ProgressBar(100, orientation='h', size=(70, 20), key='-BAR-', bar_color='#819AFC')],
        [sg.Cancel('关闭', pad=((300, 300), (20, 20)), font=('楷体', 11))]
    ]
    return sg.Window("进度", layout, font=10, finalize=True)


window1 = make_window1()
button_color = window1['-DOCX_FB-'].ButtonColor

while True:
    event, value = window1.read()
    if event is None or event == '取消':
        break

    if event == '-Radio12-':
        window1['-CT_menu-'].update(visible=True)
        window1['-XLSX-'].update(default_CT_xlsx)
        window1['-DOCX-'].update('', disabled=True)
        window1['-DOCX_FB-'].update(button_color='LightGray', disabled=True)
    if event == '-Radio11-':
        window1['-CT_menu-'].update(visible=False)
        window1['-XLSX-'].update(default_meter_xlsx)
        window1['-DOCX-'].update(default_meter_docx, disabled=False)
        window1['-DOCX_FB-'].update(button_color=button_color, disabled=False)
    if event == '-Radio13-':
        window1['-CT_menu-'].update(visible=False)
        window1['-XLSX-'].update(default_PT_xlsx)
        window1['-DOCX-'].update(default_PT_docx, disabled=False)
        window1['-DOCX_FB-'].update(button_color=button_color, disabled=False)

    if event == '确认生成':
        xlsx_name = value['-XLSX-']
        docx_name = value['-DOCX-']
        results_path = value['-PATH-']
        docx_names = value['-DOCXS-']
        if value['-Radio11-']:
            model = 'METER'
        elif value['-Radio12-']:
            # sg.Popup('还没研究呢，等着……')
            model = 'CT'
            docx_name = docx_names
        else:
            # sg.Popup('说完了还乱点？！')
            model = 'PT'
        error_dict = {}

        window2 = make_window2()
        bar = window2['-BAR-']
        window2.make_modal()    # 关闭前不能与其他窗口交互
        window1.disable()
        window1['确认生成'].update(button_color='LightGray', disabled=True)    # window1禁止用户输入，确认生成按钮变为灰色

        thread1 = threading.Thread(target=main, args=[xlsx_name, docx_name, results_path, model, error_dict])
        thread1.setDaemon(True)
        thread1.start()

        # 创建主线程循环，直到子线程结束
        while threading.active_count() >= 2:
            # 进度条窗口的循环
            while not window2.is_closed():
                event_1, value_1 = window2.read(100)
                # print(event_1, end=' ')
                if event_1 is None:
                    break
                if event_1 == '关闭':
                    window2.minimize()
                try:
                    progress_value = q.get_nowait()
                    # print(progress_value)
                except queue.Empty:
                    continue
                else:
                    window2['-P-'].Update(value=f'完成进度{progress_value}%')    # 更新进度条显示文字
                    bar.update(progress_value)   # 更新进度条
                    window2.refresh()
                    if progress_value >= 100:
                        break
            if not window2.is_closed():
                window2.close()

        window1.ding()   # 完成后，弹窗前响铃
        pop_up(error_dict, results_path, icon)
        window1.enable()
        window1['确认生成'].update(button_color=button_color, disabled=False)   # 恢复按钮颜色

window1.close()
