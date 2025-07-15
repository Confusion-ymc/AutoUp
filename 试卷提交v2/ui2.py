import pathlib
import threading
import time
from typing import OrderedDict

import PySimpleGUI as sg
from task_manager import task_manager
import tkinter as tk


# sg.theme('SystemDefaultForReal')


def update_result_count(window):
    last_task_ids = []
    print('更新线程 已启动')
    while True:
        time.sleep(0.5)
        task_dict = task_manager.TASK_DICT.copy()
        # 获取当前任务ID集合
        current_task_ids = list(task_dict)
        try:
            # 更新统计数据
            window['-SUCCESS-'].update(len(task_manager.SUCCESS_COUNT))
            window['-FAILED-'].update(len(task_manager.FAILED_COUNT))
            window['-REPEAT-'].update(len(task_manager.REPEAT_COUNT))
            window['-TASK_COUNT-'].update(task_manager.task_queue.qsize())
            window['-RUNNING-'].update(len(task_manager.RUNNING_COUNT))

            new_tasks = OrderedDict()

            # 找出需要更新的行
            rows_to_update = {}
            for row_index, task_id in enumerate(current_task_ids):
                task_info = task_dict[task_id]
                row_data = [
                    row_index + 1,
                    task_info.get('filename', ''),
                    task_info.get('status', ''),
                    task_info.get('thread', ''),
                    task_info.get('start_time').strftime("%H:%M:%S") if task_info.get('start_time') else '',
                    task_info.get('end_time').strftime("%H:%M:%S") if task_info.get('end_time') else '',
                    task_info.get('duration', ''),
                    task_info.get('error', '')
                ]
                if task_info.get('status_changed', False):
                    rows_to_update[row_index] = row_data
                    # 重置状态变化标志
                    task_info['status_changed'] = False
                if row_index >= len(last_task_ids):
                    new_tasks[row_index] = row_data

            table = window['-TASK_TABLE-'].Widget
            # 插入新增的行
            for row_index, row_data in new_tasks.items():
                table.insert("", 'end', iid=row_index, values=row_data)

            # 更新有变化的行
            for row_index, row_data in rows_to_update.items():
                if row_index in new_tasks:
                    continue
                children = table.get_children()
                table.item(children[row_index], values=row_data)
        except Exception as e:
            print(e)
        last_task_ids = current_task_ids.copy()


def main_gui(run_func):
    font = ('微软雅黑', 12)
    # 定义表格列
    headings = ['序号', '文件名', '状态', '执行线程', '开始时间', '结束时间', '耗时(秒)', '错误信息']
    col_widths = [5, 60, 10, 12, 15, 15, 10, 30]
    layout = [
        [sg.Table(
            values=[],
            headings=headings,
            max_col_width=100,
            auto_size_columns=False,
            col_widths=col_widths,
            display_row_numbers=False,
            justification='center',
            num_rows=20,
            key='-TASK_TABLE-',
            row_height=25,
            font=font,
            # enable_click_events=True,
            # tooltip='任务执行情况表',
            expand_x=True,
            expand_y=True
        )],
        [sg.Text('待执行:'), sg.Text('0', key='-TASK_COUNT-'),
         sg.Text('运行中:'), sg.Text('0', key='-RUNNING-'),
         sg.Text('成功数量:'), sg.Text('0', key='-SUCCESS-'),
         sg.Text('重复数量:'), sg.Text('0', key='-REPEAT-'),
         sg.Text('失败数量:'), sg.Text('0', key='-FAILED-'),
         sg.Button('文件夹监听中', key='-LISTEN_DIR-'),
         sg.FolderBrowse('选择上传文件夹', enable_events=True, change_submits=True, key='-CHANGE_LISTEN_DIR-',
                         target='-CHANGE_LISTEN_DIR-'),
         sg.Text(str(task_manager.UPLOAD_DIR.absolute()), key='-LISTEN_DIR_LABEL-'),
         sg.Checkbox("显示浏览器", default=task_manager.show_browser, enable_events=True, key="-SHOW_BROWSER-"),
         sg.Push(),
         sg.Text('线程数量:'), sg.InputText(str(task_manager.THREAD_COUNT), key='-THREAD_COUNT-', size=(5, 1)),
         sg.Button('开始'), sg.Button('退出'), sg.Button('关于')]
    ]
    tooltip = None

    def show_tooltip(event):
        nonlocal tooltip
        table = window['-TASK_TABLE-'].Widget
        x, y = event.x, event.y
        item = table.identify_row(y)
        col = table.identify_column(x)
        error_col_index = 7  # 错误信息列的索引，从0开始
        if tooltip:
            tooltip.destroy()
        if col == f'#{error_col_index + 1}' and item:  # tkinter列索引从 #1 开始
            task_id = int(item)
            task_info = task_manager.TASK_DICT.get(list(task_manager.TASK_DICT.keys())[task_id])
            error_info = task_info.get('error', '')
            if not error_info:
                return
            tooltip = tk.Toplevel(table)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root + 15}+{event.y_root + 10}")
            label = tk.Label(tooltip, text=error_info, background="#ffffe0", foreground="#000000", relief="solid",
                             borderwidth=1,
                             wraplength=300)
            label.pack()
            tooltip.update_idletasks()
            tooltip.deiconify()

    window = sg.Window('自动上传工具 - 任务监控', layout, resizable=True, finalize=True)
    window.set_min_size((800, 600))

    table = window['-TASK_TABLE-'].Widget
    table.bind("<Motion>", show_tooltip)

    try:
        while True:
            event, values = window.read()
            if event in (sg.WINDOW_CLOSED, '退出'):
                break
            elif event == '-LISTEN_DIR-':
                task_manager.LISTEN_DIR = not task_manager.LISTEN_DIR
                window['-LISTEN_DIR-'].update(
                    text='文件夹监听中' if task_manager.LISTEN_DIR else '文件夹未监听')
            elif event == '-CHANGE_LISTEN_DIR-':
                task_manager.change_listen_dir(pathlib.Path(values['-CHANGE_LISTEN_DIR-']))
                window['-LISTEN_DIR_LABEL-'].update(str(task_manager.UPLOAD_DIR.absolute()))
            elif event == '-SHOW_BROWSER-':
                show_browser = values['-SHOW_BROWSER-']
                if show_browser:
                    task_manager.show_browser = True
                else:
                    task_manager.show_browser = False
            elif event == '关于':
                sg.popup('有问题联系作者：ymc1107238486', title='关于')
            elif event == '开始':
                try:
                    task_manager.THREAD_COUNT = int(values['-THREAD_COUNT-'])
                    assert 0 < task_manager.THREAD_COUNT <= 50, '线程数量必须在1-50之间'
                except AssertionError as e:
                    sg.popup(e)
                    continue
                except:
                    sg.popup('线程数量必须为整数')
                    continue
                task_manager.start()
                # 启动更新线程
                threading.Thread(target=update_result_count, args=(window,), daemon=True).start()
                threading.Thread(target=run_func, daemon=True).start()

                window['开始'].update(disabled=True)
                window['-THREAD_COUNT-'].update(disabled=True)
                window['-CHANGE_LISTEN_DIR-'].update(disabled=True)
                window['-SHOW_BROWSER-'].update(disabled=True)
                print(f'任务开始，线程数量：{task_manager.THREAD_COUNT}')
        window.close()
    except Exception as e:
        print(e)
        sg.popup(f'程序出错：{e}')
