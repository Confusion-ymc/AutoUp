import threading
import time

import PySimpleGUI as sg

from task_manager import task_manager

sg.theme('SystemDefaultForReal')


def main_gui(run_func):
    font = ('微软雅黑', 15)
    layout = [[sg.Multiline(size=(80, 20), key='-OUTPUT-', autoscroll=True, reroute_stdout=True, font=font,
                            enable_events=True,  # 允许选择文本
                            write_only=False,  # 设为False才能选择和复制文本
                            disabled=False)],
              # 添加一个线程数量输入框
              [sg.Text('待执行:'), sg.Text('0', key='-TASK_COUNT-'),
               sg.Text('运行中:'), sg.Text('0', key='-RUNNING-'),
               sg.Text('成功数量:'), sg.Text('0', key='-SUCCESS-'),
               sg.Text('重复数量:'), sg.Text('0', key='-REPEAT-'),
               sg.Text('失败数量:'), sg.Text('0', key='-FAILED-'),
               sg.Button('文件夹监听中',
                         key='-listen_dir-',
                         border_width=0),
               sg.Push(),
               sg.Text('线程数量:'), sg.InputText(str(task_manager.THREAD_COUNT), key='-THREAD_COUNT-', size=(5, 1)),
               sg.Button('开始'), sg.Button('退出'), sg.Button('关于')]]

    def update_result_count():
        while True:
            time.sleep(1)
            if not task_manager.LISTEN_DIR:
                continue
            window['-SUCCESS-'].update(len(task_manager.SUCCESS_COUNT))
            window['-FAILED-'].update(len(task_manager.FAILED_COUNT))
            window['-REPEAT-'].update(len(task_manager.REPEAT_COUNT))
            window['-TASK_COUNT-'].update(task_manager.task_queue.qsize())
            window['-RUNNING-'].update(len(task_manager.RUNNING_COUNT))

    threading.Thread(target=update_result_count, daemon=True).start()
    window = sg.Window('自动上传工具', layout)
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, '退出'):
            break
        elif event == '-listen_dir-':
            task_manager.LISTEN_DIR = not task_manager.LISTEN_DIR  # state of toggle changed
            window['-listen_dir-'].update(  # update state of button
                text='文件夹监听中' if task_manager.LISTEN_DIR else '文件夹未监听')
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
            # 禁用开始按钮
            window['开始'].update(disabled=True)
            window['-THREAD_COUNT-'].update(disabled=True)
            threading.Thread(target=run_func, daemon=True).start()
            print(f'任务开始，线程数量：{task_manager.THREAD_COUNT}')
    window.close()
