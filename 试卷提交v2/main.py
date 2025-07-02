# https://www.jyeoo.com/ushare/collectCenter
import datetime
import threading
import time
from pathlib import Path
from queue import Empty
from typing import Optional

from playwright.sync_api import Locator, expect, Page
import warnings

from browser import Browser
from exception import MyRepeatError
from task_manager import task_manager
from ui2 import main_gui
from utils import Tools, FileParse

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')


class MyBrowser(Browser):
    def __init__(self):
        super().__init__(login_url='https://www.jyeoo.com/account/loginform',
                         index_url="https://www.jyeoo.com/")

    def check_login_status(self):
        expect(self.main_page.locator('.profile-drop')).to_be_visible(timeout=2000)


class AutoBrowserUpload:
    def __init__(self, task_name: str):
        self.thread_name = task_name
        self.page: Optional[Page] = None
        self.target_url = 'https://www.jyeoo.com/ushare/collectCenter'
        self.stop = False
        self.thread = None
        self.browser = MyBrowser()
        self.task_queue = task_manager.task_queue

    def open_page(self):
        self.page.goto(self.target_url)

    def start(self):
        def _start(self):
            self.browser.lunch(headless=not task_manager.show_browser)
            self.browser.login()
            self.page = self.browser.main_page
            file_path = None
            while not self.stop:
                try:
                    task_id = self.task_queue.get(timeout=1)
                except Empty:
                    time.sleep(1)
                    continue
                try:
                    file_path = task_manager.TASK_DICT[task_id]['file_path']
                    # 维护任务信息
                    task_manager.TASK_DICT[task_id]['status'] = '执行中'
                    task_manager.TASK_DICT[task_id]['status_changed'] = True
                    task_manager.TASK_DICT[task_id]['thread'] = self.thread_name
                    task_manager.TASK_DICT[task_id]['start_time'] = datetime.datetime.now()
                    task_manager.RUNNING_COUNT.append(file_path.name)
                    self.open_page()
                    self.fill_info(file_path)
                    self.check_result(file_path)

                    # 维护任务信息
                    task_manager.TASK_DICT[task_id]['status'] = '成功'
                    # 移动到成功文件夹
                    task_manager.SUCCESS_COUNT.append(file_path.name)
                    Tools.move_to_dir(file_path, task_manager.SUCCESS_DIR)
                except MyRepeatError as e:
                    # 维护任务信息
                    task_manager.TASK_DICT[task_id]['status'] = '重复'
                    task_manager.TASK_DICT[task_id]['error'] = str(e)

                    task_manager.logger.log(f'[{self.thread_name}] [重复] {file_path.name} [{e}]')
                    # 移动到重复文件夹
                    task_manager.REPEAT_COUNT.append(file_path.name)
                    Tools.move_to_dir(file_path, task_manager.REPEAT_DIR)
                except Exception as e:
                    # 维护任务信息
                    task_manager.TASK_DICT[task_id]['status'] = '失败'
                    task_manager.TASK_DICT[task_id]['error'] = str(e)

                    task_manager.logger.log(f'[{self.thread_name}] [失败] {file_path.name} [{e}]')
                    # 移动到失败文件夹
                    task_manager.FAILED_COUNT.append(file_path.name)
                    Tools.move_to_dir(file_path, task_manager.FAILED_DIR)
                finally:
                    task_manager.TASK_DICT[task_id]['status_changed'] = True
                    task_manager.TASK_DICT[task_id]['end_time'] = datetime.datetime.now()
                    task_manager.TASK_DICT[task_id]['duration'] = int((
                                                                                   task_manager.TASK_DICT[task_id][
                                                                                       'end_time'] -
                                                                                   task_manager.TASK_DICT[task_id][
                                                                                       'start_time']
                                                                           ).total_seconds())
                    try:
                        task_manager.RUNNING_COUNT.remove(file_path.name)
                    except:
                        pass
                    Tools.clear_dir(task_manager.TEMP_DIR/file_path.stem)
                    Tools.clear_dir(task_manager.TEMP_DIR/file_path.name)
            self.browser.close()

        self.thread = threading.Thread(target=_start, daemon=True, args=(self,))
        self.thread.start()

    def options_select(self, options_locator: Locator, option_text: str):
        options_locator.click()
        time.sleep(0.5)
        options_locator.select_option(label=option_text)

    def fill_info(self, data_path: Path):
        # 填写内容 以及 上传文件
        # logger.log(f'[{self.task_name}] [开始上传] {data_path.name}')
        file_parse = FileParse(data_path, task_manager.GRADE_MAP, task_manager.CLASS_MAP,
                               task_manager.CATALOG_MAP)
        if data_path.suffix == '.zip':
            # 上传类型是压缩包
            unzip_dir = Tools.unzip_file(data_path, task_manager.TEMP_DIR)
            # 找到所有的文件
            index = 0
            for_upload_files = list(Tools.list_all_files(unzip_dir))
            media_files = [file_item_path for file_item_path in for_upload_files if file_item_path.suffix == '.mp3']
            if len(media_files) == 1:
                for file_item_path in for_upload_files:
                    if file_item_path.suffix == '.mp3':
                        continue
                    index += 1
                    if index == 1:
                        self.browser.upload_file(self.page, self.page.locator('.c-btn.btn-blue.c-btn-240'), file_item_path)
                    else:
                        self.browser.upload_file(self.page, self.page.locator('.c-btn.btn-blue.c-btn-200'), file_item_path)
                    if '解析' in file_item_path.stem or '答案' in file_item_path.stem:
                        self.page.locator('.upload-after').locator('input[type=text]').all()[index - 1].fill('答案解析')
                    else:
                        self.page.locator('.upload-after').locator('input[type=text]').all()[index - 1].fill(data_path.stem)
                        if media_files:
                            self.browser.upload_file(self.page,
                                                     self.page.locator('.upload-after').get_by_text('添加音频',
                                                                                                    exact=True).all()[
                                                         index - 1],
                                                     media_files[0])
                # 标题合集
                self.page.get_by_placeholder('请输入标题', exact=True).fill(data_path.stem)
            elif len(media_files) == 0:
                # 重新压缩文件
                Tools.compress_files_to_zip(for_upload_files, task_manager.TEMP_DIR / data_path.name)
                self.browser.upload_file(self.page, self.page.locator('.c-btn.btn-blue.c-btn-240'), task_manager.TEMP_DIR / data_path.name)

            elif len(media_files) > 1:
                raise Exception('压缩包内有多个音频文件 跳过上传')

        else:
            self.browser.upload_file(self.page, self.page.locator('.c-btn.btn-blue.c-btn-240'), data_path)

        self.options_select(self.page.locator('#SubjectID'), file_parse.grade_name + file_parse.class_name)
        self.options_select(self.page.locator('#SourceID'), file_parse.catalog_name)

        self.page.locator('#privace1').click()
        self.page.locator('.c-btn.btn-blue.c-btn-320').click()

    def check_result(self, file_path: Path):
        for i in range(300):
            tips = ''
            try:
                tips = self.page.locator('.body-content').locator('.p10').text_content(timeout=500)
            except:
                pass
            if '已经上传过，被拒绝概率较大，请确认是否继续上传？' in tips:
                # self.page.locator('input[value="取消上传"]').click()
                raise MyRepeatError(tips)
            try:
                tips_message = self.page.locator('.msgtip:visible').text_content(timeout=500)
            except:
                continue
            if '操作成功' in tips_message and '上传成功' in tips_message:
                if "已经有用户上传了，不能重复上传呢!" in tips_message or "可前往【我的上传】或【个人中心-我的上传】看审核进度！" in tips_message:
                    self.page.locator('.msgtip:visible').get_by_text('我知道了', exact=True).click()
                if "已经有用户上传了，不能重复上传呢!" in tips_message:
                    raise MyRepeatError(f'部分重复! {tips_message}')
                task_manager.logger.log(f'[{self.thread_name}] [上传成功] {file_path.name}')
                return True
            else:
                raise MyRepeatError(tips_message)
        raise Exception('等待超时')

    def wait(self):
        self.thread.join()


def run():
    login_b = MyBrowser()
    login_b.lunch()
    login_b.login()
    task_manager.logger.log('登录成功！')
    login_b.close()
    runners = []
    for i in range(task_manager.THREAD_COUNT):
        robot = AutoBrowserUpload(task_name=f'线程{i + 1}')
        robot.start()
        runners.append(robot)
    for runner in runners:
        runner.wait()


if __name__ == '__main__':
    main_gui(run)
