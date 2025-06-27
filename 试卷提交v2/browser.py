import json
import os
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, Page, Locator


class Browser:
    def __init__(self, login_url, index_url):
        self.login_url = login_url
        self.index_url = index_url
        self.browser = None
        self.context = None
        self.main_page = None
        self.is_login = False

    def upload_file(self, page, add_btn_locator: Locator, file_path: Path):
        with page.expect_file_chooser() as fc_info:
            add_btn_locator.click()
            file_chooser = fc_info.value
            file_chooser.set_files(file_path)

    def lunch(self):
        p = sync_playwright().start()
        self.browser = p.chromium.launch(headless=False)
        self.context = self.browser.new_context()
        self.main_page = self.create_page()

    def create_page(self) -> Page:
        page = self.context.new_page()
        # 最大化
        page.set_viewport_size({"width": 1920, "height": 1080})
        return page

    def check_login_status(self):
        raise NotImplementedError

    def login_by_cookie(self):
        # 通过cookies.json登录
        if not os.path.exists('../cookies.json'):
            self.main_page.goto(self.login_url)
            return False
        # logger.log('尝试使用cookies.json登录')
        with open('../cookies.json', 'r', encoding='utf-8') as f:
            cookies = json.load(f)
        self.context.add_cookies(cookies)
        try:
            self.main_page.goto(self.index_url)
            self.check_login_status()
            return True
        except:
            print('登录失效，请手动登录')
            return False

    def login(self):
        try:
            if not self.login_by_cookie():
                self.main_page.goto(self.login_url)
            else:
                return True
            # 等待登录
            print('请在浏览器完成登录')
            # '商城管理' 等待可见
            while True:
                try:
                    self.check_login_status()
                    break
                except:
                    time.sleep(0.5)
            # 获取 cookies
            cookies = self.browser.contexts[0].cookies()
            # 保存 cookies 到文件
            with open('../cookies.json', 'w', encoding='utf-8') as f:
                json.dump(cookies, f)
            self.is_login = True
            return True
        except Exception as e:
            self.close()
            raise e

    def close(self):
        try:
            self.browser.close()
        except:
            pass

