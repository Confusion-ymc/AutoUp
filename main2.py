import json
import os
import shutil
import time
import zipfile
from pathlib import Path

import pandas as pd
from playwright._impl._errors import TargetClosedError
from playwright.sync_api import sync_playwright, Locator, expect
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')


def init_dir(path_str):
    new_path = Path(path_str)
    new_path.mkdir(exist_ok=True)
    return new_path


UPLOAD_DIR = init_dir('待上传')
FAILED_DIR = init_dir('上传失败')
SUCCESS_DIR = init_dir('上传成功')
TEMP_DIR = Path('temp')
UPLOAD_LOG = Path('上传成功.txt')
if not UPLOAD_LOG.exists():
    UPLOAD_LOG.touch()

WAIT_TIME = 2  # 每步等待秒


class FileParse:
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.detail_img_dir = TEMP_DIR / self.file_path.stem / '商品详情页图'
        self.main_img_dir = TEMP_DIR / self.file_path.stem / '商品主图'
        self.color_img_dir = TEMP_DIR / self.file_path.stem / '颜色图'
        self.catalog = " ＞ ".join(self.file_path.parts[1:-1])
        self.unzip_file()
        self.goods_name = ''
        self.factory_name = ''
        self.sku_data = []
        self.parse_excel()

    def unzip_file(self):
        assert self.file_path.suffix == '.zip', '只支持zip文件'
        # 清空TEMP文件夹下的文件
        if TEMP_DIR.exists():
            shutil.rmtree(TEMP_DIR)
        # 解压文件
        with zipfile.ZipFile(self.file_path, 'r') as zip_ref:
            zip_ref.extractall(TEMP_DIR)
        return TEMP_DIR

    def parse_excel(self):
        # 寻找xlsx文件
        for file_path in TEMP_DIR.rglob('*.xlsx'):
            df = pd.read_excel(file_path)
            self.goods_name = df['商品名称'][0]
            self.factory_name = df['供应商名称'][0]
            self.sku_data = [
                (
                    color,
                    group.drop(columns="颜色").to_dict("records")
                )
                for color, group in df.groupby("颜色")
            ]
            return
        raise FileNotFoundError('未找到xlsx文件')

    def detail_images(self):
        return list_all_files(self.detail_img_dir)

    def main_images(self):
        return list_all_files(self.main_img_dir)

    def color_images(self):
        return list_all_files(self.color_img_dir)


class AutoBrowserUpload:
    def __init__(self):
        self.browser, self.page = self.lunch_browser()
        self.login_url = 'https://api.hnqtyx.top/admin/login.html'
        self.index_url = 'https://api.hnqtyx.top/admin.html#/data/total.portal/index.html?spm=m-67-68-117'

    def lunch_browser(self):
        p = sync_playwright().start()
        b = p.chromium.launch(headless=False)
        page = b.new_page()
        # 最大化
        page.set_viewport_size({"width": 1920, "height": 1080})
        return b, page

    def check_login_status(self):
        login_user = self.page.locator('.layui-header.notselect').locator('a').last.text_content().strip()
        assert login_user != '立即登录'

    def login_by_cookie(self):
        # 通过cookies.json登录
        if not os.path.exists('cookies.json'):
            self.page.goto(self.login_url)
            return False
        print('尝试使用cookies.json登录')
        with open('cookies.json', 'r', encoding='utf-8') as f:
            cookies = json.load(f)
        self.browser.contexts[0].add_cookies(cookies)
        self.open_index()
        try:
            self.check_login_status()
            return True
        except:
            print('登录失效，请手动登录')
            return False

    def login(self):
        try:
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
            with open('cookies.json', 'w', encoding='utf-8') as f:
                json.dump(cookies, f)
            self.open_index()
            return True
        except Exception as e:
            self.close_browser()
            raise e

    def open_index(self):
        self.page.goto(self.index_url)

    def open_upload_page(self):
        self.page.get_by_text('添加商城商品', exact=True).click(timeout=120000)
        expect(self.page.get_by_text('添加商品数据')).to_be_visible()

    def close_browser(self):
        try:
            self.browser.close()
        except:
            pass

    def upload_file(self, add_btn_locator: Locator, file_path: Path):
        with self.page.expect_file_chooser() as fc_info:
            add_btn_locator.click()
            file_chooser = fc_info.value
            file_chooser.set_files(file_path)

    def fill_info(self, data_file: FileParse):
        self.page.locator('.layui-select-title').click()
        self.page.locator(f'.layui-anim-upbit').get_by_text(data_file.catalog).click()
        self.page.get_by_placeholder('请输入商品名称').fill('TEST----' + data_file.goods_name)
        self.page.get_by_placeholder('请输入厂家名称').fill(data_file.factory_name)

        # 上传封面和轮播图
        self.upload_file(self.page.locator('table').first.locator('a').first, data_file.main_images()[0])
        for img_path in data_file.main_images()[1:]:
            self.upload_file(self.page.locator('table').first.locator('a').last, img_path)

        self.page.locator('.layui-form-radio').get_by_text('实物').click()
        sku_set = []
        self.page.get_by_text('增加规则分组', exact=True).click()
        groups = self.page.get_by_placeholder('请输入分组名称').all()
        groups[0].fill('颜色')
        groups[1].fill('规格')
        color_index = 0
        sku_index = 0
        for color, sku_list in data_file.sku_data:
            if color_index > 0:
                self.page.get_by_text('增加', exact=True).all()[0].click()
            self.page.locator(".goods-spec-box.ng-scope").all()[0].get_by_placeholder('请输入规格').all()[
                color_index].fill(color)

            for sku in sku_list:
                if sku['规格'] not in sku_set:
                    sku_set.append(sku['规格'])
                    if sku_index > 0:
                        self.page.get_by_text('增加', exact=True).all()[1].click()
                    self.page.locator(".goods-spec-box.ng-scope").all()[1].get_by_placeholder('请输入规格').all()[
                        sku_index].fill(sku['规格'])
                    sku_index += 1
                # 填写价格
                # sku['供货价']
                tr_locator = self.page.locator(f"tr:has(td:text-is('{color}')):has(td:text-is('{sku['规格']}'))")
                tr_locator.locator("td:nth-child(4)").locator('input').fill(str(sku["供货价"] * 2.5))  # 价格在第4列
            color_index += 1

        # 先取消所有启用
        [item.click() for item in self.page.locator('.layui-table.margin-top-10').get_by_role('checkbox').all()]
        # 重新打开
        for color, sku_list in data_file.sku_data:
            for sku in sku_list:
                tr_locator = self.page.locator(f"tr:has(td:text-is('{color}')):has(td:text-is('{sku['规格']}'))")
                tr_locator.locator("td:nth-child(8)").get_by_role('checkbox').click()

        # 上传详情图片
        for img_path in data_file.detail_images():
            self.upload_file(self.page.locator('.cke_button_icon.cke_button__image_icon'), img_path)

    def confirm(self):
        self.page.locator('.layui-form-item.text-center').get_by_text('保存', exact=True).click()

    def check_result(self):
        print('  [保存中]')
        try:
            expect(self.page.get_by_text('添加商城商品', exact=True)).to_be_visible(timeout=120000)
            time.sleep(1)
        except:
            raise Exception('上传失败')


def list_all_files(directory: Path):
    """
    遍历文件夹下所有文件（包括子文件夹）

    Args:
        directory (str): 要遍历的文件夹路径

    Returns:
        list: 所有文件的完整路径列表
    """
    all_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            full_path = os.path.join(root, file)
            all_files.append(Path(full_path))
    try:
        return sorted(all_files, key=lambda x: int(x.stem.split('_')[1]))
    except:
        return all_files


class UploadLoger:
    def __init__(self, log_file_path: Path):
        self.log_file_path = log_file_path

    def log(self, file_path: Path):
        with open(self.log_file_path, 'a', encoding='utf-8') as f:
            f.write(file_path.name + '\n')


def move_to_dir(file_path: Path, dir_path: Path):
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
    os.rename(file_path, dir_path.name + '/' + file_path.name)


def run():
    browser = AutoBrowserUpload()
    upload_loger = UploadLoger(UPLOAD_LOG)
    if not browser.login_by_cookie():
        browser.login()
    print('登录成功')
    browser.page.get_by_text('商品管理', exact=True).click()
    n = 0
    for file_path in list_all_files(UPLOAD_DIR):
        if file_path.suffix != '.zip':
            continue
        n += 1
        try:
            print('-------------')
            print(f"{n}.[开始上传] {file_path.name}")
            browser.open_upload_page()
            time.sleep(WAIT_TIME)
            # 上传文件
            data_file = FileParse(file_path)
            browser.fill_info(data_file)
            browser.confirm()
            browser.check_result()
            upload_loger.log(file_path)
            print('  [上传成功]')
            # 移动到成功文件夹
            move_to_dir(file_path, SUCCESS_DIR)
        except TargetClosedError as e:
            print(f'  [上传失败] 详细信息:\n   浏览器被关闭, 停止本次上传 {e}')
            break
        except Exception as e:
            print(f'  [上传失败] 详细信息:\n   {e}')
            # 移动到失败文件夹
            move_to_dir(file_path, FAILED_DIR)
    input('上传完成, 回车退出并关闭浏览器')
    browser.close_browser()


if __name__ == '__main__':
    print('启动中...')
    try:
        run()
    except Exception as e:
        input(f'程序错误, {e}')
