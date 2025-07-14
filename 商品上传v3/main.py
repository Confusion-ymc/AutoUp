import json
import os
import threading
import time
from pathlib import Path

import pandas as pd
from playwright._impl._errors import TargetClosedError
from playwright.sync_api import sync_playwright, Locator, expect
import datetime
import warnings

# 忽略特定警告
warnings.filterwarnings('ignore', category=UserWarning, message="Workbook contains no default style")
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')
# 创建一个信号
finish_event = threading.Event()

# 用simplegui弹出提示
import PySimpleGUI as sg


def show_wx_tip():
    skip = False
    file_name = '今日不再提示.txt'
    wx = 'ymc1107238486'
    # 检查是否存在文件
    if os.path.exists(file_name):
        with open(file_name, 'r') as f:
            # 读取文件内容
            content = f.read()
            # 转换为时间戳
            content = int(content)
            # 转为datetime 通过day判断是否是同一天
            content = datetime.datetime.fromtimestamp(content)
            # 获取当前时间
            now = datetime.datetime.now()
            # 判断是否是同一天
            if content.day == now.day:
                skip = True
    if skip:
        return
    # 让提示词可以复制
    layout = [
        [sg.Text('欢迎使用自动上传工具，有问题请联系开发者微信'), ],
        [sg.Text(wx), sg.Button('复制', key='Copy')],
        # 按钮靠右
        [sg.Push(), sg.Button('今日不再提示')]
    ]
    window = sg.Window('提示', layout)
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == '关闭':
            break
        elif event == 'Copy':
            # 复制文本
            window.TKroot.clipboard_clear()
            window.TKroot.clipboard_append(wx)
        elif event == '今日不再提示':
            # 写入一个文件 下次不再提示 记录时间戳
            with open(file_name, 'w') as f:
                # 写入当前时间戳
                f.write(str(int(time.time())))
            break
    window.close()


def main_gui(run_func):
    layout = [
        [sg.Multiline(size=(80, 20), key='-OUTPUT-', autoscroll=True, reroute_stdout=True)],
        [sg.Push(), sg.Button('开始'), sg.Button('退出')],
    ]

    window = sg.Window('自动上传工具', layout)

    # 根据信号更新按钮状态
    def update_button():
        while True:
            if finish_event.is_set():
                window['开始'].update(disabled=False)
            time.sleep(0.1)

    threading.Thread(target=update_button, daemon=True).start()

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, '退出'):
            break
        elif event == '开始':
            # 禁用开始按钮
            window['开始'].update(disabled=True)
            run_func()

    window.close()


def init_dir(path_str):
    new_path = Path(path_str)
    new_path.mkdir(exist_ok=True)
    return new_path


def get_price_rate():
    """
    读取 汇率设置.txt
    :return:
    """
    try:
        with open('汇率设置.txt', 'r', encoding='utf-8') as f:
            return float(f.read().strip())
    except Exception as e:
        print(f"[警告] 读取汇率设置出错 {e}, 使用默认汇率7")
        return 7


UPLOAD_DIR = init_dir('待上传')
FAILED_DIR = init_dir('上传失败')
SUCCESS_DIR = init_dir('上传成功')
# TEMP_DIR = Path('temp')
UPLOAD_LOG = Path('上传成功.txt')

USD_RATE = get_price_rate()
USD_ADD = 15

if not UPLOAD_LOG.exists():
    UPLOAD_LOG.touch()

WAIT_TIME = 2  # 每步等待秒


class FileParse:
    def __init__(self, file_path: Path, catalog: str):
        self.file_path = file_path
        self.detail_img_dir = self.file_path / '商品详情页图'
        self.main_img_dir = self.file_path / '商品主图'
        self.color_img_dir = self.file_path / '颜色图'
        self.catalog = catalog
        self.goods_name = ''
        self.factory_name = ''
        self.sku_data = []
        self.parse_excel()

    def parse_excel(self):
        # 寻找xlsx文件
        for file_path in self.file_path.rglob('*.xlsx'):
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
        print('尝试自动登录')
        with open('cookies.json', 'r', encoding='utf-8') as f:
            cookies = json.load(f)
        self.browser.contexts[0].add_cookies(cookies)
        self.open_index()
        try:
            self.check_login_status()
            return True
        except:
            print('[警告] 登录失效，请手动登录')
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
        self.page.goto(self.index_url, timeout=10000)

    def open_upload_page(self):
        print('  [正在自动打开上传页面]')
        self.open_index()
        self.page.get_by_text('商品管理', exact=True).click(timeout=10000)
        try:
            self.page.get_by_text('添加商城商品', exact=True).click(timeout=10000)
        except Exception as e:
            print('[警告] 打开上传页面失败, 请手动打开')
        while True:
            try:
                expect(self.page.get_by_text('添加商品数据')).to_be_visible(timeout=60000)
                break
            except:
                pass

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
        time.sleep(WAIT_TIME)
        self.page.locator('.layui-anim-upbit').get_by_text(data_file.catalog).first.click()
        self.page.get_by_placeholder('请输入商品名称').first.fill(str(data_file.goods_name))
        self.page.get_by_placeholder('请输入商品名称').last.fill(str(data_file.goods_name))
        self.page.get_by_placeholder('请输入厂家名称').fill(str(data_file.factory_name))

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
                color_index].fill(str(color))

            for sku in sku_list:
                if sku['规格'] not in sku_set:
                    sku_set.append(sku['规格'])
                    if sku_index > 0:
                        self.page.get_by_text('增加', exact=True).all()[1].click()
                    self.page.locator(".goods-spec-box.ng-scope").all()[1].get_by_placeholder('请输入规格').all()[
                        sku_index].fill(str(sku['规格']))
                    sku_index += 1
                # 填写价格
                # sku['供货价']
                tr_locator = self.page.locator(f"tr:has(td:text-is('{color}')):has(td:text-is('{sku['规格']}'))")
                tr_locator.locator("td:nth-child(4)").locator('input').fill(str(sku["供货价"] * 2.5))
                tr_locator.locator("td:nth-child(5)").locator('input').fill(str(round(sku["供货价"] / USD_RATE + USD_ADD, 2)))
            color_index += 1

        # 先取消所有启用
        [item.click() for item in self.page.locator('.layui-table.margin-top-10').get_by_role('checkbox').all()]
        # 重新打开
        for color, sku_list in data_file.sku_data:
            for sku in sku_list:
                tr_locator = self.page.locator(f"tr:has(td:text-is('{color}')):has(td:text-is('{sku['规格']}'))")
                tr_locator.get_by_role('checkbox').click()

        # 上传详情图片
        for img_path in data_file.detail_images():
            self.upload_file(self.page.locator('.cke_button_icon.cke_button__image_icon'), img_path)

    def confirm(self):
        self.page.locator('.layui-form-item.text-center').get_by_text('保存', exact=True).click()

    def check_result(self):
        print('  [保存中]')
        for i in range(120):
            try:
                # js获取当前页面url
                url = self.page.evaluate('window.location.href')
                if 'index.html' in url:
                    print('  [上传成功]1')
                    return True
                if self.page.locator('.layui-layer-content.layui-layer-padding').text_content(
                        timeout=1000) == '商品编辑成功！':
                    print('  [上传成功]2')
                    return True
                else:
                    time.sleep(0.1)
            except:
                pass
        raise Exception('[错误] 上传失败, 保存数据失败')


def list_all_files(directory: Path, just_root=False):
    """
    遍历文件夹下所有文件（包括子文件夹）

    Args:
        directory (str): 要遍历的文件夹路径
        just_root (bool): 是否只返回根目录下的文件

    Returns:
        list: 所有文件的完整路径列表
    """
    all_files = []
    files = os.listdir(directory)
    if just_root:
        return [Path(directory) / file for file in files]
    for root, dirs, files in os.walk(directory):
        for file in files:
            full_path = os.path.join(root, file)
            all_files.append(Path(full_path))
    try:
        return sorted(all_files, key=lambda x: int(x.stem.split('_')[1]))
    except:
        return all_files


def list_all_dirs(directory: Path, just_root=False):
    """
    遍历文件夹下所有文件夹（包括子文件夹）
    Args:
        directory (str): 要遍历的文件夹路径
        just_root (bool): 是否只返回根目录下的文件夹
    Returns:
        list: 所有文件夹的完整路径列表
    """
    all_dirs = []
    for root, dirs, files in os.walk(directory):
        if just_root:
            return [Path(root) / dir for dir in dirs]
        for dir in dirs:
            full_path = os.path.join(root, dir)
            all_dirs.append(Path(full_path))
    return all_dirs


class UploadLoger:
    def __init__(self, log_file_path: Path):
        self.log_file_path = log_file_path

    def log(self, file_path: Path):
        with open(self.log_file_path, 'a', encoding='utf-8') as f:
            f.write(file_path.name + '\n')


def move_to_dir(file_path: Path, dir_path: Path):
    try:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
        os.rename(file_path, dir_path.name + '/' + file_path.name)
    except Exception as e:
        print(f'[警告] 移动文件失败 {e}')


def run():
    def task():
        print(f'启动中... {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        try:
            browser = AutoBrowserUpload()
            upload_loger = UploadLoger(UPLOAD_LOG)
            if not browser.login_by_cookie():
                browser.login()
            print('登录成功')
        except Exception as e:
            print(f'[错误] 系统错误 {e}')
            return
        n = 0
        for catalog_dir in list_all_dirs(UPLOAD_DIR, just_root=True):
            catalog = catalog_dir.name.replace('-', " ＞ ")
            for goods_dir in list_all_dirs(catalog_dir, just_root=True):
                n += 1
                try:
                    print('-------------')
                    print(f"{n}.[开始上传] [{catalog}] [{goods_dir.name}]")
                    browser.open_upload_page()
                    time.sleep(WAIT_TIME)
                    # 上传文件
                    data_file = FileParse(goods_dir, catalog)
                    browser.fill_info(data_file)
                    browser.confirm()
                    browser.check_result()
                    upload_loger.log(goods_dir)
                    # 移动到成功文件夹
                    move_to_dir(goods_dir, SUCCESS_DIR)
                except TargetClosedError as e:
                    print(f'  [错误] [上传失败] 详细信息:\n   浏览器被关闭, 停止本次上传 {e}')
                    break
                except Exception as e:
                    print(f'  [错误] [上传失败] 详细信息:\n   {e}')
                    # 移动到失败文件夹
                    move_to_dir(goods_dir, FAILED_DIR)
        print('上传完成')
        browser.close_browser()
        finish_event.set()

    # 启动线程
    threading.Thread(target=task, daemon=True).start()


if __name__ == '__main__':
    sg.theme('SystemDefaultForReal')
    show_wx_tip()
    main_gui(run)
