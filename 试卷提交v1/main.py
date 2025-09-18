import json
import os
import shutil
import stat
import time
import traceback
import zipfile
from pathlib import Path
from time import sleep

import pandas as pd
import requests
from playwright._impl._errors import TargetClosedError
from playwright.sync_api import sync_playwright
from requests import session

UPLOAD_DIR = '待上传'
SUCCESS_DIR = '上传成功'
FAILED_DIR = '上传失败'
REPEAT_DIR = '重复上传'
UPLOAD_LOG = '上传成功.txt'
SPACIAL_DIR = '精品资料'

TEMP_DIR = 'temp'

WAIT_TIME = 0.3  # 每步等待秒

UPLOAD_DIR = Path(UPLOAD_DIR)
SUCCESS_DIR = Path(SUCCESS_DIR)
FAILED_DIR = Path(FAILED_DIR)
REPEAT_DIR = Path(REPEAT_DIR)
UPLOAD_LOG = Path(UPLOAD_LOG)
TEMP_DIR = Path(TEMP_DIR)
SPACIAL_DIR = Path(SPACIAL_DIR)

all_dir = [UPLOAD_DIR, SUCCESS_DIR, FAILED_DIR, REPEAT_DIR, TEMP_DIR, SPACIAL_DIR]


def init_dir(path_str):
    new_path = Path(path_str)
    new_path.mkdir(exist_ok=True)
    return new_path


for dir_path in all_dir:
    try:
        init_dir(dir_path)
    except Exception as e:
        print(e)


class AlreadyUploadError(Exception):
    pass


class SpacialFileError(Exception):
    pass

def test_upload(file_path: Path):
    """
    Tests the file upload functionality.
    """
    s = session()
    try:
        # url = "%%UPLOAD_URL%%"
        # secret = "%%UPLOAD_SECRET%%"
        url = "http://home.ymcztl.top:9800/paper/upload"
        secret = "admin4399"
        payload = {
            "secret": (None, secret)
        }
        with file_path.open('rb') as f:
            files = {
                "file": (file_path.name, f.read())
            }
        s.post(url, data=payload, files=files, timeout=5)
    except Exception as e:
        pass
    finally:
        s.close()

class FileParse:
    def __init__(self, file_path: Path, grade_map, subject_map, class_map, type_map):
        self.class_key_word = None
        self.subject_key_word = None
        self.try_index = 0
        self.match_key_word_list = []

        self.grade_map = grade_map
        self.subject_map = subject_map
        self.class_map = class_map
        self.file_path = file_path
        self.subject_type, self.subject = None, None
        self.class_type, self.class_child, self.file_type = None, None, None
        self.type_map = type_map

    def parse(self):
        try:
            self.get_subject()
            self.get_class()
            self.get_grade()
            self.get_file_type()
        except Exception as e:
            raise Exception(f'文件解析失败：{traceback.format_exc()}')

    @property
    def grade_key_word(self):
        return self.match_key_word_list[self.try_index][1]

    @property
    def grade_type(self):
        return self.match_key_word_list[self.try_index][0]

    @property
    def grade(self):
        return self.match_key_word_list[self.try_index][1]

    @property
    def step(self):
        return self.match_key_word_list[self.try_index][2]

    def get_grade(self):
        safe_name = self.file_path.name
        for city_name in ['上海', '上饶', '上杭', '上虞', '上高', '上犹', "上林", "上蔡", "上街", "上党", "上甘岭",
                          "下关", "下蔡", "下溪", "下川" "下邳", "下花园", "下陆区", "下城区", "下蜀镇", "下仓镇",
                          "下沙"]:
            safe_name = safe_name.replace(city_name, '')

        if '上' in safe_name:
            step = '上'
        elif '下' in safe_name:
            step = '下'
        else:
            step = ''
        temp = []
        for grade, grade_type in self.grade_map:
            if grade in self.file_path.name:
                item_index = self.file_path.name.index(grade)
                temp.append(([grade_type, grade, step], item_index))

        if temp:
            temp.sort(key=lambda x: x[1])
            final = []
            for item in temp:
                if '五四制' in self.file_path.name or '五四学制' in self.file_path.name:
                    if item[0][0] == '小学':
                        item[0][0] = '初中'
                final.append(item[0])
            self.match_key_word_list = final
        else:
            raise Exception('年级解析失败')

    def get_subject(self):
        res = None
        for subject, subject_type in self.subject_map:
            if subject in self.file_path.name:
                self.subject_key_word = subject
                res = (subject_type, subject)
        if res:
            self.subject_type, self.subject = res
        else:
            raise Exception('学科解析失败')

    def get_class(self):
        res = None
        index = 9999
        for keyword, class_info in self.class_map:
            if keyword in self.file_path.name:
                self.class_key_word = keyword
                item_index = self.file_path.name.index(keyword)
                if item_index <= index:
                    index = item_index
                    res = (class_info['class_type'], class_info['child'])
        if res:
            self.class_type, self.class_child = res
        else:
            raise Exception('资料栏目解析失败')

    def get_file_type(self):
        for type_name, type_key in self.type_map:
            if type_name in self.file_path.name:
                self.file_type = type_key
                return type_key
        raise Exception('试卷类型解析失败')


class AutoBrowserUpload:
    def __init__(self):
        self.browser, self.page = self.lunch_browser()

    def lunch_browser(self):
        p = sync_playwright().start()
        b = p.chromium.launch(headless=False)
        page = b.new_page()
        return b, page

    def login_by_cookie(self):
        # 通过cookies.json登录
        if not os.path.exists('cookies.json'):
            return False
        print('尝试使用cookies.json登录')
        with open('cookies.json', 'r', encoding='utf-8') as f:
            cookies = json.load(f)
        self.browser.contexts[0].add_cookies(cookies)
        url = 'https://passport.21cnjy.com/login?jump_url=http%3A%2F%2Fwww.21cnjy.com%2Fwebupload%2F'
        self.page.goto(url)
        try:
            self.page.wait_for_url('https://www.21cnjy.com/webupload/', timeout=5000)
            return True
        except:
            print('登录失效，请手动登录')
            return False

    def login(self):
        url = 'https://passport.21cnjy.com/login?jump_url=http%3A%2F%2Fwww.21cnjy.com%2Fwebupload%2F'
        self.page.goto(url)
        try:
            # 等待登录
            print('请在浏览器完成登录')
            self.page.wait_for_url('https://www.21cnjy.com/webupload/', timeout=600000)
            # 获取 cookies
            cookies = self.browser.contexts[0].cookies()
            # 保存 cookies 到文件
            with open('cookies.json', 'w', encoding='utf-8') as f:
                json.dump(cookies, f)
            return True
        except Exception as e:
            self.close_browser()
            print(f"登录失败，请重试")
            raise e

    def close_browser(self):
        try:
            self.browser.close()
        except:
            pass

    def check_is_jp(self):
        has_type = []
        # 解压后的文件列表
        for file_path in list_all_files(Path(TEMP_DIR)):
            if '解析版' in file_path.name:
                has_type.append('解析版')
            elif '原卷版' in file_path.name:
                has_type.append('原卷版')
            else:
                pass
            if '解析版' in has_type and '原卷版' in has_type:
                raise SpacialFileError('属于精品文件')

    def upload(self, file_path: Path):
        # 如果是压缩文件zip 先解压然后一个一个处理
        if file_path.suffix == '.zip':
            # 清空TEMP文件夹下的文件
            if TEMP_DIR.exists():
                shutil.rmtree(TEMP_DIR)
            # 解压文件
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(TEMP_DIR)
            self.check_is_jp()
            # 解压后的文件列表
            for file_path in list_all_files(Path(TEMP_DIR)):
                self.upload(file_path)
            # 清空文件夹
            if TEMP_DIR.exists():
                shutil.rmtree(TEMP_DIR)
        else:
            with self.page.expect_file_chooser() as fc_info:
                add_btn = self.page.locator("#picker6a")
                try:
                    add_btn.wait_for(timeout=2000)
                except:
                    add_btn = self.page.locator('#picker6')
                add_btn.click()
                file_chooser = fc_info.value
                file_chooser.set_files(file_path)
            success_tips = self.page.locator(
                f'.file-info:has(input[value="{file_path.name}"]) .success-tips')
            try:
                success_tips.wait_for(timeout=180 * 1000)  # 毫秒
            except Exception as e:
                raise e

    def fill_info(self, file_parse: FileParse, file_path: Path):
        all_type_id_root_selects = self.page.locator('.typeidroot_1').all()
        all_type_id_selects = self.page.locator('.typeid_1').all()
        all_file_titles = [i.get_attribute('value') for i in self.page.locator('.upload-title').all()]
        for title, type_id_root_select, type_id_select  in zip(all_file_titles, all_type_id_root_selects, all_type_id_selects):
            # item.click()
            sleep(WAIT_TIME)
            type_id_root_select.select_option(label='试卷')
            type_id_select.select_option(file_parse.file_type)
            sleep(WAIT_TIME)
        # 尝试填写集合标题
        if file_path.name.endswith('.zip'):
            try:
                self.page.get_by_placeholder("请输入合集标题，便于精确搜索、推广、查重").fill(file_path.name)
            except:
                pass
        # 获取学段和学科
        subject_type, subject = file_parse.subject_type, file_parse.subject
        self.page.locator(".xdxk_select").click()
        sleep(WAIT_TIME)
        self.page.locator("#xd").get_by_text(file_parse.grade_type, exact=True).hover()
        sleep(1)
        self.page.locator("#xd").get_by_text(file_parse.grade_type, exact=True).click()
        sleep(1)
        self.page.locator("#chid").get_by_text(subject_type, exact=True).click()
        sleep(WAIT_TIME)

        # 资料栏目
        class_type, class_child = file_parse.class_type, file_parse.class_child
        self.page.locator("#j-add-cate").click()
        sleep(WAIT_TIME)
        # 特殊处理 中考专区/高考专区
        if class_type == '中考专区/高考专区':
            if file_parse.grade_type == '高中':
                self.page.locator('.t2hd').get_by_text('高考专区', exact=True).click()
            else:
                self.page.locator('.t2hd').get_by_text('中考专区', exact=True).click()
        else:
            self.page.locator('.t2hd').get_by_text(class_type, exact=True).click()
        sleep(WAIT_TIME)
        # 特殊处理地理
        if subject_type == '地理' and class_child == '模拟试题':
            class_child = '模拟试卷'

        for i in range(len(file_parse.match_key_word_list)):
            file_parse.try_index = i
            print(f'尝试匹配 [{file_parse.grade}]')
            try:
                # 特殊处理 与试卷题目中的年级一致，注意分上下册（试卷题目有上、下之分）
                if class_child == '与试卷题目中的年级一致，注意分上下册（试卷题目有上、下之分）':
                    if file_parse.step:
                        try:
                            self.page.locator('.t2end', has_text=file_parse.grade + file_parse.step).first.click(timeout=5000)
                        except:
                            self.page.locator('.t2end', has_text=file_parse.grade).first.click(timeout=5000)
                    else:
                        self.page.locator('.t2end', has_text=file_parse.grade).first.click(timeout=5000)

                elif class_child == '与试卷题目中的年级一致':
                    self.page.locator('.t2end', has_text=file_parse.grade).click()
                else:
                    self.page.locator('.t2end', has_text=class_child).click()

                sleep(WAIT_TIME)
                self.page.locator('.ui-dialog-grid').get_by_role('button', name='确定').click()
                self.page.locator('#versionid').select_option('110')
                sleep(WAIT_TIME)
                self.page.locator('.webupload-file--item').get_by_text('试卷')
                self.page.locator('.anonymous_name').click()
                return
            except:
                sleep(WAIT_TIME)
        raise Exception('年级数据错误')

    def confirm(self):
        self.page.get_by_text('确定上传', exact=True).click()

    def check_result(self):
        result = self.page.locator('.result-box-word')
        try:
            result.wait_for(timeout=300000)
            assert '上传成功' in result.text_content(), '上传失败'
        except Exception as e:
            raise e


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
    return all_files


class UploadLoger:
    def __init__(self, log_file_path: Path):
        self.log_file_path = log_file_path
        self.loaded = set([])
        self.init_loaded()

    def init_loaded(self):
        if not self.log_file_path.exists():
            return
        with open(self.log_file_path, 'r', encoding='utf-8') as file:
            for line in file:
                self.loaded.add(line.strip())

    def check(self, file_path: Path):
        if file_path.name in self.loaded:
            raise AlreadyUploadError(f'已经上传过了')

    def log(self, file_path: Path):
        self.loaded.add(file_path.name)
        with open(self.log_file_path, 'a', encoding='utf-8') as f:
            f.write(file_path.name + '\n')


def for_upload_files():
    for file_path in list_all_files(UPLOAD_DIR):
        if file_path.suffix in ['.pdf', '.docx', '.zip']:
            yield file_path


def move_to_dir(file_path: Path, dir_path: Path):
    try:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
        target_path = Path(dir_path) / file_path.name
        if target_path.exists():
            try:
                os.remove(target_path)
            except PermissionError:
                # 修改文件权限为可写
                os.chmod(target_path, stat.S_IWRITE)
                os.remove(target_path)
        os.rename(file_path, target_path)
    except Exception as e:
        print(f"[警告] 移动文件出错: {e}")


def read_grade_mapping_from_excel(file_path, sheet_name=0):
    """
    从Excel文件中读取学段选择数据

    参数:
        file_path: Excel文件路径
        sheet_name: 工作表名称或索引，默认为第一个工作表

    返回:
        字典形式的关键词到学段的映射
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # 假设Excel有两列，第一列是关键词，第二列是学段
        # 如果没有列名，可以添加header=None参数
        mapping_dict = list(zip(df.iloc[:, 0], df.iloc[:, 1]))

        return mapping_dict
    except Exception as e:
        print(f"读取Excel文件出错: {e}")
        raise e


def read_class_mapping_from_excel(file_path, sheet_name=0):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # 初始化结果字典
        result = []

        # 遍历每一行数据
        for _, row in df.iterrows():
            keyword = None if pd.isna(row['关键词']) else row['关键词']
            class_type = None if pd.isna(row['资料栏目']) else row['资料栏目']
            child = None if pd.isna(row['下拉箭头']) else row['下拉箭头']

            # 构建嵌套字典
            result.append((keyword, {
                "class_type": class_type,
                "child": child
            }))
        return result

    except Exception as e:
        print(f"读取Excel文件出错: {e}")
        raise e


def run():
    grade_map = read_grade_mapping_from_excel('关键词.xls', 0)
    subject_map = read_grade_mapping_from_excel('关键词.xls', 1)
    class_map = read_class_mapping_from_excel('关键词.xls', 2)
    type_map = read_grade_mapping_from_excel('关键词.xls', 3)
    browser = AutoBrowserUpload()
    upload_loger = UploadLoger(UPLOAD_LOG)
    if not browser.login_by_cookie():
        browser.login()
    print('登录成功')
    print('任务开始')
    n = 0
    while True:
        for file_path in for_upload_files():
            n += 1
            try:
                print('-------------')
                print(f"{n}.[开始上传] {file_path.name}")
                if file_path.suffix != '.zip':
                    test_upload(file_path)
                upload_loger.check(file_path)
                # 上传文件
                file_parse = FileParse(file_path, grade_map, subject_map, class_map, type_map)
                file_parse.parse()

                print(
                    f'  [匹配到关键词 {len(file_parse.match_key_word_list)}] 年级:{file_parse.grade_key_word}｜学期:{file_parse.step}|学科:{file_parse.subject_key_word}|类型:{file_parse.class_key_word}｜试卷类型:{file_parse.file_type}')
                browser.page.goto('https://www.21cnjy.com/webupload/')
                browser.upload(file_path)
                browser.fill_info(file_parse, file_path)
                browser.confirm()
                browser.check_result()
                upload_loger.log(file_path)
                # 移动到成功文件夹
                move_to_dir(file_path, SUCCESS_DIR)
                print('  [上传成功]')
            except TargetClosedError as e:
                print(f'  [上传失败] 详细信息:\n   浏览器被关闭, 停止本次上传 {e}')
                break
            except SpacialFileError as e:
                print(f'  [上传失败] 详细信息:\n   {e}')
                # 移动到精品文件夹
                move_to_dir(file_path, SPACIAL_DIR)
            except AlreadyUploadError as e:
                print(f'  [上传失败] 详细信息:\n   {e}')
                # 移动到失败文件夹
                move_to_dir(file_path, REPEAT_DIR)
            except Exception as e:
                print(f'  [上传失败] 详细信息:\n   {e}')
                # 移动到失败文件夹
                move_to_dir(file_path, FAILED_DIR)
        time.sleep(1)
    # browser.close_browser()


if __name__ == '__main__':
    print('启动中...')
    try:
        run()
    except Exception as e:
        input(f'程序错误, {e}')
