import json
import os
import shutil
import zipfile
from pathlib import Path
from time import sleep

import pandas as pd
from playwright._impl._errors import TargetClosedError
from playwright.sync_api import sync_playwright

UPLOAD_DIR = '待上传'
SUCCESS_DIR = '上传成功'
FAILED_DIR = '上传失败'
REPEAT_DIR = '重复上传'
UPLOAD_LOG = '上传成功.txt'
SPACIAL_DIR = '精品资料'

TEMP_DIR = 'temp'

WAIT_TIME = 1  # 每步等待秒

UPLOAD_DIR = Path(UPLOAD_DIR)
SUCCESS_DIR = Path(SUCCESS_DIR)
FAILED_DIR = Path(FAILED_DIR)
REPEAT_DIR = Path(REPEAT_DIR)
UPLOAD_LOG = Path(UPLOAD_LOG)
TEMP_DIR = Path(TEMP_DIR)
SPACIAL_DIR = Path(SPACIAL_DIR)


class AlreadyUploadError(Exception):
    pass


class SpacialFileError(Exception):
    pass


class FileParse:
    def __init__(self, file_path: Path, grade_map, subject_map, class_map):
        self.grade_key_word = None
        self.class_key_word = None
        self.subject_key_word = None

        self.grade_map = grade_map
        self.subject_map = subject_map
        self.class_map = class_map
        self.file_path = file_path
        self.grade_type, self.grade, self.step = self.get_grade()
        self.subject_type, self.subject = self.get_subject()
        self.class_type, self.class_child = self.get_class()

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
        res = None
        index = 9999
        for grade, grade_type in self.grade_map:
            if grade in self.file_path.name:
                self.grade_key_word = grade
                item_index = self.file_path.name.index(grade)
                if item_index <= index:
                    index = item_index
                    res = (grade_type, grade, step)
        if res:
            return res
        raise Exception('年级解析失败')

    def get_subject(self):
        res = None
        index = 9999
        for subject, subject_type in self.subject_map:
            if subject in self.file_path.name:
                self.subject_key_word = subject
                item_index = self.file_path.name.index(subject)
                if item_index <= index:
                    index = item_index
                    res = (subject_type, subject)
        if res:
            return res
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
            return res
        raise Exception('资料栏目解析失败')


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
            self.page.wait_for_url('https://www.21cnjy.com/webupload/', timeout=60000)
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
            # file_name = file_path.split('.')[0]
            success_tips = self.page.locator(
                f'.file-info:has(input[value="{file_path.name}"]) .success-tips')
            try:
                success_tips.wait_for(timeout=10000)
            except Exception as e:
                raise e

    def fill_info(self, file_parse: FileParse, file_path: Path):
        all_select_type = self.page.locator('.select_type').all()
        for item in all_select_type:
            item.click()
            sleep(WAIT_TIME)
            item.locator('span[class="7"]', has_text='试卷').click()
            sleep(WAIT_TIME)
        # 尝试填写集合标题
        if file_path.name.endswith('.zip'):
            try:
                self.page.get_by_placeholder("请输入合集标题，便于精确搜索、推广、查重").fill(file_path.name)
            except:
                pass
        # 获取学段和学科
        grade_type, grade, step = file_parse.grade_type, file_parse.grade, file_parse.step
        subject_type, subject = file_parse.subject_type, file_parse.subject
        self.page.locator(".xdxk_select").click()
        sleep(WAIT_TIME)
        self.page.locator("#xd").get_by_text(grade_type, exact=True).click()
        sleep(WAIT_TIME)
        self.page.locator("#chid").get_by_text(subject_type, exact=True).click()
        sleep(WAIT_TIME)

        # 选择栏目
        class_type, class_child = file_parse.class_type, file_parse.class_child
        self.page.locator("#j-add-cate").click()
        sleep(WAIT_TIME)
        # 特殊处理 中考专区/高考专区
        if class_type == '中考专区/高考专区':
            if grade_type == '高中':
                self.page.locator('.t2hd').get_by_text('高考专区', exact=True).click()
            else:
                self.page.locator('.t2hd').get_by_text('中考专区', exact=True).click()
        else:
            self.page.locator('.t2hd').get_by_text(class_type, exact=True).click()
        sleep(WAIT_TIME)
        # 特殊处理地理
        if subject_type == '地理' and class_child == '模拟试题':
            class_child = '模拟试卷'
        # 特殊处理 与试卷题目中的年级一致，注意分上下册（试卷题目有上、下之分）
        if class_child == '与试卷题目中的年级一致，注意分上下册（试卷题目有上、下之分）':
            if grade == '高三':
                # 高三不填上下
                self.page.locator('.t2end', has_text=grade).click()
            else:
                self.page.locator('.t2end', has_text=grade + step).click()
        elif class_child == '与试卷题目中的年级一致':
            self.page.locator('.t2end', has_text=grade).click()
        else:
            self.page.locator('.t2end', has_text=class_child).click()
        sleep(WAIT_TIME)
        self.page.locator('.ui-dialog-grid').get_by_role('button', name='确定').click()
        self.page.locator('#versionid').select_option('110')
        sleep(WAIT_TIME)
        self.page.locator('.anonymous_name').click()

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
        os.rename(file_path, dir_path.name + '/' + file_path.name)
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
    browser = AutoBrowserUpload()
    upload_loger = UploadLoger(UPLOAD_LOG)
    if not browser.login_by_cookie():
        browser.login()
    print('登录成功')
    print('任务开始')
    n = 0
    for file_path in for_upload_files():
        n += 1
        try:
            print('-------------')
            print(f"{n}.[开始上传] {file_path.name}")
            upload_loger.check(file_path)
            # 上传文件
            file_parse = FileParse(file_path, grade_map, subject_map, class_map)
            print(
                f'  [匹配到关键词] 年级:{file_parse.grade_key_word}|学科:{file_parse.subject_key_word}|类型:{file_parse.class_key_word}')
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
    input('上传完成, 回车退出并关闭浏览器')
    browser.close_browser()


if __name__ == '__main__':
    print('启动中...')
    try:
        run()
    except Exception as e:
        input(f'程序错误, {e}')
