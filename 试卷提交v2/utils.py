import os
import queue
import threading
import zipfile
from pathlib import Path

import pandas as pd

INCLUDE_FILES = ['.doc', '.docx', '.pdf', '.mp3', '.zip']


class LogHandler:
    def __init__(self):
        self.queue = queue.Queue()

    def log(self, message):
        self.queue.put(message)

    def update_logs(self):
        def _update_logs(self):
            while True:
                message = self.queue.get()
                print('---------------------------')
                print(message)

        threading.Thread(target=_update_logs, daemon=True, args=(self,)).start()


class Tools:
    @staticmethod
    def init_dir(path_str):
        new_path = Path(path_str)
        new_path.mkdir(exist_ok=True)
        return new_path

    @staticmethod
    def list_all_files(directory: Path):
        """
        遍历文件夹下所有文件（包括子文件夹）

        Args:
            directory (str): 要遍历的文件夹路径

        Returns:
            list: 所有文件的完整路径列表
        """
        for root, dirs, files in os.walk(directory):
            for file in files:
                full_path = Path(os.path.join(root, file))
                if full_path.suffix not in INCLUDE_FILES:
                    continue
                yield full_path

    @staticmethod
    def move_to_dir(file_path: Path, dir_path: Path):
        try:
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)
            os.rename(file_path, dir_path.name + '/' + file_path.name)
        except Exception as e:
            print(f"[警告] 移动文件出错 [{file_path} => {dir_path}] {e}")

    @staticmethod
    def unzip_file(file_path: Path, to_dir: Path):
        assert file_path.suffix == '.zip', '只支持zip文件'
        # 解压文件
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(to_dir / file_path.stem)
        return to_dir / file_path.stem

    @staticmethod
    def load_keywords(file_path, sheet_name):
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # 一行一行读取转为列表
            return df.values.tolist()
        except Exception as e:
            print(f"读取Excel文件出错: {e}")
            raise e


class FileParse:
    def __init__(self, file_path: Path, grade_map, class_map, catalog_map):
        self.file_path = file_path
        self.grade_map = grade_map
        self.class_map = class_map
        self.catalog_map = catalog_map
        self.grade_name = self.get_grade()
        self.class_name = self.get_class()
        self.catalog_name = self.get_catalog()

    def get_grade(self):
        for grade_name, grade in self.grade_map:
            if grade_name in self.file_path.name:
                return grade
        raise Exception('年级解析失败')

    def get_class(self):
        for class_name in self.class_map:
            if class_name in self.file_path.name:
                return class_name
        raise Exception('学科解析失败')

    def get_catalog(self):
        for grade, kw, catalog in self.catalog_map:
            if grade == self.grade_name and kw in self.file_path.name:
                return catalog
        raise Exception('目录解析失败')
