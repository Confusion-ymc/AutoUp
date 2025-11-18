import os
import queue
import shutil
import threading
import zipfile
from pathlib import Path

import pandas as pd
from pydub import AudioSegment

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
        只返回第一级目录下的文件，不包含子目录中的文件

        Args:
            directory (str): 要遍历的文件夹路径

        Returns:
            list: 所有文件的完整路径列表
        """
        # 检查目录是否存在
        # if directory.exists() and directory.is_dir():
        #     for item in directory.iterdir():
        #         if item.is_file() and item.suffix in INCLUDE_FILES:
        #             yield item

        for root, dirs, files in os.walk(directory):
            for file in files:
                full_path = Path(os.path.join(root, file))
                if full_path.suffix not in INCLUDE_FILES:
                    continue
                yield full_path

    @staticmethod
    def move_to_dir(file_path: Path, dir_path: Path):
        try:
            if not dir_path.is_dir():
                dir_path.mkdir(parents=True, exist_ok=True)
            shutil.move(file_path, dir_path / file_path.name)
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
    def clear_dir(dir_path: Path):
        """
        删除目录下的所有文件
        :param dir_path:
        :return:
        """
        try:
            if dir_path.is_dir():
                shutil.rmtree(dir_path)
            if dir_path.is_file():
                dir_path.unlink()
        except Exception as e:
            print(f"[警告] 清空目录出错 [{dir_path}] {e}")

    @staticmethod
    def compress_files_to_zip(file_paths: list[Path], zip_file_name: Path, root_dir: Path):
        """
        将指定文件路径列表中的文件压缩成一个 ZIP 文件。
        :param root_dir:
        :param file_paths: 文件路径列表，每个元素为文件的绝对或相对路径
        :param zip_file_name: 生成的 ZIP 文件的名称
        """
        with zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in file_paths:
                if file_path.is_file():
                    # 将文件添加到 ZIP 文件中，保留原文件结构
                    zipf.write(file_path, arcname=file_path.relative_to(root_dir))
                else:
                    print(f"警告: 文件 {file_path} 不存在，跳过。")

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

    @staticmethod
    def compress_media(file_path: Path):
        to_dir = file_path.parent / (file_path.name + '.compress')
        audio = AudioSegment.from_file(file_path)
        audio.export(to_dir, format="mp3", bitrate="8k")
        file_path.unlink()
        to_dir.rename(file_path)
        return to_dir


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
        print('[警告] 目录解析失败，使用默认目录 "月考试题"')
        return '月考试题'