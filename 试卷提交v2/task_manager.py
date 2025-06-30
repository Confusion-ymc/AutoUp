import queue
import threading
import time
import uuid
from collections import OrderedDict
from pathlib import Path

from utils import Tools, LogHandler


class TaskManager:
    def __init__(self):
        self.THREAD_COUNT = 1
        self.LISTEN_DIR = True
        self.UPLOAD_DIR = Tools.init_dir('待上传')
        self.FAILED_DIR = Tools.init_dir('失败')
        self.REPEAT_DIR = Tools.init_dir('重复')
        self.SUCCESS_DIR = Tools.init_dir('上传成功')
        self.TEMP_DIR = Path('temp')

        self.task_queue = queue.Queue()
        self.task_set = set([])
        self.SUCCESS_COUNT = []
        self.FAILED_COUNT = []
        self.REPEAT_COUNT = []
        self.RUNNING_COUNT = []
        self.TASK_DICT = OrderedDict()
        self.GRADE_MAP = Tools.load_keywords('关键词.xls', 'Sheet1')
        self.CATALOG_MAP = Tools.load_keywords('关键词.xls', 'Sheet2')
        self.CLASS_MAP = [
            "数学",
            "语文",
            "英语",
            "奥数",
            "科学",
            "道德与法治",
            "物理",
            "化学",
            "生物",
            "地理",
            "政治",
            "历史",
            "信息",
            "通用",
            "高等数学",
            "概率论与数理统计",
            "线性代数",
            "信息技术",
        ]
        self.logger = LogHandler()
        self.logger.update_logs()
        self.is_changed_listen_dir = False

    def change_listen_dir(self, dir_path: Path):
        self.UPLOAD_DIR = dir_path
        self.is_changed_listen_dir = True

    def start(self):
        def load_tasks(self: TaskManager):
            while True:
                if self.LISTEN_DIR:
                    # 检测上传文件是否有变动 有就把文件加入队列
                    for file_path in Tools.list_all_files(self.UPLOAD_DIR):
                        if file_path.name not in self.task_set:
                            self.task_set.add(file_path.name)
                            # 格式: {task_id: {'filename': '', 'status': '', 'thread': '', 'start_time': '', 'end_time': '', 'error': ''}}
                            task_id = uuid.uuid4().__str__()
                            self.TASK_DICT[task_id] = {
                                'file_path': file_path,
                                'filename': file_path.name,
                                'status': '',
                                'thread': '',
                                'start_time': '',
                                'end_time': '',
                                'error': '',
                                'status_changed': True
                            }
                            self.task_queue.put(task_id)
                time.sleep(2)
                if self.is_changed_listen_dir:
                    self.task_queue = queue.Queue()
                    self.task_set = set([])
                    self.TASK_DICT = OrderedDict()
                    self.is_changed_listen_dir = False

        threading.Thread(target=load_tasks, daemon=True, args=(self,)).start()

task_manager = TaskManager()