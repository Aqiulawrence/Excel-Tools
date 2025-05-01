import logging
import os
import atexit

from datetime import datetime
from logging.handlers import RotatingFileHandler

class AppLogger:
    def __init__(self, log_file='./logs/operation.log', max_bytes=10 * 1024 * 1024, backup_count=5):
        self.log_file = log_file
        self._setup_logging(max_bytes, backup_count)
        self._record_startup()
        atexit.register(self._record_shutdown)

    def _setup_logging(self, max_bytes, backup_count):
        # 创建日志目录如果不存在
        os.makedirs(os.path.dirname(self.log_file), exist_ok=True)

        # 设置日志格式
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        date_format = '%Y-%m-%d %H:%M:%S'

        # 配置日志文件处理器
        file_handler = RotatingFileHandler(
            self.log_file,
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding='utf-8'
        )
        file_handler.setFormatter(logging.Formatter(log_format, datefmt=date_format))

        # 配置logger
        self.logger = logging.getLogger('AppLogger')
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(file_handler)

    def _record_startup(self):
        startup_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.logger.info("=" * 60)
        self.logger.info(f"应用程序启动 - 时间: {startup_time}")
    def _record_shutdown(self):
        shutdown_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.logger.info(f"应用程序关闭 - 时间: {shutdown_time}")
        self.logger.info("=" * 60)

    def get_logger(self):
        return self.logger

# 日志装饰器
def log(func):
    from functools import wraps

    @wraps(func)
    def wrapper(*args, **kwargs):
        logger = logging.getLogger('AppLogger')

        # 获取调用信息
        func_name = func.__name__
        module_name = func.__module__

        # 记录开始调用
        logger.info(
            f"调用开始: {module_name}.{func_name} | "
            f"参数: {args if len(str(args)) < 100 else '<too long>'} | "
            f"关键字参数: {kwargs if len(str(kwargs)) < 100 else '<too long>'}"
        )

        try:
            # 执行函数
            result = func(*args, **kwargs)

            # 记录成功调用
            logger.info(
                f"调用成功: {module_name}.{func_name} | "
                f"返回: {result if len(str(result)) < 100 else '<too long>'}"
            )
            return result

        except Exception as e:
            # 记录调用失败
            logger.error(
                f"调用失败: {module_name}.{func_name} | "
                f"错误: {str(e)}",
                exc_info=False
            )
            raise

    return wrapper