# file_processor.py
import os
from pathlib import Path
from typing import Generator, Set
from config import Config

class FileProcessor:
    """文件处理工具类"""
    
    @staticmethod
    def get_all_files(directory: str) -> Generator[str, None, None]:
        """获取目录下所有支持的文件"""
        for root, _, files in os.walk(directory):
            for file in files:
                if Config.TEMP_FILE_PATTERN.match(file):
                    continue
                path = Path(root) / file
                if path.suffix.lower() in Config.SUPPORTED_EXTS:
                    yield str(path)

    @staticmethod
    def change_extension(path: str, new_ext: str = ".txt") -> str:
        """修改文件扩展名"""
        return str(Path(path).with_suffix(new_ext))
