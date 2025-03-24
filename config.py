# config.py
import re
from typing import Pattern, Set
import os

class Config:
    # 文件配置
    SUPPORTED_EXTS: Set[str] = {'.txt', '.doc', '.docx', '.xls', '.xlsx', '.pdf'}
    TEMP_FILE_PATTERN: Pattern = re.compile(r'^~\$')
    OUTPUT_DIR: str = "摘要文件列表"
    COMBINED_FILENAME: str = "!摘要文件总览.txt"
    ERROR_FILENAME: str = "!过滤文件总览.txt"
    
    # 处理参数
    SUMMARY_LENGTH: int = 5
    KEYWORDS_LIMIT: int = 10
    CELL_DELIMITER: str = " "
    LINE_DELIMITER: str = "\n"
    
    # 翻译模型
    TRANSLATION_MODEL: str = "Helsinki-NLP/opus-mt-en-zh"
    TRANSLATION_MAX_LENGTH: int = 512
    ENGLISH_PATTERN: Pattern = re.compile(r'^(?=.*[A-Za-z])[A-Za-z0-9\s.,!?&@#$%^*()\'"\-]+$')

    # 线程配置
    MAX_WORKERS: int = (os.cpu_count() or 2) * 2
    
    # 源目录
    SOURCE_FOLDER: str = r"D:\GitHub\NLP\source_files"
    # 目标目录
    TARGET_FOLDER: str = r"D:\GitHub\NLP\target_files"
    