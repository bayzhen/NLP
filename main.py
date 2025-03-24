# main.py
from datetime import datetime
from translator import Translator
from summary_generator import SummaryGenerator
from pathlib import Path
from config import Config
from typing import Dict

class DocumentProcessor:
    """文档处理流水线"""
    
    def __init__(self, translator=None):
        self.translator = translator or Translator()
        self.summary_gen = SummaryGenerator(self.translator)
        self.file_mgr = FileManager()

    def process(self, source_path: str, target_path: str) -> None:
        """执行完整处理流程"""
        start_time = datetime.now()
        self.summary_gen.process_files(source_path, target_path)
        mid_time = datetime.now()
        
        self.file_mgr.combine_results(target_path)
        self.file_mgr.save_errors(self.summary_gen.error_files, target_path)
        
        print(f"处理完成\n时间统计:"
              f"\n- 开始: {start_time}"
              f"\n- 分析完成: {mid_time}"
              f"\n- 总计耗时: {datetime.now() - start_time}")

class FileManager:
    """文件管理工具"""
    
    @staticmethod
    def combine_results(target_path: str) -> None:
        """合并所有结果文件"""
        source_dir = Path(target_path) / Config.OUTPUT_DIR
        combined_file = Path(target_path) / Config.COMBINED_FILENAME
        
        with combined_file.open("w", encoding="utf-8") as output:
            for file in sorted(source_dir.glob("*.txt")):
                content = file.read_text(encoding="utf-8")
                output.write(f"{content}\n\n{'*'*90}\n")

    @staticmethod
    def save_errors(error_dict: Dict, target_path: str) -> None:
        """保存错误日志"""
        error_file = Path(target_path) / Config.ERROR_FILENAME
        content = [f"文件: {path}\n错误: {err}\n" for path, err in error_dict.items()]
        error_file.write_text("\n".join(content), encoding="utf-8")

if __name__ == "__main__":
    processor = DocumentProcessor()
    processor.process(Config.SOURCE_FOLDER, Config.TARGET_FOLDER)
