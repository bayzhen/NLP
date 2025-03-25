# summary_generator.py
from snownlp import SnowNLP
from typing import Dict, Optional
from config import Config
from concurrent.futures import ThreadPoolExecutor, as_completed
from file_processor import FileProcessor
from text_extractor import TextExtractor
from pathlib import Path
from translator import Translator
import re
import os


class SummaryGenerator:
    """智能摘要生成器"""

    def __init__(self, translator: Translator):
        self.translator: Translator = translator
        self.error_files = {}
        self.source_path = ""

    def process_files(self, source_path: str, target_path: str) -> None:
        """批量处理文件"""
        self.source_path = source_path
        output_dir = Path(target_path) / Config.OUTPUT_DIR
        output_dir.mkdir(parents=True, exist_ok=True)

        with ThreadPoolExecutor(max_workers=Config.MAX_WORKERS) as executor:
            futures = {
                executor.submit(self._process_single_file, file, output_dir)
                for file in FileProcessor.get_all_files(source_path)
            }
            for future in as_completed(futures):
                future.result()

    def _process_single_file(self, file_path: str, output_dir: Path) -> None:
        """处理单个文件"""
        try:
            print(file_path)
            text = TextExtractor.extract(file_path)
            analysis = self._analyze_text(text)
            translations = self._generate_translations(file_path, analysis)
            self._save_results(file_path, output_dir, analysis, translations)
        except Exception as e:
            self.error_files[file_path] = str(e)

    def _analyze_text(self, text: str) -> Dict[str, list]:
        """执行文本分析"""
        cleaned = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9\s,\.!?，。！？]", "", text)
        s = SnowNLP(cleaned)
        return {
            "keywords": s.keywords(Config.KEYWORDS_LIMIT),
            "summary": s.summary(Config.SUMMARY_LENGTH),
        }

    def _generate_translations(self, file_path: str, analysis: Dict) -> Dict:
        """生成翻译内容"""
        return {
            "filename": self._safe_translate(Path(file_path).name),
            "keywords": ",".join(
                filter(None, (self._safe_translate(kw) for kw in analysis["keywords"]))
            ),
            "summary": self._safe_translate(",".join(analysis["summary"])),
        }

    def _safe_translate(self, text: str) -> str:
        """带安全校验的翻译"""
        return (
            self.translator.translate(text) if self.translator.is_english(text) else ""
        )

    def _save_results(
        self, src_path: str, output_dir: Path, analysis: Dict, translations: Dict
    ) -> None:
        content = [
            f"原文路径:{src_path}",
            self._format_section("文件名称翻译", translations["filename"]),
            self._format_section(
                "关键词", analysis["keywords"], translations["keywords"]
            ),
            self._format_section("摘要", analysis["summary"], translations["summary"]),
        ]

        # # 生成相对路径并创建目录结构
        # src_path_obj = Path(src_path)
        # # 获取相对于源目录的相对路径
        # relative_path = src_path_obj.relative_to(Path(self.source_path).resolve())
        # # 保持原有目录结构但去除源目录前缀
        # output_file = output_dir / relative_path.with_suffix(".txt")
        # # 创建必要的父目录
        # output_file.parent.mkdir(parents=True, exist_ok=True)

        # file_name = os.path.basename(src_path)
        # target_file_name = FileProcessor.change_extension(file_name)
        # output_file = os.path.join(output_dir, target_file_name)

        src_path_obj = Path(src_path)
        # 直接获取文件名并替换扩展名
        output_file = (output_dir / src_path_obj.name).with_suffix(".txt")
        # 确保输出目录存在
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text("\n".join(filter(None, content)), encoding="utf-8")

    @staticmethod
    def _format_section(title: str, content: list, translation: str = "") -> str:
        """格式化内容段落"""
        parts = [f"{title}:{content}"]
        if translation:
            parts.append(f"{title}翻译:\n{translation}")
        return "\n".join(parts)
