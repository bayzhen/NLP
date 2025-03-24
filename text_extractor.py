# text_extractor.py
import pdfplumber
import openpyxl
import win32com.client as win32
from docx import Document
from typing import Any, Dict
from contextlib import contextmanager
from config import Config
from pathlib import Path
from datetime import datetime


class TextExtractor:
    """多格式文本提取器"""

    @classmethod
    def extract(cls, file_path: str) -> str:
        """统一入口方法"""
        ext = Path(file_path).suffix.lower()
        handler = getattr(cls, f"_handle_{ext[1:]}", cls._handle_unsupported)
        return handler(file_path)

    @staticmethod
    def _handle_unsupported(file_path: str) -> str:
        print("不支持的文件：", file_path)
        return ""

    @staticmethod
    def _handle_txt(file_path: str) -> str:
        """处理纯文本文件"""
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    @staticmethod
    def _handle_pdf(file_path: str) -> str:
        """处理PDF文件"""
        text = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                if text_content := page.extract_text():
                    text.append(text_content)
        return "\n".join(text)

    @classmethod
    def _handle_doc(cls, file_path: str) -> str:
        """处理旧版Word文档"""
        with cls._word_context() as word:
            doc = word.Documents.Open(file_path)
            content = doc.Content.Text
            doc.Close()
            return content

    @staticmethod
    def _handle_docx(file_path: str) -> str:
        """处理新版Word文档"""
        return "\n".join(p.text for p in Document(file_path).paragraphs)

    @classmethod
    def _handle_xls(cls, file_path: str) -> str:
        """处理旧版Excel文件"""
        # 使用xlrd实现（需补充）
        ...

    @classmethod
    def _handle_xlsx(cls, file_path: str) -> str:
        """处理新版Excel文件"""
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        xlsx_str = ""
        xlsx_dict = cls._process_excel(workbook)
        for key in xlsx_dict.keys():
            value = xlsx_dict[key]
            xlsx_str += key + "\n"
            xlsx_str += value + "\n"
        return xlsx_str

    @classmethod
    def _process_excel(cls, workbook: Any) -> Dict[str, str]:
        """通用Excel处理逻辑"""
        results = {}
        for sheet in workbook.worksheets:
            lines = []
            for row in sheet.iter_rows():
                cells = [cls._format_cell(cell) for cell in row]
                if any(cells):
                    lines.append(Config.CELL_DELIMITER.join(cells))
            results[sheet.title] = Config.LINE_DELIMITER.join(lines)
        return results

    @staticmethod
    def _format_cell(cell: Any) -> str:
        """格式化单元格内容"""
        value = cell.value
        if value is None:
            return ""
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return str(value).strip()

    @contextmanager
    def _word_context():
        """Word应用程序上下文管理器"""
        try:
            word = win32.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
            yield word
        finally:
            word.Quit()
