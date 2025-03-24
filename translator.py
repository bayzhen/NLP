# translator.py
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer
from typing import Optional, Tuple
from config import Config

class Translator:
    """多语言翻译处理器"""
    
    def __init__(self):
        self.model, self.tokenizer = self._load_model()

    def _load_model(self) -> Tuple[AutoModelForSeq2SeqLM, AutoTokenizer]:
        try:
            tokenizer = AutoTokenizer.from_pretrained(Config.TRANSLATION_MODEL)
            model = AutoModelForSeq2SeqLM.from_pretrained(Config.TRANSLATION_MODEL)
            return model, tokenizer
        except Exception as e:
            raise RuntimeError(
                f"模型加载失败，请执行以下命令下载：\n"
                f"from transformers import AutoTokenizer, AutoModelForSeq2SeqLM\n"
                f"AutoTokenizer.from_pretrained('{Config.TRANSLATION_MODEL}')\n"
                f"AutoModelForSeq2SeqLM.from_pretrained('{Config.TRANSLATION_MODEL}')"
            ) from e

    def translate(self, text: str) -> Optional[str]:
        """执行翻译操作"""
        try:
            formatted_text = f">>zho_Hans<< {text}"
            inputs = self.tokenizer(
                text[:Config.TRANSLATION_MAX_LENGTH],
                return_tensors="pt",
                padding=True,
                truncation=True
            )
            outputs = self.model.generate(
                **inputs,
                max_length=Config.TRANSLATION_MAX_LENGTH,
                num_beams=5,
                early_stopping=True,
                repetition_penalty=1.5,
                no_repeat_ngram_size=2
            )
            result = self.tokenizer.decode(outputs[0], skip_special_tokens=True)
            return result
        except Exception as e:
            print(f"翻译失败: {str(e)}")
            return None

    @staticmethod
    def is_english(text: str) -> bool:
        """判断是否为英文内容"""
        return bool(Config.ENGLISH_PATTERN.match(text))
