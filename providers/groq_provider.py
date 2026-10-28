"""
مزود Groq - باستخدام OpenAI-compatible API
"""
import os
from dotenv import load_dotenv
from groq import Groq
from providers.base import BaseProvider


class GroqProvider(BaseProvider):
    """التعامل مع نماذج Groq"""

    def __init__(self, model_name: str = None):
        load_dotenv()

        api_key = os.getenv("GROQ_API_KEY")
        if not api_key:
            raise ValueError("مفتاح GROQ_API_KEY غير موجود في ملف .env")

        self.client = Groq(api_key=api_key)

        if model_name is None:
            from core.config import MODELS
            model_name = MODELS['groq']['default']

        self.model_name = model_name
        print("تم تهيئة GroqProvider باستخدام نموذج: " + model_name)

    def generate(self, prompt: str) -> str:
        try:
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            return response.choices[0].message.content
        except Exception as e:
            return "خطأ في الاتصال بـ Groq: " + str(e)

    def is_available(self) -> bool:
        try:
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {"role": "user", "content": "قل: اختبار"}
                ]
            )
            return response.choices[0].message.content is not None
        except Exception:
            return False
