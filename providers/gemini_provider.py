"""
مزود Google Gemini - باستخدام الحزمة الجديدة google-genai
"""
import os
from dotenv import load_dotenv
from google import genai
from providers.base import BaseProvider


class GeminiProvider(BaseProvider):
    """التعامل مع نموذج Gemini من Google باستخدام الحزمة الحديثة"""

    def __init__(self, model_name: str = None):
        load_dotenv()

        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("مفتاح GEMINI_API_KEY غير موجود في ملف .env")

        self.client = genai.Client(api_key=api_key)

        if model_name is None:
            from core.config import MODELS
            model_name = MODELS['gemini']['default']

        self.model_name = model_name
        print("تم تهيئة GeminiProvider باستخدام نموذج: " + model_name)

    def generate(self, prompt: str) -> str:
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt
            )
            return response.text
        except Exception as e:
            return "خطأ في الاتصال بـ Gemini: " + str(e)

    def is_available(self) -> bool:
        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents="قل: اختبار"
            )
            return response.text is not None
        except Exception:
            return False
