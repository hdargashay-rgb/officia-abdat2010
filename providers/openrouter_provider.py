"""
مزود OpenRouter - بوابة نماذج متعددة
"""
import os
from dotenv import load_dotenv
from openai import OpenAI
from providers.base import BaseProvider


class OpenRouterProvider(BaseProvider):
    """التعامل مع نماذج OpenRouter"""

    def __init__(self, model_name: str = None):
        load_dotenv()

        api_key = os.getenv("OPENROUTER_API_KEY")
        if not api_key:
            raise ValueError("مفتاح OPENROUTER_API_KEY غير موجود في ملف .env")

        self.client = OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=api_key
        )

        if model_name is None:
            from core.config import MODELS
            model_name = MODELS['openrouter']['default']

        self.model_name = model_name
        print("تم تهيئة OpenRouterProvider باستخدام نموذج: " + model_name)

    def generate(self, prompt: str) -> str:
        try:
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {
                        "role": "system",
                        "content": "أنت مساعد ذكي يتحدث العربية الفصحى بوضوح ودقة."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ]
            )
            return response.choices[0].message.content
        except Exception as e:
            return "خطأ في الاتصال بـ OpenRouter: " + str(e)

    def is_available(self) -> bool:
        try:
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {"role": "user", "content": "قل: اختبار"}
                ],
                max_tokens=10
            )
            return response.choices[0].message.content is not None
        except Exception:
            return False
