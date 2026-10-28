"""
عميل النماذج اللغوية - الواجهة الموحدة مع 3 مزودين
"""
from providers.groq_provider import GroqProvider
from providers.openrouter_provider import OpenRouterProvider
from providers.gemini_provider import GeminiProvider
from core.config import MODELS


class LLMClient:
    """واجهة واحدة للتواصل مع 3 مزودين بالترتيب: Groq → OpenRouter → Gemini"""

    def __init__(self):
        self.providers = [
            ('Groq', GroqProvider(model_name=MODELS['groq']['default'])),
            ('OpenRouter', OpenRouterProvider(model_name=MODELS['openrouter']['default'])),
            ('Gemini', GeminiProvider(model_name=MODELS['gemini']['default']))
        ]
        print("تم تهيئة LLMClient مع 3 مزودين")

    def ask(self, prompt: str) -> str:
        """إرسال سؤال مع محاولة تلقائية عبر المزودين بالترتيب"""
        for name, provider in self.providers:
            print(f"\n[محاولة] باستخدام: {name}")
            try:
                response = provider.generate(prompt)
                if not response.startswith("خطأ في الاتصال"):
                    print(f"  -> نجح الاتصال بـ {name}!")
                    return response
                print(f"  -> فشل، الانتقال للتالي...")
            except Exception as e:
                print(f"  -> استثناء: {str(e)[:50]}... الانتقال للتالي")
                continue
        
        return "فشل جميع المزودين في الاتصال"

    def check_connection(self) -> bool:
        """التحقق من أن مزوداً واحداً على الأقل يعمل"""
        for name, provider in self.providers:
            if provider.is_available():
                return True
        return False
