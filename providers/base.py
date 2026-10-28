"""
الواجهة الأساسية لجميع مزودي النماذج اللغوية
"""
from abc import ABC, abstractmethod


class BaseProvider(ABC):
    """فئة أساسية يجب أن ترثها جميع المزودين"""

    @abstractmethod
    def generate(self, prompt: str) -> str:
        """إرسال طلب والحصول على رد نصي"""
        pass

    @abstractmethod
    def is_available(self) -> bool:
        """التحقق من أن المزود يعمل ومتاح"""
        pass
