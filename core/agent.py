"""
Agent Loop - الحلقة الذكية للوكيل
تربط بين النموذج اللغوي و Executor مع التصحيح الذاتي
"""
from core.llm_client import LLMClient
from core.executor import Executor
from core.config import MAX_AUTO_FIX_ATTEMPTS


class Agent:
    def __init__(self):
        self.llm = LLMClient()
        self.executor = Executor()
        self.max_attempts = MAX_AUTO_FIX_ATTEMPTS
        print("تم تهيئة Agent - الوكيل الذكي")

    def execute_task(self, user_request: str) -> dict:
        print("\n" + "=" * 60)
        print("طلب المستخدم: " + user_request)
        print("=" * 60)

        code_prompt = self._build_code_prompt(user_request)
        
        for attempt in range(1, self.max_attempts + 1):
            print(f"\n[محاولة {attempt}/{self.max_attempts}]")
            print("  -> جاري توليد الكود...")
            generated_code = self.llm.ask(code_prompt)
            clean_code = self._extract_code(generated_code)
            
            if not clean_code:
                print("  -> فشل استخراج الكود، إعادة المحاولة...")
                code_prompt = self._build_fix_prompt(user_request, generated_code, "لم يتم استخراج كود Python صالح")
                continue

            print("  -> جاري تنفيذ الكود عبر run_python...")
            # استخدام الدالة الجديدة المتفق عليها
            result = self.executor.run_python(clean_code)
            
            if result['success']:
                print(f"\n✅ نجحت المهمة في المحاولة {attempt}!")
                return {'success': True, 'attempts': attempt, 'final_output': result['output'], 'files_created': result['files_created'], 'code': clean_code}
            
            print(f"  -> فشل التنفيذ: {result['error'][:100]}")
            if attempt < self.max_attempts:
                code_prompt = self._build_fix_prompt(user_request, clean_code, result['error'])
            else:
                print(f"\n❌ فشلت المهمة بعد {self.max_attempts} محاولات")
                return {'success': False, 'attempts': attempt, 'final_output': result['output'], 'error': result['error'], 'files_created': result['files_created'], 'code': clean_code}

    def _build_code_prompt(self, user_request: str) -> str:
        return f"""أنت مبرمج Python خبير. اكتب كود Python لتنفيذ: {user_request}
القواعد:
1. استخدم فقط: openpyxl, pandas, numpy, os, sys, math, datetime, json, re, pathlib, shutil, tempfile
2. اطبع النتائج بوضوح باستخدام print()
3. لا تستخدم input()
4. أعد فقط كود Python بدون أي شرح. ابدأ بـ ```python وانهِ بـ ```"""

    def _build_fix_prompt(self, original_request: str, previous_code: str, error: str) -> str:
        return f"""الكود فشل. الطلب: {original_request}
الكود السابق:
```python
{previous_code}
```
الخطأ: {error}
أعد فقط كود Python المُصحح بدون أي شرح. ابدأ بـ ```python وانهِ بـ ```"""

    def _extract_code(self, response: str) -> str:
        if '```python' in response:
            start = response.find('```python') + 9
            end = response.find('```', start)
            if end != -1: return response[start:end].strip()
        return response.strip()
