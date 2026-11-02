"""
Agent Loop - الحلقة الذكية للوكيل
"""
from core.llm_client import LLMClient
from core.executor import Executor
from core.config import MAX_AUTO_FIX_ATTEMPTS, WORKSPACE_DIR


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
            result = self.executor.run_python(clean_code)
            
            if result['success']:
                print(f"\n✅ نجحت المهمة في المحاولة {attempt}!")
                return {
                    'success': True, 
                    'attempts': attempt, 
                    'final_output': result['output'], 
                    'files_created': result.get('files_created', []), 
                    'code': clean_code
                }
            
            print(f"  -> فشل التنفيذ: {result['error'][:100]}")
            if attempt < self.max_attempts:
                code_prompt = self._build_fix_prompt(user_request, clean_code, result['error'])
            else:
                print(f"\n❌ فشلت المهمة بعد {self.max_attempts} محاولات")
                return {
                    'success': False, 
                    'attempts': attempt, 
                    'final_output': result['output'], 
                    'error': result['error'], 
                    'files_created': result.get('files_created', []), 
                    'code': clean_code
                }

    def _build_code_prompt(self, user_request: str) -> str:
        return f"""أنت كاتب خطابات محترف ومبرمج Python خبير.

⚠️ قاعدة إلزامية واحدة فقط:
يجب أن تستخدم الدوال الموجودة في core.formatter فقط.
ممنوع اختراع أسماء دوال غير موجودة.

📚 الدوال المتاحة حصرياً:

للخطابات (استخدم واحدة فقط):
- generate_professional_letter(...) ← دالة شاملة لإنشاء خطاب كامل
- create_professional_letter(...) ← دالة بديلة

لـ Excel:
- create_multi_sheet_excel(sheets_config, output_dir)
- create_fixed_asset_report(company, year, sections, output_dir)
- create_professional_invoice(company, client, number, date, items, output_dir)

==========================================
📝 مثال صحيح 100% (انسخ النمط):
==========================================
```python
from core.formatter import generate_professional_letter

output_path = generate_professional_letter(
    organization_name="شركة ردينت للعناية بالسيارات",
    title="مطالبة بمتأخرات الأجرة الشهرية",
    recipient="السيد أحمد محمد",
    subject="مطالبة بمتأخرات الأجرة الشهرية",
    paragraphs=[
        "بالإشارة إلى الموضوع أعلاه، نتقدم إليكم بهذا الخطاب للمطالبة بالمستحقات المتأخرة.",
        "حيث أن إجمالي المبلغ المستحق يبلغ 120,000 ريال، عبارة عن أربعة شيكات شهرية لم يتم صرفها.",
        "نرجو التكرم بسداد المبلغ المستحق في فترة أقصاها أسبوع من تاريخ هذا الخطاب."
    ],
    table_headers=["الشهر", "رقم الشيك", "القيمة (ريال)", "البنك"],
    table_data=[
        ["يوليو", "0000611", "30000", "دخان"],
        ["أغسطس", "0000612", "30000", "دخان"],
        ["سبتمبر", "0000613", "30000", "دخان"],
        ["أكتوبر", "0000614", "30000", "دخان"]
    ],
    table_caption="تفاصيل الشيكات المستحقة",
    signature_name="المحامي عبدالله",
    signature_title="المستشار القانوني",
    note="يرجى السداد خلال أسبوع من تاريخ الخطاب",
    output_dir="{WORKSPACE_DIR}"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
❌ ممنوع:
==========================================
- ممنوع: from docx import Document (استخدم generate_professional_letter)
- ممنوع: generate_document, create_letter, make_doc (أسماء غير موجودة)
- ممنوع: اختراع أي دالة غير مذكورة أعلاه

==========================================
الآن نفذ الطلب:
==========================================

{user_request}

اكتب كود Python يبدأ بـ from core.formatter import generate_professional_letter. ابدأ بـ ```python وانهِ بـ ```"""

    def _build_fix_prompt(self, original_request: str, previous_code: str, error: str) -> str:
        return f"""❌ الكود السابق فشل.

الطلب: {original_request}

الكود الذي فشل:
```python
{previous_code}
```

الخطأ: {error}

==========================================
تعليمات التصحيح:
==========================================
1. استخدم ONLY: from core.formatter import generate_professional_letter
2. لا تستخدم: generate_document, create_letter, make_doc
3. تأكد من تمرير جميع المعاملات المطلوبة
4. اطبع: print(f"SUCCESS: {{output_path}}")

أعد كتابة الكود كاملاً. ابدأ بـ ```python وانهِ بـ ```"""

    def _extract_code(self, response: str) -> str:
        if '```python' in response:
            start = response.find('```python') + 9
            end = response.find('```', start)
            if end != -1:
                return response[start:end].strip()
        if '```' in response:
            start = response.find('```') + 3
            end = response.find('```', start)
            if end != -1:
                return response[start:end].strip()
        return response.strip()
