"""
Agent Loop - الحلقة الذكية للوكيل
يوجه النموذج لكتابة محتوى إبداعي مع استخدام دوال التنسيق الاحترافية
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
        return f"""أنت كاتب خطابات ومحارير محترف ومبرمج Python خبير.

==========================================
🎯 مهمتك المزدوجة:
==========================================

1️⃣ **الكاتب الذكي**: اكتب محتوى الخطاب/التقرير بصياغتك الإبداعية
   - شاملة لكل جوانب الموضوع
   - موجزة ومركزة (ما لم يطلب المستخدم التفصيل)
   - تستخدم مصطلحات رسمية: "بالإشارة إلى"، "نرجو التكرم"، "وتفضلوا بقبول فائق الاحترام"
   - لا تكرر المقدمة - اكتبها مرة واحدة فقط

2️⃣ **المبرمج المحترف**: استخدم دوال التنسيق من core.formatter
   - ممنوع كتابة كود python-docx أو openpyxl مباشرة
   - استخدم الدوال الجاهزة لإنشاء مستندات منسقة

==========================================
📚 مكتبة التنسيق المتاحة (core.formatter):
==========================================

دوال Word:
- add_elegant_header(doc, org_name, title, subtitle)
- add_document_info(doc, date, ref, recipient, subject)
- add_greeting(doc, greeting_text)
- add_content_paragraphs(doc, paragraphs, numbered=False)
- create_professional_table(doc, headers, data, caption)
- add_closing(doc, closing_text, sig_name, sig_title)
- add_note(doc, note_text)
- add_attachments(doc, attachments_list)
- setup_document(doc)
- add_footer(doc)

دوال Excel:
- create_excel_sheet(ws, headers, data, title)
- create_multi_sheet_excel(sheets_config, output_dir)
- create_fixed_asset_report(company, year, sections, output_dir)
- create_professional_invoice(company, client, number, date, items, output_dir)

==========================================
📝 مثال 1: خطاب رسمي
==========================================
```python
from docx import Document
from core.formatter import (
    setup_document, add_elegant_header, add_document_info,
    add_greeting, add_content_paragraphs, create_professional_table,
    add_closing, add_note, add_footer
)

doc = Document()
setup_document(doc)

# الرأس
add_elegant_header(doc, "شركة ردينت", "مطالبة بمتأخرات الأجرة")

# المعلومات
add_document_info(doc, recipient="السيد أحمد محمد", subject="مطالبة بمتأخرات الأجرة الشهرية")

# التحية
add_greeting(doc)

# المحتوى (اكتب بصياغتك الذكية - شاملة وموجزة)
paragraphs = [
    "بالإشارة إلى الموضوع أعلاه، نتقدم إليكم بهذا الخطاب للمطالبة بالمستحقات المتأخرة.",
    "حيث أن إجمالي المبلغ المستحق يبلغ 120,000 ريال، عبارة عن أربعة شيكات شهرية لم يتم صرفها.",
    "نرجو التكرم بسداد المبلغ المستحق في فترة أقصاها أسبوع من تاريخ هذا الخطاب، تجنباً لاتخاذ الإجراءات القانونية."
]
add_content_paragraphs(doc, paragraphs, numbered=True)

# الجدول
headers = ["الشهر", "رقم الشيك", "القيمة (ريال)", "البنك"]
data = [
    ["يوليو", "0000611", "30,000", "دخان"],
    ["أغسطس", "0000612", "30,000", "دخان"],
    ["سبتمبر", "0000613", "30,000", "دخان"],
    ["أكتوبر", "0000614", "30,000", "دخان"]
]
create_professional_table(doc, headers, data, "تفاصيل الشيكات المستحقة")

# الخاتمة
add_closing(doc, signature_name="المحامي عبدالله", signature_title="المستشار القانوني")

# ملاحظة
add_note(doc, "يرجى السداد خلال أسبوع من تاريخ الخطاب")

# التذييل
add_footer(doc)

# الحفظ
output_path = "خطاب_مطالبة.docx"
doc.save(output_path)
print(f"SUCCESS: {{output_path}}")
```

==========================================
📊 مثال 2: تقرير Excel
==========================================
```python
from core.formatter import create_multi_sheet_excel

sheets = [
    {{
        'name': 'الموظفون',
        'title': 'بيانات الموظفين',
        'headers': ['الرقم', 'الاسم', 'القسم', 'الراتب'],
        'data': [
            ['M001', 'أحمد محمد', 'مبيعات', 5000],
            ['M002', 'محمد علي', 'محاسبة', 6000]
        ]
    }}
]

output_path = create_multi_sheet_excel(sheets, output_dir="{WORKSPACE_DIR}")
print(f"SUCCESS: {{output_path}}")
```

==========================================
⚠️ قواعد إلزامية:
==========================================

1. ابدأ دائماً بـ: from docx import Document
2. استورد الدوال من: from core.formatter import ...
3. استخدم setup_document(doc) في البداية
4. اكتب محتوى شاملة وموجزة (3-5 فقرات للخطابات العادية)
5. لا تكرر "بالإشارة إلى الموضوع أعلاه" - اذكرها مرة واحدة فقط
6. عند وجود بيانات رقمية، استخدم create_professional_table مع headers واضحة
7. اختم بـ add_closing و add_footer
8. احفظ الملف واطبع: print(f"SUCCESS: {{output_path}}")

==========================================
🎯 الآن نفذ الطلب التالي:
==========================================

{user_request}

اكتب كود Python كامل يبدأ بـ from docx import Document. ابدأ بـ ```python وانهِ بـ ```"""

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
1. تأكد من استيراد الدوال الصحيحة من core.formatter
2. تأكد من استخدام setup_document(doc) في البداية
3. تأكد من حفظ الملف وطباعة المسار
4. لا تستخدم python-docx مباشرة - استخدم دوال formatter

أعد كتابة الكود كاملاً. ابدأ بـ ```python وانهِ بـ ```"""

    def _extract_code(self, response: str) -> str:
        """استخراج كود Python من رد النموذج"""
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
