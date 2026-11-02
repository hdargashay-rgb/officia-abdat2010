"""
Agent Loop - الحلقة الذكية للوكيل
تربط بين النموذج اللغوي و Executor مع التصحيح الذاتي
متوافق تماماً مع Officia v2.1 و formatter.py
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
        return f"""أنت وكيل ذكي متخصص في إنشاء المستندات المكتبية الاحترافية باللغة العربية.

مهمتك: كتابة كود Python كامل لإنشاء الملفات المطلوبة.

==========================================
القواعد الإلزامية:
==========================================

1. **استخدم مكتبة التنسيق الاحترافي** (موجودة مسبقاً):
   - للخطابات: `from core.formatter import create_professional_letter`
   - للتقارير: `from core.formatter import create_excel_report`
   - للفواتير: `from core.formatter import create_invoice`

2. **مسار الحفظ**: استخدم `output_dir="{WORKSPACE_DIR}"` في جميع الدوال.

3. **الصياغة الرسمية للخطابات**:
   - استخدم عبارات مثل: "بالإشارة إلى الموضوع أعلاه"، "نرجو التكرم بـ"، "وتفضلوا بقبول فائق الاحترام والتقدير"
   - بنية الخطاب: عنوان → تاريخ → مرسل إليه → موضوع → تحية → مقدمة → محتوى → خاتمة → توقيع
   - استخدم جداول منسقة للبيانات الرقمية

4. **التنسيق الاحترافي**:
   - ألوان رسمية: أزرق داكن للعناوين
   - خطوط Calibri بحجم 12
   - محاذاة عربية صحيحة (RTL)
   - جداول برؤوس ملونة

5. **اطبع اسم الملف المنشأ** في النهاية باستخدام:
```python
   print(f"SUCCESS: {{output_path}}")
```

6. **المكتبات المسموحة**:
   openpyxl, pandas, numpy, python-docx, Pillow, os, sys, math, datetime, json, re, pathlib, shutil, tempfile, core.formatter

7. **لا تستخدم**: input(), subprocess, os.system, eval(), exec()

==========================================
أمثلة توضيحية:
==========================================

**مثال 1: إنشاء خطاب رسمي**
```python
from core.formatter import create_professional_letter

output_path = create_professional_letter(
    recipient="شركة ردينت للعناية بالسيارات",
    subject="مطالبة بمتأخرات الأجرة الشهرية",
    content_paragraphs=[
        "بالإشارة إلى الموضوع أعلاه، نتقدم إليكم بهذا الخطاب للمطالبة بالمستحقات المتأخرة.",
        "حيث أن إجمالي المبلغ المستحق يبلغ 120,000 ريال، عبارة عن أربعة شيكات شهرية.",
        "نرجو التكرم بسداد المبلغ المستحق في فترة أقصاها أسبوع من تاريخ هذا الخطاب."
    ],
    table_headers=["الشهر", "رقم الشيك", "القيمة", "البنك"],
    table_data=[
        ["يوليو", "0000611", "30000", "دخان"],
        ["أغسطس", "0000612", "30000", "دخان"],
        ["سبتمبر", "0000613", "30000", "دخان"],
        ["أكتوبر", "0000614", "30000", "دخان"]
    ],
    table_caption="تفاصيل الشيكات المستحقة",
    closing_note="يرجى سداد المستحقات في فترة أقصاها أسبوع من تاريخه",
    sender_name="المحامي",
    sender_organization="مكتب المحاماة",
    output_dir="{WORKSPACE_DIR}"
)
print(f"SUCCESS: {{output_path}}")
```

**مثال 2: إنشاء تقرير Excel**
```python
from core.formatter import create_excel_report

headers = ["اسم الموظف", "القسم", "الراتب الأساسي", "البدلات", "الإجمالي"]
data = [
    ["أحمد محمد", "المبيعات", 5000, 1000, 6000],
    ["محمد علي", "المحاسبة", 6000, 1200, 7200],
    ["فاطمة حسن", "الموارد البشرية", 5500, 1100, 6600],
    ["خالد سعيد", "تقنية المعلومات", 7000, 1400, 8400],
    ["نورة أحمد", "الإدارة", 8000, 1600, 9600]
]

output_path = create_excel_report(
    headers=headers,
    data=data,
    title="تقرير رواتب الموظفين - نوفمبر 2026",
    sheet_name="الرواتب",
    add_chart=True,
    freeze_panes=True,
    auto_filter=True,
    output_dir="{WORKSPACE_DIR}"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
الآن، نفذ الطلب التالي:
==========================================

{user_request}

اكتب كود Python كامل ومنسق. ابدأ بـ ```python وانهِ بـ ```"""

    def _build_fix_prompt(self, original_request: str, previous_code: str, error: str) -> str:
        return f"""❌ الكود السابق فشل في التنفيذ.

الطلب الأصلي:
{original_request}

الكود الذي فشل:
```python
{previous_code}
```

رسالة الخطأ:
{error}

==========================================
تعليمات التصحيح:
==========================================
1. حلل الخطأ بعناية
2. تأكد من استخدام `from core.formatter import ...` بشكل صحيح
3. تأكد من تمرير `output_dir="{WORKSPACE_DIR}"` في دالة الإنشاء
4. تأكد من طباعة اسم الملف في النهاية: `print(f"SUCCESS: {{output_path}}")`
5. لا تغير المنطق الأساسي، فقط صحح الخطأ

أعد كتابة الكود المُصحح بالكامل. ابدأ بـ ```python وانهِ بـ ```"""

    def _extract_code(self, response: str) -> str:
        """استخراج كود Python من رد النموذج"""
        # محاولة استخراج الكود من بين علامات ```python
        if '```python' in response:
            start = response.find('```python') + 9
            end = response.find('```', start)
            if end != -1:
                return response[start:end].strip()
        
        # محاولة بديلة: البحث عن ``` فقط
        if '```' in response:
            start = response.find('```') + 3
            end = response.find('```', start)
            if end != -1:
                return response[start:end].strip()
        
        # إذا لم توجد علامات، نعتبر أن الرد كله كود
        return response.strip()
