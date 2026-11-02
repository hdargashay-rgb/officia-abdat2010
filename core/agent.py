"""
Agent Loop - الحلقة الذكية للوكيل
يوجه النموذج لاستخدام formatter.py لإنشاء مستندات احترافية
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

==========================================
⚠️ قواعد إلزامية صارمة:
==========================================

1. **يجب استخدام مكتبة core.formatter** لإنشاء أي ملف:
   - للخطابات: `from core.formatter import create_professional_letter`
   - لتقارير الأصول الثابتة: `from core.formatter import create_fixed_asset_report`
   - لتقارير Excel متعددة الأوراق: `from core.formatter import create_multi_sheet_excel`
   - للفواتير: `from core.formatter import create_professional_invoice`

2. **مسار الحفظ**: استخدم دائماً `output_dir="{WORKSPACE_DIR}"`

3. **اطبع اسم الملف في النهاية**:
```python
   print(f"SUCCESS: {{output_path}}")
```

4. **المكتبات المسموحة**:
   core.formatter, openpyxl, pandas, numpy, python-docx, Pillow, os, sys, math, datetime, json, re, pathlib, shutil, tempfile

5. **ممنوع**: input(), subprocess, os.system, eval(), exec(), os.environ

==========================================
مثال 1: خطاب رسمي
==========================================
```python
from core.formatter import create_professional_letter

output_path = create_professional_letter(
    recipient="شركة ردينت للعناية بالسيارات",
    subject="مطالبة بمتأخرات الأجرة الشهرية",
    content_paragraphs=[
        "بالإشارة إلى الموضوع أعلاه، نتقدم إليكم بهذا الخطاب للمطالبة بالمستحقات المتأخرة.",
        "حيث أن إجمالي المبلغ المستحق يبلغ 120,000 ريال.",
        "نرجو التكرم بسداد المبلغ المستحق في فترة أقصاها أسبوع."
    ],
    table_headers=["الشهر", "رقم الشيك", "القيمة", "البنك"],
    table_data=[
        ["يوليو", "0000611", "30000", "دخان"],
        ["أغسطس", "0000612", "30000", "دخان"],
        ["سبتمبر", "0000613", "30000", "دخان"],
        ["أكتوبر", "0000614", "30000", "دخان"]
    ],
    table_caption="تفاصيل الشيكات المستحقة",
    closing_note="يرجى السداد خلال أسبوع",
    sender_name="المحامي عبدالله",
    sender_organization="مكتب المحاماة",
    output_dir="{WORKSPACE_DIR}"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
مثال 2: تقرير أصول ثابتة احترافي
==========================================
```python
from core.formatter import create_fixed_asset_report

sections = [
    {{
        'name': 'Hardware & Software',
        'depreciation_rate': 0.3333,
        'assets': [
            {{
                'sno': 2,
                'supplier': 'ATCOM TECHNOLOGY',
                'description': 'AVAYA IP TELEPHONE',
                'fa_no': '',
                'book_value': 5250.00,
                'addition': 0,
                'disposal': 0,
                'purchase_date': '3/24/22'
            }},
            {{
                'sno': 3,
                'supplier': 'THME HONEST',
                'description': 'LENOVO COMPUTER',
                'fa_no': '',
                'book_value': 5480.00,
                'addition': 0,
                'disposal': 0,
                'purchase_date': '3/28/22'
            }}
        ]
    }},
    {{
        'name': 'Tools & Equipment',
        'depreciation_rate': 0.20,
        'assets': [
            {{
                'sno': 1,
                'supplier': 'XPEL INC USA',
                'description': 'XPEL FILM CUTTING MACHINE',
                'fa_no': '',
                'book_value': 7861.00,
                'addition': 0,
                'disposal': 0,
                'purchase_date': '1/12/22'
            }}
        ]
    }}
]

output_path = create_fixed_asset_report(
    company_name="Radiant Car Care",
    year=2025,
    sections=sections,
    output_dir="{WORKSPACE_DIR}",
    filename="Fixed_Asset_2025"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
مثال 3: تقرير Excel متعدد الأوراق
==========================================
```python
from core.formatter import create_multi_sheet_excel

sheets_config = [
    {{
        'name': 'الموظفون',
        'title': 'بيانات الموظفين',
        'headers': ['رقم الموظف', 'الاسم', 'القسم', 'الراتب'],
        'data': [
            ['M001', 'أحمد محمد', 'مبيعات', 5000],
            ['M002', 'محمد علي', 'محاسبة', 6000],
            ['M003', 'فاطمة حسن', 'موارد بشرية', 5500]
        ]
    }},
    {{
        'name': 'الحضور',
        'title': 'سجل الحضور',
        'headers': ['رقم الموظف', 'أيام الحضور', 'أيام الغياب'],
        'data': [
            ['M001', 20, 2],
            ['M002', 22, 0],
            ['M003', 18, 4]
        ]
    }}
]

output_path = create_multi_sheet_excel(
    sheets_config,
    output_dir="{WORKSPACE_DIR}",
    filename="تقرير_الموظفين"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
مثال 4: فاتورة احترافية
==========================================
```python
from core.formatter import create_professional_invoice

items = [
    {{'description': 'خدمة غسيل سيارات', 'quantity': 5, 'unit_price': 100}},
    {{'description': 'خدمة تلميع', 'quantity': 2, 'unit_price': 250}},
    {{'description': 'خدمة حماية الطلاء', 'quantity': 1, 'unit_price': 1500}}
]

output_path = create_professional_invoice(
    company_name="Radiant Car Care",
    client_name="أحمد محمد",
    invoice_number="INV-2025-001",
    invoice_date="2025-11-03",
    items=items,
    tax_rate=0.05,
    notes="شكراً لتعاملكم معنا",
    output_dir="{WORKSPACE_DIR}",
    filename="فاتورة_001"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
الآن نفذ الطلب التالي:
==========================================

{user_request}

اكتب كود Python كامل يستخدم core.formatter. ابدأ بـ ```python وانهِ بـ ```"""

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
1. استخدم `from core.formatter import ...` (إحدى الدوال الأربع)
2. تأكد من `output_dir="{WORKSPACE_DIR}"`
3. اطبع: `print(f"SUCCESS: {{output_path}}")`
4. لا تكتب كود openpyxl أو python-docx مباشرة

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
