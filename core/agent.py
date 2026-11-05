"""
Agent Loop - الحلقة الذكية للوكيل
يستخدم نظام الأنماط الديناميكي لتصنيف المستندات واختيار الأنماط المناسبة
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

    def execute_task(self, user_request: str, uploaded_file_path: str = None) -> dict:
        print("\n" + "=" * 60)
        print("طلب المستخدم: " + user_request)
        if uploaded_file_path:
            print("ملف مرفوع للمعالجة: " + uploaded_file_path)
        print("=" * 60)

        code_prompt = self._build_code_prompt(user_request, uploaded_file_path)
        
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

    def _build_code_prompt(self, user_request: str, uploaded_file_path: str = None) -> str:
        
        # ==========================================
        # السيناريو الأول: وجود ملف مرفوع (محاكاة أو تحليل)
        # ==========================================
        if uploaded_file_path:
            return f"""أنت وكيل ذكي متخصص في تحليل المستندات وإنشاء ملفات مشابهة.

📎 تم رفع ملف مرجعي: {uploaded_file_path}
🎯 طلب المستخدم: {user_request}

⚠️ قاعدة إلزامية: استخدم الدوال المخصصة للملفات المرفوعة من core.formatter.

📚 الدوال المتاحة:
1️⃣ لمحاكاة الملف (إعادة بنائه بنفس التنسيق والمعادلات):
- replicate_document(source_file='{uploaded_file_path}', output_dir='{WORKSPACE_DIR}')

2️⃣ لتنسيق تحليل النموذج الذكي:
- format_analysis_report(analysis_data, report_title, company_name, period, output_format='excel', style='accounting', output_dir='{WORKSPACE_DIR}')

==========================================
📝 مثال: محاكاة ملف مرفوع
==========================================
```python
from core.formatter import replicate_document

result = replicate_document(
    source_file='{uploaded_file_path}',
    output_dir='{WORKSPACE_DIR}'
)
print("SUCCESS: " + result['output_path'])
```

⚠️ تحذير: ممنوع اختراع دوال، استخدم فقط الدوال المذكورة أعلاه للملفات المرفوعة.
ابدأ بـ ```python وانهِ بـ ```"""

        # ==========================================
        # السيناريو الثاني: إنشاء مستند جديد من الصفر
        # ==========================================
        return f"""أنت كاتب محترف ومبرمج Python خبير متخصص في إنشاء مستندات احترافية.

⚠️ قاعدة إلزامية: استخدم الدوال من core.formatter فقط حسب نوع المستند المطلوب.

📚 الدوال المتاحة:
1️⃣ للخطابات والمذكرات بأنماط متنوعة:
- generate_styled_letter(organization_name, title, recipient, subject, paragraphs, table_headers, table_data, table_caption, header_style, content_style, table_style, footer_style, signature_style, signatures, output_dir)

2️⃣ للسيرة الذاتية المتوافقة مع ATS:
- generate_ats_cv(full_name, phone, email, location, linkedin, summary, experience, education, skills, languages, certifications, projects, interests, objective, output_dir)

3️⃣ للتقارير المالية الشاملة (أي نوع تقرير):
- generate_financial_report(report_type, company_name, period, sections, currency='ريال', include_totals=True, include_formulas=True, style='accounting', output_dir)
  (report_type: 'income_statement', 'balance_sheet', 'cash_flow_direct', 'trial_balance', 'sales_report', إلخ)
  (style: 'accounting', 'modern', 'classic')

4️⃣ لتقارير الأصول الثابتة والإهلاكات (سريعة ومخصصة):
- create_fixed_asset_report(company_name, year, sections, output_dir)
  (ملاحظة: sections هي قائمة من القواميس، كل قاموس يحتوي على: 'sheet_name', 'headers', 'data')

5️⃣ لجدول الإهلاك الشهري المضغوط (طريقة الرصيد المتناقص):
- generate_depreciation_schedule(company_name, year, assets, output_dir)
  (assets: قائمة قواميس، كل قاموس: 'asset_id', 'description', 'cost', 'accumulated_dep_start', 'depreciation_rate', 'salvage_value')

==========================================
🎨 الأنماط المتاحة للخطابات:
==========================================
📌 header_style: official, modern, professional, minimal, bold, gradient, celebration, islamic
📌 content_style: paragraphs, numbered, bullets, compact, mixed, quoted, highlighted, decorative
📌 table_style: royal_blue, green_finance, dark_elegant, colorful, minimal, accounting, modern_shadows, golden_luxury, purple_royal, rose_elegant
📌 footer_style: simple, professional, contact, logo, decorative, double_line
📌 signature_style: single, dual, triple, stamp, multi_level, formal_table

==========================================
📝 مثال 1: خطاب رسمي كلاسيكي
==========================================
```python
from core.formatter import generate_styled_letter

output_path = generate_styled_letter(
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
        ["أغسطس", "0000612", "30000", "دخان"]
    ],
    header_style='official',
    content_style='paragraphs',
    table_style='royal_blue',
    footer_style='professional',
    signature_style='single',
    signatures=[{{'name': 'المحامي عبدالله', 'title': 'المستشار القانوني'}}],
    output_dir="{WORKSPACE_DIR}"
)
print("SUCCESS: " + output_path)
```

==========================================
📝 مثال 2: سيرة ذاتية متوافقة مع ATS
==========================================
```python
from core.formatter import generate_ats_cv

output_path = generate_ats_cv(
    full_name="أحمد محمد عبدالله",
    phone="+966501234567",
    email="ahmed@example.com",
    location="الرياض، السعودية",
    linkedin="linkedin.com/in/ahmed",
    objective="العمل في منصب مدير مالي في شركة رائدة.",
    summary="محاسب قانوني معتمد (CPA) بخبرة 10 سنوات في إدارة الحسابات والتحليل المالي.",
    experience=[{{'role': 'مدير الحسابات', 'company': 'شركة ردينت', 'date': '2020 - الحاضر', 'responsibilities': ['إدارة فريق محاسبة', 'إعداد التقارير المالية']}}],
    education=[{{'degree': 'بكالوريوس محاسبة', 'university': 'جامعة الملك سعود', 'year': '2014'}}],
    skills=['إدارة الحسابات', 'التحليل المالي', 'IFRS', 'Excel المتقدم', 'SAP ERP'],
    languages=[{{'name': 'العربية', 'level': 'اللغة الأم'}}, {{'name': 'الإنجليزية', 'level': 'طلاقة تامة'}}],
    certifications=[{{'name': 'CPA', 'issuer': 'الهيئة السعودية', 'year': '2019'}}],
    projects=[{{'name': 'تطوير نظام المحاسبة', 'date': '2022', 'description': 'تقليل الأخطاء بنسبة 40%'}}],
    interests=['القراءة في الاقتصاد', 'التطوير المهني'],
    output_dir="{WORKSPACE_DIR}"
)
print("SUCCESS: " + output_path)
```
⚠️ تعليمات مهمة للسيرة الذاتية: اكتب محتوى شاملاً، أضف دائماً: languages, certifications, projects, interests. المهارات كنقاط منفصلة (10-15 مهارة).

==========================================
📝 مثال 3: تقرير مالي شامل (قائمة الدخل)
==========================================
```python
from core.formatter import generate_financial_report

sections_data = [
    {{
        'title': 'الإيرادات',
        'headers': ['البند', 'القيمة الحالية', 'القيمة السابقة', 'التغيير %'],
        'data': [
            ['إيرادات المبيعات', 500000, 450000, 11.11],
            ['إيرادات الخدمات', 150000, 120000, 25.00],
            {{'is_subtotal': True}}
        ]
    }},
    {{
        'title': 'صافي الدخل',
        'headers': ['البند', 'القيمة الحالية', 'القيمة السابقة', 'التغيير %'],
        'data': [
            ['صافي الربح قبل الضريبة', 375000, 303000, 23.76],
            {{'is_total': True}}
        ]
    }}
]

output_path = generate_financial_report(
    report_type="income_statement",
    company_name="شركة ردينت للعناية بالسيارات",
    period="2026",
    sections=sections_data,
    currency="ريال",
    include_totals=True,
    include_formulas=True,
    style="accounting",
    output_dir="{WORKSPACE_DIR}"
)
print("SUCCESS: " + output_path)
```

==========================================
📝 مثال 4: تقرير الأصول الثابتة والإهلاكات
==========================================
```python
from core.formatter import create_fixed_asset_report

sections_data = [
    {{
        'sheet_name': 'المركبات',
        'headers': ['رقم الأصل', 'الوصف', 'تاريخ الشراء', 'التكلفة', 'الإهلاك', 'صافي القيمة'],
        'data': [
            ['V-001', 'تويوتا كامري 2023', '2023-01-15', 120000, 24000, 96000],
            ['V-002', 'هيونداي سوناتا 2022', '2022-06-10', 110000, 33000, 77000]
        ]
    }},
    {{
        'sheet_name': 'الأجهزة والمعدات',
        'headers': ['رقم الأصل', 'الوصف', 'تاريخ الشراء', 'التكلفة', 'الإهلاك', 'صافي القيمة'],
        'data': [
            ['E-001', 'خادم HP ProLiant', '2023-03-01', 45000, 9000, 36000],
            ['E-002', 'طابعات ليزر (5)', '2023-05-20', 15000, 3000, 12000]
        ]
    }}
]

output_path = create_fixed_asset_report(
    company_name="شركة ردينت للعناية بالسيارات",
    year="2026",
    sections=sections_data,
    output_dir="{WORKSPACE_DIR}"
)
print("SUCCESS: " + output_path)
```

==========================================
📝 مثال 5: جدول الإهلاك الشهري المضغوط
==========================================
```python
from core.formatter import generate_depreciation_schedule

assets_data = [
    {{
        'asset_id': 'V-001',
        'description': 'تويوتا هايلكس 2024',
        'cost': 185000,
        'accumulated_dep_start': 37000,
        'depreciation_rate': 33.33,
        'salvage_value': 0
    }},
    {{
        'asset_id': 'E-001',
        'description': 'خادم HP ProLiant',
        'cost': 85000,
        'accumulated_dep_start': 17000,
        'depreciation_rate': 25.00,
        'salvage_value': 5000
    }}
]

output_path = generate_depreciation_schedule(
    company_name="شركة ردينت للعناية بالسيارات",
    year="2026",
    assets=assets_data,
    output_dir="{WORKSPACE_DIR}"
)
print("SUCCESS: " + output_path)
```
        
==========================================
⚠️ تحذيرات حاسمة ومهمة جداً:
==========================================
- ممنوع إضافة أرقام (1. 2. 3.) في بداية الفقرات (النظام ينظفها تلقائياً).
- ممنوع استخدام from docx import Document مباشرة.
- ممنوع اختراع دوال غير موجودة.
- اختر الأنماط المناسبة حسب نوع المستند.

⚠️ قاعدة صارمة لـ print:
❌ ممنوع تماماً: print(f"SUCCESS: {{output_path}}") أو استخدام f-string مع المتغيرات.
✅ مسموح فقط وبشكل حصري: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

اختر الدالة المناسبة والأنماط الأنسب. ابدأ بـ ```python وانهِ بـ ```"""

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
1. استخدم الدالة المناسبة من core.formatter
2. تأكد من تمرير جميع المعاملات المطلوبة
3. اختر الأنماط المناسبة (header_style, content_style, table_style...)
4. اطبع النتيجة بهذه الصيغة حصراً: print("SUCCESS: " + output_path)
5. ممنوع إضافة أرقام في بداية الفقرات

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
