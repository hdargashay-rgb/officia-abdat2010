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

    def _route_request(self, user_request: str) -> dict:
        """
        دالة التوجيه الذكي: تحلل طلب المستخدم وتحدد الإجراء المناسب
        """
        request_lower = user_request.lower()
        
        routing_rules = {
            'خطاب': {'function': 'generate_styled_letter', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_styled_letter لإنشاء الخطاب', 'type': 'word'},
            'مذكرة': {'function': 'generate_styled_letter', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_styled_letter لإنشاء المذكرة', 'type': 'word'},
            'رسالة رسمية': {'function': 'generate_styled_letter', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_styled_letter لإنشاء الرسالة', 'type': 'word'},
            'سيرة ذاتية': {'function': 'generate_ats_cv', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_ats_cv لإنشاء السيرة الذاتية', 'type': 'word'},
            'cv': {'function': 'generate_ats_cv', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_ats_cv لإنشاء السيرة الذاتية', 'type': 'word'},
            'تقرير مالي': {'function': 'generate_financial_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_financial_report لإنشاء التقرير المالي (Excel)', 'type': 'excel'},
            'قائمة الدخل': {'function': 'generate_financial_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_financial_report لإنشاء قائمة الدخل (Excel)', 'type': 'excel'},
            'ميزان عمومية': {'function': 'generate_financial_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_financial_report لإنشاء الميزان العمومي (Excel)', 'type': 'excel'},
            'أصول ثابتة': {'function': 'create_fixed_asset_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة create_fixed_asset_report لإنشاء تقرير الأصول الثابتة (Excel متعدد الأوراق)', 'type': 'excel'},
            'تقرير الأصول': {'function': 'create_fixed_asset_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة create_fixed_asset_report لإنشاء تقرير الأصول (Excel متعدد الأوراق)', 'type': 'excel'},
            'جدول إهلاك': {'function': 'generate_depreciation_schedule', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_depreciation_schedule لإنشاء جدول الإهلاك (Excel)', 'type': 'excel'},
            'إهلاك شهري': {'function': 'generate_depreciation_schedule', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_depreciation_schedule لإنشاء جدول الإهلاك الشهري (Excel)', 'type': 'excel'},
            'ميزان مراجعة': {'function': 'generate_trial_balance', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_trial_balance لإنشاء ميزان المراجعة (Excel)', 'type': 'excel'},
            'كشف رواتب': {'function': 'generate_payroll_register', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_payroll_register لإنشاء كشف الرواتب (Excel متعدد الأوراق مع قسائم توقيع)', 'type': 'excel'},
            'مسير رواتب': {'function': 'generate_payroll_register', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_payroll_register لإنشاء مسير الرواتب (Excel متعدد الأوراق)', 'type': 'excel'},
            'رواتب الموظفين': {'function': 'generate_payroll_register', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_payroll_register لإنشاء رواتب الموظفين (Excel متعدد الأوراق)', 'type': 'excel'},
            'محاكاة ملف': {'function': 'replicate_document', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة replicate_document لمحاكاة الملف المرفوع', 'type': 'file'},
            'إعادة بناء': {'function': 'replicate_document', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة replicate_document لإعادة بناء الملف', 'type': 'file'},
            'تحليل نموذج': {'function': 'format_analysis_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة format_analysis_report لتحليل النموذج', 'type': 'excel'}
        }
        
        for keyword, rule in routing_rules.items():
            if keyword in request_lower:
                return {'found': True, 'function': rule['function'], 'instruction': rule['instruction'], 'type': rule['type']}
        
        return {'found': False, 'function': None, 'instruction': '', 'type': 'custom'}

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
        routing = self._route_request(user_request)
        routing_instruction = routing['instruction']
        
        # ==========================================
        # السيناريو الأول: وجود ملف مرفوع
        # ==========================================
        if uploaded_file_path:
            return f"""أنت وكيل ذكي متخصص في تحليل المستندات وإنشاء ملفات مشابهة.

📎 تم رفع ملف مرجعي: {uploaded_file_path}
🎯 طلب المستخدم: {user_request}

{routing_instruction}

⚠️ قاعدة إلزامية: استخدم الدوال المخصصة للملفات المرفوعة من core.formatter.

📚 الدوال المتاحة:
1️⃣ لمحاكاة الملف:
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

⚠️ تحذير: ممنوع اختراع دوال. ابدأ بـ ```python وانهِ بـ ```"""

        # ==========================================
        # السيناريو الثاني: Fallback احترافي (لا توجد دالة مخصصة)
        # ==========================================
        if not routing['found']:
            return f"""أنت مبرمج Python خبير متخصص في إنشاء مستندات احترافية بمعايير محاسبية وإدارية عالية.

🎯 طلب المستخدم: {user_request}

⚠️ لا توجد دالة مخصصة جاهزة لهذا الطلب في مكتبة core.formatter.
✅ الحل: اكتب كود Python كامل من الصفر لتنفيذ هذا الطلب بمعايير احترافية صارمة.

==========================================
🎨 بروتوكول التنسيق الاحترافي (إلزامي):
==========================================
📊 للملفات Excel (openpyxl):
1. استخدم أنماط احترافية: header_fill (لون 1F4E79), header_font (Arial, 12, bold, white), data_font (Arial, 10).
2. أضف عنواناً رئيسياً مدمجاً في الصف الأول، ورؤوس أعمدة ملونة في الصف الثاني.
3. استخدم تنسيق الأرقام: '#,##0.00' للأعمدة المالية.
4. أضف صف "إجمالي" في النهاية بخط عريض.
5. فعّل: ws.sheet_view.rightToLeft = True و ws.freeze_panes = 'A3' و ws.auto_filter.
6. اضبط عرض الأعمدة بشكل مناسب.

📄 للملفات Word (python-docx):
1. أضف عنواناً رئيسياً بخط كبير وعريض ومحاذاة وسط.
2. استخدم جداول منسقة (Table) للبيانات مع حدود واضحة.
3. أضف تذييل (Footer) يحتوي على التوقيع والتاريخ.

📁 الحفظ والطباعة:
- احفظ الملف في: {WORKSPACE_DIR}
- اطبع مسار الملف بهذه الصيغة حصراً: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

اكتب كود Python كامل بمعايير احترافية. ابدأ بـ ```python وانهِ بـ ```"""

        # ==========================================
        # السيناريو الثالث: دالة مخصصة موجودة - نعرضها فقط!
        # ==========================================
        
        # قاموس يحتوي على وصف ومثال كل دالة
        function_prompts = {
            'generate_styled_letter': f"""أنت كاتب محترف متخصص في إنشاء الخطابات الرسمية.

{routing_instruction}

⚠️ قاعدة إلزامية: استخدم دالة generate_styled_letter فقط.

📚 الدالة المتاحة:
- generate_styled_letter(organization_name, title, recipient, subject, paragraphs, table_headers, table_data, table_caption, header_style, content_style, table_style, footer_style, signature_style, signatures, output_dir)

🎨 الأنماط المتاحة:
📌 header_style: official, modern, professional, minimal, bold, gradient, celebration, islamic
📌 content_style: paragraphs, numbered, bullets, compact, mixed, quoted, highlighted, decorative
📌 table_style: royal_blue, green_finance, dark_elegant, colorful, minimal, accounting, modern_shadows, golden_luxury, purple_royal, rose_elegant
📌 footer_style: simple, professional, contact, logo, decorative, double_line
📌 signature_style: single, dual, triple, stamp, multi_level, formal_table

==========================================
📝 مثال: خطاب رسمي
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

⚠️ تحذيرات:
- ممنوع إضافة أرقام في بداية الفقرات.
- ممنوع اختراع دوال غير موجودة.
- اطبع النتيجة بهذه الصيغة حصراً: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

{routing_instruction}

ابدأ بـ ```python وانهِ بـ ```""",

            'generate_ats_cv': f"""أنت متخصص في إنشاء السير الذاتية المتوافقة مع ATS.

{routing_instruction}

⚠️ قاعدة إلزامية: استخدم دالة generate_ats_cv فقط.

📚 الدالة المتاحة:
- generate_ats_cv(full_name, phone, email, location, linkedin, summary, experience, education, skills, languages, certifications, projects, interests, objective, output_dir)

==========================================
📝 مثال: سيرة ذاتية
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

⚠️ تحذيرات:
- اكتب محتوى شاملاً.
- أضف دائماً: languages, certifications, projects, interests.
- اطبع النتيجة بهذه الصيغة حصراً: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

{routing_instruction}

ابدأ بـ ```python وانهِ بـ ```""",

            'generate_financial_report': f"""أنت محاسب خبير متخصص في إنشاء التقارير المالية.

{routing_instruction}

⚠️ قاعدة إلزامية: استخدم دالة generate_financial_report فقط.

📚 الدالة المتاحة:
- generate_financial_report(report_type, company_name, period, sections, currency='ريال', include_totals=True, include_formulas=True, style='accounting', output_dir)
  (report_type: 'income_statement', 'balance_sheet', 'cash_flow_direct', 'trial_balance', 'sales_report')

==========================================
📝 مثال: قائمة الدخل
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

⚠️ تحذيرات:
- اطبع النتيجة بهذه الصيغة حصراً: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

{routing_instruction}

ابدأ بـ ```python وانهِ بـ ```""",

            'create_fixed_asset_report': f"""أنت محاسب خبير متخصص في إنشاء تقارير الأصول الثابتة.

{routing_instruction}

⚠️ قاعدة إلزامية: استخدم دالة create_fixed_asset_report فقط.

📚 الدالة المتاحة:
- create_fixed_asset_report(company_name, year, sections, assets=None, include_depreciation_schedule=True, output_dir)
  (sections: قائمة الأقسام. assets: قائمة تفاصيل الأصول لحساب الإهلاك تلقائياً)

==========================================
📝 مثال: تقرير الأصول الثابتة
==========================================
```python
from core.formatter import create_fixed_asset_report

sections_data = [
    {{
        'sheet_name': 'المركبات',
        'headers': ['رقم الأصل', 'الوصف', 'التكلفة', 'الإهلاك', 'الصافي'],
        'data': [['V-001', 'تويوتا هايلكس 2024', 185000, 37000, 148000]]
    }},
    {{
        'sheet_name': 'الأجهزة والمعدات',
        'headers': ['رقم الأصل', 'الوصف', 'التكلفة', 'الإهلاك', 'الصافي'],
        'data': [['E-001', 'خادم HP ProLiant', 85000, 17000, 68000]]
    }}
]

assets_data = [
    {{'asset_id': 'V-001', 'description': 'تويوتا هايلكس 2024', 'category': 'المركبات', 'cost': 185000, 'accumulated_dep_start': 37000}},
    {{'asset_id': 'E-001', 'description': 'خادم HP ProLiant', 'category': 'الأجهزة والمعدات', 'cost': 85000, 'accumulated_dep_start': 17000}}
]

output_path = create_fixed_asset_report(
    company_name="شركة ردينت للعناية بالسيارات",
    year="2026",
    sections=sections_data,
    assets=assets_data,
    include_depreciation_schedule=True,
    output_dir="{WORKSPACE_DIR}"
)
print("SUCCESS: " + output_path)
```

⚠️ تحذيرات:
- اطبع النتيجة بهذه الصيغة حصراً: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

{routing_instruction}

ابدأ بـ ```python وانهِ بـ ```""",

            'generate_depreciation_schedule': f"""أنت محاسب خبير متخصص في إنشاء جداول الإهلاك.

{routing_instruction}

⚠️ قاعدة إلزامية: استخدم دالة generate_depreciation_schedule فقط.

📚 الدالة المتاحة:
- generate_depreciation_schedule(company_name, year, assets, output_dir)
  (assets: قائمة قواميس، كل قاموس: 'asset_id', 'description', 'cost', 'accumulated_dep_start', 'depreciation_rate', 'salvage_value')

==========================================
📝 مثال: جدول الإهلاك
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

⚠️ تحذيرات:
- اطبع النتيجة بهذه الصيغة حصراً: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

{routing_instruction}

ابدأ بـ ```python وانهِ بـ ```""",

            'generate_trial_balance': f"""أنت محاسب خبير متخصص في إنشاء ميزان المراجعة.

{routing_instruction}

⚠️ قاعدة إلزامية: استخدم دالة generate_trial_balance فقط.
⚠️ مهم جداً: accounts يجب أن تكون LIST من DICTS، وليست string!

📚 الدالة المتاحة:
- generate_trial_balance(company_name, period, accounts, output_dir)
  (accounts: قائمة قواميس، كل قاموس: 'account_number', 'account_name', 'opening_debit', 'opening_credit', 'debit_movement', 'credit_movement')

==========================================
📝 مثال: ميزان المراجعة
==========================================
```python
from core.formatter import generate_trial_balance

accounts_data = [
    {{
        'account_number': '1101',
        'account_name': 'الصندوق',
        'opening_debit': 50000,
        'opening_credit': 0,
        'debit_movement': 120000,
        'credit_movement': 95000
    }},
    {{
        'account_number': '2101',
        'account_name': 'الموردون',
        'opening_debit': 0,
        'opening_credit': 140000,
        'debit_movement': 100000,
        'credit_movement': 130000
    }}
]

output_path = generate_trial_balance(
    company_name="شركة ردينت للعناية بالسيارات",
    period="2026",
    accounts=accounts_data,
    output_dir="{WORKSPACE_DIR}"
)
print("SUCCESS: " + output_path)
```

⚠️ تحذيرات:
- accounts يجب أن تكون LIST من DICTS!
- اطبع النتيجة بهذه الصيغة حصراً: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

{routing_instruction}

ابدأ بـ ```python وانهِ بـ ```""",

            'generate_payroll_register': f"""أنت محاسب خبير متخصص في إنشاء كشوف الرواتب.

{routing_instruction}

⚠️ قاعدة إلزامية: استخدم دالة generate_payroll_register فقط.
⚠️ مهم جداً: هذه الدالة تنشئ ملف Excel متعدد الأوراق، وليس خطاب Word!
⚠️ employees يجب أن تكون LIST من DICTS!

📚 الدالة المتاحة:
- generate_payroll_register(company_name, period, employees, output_dir, include_signature_slips=True)
  (employees: قائمة قواميس، كل قاموس: 'employee_id', 'employee_name', 'department', 'basic_salary', 'housing_allowance', 'transport_allowance', 'other_allowances', 'gross_pay', 'insurance_deductions', 'other_deductions', 'net_pay', 'payment_date')

==========================================
📝 مثال: كشف الرواتب
==========================================
```python
from core.formatter import generate_payroll_register

employees_data = [
    {{
        'employee_id': 'EMP001',
        'employee_name': 'أحمد محمد',
        'department': 'إداري',
        'basic_salary': 8000,
        'housing_allowance': 2000,
        'transport_allowance': 500,
        'other_allowances': 0,
        'gross_pay': 10500,
        'insurance_deductions': 525,
        'other_deductions': 0,
        'net_pay': 9975,
        'payment_date': '2026-01-31'
    }},
    {{
        'employee_id': 'EMP002',
        'employee_name': 'خالد علي',
        'department': 'تشغيلي',
        'basic_salary': 5000,
        'housing_allowance': 1250,
        'transport_allowance': 300,
        'other_allowances': 200,
        'gross_pay': 6750,
        'insurance_deductions': 337.5,
        'other_deductions': 50,
        'net_pay': 6362.5,
        'payment_date': '2026-01-31'
    }}
]

output_path = generate_payroll_register(
    company_name="شركة ردينت للعناية بالسيارات",
    period="يناير 2026",
    employees=employees_data,
    include_signature_slips=True,
    output_dir="{WORKSPACE_DIR}"
)
print("SUCCESS: " + output_path)
```

⚠️ تحذيرات:
- employees يجب أن تكون LIST من DICTS!
- هذه الدالة تنشئ ملف Excel، وليس Word!
- اطبع النتيجة بهذه الصيغة حصراً: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

{routing_instruction}

ابدأ بـ ```python وانهِ بـ ```"""
        }
        
        # إرجاع الـ prompt الخاص بالدالة المحددة فقط
        function_name = routing['function']
        if function_name in function_prompts:
            return function_prompts[function_name]
        
        # Fallback إذا لم توجد الدالة في القاموس
        return f"""أنت مبرمج Python خبير.

🎯 طلب المستخدم: {user_request}

{routing_instruction}

اكتب كود Python كامل. ابدأ بـ ```python وانهِ بـ ```"""

==========================================
⚠️ تحذيرات حاسمة ومهمة جداً:
==========================================
- ممنوع إضافة أرقام (1. 2. 3.) في بداية الفقرات.
- ممنوع استخدام from docx import Document مباشرة.
- ممنوع اختراع دوال غير موجودة.
- لكشف الرواتب، استخدم generate_payroll_register (Excel) وليس generate_styled_letter (Word).

⚠️ قاعدة صارمة لـ print:
❌ ممنوع تماماً: print(f"SUCCESS: {{output_path}}")
✅ مسموح فقط وبشكل حصري: print("SUCCESS: " + output_path)

==========================================
الآن نفذ الطلب:
==========================================
{user_request}

{routing_instruction}

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
3. اختر الأنماط المناسبة
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
