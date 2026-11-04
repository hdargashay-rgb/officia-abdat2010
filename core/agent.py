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
        return f"""أنت كاتب محترف ومبرمج Python خبير متخصص في إنشاء مستندات احترافية.

⚠️ قاعدة إلزامية:
استخدم الدوال من core.formatter فقط حسب نوع المستند المطلوب.

📚 الدوال المتاحة:

1️⃣ للخطابات الرسمية بأنماط متنوعة:
- generate_styled_letter(organization_name, title, recipient, subject, paragraphs, table_headers, table_data, table_caption, header_style, content_style, table_style, footer_style, signature_style, signatures, output_dir)

2️⃣ للسيرة الذاتية المتوافقة مع ATS:
- generate_ats_cv(full_name, phone, email, location, linkedin, summary, experience, education, skills, output_dir)

3️⃣ لتقارير Excel:
- create_multi_sheet_excel(sheets_config, output_dir)
- create_fixed_asset_report(company, year, sections, output_dir)
- create_professional_invoice(company, client, number, date, items, output_dir)

4️⃣ للتقارير المالية المعقدة (Excel متعدد الأوراق):
- create_fixed_asset_report(company_name, year, sections, output_dir)
  (ملاحظة: sections هي قائمة من القواميس، كل قاموس يحتوي على 'sheet_name', 'headers', 'data')

==========================================
🎨 الأنماط المتاحة للخطابات:
==========================================

📌 header_style (الرأس):
- official: رسمي كلاسيكي (خطوط مزدوجة + ذهبية)
- modern: حديث بسيط
- professional: احترافي مع شعار
- minimal: بسيط جداً
- bold: عريض قوي
- gradient: تجاري متدرج
- celebration: احتفالي مزخرف
- islamic: إسلامي تراثي (مع بسملة)

📌 content_style (المحتوى):
- paragraphs: فقرات تقليدية
- numbered: نقاط مرقمة
- bullets: نقاط برمز •
- compact: مضغوط
- mixed: مزيج مرن
- quoted: مقتبس
- highlighted: بارز ومميز
- decorative: زخرفي

📌 table_style (الجدول):
- royal_blue: أزرق ملكي
- green_finance: أخضر مالي
- dark_elegant: أسود أنيق
- colorful: ملون متدرج
- minimal: بسيط بدون ألوان
- accounting: محاسبي Excel
- modern_shadows: حديث بظلال
- golden_luxury: ذهبي فاخر
- purple_royal: بنفسجي ملكي
- rose_elegant: وردي أنيق

📌 footer_style (التذييل):
- simple: بسيط
- professional: احترافي
- contact: مع معلومات الاتصال
- logo: شعار + نص
- decorative: زخرفي
- double_line: خط مزدوج

📌 signature_style (التوقيع):
- single: توقيع واحد
- dual: توقيعان
- triple: ثلاثة توقيعات
- stamp: ختم + توقيع
- multi_level: متعدد المستويات
- formal_table: جدول توقيعات رسمي

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
        ["أغسطس", "0000612", "30000", "دخان"],
        ["سبتمبر", "0000613", "30000", "دخان"],
        ["أكتوبر", "0000614", "30000", "دخان"]
    ],
    table_caption="تفاصيل الشيكات المستحقة",
    header_style='official',
    content_style='paragraphs',
    table_style='royal_blue',
    footer_style='professional',
    signature_style='single',
    signatures=[{{'name': 'المحامي عبدالله', 'title': 'المستشار القانوني'}}],
    output_dir="{WORKSPACE_DIR}"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
📝 مثال 2: مذكرة داخلية حديثة
==========================================
```python
from core.formatter import generate_styled_letter

output_path = generate_styled_letter(
    organization_name="شركة ردينت",
    title="مذكرة داخلية",
    recipient="جميع الموظفين",
    subject="تحديث مواعيد العمل",
    paragraphs=[
        "نود إعلامكم بتعديل مواعيد العمل خلال شهر رمضان المبارك.",
        "سيكون دوام الموظفين من الساعة 9 صباحاً حتى 3 ظهراً.",
        "نرجو من الجميع الالتزام بالمواعيد الجديدة."
    ],
    header_style='modern',
    content_style='bullets',
    table_style='minimal',
    footer_style='simple',
    signature_style='none',
    output_dir="{WORKSPACE_DIR}"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
==========================================
📝 مثال 3: سيرة ذاتية متوافقة مع ATS
==========================================
```python
from core.formatter import generate_ats_cv

output_path = generate_ats_cv(
    full_name="أحمد محمد عبدالله",
    phone="+966501234567",
    email="ahmed@example.com",
    location="الرياض، السعودية",
    linkedin="linkedin.com/in/ahmed",
    objective="العمل في منصب مدير مالي في شركة رائدة حيث يمكنني توظيف خبرتي في التحليل المالي وإدارة الميزانيات لتحقيق أهداف النمو.",
    summary="محاسب قانوني معتمد (CPA) بخبرة 10 سنوات في إدارة الحسابات والتحليل المالي. متخصص في إعداد التقارير المالية وفق معايير IFRS، ومراجعة الحسابات الداخلية، وإدارة فرق العمل المحاسبية. حققت زيادة في كفاءة العمليات المالية بنسبة 25% في آخر منصب.",
    experience=[
        {
            'role': 'مدير الحسابات',
            'company': 'شركة ردينت للعناية بالسيارات',
            'date': '2020 - الحاضر',
            'responsibilities': [
                'إدارة فريق محاسبة مكون من 5 موظفين والإشراف على عملياتهم اليومية',
                'إعداد التقارير المالية الشهرية والسنوية وعرضها على مجلس الإدارة',
                'مراجعة الحسابات وضمان دقة البيانات المالية وفق معايير IFRS',
                'تطوير نظام محاسبي جديد قلل الوقت المستغرق في الإغلاق الشهري بنسبة 30%',
                'التنسيق مع المدققين الخارجيين وضمان الامتثال للأنظمة الضريبية'
            ]
        },
        {
            'role': 'محاسب أول',
            'company': 'شركة الخليج التجارية',
            'date': '2016 - 2020',
            'responsibilities': [
                'إدارة الحسابات المدينة والدائنة ومعالجة أكثر من 500 فاتورة شهرياً',
                'إعداد موازنات الأقسام ومتابعة الانحرافات',
                'تحليل البيانات المالية وتقديم توصيات لتحسين الأداء'
            ]
        }
    ],
    education=[
        {
            'degree': 'بكالوريوس محاسبة',
            'university': 'جامعة الملك سعود',
            'year': '2014'
        },
        {
            'degree': 'دبلوم معايير المحاسبة الدولية IFRS',
            'university': 'الهيئة السعودية للمحاسبين القانونيين',
            'year': '2018'
        }
    ],
    skills=[
        'إدارة الحسابات والميزانيات',
        'التحليل المالي وإعداد التقارير',
        'معايير المحاسبة الدولية IFRS',
        'المراجعة الداخلية والامتثال',
        'Excel المتقدم (Pivot Tables, VBA)',
        'نظام SAP ERP',
        'QuickBooks و Oracle Financials',
        'إدارة فرق العمل والقيادة',
        'التخطيط المالي والتنبؤ',
        'اللغة الإنجليزية (تحدثاً وكتابة)'
    ],
    languages=[
        {'name': 'العربية', 'level': 'اللغة الأم'},
        {'name': 'الإنجليزية', 'level': 'طلاقة تامة (IELTS 7.5)'},
        {'name': 'الفرنسية', 'level': 'مستوى متوسط'}
    ],
    certifications=[
        {'name': 'محاسب قانوني معتمد CPA', 'issuer': 'الهيئة السعودية', 'year': '2019'},
        {'name': 'شهادة CMA - محاسب إداري معتمد', 'issuer': 'IMA', 'year': '2021'},
        {'name': 'شهادة IFRS Diploma', 'issuer': 'ACCA', 'year': '2018'}
    ],
    projects=[
        {
            'name': 'تطوير نظام المحاسبة السحابي',
            'date': '2022',
            'description': 'قيادة مشروع الانتقال من النظام التقليدي إلى SAP Cloud، مما أدى إلى تقليل الأخطاء بنسبة 40% وتحسين سرعة الإغلاق الشهري.'
        }
    ],
    interests=['القراءة في الاقتصاد العالمي', 'التطوير المهني المستمر', 'التطوع في تعليم المحاسبة للشباب'],
    output_dir="{WORKSPACE_DIR}"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
📝 مثال 4: تقرير مالي متعدد الأوراق (Fixed Asset)
==========================================
```python
from core.formatter import create_fixed_asset_report

sections_data = [
    {
        'sheet_name': 'المركبات',
        'headers': ['رقم الأصل', 'الوصف', 'تاريخ الشراء', 'التكلفة', 'الإهلاك', 'صافي القيمة'],
        'data': [
            ['V-001', 'تويوتا كامري 2023', '2023-01-15', 120000, 24000, 96000],
            ['V-002', 'هيونداي سوناتا 2022', '2022-06-10', 110000, 33000, 77000]
        ]
    },
    {
        'sheet_name': 'الأجهزة والمعدات',
        'headers': ['رقم الأصل', 'الوصف', 'تاريخ الشراء', 'التكلفة', 'الإهلاك', 'صافي القيمة'],
        'data': [
            ['E-001', 'خادم HP ProLiant', '2023-03-01', 45000, 9000, 36000],
            ['E-002', 'طابعات ليزر (5)', '2023-05-20', 15000, 3000, 12000]
        ]
    }
]

output_path = create_fixed_asset_report(
    company_name="شركة ردينت للعناية بالسيارات",
    year="2026",
    sections=sections_data,
    output_dir="{WORKSPACE_DIR}"
)
print(f"SUCCESS: {{output_path}}")
```

==========================================
⚠️ تحذيرات مهمة:
==========================================
- ممنوع إضافة أرقام (1. 2. 3.) في بداية الفقرات
- ممنوع استخدام from docx import Document مباشرة
- ممنوع اختراع دوال غير موجودة
- اختر الأنماط المناسبة حسب نوع المستند
- للسيرة الذاتية: استخدم generate_ats_cv فقط

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
4. اطبع: print(f"SUCCESS: {{output_path}}")
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

==========================================
📝 مثال 4: تقرير مالي شامل (قائمة الدخل)
==========================================
```python
from core.formatter import generate_financial_report

sections_data = [
    {
        'title': 'الإيرادات',
        'headers': ['البند', 'القيمة الحالية', 'القيمة السابقة', 'التغيير %'],
        'data': [
            ['إيرادات المبيعات', 500000, 450000, 11.11],
            ['إيرادات الخدمات', 150000, 120000, 25.00],
            ['إيرادات أخرى', 50000, 30000, 66.67],
            {'is_subtotal': True}  # علامة الإجمالي الفرعي
        ]
    },
    {
        'title': 'المصروفات التشغيلية',
        'headers': ['البند', 'القيمة الحالية', 'القيمة السابقة', 'التغيير %'],
        'data': [
            ['رواتب وأجور', 200000, 180000, 11.11],
            ['إيجارات', 60000, 60000, 0.00],
            ['مرافق وكهرباء', 25000, 22000, 13.64],
            ['مصروفات إدارية', 40000, 35000, 14.29],
            {'is_subtotal': True}
        ]
    },
    {
        'title': 'صافي الدخل',
        'headers': ['البند', 'القيمة الحالية', 'القيمة السابقة', 'التغيير %'],
        'data': [
            ['صافي الربح قبل الضريبة', 375000, 303000, 23.76],
            ['الضريبة (15%)', 56250, 45450, 23.76],
            {'is_total': True}  # علامة الإجمالي الكلي
        ]
    }
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
print(f"SUCCESS: {{output_path}}")


==========================================
⚠️ تعليمات مهمة للسيرة الذاتية:
==========================================
- اكتب محتوى شاملاً واحترافياً (النموذج متدرب على ملايين السير الذاتية)
- أضف دائماً: languages, certifications, projects, interests
- المهارات: اكتبها كنقاط منفصلة (10-15 مهارة)
- الخبرة: 2-3 وظائف على الأقل، كل واحدة بـ 3-5 مسؤوليات
- التعليم: درجة علمية + شهادات مهنية
- اللغات: 2-3 لغات مع المستوى
- الشهادات: شهادات مهنية مع الجهة والسنة
- المشاريع: مشروع واحد على الأقل مع وصف
- الهدف المهني: فقرة قصيرة (2-3 أسطر)
- الملخص المهني: فقرة قوية (4-5 أسطر) تبرز الإنجازات بالأرقام
