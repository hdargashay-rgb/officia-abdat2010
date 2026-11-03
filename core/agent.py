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
    summary="محاسب قانوني معتمد بخبرة 10 سنوات في إدارة الحسابات والتحليل المالي، متخصص في إعداد التقارير المالية ومراجعة الحسابات.",
    experience=[
        {{
            'role': 'مدير الحسابات',
            'company': 'شركة ردينت',
            'date': '2020 - الحاضر',
            'responsibilities': [
                'إدارة فريق محاسبة مكون من 5 موظفين',
                'إعداد التقارير المالية الشهرية والسنوية',
                'مراجعة الحسابات وضمان دقة البيانات'
            ]
        }}
    ],
    education=[
        {{
            'degree': 'بكالوريوس محاسبة',
            'university': 'جامعة الملك سعود',
            'year': '2014'
        }}
    ],
    skills=["Excel المتقدم", "SAP", "التحليل المالي", "IFRS", "المراجعة الداخلية"],
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
