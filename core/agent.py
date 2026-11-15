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
        """دالة التوجيه الذكي الهجين: كلمات مفتاحية أولاً، ثم تصنيف ذكي بالذكاء الاصطناعي"""
        request_lower = user_request.lower()
        
        routing_rules = {
            'خطاب': {'function': 'generate_styled_letter', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_styled_letter', 'type': 'word'},
            'مذكرة': {'function': 'generate_styled_letter', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_styled_letter', 'type': 'word'},
            'رسالة رسمية': {'function': 'generate_styled_letter', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_styled_letter', 'type': 'word'},
            'سيرة ذاتية': {'function': 'generate_ats_cv', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_ats_cv', 'type': 'word'},
            'cv': {'function': 'generate_ats_cv', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_ats_cv', 'type': 'word'},
            'تقرير مالي': {'function': 'generate_financial_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_financial_report', 'type': 'excel'},
            'قائمة الدخل': {'function': 'generate_financial_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_financial_report', 'type': 'excel'},
            'ميزان عمومية': {'function': 'generate_financial_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_financial_report', 'type': 'excel'},
            'أصول ثابتة': {'function': 'create_fixed_asset_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة create_fixed_asset_report', 'type': 'excel'},
            'تقرير الأصول': {'function': 'create_fixed_asset_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة create_fixed_asset_report', 'type': 'excel'},
            'جدول إهلاك': {'function': 'generate_depreciation_schedule', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_depreciation_schedule', 'type': 'excel'},
            'إهلاك شهري': {'function': 'generate_depreciation_schedule', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_depreciation_schedule', 'type': 'excel'},
            'ميزان مراجعة': {'function': 'generate_trial_balance', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_trial_balance', 'type': 'excel'},
            'كشف رواتب': {'function': 'generate_payroll_register', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_payroll_register', 'type': 'excel'},
            'مسير رواتب': {'function': 'generate_payroll_register', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_payroll_register', 'type': 'excel'},
            'رواتب الموظفين': {'function': 'generate_payroll_register', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_payroll_register', 'type': 'excel'},
            'payroll': {'function': 'generate_payroll_register', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة generate_payroll_register', 'type': 'excel'},
            'محاكاة ملف': {'function': 'replicate_document', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة replicate_document', 'type': 'file'},
            'إعادة بناء': {'function': 'replicate_document', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة replicate_document', 'type': 'file'},
            'تحويل pdf': {'function': 'convert_pdf_to_word', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة convert_pdf_to_word', 'type': 'file'},
            'pdf to word': {'function': 'convert_pdf_to_word', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة convert_pdf_to_word', 'type': 'file'},
            'pdf to excel': {'function': 'convert_pdf_to_excel', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة convert_pdf_to_excel', 'type': 'file'},
            'استخراج جداول': {'function': 'convert_pdf_to_excel', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة convert_pdf_to_excel', 'type': 'file'},
            'تحليل نموذج': {'function': 'format_analysis_report', 'instruction': '⚠️ توجيه تلقائي: استخدم دالة format_analysis_report', 'type': 'excel'}
        }
        
        for keyword, rule in routing_rules.items():
            if keyword in request_lower:
                return {'found': True, 'function': rule['function'], 'instruction': rule['instruction'], 'type': rule['type'], 'method': 'keyword'}
        
        return self._llm_classify(user_request)

    def _llm_classify(self, user_request: str) -> dict:
        """استخدام النموذج اللغوي لفهم نية المستخدم واختيار الدالة الأنسب"""
        available_functions = """
        1. generate_styled_letter: خطابات، مذكرات، ورسائل رسمية (Word).
        2. generate_ats_cv: سير ذاتية احترافية متوافقة مع أنظمة التوظيف (Word).
        3. generate_financial_report: تقارير مالية، قوائم دخل، وميزانيات عمومية (Excel).
        4. create_fixed_asset_report: تقارير الأصول الثابتة وجداول الإهلاك (Excel).
        5. generate_trial_balance: ميزان مراجعة محاسبي (أرصدة افتتاحية وحركة وختامية) (Excel).
        6. generate_payroll_register: كشوف ورواتب ومسيرات موظفين شاملة (Excel).
        7. convert_pdf_to_word: تحويل ملفات PDF إلى مستندات Word قابلة للتعديل.
        8. convert_pdf_to_excel: استخراج الجداول والبيانات من ملفات PDF إلى Excel.
        9. replicate_document: محاكاة أو إعادة بناء ملف مرفوع (Word/Excel).
        10. format_analysis_report: تحليل بيانات نموذج Excel مرفوع وتنسيقه.
        """
        
        prompt = f"""
        أنت مصنف ذكي لنظام Officia. مهمتك هي اختيار رقم الدالة الأنسب لطلب المستخدم من القائمة أدناه.
        
        طلب المستخدم: "{user_request}"
        
        الدوال المتاحة:
        {available_functions}
        
        تعليمات الإجابة:
        - اكتب رقم الدالة الأنسب فقط (من 1 إلى 10).
        - إذا كان الطلب لا ينطبق على أي دالة ويحتاج إلى برمجة مخصصة من الصفر، اكتب 0.
        - لا تكتب أي شرح أو نص إضافي، فقط الرقم.
        """
        
        try:
            response = self.llm.ask(prompt).strip()
            import re
            numbers = re.findall(r'\d+', response)
            if numbers:
                selected_num = int(numbers[0])
            else:
                selected_num = 0
                
            function_map = {
                1: 'generate_styled_letter',
                2: 'generate_ats_cv',
                3: 'generate_financial_report',
                4: 'create_fixed_asset_report',
                5: 'generate_trial_balance',
                6: 'generate_payroll_register',
                7: 'convert_pdf_to_word',
                8: 'convert_pdf_to_excel',
                9: 'replicate_document',
                10: 'format_analysis_report'
            }
            
            if selected_num in function_map:
                func_name = function_map[selected_num]
                return {
                    'found': True, 
                    'function': func_name, 
                    'instruction': f'🧠 توجيه ذكي: تم تحليل طلبك واختيار دالة {func_name} الأنسب.', 
                    'type': 'smart_classified',
                    'method': 'llm'
                }
        except Exception as e:
            print(f"⚠️ فشل التصنيف الذكي: {str(e)}")
            
        return {'found': False, 'function': None, 'instruction': '', 'type': 'custom', 'method': 'fallback'}

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
    target_function = routing.get('function', 'generate_financial_report')
    
    # ⛔ قائمة الدوال المسموحة فقط
    ALLOWED_FUNCTIONS = """
⛔ قائمة الدوال المسموحة حصراً (ممنوع استخدام أي دالة خارجها):
  • generate_financial_report (لجميع التقارير المالية: قائمة الدخل، الميزانية، التدفقات)
  • generate_styled_letter (للخطابات والمذكرات)
  • generate_ats_cv (للسير الذاتية)
  • create_fixed_asset_report (للأصول الثابتة)
  • generate_trial_balance (لميزان المراجعة)
  • generate_payroll_register (للرواتب)
  • generate_depreciation_schedule (للإهلاك)
  • convert_pdf_to_word / convert_pdf_to_excel (لتحويل PDF)
  • replicate_document (لمحاكاة الملفات)
  • format_analysis_report (لتحليل النماذج)
  • create_excel_sheet / create_multi_sheet_excel (لجداول Excel عامة)
"""
    
    # ⚠️ قاعدة إلزامية صارمة
    STRICT_RULE = f"""
⚠️⚠️⚠️ قاعدة إلزامية صارمة - مخالفتها تسبب فشل المهمة ⚠️⚠️⚠️
1. الدالة المستهدفة لطلبك هي: **{target_function}**
2. يجب استخدام هذه الدالة حصراً - ممنوع اختراع دوال جديدة
3. ممنوع استدعاء: generate_balance_sheet, generate_income_statement, generate_cash_flow
4. لجميع التقارير المالية استخدم: generate_financial_report(report_type='...')
5. قيم report_type المسموحة: 'income_statement', 'balance_sheet', 'cash_flow', 'equity_statement'
"""
    
    # 🎨 بروتوكول التنسيق الإلزامي
    STYLE_PROTOCOL = """
🎨 بروتوكول التنسيق الإلزامي (يجب تطبيقه في كل ملف Excel):
  • عمود النصوص (البيان): عرض 45
  • أعمدة الأرقام: عرض 20-22
  • اللون الأزرق الغامق 1F4E79: للعناوين والرؤوس وصف الإجمالي
  • اللون الأزرق الفاتح D6E4F0: للمجاميع الفرعية
  • اللون الأبيض: لصفوف البيانات العادية
  • تلوين متبادل للبيانات (أبيض / أزرق فاتح جداً E8F0F7)
  • حدود مزدوجة (double border) تحت صف الإجمالي النهائي
  • معادلات Excel حية (=SUM) لكل إجمالي
  • اتجاه RTL للورقة
  • تجميد الرؤوس (freeze_panes)
"""
    
    if uploaded_file_path:
        return f"""أنت وكيل ذكي متخصص في تحليل المستندات.

📎 ملف مرفوع: {uploaded_file_path}
🎯 الطلب: {user_request}

{STRICT_RULE}
{ALLOWED_FUNCTIONS}
{STYLE_PROTOCOL}

{routing_instruction}

⚠️ استخدم الدوال من core.formatter أو core.pdf_converter فقط.

📚 الدوال المتاحة للملفات المرفوعة:
- convert_pdf_to_word(pdf_path='{uploaded_file_path}', output_dir='{WORKSPACE_DIR}')
- convert_pdf_to_excel(pdf_path='{uploaded_file_path}', output_dir='{WORKSPACE_DIR}')
- replicate_document(source_file='{uploaded_file_path}', output_dir='{WORKSPACE_DIR}')
- format_analysis_report(analysis_data, report_title, company_name, period, output_format='excel', style='accounting', output_dir='{WORKSPACE_DIR}')

========================================
📝 مثال 1: تحويل PDF إلى Word
========================================
```python
from core.pdf_converter import convert_pdf_to_word

output_path = convert_pdf_to_word(
    pdf_path='{uploaded_file_path}',
    output_dir='{WORKSPACE_DIR}'
)
print("SUCCESS: " + output_path)
```

========================================
📝 مثال 2: استخراج جداول من PDF إلى Excel
========================================
```python
from core.pdf_converter import convert_pdf_to_excel

output_path = convert_pdf_to_excel(
    pdf_path='{uploaded_file_path}',
    output_dir='{WORKSPACE_DIR}'
)
print("SUCCESS: " + output_path)
```

اكتب الكود الكامل الآن.
"""
    
    else:
        return f"""أنت وكيل ذكي متخصص في إنشاء المستندات والتقارير.

🎯 الطلب: {user_request}

{STRICT_RULE}
{ALLOWED_FUNCTIONS}
{STYLE_PROTOCOL}

{routing_instruction}

📚 أمثلة على الاستخدام الصحيح:

========================================
📝 مثال 1: قائمة الدخل
========================================
```python
from core.formatter import generate_financial_report

result = generate_financial_report(
    report_type='income_statement',
    company_name='شركة ردينت',
    period='2026',
    sections=[
        {{'title': 'الإيرادات', 'headers': ['البيان', 'المبلغ'], 'data': [['مبيعات', 5000000], ['خدمات', 2000000]]}},
        {{'title': 'المصروفات', 'headers': ['البيان', 'المبلغ'], 'data': [['رواتب', 1500000], ['إيجارات', 300000]]}}
    ]
)
print("SUCCESS: " + result)
```

========================================
📝 مثال 2: الميزانية العمومية
========================================
```python
from core.formatter import generate_financial_report

result = generate_financial_report(
    report_type='balance_sheet',
    company_name='شركة ردينت',
    period='2026',
    sections=[
        {{'title': 'الأصول المتداولة', 'headers': ['البيان', 'المبلغ'], 'data': [['نقدية', 1500000], ['مخزون', 800000]]}},
        {{'title': 'الخصوم', 'headers': ['البيان', 'المبلغ'], 'data': [['دائنون', 900000], ['قروض', 2000000]]}}
    ]
)
print("SUCCESS: " + result)
```

========================================
📝 مثال 3: كشف رواتب
========================================
```python
from core.formatter import generate_payroll_register

result = generate_payroll_register(
    employees=[...],
    period='2026-01'
)
print("SUCCESS: " + result)
```

========================================
🚫 أمثلة خاطئة (ممنوعة تماماً):
========================================
```python
# ❌ generate_balance_sheet(...) - دالة غير موجودة
# ❌ generate_income_statement(...) - دالة غير موجودة
# ❌ from core.formatter import generate_balance_sheet - استيراد خاطئ
```

⚠️ تذكير أخير: استخدم **{target_function}** حصراً.

اكتب الكود الكامل الآن.
"""

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
