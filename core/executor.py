"""
Executor - ورشة العمل المعزولة
ينفذ أكواد Python في مجلد workspace مع مهلة زمنية وفحص أمني

الإصدار المحسّن: حلول جذرية لا مؤقتة
1. مصدر واحد للحقيقة (Registry) - لا تكرار لقوائم الدوال
2. فحص شامل لكل الوحدات (formatter, table_engine, formula_helpers, styles, pdf_converter)
3. طبقة تحقق حقيقية من بروتوكول التنسيق (ليست مجرد تذكير في prompt)
4. فحص أمني متوازن (يسمح بـ requests/urllib الضرورية لـ pdf_converter)
"""
import os
import sys
import subprocess
import shutil
import re
from datetime import datetime
from core.config import WORKSPACE_DIR, EXECUTION_TIMEOUT, ALLOWED_LIBRARIES

# ============================================================
# 🎯 مصدر واحد للحقيقة: Registry
# إذا كان Registry متاحاً نستخدمه، وإلا نستخدم قائمة محلية كـ fallback
# ============================================================
try:
    from core.registry import REGISTRY, ENGINE_MODULES
    _HAS_REGISTRY = True
except ImportError:
    _HAS_REGISTRY = False
    REGISTRY = None
    ENGINE_MODULES = []


class Executor:
    """تنفيذ أكواد Python بأمان في بيئة معزولة"""
    
    def __init__(self):
        self.workspace = WORKSPACE_DIR
        self.timeout = EXECUTION_TIMEOUT
        os.makedirs(self.workspace, exist_ok=True)
        print("تم تهيئة Executor - مجلد العمل: " + self.workspace)
    
    # ============================================================
    # 🔐 بناء قائمة الدوال المسموحة ديناميكياً من Registry
    # الجذر: لا نكرر القوائم - نقرأها من مصدر واحد
    # ============================================================
    def _get_allowed_functions(self) -> set:
        """
        يرجع مجموعة كل الدوال المسموحة في النظام.
        الجذر: إذا أضفت دالة جديدة في Registry، تظهر هنا تلقائياً.
        """
        allowed = set()
        
        if _HAS_REGISTRY and REGISTRY:
            # كل الدوال المسجلة في Registry
            allowed.update(REGISTRY.keys())
        
        # دوال إضافية ضرورية (ليست في Registry لأنها دوال داخلية أو مساعدة)
        allowed.update({
            'generate_professional_letter',   # letter_planner يحتاجها
            'create_excel_sheet',             # مستخدمة في formatter
            'generate_receipt',               # إيصال
            'generate_memo',                  # مذكرة
        })
        
        # دوال وحدات الأدوات (table_engine, formula_helpers)
        for mod_name in (ENGINE_MODULES if _HAS_REGISTRY else ['core.table_engine', 'core.formula_helpers']):
            try:
                import importlib
                mod = importlib.import_module(mod_name)
                allowed.update({
                    name for name, obj in vars(mod).items()
                    if callable(obj) and not name.startswith('_')
                })
            except Exception:
                pass
        
        # دوال styles (مستخدمة في generate_styled_letter)
        try:
            from core import styles as styles_mod
            allowed.update({
                'apply_header_style', 'apply_content_style', 'apply_table_style',
                'apply_footer_style', 'apply_signature_style',
                'classify_document_type', 'get_recommended_styles'
            })
        except Exception:
            pass
        
        return allowed
    
    # ============================================================
    # 🧹 إزالة الترقيم التلقائي من نصوص الفقرات
    # ============================================================
    def remove_numbering_from_paragraphs(self, code: str) -> str:
        """إزالة الترقيم التلقائي (1. 2. 3.) من نصوص الفقرات"""
        result = code
        patterns = [
            r'(["\'])(\d+\.\s+)([^"\']*?)(["\'])',
            r'(["\'])(\d+-\s+)([^"\']*?)(["\'])',
            r'(["\'])(\d+\)\s+)([^"\']*?)(["\'])',
            r'(["\'])([\u0660-\u0669]+[\.\-\)]\s+)([^"\']*?)(["\'])',
        ]
        for pattern in patterns:
            result = re.sub(pattern, r'\1\3\4', result)
        return result
    
    # ============================================================
    # 🔧 تصحيح الأخطاء الشائعة
    # الجذر: نعالج الأنماط المتكررة التي يولدها النموذج
    # ============================================================
    def fix_common_code_errors(self, code: str) -> str:
        """تصحيح الأخطاء الشائعة في الكود المولد تلقائياً"""
        # 1. f-string المعقد في print → concatenation
        code = re.sub(
            r'print\(f["\']SUCCESS:\s*\{(\w+)\}["\']\)',
            r'print("SUCCESS: " + str(\1))',
            code
        )
        # 2. print مع متغير خاطئ (قاموس أو قائمة)
        code = re.sub(
            r'print\((assets|data|sections|headers|sections_data|assets_data)\)',
            r'pass  # تم حذف طباعة المتغير الخاطئ',
            code
        )
        # 3. الأقواس المعقوفة المزدوجة من f-string
        code = code.replace('{{', '{').replace('}}', '}')
        # 4. print مع dict مباشرة
        code = re.sub(
            r'print\((\w+)\["output_path"\]\)',
            r'print("SUCCESS: " + str(\1["output_path"]))',
            code
        )
        return code
    
    # ============================================================
    # 🛡️ التحقق من الاستيرادات - فحص شامل لكل الوحدات
    # الجذر: لا نفحص formatter فقط - نفحص كل الوحدات
    # ============================================================
    def validate_and_fix_imports(self, code: str) -> tuple:
        """
        التحقق من صحة الاستيرادات وإيقاف التنفيذ فوراً عند اكتشاف دوال مخترعة.
        
        الجذر: نفحص 5 وحدات:
        - core.formatter
        - core.table_engine
        - core.formula_helpers
        - core.styles
        - core.pdf_converter
        """
        allowed = self._get_allowed_functions()
        
        # قاموس البدائل: اسم مخترع → البديل الصحيح
        ALTERNATIVES = {
            'generate_balance_sheet': "generate_financial_report(report_type='balance_sheet')",
            'create_balance_sheet': "generate_financial_report(report_type='balance_sheet')",
            'generate_income_statement': "generate_financial_report(report_type='income_statement')",
            'create_income_statement': "generate_financial_report(report_type='income_statement')",
            'generate_cash_flow': "generate_financial_report(report_type='cash_flow')",
            'generate_cash_flow_statement': "generate_financial_report(report_type='cash_flow')",
            'create_cash_flow_statement': "generate_financial_report(report_type='cash_flow')",
            'generate_financial_statement': "generate_financial_report",
            'create_financial_report': "generate_financial_report",
            'generate_letter': "generate_styled_letter",
            'create_letter': "generate_styled_letter",
            'generate_official_letter': "generate_styled_letter",
            'generate_cv': "generate_ats_cv",
            'create_cv': "generate_ats_cv",
            'generate_resume': "generate_ats_cv",
            'generate_invoice': "create_professional_invoice",
            'create_invoice': "create_professional_invoice",
            'generate_depreciation': "generate_depreciation_schedule",
            'create_depreciation_schedule': "generate_depreciation_schedule",
            'create_trial_balance': "generate_trial_balance",
            'generate_payroll': "generate_payroll_register",
            'create_payroll_register': "generate_payroll_register",
            'create_payroll': "generate_payroll_register",
            'generate_fixed_asset_report': "create_fixed_asset_report",
            'generate_fixed_assets_report': "create_fixed_asset_report",
            'create_fixed_assets_report': "create_fixed_asset_report",
        }
        
        # الوحدات التي نفحص استيراداتها
        MODULES_TO_CHECK = [
            'core.formatter',
            'core.table_engine',
            'core.formula_helpers',
            'core.styles',
            'core.pdf_converter',
        ]
        
        errors = []
        critical_errors = []
        
        for module_name in MODULES_TO_CHECK:
            # نمط: from module_name import func1, func2
            pattern = rf'from\s+{re.escape(module_name)}\s+import\s+([^\n]+)'
            matches = re.findall(pattern, code)
            
            for match in matches:
                # تجاهل import *
                if match.strip() == '*':
                    errors.append(f"❌ ممنوع 'from {module_name} import *'")
                    critical_errors.append('*')
                    continue
                
                functions = [f.strip() for f in match.split(',')]
                for func in functions:
                    func = func.strip()
                    if not func:
                        continue
                    
                    if func not in allowed:
                        error_msg = f"❌ دالة غير موجودة في {module_name}: '{func}'"
                        errors.append(error_msg)
                        critical_errors.append(func)
                        
                        if func in ALTERNATIVES:
                            errors.append(f"   💡 البديل الصحيح: {ALTERNATIVES[func]}")
                        else:
                            # اقتراح أقرب دالة من القائمة المسموحة
                            close_matches = [
                                a for a in allowed 
                                if func[:5] in a or a[:5] in func
                            ][:3]
                            if close_matches:
                                errors.append(f"   📋 دوال مشابهة: {', '.join(close_matches)}")
                            else:
                                errors.append(f"   📋 الدوال المتاحة: {', '.join(sorted(allowed))}")
        
        # إذا كانت هناك دوال مخترعة، أرجع code=None لإيقاف التنفيذ فوراً
        if critical_errors:
            return None, errors
        
        return code, errors
    
    # ============================================================
    # 🎨 طبقة التحقق من بروتوكول التنسيق (فحص حقيقي)
    # الجذر: ليس مجرد تذكير في prompt - فحص برمجي فعلي
    # ============================================================
    def _verify_formatting_protocol(self, code: str) -> dict:
        """
        التحقق من تطبيق بروتوكول التنسيق الإلزامي في الكود المخصص.
        
        الجذر: نرفض الكود الذي لا يطبق البروتوكول، وليس فقط نطلب من النموذج تطبيقه.
        """
        # استثناء: الكود الذي يستخدم table_engine يطبق التنسيق تلقائياً
        if 'from core.table_engine import' in code or 'from table_engine import' in code:
            return {'valid': True, 'problems': []}
        
        # استثناء: الكود الذي يستخدم دوال formatter الجاهزة (تنسّق نفسها)
        if 'from core.formatter import' in code:
            return {'valid': True, 'problems': []}
        
        # استثناء: كود PDF converter (ليس تنسيق Excel/Word)
        if 'from core.pdf_converter import' in code:
            return {'valid': True, 'problems': []}
        
        problems = []
        
        # 1. التحقق من وجود openpyxl أو python-docx
        if 'openpyxl' not in code and 'docx' not in code:
            problems.append("الكود لا يستخدم openpyxl أو python-docx")
        
        # 2. التحقق من RTL للمستندات العربية
        if 'rightToLeft' not in code and 'WD_ALIGN_PARAGRAPH.RIGHT' not in code:
            problems.append("الكود لا يطبق اتجاه RTL")
        
        # 3. التحقق من الألوان الاحترافية
        if '1F4E79' not in code and 'royal_dark' not in code:
            problems.append("الكود لا يستخدم الألوان الاحترافية (1F4E79)")
        
        # 4. التحقق من معادلات SUM للإجماليات (للـ Excel فقط)
        if 'openpyxl' in code and '=SUM(' not in code:
            problems.append("الكود لا يحتوي على معادلات SUM للإجماليات")
        
        # 5. التحقق من عرض الأعمدة
        if 'column_dimensions' not in code and 'width' not in code:
            problems.append("الكود لا يضبط عرض الأعمدة")
        
        return {
            'valid': len(problems) == 0,
            'problems': problems
        }
    
    # ============================================================
    # 🔒 الفحص الأمني - متوازن وليس صارماً جداً
    # الجذر: نسمح بـ requests/urllib لأن pdf_converter يحتاجها
    # ============================================================
    def _security_check(self, script: str) -> dict:
        """
        فحص أمني - محظورات حقيقية فقط.
        
        الجذر: نسمح بـ:
        - requests, urllib (ضرورية لـ pdf_converter)
        - os.environ, os.getenv (ضرورية لبعض الإعدادات)
        
        نمنع فقط:
        - subprocess, os.system (تنفيذ أوامر نظام)
        - eval, exec (تنفيذ كود ديناميكي)
        - shutil.rmtree, os.remove (حذف ملفات)
        """
        forbidden_patterns = [
            'import subprocess',
            'os.system(',
            'eval(',
            'exec(',
            'shutil.rmtree',
            'os.remove(',
            'os.unlink(',
            'os.rmdir(',
        ]
        for pattern in forbidden_patterns:
            if pattern in script:
                return {'allowed': False, 'reason': f'نمط محظور: {pattern}'}
        return {'allowed': True, 'reason': ''}
    
    # ============================================================
    # 🚀 التنفيذ الرئيسي
    # ============================================================
    def run_script(self, script_content: str, timeout: int = None) -> dict:
        if timeout is None:
            timeout = self.timeout
        
        # 🔍 التحقق من الاستيرادات (طبقة 1)
        script_content, import_errors = self.validate_and_fix_imports(script_content)
        if script_content is None:
            print("\n" + "=" * 60)
            print("🛑 تم إيقاف التنفيذ - دوال غير موجودة:")
            for err in import_errors:
                print(f"  {err}")
            print("=" * 60 + "\n")
            return {
                'success': False,
                'output': '',
                'error': '\n'.join(import_errors),
                'files_created': []
            }
        
        # 🔧 تصحيح الأخطاء الشائعة (طبقة 2)
        script_content = self.fix_common_code_errors(script_content)
        script_content = self.remove_numbering_from_paragraphs(script_content)
        
        # 🎨 التحقق من بروتوكول التنسيق (طبقة 3)
        formatting_check = self._verify_formatting_protocol(script_content)
        if not formatting_check['valid']:
            print("\n" + "=" * 60)
            print("🎨 تم إيقاف التنفيذ - الكود لا يطبق بروتوكول التنسيق:")
            for prob in formatting_check['problems']:
                print(f"  ❌ {prob}")
            print("=" * 60 + "\n")
            return {
                'success': False,
                'output': '',
                'error': '❌ الكود لا يطبق بروتوكول التنسيق الإلزامي:\n- ' + '\n- '.join(formatting_check['problems']),
                'files_created': []
            }
        
        # 🔒 الفحص الأمني (طبقة 4)
        security_check = self._security_check(script_content)
        if not security_check['allowed']:
            return {'success': False, 'output': '', 'error': security_check['reason'], 'files_created': []}
        
        # 🚀 التنفيذ الفعلي
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        script_path = os.path.join(self.workspace, f"temp_script_{timestamp}.py")
        
        try:
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)
            
            files_before = set(os.listdir(self.workspace))
            safe_env = {
                'PATH': os.environ.get('PATH', ''),
                'PYTHONUTF8': '1',
                'PYTHONIOENCODING': 'utf-8',
                'HOME': os.environ.get('HOME', '/tmp'),
                'USER': os.environ.get('USER', 'officia'),
            }
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            safe_env['PYTHONPATH'] = project_root
            
            result = subprocess.run(
                [sys.executable, script_path],
                capture_output=True,
                text=True,
                timeout=timeout,
                cwd=self.workspace,
                env=safe_env
            )
            
            files_after = set(os.listdir(self.workspace))
            new_files = [f for f in list(files_after - files_before) if not f.startswith('temp_script_')]
            
            if result.returncode == 0:
                return {'success': True, 'output': result.stdout, 'error': '', 'files_created': new_files}
            else:
                return {'success': False, 'output': result.stdout, 'error': result.stderr, 'files_created': new_files}
        
        except subprocess.TimeoutExpired:
            return {'success': False, 'output': '', 'error': f'انتهت المهلة الزمنية ({timeout} ثانية)', 'files_created': []}
        except Exception as e:
            return {'success': False, 'output': '', 'error': str(e), 'files_created': []}
        finally:
            if os.path.exists(script_path):
                try: os.remove(script_path)
                except: pass
    
    def run_python(self, script_content: str, timeout: int = None) -> dict:
        """واجهة مبسطة ونظيفة لاستدعاء تنفيذ الأكواد"""
        return self.run_script(script_content, timeout)
    
    def create_backup(self, filepath: str) -> str:
        if not os.path.exists(filepath): return None
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = filepath + f".backup_{timestamp}"
        shutil.copy2(filepath, backup_path)
        return backup_path
