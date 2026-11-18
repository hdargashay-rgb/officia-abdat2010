"""
Executor - ورشة العمل المعزولة
ينفذ أكواد Python في مجلد workspace مع مهلة زمنية وفحص أمني
"""
import os
import sys
import subprocess
import shutil
import re
from datetime import datetime
from core.config import WORKSPACE_DIR, EXECUTION_TIMEOUT, ALLOWED_LIBRARIES

class Executor:
    """تنفيذ أكواد Python بأمان في بيئة معزولة"""
    
    def __init__(self):
        self.workspace = WORKSPACE_DIR
        self.timeout = EXECUTION_TIMEOUT
        os.makedirs(self.workspace, exist_ok=True)
        print("تم تهيئة Executor - مجلد العمل: " + self.workspace)
    
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
    
    def fix_common_code_errors(self, code: str) -> str:
        """تصحيح الأخطاء الشائعة في الكود المولد تلقائياً"""
        code = re.sub(
            r'print\(f["\']SUCCESS:\s*\{(\w+)\}["\']\)',
            r'print("SUCCESS: " + str(\1))',
            code
        )
        code = re.sub(
            r'print\((assets|data|sections|headers|sections_data|assets_data)\)',
            r'pass  # تم حذف طباعة المتغير الخاطئ',
            code
        )
        code = code.replace('{{', '{').replace('}}', '}')
        code = re.sub(
            r'print\((\w+)\["output_path"\]\)',
            r'print("SUCCESS: " + str(\1["output_path"]))',
            code
        )
        return code
    
    def validate_and_fix_imports(self, code: str) -> tuple:
        """التحقق من صحة الاستيرادات وإيقاف التنفيذ فوراً عند اكتشاف دوال مخترعة"""
        ALLOWED_FUNCTIONS = {
            'generate_styled_letter',
            'generate_professional_letter',  # ✅ مُضاف
            'generate_ats_cv',
            'generate_financial_report',
            'create_fixed_asset_report',
            'generate_depreciation_schedule',
            'generate_trial_balance',
            'generate_payroll_register',
            'replicate_document',
            'format_analysis_report',
            'create_multi_sheet_excel',
            'create_professional_invoice',
        }
        
        errors = []
        critical_errors = []
        import_pattern = r'from\s+core\.formatter\s+import\s+([^\n]+)'
        matches = re.findall(import_pattern, code)
        
        for match in matches:
            functions = [f.strip() for f in match.split(',')]
            for func in functions:
                func = func.strip()
                if not func:
                    continue
                if func not in ALLOWED_FUNCTIONS:
                    error_msg = f"❌ دالة غير موجودة: '{func}'"
                    errors.append(error_msg)
                    critical_errors.append(func)
                    alternatives = {
                        'generate_income_statement_report': "generate_financial_report(report_type='income_statement')",
                        'generate_balance_sheet_report': "generate_financial_report(report_type='balance_sheet')",
                        'generate_cash_flow_report': "generate_financial_report(report_type='cash_flow')",
                        'generate_letter': "generate_styled_letter",
                        'create_letter': "generate_styled_letter",
                        'generate_cv': "generate_ats_cv",
                        'create_cv': "generate_ats_cv",
                    }
                    if func in alternatives:
                        errors.append(f"   💡 البديل الصحيح: {alternatives[func]}")
                    else:
                        errors.append(f"   📋 الدوال المتاحة: {', '.join(sorted(ALLOWED_FUNCTIONS))}")
        
        if critical_errors:
            return None, errors
        return code, errors
    
    def _verify_formatting_protocol(self, code: str) -> dict:
        """
        ✅ طبقة جديدة: التحقق من تطبيق بروتوكول التنسيق الإلزامي
        """
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
    
    def _security_check(self, script: str) -> dict:
        """فحص أمني - محظورات حقيقية فقط"""
        # ✅ تم تخفيف القيود للسماح بـ requests/urllib (مهمة لـ pdf_converter)
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
    
    def run_script(self, script_content: str, timeout: int = None) -> dict:
        if timeout is None:
            timeout = self.timeout
        
        # 🔍 التحقق من الاستيرادات
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
        
        # 🔧 تصحيح الأخطاء الشائعة
        script_content = self.fix_common_code_errors(script_content)
        script_content = self.remove_numbering_from_paragraphs(script_content)
        
        # ✅ طبقة التحقق من بروتوكول التنسيق
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
        
        # 🔒 الفحص الأمني
        security_check = self._security_check(script_content)
        if not security_check['allowed']:
            return {'success': False, 'output': '', 'error': security_check['reason'], 'files_created': []}
        
        # 🚀 التنفيذ
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
