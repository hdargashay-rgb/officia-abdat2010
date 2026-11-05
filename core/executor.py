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
        
        # أنماط متعددة للترقيم
        patterns = [
            # "1. نص" أو '1. نص'
            r'(["\'])(\d+\.\s+)([^"\']*?)(["\'])',
            # "1- نص" أو '1- نص'
            r'(["\'])(\d+-\s+)([^"\']*?)(["\'])',
            # "1) نص" أو '1) نص'
            r'(["\'])(\d+\)\s+)([^"\']*?)(["\'])',
            # أرقام عربية: "١. نص"
            r'(["\'])([\u0660-\u0669]+[\.\-\)]\s+)([^"\']*?)(["\'])',
        ]
        
        for pattern in patterns:
            result = re.sub(pattern, r'\1\3\4', result)
        
        return result

    def fix_common_code_errors(self, code: str) -> str:
        """
        تصحيح الأخطاء الشائعة في الكود المولد تلقائياً
        
        هذه الطبقة الأولى من الحماية ضد أخطاء النموذج في كتابة الكود.
        """
        # 1. تصحيح f-string المعقد: print(f"SUCCESS: {var}") → print("SUCCESS: " + str(var))
        code = re.sub(
            r'print\(f["\']SUCCESS:\s*\{(\w+)\}["\']\)',
            r'print("SUCCESS: " + str(\1))',
            code
        )
        
        # 2. تصحيح print مع متغير خاطئ (قاموس أو قائمة بدلاً من output_path)
        code = re.sub(
            r'print\((assets|data|sections|headers|sections_data|assets_data)\)',
            r'pass  # تم حذف طباعة المتغير الخاطئ',
            code
        )
        
        # 3. تصحيح الأقواس المعقوفة المزدوجة: {{var}} → {var}
        code = code.replace('{{', '{').replace('}}', '}')
        
        # 4. تصحيح print مع dict مباشرة
        code = re.sub(
            r'print\((\w+)\["output_path"\]\)',
            r'print("SUCCESS: " + str(\1["output_path"]))',
            code
        )
        
        return code

    def validate_and_fix_imports(self, code: str) -> tuple:
        """
        التحقق من صحة الاستيرادات وإيقاف التنفيذ فوراً عند اكتشاف دوال مخترعة
        
        ترجع: (code, errors) - إذا كانت هناك دوال مخترعة، code يصبح None
        """
        # قائمة الدوال المسموح بها من core.formatter
        ALLOWED_FUNCTIONS = ALLOWED_FUNCTIONS = ALLOWED_FUNCTIONS = {
    'generate_styled_letter',
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
    'generate_receipt',
    'generate_memo'
}
    
        
        errors = []
        critical_errors = []
        
        # استخراج جميع استيرادات formatter
        import_pattern = r'from\s+core\.formatter\s+import\s+([^\n]+)'
        matches = re.findall(import_pattern, code)
        
        for match in matches:
            functions = [f.strip() for f in match.split(',')]
            
            for func in functions:
                func = func.strip()
                if not func:
                    continue
                    
                if func not in ALLOWED_FUNCTIONS:
                    # دالة مخترعة - خطأ حرج
                    error_msg = f"❌ دالة غير موجودة: '{func}'"
                    errors.append(error_msg)
                    critical_errors.append(func)
                    
                    # قاموس البدائل للمساعدة
                    alternatives = {
                        'generate_income_statement_report': "generate_financial_report(report_type='income_statement')",
                        'generate_balance_sheet_report': "generate_financial_report(report_type='balance_sheet')",
                        'generate_cash_flow_report': "generate_financial_report(report_type='cash_flow_direct')",
                        'generate_equity_report': "generate_financial_report(report_type='equity_statement')",
                        'generate_trial_balance_report': "generate_financial_report(report_type='trial_balance')",
                        'generate_sales_report': "generate_financial_report(report_type='sales_report')",
                        'generate_expenses_report': "generate_financial_report(report_type='expenses_report')",
                        'generate_account_statement': "generate_financial_report(report_type='account_statement')",
                        'generate_bank_reconciliation': "generate_financial_report(report_type='bank_reconciliation')",
                        'generate_statement': "generate_financial_report",
                        'generate_financial_statement': "generate_financial_report",
                        'generate_report': "generate_financial_report",
                        'create_income_statement': "generate_financial_report",
                        'create_financial_report': "generate_financial_report",
                        'generate_depreciation': "generate_depreciation_schedule",
                        'generate_depreciation_report': "generate_depreciation_schedule",
                        'create_depreciation_schedule': "generate_depreciation_schedule",
                    }
                    
                    if func in alternatives:
                        errors.append(f"   💡 البديل الصحيح: {alternatives[func]}")
                    else:
                        errors.append(f"   📋 الدوال المتاحة: {', '.join(sorted(ALLOWED_FUNCTIONS))}")
        
        # إذا كانت هناك دوال مخترعة، أرجع code=None لإيقاف التنفيذ فوراً
        if critical_errors:
            return None, errors
        
        return code, errors

    def run_script(self, script_content: str, timeout: int = None) -> dict:
        if timeout is None:
            timeout = self.timeout

        # 🔍 التحقق من الاستيرادات وإيقاف فوري عند اكتشاف دوال مخترعة
        script_content, import_errors = self.validate_and_fix_imports(script_content)
        
        if script_content is None:
            # هناك دوال مخترعة - أوقف التنفيذ فوراً
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

        # 🔧 تصحيح الأخطاء الشائعة في الكود المولد (طبقة حماية إضافية)
        script_content = self.fix_common_code_errors(script_content)

        # تطبيق إزالة الترقيم
        script_content = self.remove_numbering_from_paragraphs(script_content)

        security_check = self._security_check(script_content)
        if not security_check['allowed']:
            return {'success': False, 'output': '', 'error': security_check['reason'], 'files_created': []}

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

    def _security_check(self, script: str) -> dict:
        forbidden_patterns = ['import subprocess', 'os.system(', 'eval(', 'exec(', 'import requests', 'import urllib', 'shutil.rmtree', 'os.environ', 'os.getenv']
        for pattern in forbidden_patterns:
            if pattern in script:
                return {'allowed': False, 'reason': f'نمط محظور: {pattern}'}
        return {'allowed': True, 'reason': ''}

    def create_backup(self, filepath: str) -> str:
        if not os.path.exists(filepath): return None
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = filepath + f".backup_{timestamp}"
        shutil.copy2(filepath, backup_path)
        return backup_path
