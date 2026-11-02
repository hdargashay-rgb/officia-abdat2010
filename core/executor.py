"""
Executor - ورشة العمل المعزولة
ينفذ أكواد Python في مجلد workspace مع مهلة زمنية وفحص أمني
متوافق تماماً مع Officia v2.1 و formatter.py
"""
import os
import sys
import subprocess
import shutil
from datetime import datetime
from core.config import WORKSPACE_DIR, EXECUTION_TIMEOUT, ALLOWED_LIBRARIES


class Executor:
    """تنفيذ أكواد Python بأمان في بيئة معزولة"""

    def __init__(self):
        self.workspace = WORKSPACE_DIR
        self.timeout = EXECUTION_TIMEOUT
        os.makedirs(self.workspace, exist_ok=True)
        print("تم تهيئة Executor - مجلد العمل: " + self.workspace)

    def run_script(self, script_content: str, timeout: int = None) -> dict:
        """تنفيذ كود Python في بيئة معزولة وآمنة"""
        if timeout is None:
            timeout = self.timeout

        # الفحص الأمني
        security_check = self._security_check(script_content)
        if not security_check['allowed']:
            return {'success': False, 'output': '', 'error': security_check['reason'], 'files_created': []}

        # إنشاء ملف مؤقت
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        script_path = os.path.join(self.workspace, f"temp_script_{timestamp}.py")
        
        try:
            # كتابة الكود إلى الملف
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)

            # تسجيل الملفات قبل التنفيذ
            files_before = set(os.listdir(self.workspace))

            # ==========================================
            # بيئة معزولة وآمنة (بدون API Keys)
            # ==========================================
            safe_env = {
                'PATH': os.environ.get('PATH', ''),
                'PYTHONUTF8': '1',
                'PYTHONIOENCODING': 'utf-8',
                'HOME': os.environ.get('HOME', '/tmp'),
                'USER': os.environ.get('USER', 'officia'),
            }
            
            # إضافة PYTHONPATH ليشمل مجلد المشروع (لدعم core.formatter)
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            safe_env['PYTHONPATH'] = project_root

            # تنفيذ الكود
            result = subprocess.run(
                [sys.executable, script_path],
                capture_output=True,
                text=True,
                timeout=timeout,
                cwd=self.workspace,
                env=safe_env
            )

            # تسجيل الملفات بعد التنفيذ
            files_after = set(os.listdir(self.workspace))
            new_files = [f for f in list(files_after - files_before) if not f.startswith('temp_script_')]

            if result.returncode == 0:
                return {
                    'success': True, 
                    'output': result.stdout, 
                    'error': '', 
                    'files_created': new_files
                }
            else:
                return {
                    'success': False, 
                    'output': result.stdout, 
                    'error': result.stderr, 
                    'files_created': new_files
                }

        except subprocess.TimeoutExpired:
            return {
                'success': False, 
                'output': '', 
                'error': f'انتهت المهلة الزمنية ({timeout} ثانية)', 
                'files_created': []
            }
        except Exception as e:
            return {
                'success': False, 
                'output': '', 
                'error': str(e), 
                'files_created': []
            }
        finally:
            # حذف الملف المؤقت
            if os.path.exists(script_path):
                try: 
                    os.remove(script_path)
                except: 
                    pass

    def run_python(self, script_content: str, timeout: int = None) -> dict:
        """واجهة مبسطة ونظيفة لاستدعاء تنفيذ الأكواد"""
        return self.run_script(script_content, timeout)

    def _security_check(self, script: str) -> dict:
        """فحص أمني للكود المُنشأ"""
        # الأنماط المحظورة (خطرة)
        forbidden_patterns = [
            'import subprocess',
            'os.system(',
            'eval(',
            'exec(',
            'import requests',
            'import urllib',
            'shutil.rmtree',
            'os.environ',  # منع قراءة متغيرات البيئة (API Keys)
            'os.getenv',    # منع قراءة متغيرات البيئة
            'open(',        # منع قراءة ملفات النظام (اختياري)
        ]
        
        for pattern in forbidden_patterns:
            if pattern in script:
                return {
                    'allowed': False, 
                    'reason': f'نمط محظور: {pattern} - هذا النمط غير مسموح لأسباب أمنية'
                }
        
        return {'allowed': True, 'reason': ''}

    def create_backup(self, filepath: str) -> str:
        """إنشاء نسخة احتياطية من ملف"""
        if not os.path.exists(filepath): 
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = filepath + f".backup_{timestamp}"
        shutil.copy2(filepath, backup_path)
        return backup_path
