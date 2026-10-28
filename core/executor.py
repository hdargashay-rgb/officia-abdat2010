"""
Executor - ورشة العمل المعزولة
ينفذ أكواد Python في مجلد workspace مع مهلة زمنية وفحص أمني
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
        if timeout is None:
            timeout = self.timeout

        security_check = self._security_check(script_content)
        if not security_check['allowed']:
            return {'success': False, 'output': '', 'error': security_check['reason'], 'files_created': []}

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        script_path = os.path.join(self.workspace, f"temp_script_{timestamp}.py")
        
        try:
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(script_content)

            files_before = set(os.listdir(self.workspace))

            result = subprocess.run(
                [sys.executable, script_path],
                capture_output=True,
                text=True,
                timeout=timeout,
                cwd=self.workspace,
                env={**os.environ, 'PYTHONUTF8': '1'}
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
        forbidden_patterns = ['import subprocess', 'os.system(', 'eval(', 'exec(', 'import requests', 'import urllib', 'shutil.rmtree']
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
