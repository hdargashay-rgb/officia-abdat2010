"""
اعدادات مشروع Officia الاساسية
"""
import os

# 1. تحديد مسارات المشروع
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE_DIR = os.path.join(BASE_DIR, 'workspace')
LOGS_DIR = os.path.join(BASE_DIR, 'logs')
PROMPTS_DIR = os.path.join(BASE_DIR, 'prompts')

# 2. اعدادات الامان والحدود
MAX_AUTO_FIX_ATTEMPTS = 3
MAX_AGENT_STEPS = 10
EXECUTION_TIMEOUT = 30

# 3. اعدادات النماذج (Groq هو الرئيسي الموثوق 100%)
MODELS = {
    'groq': {
        'default': 'llama-3.3-70b-versatile'
    },
    'openrouter': {
        'default': 'meta-llama/llama-3.3-70b-instruct',
        'fallback': 'qwen/qwen-2.5-72b-instruct'
    },
    'gemini': {
        'default': 'gemini-3.8-flash'
    }
}

# 4. المكتبات المسموح للنموذج باستدعائها في السكربتات
ALLOWED_LIBRARIES = {
    'openpyxl', 'pandas', 'numpy', 'python-docx', 'Pillow',
    'os', 'sys', 'math', 'datetime', 'json', 're', 'pathlib', 'shutil', 'tempfile'
}

# 5. التاكد من وجود مجلدات العمل والسجلات
os.makedirs(WORKSPACE_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)
