"""
Officia - واجهة الويب للوكيل الذكي
"""
import streamlit as st
import os
import sys
import time
from datetime import datetime

# إضافة مسار المشروع
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.agent import Agent
from core.config import WORKSPACE_DIR, MODELS

# ==========================================
# إعدادات إدارة الملفات
# ==========================================
MAX_FILES = 20  # الحد الأقصى للملفات
MAX_AGE_HOURS = 24  # عمر الملف الأقصى بالساعات

# إعداد الصفحة
st.set_page_config(
    page_title="Officia - الوكيل الذكي",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS مخصص للعربية
st.markdown("""
<style>
    .main {
        direction: rtl;
        text-align: right;
    }
    .stButton>button {
        width: 100%;
    }
    .chat-message {
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
    }
    .user-message {
        background-color: #e3f2fd;
        border-right: 4px solid #2196f3;
    }
    .bot-message {
        background-color: #f5f5f5;
        border-right: 4px solid #4caf50;
    }
    .provider-info {
        background-color: #f8f9fa;
        padding: 0.5rem;
        border-radius: 0.3rem;
        margin-bottom: 0.3rem;
        font-size: 0.85rem;
    }
    .file-card {
        background-color: #e8f5e9;
        padding: 1rem;
        border-radius: 0.5rem;
        border-right: 4px solid #4caf50;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3e0;
        padding: 0.8rem;
        border-radius: 0.5rem;
        border-right: 4px solid #ff9800;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# دوال إدارة الملفات الذكية
# ==========================================
def cleanup_old_files():
    """حذف الملفات القديمة جداً أو تجاوز الحد الأقصى"""
    if not os.path.exists(WORKSPACE_DIR):
        return []
    
    files = []
    for f in os.listdir(WORKSPACE_DIR):
        file_path = os.path.join(WORKSPACE_DIR, f)
        if os.path.isfile(file_path):
            mtime = os.path.getmtime(file_path)
            age_hours = (time.time() - mtime) / 3600
            files.append({
                'name': f,
                'path': file_path,
                'mtime': mtime,
                'age_hours': age_hours,
                'size': os.path.getsize(file_path)
            })
    
    # ترتيب من الأقدم إلى الأحدث
    files.sort(key=lambda x: x['mtime'])
    
    deleted = []
    
    # 1. حذف الملفات الأقدم من MAX_AGE_HOURS
    for file in files[:]:
        if file['age_hours'] > MAX_AGE_HOURS:
            try:
                os.remove(file['path'])
                deleted.append(file['name'])
                files.remove(file)
            except:
                pass
    
    # 2. إذا تجاوزنا الحد، نحذف الأقدم
    while len(files) > MAX_FILES:
        oldest = files.pop(0)
        try:
            os.remove(oldest['path'])
            deleted.append(oldest['name'])
        except:
            pass
    
    return deleted

def get_workspace_files():
    """جلب الملفات مع معلوماتها"""
    if not os.path.exists(WORKSPACE_DIR):
        return []
    
    files = []
    for f in os.listdir(WORKSPACE_DIR):
        file_path = os.path.join(WORKSPACE_DIR, f)
        if os.path.isfile(file_path):
            mtime = os.path.getmtime(file_path)
            files.append({
                'name': f,
                'path': file_path,
                'mtime': mtime,
                'size': os.path.getsize(file_path),
                'created': datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
            })
    
    # ترتيب من الأحدث إلى الأقدم
    files.sort(key=lambda x: x['mtime'], reverse=True)
    return files

def get_total_size():
    """حساب الحجم الكلي للملفات"""
    total = 0
    if os.path.exists(WORKSPACE_DIR):
        for f in os.listdir(WORKSPACE_DIR):
            file_path = os.path.join(WORKSPACE_DIR, f)
            if os.path.isfile(file_path):
                total += os.path.getsize(file_path)
    return total

# تهيئة الجلسة
if 'agent' not in st.session_state:
    st.session_state.agent = Agent()
    st.session_state.chat_history = []

# ==========================================
# الشريط الجانبي
# ==========================================
with st.sidebar:
    st.title("⚙️ الإعدادات")
    st.markdown("---")
    
    st.subheader("📊 معلومات النظام")
    
    st.markdown("**المزودون والنماذج المتاحة:**")
    
    st.markdown(f"""
    <div class="provider-info">
    🟢 <b>Groq:</b> {MODELS['groq']['default']}
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="provider-info">
    🟡 <b>OpenRouter:</b> {MODELS['openrouter']['default']}
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="provider-info">
    🔵 <b>Gemini:</b> {MODELS['gemini']['default']}
    </div>
    """, unsafe_allow_html=True)
    
    st.info("💡 النظام ينتقل تلقائياً بين المزودين عند فشل أي منهم")
    
    st.write(f"**مجلد العمل:**")
    st.code(WORKSPACE_DIR, language="text")
    
    st.markdown("---")
    
    # ==========================================
    # إدارة مساحة العمل الذكية
    # ==========================================
    st.subheader("📁 مساحة العمل")
    
    # تنظيف تلقائي عند تحميل الصفحة
    deleted_files = cleanup_old_files()
    if deleted_files:
        st.warning(f"🗑️ تم حذف {len(deleted_files)} ملف قديم تلقائياً")
    
    workspace_files = get_workspace_files()
    total_size = get_total_size()
    
    # عرض إحصائيات
    st.write(f"**عدد الملفات:** {len(workspace_files)} / {MAX_FILES}")
    st.write(f"**الحجم الكلي:** {total_size/1024:.1f} KB")
    
    # شريط تقدم
    usage_percent = (len(workspace_files) / MAX_FILES) * 100
    if usage_percent > 80:
        st.warning(f"⚠️ اقتربت من الحد الأقصى ({usage_percent:.0f}%)")
    else:
        st.progress(usage_percent / 100)
    
    # تنبيه مهم
    st.markdown("""
    <div class="warning-box">
    <small>⚠️ <b>تنبيه:</b> الملفات تُحذف تلقائياً بعد 24 ساعة أو عند إعادة تشغيل التطبيق. حمّل ملفاتك المهمة فوراً!</small>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # زر مسح جميع الملفات
    if st.button("️ مسح جميع الملفات", use_container_width=True):
        if os.path.exists(WORKSPACE_DIR):
            for f in os.listdir(WORKSPACE_DIR):
                file_path = os.path.join(WORKSPACE_DIR, f)
                if os.path.isfile(file_path):
                    os.remove(file_path)
        st.session_state.chat_history = []
        st.success("✅ تم مسح جميع الملفات")
        st.rerun()
    
    st.markdown("---")
    
    if st.button("🗑️ مسح المحادثة فقط"):
        st.session_state.chat_history = []
        st.rerun()

# ==========================================
# الواجهة الرئيسية
# ==========================================
st.title("🤖 Officia - الوكيل الذكي للمهام المكتبية")
st.markdown("مرحباً بك! أنا مساعدك الذكي لإنجاز المهام المكتبية. اكتب طلبك بالعربية وسأنفذه لك.")

# عرض المحادثة
chat_container = st.container()
with chat_container:
    for message in st.session_state.chat_history:
        if message['role'] == 'user':
            st.markdown(f"""
            <div class="chat-message user-message">
                <strong> أنت:</strong><br>
                {message['content']}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="chat-message bot-message">
                <strong>🤖 Officia:</strong><br>
                {message['content']}
            </div>
            """, unsafe_allow_html=True)

# ==========================================
# عرض أزرار التحميل في المنطقة الرئيسية
# ==========================================
workspace_files = get_workspace_files()

if workspace_files:
    st.markdown("---")
    st.markdown('<div class="file-card"><strong>📂 الملفات الجاهزة للتحميل:</strong></div>', unsafe_allow_html=True)
    
    # عرض في أعمدة (2 لكل صف)
    cols = st.columns(2)
    for idx, file_info in enumerate(workspace_files):
        file = file_info['name']
        file_path = file_info['path']
        
        with open(file_path, "rb") as f:
            file_bytes = f.read()
        
        mime_type = "application/octet-stream"
        if file.endswith('.docx'):
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif file.endswith('.xlsx'):
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif file.endswith('.pdf'):
            mime_type = "application/pdf"
        elif file.endswith('.txt'):
            mime_type = "text/plain"
        
        with cols[idx % 2]:
            st.markdown(f"**{file}**")
            st.caption(f"📅 {file_info['created']} | 💾 {file_info['size']/1024:.1f} KB")
            st.download_button(
                label=f"⬇️ تحميل",
                data=file_bytes,
                file_name=file,
                mime=mime_type,
                key=f"main_dl_{file}_{int(file_info['mtime'])}",
                use_container_width=True
            )
else:
    st.info("📭 لا توجد ملفات حالياً. اطلب من الوكيل إنشاء ملف!")

# إدخال الطلب
st.markdown("---")
user_input = st.text_input("اكتب طلبك هنا:", placeholder="مثال: أنشئ ملف Excel بأسماء 5 موظفين ورواتبهم", key="input")

col1, col2, col3 = st.columns([3, 1, 1])

with col1:
    pass

with col2:
    submit = st.button("🚀 تنفيذ", type="primary")

with col3:
    if st.button("⏹️ إيقاف"):
        st.stop()

# معالجة الطلب
if submit and user_input:
    st.session_state.chat_history.append({
        'role': 'user',
        'content': user_input
    })
    
    files_before = set(os.listdir(WORKSPACE_DIR)) if os.path.exists(WORKSPACE_DIR) else set()
    
    with st.spinner('جاري تنفيذ المهمة... يرجى الانتظار'):
        try:
            result = st.session_state.agent.execute_task(user_input)
            
            time.sleep(1)
            
            files_after = set(os.listdir(WORKSPACE_DIR)) if os.path.exists(WORKSPACE_DIR) else set()
            new_files = files_after - files_before
            
            if result['success']:
                response = f"✅ **تم التنفيذ بنجاح في المحاولة {result['attempts']}**\n\n"
                response += f"**المخرجات:**\n```\n{result['final_output']}\n```\n\n"
                
                if new_files:
                    response += f"**الملفات المُنشأة الجديدة:**\n"
                    for file in new_files:
                        response += f"📄 {file}\n"
                    response += "\n💡 **يمكنك تحميل الملفات من الأزرار أعلاه.**"
                else:
                    response += "**لم يتم إنشاء ملفات جديدة.**\n"
            else:
                response = f"❌ **فشل التنفيذ بعد {result['attempts']} محاولات**\n\n"
                response += f"**الخطأ:**\n```\n{result.get('error', 'خطأ غير محدد')}\n```\n"
            
            st.session_state.chat_history.append({
                'role': 'bot',
                'content': response
            })
            
        except Exception as e:
            error_msg = f"❌ **حدث خطأ غير متوقع:**\n```\n{str(e)}\n```"
            st.session_state.chat_history.append({
                'role': 'bot',
                'content': error_msg
            })
    
    st.rerun()

# تذييل الصفحة
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888;'>
    <p>Officia v2.1 - الوكيل الذكي للمهام المكتبية | إدارة ذكية للمساحة</p>
</div>
""", unsafe_allow_html=True)
