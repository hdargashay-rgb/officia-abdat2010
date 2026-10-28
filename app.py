"""
Officia - واجهة الويب للوكيل الذكي
"""
import streamlit as st
import os
import sys

# إضافة مسار المشروع
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.agent import Agent
from core.config import WORKSPACE_DIR, MODELS

# إعداد الصفحة
st.set_page_config(
    page_title="Officia - الوكيل الذكي",
    page_icon="🤖",
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
    .download-box {
        background-color: #e8f5e9;
        padding: 1rem;
        border-radius: 0.5rem;
        border-right: 4px solid #4caf50;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# تهيئة الجلسة
if 'agent' not in st.session_state:
    st.session_state.agent = Agent()
    st.session_state.chat_history = []
    st.session_state.last_files_created = []  # لتخزين الملفات للتحميل

# الشريط الجانبي
with st.sidebar:
    st.title("⚙️ الإعدادات")
    st.markdown("---")
    
    st.subheader("📊 معلومات النظام")
    
    # عرض جميع المزودين والنماذج المتاحة
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
    
    st.info("💡 النظام ينتقل تلقائياً بين المزودين عند فشل أي منهم", icon="ℹ️")
    
    st.write(f"**مجلد العمل:**")
    st.code(WORKSPACE_DIR, language="text")
    
    st.markdown("---")
    
    st.subheader("📁 الملفات في مساحة العمل")
    if os.path.exists(WORKSPACE_DIR):
        files = os.listdir(WORKSPACE_DIR)
        if files:
            for file in files:
                file_path = os.path.join(WORKSPACE_DIR, file)
                size = os.path.getsize(file_path)
                st.write(f"📄 {file} ({size/1024:.1f} KB)")
        else:
            st.info("لا توجد ملفات")
    
    st.markdown("---")
    
    if st.button("🗑️ مسح المحادثة والذاكرة"):
        st.session_state.chat_history = []
        st.session_state.last_files_created = []
        st.rerun()

# الواجهة الرئيسية
st.title("🤖 Officia - الوكيل الذكي للمهام المكتبية")
st.markdown("مرحباً بك! أنا مساعدك الذكي لإنجاز المهام المكتبية. اكتب طلبك بالعربية وسأنفذه لك.")

# عرض المحادثة
chat_container = st.container()
with chat_container:
    for message in st.session_state.chat_history:
        if message['role'] == 'user':
            st.markdown(f"""
            <div class="chat-message user-message">
                <strong>👤 أنت:</strong><br>
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

# عرض أزرار تحميل الملفات (إذا وجدت)
if st.session_state.get('last_files_created'):
    st.markdown("---")
    st.markdown('<div class="download-box"><strong>📂 الملفات الجاهزة للتحميل:</strong></div>', unsafe_allow_html=True)
    
    for file in st.session_state.last_files_created:
        file_path = os.path.join(WORKSPACE_DIR, file)
        if os.path.exists(file_path):
            with open(file_path, "rb") as f:
                file_bytes = f.read()
            
            st.download_button(
                label=f"⬇️ تحميل الملف: {file}",
                data=file_bytes,
                file_name=file,
                mime="application/octet-stream",
                key=f"download_btn_{file}",
                use_container_width=True
            )

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
    # إضافة رسالة المستخدم
    st.session_state.chat_history.append({
        'role': 'user',
        'content': user_input
    })
    
    # عرض مؤشر التحميل
    with st.spinner('جاري تنفيذ المهمة... يرجى الانتظار'):
        try:
            result = st.session_state.agent.execute_task(user_input)
            
            if result['success']:
                response = f"✅ **تم التنفيذ بنجاح في المحاولة {result['attempts']}**\n\n"
                response += f"**المخرجات:**\n```\n{result['final_output']}\n```\n\n"
                
                if result.get('files_created'):
                    response += f"**الملفات المُنشأة:**\n"
                    for file in result['files_created']:
                        response += f"📄 {file}\n"
                    # حفظ الملفات في الجلسة لعرض أزرار التحميل
                    st.session_state.last_files_created = result['files_created']
                else:
                    st.session_state.last_files_created = []
            else:
                response = f"❌ **فشل التنفيذ بعد {result['attempts']} محاولات**\n\n"
                response += f"**الخطأ:**\n```\n{result.get('error', 'خطأ غير محدد')}\n```\n"
                st.session_state.last_files_created = []
            
            # إضافة رد الوكيل
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
            st.session_state.last_files_created = []
    
    # إعادة رسم الصفحة لتحديث الواجهة
    st.rerun()

# تذييل الصفحة
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888;'>
    <p>Officia v1.1 - الوكيل الذكي للمهام المكتبية</p>
</div>
""", unsafe_allow_html=True)
