"""
Officia - واجهة الويب للوكيل الذكي
"""
import streamlit as st
import os
import sys

# إضافة مسار المشروع
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.agent import Agent
from core.config import WORKSPACE_DIR

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
</style>
""", unsafe_allow_html=True)

# تهيئة الجلسة
if 'agent' not in st.session_state:
    st.session_state.agent = Agent()
    st.session_state.chat_history = []

# الشريط الجانبي
with st.sidebar:
    st.title("⚙️ الإعدادات")
    st.markdown("---")
    
    st.subheader("📊 معلومات النظام")
    st.write(f"**المزود:** Groq")
    st.write(f"**النموذج:** qwen/qwen3.8-27b")
    st.write(f"**مجلد العمل:**")
    st.code(WORKSPACE_DIR, language="text")
    
    st.markdown("---")
    
    st.subheader("📁 الملفات المُنشأة")
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
    
    if st.button("🗑️ مسح المحادثة"):
        st.session_state.chat_history = []
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

# إدخال الطلب
st.markdown("---")
user_input = st.text_input("اكتب طلبك هنا:", placeholder="مثال: أنشئ ملف Excel بأسماء 5 موظفين ورواتبهم", key="input")

col1, col2, col3 = st.columns([3, 1, 1])

with col1:
    pass

with col2:
    submit = st.button("🚀 تنفيذ", type="primary")

with col3:
    if st.button("️ إيقاف"):
        st.stop()

# معالجة الطلب
if submit and user_input:
    # إضافة رسالة المستخدم
    st.session_state.chat_history.append({
        'role': 'user',
        'content': user_input
    })
    
    # عرض مؤشر التحميل
    with st.spinner('جاري تنفيذ المهمة...'):
        try:
            result = st.session_state.agent.execute_task(user_input)
            
            if result['success']:
                response = f"✅ **تم التنفيذ بنجاح في المحاولة {result['attempts']}**\n\n"
                response += f"**المخرجات:**\n```\n{result['final_output']}\n```\n\n"
                
                if result.get('files_created'):
                    response += f"**الملفات المُنشأة:**\n"
                    for file in result['files_created']:
                        response += f"📄 {file}\n"
            else:
                response = f"❌ **فشل التنفيذ بعد {result['attempts']} محاولات**\n\n"
                response += f"**الخطأ:**\n```\n{result.get('error', 'خطأ غير محدد')}\n```\n"
            
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
    
    # إعادة رسم الصفحة
    st.rerun()

# تذييل الصفحة
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888;'>
    <p>Officia v1.0 - الوكيل الذكي للمهام المكتبية</p>
</div>
""", unsafe_allow_html=True)