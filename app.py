import streamlit as st
import pandas as pd
import os

# إعدادات الصفحة
st.set_page_config(page_title="توزيع أماكن الامتحانات", layout="wide")

# ==========================================
# ستايل الواجهة الأساسي
# ==========================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap');
    * { font-family: 'Tajawal', sans-serif !important; }
    .stApp, .stMarkdown, .stText, p, label, input, .stSelectbox { 
        text-align: right !important; direction: rtl !important; font-size: 16px !important;
    }
    h1 { text-align: center !important; color: #1E3A8A !important; font-weight: 700 !important; }
    h3 { 
        color: #004d40 !important; font-weight: 700 !important; padding-top: 15px; 
        border-bottom: 2px solid #b6e3f4; padding-bottom: 10px; margin-bottom: 20px; 
        text-align: right !important; direction: rtl !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- الهيدر (الشعارات والعنوان الجديد) ---
col_left, col_space, col_right = st.columns([1, 3, 1])
with col_left:
    if os.path.exists("logo_faculty.png"): st.image("logo_faculty.png", use_container_width=True)
with col_space:
    st.markdown("<div style='display: flex; justify-content: center; align-items: center; height: 100%; margin-top: 20px;'><h1 style='margin: 0;'>توزيع أماكن الامتحانات (الشجرة)</h1></div>", unsafe_allow_html=True)
with col_right:
    if os.path.exists("logo_unit.png"): st.image("logo_unit.png", use_container_width=True)

st.markdown("---")

# --- تهيئة الذاكرة المؤقتة ---
if 'rooms_df' not in st.session_state: st.session_state.rooms_df = None
if 'students_df' not in st.session_state: st.session_state.students_df = None

# --- قسم رفع الملفات ---
st.markdown("<h3>الخطوة 1: رفع ملفات البيانات الأساسية</h3>", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1: 
    st.markdown("**1. ملف أماكن اللجان (رقم اللجنة، مكان اللجنة، سعة اللجنة)**")
    rooms_file = st.file_uploader("ارفع ملف اللجان بصيغة Excel", type=['xlsx'], key="rooms_uploader")
    
    if rooms_file:
        try:
            df_rooms = pd.read_excel(rooms_file)
            df_rooms.columns = df_rooms.columns.str.strip()
            required_rooms = ["رقم اللجنة", "مكان اللجنة", "سعة اللجنة"]
            
            if all(col in df_rooms.columns for col in required_rooms):
                st.session_state.rooms_df = df_rooms[required_rooms].copy()
                st.success(f"✅ تم قراءة ملف اللجان بنجاح! (عدد اللجان: {len(df_rooms)})")
            else:
                st.error("❌ تأكد من وجود الأعمدة: رقم اللجنة، مكان اللجنة، سعة اللجنة")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة ملف اللجان: {e}")

with col2: 
    st.markdown("**2. ملف الطلبة المسجلين (يحتوي على شيتات المقررات)**")
    students_file = st.file_uploader("ارفع ملف الطلبة بصيغة Excel", type=['xlsx'], key="students_uploader")
    
    if students_file:
        try:
            all_sheets = pd.read_excel(students_file, sheet_name=None)
            all_students = []
            
            for sheet_name, df in all_sheets.items():
                df.columns = df.columns.str.strip()
                if "المستوى" in df.columns: 
                    df.rename(columns={"المستوى": "المستوي"}, inplace=True)
                
                required_students = ["رقم الجلوس", "اسم المقرر", "المستوي"]
                
                if all(col in df.columns for col in required_students):
                    all_students.append(df[required_students])
                else:
                    st.warning(f"⚠️ الشيت '{sheet_name}' تم تجاهله لعدم وجود الأعمدة المطلوبة.")
            
            if all_students:
                st.session_state.students_df = pd.concat(all_students, ignore_index=True)
                st.session_state.students_df['رقم الجلوس'] = pd.to_numeric(st.session_state.students_df['رقم الجلوس'], errors='coerce')
                st.session_state.students_df.dropna(subset=['رقم الجلوس'], inplace=True)
                st.session_state.students_df = st.session_state.students_df.sort_values(by="رقم الجلوس").reset_index(drop=True)
                
                st.success(f"✅ تم دمج بيانات الطلبة بنجاح! (إجمالي الطلبة: {len(st.session_state.students_df)})")
            else:
                st.error("❌ لم يتم العثور على الأعمدة المطلوبة في أي شيت.")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة ملف الطلبة: {e}")

# --- رسالة التأكيد للانتقال للخطوة التالية ---
if st.session_state.rooms_df is not None and st.session_state.students_df is not None:
    st.markdown("---")
    st.info("🚀 ممتاز! تم رفع الملفين بنجاح. نحن جاهزون الآن لعملية التوزيع الذكية.")
