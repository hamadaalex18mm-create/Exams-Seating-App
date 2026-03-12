import streamlit as st
import pandas as pd
import io
import os

# إعدادات الصفحة
st.set_page_config(page_title="توزيع أماكن الامتحانات", layout="wide")

# ==========================================
# ستايل الواجهة
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
    .stButton > button { font-size: 18px !important; font-weight: bold !important; width: 100% !important; }
    
    [data-testid="stDataFrame"] div[role="gridcell"], 
    [data-testid="stDataFrame"] div[role="columnheader"] {
        text-align: right !important; justify-content: flex-end !important; font-size: 15px !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- الهيدر ---
col_left, col_space, col_right = st.columns([1, 3, 1])
with col_left:
    if os.path.exists("logo_faculty.png"): st.image("logo_faculty.png", use_container_width=True)
    elif os.path.exists("logo_faculty.jpg"): st.image("logo_faculty.jpg", use_container_width=True)
with col_space:
    st.markdown("<div style='display: flex; justify-content: center; align-items: center; height: 100%; margin-top: 20px;'><h1 style='margin: 0;'>توزيع أماكن الامتحانات (الشجرة الموحدة)</h1></div>", unsafe_allow_html=True)
with col_right:
    if os.path.exists("logo_unit.png"): st.image("logo_unit.png", use_container_width=True)
    elif os.path.exists("logo_unit.jpg"): st.image("logo_unit.jpg", use_container_width=True)

st.markdown("---")

if 'rooms_df' not in st.session_state: st.session_state.rooms_df = None
if 'students_df' not in st.session_state: st.session_state.students_df = None

# ==========================================
# الخطوة 1: رفع الملفات
# ==========================================
if st.session_state.rooms_df is None or st.session_state.students_df is None:
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
                    if "المستوى" in df.columns: df.rename(columns={"المستوى": "المستوي"}, inplace=True)
                    required_students = ["رقم الجلوس", "اسم المقرر", "المستوي"]
                    
                    if all(col in df.columns for col in required_students):
                        all_students.append(df[required_students])
                    else:
                        st.warning(f"⚠️ الشيت '{sheet_name}' تم تجاهله لعدم وجود الأعمدة المطلوبة.")
                
                if all_students:
                    df_all = pd.concat(all_students, ignore_index=True)
                    df_all['رقم الجلوس'] = pd.to_numeric(df_all['رقم الجلوس'], errors='coerce')
                    df_all.dropna(subset=['رقم الجلوس'], inplace=True)
                    df_all['رقم الجلوس'] = df_all['رقم الجلوس'].astype(int)
                    st.session_state.students_df = df_all
                    st.success(f"✅ تم دمج بيانات الطلبة بنجاح! (إجمالي السجلات: {len(df_all)})")
                    st.rerun()
                else:
                    st.error("❌ لم يتم العثور على الأعمدة المطلوبة في أي شيت.")
            except Exception as e:
                st.error(f"حدث خطأ أثناء قراءة ملف الطلبة: {e}")

# ==========================================
# الخطوة 2: خوارزمية الشجرة الموحدة
# ==========================================
if st.session_state.rooms_df is not None and st.session_state.students_df is not None:
    st.markdown("<h3>الخطوة 2: توليد خريطة اللجان الموحدة</h3>", unsafe_allow_html=True)
    
    # 1. استخراج أرقام الجلوس الفريدة وترتيبها
    df_students = st.session_state.students_df
    unique_seats = sorted(df_students['رقم الجلوس'].unique())
    total_unique_students = len(unique_seats)
    
    # 2. إنشاء قاموس سريع لمعرفة كل طالب بيمتحن إيه (عشان الأداء يكون سريع جداً)
    seat_courses = df_students.groupby('رقم الجلوس')['اسم المقرر'].apply(list).to_dict()
    
    st.info(f"إجمالي عدد الطلبة (بدون تكرار) المطلوب توزيعهم: **{total_unique_students}** طالب.")
    
    if st.button("🚀 بدء التوزيع الذكي الموحد", type="primary"):
        with st.spinner("جاري التوزيع والتأكد من عدم تخطي السعة في أي مقرر..."):
            result_data = []
            curr_student_idx = 0
            rooms_list = st.session_state.rooms_df.to_dict('records')
            
            for room in rooms_list:
                room_num = room['رقم اللجنة']
                room_loc = room['مكان اللجنة']
                try: room_cap = int(room['سعة اللجنة'])
                except: room_cap = 0
                
                # لو الطلبة خلصوا، قفل باقي اللجان على إنها فارغة
                if curr_student_idx >= total_unique_students or room_cap <= 0:
                    result_data.append({
                        'رقم اللجنة': room_num,
                        'مكان اللجنة': room_loc,
                        'من': '-',
                        'إلى': '-',
                        'ملاحظات': 'لجنة فارغة'
                    })
                    continue
                
                # خوارزمية التعبئة بدون كسر سعة أي مقرر
                course_counts = {}
                max_idx = curr_student_idx
                
                for i in range(curr_student_idx, total_unique_students):
                    seat = unique_seats[i]
                    courses_for_seat = seat_courses.get(seat, [])
                    
                    # هل إضافة الطالب ده هتكسر السعة في أي مادة؟
                    can_add = True
                    for c in courses_for_seat:
                        if course_counts.get(c, 0) + 1 > room_cap:
                            can_add = False
                            break
                    
                    if not can_add:
                        break # نوقف قبل ما نضيف الطالب ده
                    
                    # لو السعة تمام، ضيف الطالب للمواد بتاعته وكمل
                    for c in courses_for_seat:
                        course_counts[c] = course_counts.get(c, 0) + 1
                    max_idx = i
                
                # تطبيق اللمسة الإبداعية: التقفيل على 0 أو 5 (الرجوع للخلف كحد أقصى 9 خطوات)
                best_end_idx = max_idx
                for p in range(10): # من 0 لـ 9
                    check_idx = max_idx - p
                    if check_idx < curr_student_idx: # عشان منرجعش للجنة اللي قبلها
                        break
                    seat = unique_seats[check_idx]
                    if seat % 5 == 0:
                        best_end_idx = check_idx
                        break
                
                # تسجيل بيانات اللجنة
                start_seat = unique_seats[curr_student_idx]
                end_seat = unique_seats[best_end_idx]
                global_count = best_end_idx - curr_student_idx + 1
                
                # حساب أقصى كثافة فعلية بعد التقفيل (عشان الملاحظات)
                final_course_counts = {}
                for i in range(curr_student_idx, best_end_idx + 1):
                    for c in seat_courses.get(unique_seats[i], []):
                         final_course_counts[c] = final_course_counts.get(c, 0) + 1
                max_course_load = max(final_course_counts.values()) if final_course_counts else 0
                
                result_data.append({
                    'رقم اللجنة': room_num,
                    'مكان اللجنة': room_loc,
                    'من': start_seat,
                    'إلى': end_seat,
                    'ملاحظات': f'نطاق: {global_count} طالب | أعلى كثافة لمادة: {max_course_load} (سعة: {room_cap})'
                })
                
                # تحريك المؤشر للطلبة اللي بعدهم
                curr_student_idx = best_end_idx + 1
            
            final_df = pd.DataFrame(result_data)
            
            if curr_student_idx < total_unique_students:
                st.error(f"⚠️ تحذير: إجمالي سعة اللجان لا تكفي! متبقي {total_unique_students - curr_student_idx} طالب بدون أماكن.")
            else:
                st.success("✅ تم الانتهاء من التوزيع الموحد بنجاح!")
            
            # عرض النتيجة
            styled_df = final_df.style.set_properties(**{'text-align': 'right'}).set_table_styles([dict(selector='th', props=[('text-align', 'right')])])
            st.dataframe(styled_df, hide_index=True, use_container_width=True)
            
            # --- ملف الإكسيل للطباعة ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='خريطة اللجان')
                workbook = writer.book
                worksheet = writer.sheets['خريطة اللجان']
                worksheet.sheet_view.rightToLeft = True 
                
                from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                center_align = Alignment(horizontal='center', vertical='center')
                header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True, size=12)
                empty_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                
                for col_num in range(1, 6):
                    cell = worksheet.cell(row=1, column=col_num)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center_align
                    cell.border = thin_border
                
                for r_idx in range(2, worksheet.max_row + 1):
                    is_empty = (worksheet.cell(row=r_idx, column=3).value == '-')
                    for c_idx in range(1, 6):
                        cell = worksheet.cell(row=r_idx, column=c_idx)
                        cell.border = thin_border
                        cell.alignment = center_align
                        if is_empty:
                            cell.fill = empty_fill
                            
                worksheet.column_dimensions['A'].width = 12 
                worksheet.column_dimensions['B'].width = 35 
                worksheet.column_dimensions['C'].width = 15 
                worksheet.column_dimensions['D'].width = 15 
                worksheet.column_dimensions['E'].width = 45 
                
                worksheet.print_area = f"A1:E{worksheet.max_row}"
                worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
                worksheet.sheet_properties.pageSetUpPr.fitToPage = True
                worksheet.page_setup.fitToWidth = 1
                worksheet.page_setup.fitToHeight = 0
                worksheet.print_options.horizontalCentered = True
            
            st.markdown("<div style='display: flex; justify-content: flex-end; width: 100%; margin-top: 15px;'>", unsafe_allow_html=True)
            st.download_button(
                label="📥 تحميل خريطة اللجان الموحدة (Excel)", 
                data=output.getvalue(), 
                file_name="توزيع_أماكن_الامتحانات_الموحد.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                type="primary"
            )
            st.markdown("</div>", unsafe_allow_html=True)
            
    st.markdown("---")
    if st.button("رجوع ورفع ملفات جديدة"):
        st.session_state.rooms_df = None
        st.session_state.students_df = None
        st.rerun()
