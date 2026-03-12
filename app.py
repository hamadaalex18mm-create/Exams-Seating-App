import streamlit as st
import pandas as pd
import io
import os
import math

# إعدادات الصفحة
st.set_page_config(page_title="توزيع أماكن الامتحانات", layout="wide")

# ==========================================
# دالة ذكية لتحويل أرقام المستويات لنصوص (بدون تكرار)
# ==========================================
def format_level(val):
    s = str(val).strip()
    s = s.replace("المستوي", "").replace("المستوى", "").strip()
    
    parts = s.split()
    if parts:
        if parts[0] == '1': parts[0] = 'الاول'
        elif parts[0] == '2': parts[0] = 'الثاني'
        elif parts[0] == '3': parts[0] = 'الثالث'
        elif parts[0] == '4': parts[0] = 'الرابع'
        
    return " ".join(parts)

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
    .stButton > button { font-size: 18px !important; font-weight: bold !important; width: 100% !important; }
    
    [data-testid="stDataFrame"] div[role="gridcell"], 
    [data-testid="stDataFrame"] div[role="columnheader"] {
        text-align: right !important; justify-content: flex-end !important; font-size: 15px !important;
    }
    
    .stTabs [data-baseweb="tab-list"] { gap: 10px; justify-content: flex-end; }
    .stTabs [data-baseweb="tab"] { font-size: 18px; font-weight: bold; background-color: #f0f4f8; border-radius: 5px 5px 0 0; padding: 10px 20px; }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; }
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
    st.markdown("<div style='display: flex; justify-content: center; align-items: center; height: 100%; margin-top: 20px;'><h1 style='margin: 0;'>توزيع أماكن الامتحانات (الشجرة)</h1></div>", unsafe_allow_html=True)
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
# الخطوة 2 و 3: إدخال البيانات والخوارزمية
# ==========================================
if st.session_state.rooms_df is not None and st.session_state.students_df is not None:
    
    # --- الخطوة 2 ---
    st.markdown("<h3>الخطوة 2: إدخال بيانات الخريطة</h3>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        exam_period = st.selectbox("أماكن امتحانات:", ["منتصف فصل الخريف", "نهاية فصل الخريف", "منتصف فصل الربيع", "نهاية فصل الربيع"])
    with c2:
        academic_year = st.selectbox("العام الجامعي:", ["2025 - 2026", "2026 - 2027", "2027 - 2028", "2028 - 2029", "2029 - 2030"])
    
    level_courses = st.text_input("مقررات المستوي:")
    
    # --- الخطوة 3 ---
    st.markdown("<h3>الخطوة 3: توليد خريطة اللجان الموحدة</h3>", unsafe_allow_html=True)
    
    df_students = st.session_state.students_df
    unique_seats = sorted(df_students['رقم الجلوس'].unique())
    total_unique_students = len(unique_seats)
    all_subjects = sorted(df_students['اسم المقرر'].unique())
    
    seat_courses = df_students.groupby('رقم الجلوس')['اسم المقرر'].apply(lambda x: list(set(x))).to_dict()
    seat_levels = df_students.groupby('رقم الجلوس')['المستوي'].apply(lambda x: list(set(x))).to_dict()
    
    st.info(f"إجمالي عدد الطلبة (بدون تكرار) المطلوب توزيعهم: **{total_unique_students}** طالب.")
    
    if st.button("🚀 بدء التوزيع وتوليد الوثائق الرسمية", type="primary"):
        with st.spinner("جاري التوزيع والتنسيق الاحترافي لملفات الإكسيل..."):
            result_data = []
            curr_student_idx = 0
            
            first_seat = unique_seats[0] if total_unique_students > 0 else 0
            current_range_start = math.floor(first_seat / 5.0) * 5 if first_seat > 0 else 0
            if current_range_start < first_seat and current_range_start % 10 == 0:
                current_range_start += 1
            
            rooms_list = st.session_state.rooms_df.to_dict('records')
            
            for room in rooms_list:
                room_num = room['رقم اللجنة']
                room_loc = room['مكان اللجنة']
                try: room_cap = int(room['سعة اللجنة'])
                except: room_cap = 0
                
                if curr_student_idx >= total_unique_students or room_cap <= 0:
                    empty_room = {
                        'رقم اللجنة': room_num, 'مكان اللجنة': room_loc, 'سعة اللجنة': room_cap,
                        'من': '-', 'إلى': '-', 'ملاحظات': 'لجنة فارغة'
                    }
                    for subj in all_subjects: empty_room[subj] = '-'
                    result_data.append(empty_room)
                    continue
                
                course_counts = {}
                max_c = 0
                for i in range(curr_student_idx, total_unique_students):
                    seat = unique_seats[i]
                    courses_for_seat = seat_courses.get(seat, [])
                    
                    can_add = True
                    for c in courses_for_seat:
                        if course_counts.get(c, 0) + 1 > room_cap:
                            can_add = False
                            break
                            
                    remaining_total_students = total_unique_students - i
                    if not can_add and remaining_total_students < 5:
                        can_add = True 
                    
                    if not can_add:
                        break 
                        
                    for c in courses_for_seat:
                        course_counts[c] = course_counts.get(c, 0) + 1
                    max_c += 1
                
                final_end = None
                best_c = max_c
                
                if curr_student_idx + max_c == total_unique_students:
                    last_actual = unique_seats[curr_student_idx + max_c - 1]
                    final_end = math.ceil(last_actual / 5.0) * 5
                    best_c = max_c
                else:
                    for rollback in range(min(10, max_c)): 
                        test_c = max_c - rollback
                        if test_c <= 0: continue
                        
                        last_included = unique_seats[curr_student_idx + test_c - 1]
                        next_actual = unique_seats[curr_student_idx + test_c]
                        
                        largest_multiple_of_5 = math.floor((next_actual - 1) / 5.0) * 5
                        
                        if largest_multiple_of_5 >= last_included:
                            final_end = largest_multiple_of_5
                            best_c = test_c
                            break
                    
                    if final_end is None:
                        best_c = max_c
                        next_actual = unique_seats[curr_student_idx + best_c]
                        final_end = next_actual - 1
                
                final_course_counts = {}
                room_levels = set()
                
                for i in range(curr_student_idx, curr_student_idx + best_c):
                    current_seat = unique_seats[i]
                    for c in seat_courses.get(current_seat, []):
                         final_course_counts[c] = final_course_counts.get(c, 0) + 1
                    
                    for lvl in seat_levels.get(current_seat, []):
                        room_levels.add(str(lvl))
                
                if room_levels:
                    sorted_raw_levels = sorted(list(room_levels))
                    formatted_levels = []
                    for l in sorted_raw_levels:
                        f_lvl = format_level(l)
                        if f_lvl not in formatted_levels:
                            formatted_levels.append(f_lvl)
                    
                    if len(formatted_levels) > 2:
                        simplified_levels = []
                        for lvl in formatted_levels:
                            words = lvl.split()
                            if len(words) >= 3:
                                simplified = f"{words[0]} {words[-1]}" 
                                if simplified not in simplified_levels:
                                    simplified_levels.append(simplified)
                            else:
                                if lvl not in simplified_levels:
                                    simplified_levels.append(lvl)
                        notes_text = "المستوي " + " & ".join(simplified_levels)
                    else:
                        notes_text = "المستوي " + " & ".join(formatted_levels)
                else:
                    notes_text = ""
                
                room_data = {
                    'رقم اللجنة': room_num,
                    'مكان اللجنة': room_loc,
                    'سعة اللجنة': room_cap,
                    'من': current_range_start,
                    'إلى': final_end,
                    'ملاحظات': notes_text
                }
                for subj in all_subjects:
                    room_data[subj] = final_course_counts.get(subj, 0)
                    
                result_data.append(room_data)
                
                current_range_start = final_end + 1
                curr_student_idx += best_c
            
            final_df = pd.DataFrame(result_data)
            
            summary_data = []
            for row in result_data:
                summary_data.append({
                    'رقم اللجنة': row['رقم اللجنة'],
                    'مكان اللجنة': row['مكان اللجنة'],
                    'بداية اللجنة (من)': row['من'],
                    'نهاية اللجنة (إلى)': row['إلى'],
                    'ملاحظات': row['ملاحظات']
                })
            summary_df = pd.DataFrame(summary_data)

            if curr_student_idx < total_unique_students:
                st.error(f"⚠️ تحذير: إجمالي سعة اللجان لا تكفي! متبقي {total_unique_students - curr_student_idx} طالب بدون أماكن.")
            else:
                st.success("✅ تم الانتهاء من التوزيع وإنشاء الشيتين بنجاح!")

            tab1, tab2 = st.tabs(["📄 خريطة اللجان (ملخص)", "📊 الخريطة التفصيلية (بالمواد)"])
            with tab1:
                display_summary_cols = summary_df.columns.tolist()[::-1]
                styled_summary = summary_df[display_summary_cols].style.set_properties(**{'text-align': 'right'}).set_table_styles([dict(selector='th', props=[('text-align', 'right')])])
                st.dataframe(styled_summary, hide_index=True, use_container_width=True)
            with tab2:
                display_final_cols = final_df.columns.tolist()[::-1]
                styled_final = final_df[display_final_cols].style.set_properties(**{'text-align': 'right'}).set_table_styles([dict(selector='th', props=[('text-align', 'right')])])
                st.dataframe(styled_final, hide_index=True, use_container_width=True)
            
            # --- توليد ملف الإكسيل الاحترافي باللوجوهات والفوتر ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                
                sheets_info = [
                    {'df': summary_df, 'name': 'خريطة اللجان', 'orientation': 'portrait'},
                    {'df': final_df, 'name': 'الخريطة التفصيلية', 'orientation': 'landscape'}
                ]
                
                from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
                from openpyxl.worksheet.table import Table, TableStyleInfo
                from openpyxl.utils import get_column_letter
                from openpyxl.drawing.image import Image as xlImage
                
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                center_align = Alignment(horizontal='center', vertical='center', readingOrder=2)
                right_align = Alignment(horizontal='right', vertical='center', readingOrder=2)
                
                empty_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                header_font_white = Font(color="FFFFFF", bold=True, size=12)
                
                # جعلنا الفونت الأساسي للبيانات Bold
                data_font = Font(bold=True, size=12) 
                # فونت الترويسة
                meta_font = Font(bold=True, size=14, color="1E3A8A") 
                
                # فحص وجود الشعارات
                fac_logo = "logo_faculty.png" if os.path.exists("logo_faculty.png") else "logo_faculty.jpg" if os.path.exists("logo_faculty.jpg") else None
                unit_logo = "logo_unit.png" if os.path.exists("logo_unit.png") else "logo_unit.jpg" if os.path.exists("logo_unit.jpg") else None
                
                for idx, info in enumerate(sheets_info):
                    current_df = info['df']
                    sheet_name = info['name']
                    orientation = info['orientation']
                    
                    total_columns = len(current_df.columns)
                    last_col_letter = get_column_letter(total_columns)
                    
                    # البيانات تبدأ من الصف السادس (هنسيب أول 4 صفوف للترويسة والشعارات)
                    current_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=5)
                    worksheet = writer.sheets[sheet_name]
                    worksheet.sheet_view.rightToLeft = True 
                    
                    # 1. إدراج الشعارات (لو موجودة)
                    try:
                        if fac_logo:
                            img1 = xlImage(fac_logo)
                            img1.width, img1.height = 90, 90
                            worksheet.add_image(img1, 'A1') # يمين (RTL)
                        if unit_logo:
                            img2 = xlImage(unit_logo)
                            img2.width, img2.height = 90, 90
                            worksheet.add_image(img2, f'{last_col_letter}1') # يسار
                    except Exception:
                        pass
                    
                    # 2. تصميم الترويسة الاحترافية بدون حدود (Borders) في المنتصف
                    if total_columns > 2:
                        merge_end = get_column_letter(total_columns - 1)
                    else:
                        merge_end = 'A'
                        
                    worksheet.merge_cells(f'B1:{merge_end}1')
                    worksheet.merge_cells(f'B2:{merge_end}2')
                    worksheet.merge_cells(f'B3:{merge_end}3')
                    
                    worksheet['B1'] = f"أماكن امتحانات: {exam_period}"
                    worksheet['B2'] = f"العام الجامعي: {academic_year}"
                    worksheet['B3'] = f"مقررات المستوي: {level_courses}"
                    
                    for r in range(1, 6):
                        worksheet.row_dimensions[r].height = 26.25 
                        if r <= 3:
                            cell = worksheet[f'B{r}']
                            cell.alignment = center_align
                            cell.font = meta_font
                    
                    last_row = worksheet.max_row
                    
                    # 3. إنشاء جدول الإكسيل للبيانات
                    table_ref = f"A6:{last_col_letter}{last_row}"
                    tab = Table(displayName=f"TableMap_{idx}", ref=table_ref)
                    style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                    tab.tableStyleInfo = style
                    worksheet.add_table(tab)
                    
                    # 4. تنسيق الخلايا 
                    for r_idx in range(6, last_row + 1):
                        worksheet.row_dimensions[r_idx].height = 26.25 
                        
                        is_empty = False
                        if r_idx > 6:
                            if sheet_name == 'خريطة اللجان':
                                is_empty = (worksheet.cell(row=r_idx, column=3).value == '-') 
                            else:
                                is_empty = (worksheet.cell(row=r_idx, column=4).value == '-') 
                                
                        for c_idx in range(1, total_columns + 1):
                            cell = worksheet.cell(row=r_idx, column=c_idx)
                            cell.border = thin_border
                            
                            if r_idx == 6: # صف عناوين الجدول
                                cell.font = header_font_white
                                cell.alignment = center_align
                            else:
                                cell.font = data_font # خط عريض لكل البيانات
                                
                                if c_idx == 2: # مكان اللجنة
                                    cell.alignment = right_align
                                else:
                                    cell.alignment = center_align
                                    
                                if is_empty:
                                    cell.fill = empty_fill
                                
                    # 5. عرض الأعمدة
                    if sheet_name == 'خريطة اللجان':
                        worksheet.column_dimensions['A'].width = 15 
                        worksheet.column_dimensions['B'].width = 45 
                        worksheet.column_dimensions['C'].width = 20 
                        worksheet.column_dimensions['D'].width = 20 
                        worksheet.column_dimensions['E'].width = 45 
                    else:
                        worksheet.column_dimensions['A'].width = 15 
                        worksheet.column_dimensions['B'].width = 35 
                        worksheet.column_dimensions['C'].width = 12 
                        worksheet.column_dimensions['D'].width = 15 
                        worksheet.column_dimensions['E'].width = 15 
                        worksheet.column_dimensions['F'].width = 40 
                        for i in range(7, total_columns + 1):
