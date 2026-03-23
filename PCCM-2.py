import streamlit as st
import pandas as pd
import re
from thefuzz import process
from datetime import datetime
from collections import defaultdict
import io
import base64

# ==========================================
# 0. CHUẨN HÓA TÊN CỘT ĐẦU VÀO
# ==========================================
COLUMN_ALIASES = {
    'TT': ['tt', 'stt', 'số thứ tự', 'sô thứ tự', 'so thu tu'],
    'Họ và tên': ['họ và tên', 'họ tên', 'giáo viên', 'tên giáo viên', 'họ tên gv', 'nhân sự', 'ho va ten', 'ho ten'],
    'Ngày sinh': ['ngày sinh', 'năm sinh', 'ns', 'ngaysinh', 'ngay sinh'],
    'PCCM': ['pccm', 'phân công chuyên môn', 'phân công', 'giảng dạy lớp', 'dạy lớp', 'chuyên môn', 'phan cong']
}

def standardize_columns(df):
    """Tự động nhận diện và đổi tên các cột biến thể về tên chuẩn"""
    rename_dict = {}
    for col in df.columns:
        # Làm sạch tên cột của file excel: viết thường, xóa khoảng trắng đầu cuối
        clean_col = str(col).lower().strip()
        
        # Tìm xem cột này thuộc nhóm chuẩn nào
        for std_name, aliases in COLUMN_ALIASES.items():
            if clean_col in aliases:
                rename_dict[col] = std_name
                break
                
    # Tiến hành đổi tên cột
    return df.rename(columns=rename_dict)


# ==========================================
# 1. TỪ ĐIỂN MAP MÃ MÔN
# ==========================================
SUBJECT_MAPPING = {
    "NGUVAN": ["ngữ văn", "văn"],
    "TOAN": ["toán", "toán học"],
    "ANH": ["tiếng anh", "ngoại ngữ 1", "ngoại ngữ 2", "ngoại ngữ", "anh"],
    "LICHSU": ["lịch sử", "sử"],
    "GDTC": ["giáo dục thể chất", "thể dục", "gdtc"],
    "GDQP": ["giáo dục quốc phòng và an ninh", "giáo dục quốc phòng", "qpan", "gdqp"],
    "DIALY": ["địa lí", "địa lý", "địa"],
    "GDKTPL": ["giáo dục kinh tế và pháp luật", "gdktpl", "kinh tế pháp luật", "ktpl"],
    "VATLY": ["vật lí", "vật lý", "lí", "lý"],
    "HOAHOC": ["hóa học", "hoá học", "hóa"],
    "SINH": ["sinh học", "sinh"],
    "CONGNGHE(NN)": ["cnnn", "nông nghiệp", "công nghệ (nn)", "công nghệ(nn)"],
    "CONGNGHE(CN)": ["cncn", "công nghiệp", "công nghệ (cn)", "công nghệ(cn)"],
    "CONGNGHE": ["công nghệ", "cncn", "cnnn"],
    "TINHOC": ["tin học", "tin"],
    "NDGDDP": ["nội dung giáo dục địa phương", "giáo dục địa phương", "gdđp", "gddp"],
    "TNHN": ["hoạt động trải nghiệm, hướng nghiệp", "hoạt động trải nghiệm", "hđ trải nghiệm", "hđtn", "hđtn hn"],
    "TIENGPHAP": ["tiếng pháp"], 
    "TIENGNGA": ["tiếng nga"], 
    "TIENGNHAT": ["tiếng nhật"],
    "TIENGTRUNG": ["tiếng trung"], 
    "TIENGHAN": ["tiếng hàn"], 
    "NGHEPHOTHONG": ["nghề phổ thông", "nghề"],
    "AMNHAC": ["âm nhạc", "nhạc"], 
    "MYTHUAT": ["mỹ thuật", "mĩ thuật", "mt"],
    "LICHSUDIALI": ["lịch sử và địa lí", "lịch sử và địa lý", "ls&đl"],
    "KHTN": ["khoa học tự nhiên", "khtn"],
    "GDCD": ["giáo dục công dân", "gdcd"],
    "HDNGLL": ["hoạt động ngoài giờ lên lớp", "hđngll"],
    "TDTTS": ["tiếng dân tộc thiểu số"], 
    "NGHETHUAT": ["nghệ thuật"]
}

REVERSE_MAP = {variant: code for code, variants in SUBJECT_MAPPING.items() for variant in variants}

def get_subject_code(raw_subject):
    raw_subject = str(raw_subject).lower().strip()
    if raw_subject in REVERSE_MAP:
        return REVERSE_MAP[raw_subject]
    
    best_match, score = process.extractOne(raw_subject, list(REVERSE_MAP.keys()))
    if score >= 75:
        return REVERSE_MAP[best_match]
    
    return raw_subject.upper()

# ==========================================
# 2. HÀM XỬ LÝ CHUỖI LỚP HỌC
# ==========================================
def parse_classes(class_string):
    if pd.isna(class_string) or not str(class_string).strip():
        return []
    
    s = str(class_string).strip()
    s = re.sub(r'\(\s*\d+\s*\)', '', s)
    s = s.replace('(', '').replace(')', '')
    
    result = []
    
    # Bắt các dải lớp
    range_pattern = re.compile(r'(\d+[a-zA-Z\.]+)(\d+)\s*(?:đến|-)\s*(?:\d+[a-zA-Z\.]+)?(\d+)', re.IGNORECASE)
    for match in range_pattern.finditer(s):
        prefix = match.group(1).upper()
        start_num = int(match.group(2))
        end_num = int(match.group(3))
        for i in range(start_num, end_num + 1):
            result.append(f"{prefix}{i}")
    
    s = range_pattern.sub(' ', s)
    parts = re.split(r'\s+', s)
    
    last_prefix = "" 
    for part in parts:
        part = part.strip()
        if not part: continue
        
        match_concat = re.match(r'^(\d+[a-zA-Z\.]+)(\d{2,})$', part)
        if match_concat:
            prefix, digits = match_concat.group(1).upper(), match_concat.group(2)
            for d in digits: result.append(f"{prefix}{d}")
            last_prefix = prefix
            continue
        
        if re.match(r'^\d+$', part) and last_prefix:
            result.append(f"{last_prefix}{part}")
            continue
        
        part_clean = part.upper()
        result.append(part_clean)
        
        prefix_match = re.match(r'^(\d+[a-zA-Z\.]+)', part_clean)
        if prefix_match: last_prefix = prefix_match.group(1)
    
    return result

def extract_all_classes(combo_list):
    """Trích xuất tất cả các lớp từ danh sách combo"""
    all_classes = set()
    for combo in combo_list:
        class_part = combo.split('-')[0]
        class_part = re.sub(r'\(.*\)', '', class_part)
        all_classes.add(class_part)
    return sorted(list(all_classes))

# ==========================================
# 3. HÀM XỬ LÝ DỮ LIỆU CHÍNH
# ==========================================
def process_data(df, nien_khoa):
    """Xử lý dữ liệu chính"""
    intermediate_data = []
    combo_tracker = defaultdict(set)
    all_combos_list = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for index, row in df.iterrows():
        stt = row.get('TT', index + 1)
        name = str(row.get('Họ và tên', '')).strip()
        
        # Xử lý ngày sinh
        dob = row.get('Ngày sinh', '')
        if pd.notna(dob):
            if isinstance(dob, datetime): 
                dob = dob.strftime("%d/%m/%Y")
            else:
                try: 
                    dob = pd.to_datetime(dob).strftime("%d/%m/%Y")
                except: 
                    dob = str(dob)
        else: 
            dob = ''
            
        pccm = row.get('PCCM', '')
        teacher_subjects = set()
        teacher_raw_combos = []
        
        if pd.notna(pccm) and str(pccm).strip() not in ['nan', '']:
            current_subj = ""
            parts = re.split(r'[;,\+\n]+', str(pccm))
            
            for part in parts:
                part = part.strip()
                if not part: continue
                
                match = re.search(r'\b\d{1,2}[a-zA-Z]+\.?\d*|\b\d{1,2}\.\d+', part)
                
                if match:
                    split_idx = match.start()
                    raw_subj = part[:split_idx].strip()
                    raw_classes = part[split_idx:].strip()
                    
                    raw_subj_clean = re.sub(r'[\(\):]', '', raw_subj).strip()
                    
                    if raw_subj_clean and re.search(r'[a-zA-Z]', raw_subj_clean):
                        current_subj = get_subject_code(raw_subj_clean)
                        
                    if current_subj:
                        teacher_subjects.add(current_subj)
                        parsed_classes = parse_classes(raw_classes)
                        for cls in parsed_classes:
                            combo = f"{cls}-{current_subj}"
                            teacher_raw_combos.append(combo)
                            all_combos_list.append(combo)
                else:
                    raw_subj_clean = re.sub(r'[\(\):]', '', part).strip()
                    if raw_subj_clean and re.search(r'[a-zA-Z]', raw_subj_clean):
                        current_subj = get_subject_code(raw_subj_clean)
            
            unique_combos = list(dict.fromkeys(teacher_raw_combos))
            for combo in unique_combos: 
                combo_tracker[combo].add(name)
                
            intermediate_data.append({
                'STT': stt, 'Name': name, 'DOB': dob,
                'Subjs': list(teacher_subjects), 'Combos': unique_combos
            })
        
        # Cập nhật progress
        progress_bar.progress((index + 1) / len(df))
        status_text.text(f"Đang xử lý: {index + 1}/{len(df)} dòng")
    
    # Xử lý trùng lặp
    processed_data = []
    
    for item in intermediate_data:
        final_class_subject = []
        teacher_name = item['Name']
        
        for combo in item['Combos']:
            if len(combo_tracker[combo]) > 1:
                final_class_subject.append(f"{combo}({teacher_name})")
            else:
                final_class_subject.append(combo)
                
        processed_data.append({
            'STT': item['STT'],
            'Họ và tên giáo viên': teacher_name,
            'Ngày sinh': item['DOB'],
            'SĐT': '', 
            'Môn học giảng dạy': ", ".join(item['Subjs']),
            'Trưởng BM': '', 
            'Lớp CN': '', 
            'Chi tiết Môn-Lớp': ",".join(final_class_subject)
        })
    
    # Tạo sheet Class
    all_classes = extract_all_classes(all_combos_list)
    
    class_data = []
    class_data.append({'A': 'Niên khóa', 'B': nien_khoa})
    class_data.append({'A': 'Lớp', 'B': 'Khối'})
    for cls in all_classes:
        match = re.search(r'^(\d+)', cls)
        khoi = match.group(1) if match else ''
        class_data.append({'A': cls, 'B': khoi})
    
    class_df = pd.DataFrame(class_data)
    
    # Tạo sheet Students
    students_columns = [
        'STT', 'Mã HS', 'Họ tên', 'Lớp', 'Giới tính', 
        'Ngày sinh', 'Số điện thoại', 'Email (nếu có)', 'Tài khoản'
    ]
    students_df = pd.DataFrame(columns=students_columns)
    
    teachers_df = pd.DataFrame(processed_data)
    
    progress_bar.empty()
    status_text.empty()
    
    return teachers_df, class_df, students_df, len(all_classes)

def get_download_link(teachers_df, class_df, students_df, nien_khoa):
    """Tạo link tải file Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        class_df.to_excel(writer, sheet_name='Class', index=False, header=False)
        teachers_df.to_excel(writer, sheet_name='Teachers', index=False)
        students_df.to_excel(writer, sheet_name='Students', index=False)
    
    output.seek(0)
    filename = f"Import_{nien_khoa.replace('-', '_')}.xlsx"
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" class="download-button">📥 Tải file Excel ({filename})</a>'
    return href

# ==========================================
# 4. GIAO DIỆN WEB VỚI STREAMLIT
# ==========================================
def main():
    st.set_page_config(
        page_title="Xử lý dữ liệu PCCM - Giáo viên",
        page_icon="📚",
        layout="wide"
    )
    
    # CSS tùy chỉnh
    st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .download-button {
        display: inline-block;
        background-color: #28a745;
        color: white;
        padding: 0.75rem 1.5rem;
        text-decoration: none;
        border-radius: 5px;
        font-weight: bold;
        margin: 1rem 0;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.markdown('<div class="main-header"><h1>📊 XỬ LÝ DỮ LIỆU PCCM</h1><p>Phân công chuyên môn giáo viên</p></div>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ CẤU HÌNH")
        
        nien_khoa = st.selectbox(
            "Niên khóa",
            ["2025-2026", "2026-2027", "2027-2028"],
            index=0
        )
        
        st.markdown("---")
        st.header("📁 TẢI LÊN FILE")
        
        uploaded_file = st.file_uploader(
            "Chọn file Excel đầu vào",
            type=['xlsx', 'xls'],
            help="File cần có sheet 'Data' với các cột: TT, Họ và tên, Ngày sinh, PCCM"
        )
        
        st.markdown("---")
        st.header("📝 HƯỚNG DẪN")
        st.markdown("""
        1. Chuẩn bị file Excel có sheet **Data**
        2. Sheet Data cần có các cột:
           - **TT**: Số thứ tự
           - **Họ và tên**: Tên giáo viên
           - **Ngày sinh**: Ngày sinh
           - **PCCM**: Phân công chuyên môn
        3. Tải file lên hệ thống
        4. Chọn niên khóa
        5. Nhấn **XỬ LÝ DỮ LIỆU**
        6. Tải file kết quả
        """)
    
    # Main content
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Data')
            
            
            # ---> THÊM DÒNG NÀY VÀO ĐỂ CHUẨN HÓA CỘT <---
            df = standardize_columns(df)
            
            # Hiển thị thông tin file
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📊 Số dòng dữ liệu", len(df))
            with col2:
                st.metric("📋 Số cột", len(df.columns))
            with col3:
                required_cols = ['TT', 'Họ và tên', 'Ngày sinh', 'PCCM']
                has_all = all(col in df.columns for col in required_cols)
                if has_all:
                    st.metric("✅ Định dạng", "Hợp lệ")
                else:
                    st.metric("❌ Định dạng", "Thiếu cột")
            
            # Xem trước dữ liệu
            with st.expander("🔍 Xem trước dữ liệu (10 dòng đầu)"):
                st.dataframe(df.head(10), use_container_width=True)
            
            # Nút xử lý
            if st.button("🚀 XỬ LÝ DỮ LIỆU", type="primary", use_container_width=True):
                with st.spinner("Đang xử lý dữ liệu..."):
                    try:
                        teachers_df, class_df, students_df, num_classes = process_data(df, nien_khoa)
                        
                        # Hiển thị kết quả
                        st.markdown('<div class="success-box">', unsafe_allow_html=True)
                        st.success("✅ XỬ LÝ THÀNH CÔNG!")
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("🏫 Niên khóa", nien_khoa)
                        with col2:
                            st.metric("📚 Số lớp", num_classes)
                        with col3:
                            st.metric("👩‍🏫 Số giáo viên", len(teachers_df))
                        st.markdown('</div>', unsafe_allow_html=True)
                        
                        # Tab hiển thị kết quả
                        tab1, tab2, tab3 = st.tabs(["👩‍🏫 Teachers", "🏫 Class", "👨‍🎓 Students"])
                        
                        with tab1:
                            st.dataframe(teachers_df, use_container_width=True, height=400)
                        with tab2:
                            st.dataframe(class_df, use_container_width=True, height=400)
                        with tab3:
                            st.dataframe(students_df, use_container_width=True)
                        
                        # Tải file
                        st.markdown("---")
                        st.markdown("### 💾 TẢI KẾT QUẢ")
                        download_link = get_download_link(teachers_df, class_df, students_df, nien_khoa)
                        st.markdown(download_link, unsafe_allow_html=True)
                        
                    except Exception as e:
                        st.error(f"❌ Lỗi xử lý: {str(e)}")
                        st.exception(e)
        
        except Exception as e:
            st.error(f"❌ Lỗi đọc file: {str(e)}")
            st.info("Vui lòng đảm bảo file có sheet 'Data' và đúng định dạng")
    else:
        st.info("👈 Vui lòng tải lên file Excel từ thanh bên trái để bắt đầu")

if __name__ == "__main__":
    main()