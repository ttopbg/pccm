"""
app.py  –  Streamlit web app, KHÔNG cần API key
Deploy: push lên GitHub (app.py + teacher_core.py + requirements.txt) -> streamlit.io/cloud
"""
import io
import streamlit as st
import pandas as pd
from teacher_core import (process_data, detect_header_row, find_column,
                           detect_ambiguous_in_data, detect_unknown_subjects,
                           _ALL_CODES, CAP_HOC_OPTIONS,
                           build_known_classes_from_gvcn)

NIEN_KHOA_OPTIONS = ["2025-2026", "2026-2027", "2027-2028"]

st.set_page_config(page_title="PCCM", page_icon="🔞", layout="wide")

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
:root {
  --step-bg:        rgba(46, 117, 182, 0.10);
  --step-border:    #2E75B6;
  --success-bg:     rgba(67, 160, 71, 0.12);
  --success-border: #43a047;
  --warn-bg:        rgba(255, 152, 0, 0.10);
  --warn-border:    #ff9800;
  --help-bg:        rgba(46, 117, 182, 0.07);
  --help-border:    rgba(46, 117, 182, 0.30);
  --code-bg:        rgba(0,0,0,0.06);
  --tag-green-bg:   rgba(67,160,71,0.15);
  --tag-green-fg:   #2e7d32;
  --tag-blue-bg:    rgba(46,117,182,0.15);
  --tag-blue-fg:    #1a5296;
  --tag-orange-bg:  rgba(230,119,0,0.15);
  --tag-orange-fg:  #a05000;
}
@media (prefers-color-scheme: dark) {
  :root {
    --step-bg:        rgba(46, 117, 182, 0.20);
    --step-border:    #5ba3d9;
    --success-bg:     rgba(67, 160, 71, 0.20);
    --success-border: #66bb6a;
    --warn-bg:        rgba(255, 152, 0, 0.18);
    --warn-border:    #ffa726;
    --help-bg:        rgba(46, 117, 182, 0.12);
    --help-border:    rgba(91,163,217,0.35);
    --code-bg:        rgba(255,255,255,0.08);
    --tag-green-bg:   rgba(67,160,71,0.22);
    --tag-green-fg:   #81c784;
    --tag-blue-bg:    rgba(46,117,182,0.22);
    --tag-blue-fg:    #90caf9;
    --tag-orange-bg:  rgba(230,119,0,0.22);
    --tag-orange-fg:  #ffb74d;
  }
}
[data-theme="dark"] {
  --step-bg:rgba(46,117,182,.20);--step-border:#5ba3d9;
  --success-bg:rgba(67,160,71,.20);--success-border:#66bb6a;
  --warn-bg:rgba(255,152,0,.18);--warn-border:#ffa726;
  --help-bg:rgba(46,117,182,.12);--help-border:rgba(91,163,217,.35);
  --code-bg:rgba(255,255,255,.08);
  --tag-green-bg:rgba(67,160,71,.22);--tag-green-fg:#81c784;
  --tag-blue-bg:rgba(46,117,182,.22);--tag-blue-fg:#90caf9;
  --tag-orange-bg:rgba(230,119,0,.22);--tag-orange-fg:#ffb74d;
}
.main-header{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:white;
  padding:1.4rem 2rem;border-radius:12px;margin-bottom:1.4rem;text-align:center}
.main-header h1{margin:0;font-size:1.7rem}
.main-header p{margin:.35rem 0 0;opacity:.85;font-size:.9rem}
.step-box{background:var(--step-bg);border-left:4px solid var(--step-border);
  padding:.75rem 1rem;border-radius:0 8px 8px 0;margin-bottom:1rem}
.success-box{background:var(--success-bg);border-left:4px solid var(--success-border);
  padding:.75rem 1rem;border-radius:0 8px 8px 0}
.warn-box{background:var(--warn-bg);border-left:4px solid var(--warn-border);
  padding:.75rem 1rem;border-radius:0 8px 8px 0;margin-bottom:1.2rem}
.ambig-card{background:var(--help-bg);border:1.5px solid var(--warn-border);
  border-radius:10px;padding:.9rem 1.1rem;margin-bottom:1.1rem}
.ambig-token{font-size:1.2rem;font-weight:800;font-family:monospace;color:var(--warn-border)}
.ambig-ctx{font-size:.82rem;opacity:.7;margin:.2rem 0 .5rem;font-style:italic}
.help-section{background:var(--help-bg);border:1px solid var(--help-border);
  border-radius:10px;padding:1rem 1.1rem;margin-bottom:.9rem;font-size:.88rem;line-height:1.6}
.help-section h4{margin:0 0 .5rem;font-size:.95rem;font-weight:700}
.help-section ul{margin:.3rem 0 0 1rem;padding:0}
.help-section li{margin:.25rem 0}
.help-section code{background:var(--code-bg);padding:.1rem .35rem;
  border-radius:4px;font-size:.82rem}
.help-section table{width:100%;border-collapse:collapse;font-size:.82rem;margin-top:.5rem}
.help-section th{text-align:left;padding:.3rem .5rem;opacity:.7;font-weight:600;
  border-bottom:1px solid var(--help-border)}
.help-section td{padding:.3rem .5rem;border-bottom:1px solid var(--help-border)}
.tag{display:inline-block;padding:.05rem .45rem;border-radius:20px;font-size:.78rem;font-weight:600}
.tag-green{background:var(--tag-green-bg);color:var(--tag-green-fg)}
.tag-blue {background:var(--tag-blue-bg); color:var(--tag-blue-fg)}
.tag-orange{background:var(--tag-orange-bg);color:var(--tag-orange-fg)}
.example-row{font-family:monospace;font-size:.8rem;word-break:break-all}
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📖 Hướng dẫn sử dụng")

    with st.expander("📁  1. Chuẩn bị file đầu vào", expanded=False):
        st.markdown("""
<div class="help-section">
<h4>Yêu cầu file Excel</h4>
<ul>
  <li>Định dạng: <code>.xlsx</code>, <code>.xls</code>, <code>.xlsm</code></li>
  <li>Phải có sheet tên <b>Data</b> (không phân biệt hoa/thường)</li>
  <li>Nếu không có sheet Data, sheet đầu tiên sẽ được dùng</li>
</ul>
<h4>Các cột bắt buộc trong sheet Data</h4>
<table>
  <tr><th>Cột</th><th>Tên chấp nhận được</th></tr>
  <tr><td><span class="tag tag-blue">Họ tên</span></td><td>Họ tên, Họ và tên, Giáo viên…</td></tr>
  <tr><td><span class="tag tag-blue">PCCM</span></td><td>PCCM, Phân công, Môn học giảng dạy…</td></tr>
  <tr><td><span class="tag tag-orange">GVCN</span></td><td>GVCN, Chủ nhiệm, Chủ nhiệm lớp, CN…</td></tr>
</table>
<h4>Vai trò của cột GVCN</h4>
<ul>
  <li>Điền vào cột <b>CN</b> trong file output</li>
  <li>Xây dựng <b>danh sách lớp hợp lệ</b> để tách PCCM chính xác</li>
  <li>Khi gặp chuỗi có nhiều cách tách, hệ thống sẽ <b>hỏi người dùng</b></li>
</ul>
</div>
""", unsafe_allow_html=True)

    with st.expander("🏫  2. Chọn cấp học", expanded=False):
        st.markdown("""
<div class="help-section">
<h4>Tại sao phải chọn cấp học?</h4>
<p>Mỗi cấp có bảng mã môn học riêng. Cùng tên môn nhưng mã sẽ khác nhau tùy cấp.</p>
<table>
  <tr><th>Cấp</th><th>Khối</th><th>Ví dụ mã môn đặc trưng</th></tr>
  <tr><td><b>Tiểu học</b></td><td>1–5</td><td>TIENGVIET, TUNHIENVAXAHOI, DAODUC, THCN, KHOAHOC, HDTN</td></tr>
  <tr><td><b>THCS</b></td><td>6–9</td><td>KHTN(VATLY), KHTN(HOAHOC), KHTN(SINH), LICHSUDIALI(SU), LICHSUDIALI(DIA), GDCD, TNHN</td></tr>
  <tr><td><b>THPT</b></td><td>10–12</td><td>VATLY, HOAHOC, SINH, LICHSU, DIALY, GDKTPL, TNHN</td></tr>
</table>
<h4>Lưu ý khi chọn THCS</h4>
<ul>
  <li>Vật lý / Lý → <code>KHTN(VATLY)</code> (không phải VATLY như THPT)</li>
  <li>Lịch sử → <code>LICHSUDIALI(SU)</code>, Địa lý → <code>LICHSUDIALI(DIA)</code></li>
  <li>Hoạt động trải nghiệm (HDTN) → <code>TNHN</code></li>
</ul>
</div>
""", unsafe_allow_html=True)

    with st.expander("✍️  2. Chú ý về cột PCCM", expanded=False):
        st.markdown("""
<div class="help-section">
<h4>Cấu trúc cơ bản</h4>
<p>Dữ liệu theo dạng <b>Tên môn: danh sách lớp</b>, nhiều môn ngăn cách bằng <code>+</code></p>
<div class="example-row">Hóa: 10A1, 10A2 + Sử: 10D1, 10D2</div>

<h4>Tách lớp thông minh nhờ cột GVCN</h4>
<p>Khi có cột GVCN, hệ thống xây dựng danh sách lớp hợp lệ. Nếu một chuỗi có <b>nhiều cách tách hợp lệ</b>, hệ thống hiển thị hộp hỏi để bạn chọn.</p>
<table>
  <tr><th>Chuỗi</th><th>known = {10A1…10A12}</th></tr>
  <tr><td><code>10A123</code></td><td>⚠️ Hỏi: [10A1,2,3] hay [10A12,3]?</td></tr>
  <tr><td><code>10A12</code></td><td>→ 10A12 ✅ (chỉ 1 cách)</td></tr>
  <tr><td><code>10A1-10A5</code></td><td>→ 10A1..10A5 (range rõ ràng)</td></tr>
  <tr><td><code>10ABC, 10A,B,C</code></td><td><b>Chưa phâm biệt được chữ, cần thêm lớp ở trước!</b></td></tr>
</table>

<h4>Các định dạng lớp được hỗ trợ</h4>
<table>
  <tr><th>Dạng viết</th><th>Kết quả</th></tr>
  <tr><td><code>10A1, 10A2, 10A3</code></td><td>Liệt kê thường</td></tr>
  <tr><td><code>10A123</code></td><td>→ tách theo known (hoặc hỏi)</td></tr>
  <tr><td><code>10A1-10A5</code></td><td>→ 10A1 đến 10A5</td></tr>
  <tr><td><code>11A3,4</code></td><td>→ 11A3, 11A4</td></tr>
  <tr><td><code>11A1(52)</code></td><td>→ 11A1 (bỏ sĩ số)</td></tr>
</table>
</div>
""", unsafe_allow_html=True)

#     with st.expander("🔤  3. Nhận diện tên môn học", expanded=False):
#         st.markdown("""
# <div class="help-section">
# <h4>Bảng mã môn học</h4>
# <table>
#   <tr><th>Tên môn</th><th>Mã</th></tr>
#   <tr><td>Ngữ văn / Văn</td><td><code>NGUVAN</code></td></tr>
#   <tr><td>Toán / Toán học</td><td><code>TOAN</code></td></tr>
#   <tr><td>Tiếng Anh / Anh / NN1</td><td><code>ANH</code></td></tr>
#   <tr><td>Lịch sử / Sử</td><td><code>LICHSU</code></td></tr>
#   <tr><td>Địa lý / Địa</td><td><code>DIALY</code></td></tr>
#   <tr><td>Vật lý / Lý</td><td><code>VATLY</code></td></tr>
#   <tr><td>Hóa học / Hóa</td><td><code>HOAHOC</code></td></tr>
#   <tr><td>Sinh học / Sinh</td><td><code>SINH</code></td></tr>
#   <tr><td>Tin học / Tin</td><td><code>TINHOC</code></td></tr>
#   <tr><td>GDTC / Thể dục</td><td><code>GDTC</code></td></tr>
#   <tr><td>GDQP / Quốc phòng</td><td><code>GDQP</code></td></tr>
#   <tr><td>KTPL / GDKTPL</td><td><code>GDKTPL</code></td></tr>
#   <tr><td>GDĐP / GDDP</td><td><code>NDGDDP</code></td></tr>
#   <tr><td>HĐTN / TNHN</td><td><code>TNHN</code></td></tr>
#   <tr><td>Công nghệ</td><td><code>CONGNGHE</code></td></tr>
#   <tr><td>KHTN</td><td><code>KHTN</code></td></tr>
#   <tr><td>Lịch sử &amp; Địa lý</td><td><code>LICHSUDIALI</code></td></tr>
# </table>
# </div>
# """, unsafe_allow_html=True)

#     with st.expander("📊  4. Cấu trúc file đầu ra", expanded=False):
#         st.markdown("""
# <div class="help-section">
# <h4>File output gồm 3 sheet</h4>
# <p><span class="tag tag-blue">Sheet 1: Class</span> — danh sách lớp, sắp theo khối</p>
# <p><span class="tag tag-green">Sheet 2: Teachers</span></p>
# <table>
#   <tr><th>Cột</th><th>Nội dung</th></tr>
#   <tr><td>STT</td><td>Số thứ tự</td></tr>
#   <tr><td>Họ tên</td><td>Tên giáo viên</td></tr>
#   <tr><td>Ngày sinh</td><td>dd/mm/yyyy</td></tr>
#   <tr><td>SĐT</td><td>Để trống</td></tr>
#   <tr><td>Môn dạy</td><td>Mã môn, cách nhau dấu phẩy</td></tr>
#   <tr><td>TBM</td><td>Để trống</td></tr>
#   <tr><td>CN</td><td>Lớp chủ nhiệm (từ cột GVCN)</td></tr>
#   <tr><td>PCCM</td><td><code>10A1-TOAN,11B2-ANH</code></td></tr>
# </table>
# <p><span class="tag tag-orange">Sheet 3: Students</span> — tiêu đề cố định, dữ liệu trống</p>
# </div>
# """, unsafe_allow_html=True)

#     with st.expander("⚠️  5. Xử lý trùng lặp & lưu ý", expanded=False):
#         st.markdown("""
# <div class="help-section">
# <h4>Xử lý tổ hợp môn-lớp trùng</h4>
# <ul>
#   <li><b>Trùng trong cùng 1 GV:</b> bỏ tự động</li>
#   <li><b>Trùng giữa 2+ GV:</b> thêm tên GV để phân biệt</li>
# </ul>
# <div class="example-row">12A2-HOAHOC(Nguyễn Tuấn Anh)<br>12A2-HOAHOC(Đoàn Văn Chiến)</div>
# <h4>Không có cột GVCN</h4>
# <ul>
#   <li>Cột CN trong output để trống</li>
#   <li>Không có từ điển lớp → không hỏi ambiguous, dùng logic tách cũ</li>
# </ul>
# </div>
# """, unsafe_allow_html=True)


# ── SESSION STATE ─────────────────────────────────────────────────────────────
for k, v in [("phase","upload"),("ambig_list",[]),("resolved",{}),
             ("unknown_list",[]),("resolved_subjects",{}),
             ("raw_bytes",None),("cap_hoc","THPT"),("nien_khoa",NIEN_KHOA_OPTIONS[0]),
             ("known_classes",set()),("result_bytes",None),("result_filename","")]:
    if k not in st.session_state:
        st.session_state[k] = v


def _reset():
    for k in ["phase","ambig_list","resolved","unknown_list","resolved_subjects",
              "raw_bytes","cap_hoc","nien_khoa","known_classes","result_bytes","result_filename"]:
        st.session_state.pop(k, None)
    st.rerun()


def _load_df_and_known(raw_bytes):
    xl  = pd.ExcelFile(io.BytesIO(raw_bytes))
    sn  = next((s for s in xl.sheet_names if s.strip().lower()=="data"), xl.sheet_names[0])
    rdf = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sn, header=None)
    hri = detect_header_row(rdf)
    df  = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sn, header=hri)
    df.columns = [str(c).strip() for c in df.columns]
    col_gvcn = find_column(df, ["gvcn","chủ nhiệm","chu nhiem","chủ nhiệm lớp",
                                 "chu nhiem lop","lớp chủ nhiệm","lop chu nhiem","cn"])
    known = build_known_classes_from_gvcn(df, col_gvcn)
    return df, col_gvcn, known


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🙃 Tạo file Import PCCM 🙃</h1>
  <p>File Input cần có sheet <b>Data</b> (hoặc sẽ lấy sheet đầu tiên của file)</p>
</div>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# PHASE 1 — UPLOAD
# ════════════════════════════════════════════════════════════════════════════
if st.session_state.phase == "upload":

    st.markdown('<div class="step-box"><b>1️⃣</b> Tải lên file cần chuyển đổi PCCM</div>',
                unsafe_allow_html=True)
    uploaded = st.file_uploader("Chọn file Excel", type=["xlsx","xls","xlsm"],
                                 label_visibility="collapsed")

    st.markdown('<div class="step-box"><b>2️⃣</b> Chọn cấp học</div>', unsafe_allow_html=True)
    _CAP_HOC_LABELS = {"THPT": "🎓 THPT (khối 10–12)", "THCS": "📚 THCS (khối 6–9)", "TH": "🏫 Tiểu học (khối 1–5)"}
    cap_hoc = st.radio(
        "Cấp học",
        options=CAP_HOC_OPTIONS,
        format_func=lambda x: _CAP_HOC_LABELS[x],
        horizontal=True,
        label_visibility="collapsed",
    )

    st.markdown('<div class="step-box"><b>3️⃣</b> Chọn niên khóa</div>', unsafe_allow_html=True)
    nien_khoa = st.selectbox("Niên khóa", options=NIEN_KHOA_OPTIONS, label_visibility="collapsed")

    st.markdown('<div class="step-box"><b>4️⃣</b> Nhấn <code><b>Chuyển đổi</b></code> để xử lý</div>', unsafe_allow_html=True)
    run_btn = st.button("▶  Chuyển đổi", type="primary", use_container_width=True,
                        disabled=(uploaded is None))

    if uploaded is None:
        st.info("✌️ Vui lòng tải lên file Excel đầu vào ✌️")

    if run_btn and uploaded:
        raw_bytes = uploaded.read()
        fname     = uploaded.name.rsplit(".",1)[0]

        with st.spinner("Đang phân tích file..."):
            try:
                df, col_gvcn, known = _load_df_and_known(raw_bytes)
                col_pccm = find_column(df, ["pccm","phân công chuyên môn","phân công",
                                             "giảng dạy lớp","môn học giảng dạy","pcan cong","giang day"])
                ambig_list   = []
                unknown_list = []
                if col_pccm:
                    if known:
                        ambig_list = detect_ambiguous_in_data(df, col_pccm, col_gvcn, known)
                    unknown_list = detect_unknown_subjects(df, col_pccm, cap_hoc)
            except Exception as e:
                st.error(f"❌ Lỗi đọc file: {e}")
                st.stop()

        st.session_state.raw_bytes       = raw_bytes
        st.session_state.cap_hoc         = cap_hoc
        st.session_state.nien_khoa       = nien_khoa
        st.session_state.known_classes   = known
        st.session_state.result_filename = f"Import_{fname}.xlsx"

        if ambig_list:
            st.session_state.ambig_list    = ambig_list
            st.session_state.unknown_list  = unknown_list
            st.session_state.resolved      = {}
            st.session_state.phase         = "confirm_ambig"
        elif unknown_list:
            st.session_state.ambig_list    = []
            st.session_state.unknown_list  = unknown_list
            st.session_state.resolved_subjects = {}
            st.session_state.phase         = "confirm_subjects"
        else:
            st.session_state.ambig_list    = []
            st.session_state.unknown_list  = []
            st.session_state.phase         = "processing"
        st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# PHASE 2 — HỎI AMBIGUOUS
# ════════════════════════════════════════════════════════════════════════════
elif st.session_state.phase == "confirm_ambig":
    ambig_list = st.session_state.ambig_list

    st.markdown(f"""
<div class="warn-box">
⚠️ <b>Phát hiện {len(ambig_list)} chuỗi lớp có thể tách theo nhiều cách.</b><br>
Vui lòng chọn cách tách đúng cho từng chuỗi bên dưới, sau đó nhấn <b>Xác nhận & Tiếp tục</b>.
</div>
""", unsafe_allow_html=True)

    choices = {}

    for item in ambig_list:
        token  = item["token"]
        splits = item["splits"]
        occs   = item["occurrences"]
        option_labels = [", ".join(s) for s in splits]

        ctx_text = " &nbsp;|&nbsp; ".join(occs[:3]) + ("…" if len(occs) > 3 else "")
        st.markdown(f"""
<div class="ambig-card">
  <div>Chuỗi gốc: <span class="ambig-token">{token}</span></div>
  <div class="ambig-ctx">📍 {ctx_text}</div>
</div>
""", unsafe_allow_html=True)

        chosen = st.radio(
            f"**`{token}`** là:",
            options=option_labels,
            key=f"radio_{token}",
            horizontal=len(option_labels) <= 4,
        )
        choices[token] = splits[option_labels.index(chosen)]
        st.markdown("---")

    col_back, col_confirm = st.columns([1, 3])
    with col_back:
        if st.button("← Quay lại", use_container_width=True):
            _reset()
    with col_confirm:
        if st.button("✅  Xác nhận & Tiếp tục", type="primary", use_container_width=True):
            st.session_state.resolved = choices
            # Sau ambig → kiểm tra xem có môn không nhận ra không
            if st.session_state.get("unknown_list"):
                st.session_state.resolved_subjects = {}
                st.session_state.phase = "confirm_subjects"
            else:
                st.session_state.phase = "processing"
            st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# PHASE 2b — HỎI MÔN KHÔNG NHẬN RA
# ════════════════════════════════════════════════════════════════════════════
elif st.session_state.phase == "confirm_subjects":
    unknown_list = st.session_state.unknown_list
    _cap = st.session_state.get("cap_hoc", "THPT")
    _CAP_LABEL = {"TH": "Tiểu học", "THCS": "THCS", "THPT": "THPT"}.get(_cap, _cap)
    # Chỉ hiển thị mã môn phù hợp với cấp học đang xử lý
    from teacher_core import _get_subject_map
    _level_codes = sorted(set(_get_subject_map(_cap).values()))
    KEEP_RAW = "— Giữ nguyên (không map) —"
    code_options = [KEEP_RAW] + _level_codes

    st.markdown(f"""
<div class="warn-box">
⚠️ <b>Phát hiện {len(unknown_list)} tên môn chưa nhận diện được.</b><br>
Vui lòng chọn mã môn tương ứng cho từng tên bên dưới, sau đó nhấn <b>Xác nhận & Tiếp tục</b>.
</div>
""", unsafe_allow_html=True)

    subj_choices = {}  # raw_lower → code

    for item in unknown_list:
        raw        = item["raw"]
        suggestion = item["suggestion"]
        occs       = item["occurrences"]

        # Index mặc định: dùng gợi ý nếu có
        default_idx = code_options.index(suggestion) if suggestion and suggestion in code_options else 0

        ctx_text = " &nbsp;|&nbsp; ".join(occs[:3]) + ("…" if len(occs) > 3 else "")
        st.markdown(f"""
<div class="ambig-card">
  <div>Tên môn gốc: <span class="ambig-token">{raw}</span>
    {"&nbsp; 💡 Gợi ý: <b>" + suggestion + "</b>" if suggestion else ""}
  </div>
  <div class="ambig-ctx">📍 {ctx_text}</div>
</div>
""", unsafe_allow_html=True)

        chosen = st.selectbox(
            f"Mã môn cho **\"{raw}\"**:",
            options=code_options,
            index=default_idx,
            key=f"subj_{raw}",
        )
        subj_choices[raw.lower().strip()] = None if chosen == KEEP_RAW else chosen
        st.markdown("---")

    col_back, col_confirm = st.columns([1, 3])
    with col_back:
        if st.button("← Quay lại", use_container_width=True):
            # Nếu trước đó có ambig thì về confirm_ambig, không thì về upload
            if st.session_state.get("ambig_list"):
                st.session_state.phase = "confirm_ambig"
            else:
                _reset()
            st.rerun()
    with col_confirm:
        if st.button("✅  Xác nhận & Tiếp tục", type="primary", use_container_width=True):
            # Chỉ lưu những môn đã chọn mã (bỏ qua KEEP_RAW)
            st.session_state.resolved_subjects = {
                k: v for k, v in subj_choices.items() if v is not None
            }
            st.session_state.phase = "processing"
            st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# PHASE 3 — XỬ LÝ
# ════════════════════════════════════════════════════════════════════════════
elif st.session_state.phase == "processing":
    log_area  = st.empty()
    prog_bar  = st.progress(0)
    log_lines = []

    raw_bytes = st.session_state.raw_bytes
    nien_khoa = st.session_state.nien_khoa
    cap_hoc   = st.session_state.get("cap_hoc", "THPT")
    resolved          = st.session_state.get("resolved", {})
    resolved_subjects = st.session_state.get("resolved_subjects", {})

    try:
        _xl  = pd.ExcelFile(io.BytesIO(raw_bytes))
        _sn  = next((s for s in _xl.sheet_names if s.strip().lower()=="data"), _xl.sheet_names[0])
        _rdf = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=_sn, header=None)
        _hri = detect_header_row(_rdf)
        _df  = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=_sn, header=_hri)
        _df.columns = [str(c).strip() for c in _df.columns]
        _ch  = find_column(_df, ["họ tên","họ và tên","tên","giáo viên","ho ten"])
        total_t = len(_df[_df[_ch].notna()]) if _ch else 1
    except Exception:
        total_t = 1

    processed = [0]
    def progress_cb(msg):
        log_lines.append(msg)
        log_area.code("\n".join(log_lines[-20:]), language=None)
        if "Xử lý giáo viên" in msg:
            processed[0] += 1
            prog_bar.progress(min(int(processed[0]/total_t*90), 90))

    try:
        result_bytes = process_data(
            io.BytesIO(raw_bytes), nien_khoa,
            cap_hoc=cap_hoc,
            progress_cb=progress_cb,
            resolved_ambiguities=resolved,
            resolved_subjects=resolved_subjects,
        )
        prog_bar.progress(100)
        st.session_state.result_bytes = result_bytes
        st.session_state.phase        = "done"
        st.rerun()
    except Exception as e:
        prog_bar.empty()
        st.error(f"❌ Lỗi: {e}")
        if st.button("← Làm lại từ đầu"):
            _reset()


# ════════════════════════════════════════════════════════════════════════════
# PHASE 4 — DONE
# ════════════════════════════════════════════════════════════════════════════
elif st.session_state.phase == "done":
    _cap_done  = st.session_state.get("cap_hoc", "THPT")
    _nk_done   = st.session_state.get("nien_khoa", "")
    _cap_label = {"TH": "Tiểu học", "THCS": "THCS", "THPT": "THPT"}.get(_cap_done, _cap_done)

    summary_lines = []
    if st.session_state.get("resolved"):
        for tok, cls_list in st.session_state.resolved.items():
            summary_lines.append(f"• Lớp `{tok}` → **{', '.join(cls_list)}**")
    if st.session_state.get("resolved_subjects"):
        for raw, code in st.session_state.resolved_subjects.items():
            summary_lines.append(f"• Môn **\"{raw}\"** → `{code}`")
    if summary_lines:
        with st.expander(f"ℹ️ {len(summary_lines)} mục đã xác nhận thủ công", expanded=False):
            st.markdown("\n".join(summary_lines))

    st.markdown(f'<div class="success-box">✅ <b>Chuyển đổi thành công!</b> &nbsp;'
                f'Cấp học: <b>{_cap_label}</b> &nbsp;|&nbsp; Niên khóa: <b>{_nk_done}</b></div>',
                unsafe_allow_html=True)
    st.download_button(
        "⬇️  Tải xuống file Excel",
        data=st.session_state.result_bytes,
        file_name=st.session_state.result_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown("---")
    if st.button("🔄  Chuyển đổi file khác", use_container_width=True):
        _reset()
