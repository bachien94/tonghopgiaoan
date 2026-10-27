import datetime
import io
import copy
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docxcompose.composer import Composer  # Thư viện chuyên dùng để ghép file docx chuẩn xác (giữ nguyên ảnh, bảng, style)

# Cấu hình trang rộng
st.set_page_config(page_title="Tổng Hợp Giáo Án Theo Tuần", page_icon="📚", layout="wide")

# --- CSS TÙY CHỈNH GIAO DIỆN GỌN GÀNG, MƯỢT MÀ ---
st.markdown("""
<style>
    .element-container {
        margin-bottom: -4px !important;
    }
    h2 {
        margin-bottom: -15px !important;
    }
    hr {
        margin-top: 4px !important;
        margin-bottom: 4px !important;
    }
    div[data-baseweb="select"] > div {
        min-height: 30px !important;
        padding-top: 0px !important;
        padding-bottom: 0px !important;
    }
    div[data-baseweb="select"] div {
        font-size: 13px !important;
    }
    .stFileUploader section {
        padding: 2px 6px !important;
        font-size: 11px !important;
        min-height: 30px !important;
    }
    div[data-testid="column"] {
        padding: 0px 3px !important;
    }
    .stTabs [data-baseweb="tab-panel"] {
        padding-top: 8px !important;
        padding-bottom: 4px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- TIÊU ĐỀ ỨNG DỤNG ---
st.markdown("## 📊 Phần Mềm Tổng Hợp Giáo Án Theo Tuần")
st.markdown("---")

# --- PHẦN 1: ĐIỀU KHIỂN CHUNG & TÙY CHỈNH HEADER/FOOTER ---
col_c1, col_c2, col_c3 = st.columns([1, 1.5, 2])
with col_c1:
    tuan_so = st.number_input("📅 Tuần số:", min_value=1, max_value=52, value=1)
with col_c2:
    ngay_dau_tuan = st.date_input("🗓️ Ngày Thứ Hai đầu tuần:", value=datetime.date.today())
with col_c3:
    if 'generated_docx' not in st.session_state:
        st.session_state['generated_docx'] = None

# Tùy chỉnh Header và Footer
with st.expander("⚙️ Tùy chỉnh Tiêu đề trang (Header) & Chân trang (Footer) cho file Word"):
    st.markdown("##### 📌 Tùy chỉnh Header (Đầu trang - Có kẻ ngang sát chữ)")
    col_hf1, col_hf2 = st.columns(2)
    with col_hf1:
        header_left = st.text_input("📝 Header (Bên trái):", value="KBBD: Đa môn")
    with col_hf2:
        header_right = st.text_input("📝 Header (Bên phải):", value="Năm học: 2026 - 2027")
    
    st.markdown("##### 📌 Tùy chỉnh Footer (Chân trang - Cỡ chữ 13, có kẻ phía trên)")
    col_hf3, col_hf4 = st.columns(2)
    with col_hf3:
        footer_left = st.text_input("📝 Footer (Bên trái):", value="Tên Giáo viên")
    with col_hf4:
        footer_right = st.text_input("📝 Footer (Bên phải):", value="Trường Tiểu học Ba Tơ")

st.markdown("---")

# --- DANH MỤC DỮ LIỆU ---
danh_sach_mon = [
    "-- Chọn môn học --", "Tin học", "Toán", "Tiếng Việt", "Đạo đức", 
    "Tự nhiên và Xã hội", "Khoa học", "Lịch sử và Địa lý", 
    "Mỹ thuật", "Âm nhạc", "Thể dục", "Hoạt động trải nghiệm", "Tiếng Anh"
]

danh_sach_khoi = ["-- Chọn khối --", "Khối 1", "Khối 2", "Khối 3", "Khối 4", "Khối 5"]

mapping_lop = {
    "Khối 1": [f"1A{i}" for i in range(1, 9)],
    "Khối 2": [f"2A{i}" for i in range(1, 8)],
    "Khối 3": [f"3A{i}" for i in range(1, 8)],
    "Khối 4": [f"4A{i}" for i in range(1, 9)],
    "Khối 5": [f"5A{i}" for i in range(1, 8)]
}

cac_thu = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]

ngay_offset = {
    "Thứ Hai": 0, "Thứ Ba": 1, "Thứ Tư": 2, "Thứ Năm": 3, "Thứ Sáu": 4
}

dieu_phoi_tiet = [
    ("Buổi Sáng", "Tiết 1"), ("Buổi Sáng", "Tiết 2"), ("Buổi Sáng", "Tiết 3"), 
    ("Buổi Sáng", "Tiết 4"), ("Buổi Sáng", "Tiết 5"),
    ("Buổi Chiều", "Tiết 1"), ("Buổi Chiều", "Tiết 2"),
]

# --- PHẦN 2: GIAO DIỆN NHẬP LIỆU DẠNG TAB ---
st.markdown("### 📋 Nhập Thời Khóa Biểu & Giáo Án Theo Từng Ngày")
st.caption("💡 *Mẹo: Bấm vào từng Thẻ (Tab) bên dưới để nhập liệu riêng cho từng ngày trong tuần.*")

tabs = st.tabs([f"📌 {thu}" for thu in cac_thu])
tkb_data = {}

for idx, thu in enumerate(cac_thu):
    with tabs[idx]:
        st.markdown(f"<p style='margin-bottom: 2px; font-weight: bold;'>Thời khóa biểu ngày {thu}:</p>", unsafe_allow_html=True)
        
        col_h = st.columns([1.1, 2.0, 1.8, 1.8, 3.4])
        col_h[0].markdown("<div style='font-size:12px; font-weight:bold; line-height: 1.1;'>Buổi / Tiết</div>", unsafe_allow_html=True)
        col_h[1].markdown("<div style='font-size:12px; font-weight:bold; line-height: 1.1;'>Môn học</div>", unsafe_allow_html=True)
        col_h[2].markdown("<div style='font-size:12px; font-weight:bold; line-height: 1.1;'>Khối</div>", unsafe_allow_html=True)
        col_h[3].markdown("<div style='font-size:12px; font-weight:bold; line-height: 1.1;'>Lớp</div>", unsafe_allow_html=True)
        col_h[4].markdown("<div style='font-size:12px; font-weight:bold; line-height: 1.1;'>Đính kèm file giáo án (.docx)</div>", unsafe_allow_html=True)
        
        st.markdown("<hr style='margin-top: 2px; margin-bottom: 4px;'>", unsafe_allow_html=True)
        
        for buoi, tiet in dieu_phoi_tiet:
            row_cols = st.columns([1.1, 2.0, 1.8, 1.8, 3.4])
            key_prefix = f"{thu}_{buoi}_{tiet}"
            
            mon_val = st.session_state.get(f"mon_{key_prefix}", "-- Chọn môn học --")
            khoi_val = st.session_state.get(f"khoi_{key_prefix}", "-- Chọn khối --")
            lop_val = st.session_state.get(f"lop_{key_prefix}", "-- Chọn lớp --")
            is_filled = (mon_val != "-- Chọn môn học --") and (khoi_val != "-- Chọn khối --") and (lop_val != "-- Chọn lớp --")
            
            with row_cols[0]:
                if is_filled:
                    st.markdown(f"<div style='background-color: #27ae60; color: white; padding: 4px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; text-align: center; box-shadow: 0 0 6px rgba(39, 174, 96, 0.8);'>✅ {buoi}<br>{tiet}</div>", unsafe_allow_html=True)
                else:
                    if "Sáng" in buoi:
                        st.markdown(f"<div style='background-color: #3182ce; color: white; padding: 4px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; text-align: center;'>{buoi}<br>{tiet}</div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div style='background-color: #fffdf0; color: red; padding: 4px 6px; border-radius: 4px; border: 1px solid #d69e2e; font-size: 11px; font-weight: bold; text-align: center;'>{buoi}<br>{tiet}</div>", unsafe_allow_html=True)
                
            with row_cols[1]:
                mon = st.selectbox(f"Môn {key_prefix}", danh_sach_mon, label_visibility="collapsed", key=f"mon_{key_prefix}")
                
            with row_cols[2]:
                khoi = st.selectbox(f"Khối {key_prefix}", danh_sach_khoi, label_visibility="collapsed", key=f"khoi_{key_prefix}")
                
            with row_cols[3]:
                if khoi != "-- Chọn khối --":
                    danh_sach_lop_theo_khoi = ["-- Chọn lớp --"] + mapping_lop[khoi]
                else:
                    danh_sach_lop_theo_khoi = ["-- Chọn lớp --"]
                
                lop = st.selectbox(f"Lớp {key_prefix}", danh_sach_lop_theo_khoi, label_visibility="collapsed", key=f"lop_{key_prefix}")
                
            with row_cols[4]:
                file_ga = st.file_uploader(f"File {key_prefix}", type=["docx", "pdf"], label_visibility="collapsed", key=f"file_{key_prefix}")
            
            if is_filled:
                tkb_data[key_prefix] = {
                    "thu": thu, "buoi": buoi, "tiet": tiet, "mon": mon, "khoi": khoi, "lop": lop, "file": file_ga
                }
            
            st.markdown("<div style='margin-bottom: 2px;'></div>", unsafe_allow_html=True)

data_separator = "--------------------------------------------------"
st.markdown("---")

# --- HÀM XỬ LÝ GHÉP FILE WORD (DÙNG DOCXCOMPOSE BẢO TOÀN ẢNH & FONT) ---
def generate_master_docx(tkb_dict, tuan, ngay_bd, h_left, h_right, f_left, f_right):
    master_doc = Document()
    
    # Thiết lập Style chuẩn Times New Roman ngay từ đầu
    style = master_doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)

    for section in master_doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)
        
        section.header_distance = Cm(1.0)
        section.footer_distance = Cm(1.0)
        
        body_width = section.page_width - section.left_margin - section.right_margin
        
        # --- 1. HEADER ---
        header = section.header
        header.paragraphs[0].text = ""
        header_table = header.add_table(rows=1, cols=2, width=body_width)
        header_table.autofit = False
        
        tblPr = header_table._tbl.tblPr
        tblBorders = OxmlElement('w:tblBorders')
        for b_name in ['left', 'top', 'right', 'bottom', 'insideH', 'insideV']:
            b = OxmlElement(f'w:{b_name}')
            b.set(qn('w:val'), 'none')
            tblBorders.append(b)
        tblPr.append(tblBorders)
        
        cell_hl = header_table.cell(0, 0)
        cell_hr = header_table.cell(0, 1)
        cell_hl.width = body_width / 2
        cell_hr.width = body_width / 2
        
        cell_hl.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_hr.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        ph_l = cell_hl.paragraphs[0]
        ph_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
        ph_l.paragraph_format.space_before = Pt(0)
        ph_l.paragraph_format.space_after = Pt(0)
        rh_l = ph_l.add_run(h_left)
        rh_l.font.name = 'Times New Roman'
        rh_l.font.size = Pt(13)
        
        ph_r = cell_hr.paragraphs[0]
        ph_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ph_r.paragraph_format.space_before = Pt(0)
        ph_r.paragraph_format.space_after = Pt(0)
        rh_r = ph_r.add_run(h_right)
        rh_r.font.name = 'Times New Roman'
        rh_r.font.size = Pt(13)
        
        p_line_header = header.add_paragraph()
        p_line_header.paragraph_format.space_before = Pt(0)
        p_line_header.paragraph_format.space_after = Pt(0)
        p_line_header.paragraph_format.line_spacing = Pt(1)
        
        run_line = p_line_header.add_run()
        run_line.font.size = Pt(1)
        
        pPr = p_line_header._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom_border = OxmlElement('w:bottom')
        bottom_border.set(qn('w:val'), 'single')
        bottom_border.set(qn('w:sz'), '12')
        bottom_border.set(qn('w:space'), '1')
        bottom_border.set(qn('w:color'), '000000')
        pBdr.append(bottom_border)
        pPr.append(pBdr)
            
        # --- 2. FOOTER (3 Cột: Trái, Giữa - Số trang, Phải) ---
        footer = section.footer
        footer.paragraphs[0].text = ""
        footer_table = footer.add_table(rows=1, cols=3, width=body_width)
        footer_table.autofit = False
        
        tblPr_f = footer_table._tbl.tblPr
        tblBorders_f = OxmlElement('w:tblBorders')
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'single')
        top.set(qn('w:sz'), '6')
        top.set(qn('w:space'), '4')
        top.set(qn('w:color'), '000000')
        tblBorders_f.append(top)
        
        for b_name in ['left', 'bottom', 'right', 'insideH', 'insideV']:
            b = OxmlElement(f'w:{b_name}')
            b.set(qn('w:val'), 'none')
            tblBorders_f.append(b)
        tblPr_f.append(tblBorders_f)
        
        cell_fl = footer_table.cell(0, 0)
        cell_fm = footer_table.cell(0, 1)
        cell_fr = footer_table.cell(0, 2)
        
        cell_fl.width = body_width * 0.4
        cell_fm.width = body_width * 0.2
        cell_fr.width = body_width * 0.4
        
        cell_fl.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_fm.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_fr.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        pf_l = cell_fl.paragraphs[0]
        pf_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf_l.paragraph_format.space_before = Pt(2)
        pf_l.paragraph_format.space_after = Pt(0)
        rf_l = pf_l.add_run(f_left)
        rf_l.font.name = 'Times New Roman'
        rf_l.font.size = Pt(13)
        
        pf_m = cell_fm.paragraphs[0]
        pf_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf_m.paragraph_format.space_before = Pt(2)
        pf_m.paragraph_format.space_after = Pt(0)
        
        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        pf_m._p.append(fldSimple)
        
        for run in pf_m.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(13)

        pf_r = cell_fr.paragraphs[0]
        pf_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pf_r.paragraph_format.space_before = Pt(2)
        pf_r.paragraph_format.space_after = Pt(0)
        rf_r = pf_r.add_run(f_right)
        rf_r.font.name = 'Times New Roman'
        rf_r.font.size = Pt(13)

    p_title = master_doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(f"TỔNG HỢP GIÁO ÁN - TUẦN {tuan}")
    r_title.font.name = 'Times New Roman'
    r_title.font.size = Pt(16)
    r_title.bold = True
    
    master_doc.add_page_break()
    last_written_day = None

    composer = Composer(master_doc)

    for thu in cac_thu:
        ngay_cu_the = ngay_bd + datetime.timedelta(days=ngay_offset[thu])
        tiet_trong_ngay = [(b, t) for b, t in dieu_phoi_tiet if f"{thu}_{b}_{t}" in tkb_dict]
        
        if tiet_trong_ngay:
            if last_written_day is not None:
                master_doc.add_page_break()
                
            p_date = master_doc.add_paragraph()
            p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_date = p_date.add_run(f"{thu}, ngày {ngay_cu_the.day} tháng {ngay_cu_the.month} năm {ngay_cu_the.year}")
            r_date.font.name = 'Times New Roman'
            r_date.font.size = Pt(14)
            r_date.italic = True
            
            last_written_day = thu
            
            total_classes_today = len(tiet_trong_ngay)
            current_class_today_idx = 0

            has_morning = any(b == "Buổi Sáng" for b, t in tiet_trong_ngay)
            has_afternoon = any(b == "Buổi Chiều" for b, t in tiet_trong_ngay)
            has_both = has_morning and has_afternoon
            inserted_buoi_chieu_title = False

            for buoi, tiet in tiet_trong_ngay:
                key = f"{thu}_{buoi}_{tiet}"
                item = tkb_dict[key]
                current_class_today_idx += 1
                
                if has_both and buoi == "Buổi Chiều" and not inserted_buoi_chieu_title:
                    p_bc = master_doc.add_paragraph()
                    p_bc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_bc = p_bc.add_run("BUỔI CHIỀU")
                    r_bc.font.name = 'Times New Roman'
                    r_bc.font.size = Pt(14)
                    r_bc.bold = True
                    inserted_buoi_chieu_title = True

                p_lop = master_doc.add_paragraph()
                p_lop.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_lop = p_lop.add_run(f"LỚP: {item['lop']}")
                r_lop.font.name = 'Times New Roman'
                r_lop.font.size = Pt(14)
                r_lop.bold = True
                
                p_mon = master_doc.add_paragraph()
                p_mon.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_mon = p_mon.add_run(f"MÔN {item['mon'].upper()}")
                r_mon.font.name = 'Times New Roman'
                r_mon.font.size = Pt(14)
                r_mon.bold = True
                
                uploaded_file = item["file"]
                if uploaded_file is not None and uploaded_file.name.endswith(".docx"):
                    try:
                        # Sử dụng docxcompose để gộp file giáo án (giữ nguyên hình ảnh, định dạng, bảng biểu)
                        sub_doc = Document(uploaded_file)
                        composer.append(sub_doc)
                    except Exception as e:
                        p_err = master_doc.add_paragraph(f"[Lỗi đọc file: {str(e)}]")
                        p_err.runs[0].font.name = 'Times New Roman'
                else:
                    p_none = master_doc.add_paragraph("[Chưa đính kèm file Word giáo án cho tiết này]")
                    p_none.runs[0].font.name = 'Times New Roman'
                    p_none.runs[0].italic = True
                
                if current_class_today_idx < total_classes_today:
                    p_line = master_doc.add_paragraph()
                    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_line = p_line.add_run(data_separator)
                    r_line.font.name = 'Times New Roman'
                    r_line.font.size = Pt(12)
                    r_line.bold = True

    # --- Ép toàn bộ chữ sang Times New Roman ---
    for p in master_doc.paragraphs:
        for run in p.runs:
            run.font.name = 'Times New Roman'
            
    for table in master_doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.name = 'Times New Roman'

    file_stream = io.BytesIO()
    master_doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- PHẦN 3: NÚT XUẤT VÀ TẢI FILE ---
action_container = st.container()
with action_container:
    col_a, col_b, col_c = st.columns([1, 2, 1])
    with col_b:
        if st.button("🚀 XỬ LÝ & TẠO FILE WORD TỔNG HỢP", type="primary", use_container_width=True):
            has_error = False
            error_messages = []
            
            for thu in cac_thu:
                for buoi, tiet in dieu_phoi_tiet:
                    key_prefix = f"{thu}_{buoi}_{tiet}"
                    mon_val = st.session_state.get(f"mon_{key_prefix}", "-- Chọn môn học --")
                    lop_val = st.session_state.get(f"lop_{key_prefix}", "-- Chọn lớp --")
                    
                    if mon_val != "-- Chọn môn học --" and lop_val == "-- Chọn lớp --":
                        has_error = True
                        error_messages.append(f"📌 **{thu} ({buoi} - {tiet})**: Đã chọn môn **'{mon_val}'** nhưng chưa chọn lớp!")
                    elif lop_val != "-- Chọn lớp --" and mon_val == "-- Chọn môn học --":
                        has_error = True
                        error_messages.append(f"📌 **{thu} ({buoi} - {tiet})**: Đã chọn lớp **'{lop_val}'** nhưng chưa chọn môn học!")

            if has_error:
                st.error("⚠️ **Phát hiện lỗi nhập liệu thời khóa biểu:** Vui lòng kiểm tra lại các tiết học sau:")
                for err in error_messages:
                    st.markdown(err)
            elif not tkb_data:
                st.warning("⚠️ Vui lòng chọn ít nhất một tiết học đầy đủ Môn học và Lớp!")
            else:
                with st.spinner("⏳ Đang tổng hợp các file giáo án, vui lòng chờ trong giây lát..."):
                    docx_stream = generate_master_docx(tkb_data, tuan_so, ngay_dau_tuan, header_left, header_right, footer_left, footer_right)
                    st.session_state['generated_docx'] = docx_stream.getvalue()
                    st.success(f"🎉 Đã tổng hợp thành công giáo án cho Tuần {tuan_so}!")

        if st.session_state['generated_docx'] is not None:
            st.markdown("<div style='margin-top: 5px;'></div>", unsafe_allow_html=True)
            st.download_button(
                label="📥 BẤM VÀO ĐÂY ĐỂ TẢI FILE WORD VỀ MÁY",
                data=st.session_state['generated_docx'],
                file_name=f"Tong_Hop_Giao_An_Tuan_{tuan_so}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="secondary",
                use_container_width=True
            )
