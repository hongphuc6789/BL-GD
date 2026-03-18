import streamlit as st
import pdfplumber
import re
import os
import tempfile
import subprocess
import sys
from docxtpl import DocxTemplate

# Đường dẫn mặc định tới file Template cố định
TEMPLATE_PATH = "TEMPLATE.docx"

# Hàm chuyển đổi DOCX sang PDF linh hoạt theo hệ điều hành (Local Windows vs Streamlit Cloud Linux)
def convert_docx_to_pdf(docx_path, out_dir):
    # Nếu chạy trên Windows (Local)
    if sys.platform == "win32":
        import pythoncom
        from docx2pdf import convert
        pythoncom.CoInitialize() 
        pdf_path = docx_path.replace(".docx", ".pdf")
        convert(docx_path, pdf_path)
        return pdf_path
    
    # Nếu chạy trên Linux (Streamlit Community Cloud)
    else:
        # Gọi LibreOffice chạy ngầm để xuất PDF
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", out_dir, docx_path
        ], check=True)
        return docx_path.replace(".docx", ".pdf")

def extract_crew_data(pdf_file, target_flight):
    crew_list = []
    route_crews = []
    extracting = False
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Làm sạch dữ liệu, thay thế xuống dòng bằng khoảng trắng
                    clean_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                    
                    # Bỏ qua các hàng trống hoàn toàn
                    if not any(clean_row):
                        continue
                    
                    flights_col = clean_row[0] if len(clean_row) > 0 else ""
                    
                    # BƯỚC 1: TÌM BLOCK CHUYẾN BAY DỰA VÀO CỘT FLIGHTS
                    if "BL" in flights_col and "Flights" not in flights_col:
                        if target_flight in flights_col:
                            extracting = True
                            
                            # Quét tìm FE/OBS ở cột Route
                            route_col = clean_row[1] if len(clean_row) > 1 else ""
                            parts = route_col.split('/') 
                            for part in parts:
                                if 'FE' in part or 'OBS' in part:
                                    if target_flight in part:
                                        match = re.search(r'(FE|OBS).*', part)
                                        if match:
                                            r_crew = match.group(0).strip()
                                            if r_crew not in route_crews:
                                                route_crews.append(r_crew)
                                    elif not re.search(r'BL\d+', part):
                                        match = re.search(r'(FE|OBS).*', part)
                                        if match:
                                            r_crew = match.group(0).strip()
                                            if r_crew not in route_crews:
                                                route_crews.append(r_crew)
                        else:
                            extracting = False # Đã sang chuyến bay khác -> Tắt cờ
                    
                    # BƯỚC 2: GOM CREW KHI CỜ ĐANG BẬT
                    if extracting:
                        if len(clean_row) >= 2:
                            rank = clean_row[-2].strip()
                            member = clean_row[-1].strip()
                            
                            # Loại bỏ các dòng trống hoặc dòng Header (Tiêu đề bảng)
                            if not rank or not member or rank.lower() == 'rank' or member.lower() == 'crew member':
                                continue
                                
                            # Loại bỏ ngay lập tức nếu cột "Rank" chứa chữ số hoặc dấu gạch ngang (Đây là dữ liệu ngày giờ Duty bị vướng vào)
                            if re.search(r'\d', rank) or '-' in rank:
                                continue
                                
                            crew_str = f"{rank} {member}"
                            if crew_str not in crew_list: 
                                crew_list.append(crew_str)

    route_info = "\n".join(route_crews)
    return "\n".join(crew_list), route_info

# --- GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="Crew GD Generator", layout="centered", page_icon="✈️")
st.title("✈️ General Declaration Generator")

# Kiểm tra file template cứng
if not os.path.exists(TEMPLATE_PATH):
    st.error(f"❌ Lỗi hệ thống: Không tìm thấy file gốc '{TEMPLATE_PATH}'. Vui lòng kiểm tra lại thư mục chứa code.")
    st.stop()

st.header("1. Upload Crew List")
pdf_file = st.file_uploader("Upload file PDF Crew List", type=["pdf"])

st.header("2. Nhập thông tin chuyến bay")
with st.form("gd_form"):
    col1, col2 = st.columns(2)
    with col1:
        flight_no = st.text_input("Số hiệu chuyến bay (vd: BL6080)")
        reg_no = st.text_input("Đăng ký tàu bay (vd: 363 cho VN-A363)")
    with col2:
        arr_port = st.text_input("Arrival Port (vd: HAN)")
        flight_date = st.text_input("Ngày bay (vd: 17-MAR-2026)")
        
    submit_btn = st.form_submit_button("Create GD", type="primary")

if submit_btn:
    if not pdf_file:
        st.error("⚠️ Vui lòng upload file PDF Crew List.")
    elif not flight_no:
        st.error("⚠️ Vui lòng nhập số hiệu chuyến bay.")
    else:
        with st.spinner("Đang xử lý dữ liệu và tạo PDF (Có thể mất vài giây)..."):
            crew_str, route_info = extract_crew_data(pdf_file, flight_no)
            
            if not crew_str:
                st.warning(f"Không tìm thấy dữ liệu tổ bay cho chuyến {flight_no}. Vui lòng kiểm tra lại số hiệu.")
            else:
                # Render dữ liệu vào Word Template
                doc = DocxTemplate(TEMPLATE_PATH)
                context = {
                    "Fltn": flight_no.replace("BL", ""), 
                    "REG": reg_no,
                    "arr": arr_port,
                    "date": flight_date,
                    "rank": crew_str,      
                    "route": route_info    
                }
                doc.render(context)
                
                # Lưu file tạm thời
                temp_dir = tempfile.mkdtemp()
                docx_path = os.path.join(temp_dir, f"GD_{flight_no}.docx")
                doc.save(docx_path)
                
                # Chuyển đổi DOCX sang PDF
                pdf_converted = False
                pdf_path = ""
                try:
                    pdf_path = convert_docx_to_pdf(docx_path, temp_dir)
                    pdf_converted = True
                except Exception as e:
                    st.error(f"Tính năng xuất PDF gặp sự cố: {e}")
                
                st.success(f"Tạo General Declaration cho chuyến {flight_no} thành công! 🎉")
                
                st.header("3. Download Kết Quả")
                col3, col4 = st.columns(2)
                
                with open(docx_path, "rb") as d_file:
                    col3.download_button(
                        label="📄 Download DOCX",
                        data=d_file,
                        file_name=f"GD_{flight_no}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                
                if pdf_converted and os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as p_file:
                        col4.download_button(
                            label="📕 Download PDF",
                            data=p_file,
                            file_name=f"GD_{flight_no}.pdf",
                            mime="application/pdf"
                        )
