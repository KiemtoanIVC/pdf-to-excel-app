import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
from io import BytesIO
from datetime import datetime
from PIL import Image

# Load logo
logo = Image.open('assets/logo.png')

# Cấu hình trang với logo
st.set_page_config(
    page_title="IVC - PDF to Excel Converter",
    page_icon=logo,
    layout="wide"
)

# Tiêu đề ứng dụng với logo
col1, col2 = st.columns([1, 4])
with col1:
    st.image(logo, width=100)
with col2:
    st.title("Chuyển đổi C12 -TB BHXH - PDF sang Excel")
st.markdown("---")

# Các hàm xử lý từ file gốc
def extract_month_from_text(text):
    match = re.search(r'Tháng (\d+) năm (\d+)', text)
    if match:
        return f"{match.group(1)}/{match.group(2)}"
    return ""

def create_filter_value(stt, cur_level1):
    if not stt or not str(stt).strip():
        return "", cur_level1
    
    stt = str(stt).strip()
    main_sections = ['A', 'B', 'C', 'D', 'Đ']
    
    if stt in main_sections:
        cur_level1 = stt
        return stt, cur_level1
    
    if cur_level1:
        return f"{cur_level1}_{stt}", cur_level1
        
    return stt, cur_level1

def process_pdf_files(uploaded_files):
    all_companies = []
    all_tables = []
    
    # Hiển thị progress bar
    progress_bar = st.progress(0)
    progress_text = st.empty()
    
    for i, uploaded_file in enumerate(uploaded_files):
        try:
            # Đọc file PDF
            bytes_data = uploaded_file.read()
            
            with pdfplumber.open(BytesIO(bytes_data)) as pdf:
                # Xử lý thông tin công ty
                page = pdf.pages[0]
                text = page.extract_text()
                month = extract_month_from_text(text)
                
                company_data = {
                    'Tháng': [month],
                    'Tên công ty': [''],
                    'Mã đơn vị': [''],
                    'Điện thoại': ['']
                }
                
                lines = text.split('\n')
                for line in lines:
                    if 'Kính gửi:' in line:
                        company_data['Tên công ty'][0] = line.replace('Kính gửi:', '').strip()
                    elif 'Mã đơn vị:' in line:
                        company_data['Mã đơn vị'][0] = line.split('Mã đơn vị:')[1].split('Điện')[0].strip()
                    elif 'Điện thoại:' in line:
                        company_data['Điện thoại'][0] = line.split('Điện thoại:')[1].strip()
                
                company_df = pd.DataFrame(company_data)
                all_companies.append(company_df)
                
                # Xử lý bảng dữ liệu
                table_data = []
                cur_level1 = None
                
                for page_num in range(len(pdf.pages)):
                    page = pdf.pages[page_num]
                    tables = page.extract_tables()
                    
                    if tables:
                        main_table = tables[0]
                        for row in main_table:
                            if row and len(row) >= 8:
                                if any(cell and cell.strip() and '===' not in str(cell) for cell in row):
                                    filter_value, cur_level1 = create_filter_value(row[0], cur_level1)
                                    formatted_row = {
                                        'Tháng': month,
                                        'Filter': filter_value,
                                        'STT': row[0],
                                        'Nội dung': row[1],
                                        'BHXH_OD_TS': row[2],
                                        'BHXH_HTTT': row[3],
                                        'BHYT': row[4],
                                        'BHTN': row[5],
                                        'BHTNLD_BNN': row[6],
                                        'Cộng': row[7]
                                    }
                                    table_data.append(formatted_row)
                
                if table_data:
                    columns_order = ['Tháng', 'Filter', 'STT', 'Nội dung', 'BHXH_OD_TS', 
                                   'BHXH_HTTT', 'BHYT', 'BHTN', 'BHTNLD_BNN', 'Cộng']
                    table_df = pd.DataFrame(table_data, columns=columns_order)
                    all_tables.append(table_df)
            
            # Cập nhật progress
            progress = (i + 1) / len(uploaded_files)
            progress_bar.progress(progress)
            progress_text.text(f'Đang xử lý: {uploaded_file.name}')
            
        except Exception as e:
            st.error(f"Lỗi khi xử lý file {uploaded_file.name}: {str(e)}")
    
    # Xóa progress bar và text
    progress_bar.empty()
    progress_text.empty()
    
    if all_companies and all_tables:
        # Gộp và sắp xếp dữ liệu
        combined_companies = pd.concat(all_companies, ignore_index=True)
        combined_tables = pd.concat(all_tables, ignore_index=True)
        
        # Sắp xếp theo tháng
        combined_companies['Sort_Date'] = pd.to_datetime(combined_companies['Tháng'].str.split('/').str[0] + '/01/' + 
                                                       combined_companies['Tháng'].str.split('/').str[1])
        combined_tables['Sort_Date'] = pd.to_datetime(combined_tables['Tháng'].str.split('/').str[0] + '/01/' + 
                                                    combined_tables['Tháng'].str.split('/').str[1])
        
        combined_companies = combined_companies.sort_values('Sort_Date').drop('Sort_Date', axis=1)
        combined_tables = combined_tables.sort_values('Sort_Date').drop('Sort_Date', axis=1)
        
        # Tạo các sheet phụ
        ps_trong_ky = combined_tables[combined_tables['Filter'] == 'B'].copy()
        sl_lao_dong = combined_tables[combined_tables['Filter'] == 'Đ_1'].copy()
        
        return combined_companies, combined_tables, ps_trong_ky, sl_lao_dong
    
    return None, None, None, None

# Giao diện upload files
uploaded_files = st.file_uploader("Chọn các file PDF cần xử lý", 
                                type=['pdf'], 
                                accept_multiple_files=True)

if uploaded_files:
    if st.button("Xử lý"):
        with st.spinner('Đang xử lý...'):
            # Xử lý files
            company_df, table_df, ps_trong_ky, sl_lao_dong = process_pdf_files(uploaded_files)
            
            if company_df is not None:
                # Tạo Excel file trong memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    company_df.to_excel(writer, sheet_name='Thong tin cong ty', index=False)
                    table_df.to_excel(writer, sheet_name='Bang du lieu', index=False)
                    ps_trong_ky.to_excel(writer, sheet_name='PS trong ky', index=False)
                    sl_lao_dong.to_excel(writer, sheet_name='SL lao dong', index=False)
                
                # Tạo tên file với timestamp
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_filename = f"ketqua_tong_hop_{timestamp}.xlsx"
                
                # Download button
                st.success("✅ Xử lý thành công!")
                st.download_button(
                    label="📥 Tải file Excel",
                    data=output.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Hiển thị preview dữ liệu
                with st.expander("Xem trước dữ liệu"):
                    st.subheader("Thông tin công ty")
                    st.dataframe(company_df)
                    
                    st.subheader("Bảng dữ liệu")
                    st.dataframe(table_df)
                    
                    st.subheader("PS trong kỳ")
                    st.dataframe(ps_trong_ky)
                    
                    st.subheader("SL lao động")
                    st.dataframe(sl_lao_dong)
            else:
                st.error("❌ Không có dữ liệu để xử lý")

# Thêm thông tin footer
st.markdown("---")
st.markdown("Made with ❤️ by IVC Audit PDH") 