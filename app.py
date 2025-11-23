import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
from fuzzywuzzy import fuzz, process 
import xlsxwriter # Cần thiết cho việc tạo và tải file Excel

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Tiện Ích Chuẩn Hóa (FAST)", layout="centered")
st.title("🚀 TIỆN ÍCH CHUẨN HÓA DỮ LIỆU CHÍNH XÁC CAO (Hoàn Chỉnh)")

# --- HÀM TIỆN ÍCH CHUẨN HÓA ---
@st.cache_data
def xoa_dau_tieng_viet(text):
    """Xóa dấu tiếng Việt, chuyển về chữ thường và loại bỏ khoảng trắng dư thừa."""
    if not isinstance(text, str): 
        return str(text).lower().strip()
    text = unicodedata.normalize('NFD', text)
    text = re.sub(r'[\u0300-\u036f]', '', text)
    text = text.lower().strip()
    text = re.sub(r'\s+', ' ', text)
    return text

# --- HÀM 1A: ĐỌC FILE (Đã Tối Ưu Cache) ---
@st.cache_data(show_spinner="Đang tải và đọc file lớn (Chỉ chạy lần đầu)...")
def doc_file_data(uploaded_file):
    """Hàm cache chuyên đọc file, chỉ chạy lại khi file thay đổi."""
    try:
        engine = 'pyxlsb' if uploaded_file.name.endswith('.xlsb') else 'openpyxl'
        # Dùng io.BytesIO để đảm bảo cache hoạt động tốt với file object
        df = pd.read_excel(BytesIO(uploaded_file.getvalue()), engine=engine)
        return df
    except Exception as e:
        st.error(f"❌ Lỗi đọc file: {e}")
        return None

# --- HÀM HỖ TRỢ EXCEL (Tạo file tải về) ---
@st.cache_data
def tao_file_excel(df_input):
    """Tạo file Excel từ DataFrame để tải xuống."""
    output = BytesIO()
    # Dùng xlsxwriter cho tốc độ và khả năng tương thích
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_input.to_excel(writer, index=False, sheet_name='DanhSachTrungLap')
    writer.close()
    return output

# --- BƯỚC 1: NẠP VÀ LỰA CHỌN DỮ LIỆU ---
def hien_thi_nhap_lieu():
    uploaded_file = st.file_uploader("📂 Tải lên file Excel (.xlsx, .xlsb)", type=['xlsx', 'xlsb'])
    df = None
    selected_col = None

    if uploaded_file is not None:
        st.success(f"✅ Đã tải lên file: {uploaded_file.name}")
        
        df = doc_file_data(uploaded_file)
        
        if df is not None:
            cols = df.columns.tolist()
            default_index = cols.index('hoTen') if 'hoTen' in cols else 0

            selected_col = st.selectbox(
                "📋 Chọn cột dữ liệu cần Chuẩn hóa (Ví dụ: hoTen, diaChi):", 
                options=cols,
                index=default_index
            )
            
    return df, selected_col

# --- BƯỚC 2: CHUẨN HÓA CƠ BẢN (Tối ưu Vector hóa) ---
@st.cache_data(show_spinner="Đang chuẩn hóa toàn bộ dữ liệu (Vector hóa, chỉ chạy 1 lần)...")
def xu_ly_chuan_hoa_co_ban(df, ten_cot_goc):
    """Thực hiện chuẩn hóa nhanh bằng phương pháp vector hóa Pandas."""
    if df is None or ten_cot_goc is None or ten_cot_goc not in df.columns:
        return df, None

    # Tối ưu hóa: Xử lý chuỗi bằng phương pháp Vector hóa của Pandas (.str)
    # Bước 1: Chuẩn hóa Unicode (Xóa dấu)
    df[ten_cot_goc] = df[ten_cot_goc].astype(str).str.normalize('NFD').str.replace(r'[\u0300-\u036f]', '', regex=True).fillna('')
    
    # Bước 2: Chuyển về chữ thường và loại bỏ khoảng trắng dư thừa
    ten_cot_moi = f"{ten_cot_goc}_khongdau"
    df[ten_cot_moi] = df[ten_cot_goc].str.lower().str.replace(r'\s+', ' ', regex=True).str.strip()
    
    st.success(f"✅ Đã tạo cột chuẩn hóa: **{ten_cot_moi}**. Tốc độ được cải thiện nhiều!")
    return df, ten_cot_moi

# --- BƯỚC 3: TÌM KIẾM GẦN ĐÚNG (FUZZY MATCHING) ---
def tim_kiem_gan_dung(df_input, cot_cleaned):
    """Thực hiện tìm kiếm gần đúng dựa trên FuzzyWuzzy."""
    st.subheader("🔎 Tìm Kiếm Gần Đúng (Fuzzy Search)")
    
    c1, c2 = st.columns([3, 1])
    search_term = c1.text_input("Nhập Tên/Từ khóa tìm kiếm gần đúng:", placeholder="vd: nguyen thi hoa")
    min_score = c2.slider("Ngưỡng khớp:", min_value=50, max_value=100, value=85, step=1)
    
    if search_term and df_input is not None and cot_cleaned in df_input.columns:
        term_cleaned = xoa_dau_tieng_viet(search_term)
        choices = df_input[cot_cleaned].unique().tolist()
        
        with st.spinner(f"Đang tìm kiếm gần đúng cho '{search_term}'..."):
            results = process.extract(term_cleaned, choices, scorer=fuzz.token_sort_ratio)
            filtered_results = [r for r in results if r[1] >= min_score]

        if filtered_results:
            matched_values = [r[0] for r in filtered_results]
            score_map = {r[0]: r[1] for r in filtered_results}
            
            df_ket_qua = df_input[df_input[cot_cleaned].isin(matched_values)].copy()
            
            df_ket_qua['Diem_Khop'] = df_ket_qua[cot_cleaned].map(score_map)
            df_ket_qua = df_ket_qua.sort_values(by='Diem_Khop', ascending=False)
            
            st.success(f"Tìm thấy **{len(df_ket_qua)}** hồ sơ có điểm khớp >= {min_score}!")
            hien_thi_cols = [col for col in df_input.columns if col not in [cot_cleaned]]
            st.dataframe(df_ket_qua[['Diem_Khop'] + hien_thi_cols].head(50), use_container_width=True)

        else:
            st.warning(f"Không tìm thấy kết quả nào khớp với '{search_term}' ở mức điểm {min_score} trở lên.")
    
    return

# --- BƯỚC 4: KIỂM TRA TRÙNG LẶP (Chỉ trả về Data) ---
@st.cache_data(show_spinner="Đang kiểm tra trùng lặp trên tổ hợp...")
def kiem_tra_trung_lap(df, list_cot_kiem_tra):
    if not list_cot_kiem_tra:
        return pd.DataFrame() # Trả về DataFrame rỗng nếu không có cột nào được chọn
        
    # Dùng .duplicated(subset=list, keep=False) để đánh dấu TẤT CẢ các bản ghi trùng lặp
    is_duplicate = df.duplicated(subset=list_cot_kiem_tra, keep=False)
    
    # Lọc ra các bản ghi bị trùng lặp
    df_trung = df[is_duplicate].sort_values(by=list_cot_kiem_tra)
    
    return df_trung # CHỈ TRẢ VỀ DATAFRAME

# --- HÀM GIAO DIỆN KIỂM TRA TRÙNG LẶP NÂNG CAO (Xử lý UI và Download) ---
def hien_thi_kiem_tra_trung_lap_nang_cao(df):
    st.markdown("---")
    st.subheader("🛠️ KIỂM TRA TRÙNG LẶP NÂNG CAO (Nhiều Cột)")

    all_cols = df.columns.tolist() 
    default_selection = [c for c in ['hoTen_khongdau', 'ngaySinh', 'soCmnd'] if c in all_cols]
    
    list_cot_kiem_tra = st.multiselect(
        "Chọn các cột để tạo tổ hợp trùng lặp (Ví dụ: Tên chuẩn hóa + Ngày sinh + Số CMND):",
        options=all_cols,
        default=default_selection
    )
    
    if st.button("🔍 PHÂN TÍCH TRÙNG LẶP"):
        if list_cot_kiem_tra:
            # 1. GỌI HÀM CACHE ĐỂ LẤY DỮ LIỆU
            df_trung = kiem_tra_trung_lap(df, list_cot_kiem_tra)
            
            ten_to_hop = " + ".join(list_cot_kiem_tra)
            
            # 2. HIỂN THỊ KẾT QUẢ VÀ WIDGET (NGOÀI CACHE)
            if not df_trung.empty:
                st.error(f"🔴 Tìm thấy **{len(df_trung)}** bản ghi KHẢ NĂNG TRÙNG LẶP dựa trên tổ hợp **{ten_to_hop}**!")
                
                excel_data = tao_file_excel(df_trung) 
                st.download_button(
                    label="📥 Tải danh sách Trùng lặp (Excel)",
                    data=excel_data.getvalue(),
                    file_name=f"trung_lap_nang_cao_{ten_to_hop}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.dataframe(df_trung, use_container_width=True, height=500)
            else:
                st.success("✅ Không tìm thấy bản ghi trùng lặp nào dựa trên tổ hợp đã chọn.")
        else:
            st.warning("Vui lòng chọn ít nhất một cột để chạy phân tích trùng lặp.")

# --- HÀM MAIN CHÍNH (Đã cập nhật) ---
def main():
    df_data, cot_chon = hien_thi_nhap_lieu()
    st.markdown("---")

    if df_data is not None and cot_chon:
        st.info(f"Tổng cộng **{len(df_data)}** hồ sơ. Đang xử lý cột: **{cot_chon}**")
        
        # Bước Chuẩn hóa
        df_cleaned, cot_cleaned = xu_ly_chuan_hoa_co_ban(df_data.copy(), cot_chon) 

        if df_cleaned is not None and cot_cleaned:
            st.subheader("Xem trước Dữ liệu đã Chuẩn hóa")
            st.dataframe(df_cleaned[[cot_chon, cot_cleaned]].head(20), use_container_width=True)
            st.markdown("---")
            
            # Bước Tìm kiếm Gần đúng
            tim_kiem_gan_dung(df_cleaned, cot_cleaned)
            
            # Bước Kiểm tra Trùng lặp Nâng cao
            hien_thi_kiem_tra_trung_lap_nang_cao(df_cleaned.copy())

# --- CHẠY CHƯƠNG TRÌNH ---
if __name__ == "__main__":
    main()