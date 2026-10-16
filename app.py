import base64
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import io
import zipfile
import os
import pathlib

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Rà soát & Kiểm định XML BHYT - Bệnh Viện Quân Y 5",
    page_icon="🏥",
    layout="wide"
)

# --- TUỲ CHỈNH GIAO DIỆN CHUYÊN NGHIỆP (CSS) ---
st.markdown("""
    <style>
        /* Màu chủ đạo xanh quân y & bố cục */
        .main {
            background-color: #f8f9fa;
        }
        .stSidebar {
            background-color: #f1f3f5;
            border-right: 1px solid #dee2e6;
        }
        /* Header tiêu đề chính */
        .hospital-header {
            background: linear-gradient(135deg, #1b4d3e 0%, #2c6b56 100%);
            color: white;
            padding: 25px 30px;
            border-radius: 10px;
            margin-bottom: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .hospital-header h1 {
            color: white !important;
            font-size: 1.8rem;
            margin: 0;
            font-weight: 700;
            letter-spacing: 0.5px;
        }
        .hospital-header p {
            color: #e2e8f0;
            font-size: 1rem;
            margin-top: 8px;
            margin-bottom: 0;
        }
        /* Card thống kê */
        .metric-card {
            background: white;
            padding: 15px 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border-left: 4px solid #1b4d3e;
        }
        /* Nút bấm tinh chỉnh */
        .stButton>button {
            background-color: #1b4d3e;
            color: white;
            border-radius: 6px;
            font-weight: 600;
            border: none;
            padding: 0.5rem 1rem;
            transition: all 0.3s ease;
        }
        .stButton>button:hover {
            background-color: #2c6b56;
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
    </style>
""", unsafe_allow_html=True)

# --- THƯ MỤC LƯU TRỮ DANH MỤC CỐ ĐỊNH TRÊN Ổ CỨNG ---
CATEGORY_DIR = "saved_categories"
os.makedirs(CATEGORY_DIR, exist_ok=True)

def load_categories_from_disk():
    cats = {
        "Danh mục Dịch vụ kỹ thuật (XML3)": pd.DataFrame(),
        "Danh mục Thuốc (XML2)": pd.DataFrame(),
        "Danh mục Vật tư y tế": pd.DataFrame()
    }
    for file_path in pathlib.Path(CATEGORY_DIR).glob("*.*"):
        cat_name = file_path.stem
        try:
            if file_path.suffix == '.csv':
                cats[cat_name] = pd.read_csv(file_path)
            elif file_path.suffix in ['.xlsx', '.xls']:
                cats[cat_name] = pd.read_excel(file_path)
        except Exception:
            pass
    return cats

if 'categories' not in st.session_state:
    st.session_state['categories'] = load_categories_from_disk()

def save_category_to_disk(cat_name, df):
    safe_name = "".join(c for c in cat_name if c.isalnum() or c in (' ', '_', '-')).strip()
    file_path = os.path.join(CATEGORY_DIR, f"{safe_name}.csv")
    df.to_csv(file_path, index=False, encoding='utf-8-sig')

def delete_category_from_disk(cat_name):
    safe_name = "".join(c for c in cat_name if c.isalnum() or c in (' ', '_', '-')).strip()
    file_path = os.path.join(CATEGORY_DIR, f"{safe_name}.csv")
    if os.path.exists(file_path):
        os.remove(file_path)

def decode_base64_xml(base64_str):
    try:
        decoded_bytes = base64.b64decode(base64_str.strip())
        try:
            decoded_text = decoded_bytes.decode('utf-8')
        except UnicodeDecodeError:
            decoded_text = decoded_bytes.decode('windows-1258', errors='ignore')
        return ET.fromstring(decoded_text)
    except Exception:
        return None

def deep_parse_element(element, parent_tag=""):
    row_data = {}
    def recursive_traverse(node, prefix=""):
        has_sub = False
        for child in node:
            child_tag = child.tag.split('}')[-1]
            tag_name = f"{prefix}_{child_tag}" if prefix else child_tag
            if len(list(child)) > 0:
                recursive_traverse(child, tag_name)
                has_sub = True
            else:
                val = child.text.strip() if child.text else ""
                if val:
                    row_data[tag_name] = val
        return has_sub
    recursive_traverse(element, parent_tag)
    return row_data

def extract_xml_records(root_element):
    rows = []
    candidates = []
    for elem in root_element.iter():
        if len(list(elem)) > 0 and all(len(list(sub)) == 0 for sub in elem):
            candidates.append(elem)
            
    if candidates:
        for elem in candidates:
            row_data = {child.tag.split('}')[-1]: (child.text.strip() if child.text else "") for child in elem if len(list(child)) == 0}
            if any(row_data.values()):
                rows.append(row_data)
    
    if not rows:
        for child in root_element:
            res = deep_parse_element(child)
            if res:
                rows.append(res)
    return rows

def parse_xml1_to_dict(element):
    result = {}
    for child in element.iter():
        if len(list(child)) == 0:
            if child.text and child.text.strip():
                tag_name = child.tag.split('}')[-1]
                result[tag_name] = child.text.strip()
    return result

def parse_xml_safely(uploaded_file_or_path):
    content_bytes = None
    if hasattr(uploaded_file_or_path, "read"):
        uploaded_file_or_path.seek(0)
        content_bytes = uploaded_file_or_path.read()
    else:
        with open(uploaded_file_or_path, "rb") as f:
            content_bytes = f.read()
            
    for enc in ['utf-8', 'windows-1258', 'latin-1']:
        try:
            text_content = content_bytes.decode(enc)
            return ET.fromstring(text_content), True, ""
        except Exception:
            continue
            
    return None, False, "Không thể giải mã định dạng file (lỗi mã hóa font)."

def process_xml_package(uploaded_file_or_path):
    sub_files_data = {}
    is_valid_format = True
    error_message = ""
    
    root, success, err_msg = parse_xml_safely(uploaded_file_or_path)
    if not success:
        return sub_files_data, False, err_msg
        
    try:
        hoso_list = root.findall('.//FILEHOSO')
        if not hoso_list:
            root_tag = root.tag.split('}')[-1].upper()
            if not root_tag.startswith('XML'):
                is_valid_format = False
                error_message = "Không tìm thấy thẻ <FILEHOSO> hoặc cấu trúc gốc không khớp chuẩn XML BHYT."
            sub_files_data[root_tag] = extract_xml_records(root)
            return sub_files_data, is_valid_format, error_message

        for fh in hoso_list:
            loai_hoso = fh.find('LOAIHOSO')
            noi_dung = fh.find('NOIDUNGFILE')
            
            if loai_hoso is not None and noi_dung is not None:
                l_type = loai_hoso.text.strip().upper()
                if noi_dung.text:
                    sub_root = decode_base64_xml(noi_dung.text)
                    if sub_root is not None:
                        if 'XML1' in l_type:
                            sub_files_data[l_type] = parse_xml1_to_dict(sub_root)
                        else:
                            extracted_rows = extract_xml_records(sub_root)
                            clean_rows = []
                            for r in extracted_rows:
                                clean_r = {str(k): str(v) for k, v in r.items()}
                                if clean_r and clean_r not in clean_rows:
                                    clean_rows.append(clean_r)
                            sub_files_data[l_type] = clean_rows
                    else:
                        is_valid_format = False
                        error_message = f"Không thể giải mã Base64 nội dung của loại hồ sơ: {l_type}"
            else:
                is_valid_format = False
                error_message = "Thẻ <FILEHOSO> thiếu <LOAIHOSO> hoặc <NOIDUNGFILE>."
    except Exception as e:
        is_valid_format = False
        error_message = f"Lỗi cấu trúc cú pháp XML (Syntax Error): {str(e)}"
        
    return sub_files_data, is_valid_format, error_message

def audit_cls_machine_code(all_tables):
    has_cls_missing_machine = False
    details = []
    
    for t_name, t_data in all_tables.items():
        if any(keyword in t_name for keyword in ['XML3', 'DVKTV', 'CANLAMSANG', 'CLS']):
            if isinstance(t_data, list):
                for idx, row in enumerate(t_data):
                    ma_may = row.get('MA_MAY', row.get('ma_may', ''))
                    ma_dich_vu = row.get('MA_DICH_VU', row.get('MA_VАТ_TU', row.get('MA_DVKT', '')))
                    ten_dich_vu = row.get('TEN_DICH_VU', row.get('TEN_DVKT', row.get('TEN_VAT_TU', row.get('TEN_VАТ_TU', ''))))
                    
                    if not ma_may or ma_may.strip() == "":
                        has_cls_missing_machine = True
                        details.append({
                            "Bảng": t_name,
                            "Dòng số": idx + 1,
                            "Mã dịch vụ/VTYT": ma_dich_vu,
                            "Tên dịch vụ": ten_dich_vu,
                            "Lỗi": "Thiếu mã máy (MA_MAY trống)"
                        })
    return has_cls_missing_machine, details

# --- TIÊU ĐỀ GIAO DIỆN CHÍNH ---
st.markdown("""
    <div class="hospital-header" translate="no">
        <h1>🏥 BỆNH VIỆN QUÂN Y 5</h1>
        <p>HỆ THỐNG RÀ SOÁT & KIỂM ĐỊNH DỮ LIỆU XML BHYT (THEO QĐ 4750 & QĐ 3176)</p>
    </div>
""", unsafe_allow_html=True)

# --- THANH SIDEBAR SETTING TỔNG HỢP ---
with st.sidebar:
    st.markdown("### ⚙️ QUẢN TRỊ HỆ THỐNG")
    with st.expander("🛠️ Cấu hình & Danh mục chuẩn", expanded=True):
        source_mode = st.radio("Chọn chức năng chính:", ["Rà soát tệp/thư mục XML", "📦 Quản lý & Xem chi tiết Danh mục"])
        
        st.markdown("---")
        st.markdown("##### 📚 Quản lý danh mục chuẩn BHYT")
        
        new_cat_name = st.text_input("Tên danh mục mới cần thêm:")
        if st.button("➕ Thêm danh mục"):
            if new_cat_name and new_cat_name not in st.session_state['categories']:
                st.session_state['categories'][new_cat_name] = pd.DataFrame()
                save_category_to_disk(new_cat_name, pd.DataFrame())
                st.success(f"Đã thêm danh mục: {new_cat_name}")
                st.rerun()
            elif not new_cat_name:
                st.warning("Vui lòng nhập tên danh mục.")
            else:
                st.warning("Danh mục này đã tồn tại.")

        existing_cats = list(st.session_state['categories'].keys())
        if existing_cats:
            cat_to_delete = st.selectbox("Chọn danh mục để xóa:", existing_cats, key="del_cat_select")
            if st.button("🗑️ Xóa danh mục đã chọn"):
                if len(existing_cats) > 1:
                    del st.session_state['categories'][cat_to_delete]
                    delete_category_from_disk(cat_to_delete)
                    st.success(f"Đã xóa danh mục: {cat_to_delete}")
                    st.rerun()
                else:
                    st.error("Phải giữ lại ít nhất một danh mục!")

        st.markdown("---")
        st.markdown("##### 📥 Đẩy file Excel/CSV vào danh mục")
        target_cat = st.selectbox("Chọn danh mục đích:", list(st.session_state['categories'].keys()), key="target_cat_select")
        ref_catalog_file = st.file_uploader(f"Tải file lên [{target_cat}]:", type=['xlsx', 'xls', 'csv'], key="upload_cat_file")
        
        if ref_catalog_file is not None:
            if st.button("🚀 Cập nhật dữ liệu"):
                try:
                    if ref_catalog_file.name.endswith('.csv'):
                        df_uploaded = pd.read_csv(ref_catalog_file)
                    else:
                        df_uploaded = pd.read_excel(ref_catalog_file)
                    
                    st.session_state['categories'][target_cat] = df_uploaded
                    save_category_to_disk(target_cat, df_uploaded)
                    st.success(f"✅ Đã lưu {len(df_uploaded)} dòng vào [{target_cat}]!")
                except Exception as e:
                    st.error(f"Lỗi đọc file (kiểm tra openpyxl): {e}")

# --- HIỂN THỊ NHANH TRẠNG THÁI DANH MỤC TRÊN SIDEBAR ---
with st.sidebar.expander("📦 Danh sách danh mục hiện có"):
    for c_name, c_df in st.session_state['categories'].items():
        st.write(f"- **{c_name}**: {len(c_df)} bản ghi")

# --- XỬ LÝ GIAO DIỆN CHÍNH ---
if source_mode == "📦 Quản lý & Xem chi tiết Danh mục":
    st.subheader("📦 Quản lý và Xem chi tiết Danh mục chuẩn BHYT")
    
    selected_view_cat = st.selectbox("Chọn danh mục cần xem chi tiết:", list(st.session_state['categories'].keys()))
    
    if selected_view_cat:
        cat_df = st.session_state['categories'][selected_view_cat]
        if not cat_df.empty:
            st.info(f"📊 Tổng số bản ghi trong **{selected_view_cat}**: **{len(cat_df)}** dòng | **{len(cat_df.columns)}** cột")
            st.dataframe(cat_df, use_container_width=True)
            
            csv_cat_data = cat_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label=f"📥 Tải xuống danh mục [{selected_view_cat}] (CSV)",
                data=csv_cat_data,
                file_name=f"{selected_view_cat}.csv",
                mime="text/csv"
            )
        else:
            st.warning(f"⚠️ Danh mục **{selected_view_cat}** hiện đang trống. Hãy tải file lên ở thanh cài đặt bên trái.")

else:
    sub_source_mode = st.radio("Chọn phương thức đầu vào rà soát:", ["Quét tự động theo thư mục (Folder Scan)", "Tải lên từng tệp (File Uploader)"], horizontal=True)

    if sub_source_mode == "Quét tự động theo thư mục (Folder Scan)":
        st.subheader("📁 Quét và Kiểm định hàng loạt tệp XML theo thư mục")
        folder_path = st.text_input("Nhập đường dẫn tuyệt đối tới thư mục chứa tệp XML (Ví dụ: C:/DataXML):", value="")
        
        if folder_path and os.path.exists(folder_path):
            xml_files = list(pathlib.Path(folder_path).glob("*.xml"))
            
            if xml_files:
                st.info(f"🔍 Tìm thấy **{len(xml_files)}** tệp XML trong thư mục.")
                if st.button("🚀 Bắt đầu rà soát toàn bộ"):
                    folder_summary = []
                    all_folder_results = {}
                    progress_bar = st.progress(0)
                    
                    for idx, x_file in enumerate(xml_files):
                        file_results, is_valid, err_msg = process_xml_package(x_file)
                        all_folder_results[x_file.name] = file_results
                        
                        missing_machine_status = "Không có"
                        cls_details = []
                        if is_valid:
                            has_missing, cls_details = audit_cls_machine_code(file_results)
                            if has_missing:
                                missing_machine_status = f"Có thiếu ({len(cls_details)} dịch vụ)"
                            else:
                                missing_machine_status = "Đầy đủ mã máy"
                        else:
                            missing_machine_status = "Không kiểm tra (Sai định dạng)"
                        
                        folder_summary.append({
                            "Tên tệp": x_file.name,
                            "Trạng thái định dạng": "✅ Đúng định dạng (QĐ 4750/3176)" if is_valid else f"❌ Sai định dạng ({err_msg})",
                            "Kiểm tra mã máy CLS": missing_machine_status,
                            "Chi tiết lỗi CLS": cls_details
                        })
                        progress_bar.progress((idx + 1) / len(xml_files))
                    
                    st.success("🎉 Đã hoàn tất rà soát toàn bộ thư mục!")
                    st.session_state['folder_summary'] = pd.DataFrame(folder_summary)
                    st.session_state['folder_results'] = all_folder_results
            else:
                st.warning("⚠️ Không tìm thấy tệp XML nào trong thư mục này.")
        elif folder_path:
            st.error("❌ Đường dẫn thư mục không tồn tại hoặc không hợp lệ.")

        if 'folder_summary' in st.session_state:
            st.markdown("---")
            st.markdown("### 📊 Bảng tổng hợp kết quả rà soát thư mục")
            df_summary = st.session_state['folder_summary']
            
            display_df = df_summary[['Tên tệp', 'Trạng thái định dạng', 'Kiểm tra mã máy CLS']]
            st.dataframe(display_df, use_container_width=True)
            
            valid_count = len(df_summary[df_summary['Trạng thái định dạng'].str.contains('✅')])
            invalid_count = len(df_summary[df_summary['Trạng thái định dạng'].str.contains('❌')])
            col1, col2, col3 = st.columns(3)
            col1.metric("Tổng số tệp quét", len(df_summary))
            col2.metric("Số tệp Đúng định dạng", valid_count)
            col3.metric("Số tệp Sai định dạng", invalid_count)
            
            st.markdown("---")
            st.markdown("### 🔍 Xem chi tiết từng tệp và lỗi dịch vụ CLS")
            selected_file_name = st.selectbox("Chọn tệp XML để kiểm tra chi tiết:", df_summary['Tên tệp'].tolist())
            
            if selected_file_name:
                row_info = df_summary[df_summary['Tên tệp'] == selected_file_name].iloc[0]
                st.info(f"**Trạng thái tệp {selected_file_name}:** {row_info['Trạng thái định dạng']} | **Mã máy CLS:** {row_info['Kiểm tra mã máy CLS']}")
                
                if row_info['Chi tiết lỗi CLS']:
                    st.warning("⚠️ Phát hiện các dịch vụ cận lâm sàng thiếu mã máy:")
                    st.dataframe(pd.DataFrame(row_info['Chi tiết lỗi CLS']), use_container_width=True)
                else:
                    st.success("✨ Không phát hiện lỗi thiếu mã máy CLS trong các dịch vụ của file này.")
                    
                all_tables = st.session_state['folder_results'][selected_file_name]
                if all_tables:
                    sorted_tables = sorted(list(all_tables.keys()))
                    selected_table = st.selectbox("Xem cấu trúc bảng dữ liệu của tệp:", sorted_tables, key="sel_tbl_folder")
                    if selected_table:
                        data = all_tables[selected_table]
                        if isinstance(data, list) and len(data) > 0:
                            df_data = pd.DataFrame(data)
                            
                            compare_cat_name = st.selectbox("Chọn danh mục chuẩn để đối chiếu:", ["-- Không đối chiếu --"] + list(st.session_state['categories'].keys()))
                            
                            if compare_cat_name != "-- Không đối chiếu --":
                                standard_catalog = st.session_state['categories'][compare_cat_name]
                                if not standard_catalog.empty:
                                    st.markdown(f"#### 🔄 Kết quả đối chiếu với [{compare_cat_name}]:")
                                    matching_cols = [c for c in df_data.columns if 'MA_' in c.upper()]
                                    cat_cols = [c for c in standard_catalog.columns if 'MA_' in c.upper()]
                                    
                                    if matching_cols and cat_cols:
                                        match_col = matching_cols[0]
                                        cat_col = cat_cols[0]
                                        
                                        standard_codes = set(standard_catalog[cat_col].astype(str).str.strip().values)
                                        df_data['So_sanh_danh_muc'] = df_data[match_col].astype(str).str.strip().apply(
                                            lambda x: "✅ Khớp danh mục" if x in standard_codes else "❌ Không có trong danh mục chuẩn"
                                        )
                                        st.dataframe(df_data, use_container_width=True)
                                        
                                        invalid_cat_count = len(df_data[df_data['So_sanh_danh_muc'].str.contains('❌')])
                                        if invalid_cat_count > 0:
                                            st.warning(f"⚠️ Phát hiện {invalid_cat_count} bản ghi có mã không khớp với [{compare_cat_name}]!")
                                        else:
                                            st.success("✨ Tất cả mã đều khớp với danh mục chuẩn.")
                                    else:
                                        st.warning("Không tìm thấy cột mã tương thích (MA_...) để đối chiếu.")
                                        st.dataframe(df_data, use_container_width=True)
                                else:
                                    st.warning(f"Danh mục [{compare_cat_name}] hiện đang trống.")
                                    st.dataframe(df_data, use_container_width=True)
                            else:
                                st.dataframe(df_data, use_container_width=True)
                                
                        elif isinstance(data, dict):
                            st.dataframe(pd.DataFrame([{"Trường dữ liệu": k, "Giá trị": v} for k, v in data.items()]), use_container_width=True)

    else:
        uploaded_file = st.file_uploader("Tải lên tệp XML tổng hợp hoặc tệp XML thành phần", type=['xml'])
        if uploaded_file:
            file_results, is_valid, err_msg = process_xml_package(uploaded_file)
            
            if is_valid:
                st.success(f"✅ Tệp **{uploaded_file.name}** đúng định dạng chuẩn XML BHYT.")
                has_missing, cls_details = audit_cls_machine_code(file_results)
                if has_missing:
                    st.warning(f"⚠️ Phát hiện **{len(cls_details)}** dịch vụ cận lâm sàng bị **thiếu mã máy (`MA_MAY`)**!")
                    st.dataframe(pd.DataFrame(cls_details), use_container_width=True)
                else:
                    st.success("✨ Kiểm tra CLS: Tất cả các dịch vụ đều có đủ mã máy.")
            else:
                st.error(f"❌ Tệp **{uploaded_file.name}** **SAI ĐỊNH DẠNG**. Lỗi chi tiết: {err_msg}")